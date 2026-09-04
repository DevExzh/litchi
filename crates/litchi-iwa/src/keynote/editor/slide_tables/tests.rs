use super::*;
use crate::keynote::KeynoteDocumentBuilder;
use crate::numbers::cell::CellValue;
use litchi_iwa_common::table::cell::BorderSide;
use litchi_iwa_common::table::cell::conditional_highlight::{
    Condition, Rule, Style as ConditionalHighlightStyle, Text as ConditionalText,
};
use litchi_keynote::slide::table::title::Settings as TableTitleSettings;
use litchi_keynote::slide::table::{
    dimension::{Dimension, Size},
    formula::{FormulaCachedValue, FormulaCellReference, FormulaExpression},
};
use litchi_numbers::cell::data_format::control::{Range as ControlRange, Slider, Stepper};
use litchi_numbers::cell::data_format::custom::{
    Name as CustomFormatName, Number as CustomNumber, NumberPattern,
};
use litchi_numbers::cell::data_format::duration::{Duration, Style as DurationStyle, UnitRange};
use litchi_numbers::cell::data_format::number::{
    Currency, CurrencyCode, CurrencyStyle, DecimalPlaces, FixedDecimalPlaces, Fraction,
    FractionAccuracy, NegativeStyle, Number, Percentage, Scientific, ThousandsSeparator,
};
use litchi_numbers::cell::data_format::numeral_system::{
    Base, FixedPlaces, NegativeStyle as NumeralSystemNegativeStyle, Places,
};
use litchi_numbers::cell::data_format::pop_up_menu::PopUpMenu;
use litchi_numbers::cell::data_format::{Checkbox, DataFormat, StarRating, Text as TextFormat};
use litchi_numbers::table::headers::{Count as HeaderCount, Settings as HeaderSettings};

// Raw-ID header access remains only as compatibility coverage inside the
// host's cfg(test) module. Production callers use litchi-keynote's checked
// slide/table selectors and package transactions.
impl KeynoteEditor {
    fn slide_table_header_settings(
        &self,
        slide_index: usize,
        model_object_id: u64,
    ) -> Result<HeaderSettings> {
        require_table_model(self, slide_index, model_object_id)?;
        crate::numbers::editor::table_header_settings_in_package(self.package(), model_object_id)
    }

    fn set_slide_table_header_settings(
        &mut self,
        slide_index: usize,
        model_object_id: u64,
        settings: HeaderSettings,
    ) -> Result<()> {
        require_table_model(self, slide_index, model_object_id)?;
        let mut staged = self.package().clone();
        crate::numbers::editor::set_table_header_settings_in_package(
            &mut staged,
            model_object_id,
            settings,
        )?;
        *self = KeynoteEditor::from_package(staged)?;
        Ok(())
    }
}

fn focused_table_index(
    editor: &KeynoteEditor,
    slide_index: usize,
    model_object_id: u64,
) -> Result<usize> {
    let tables = editor.slide_tables(slide_index)?;
    let mut table_index = None;
    for (index, table) in tables.iter().enumerate() {
        if table.model_object_id != model_object_id {
            continue;
        }
        if table_index.replace(index).is_some() {
            return Err(Error::ParseError(format!(
                "Keynote object {model_object_id} has ambiguous table ownership on slide {slide_index}"
            )));
        }
    }
    table_index.ok_or_else(|| {
        Error::ParseError(format!(
            "Keynote table model {model_object_id} is not owned by slide {slide_index}"
        ))
    })
}

fn focused_table_package(editor: &KeynoteEditor) -> Result<litchi_keynote::Package> {
    litchi_keynote::Package::from_bytes(&editor.to_bytes()?).map_err(|error| {
        Error::InvalidFormat(format!("focused Keynote table source failed: {error}"))
    })
}

fn set_focused_table_name(
    editor: &mut KeynoteEditor,
    slide_index: usize,
    model_object_id: u64,
    before_name: &str,
    name: &str,
) -> Result<bool> {
    let before = editor.to_bytes()?;
    let table_index = focused_table_index(editor, slide_index, model_object_id)?;
    let package = focused_table_package(editor)?;
    let semantic_name = litchi_keynote::slide::table::name::Name::new(name).map_err(|error| {
        Error::InvalidFormat(format!("focused Keynote name value failed: {error}"))
    })?;
    let result = package
        .edit_slide_table_name(
            litchi_keynote::SlideSelector::index(slide_index),
            litchi_keynote::TableSelector::index(table_index),
        )
        .and_then(|edit| edit.set(semantic_name).commit());
    let commit = match result {
        Ok(commit) => commit,
        Err(
            litchi_keynote::SlideTableNameError::InvalidSource { .. }
            | litchi_keynote::SlideTableNameError::UnsupportedDependency
            | litchi_keynote::SlideTableNameError::UnsupportedSource,
        ) => {
            assert_eq!(editor.to_bytes()?, before);
            assert_eq!(
                editor.slide_table(slide_index, model_object_id)?.info.name,
                before_name
            );
            return Ok(false);
        },
        Err(error) => {
            return Err(Error::InvalidFormat(format!(
                "focused Keynote name edit failed: {error}"
            )));
        },
    };
    let mut bytes = Vec::new();
    commit.package().write_to(&mut bytes).map_err(|error| {
        Error::InvalidFormat(format!("focused Keynote name write failed: {error}"))
    })?;
    *editor = KeynoteEditor::from_bytes(&bytes)?;
    let package = focused_table_package(editor)?;
    let focused_name = package
        .slide_table_name(
            litchi_keynote::SlideSelector::index(slide_index),
            litchi_keynote::TableSelector::index(table_index),
        )
        .map_err(|error| {
            Error::InvalidFormat(format!("focused Keynote name read failed: {error}"))
        })?;
    assert_eq!(focused_name.as_str(), name);
    Ok(true)
}

fn focused_table_title_settings(
    editor: &KeynoteEditor,
    slide_index: usize,
    model_object_id: u64,
) -> Result<TableTitleSettings> {
    let table_index = focused_table_index(editor, slide_index, model_object_id)?;
    focused_table_package(editor)?
        .slide_table_title_settings(
            litchi_keynote::SlideSelector::index(slide_index),
            litchi_keynote::TableSelector::index(table_index),
        )
        .map_err(|error| {
            Error::InvalidFormat(format!("focused Keynote title read failed: {error}"))
        })
}

fn set_focused_table_title_settings(
    editor: &mut KeynoteEditor,
    slide_index: usize,
    model_object_id: u64,
    settings: TableTitleSettings,
) -> Result<()> {
    let table_index = focused_table_index(editor, slide_index, model_object_id)?;
    let commit = focused_table_package(editor)?
        .edit_slide_table_title(
            litchi_keynote::SlideSelector::index(slide_index),
            litchi_keynote::TableSelector::index(table_index),
        )
        .map_err(|error| {
            Error::InvalidFormat(format!("focused Keynote title edit failed: {error}"))
        })?
        .set(settings)
        .commit()
        .map_err(|error| {
            Error::InvalidFormat(format!("focused Keynote title commit failed: {error}"))
        })?;
    if commit.patch().is_noop() {
        return Ok(());
    }
    let mut bytes = Vec::new();
    commit.package().write_to(&mut bytes).map_err(|error| {
        Error::InvalidFormat(format!("focused Keynote title write failed: {error}"))
    })?;
    *editor = KeynoteEditor::from_bytes(&bytes)?;
    Ok(())
}

fn assert_focused_table_dimension_rejected(
    editor: &KeynoteEditor,
    slide_index: usize,
    model_object_id: u64,
    dimension: Dimension,
    size: Size,
) -> Result<()> {
    let before = editor.to_bytes()?;
    let table_index = focused_table_index(editor, slide_index, model_object_id)?;
    let result = focused_table_package(editor)?
        .edit_slide_table_dimension_size(
            litchi_keynote::SlideSelector::index(slide_index),
            litchi_keynote::TableSelector::index(table_index),
            dimension,
        )
        .and_then(|edit| edit.set(size).commit());
    assert!(
        matches!(
            result,
            Err(litchi_keynote::SlideTableDimensionError::UnsupportedDependency)
        ),
        "focused Keynote dimension edit unexpectedly admitted builder snapshot: {result:?}"
    );
    assert_eq!(editor.to_bytes()?, before);
    Ok(())
}

fn table_geometry() -> (DrawablePoint, DrawableSize) {
    (
        DrawablePoint { x: 120.0, y: 180.0 },
        DrawableSize {
            width: 840.0,
            height: 360.0,
        },
    )
}

#[test]
fn source_built_table_has_no_conditional_highlighting_and_clear_is_idempotent() {
    let mut editor = KeynoteDocumentBuilder::new().build().unwrap();
    let (position, size) = table_geometry();
    let table = editor
        .add_slide_table(0, "Conditional", 2, 2, position, size)
        .unwrap();

    assert!(
        editor
            .slide_table_cell_conditional_highlighting(0, table.model_object_id, 1, 1)
            .unwrap()
            .is_none()
    );
    editor
        .clear_slide_table_cell_conditional_highlighting(0, table.model_object_id, 1, 1)
        .unwrap();
    assert!(
        KeynoteEditor::from_bytes(&editor.to_bytes().unwrap())
            .unwrap()
            .slide_table_cell_conditional_highlighting(0, table.model_object_id, 1, 1)
            .unwrap()
            .is_none()
    );
}

#[test]
fn source_built_table_creates_and_replaces_conditional_highlighting() {
    let mut editor = KeynoteDocumentBuilder::new().build().unwrap();
    let (position, size) = table_geometry();
    let table = editor
        .add_slide_table(0, "Conditional", 2, 2, position, size)
        .unwrap();
    editor
        .set_slide_table_cell(
            0,
            table.model_object_id,
            1,
            1,
            CellValue::Text("Organic Grain".to_owned()),
        )
        .unwrap();
    let rule = Rule::new(
        Condition::TextContains(ConditionalText::new("grain").unwrap()),
        ConditionalHighlightStyle::with_text_color(
            crate::shapes::RgbaColor::new(0.1, 0.6, 0.2, 1.0, crate::shapes::RgbColorSpace::Srgb)
                .unwrap(),
        ),
    );
    let created = editor
        .set_slide_table_cell_conditional_highlighting(
            0,
            table.model_object_id,
            1,
            1,
            std::slice::from_ref(&rule),
        )
        .unwrap();
    assert_eq!(created.table_id, table.model_object_id);
    assert_eq!((created.row, created.column), (1, 1));
    assert_eq!(created.rule_count, 1);
    assert_eq!(
        editor
            .slide_table_cell_conditional_highlighting(0, table.model_object_id, 1, 1)
            .unwrap()
            .unwrap()
            .rule_count,
        1
    );
    assert_eq!(
        editor
            .slide_table_cell_conditional_highlight_rules(0, table.model_object_id, 1, 1)
            .unwrap()
            .unwrap(),
        vec![rule]
    );
}

#[test]
fn source_built_table_roundtrips_cell_border_crud() {
    let mut editor = KeynoteDocumentBuilder::new().build().unwrap();
    let (position, size) = table_geometry();
    let table = editor
        .add_slide_table(0, "Borders", 3, 3, position, size)
        .unwrap();
    let stroke = crate::shapes::Stroke::new(
        crate::shapes::RgbaColor::new(0.9, 0.2, 0.1, 1.0, crate::shapes::RgbColorSpace::Srgb)
            .unwrap(),
        crate::shapes::Width::new(3.0).unwrap(),
        crate::shapes::Pattern::Solid,
    );
    editor
        .set_slide_table_cell_border(0, table.model_object_id, 2, 1, BorderSide::Top, stroke)
        .unwrap();

    let mut reopened = KeynoteEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .slide_table_cell_borders(0, table.model_object_id, 2, 1)
            .unwrap()
            .top,
        Some(stroke)
    );
    reopened
        .clear_slide_table_cell_border(0, table.model_object_id, 2, 1, BorderSide::Top)
        .unwrap();
    assert_eq!(
        reopened
            .slide_table_cell_borders(0, table.model_object_id, 2, 1)
            .unwrap()
            .top,
        None
    );
}

#[test]
fn source_built_table_roundtrips_cell_fill_crud() {
    let mut editor = KeynoteDocumentBuilder::new().build().unwrap();
    let (position, size) = table_geometry();
    let table = editor
        .add_slide_table(0, "Fills", 3, 3, position, size)
        .unwrap();
    let inherited = editor
        .slide_table_cell_fill(0, table.model_object_id, 1, 1)
        .unwrap();
    let fill = crate::shapes::ShapeFill::Solid(
        crate::shapes::RgbaColor::new(0.95, 0.65, 0.1, 1.0, crate::shapes::RgbColorSpace::Srgb)
            .unwrap(),
    );
    editor
        .set_slide_table_cell_fill(0, table.model_object_id, 1, 1, &fill)
        .unwrap();

    let mut reopened = KeynoteEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .slide_table_cell_fill(0, table.model_object_id, 1, 1)
            .unwrap(),
        fill
    );
    assert!(
        reopened
            .reset_slide_table_cell_fill(0, table.model_object_id, 1, 1)
            .unwrap()
    );
    assert_eq!(
        reopened
            .slide_table_cell_fill(0, table.model_object_id, 1, 1)
            .unwrap(),
        inherited
    );
}

#[test]
fn source_built_table_roundtrips_cell_layout_crud() {
    let mut editor = KeynoteDocumentBuilder::new().build().unwrap();
    let (position, size) = table_geometry();
    let table = editor
        .add_slide_table(0, "Layout", 3, 3, position, size)
        .unwrap();
    let inherited = editor
        .slide_table_cell_layout(0, table.model_object_id, 1, 1)
        .unwrap();
    let layout = KeynoteTableCellLayout::default()
        .with_text_wrap(KeynoteTableCellTextWrap::Wrapped)
        .with_vertical_alignment(KeynoteTableCellVerticalAlignment::Middle)
        .with_insets(KeynoteTableCellInsets::uniform(
            KeynoteTableCellInset::from_points(7.0).unwrap(),
        ));
    editor
        .set_slide_table_cell_layout(0, table.model_object_id, 1, 1, layout)
        .unwrap();

    let mut reopened = KeynoteEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .slide_table_cell_layout(0, table.model_object_id, 1, 1)
            .unwrap(),
        layout
    );
    assert!(
        reopened
            .reset_slide_table_cell_layout(0, table.model_object_id, 1, 1)
            .unwrap()
    );
    assert_eq!(
        reopened
            .slide_table_cell_layout(0, table.model_object_id, 1, 1)
            .unwrap(),
        inherited
    );
}

#[test]
fn source_built_table_roundtrips_number_data_format_crud() {
    let mut editor = KeynoteDocumentBuilder::new().build().unwrap();
    let (position, size) = table_geometry();
    let table = editor
        .add_slide_table(0, "Formats", 3, 3, position, size)
        .unwrap();
    let format = Number::new(
        DecimalPlaces::fixed(2).unwrap(),
        NegativeStyle::Parentheses,
        ThousandsSeparator::Shown,
    );
    let data_format = DataFormat::Number(format);
    editor
        .set_slide_table_cell(
            0,
            table.model_object_id,
            1,
            1,
            CellValue::number(1_234.5).expect("finite cell number"),
        )
        .unwrap();
    editor
        .set_slide_table_cell_data_format(0, table.model_object_id, 1, 1, data_format.clone())
        .unwrap();

    let mut reopened = KeynoteEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .slide_table_cell_data_format(0, table.model_object_id, 1, 1)
            .unwrap(),
        data_format
    );
    reopened
        .set_slide_table_cell_data_format(0, table.model_object_id, 1, 1, DataFormat::Automatic)
        .unwrap();
    assert_eq!(
        reopened
            .slide_table_cell_data_format(0, table.model_object_id, 1, 1)
            .unwrap(),
        DataFormat::Automatic
    );
}

#[test]
fn source_built_table_roundtrips_percentage_data_format() {
    let mut editor = KeynoteDocumentBuilder::new().build().unwrap();
    let (position, size) = table_geometry();
    let table = editor
        .add_slide_table(0, "Percentages", 3, 3, position, size)
        .unwrap();
    let format = Percentage::new(
        DecimalPlaces::fixed(2).unwrap(),
        NegativeStyle::Parentheses,
        ThousandsSeparator::Shown,
    );
    editor
        .set_slide_table_cell(
            0,
            table.model_object_id,
            1,
            1,
            CellValue::number(-12.345).expect("finite cell number"),
        )
        .unwrap();
    editor
        .set_slide_table_cell_percentage_format(0, table.model_object_id, 1, 1, format)
        .unwrap();

    let mut reopened = KeynoteEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .slide_table_cell_data_format(0, table.model_object_id, 1, 1)
            .unwrap(),
        DataFormat::Percentage(format)
    );
    assert_eq!(
        reopened
            .slide_table_cell_percentage_format(0, table.model_object_id, 1, 1)
            .unwrap(),
        Some(format)
    );
    reopened
        .set_slide_table_cell_data_format(0, table.model_object_id, 1, 1, DataFormat::Automatic)
        .unwrap();
    assert_eq!(
        reopened
            .slide_table_cell_data_format(0, table.model_object_id, 1, 1)
            .unwrap(),
        DataFormat::Automatic
    );
}

#[test]
fn source_built_table_roundtrips_currency_format_crud() {
    let mut editor = KeynoteDocumentBuilder::new().build().unwrap();
    let (position, size) = table_geometry();
    let table = editor
        .add_slide_table(0, "Currencies", 3, 3, position, size)
        .unwrap();
    let format = Currency::new(
        CurrencyCode::GBP,
        DecimalPlaces::fixed(2).unwrap(),
        NegativeStyle::Parentheses,
        ThousandsSeparator::Shown,
        CurrencyStyle::Accounting,
    );
    editor
        .set_slide_table_cell(
            0,
            table.model_object_id,
            1,
            1,
            CellValue::number(-1_234.5).expect("finite cell number"),
        )
        .unwrap();
    editor
        .set_slide_table_cell_currency_format(0, table.model_object_id, 1, 1, format)
        .unwrap();

    let mut reopened = KeynoteEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .slide_table_cell_currency_format(0, table.model_object_id, 1, 1)
            .unwrap(),
        Some(format)
    );
    assert!(
        reopened
            .reset_slide_table_cell_currency_format(0, table.model_object_id, 1, 1)
            .unwrap()
    );
    assert_eq!(
        reopened
            .slide_table_cell_data_format(0, table.model_object_id, 1, 1)
            .unwrap(),
        DataFormat::Automatic
    );
}

#[test]
fn source_built_table_roundtrips_scientific_format_crud() {
    let mut editor = KeynoteDocumentBuilder::new().build().unwrap();
    let (position, size) = table_geometry();
    let table = editor
        .add_slide_table(0, "Scientific", 3, 3, position, size)
        .unwrap();
    let format = Scientific::new(FixedDecimalPlaces::new(5).unwrap());
    editor
        .set_slide_table_cell(
            0,
            table.model_object_id,
            1,
            1,
            CellValue::number(-1_234.5).expect("finite cell number"),
        )
        .unwrap();
    editor
        .set_slide_table_cell_scientific_format(0, table.model_object_id, 1, 1, format)
        .unwrap();

    let mut reopened = KeynoteEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .slide_table_cell_scientific_format(0, table.model_object_id, 1, 1)
            .unwrap(),
        Some(format)
    );
    assert!(
        reopened
            .reset_slide_table_cell_scientific_format(0, table.model_object_id, 1, 1)
            .unwrap()
    );
    assert_eq!(
        reopened
            .slide_table_cell_data_format(0, table.model_object_id, 1, 1)
            .unwrap(),
        DataFormat::Automatic
    );
}

#[test]
fn source_built_table_roundtrips_fraction_format_crud() {
    let mut editor = KeynoteDocumentBuilder::new().build().unwrap();
    let (position, size) = table_geometry();
    let table = editor
        .add_slide_table(0, "Fractions", 3, 3, position, size)
        .unwrap();
    let format = Fraction::new(FractionAccuracy::Eighths);
    editor
        .set_slide_table_cell(
            0,
            table.model_object_id,
            1,
            1,
            CellValue::number(-12.375).expect("finite cell number"),
        )
        .unwrap();
    editor
        .set_slide_table_cell_fraction_format(0, table.model_object_id, 1, 1, format)
        .unwrap();

    let mut reopened = KeynoteEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .slide_table_cell_fraction_format(0, table.model_object_id, 1, 1)
            .unwrap(),
        Some(format)
    );
    assert!(
        reopened
            .reset_slide_table_cell_fraction_format(0, table.model_object_id, 1, 1)
            .unwrap()
    );
    assert_eq!(
        reopened
            .slide_table_cell_data_format(0, table.model_object_id, 1, 1)
            .unwrap(),
        DataFormat::Automatic
    );
}

#[test]
fn source_built_table_roundtrips_numeral_system_format_crud() {
    let mut editor = KeynoteDocumentBuilder::new().build().unwrap();
    let (position, size) = table_geometry();
    let table = editor
        .add_slide_table(0, "Numeral Systems", 3, 3, position, size)
        .unwrap();
    let format = NumeralSystem::new(
        Base::HEXADECIMAL,
        Places::Fixed(FixedPlaces::EIGHT),
        NumeralSystemNegativeStyle::TwosComplement,
    )
    .unwrap();
    editor
        .set_slide_table_cell(
            0,
            table.model_object_id,
            1,
            1,
            CellValue::number(-1_234.5).expect("finite cell number"),
        )
        .unwrap();
    editor
        .set_slide_table_cell_numeral_system_format(0, table.model_object_id, 1, 1, format)
        .unwrap();

    let mut reopened = KeynoteEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .slide_table_cell_numeral_system_format(0, table.model_object_id, 1, 1)
            .unwrap(),
        Some(format)
    );
    assert!(
        reopened
            .reset_slide_table_cell_numeral_system_format(0, table.model_object_id, 1, 1)
            .unwrap()
    );
    assert_eq!(
        reopened
            .slide_table_cell_data_format(0, table.model_object_id, 1, 1)
            .unwrap(),
        DataFormat::Automatic
    );
}

#[test]
fn source_built_table_roundtrips_date_time_format_crud() {
    let mut editor = KeynoteDocumentBuilder::new().build().unwrap();
    let (position, size) = table_geometry();
    let table = editor
        .add_slide_table(0, "Dates", 3, 3, position, size)
        .unwrap();
    let format = DateTime::iso_date_time_24_hour_with_seconds();
    editor
        .set_slide_table_cell(
            0,
            table.model_object_id,
            1,
            1,
            CellValue::date(789_332_889.0).expect("finite cell date"),
        )
        .unwrap();
    editor
        .set_slide_table_cell_date_time_format(0, table.model_object_id, 1, 1, format.clone())
        .unwrap();

    let mut reopened = KeynoteEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .slide_table_cell_date_time_format(0, table.model_object_id, 1, 1)
            .unwrap(),
        Some(format)
    );
    assert!(
        reopened
            .reset_slide_table_cell_date_time_format(0, table.model_object_id, 1, 1)
            .unwrap()
    );
    assert_eq!(
        reopened
            .slide_table_cell_data_format(0, table.model_object_id, 1, 1)
            .unwrap(),
        DataFormat::Automatic
    );
}

#[test]
fn source_built_table_roundtrips_duration_format_crud() {
    let mut editor = KeynoteDocumentBuilder::new().build().unwrap();
    let (position, size) = table_geometry();
    let table = editor
        .add_slide_table(0, "Durations", 3, 3, position, size)
        .unwrap();
    let range = UnitRange::hours_to_milliseconds();
    let format = Duration::custom(DurationStyle::Abbreviated, range);
    editor
        .set_slide_table_cell(
            0,
            table.model_object_id,
            1,
            1,
            CellValue::duration(3_723.5).expect("finite cell duration"),
        )
        .unwrap();
    editor
        .set_slide_table_cell_duration_format(0, table.model_object_id, 1, 1, format)
        .unwrap();

    let mut reopened = KeynoteEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .slide_table_cell_duration_format(0, table.model_object_id, 1, 1)
            .unwrap(),
        Some(format)
    );
    assert!(
        reopened
            .reset_slide_table_cell_duration_format(0, table.model_object_id, 1, 1)
            .unwrap()
    );
    assert_eq!(
        reopened
            .slide_table_cell_data_format(0, table.model_object_id, 1, 1)
            .unwrap(),
        DataFormat::Automatic
    );
}

#[test]
fn source_built_table_roundtrips_checkbox_format_crud() {
    let mut editor = KeynoteDocumentBuilder::new().build().unwrap();
    let (position, size) = table_geometry();
    let table = editor
        .add_slide_table(0, "Checkboxes", 3, 3, position, size)
        .unwrap();
    editor
        .set_slide_table_cell_checkbox_format(0, table.model_object_id, 1, 1, Checkbox)
        .unwrap();

    let mut reopened = KeynoteEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .slide_table_cell_checkbox_format(0, table.model_object_id, 1, 1)
            .unwrap(),
        Some(Checkbox)
    );
    assert_eq!(
        reopened
            .slide_table(0, table.model_object_id)
            .unwrap()
            .get_cell(1, 1),
        Some(&KeynoteTableCellValue::Boolean(false))
    );
    assert!(
        reopened
            .reset_slide_table_cell_checkbox_format(0, table.model_object_id, 1, 1)
            .unwrap()
    );
    assert_eq!(
        reopened
            .slide_table_cell_data_format(0, table.model_object_id, 1, 1)
            .unwrap(),
        DataFormat::Automatic
    );
}

#[test]
fn source_built_table_roundtrips_star_rating_format_crud() {
    let mut editor = KeynoteDocumentBuilder::new().build().unwrap();
    let (position, size) = table_geometry();
    let table = editor
        .add_slide_table(0, "Ratings", 3, 3, position, size)
        .unwrap();
    editor
        .set_slide_table_cell_star_rating_format(0, table.model_object_id, 1, 1, StarRating)
        .unwrap();

    let mut reopened = KeynoteEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .slide_table_cell_star_rating_format(0, table.model_object_id, 1, 1)
            .unwrap(),
        Some(StarRating)
    );
    assert_eq!(
        reopened
            .slide_table(0, table.model_object_id)
            .unwrap()
            .get_cell(1, 1),
        Some(&CellValue::number(0.0).expect("finite cell number"))
    );
    assert!(
        reopened
            .reset_slide_table_cell_star_rating_format(0, table.model_object_id, 1, 1)
            .unwrap()
    );
    assert_eq!(
        reopened
            .slide_table_cell_data_format(0, table.model_object_id, 1, 1)
            .unwrap(),
        DataFormat::Automatic
    );
}

#[test]
fn source_built_table_roundtrips_slider_format_crud() {
    let mut editor = KeynoteDocumentBuilder::new().build().unwrap();
    let (position, size) = table_geometry();
    let table = editor
        .add_slide_table(0, "Sliders", 3, 3, position, size)
        .unwrap();
    let range = ControlRange::new(-10.0, 30.0, 0.5).unwrap();
    let format = Slider::new(range, Number::default().into());
    editor
        .set_slide_table_cell_slider_format(0, table.model_object_id, 1, 1, format.clone())
        .unwrap();

    let mut reopened = KeynoteEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .slide_table_cell_slider_format(0, table.model_object_id, 1, 1)
            .unwrap(),
        Some(format)
    );
    assert_eq!(
        reopened
            .slide_table(0, table.model_object_id)
            .unwrap()
            .get_cell(1, 1),
        Some(&CellValue::number(10.0).expect("finite cell number"))
    );
    assert!(
        reopened
            .reset_slide_table_cell_slider_format(0, table.model_object_id, 1, 1)
            .unwrap()
    );
    assert_eq!(
        reopened
            .slide_table_cell_data_format(0, table.model_object_id, 1, 1)
            .unwrap(),
        DataFormat::Automatic
    );
}

#[test]
fn source_built_table_roundtrips_stepper_format_crud() {
    let mut editor = KeynoteDocumentBuilder::new().build().unwrap();
    let (position, size) = table_geometry();
    let table = editor
        .add_slide_table(0, "Steppers", 3, 3, position, size)
        .unwrap();
    let range = ControlRange::new(-10.0, 30.0, 0.5).unwrap();
    let format = Stepper::new(range, Number::default().into());
    editor
        .set_slide_table_cell_stepper_format(0, table.model_object_id, 1, 1, format.clone())
        .unwrap();

    let mut reopened = KeynoteEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .slide_table_cell_stepper_format(0, table.model_object_id, 1, 1)
            .unwrap(),
        Some(format)
    );
    assert_eq!(
        reopened
            .slide_table(0, table.model_object_id)
            .unwrap()
            .get_cell(1, 1),
        Some(&CellValue::number(-10.0).expect("finite cell number"))
    );
    assert!(
        reopened
            .reset_slide_table_cell_stepper_format(0, table.model_object_id, 1, 1)
            .unwrap()
    );
    assert_eq!(
        reopened
            .slide_table_cell_data_format(0, table.model_object_id, 1, 1)
            .unwrap(),
        DataFormat::Automatic
    );
}

#[test]
fn source_built_table_roundtrips_text_format_crud() {
    let mut editor = KeynoteDocumentBuilder::new().build().unwrap();
    let (position, size) = table_geometry();
    let table = editor
        .add_slide_table(0, "Text", 3, 3, position, size)
        .unwrap();
    editor
        .set_slide_table_cell(
            0,
            table.model_object_id,
            1,
            1,
            KeynoteTableCellValue::Text("00123".to_owned()),
        )
        .unwrap();
    editor
        .set_slide_table_cell_text_format(0, table.model_object_id, 1, 1)
        .unwrap();

    let mut reopened = KeynoteEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .slide_table_cell_text_format(0, table.model_object_id, 1, 1)
            .unwrap(),
        Some(TextFormat)
    );
    assert_eq!(
        reopened
            .slide_table(0, table.model_object_id)
            .unwrap()
            .get_cell(1, 1),
        Some(&KeynoteTableCellValue::Text("00123".to_owned()))
    );
    assert!(
        reopened
            .reset_slide_table_cell_text_format(0, table.model_object_id, 1, 1)
            .unwrap()
    );
    assert_eq!(
        reopened
            .slide_table_cell_data_format(0, table.model_object_id, 1, 1)
            .unwrap(),
        DataFormat::Automatic
    );
}

#[test]
fn source_built_table_roundtrips_custom_format_crud() {
    let mut editor = KeynoteDocumentBuilder::new().build().unwrap();
    let (position, size) = table_geometry();
    let table = editor
        .add_slide_table(0, "Custom Numbers", 3, 3, position, size)
        .unwrap();
    let format = Custom::Number(CustomNumber::new(
        CustomFormatName::try_new("Grouped Integer").unwrap(),
        NumberPattern::try_new("#,###").unwrap(),
    ));
    editor
        .set_slide_table_cell(
            0,
            table.model_object_id,
            1,
            1,
            CellValue::number(12_345.0).expect("finite cell number"),
        )
        .unwrap();
    editor
        .set_slide_table_cell_custom_format(0, table.model_object_id, 1, 1, format.clone())
        .unwrap();

    let mut reopened = KeynoteEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .slide_table_cell_custom_format(0, table.model_object_id, 1, 1)
            .unwrap(),
        Some(format)
    );
    assert!(
        reopened
            .reset_slide_table_cell_custom_format(0, table.model_object_id, 1, 1)
            .unwrap()
    );
    assert_eq!(
        reopened
            .slide_table_cell_data_format(0, table.model_object_id, 1, 1)
            .unwrap(),
        DataFormat::Automatic
    );
}

#[test]
fn source_built_table_roundtrips_pop_up_menu_format_crud() {
    let mut editor = KeynoteDocumentBuilder::new().build().unwrap();
    let (position, size) = table_geometry();
    let table = editor
        .add_slide_table(0, "Menus", 3, 3, position, size)
        .unwrap();
    let format = PopUpMenu::new(["Draft", "Published"]).unwrap();
    editor
        .set_slide_table_cell_pop_up_menu_format(0, table.model_object_id, 1, 1, format.clone())
        .unwrap();

    let mut reopened = KeynoteEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .slide_table_cell_pop_up_menu_format(0, table.model_object_id, 1, 1)
            .unwrap(),
        Some(format)
    );
    assert_eq!(
        reopened
            .slide_table(0, table.model_object_id)
            .unwrap()
            .get_cell(1, 1),
        Some(&KeynoteTableCellValue::Text("Draft".to_owned()))
    );
    assert!(
        reopened
            .reset_slide_table_cell_pop_up_menu_format(0, table.model_object_id, 1, 1)
            .unwrap()
    );
    assert_eq!(
        reopened
            .slide_table_cell_data_format(0, table.model_object_id, 1, 1)
            .unwrap(),
        DataFormat::Automatic
    );
}

#[test]
fn source_built_table_roundtrips_full_crud() {
    let editor = KeynoteDocumentBuilder::new().build().unwrap();
    assert!(editor.slide_tables(0).unwrap().is_empty());
    let mut package = editor.into_package();
    let engine = package.remove_entry("Index/CalculationEngine.iwa").unwrap();
    package
        .insert_entry("Index/CalculationEngine-81.iwa", engine)
        .unwrap();
    let mut editor = KeynoteEditor::from_package(package).unwrap();
    let (position, size) = table_geometry();
    let table = editor
        .add_slide_table(0, "Forecast", 3, 3, position, size)
        .unwrap();
    assert_eq!(table.name, "Forecast");
    assert_eq!((table.rows, table.columns), (3, 3));

    assert_eq!(
        editor
            .set_slide_table_cells(
                0,
                table.model_object_id,
                [
                    KeynoteTableCellUpdate::new(
                        0,
                        0,
                        KeynoteTableCellValue::Text("Region".to_owned()),
                    ),
                    KeynoteTableCellUpdate::new(
                        1,
                        1,
                        CellValue::number(42.5).expect("finite cell number"),
                    ),
                ],
            )
            .unwrap(),
        2
    );
    let before_invalid_batch = editor.to_bytes().unwrap();
    assert!(
        editor
            .set_slide_table_cells(
                0,
                table.model_object_id,
                [
                    KeynoteTableCellUpdate::new(2, 0, KeynoteTableCellValue::Boolean(true),),
                    KeynoteTableCellUpdate::clear(2, 0),
                ],
            )
            .is_err()
    );
    assert_eq!(editor.to_bytes().unwrap(), before_invalid_batch);
    let focused_name_changed =
        set_focused_table_name(&mut editor, 0, table.model_object_id, "Forecast", "Outlook")
            .unwrap();
    editor
        .resize_slide_table(0, table.model_object_id, 4, 4)
        .unwrap();
    let replacement = DrawableGeometry {
        position: Some(DrawablePoint { x: 180.0, y: 210.0 }),
        size: Some(DrawableSize {
            width: 900.0,
            height: 420.0,
        }),
        flags: Some(TABLE_GEOMETRY_FLAGS),
        angle: Some(TABLE_ANGLE_DEGREES),
    };
    editor
        .set_slide_table_geometry(0, table.drawable_object_id, replacement)
        .unwrap();
    let headers = HeaderSettings {
        header_rows: Some(HeaderCount::TWO),
        header_columns: None,
        footer_rows: Some(HeaderCount::ONE),
        ..Default::default()
    };
    editor
        .set_slide_table_header_settings(0, table.model_object_id, headers)
        .unwrap();
    assert_focused_table_dimension_rejected(
        &editor,
        0,
        table.model_object_id,
        Dimension::Column(0),
        Size::points(300.0).unwrap(),
    )
    .unwrap();
    assert_focused_table_dimension_rejected(
        &editor,
        0,
        table.model_object_id,
        Dimension::Row(0),
        Size::points(140.0).unwrap(),
    )
    .unwrap();
    let laid_out_geometry = replacement;

    let bytes = editor.to_bytes().unwrap();
    let mut reopened = KeynoteEditor::from_bytes(&bytes).unwrap();
    let materialized = reopened.slide_table(0, table.model_object_id).unwrap();
    assert_eq!(
        materialized.info.name,
        if focused_name_changed {
            "Outlook"
        } else {
            "Forecast"
        }
    );
    assert_eq!((materialized.info.rows, materialized.info.columns), (4, 4));
    assert_eq!(materialized.info.geometry, laid_out_geometry);
    assert_eq!(
        reopened
            .slide_table_header_settings(0, table.model_object_id)
            .unwrap(),
        headers
    );
    assert_eq!(
        materialized.get_cell(0, 0),
        Some(&KeynoteTableCellValue::Text("Region".to_owned()))
    );
    assert_eq!(
        materialized.get_cell(1, 1),
        Some(&CellValue::number(42.5).expect("finite cell number"))
    );

    reopened
        .clear_slide_table_cell(0, table.model_object_id, 1, 1)
        .unwrap();
    assert!(
        reopened
            .slide_table(0, table.model_object_id)
            .unwrap()
            .get_cell(1, 1)
            .is_none_or(KeynoteTableCellValue::is_empty)
    );
    let removed = reopened
        .remove_slide_table(0, table.drawable_object_id)
        .unwrap();
    assert_eq!(removed.table.model_object_id, table.model_object_id);
    assert!(reopened.slide_tables(0).unwrap().is_empty());
    assert!(KeynoteEditor::from_bytes(&reopened.to_bytes().unwrap()).is_ok());
}

#[test]
fn source_built_table_duplication_clones_formula_storage_and_geometry() {
    let mut editor = KeynoteDocumentBuilder::new().build().unwrap();
    let (position, size) = table_geometry();
    let source = editor
        .add_slide_table(0, "Budget", 3, 2, position, size)
        .unwrap();
    editor
        .set_slide_table_cells(
            0,
            source.model_object_id,
            [
                KeynoteTableCellUpdate::new(
                    0,
                    0,
                    KeynoteTableCellValue::Text("Category".to_owned()),
                ),
                KeynoteTableCellUpdate::new(1, 0, KeynoteTableCellValue::Text("Travel".to_owned())),
                KeynoteTableCellUpdate::new(
                    1,
                    1,
                    CellValue::number(125.0).expect("finite cell number"),
                ),
            ],
        )
        .unwrap();
    editor
        .set_slide_table_formula(
            0,
            source.model_object_id,
            2,
            1,
            FormulaExpression::function(
                "SUM",
                [
                    FormulaExpression::Number(100.0),
                    FormulaExpression::Number(25.0),
                ],
            ),
            FormulaCachedValue::number(125.0).expect("finite cached number"),
        )
        .unwrap();

    let copied = editor
        .duplicate_slide_table(0, source.drawable_object_id)
        .unwrap();
    assert_ne!(copied.drawable_object_id, source.drawable_object_id);
    assert_ne!(copied.model_object_id, source.model_object_id);
    assert_eq!(copied.name, "Budget copy");
    assert_eq!((copied.rows, copied.columns), (source.rows, source.columns));
    let mut expected_geometry = source.geometry;
    if let Some(position) = expected_geometry.position.as_mut() {
        position.x += TABLE_DUPLICATE_OFFSET;
        position.y += TABLE_DUPLICATE_OFFSET;
    }
    assert_eq!(copied.geometry, expected_geometry);
    assert_eq!(
        editor
            .slide_table(copied.slide_index, copied.model_object_id)
            .unwrap()
            .get_cell(1, 0),
        Some(&KeynoteTableCellValue::Text("Travel".to_owned()))
    );
    assert_eq!(
        editor
            .slide_table_formula(copied.slide_index, copied.model_object_id, 2, 1)
            .unwrap()
            .as_deref(),
        Some("=SUM(100,25)")
    );

    editor
        .set_slide_table_cell(
            copied.slide_index,
            copied.model_object_id,
            1,
            0,
            KeynoteTableCellValue::Text("Lodging".to_owned()),
        )
        .unwrap();
    assert_eq!(
        editor
            .slide_table(source.slide_index, source.model_object_id)
            .unwrap()
            .get_cell(1, 0),
        Some(&KeynoteTableCellValue::Text("Travel".to_owned()))
    );
    assert_eq!(
        editor
            .duplicate_slide_table(0, source.drawable_object_id)
            .unwrap()
            .name,
        "Budget copy 2"
    );

    let mut reopened = KeynoteEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(reopened.slide_tables(0).unwrap().len(), 3);
    reopened
        .remove_slide_table(copied.slide_index, copied.drawable_object_id)
        .unwrap();
    assert_eq!(reopened.slide_tables(0).unwrap().len(), 2);
    assert_eq!(
        reopened
            .slide_table(source.slide_index, source.model_object_id)
            .unwrap()
            .get_cell(1, 0),
        Some(&KeynoteTableCellValue::Text("Travel".to_owned()))
    );
}

#[test]
fn source_built_table_roundtrips_formula_crud_transactionally() {
    let editor = KeynoteDocumentBuilder::new().build().unwrap();
    let mut package = editor.into_package();
    let engine = package.remove_entry("Index/CalculationEngine.iwa").unwrap();
    package
        .insert_entry("Index/CalculationEngine-82.iwa", engine)
        .unwrap();
    let mut editor = KeynoteEditor::from_package(package).unwrap();
    let (position, size) = table_geometry();
    let table = editor
        .add_slide_table(0, "Formula", 3, 2, position, size)
        .unwrap();
    editor
        .set_slide_table_formula(
            0,
            table.model_object_id,
            2,
            1,
            FormulaExpression::function(
                "SUM",
                [
                    FormulaExpression::Number(1.0),
                    FormulaExpression::Number(2.0),
                ],
            ),
            FormulaCachedValue::number(3.0).expect("finite cached number"),
        )
        .unwrap();

    let mut reopened = KeynoteEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .slide_table_formula(0, table.model_object_id, 2, 1)
            .unwrap()
            .as_deref(),
        Some("=SUM(1,2)")
    );
    reopened
        .set_slide_table_formula(
            0,
            table.model_object_id,
            2,
            1,
            FormulaExpression::function(
                "SUM",
                [
                    FormulaExpression::Number(3.0),
                    FormulaExpression::Number(4.0),
                ],
            ),
            FormulaCachedValue::number(7.0).expect("finite cached number"),
        )
        .unwrap();
    assert_eq!(
        reopened
            .slide_table_formula(0, table.model_object_id, 2, 1)
            .unwrap()
            .as_deref(),
        Some("=SUM(3,4)")
    );

    let before = reopened.to_bytes().unwrap();
    assert!(
        reopened
            .set_slide_table_formula(
                0,
                table.model_object_id,
                usize::MAX,
                1,
                FormulaExpression::Number(1.0),
                FormulaCachedValue::number(1.0).expect("finite cached number"),
            )
            .is_err()
    );
    assert_eq!(reopened.to_bytes().unwrap(), before);
    assert_eq!(
        reopened
            .clear_slide_table_formula(0, table.model_object_id, 2, 1)
            .unwrap(),
        "=SUM(3,4)"
    );
    assert_eq!(
        reopened
            .slide_table_formula(0, table.model_object_id, 2, 1)
            .unwrap(),
        None
    );
    let cleared = reopened.to_bytes().unwrap();
    assert!(
        reopened
            .clear_slide_table_formula(0, table.model_object_id, 2, 1)
            .is_err()
    );
    assert_eq!(reopened.to_bytes().unwrap(), cleared);
}

#[test]
fn source_built_table_roundtrips_section_relative_axis_crud_transactionally() {
    let mut editor = KeynoteDocumentBuilder::new().build().unwrap();
    let (position, size) = table_geometry();
    let table = editor
        .add_slide_table(0, "Topology", 4, 4, position, size)
        .unwrap();
    let model_id = table.model_object_id;
    let row_size = Size::points(88.0).unwrap();
    let column_size = Size::points(144.0).unwrap();
    editor
        .set_slide_table_cell(
            0,
            model_id,
            1,
            1,
            KeynoteTableCellValue::Text("shift me".to_owned()),
        )
        .unwrap();
    editor
        .set_slide_table_formula(
            0,
            model_id,
            2,
            2,
            FormulaExpression::cell(FormulaCellReference::relative(1, 1)),
            FormulaCachedValue::number(7.0).expect("finite cached number"),
        )
        .unwrap();
    assert_focused_table_dimension_rejected(&editor, 0, model_id, Dimension::Row(1), row_size)
        .unwrap();
    assert_focused_table_dimension_rejected(
        &editor,
        0,
        model_id,
        Dimension::Column(1),
        column_size,
    )
    .unwrap();
    let baseline_geometry = editor.slide_tables(0).unwrap()[0].geometry;
    let baseline = editor.to_bytes().unwrap();

    editor
        .insert_slide_table_row(0, model_id, RowInsertion::body(1))
        .unwrap();
    editor
        .insert_slide_table_column(0, model_id, ColumnInsertion::body(1))
        .unwrap();
    let reopened = KeynoteEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    let shifted = reopened.slide_table(0, model_id).unwrap();
    assert_eq!((shifted.info.rows, shifted.info.columns), (5, 5));
    assert_ne!(shifted.info.geometry.size, baseline_geometry.size);
    assert_eq!(
        shifted.get_cell(1, 1),
        Some(&KeynoteTableCellValue::Text("shift me".to_owned()))
    );
    assert_eq!(
        shifted.get_cell(3, 3),
        Some(&KeynoteTableCellValue::Formula("=B2".to_owned()))
    );
    editor
        .remove_slide_table_column(0, model_id, ColumnDeletion::body(1))
        .unwrap();
    editor
        .remove_slide_table_row(0, model_id, RowDeletion::body(1))
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);
    assert_eq!(
        editor.slide_tables(0).unwrap()[0].geometry,
        baseline_geometry
    );

    let before_error = editor.to_bytes().unwrap();
    assert!(
        editor
            .insert_slide_table_row(1, model_id, RowInsertion::body(usize::MAX))
            .is_err()
    );
    assert!(
        editor
            .remove_slide_table_column(0, model_id, ColumnDeletion::body(usize::MAX))
            .is_err()
    );
    assert_eq!(editor.to_bytes().unwrap(), before_error);
}

#[test]
fn source_built_footer_formula_expands_and_contracts_with_body_rows() {
    let mut editor = KeynoteDocumentBuilder::new().build().unwrap();
    let (position, size) = table_geometry();
    let table = editor
        .add_slide_table(0, "Footer aggregate", 4, 3, position, size)
        .unwrap();
    let model_id = table.model_object_id;
    editor
        .set_slide_table_header_settings(
            0,
            model_id,
            HeaderSettings {
                footer_rows: Some(HeaderCount::ONE),
                ..Default::default()
            },
        )
        .unwrap();
    editor
        .set_slide_table_formula(
            0,
            model_id,
            3,
            1,
            FormulaExpression::function(
                "SUM",
                [FormulaExpression::range(
                    FormulaCellReference::relative(1, 1),
                    FormulaCellReference::relative(2, 1),
                )],
            ),
            FormulaCachedValue::number(3.0).expect("finite cached number"),
        )
        .unwrap();
    let mut package = editor.into_package();
    let engine = package.remove_entry("Index/CalculationEngine.iwa").unwrap();
    package
        .insert_entry("Index/CalculationEngine-42-2.iwa", engine)
        .unwrap();
    let mut editor = KeynoteEditor::from_package(package).unwrap();
    let baseline = editor.to_bytes().unwrap();

    editor
        .insert_slide_table_row(0, model_id, RowInsertion::body(3))
        .unwrap();
    assert_eq!(
        editor
            .slide_table_formula(0, model_id, 4, 1)
            .unwrap()
            .as_deref(),
        Some("=SUM(B2:B4)")
    );
    editor
        .remove_slide_table_row(0, model_id, RowDeletion::body(3))
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);
}

#[test]
fn source_built_fixed_table_sections_roundtrip_full_axis_crud() {
    let mut editor = KeynoteDocumentBuilder::new().build().unwrap();
    let (position, size) = table_geometry();
    let table = editor
        .add_slide_table(0, "Fixed sections", 4, 4, position, size)
        .unwrap();
    let model_id = table.model_object_id;
    editor
        .set_slide_table_header_settings(
            0,
            model_id,
            HeaderSettings {
                header_rows: Some(HeaderCount::ONE),
                header_columns: Some(HeaderCount::ONE),
                footer_rows: Some(HeaderCount::ONE),
                ..Default::default()
            },
        )
        .unwrap();
    let baseline = editor.to_bytes().unwrap();

    editor
        .insert_slide_table_row(0, model_id, RowInsertion::header(1))
        .unwrap();
    editor
        .insert_slide_table_row(0, model_id, RowInsertion::footer(0))
        .unwrap();
    editor
        .insert_slide_table_column(0, model_id, ColumnInsertion::header(1))
        .unwrap();
    let settings = editor.slide_table_header_settings(0, model_id).unwrap();
    assert_eq!(settings.header_row_count(), 2);
    assert_eq!(settings.footer_row_count(), 2);
    assert_eq!(settings.header_column_count(), 2);
    let table = editor.slide_table(0, model_id).unwrap();
    assert_eq!((table.info.rows, table.info.columns), (6, 5));

    editor
        .remove_slide_table_column(0, model_id, ColumnDeletion::header(1))
        .unwrap();
    editor
        .remove_slide_table_row(0, model_id, RowDeletion::footer(0))
        .unwrap();
    editor
        .remove_slide_table_row(0, model_id, RowDeletion::header(1))
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    editor
        .set_slide_table_cell(
            0,
            model_id,
            0,
            0,
            KeynoteTableCellValue::Text("Header".to_owned()),
        )
        .unwrap();
    editor
        .set_slide_table_cell(
            0,
            model_id,
            1,
            1,
            KeynoteTableCellValue::Text("Body".to_owned()),
        )
        .unwrap();
    editor
        .set_slide_table_cell(
            0,
            model_id,
            3,
            2,
            KeynoteTableCellValue::Text("Footer".to_owned()),
        )
        .unwrap();
    editor
        .remove_slide_table_row(0, model_id, RowDeletion::header(0))
        .unwrap();
    editor
        .remove_slide_table_row(0, model_id, RowDeletion::footer(0))
        .unwrap();
    editor
        .remove_slide_table_column(0, model_id, ColumnDeletion::header(0))
        .unwrap();
    let settings = editor.slide_table_header_settings(0, model_id).unwrap();
    assert_eq!(settings.header_row_count(), 0);
    assert_eq!(settings.footer_row_count(), 0);
    assert_eq!(settings.header_column_count(), 0);
    let table = editor.slide_table(0, model_id).unwrap();
    assert_eq!((table.info.rows, table.info.columns), (2, 3));
    assert_eq!(
        table.get_cell(0, 0),
        Some(&KeynoteTableCellValue::Text("Body".to_owned()))
    );
    assert!(!table.iter_cells().any(|(_, value)| matches!(
        value,
        KeynoteTableCellValue::Text(text) if text == "Header" || text == "Footer"
    )));
}

#[test]
fn source_built_table_roundtrips_title_settings_transactionally() {
    let mut editor = KeynoteDocumentBuilder::new().build().unwrap();
    let (position, size) = table_geometry();
    let table = editor
        .add_slide_table(0, "Forecast", 2, 2, position, size)
        .unwrap();
    let visible = TableTitleSettings::new(Some(true), Some(true));
    let initially_hidden = TableTitleSettings::new(Some(false), None);
    assert_eq!(
        focused_table_title_settings(&editor, 0, table.model_object_id).unwrap(),
        initially_hidden
    );
    set_focused_table_title_settings(&mut editor, 0, table.model_object_id, visible).unwrap();

    let mut reopened = KeynoteEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        focused_table_title_settings(&reopened, 0, table.model_object_id).unwrap(),
        visible
    );
    let unchanged = reopened.to_bytes().unwrap();
    set_focused_table_title_settings(&mut reopened, 0, table.model_object_id, visible).unwrap();
    assert_eq!(reopened.to_bytes().unwrap(), unchanged);

    let explicit_hidden = TableTitleSettings::new(Some(false), Some(false));
    set_focused_table_title_settings(&mut reopened, 0, table.model_object_id, explicit_hidden)
        .unwrap();
    assert_eq!(
        focused_table_title_settings(&reopened, 0, table.model_object_id).unwrap(),
        explicit_hidden
    );
    set_focused_table_title_settings(
        &mut reopened,
        0,
        table.model_object_id,
        TableTitleSettings::default(),
    )
    .unwrap();
    assert_eq!(
        focused_table_title_settings(&reopened, 0, table.model_object_id).unwrap(),
        TableTitleSettings::default()
    );

    let before_error = reopened.to_bytes().unwrap();
    assert!(
        set_focused_table_title_settings(&mut reopened, 1, table.model_object_id, visible).is_err()
    );
    assert_eq!(reopened.to_bytes().unwrap(), before_error);
}

#[test]
fn tables_on_multiple_slides_remain_isolated() {
    let mut editor = KeynoteDocumentBuilder::new().build().unwrap();
    let layout = editor
        .slide_layouts()
        .unwrap()
        .into_iter()
        .find(|layout| layout.is_default)
        .unwrap();
    editor.add_slide(layout.id).unwrap();
    let (position, size) = table_geometry();
    let first = editor
        .add_slide_table(0, "First", 2, 2, position, size)
        .unwrap();
    let second = editor
        .add_slide_table(1, "Second", 2, 2, position, size)
        .unwrap();
    assert_ne!(first.drawable_object_id, second.drawable_object_id);
    assert_ne!(first.model_object_id, second.model_object_id);
    assert_eq!(editor.slide_tables(0).unwrap()[0].name, "First");
    assert_eq!(editor.slide_tables(1).unwrap()[0].name, "Second");
}

#[test]
fn source_built_tables_list_in_z_order_and_roundtrip_through_catalog() {
    let mut editor = KeynoteDocumentBuilder::new().build().unwrap();
    let (position, size) = table_geometry();
    editor
        .add_slide_table(0, "First", 2, 2, position, size)
        .unwrap();
    editor
        .add_slide_table(0, "Second", 3, 2, position, size)
        .unwrap();
    editor
        .add_slide_table(0, "Third", 4, 3, position, size)
        .unwrap();

    let expected = [("First", 2, 2), ("Second", 3, 2), ("Third", 4, 3)];
    let listed = editor.slide_tables(0).unwrap();
    assert_eq!(listed.len(), expected.len());
    for (table, (name, rows, columns)) in listed.iter().zip(expected) {
        assert_eq!(table.name, name);
        assert_eq!((table.rows, table.columns), (rows, columns));
    }

    // Listing is read-only and the compact catalog is rebuilt from the same
    // immutable package snapshot on reopen; no table identity or z-order is
    // inferred from the generated model's allocation order.
    let source = editor.to_bytes().unwrap();
    let reopened = KeynoteEditor::from_bytes(&source).unwrap();
    let reread = reopened.slide_tables(0).unwrap();
    assert_eq!(reread.len(), listed.len());
    for (before, after) in listed.iter().zip(reread) {
        assert_eq!(
            (
                before.name.as_str(),
                before.rows,
                before.columns,
                before.drawable_object_id,
                before.model_object_id,
            ),
            (
                after.name.as_str(),
                after.rows,
                after.columns,
                after.drawable_object_id,
                after.model_object_id,
            )
        );
    }
    assert_eq!(reopened.to_bytes().unwrap(), source);
}
