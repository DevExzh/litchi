//! Snapshot and transaction coverage for the Pages table facade.

use super::*;
use crate::archive::RawMessage;
use crate::numbers::CellValue;
use crate::pages::PagesDocumentBuilder;
use litchi_iwa_common::table::cell::conditional_highlight::{
    Condition, Rule, Style, Text,
};
use crate::table_cell_data_format::{
    TableCellCurrencyCode, TableCellCurrencyStyle, TableCellCustomFormatName,
    TableCellCustomTextFormat, TableCellFractionAccuracy, TableCellNumeralSystemBase,
    TableCellNumeralSystemFixedPlaces, TableCellNumeralSystemNegativeStyle,
    TableCellNumeralSystemPlaces,
};
use litchi_numbers::table::headers::{Count as HeaderCount, Settings as HeaderSettings};
use litchi_iwa_common::table::cell::number_format::{
    DecimalPlaces as NumberDecimalPlaces, NegativeStyle as NumberNegativeStyle,
    ThousandsSeparator as NumberThousandsSeparator,
};

const SOURCE_BUILT_TABLE_INFO_OBJECT_ID: u64 = 9;

#[test]
fn malformed_table_model_payload_is_reported() {
    let messages = [RawMessage {
        type_: TABLE_MODEL_MESSAGE_TYPES[0],
        data: vec![0x80],
    }];
    let error = decode_table_models(messages.iter(), 41).expect_err("malformed payload");
    assert!(matches!(
        error,
        Error::InvalidFormat(message)
            if message.contains("Pages table model 41")
                && message.contains("malformed table-model payload")
    ));
}

#[test]
fn source_built_table_has_no_conditional_highlighting_and_clear_is_idempotent() {
    let mut editor = PagesDocumentBuilder::new()
        .body_table("Conditional", 2, 2)
        .build()
        .unwrap();
    let model_id = editor.tables().unwrap()[0].model_object_id;

    assert!(
        editor
            .table_cell_conditional_highlighting(model_id, 1, 1)
            .unwrap()
            .is_none()
    );
    editor
        .clear_table_cell_conditional_highlighting(model_id, 1, 1)
        .unwrap();
    assert!(
        PagesEditor::from_bytes(&editor.to_bytes().unwrap())
            .unwrap()
            .table_cell_conditional_highlighting(model_id, 1, 1)
            .unwrap()
            .is_none()
    );
}

#[test]
fn source_built_table_creates_and_replaces_conditional_highlighting() {
    let mut editor = PagesDocumentBuilder::new()
        .body_table("Conditional", 2, 2)
        .build()
        .unwrap();
    let model_id = editor.tables().unwrap()[0].model_object_id;
    editor
        .set_table_cell(model_id, 1, 1, CellValue::Text("Organic Grain".to_owned()))
        .unwrap();
    let rule = Rule::new(
        Condition::TextContains(
            Text::new("grain").unwrap(),
        ),
        Style::with_fill(
            crate::shapes::RgbaColor::new(0.9, 0.1, 0.1, 1.0, crate::shapes::RgbColorSpace::Srgb)
                .unwrap(),
        ),
    );
    let created = editor
        .set_table_cell_conditional_highlighting(model_id, 1, 1, std::slice::from_ref(&rule))
        .unwrap();
    assert_eq!(created.table_id, model_id);
    assert_eq!((created.row, created.column), (1, 1));
    assert_eq!(created.rule_count, 1);
    assert_eq!(
        editor
            .table_cell_conditional_highlighting(model_id, 1, 1)
            .unwrap()
            .unwrap()
            .rule_count,
        1
    );
    assert_eq!(
        editor
            .table_cell_conditional_highlight_rules(model_id, 1, 1)
            .unwrap()
            .unwrap(),
        vec![rule]
    );
}

#[test]
fn source_built_table_roundtrips_cell_border_crud() {
    let mut editor = PagesDocumentBuilder::new()
        .body_table("Borders", 3, 3)
        .build()
        .unwrap();
    let model_id = editor.tables().unwrap()[0].model_object_id;
    let stroke = crate::shapes::ShapeStroke::new(
        crate::shapes::RgbaColor::new(0.1, 0.3, 0.9, 1.0, crate::shapes::RgbColorSpace::Srgb)
            .unwrap(),
        crate::shapes::StrokeWidth::new(2.0).unwrap(),
        crate::shapes::StrokePattern::RoundedDash,
    );
    editor
        .set_table_cell_border(model_id, 1, 1, PagesTableCellBorderSide::Right, stroke)
        .unwrap();

    let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened.table_cell_borders(model_id, 1, 1).unwrap().right,
        Some(stroke)
    );
    reopened
        .clear_table_cell_border(model_id, 1, 1, PagesTableCellBorderSide::Right)
        .unwrap();
    assert_eq!(
        reopened.table_cell_borders(model_id, 1, 1).unwrap().right,
        None
    );
}

#[test]
fn source_built_table_roundtrips_cell_fill_crud() {
    let mut editor = PagesDocumentBuilder::new()
        .body_table("Fills", 3, 3)
        .build()
        .unwrap();
    let model_id = editor.tables().unwrap()[0].model_object_id;
    let inherited = editor.table_cell_fill(model_id, 1, 1).unwrap();
    let fill = crate::shapes::ShapeFill::Solid(
        crate::shapes::RgbaColor::new(0.2, 0.75, 0.35, 1.0, crate::shapes::RgbColorSpace::Srgb)
            .unwrap(),
    );
    editor.set_table_cell_fill(model_id, 1, 1, &fill).unwrap();

    let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(reopened.table_cell_fill(model_id, 1, 1).unwrap(), fill);
    assert!(reopened.reset_table_cell_fill(model_id, 1, 1).unwrap());
    assert_eq!(reopened.table_cell_fill(model_id, 1, 1).unwrap(), inherited);
}

#[test]
fn source_built_table_roundtrips_cell_layout_crud() {
    let mut editor = PagesDocumentBuilder::new()
        .body_table("Layout", 3, 3)
        .build()
        .unwrap();
    let model_id = editor.tables().unwrap()[0].model_object_id;
    let inherited = editor.table_cell_layout(model_id, 1, 1).unwrap();
    let layout = PagesTableCellLayout::default()
        .with_text_wrap(PagesTableCellTextWrap::Wrapped)
        .with_vertical_alignment(PagesTableCellVerticalAlignment::Bottom)
        .with_insets(PagesTableCellInsets::uniform(
            PagesTableCellInset::from_points(5.0).unwrap(),
        ));
    editor
        .set_table_cell_layout(model_id, 1, 1, layout)
        .unwrap();

    let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(reopened.table_cell_layout(model_id, 1, 1).unwrap(), layout);
    assert!(reopened.reset_table_cell_layout(model_id, 1, 1).unwrap());
    assert_eq!(
        reopened.table_cell_layout(model_id, 1, 1).unwrap(),
        inherited
    );
}

#[test]
fn source_built_table_roundtrips_number_format_crud() {
    let mut editor = PagesDocumentBuilder::new()
        .body_table("Formats", 3, 3)
        .build()
        .unwrap();
    let model_id = editor.tables().unwrap()[0].model_object_id;
    let format = PagesTableCellNumberFormat::new(
        NumberDecimalPlaces::fixed(2).unwrap(),
        NumberNegativeStyle::Parentheses,
        NumberThousandsSeparator::Shown,
    );
    editor
        .set_table_cell(
            model_id,
            1,
            1,
            PagesCellValue::number(1_234.5).expect("finite test number"),
        )
        .unwrap();
    editor
        .set_table_cell_number_format(model_id, 1, 1, format)
        .unwrap();

    let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened.table_cell_number_format(model_id, 1, 1).unwrap(),
        Some(format)
    );
    assert!(
        reopened
            .reset_table_cell_number_format(model_id, 1, 1)
            .unwrap()
    );
    assert_eq!(
        reopened.table_cell_number_format(model_id, 1, 1).unwrap(),
        None
    );
}

#[test]
fn source_built_table_roundtrips_percentage_data_format() {
    let mut editor = PagesDocumentBuilder::new()
        .body_table("Percentages", 3, 3)
        .build()
        .unwrap();
    let model_id = editor.tables().unwrap()[0].model_object_id;
    let format = PagesTableCellPercentageFormat::new(
        PagesTableCellDecimalPlaces::fixed(2).unwrap(),
        PagesTableCellNegativeNumberStyle::Parentheses,
        PagesTableCellThousandsSeparator::Shown,
    );
    editor
        .set_table_cell(
            model_id,
            1,
            1,
            PagesCellValue::number(-12.345).expect("finite test number"),
        )
        .unwrap();
    editor
        .set_table_cell_percentage_format(model_id, 1, 1, format)
        .unwrap();

    let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened.table_cell_data_format(model_id, 1, 1).unwrap(),
        PagesTableCellDataFormat::Percentage(format)
    );
    assert_eq!(
        reopened
            .table_cell_percentage_format(model_id, 1, 1)
            .unwrap(),
        Some(format)
    );
    reopened
        .set_table_cell_data_format(model_id, 1, 1, PagesTableCellDataFormat::Automatic)
        .unwrap();
    assert_eq!(
        reopened.table_cell_data_format(model_id, 1, 1).unwrap(),
        PagesTableCellDataFormat::Automatic
    );
}

#[test]
fn source_built_table_roundtrips_currency_format_crud() {
    let mut editor = PagesDocumentBuilder::new()
        .body_table("Currencies", 3, 3)
        .build()
        .unwrap();
    let model_id = editor.tables().unwrap()[0].model_object_id;
    let format = PagesTableCellCurrencyFormat::new(
        TableCellCurrencyCode::EUR,
        PagesTableCellDecimalPlaces::fixed(2).unwrap(),
        PagesTableCellNegativeNumberStyle::Parentheses,
        PagesTableCellThousandsSeparator::Shown,
        TableCellCurrencyStyle::Accounting,
    );
    editor
        .set_table_cell(
            model_id,
            1,
            1,
            PagesCellValue::number(-1_234.5).expect("finite test number"),
        )
        .unwrap();
    editor
        .set_table_cell_currency_format(model_id, 1, 1, format)
        .unwrap();

    let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened.table_cell_currency_format(model_id, 1, 1).unwrap(),
        Some(format)
    );
    assert!(
        reopened
            .reset_table_cell_currency_format(model_id, 1, 1)
            .unwrap()
    );
    assert_eq!(
        reopened.table_cell_data_format(model_id, 1, 1).unwrap(),
        PagesTableCellDataFormat::Automatic
    );
}

#[test]
fn source_built_table_roundtrips_scientific_format_crud() {
    let mut editor = PagesDocumentBuilder::new()
        .body_table("Scientific", 3, 3)
        .build()
        .unwrap();
    let model_id = editor.tables().unwrap()[0].model_object_id;
    let format =
        PagesTableCellScientificFormat::new(PagesTableCellFixedDecimalPlaces::new(5).unwrap());
    editor
        .set_table_cell(
            model_id,
            1,
            1,
            PagesCellValue::number(-1_234.5).expect("finite test number"),
        )
        .unwrap();
    editor
        .set_table_cell_scientific_format(model_id, 1, 1, format)
        .unwrap();

    let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .table_cell_scientific_format(model_id, 1, 1)
            .unwrap(),
        Some(format)
    );
    assert!(
        reopened
            .reset_table_cell_scientific_format(model_id, 1, 1)
            .unwrap()
    );
    assert_eq!(
        reopened.table_cell_data_format(model_id, 1, 1).unwrap(),
        PagesTableCellDataFormat::Automatic
    );
}

#[test]
fn source_built_table_roundtrips_fraction_format_crud() {
    let mut editor = PagesDocumentBuilder::new()
        .body_table("Fractions", 3, 3)
        .build()
        .unwrap();
    let model_id = editor.tables().unwrap()[0].model_object_id;
    let format = PagesTableCellFractionFormat::new(TableCellFractionAccuracy::Eighths);
    editor
        .set_table_cell(
            model_id,
            1,
            1,
            PagesCellValue::number(-12.375).expect("finite test number"),
        )
        .unwrap();
    editor
        .set_table_cell_fraction_format(model_id, 1, 1, format)
        .unwrap();

    let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened.table_cell_fraction_format(model_id, 1, 1).unwrap(),
        Some(format)
    );
    assert!(
        reopened
            .reset_table_cell_fraction_format(model_id, 1, 1)
            .unwrap()
    );
    assert_eq!(
        reopened.table_cell_data_format(model_id, 1, 1).unwrap(),
        PagesTableCellDataFormat::Automatic
    );
}

#[test]
fn source_built_table_roundtrips_numeral_system_format_crud() {
    let mut editor = PagesDocumentBuilder::new()
        .body_table("Numeral Systems", 3, 3)
        .build()
        .unwrap();
    let model_id = editor.tables().unwrap()[0].model_object_id;
    let format = PagesTableCellNumeralSystemFormat::new(
        TableCellNumeralSystemBase::HEXADECIMAL,
        TableCellNumeralSystemPlaces::Fixed(TableCellNumeralSystemFixedPlaces::EIGHT),
        TableCellNumeralSystemNegativeStyle::TwosComplement,
    )
    .unwrap();
    editor
        .set_table_cell(
            model_id,
            1,
            1,
            PagesCellValue::number(-1_234.5).expect("finite test number"),
        )
        .unwrap();
    editor
        .set_table_cell_numeral_system_format(model_id, 1, 1, format)
        .unwrap();

    let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .table_cell_numeral_system_format(model_id, 1, 1)
            .unwrap(),
        Some(format)
    );
    assert!(
        reopened
            .reset_table_cell_numeral_system_format(model_id, 1, 1)
            .unwrap()
    );
    assert_eq!(
        reopened.table_cell_data_format(model_id, 1, 1).unwrap(),
        PagesTableCellDataFormat::Automatic
    );
}

#[test]
fn source_built_table_roundtrips_date_time_format_crud() {
    let mut editor = PagesDocumentBuilder::new()
        .body_table("Dates", 3, 3)
        .build()
        .unwrap();
    let model_id = editor.tables().unwrap()[0].model_object_id;
    let format = PagesTableCellDateTimeFormat::iso_date_time_24_hour_with_seconds();
    editor
        .set_table_cell(
            model_id,
            1,
            1,
            PagesCellValue::date(789_332_889.0).expect("finite test date"),
        )
        .unwrap();
    editor
        .set_table_cell_date_time_format(model_id, 1, 1, format.clone())
        .unwrap();

    let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .table_cell_date_time_format(model_id, 1, 1)
            .unwrap(),
        Some(format)
    );
    assert!(
        reopened
            .reset_table_cell_date_time_format(model_id, 1, 1)
            .unwrap()
    );
    assert_eq!(
        reopened.table_cell_data_format(model_id, 1, 1).unwrap(),
        PagesTableCellDataFormat::Automatic
    );
}

#[test]
fn source_built_table_roundtrips_duration_format_crud() {
    let mut editor = PagesDocumentBuilder::new()
        .body_table("Durations", 3, 3)
        .build()
        .unwrap();
    let model_id = editor.tables().unwrap()[0].model_object_id;
    let range = PagesTableCellDurationUnitRange::hours_to_milliseconds();
    let format =
        PagesTableCellDurationFormat::custom(PagesTableCellDurationStyle::Abbreviated, range);
    editor
        .set_table_cell(
            model_id,
            1,
            1,
            PagesCellValue::duration(3_723.5).expect("finite test duration"),
        )
        .unwrap();
    editor
        .set_table_cell_duration_format(model_id, 1, 1, format)
        .unwrap();

    let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened.table_cell_duration_format(model_id, 1, 1).unwrap(),
        Some(format)
    );
    assert!(
        reopened
            .reset_table_cell_duration_format(model_id, 1, 1)
            .unwrap()
    );
    assert_eq!(
        reopened.table_cell_data_format(model_id, 1, 1).unwrap(),
        PagesTableCellDataFormat::Automatic
    );
}

#[test]
fn source_built_table_roundtrips_checkbox_format_crud() {
    let mut editor = PagesDocumentBuilder::new()
        .body_table("Checkboxes", 3, 3)
        .build()
        .unwrap();
    let model_id = editor.tables().unwrap()[0].model_object_id;
    editor
        .set_table_cell_checkbox_format(model_id, 1, 1, PagesTableCellCheckboxFormat)
        .unwrap();

    let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened.table_cell_checkbox_format(model_id, 1, 1).unwrap(),
        Some(PagesTableCellCheckboxFormat)
    );
    assert_eq!(
        reopened.table(model_id).unwrap().get_cell(1, 1),
        Some(&PagesCellValue::Boolean(false))
    );
    assert!(
        reopened
            .reset_table_cell_checkbox_format(model_id, 1, 1)
            .unwrap()
    );
    assert_eq!(
        reopened.table_cell_data_format(model_id, 1, 1).unwrap(),
        PagesTableCellDataFormat::Automatic
    );
}

#[test]
fn source_built_table_roundtrips_star_rating_format_crud() {
    let mut editor = PagesDocumentBuilder::new()
        .body_table("Ratings", 3, 3)
        .build()
        .unwrap();
    let model_id = editor.tables().unwrap()[0].model_object_id;
    editor
        .set_table_cell_star_rating_format(model_id, 1, 1, PagesTableCellStarRatingFormat)
        .unwrap();

    let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .table_cell_star_rating_format(model_id, 1, 1)
            .unwrap(),
        Some(PagesTableCellStarRatingFormat)
    );
    assert_eq!(
        reopened.table(model_id).unwrap().get_cell(1, 1),
        Some(&PagesCellValue::number(0.0).expect("finite test number"))
    );
    assert!(
        reopened
            .reset_table_cell_star_rating_format(model_id, 1, 1)
            .unwrap()
    );
    assert_eq!(
        reopened.table_cell_data_format(model_id, 1, 1).unwrap(),
        PagesTableCellDataFormat::Automatic
    );
}

#[test]
fn source_built_table_roundtrips_slider_format_crud() {
    let mut editor = PagesDocumentBuilder::new()
        .body_table("Sliders", 3, 3)
        .build()
        .unwrap();
    let model_id = editor.tables().unwrap()[0].model_object_id;
    let range = PagesTableCellSliderRange::new(-10.0, 30.0, 0.5).unwrap();
    let format = PagesTableCellSliderFormat::new(
        range,
        crate::table_cell_data_format::TableCellNumberFormat::default().into(),
    );
    editor
        .set_table_cell_slider_format(model_id, 1, 1, format.clone())
        .unwrap();

    let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened.table_cell_slider_format(model_id, 1, 1).unwrap(),
        Some(format)
    );
    assert_eq!(
        reopened.table(model_id).unwrap().get_cell(1, 1),
        Some(&PagesCellValue::number(10.0).expect("finite test number"))
    );
    assert!(
        reopened
            .reset_table_cell_slider_format(model_id, 1, 1)
            .unwrap()
    );
    assert_eq!(
        reopened.table_cell_data_format(model_id, 1, 1).unwrap(),
        PagesTableCellDataFormat::Automatic
    );
}

#[test]
fn source_built_table_roundtrips_stepper_format_crud() {
    let mut editor = PagesDocumentBuilder::new()
        .body_table("Steppers", 3, 3)
        .build()
        .unwrap();
    let model_id = editor.tables().unwrap()[0].model_object_id;
    let range = PagesTableCellStepperRange::new(-10.0, 30.0, 0.5).unwrap();
    let format = PagesTableCellStepperFormat::new(
        range,
        crate::table_cell_data_format::TableCellNumberFormat::default().into(),
    );
    editor
        .set_table_cell_stepper_format(model_id, 1, 1, format.clone())
        .unwrap();

    let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened.table_cell_stepper_format(model_id, 1, 1).unwrap(),
        Some(format)
    );
    assert_eq!(
        reopened.table(model_id).unwrap().get_cell(1, 1),
        Some(&PagesCellValue::number(-10.0).expect("finite test number"))
    );
    assert!(
        reopened
            .reset_table_cell_stepper_format(model_id, 1, 1)
            .unwrap()
    );
    assert_eq!(
        reopened.table_cell_data_format(model_id, 1, 1).unwrap(),
        PagesTableCellDataFormat::Automatic
    );
}

#[test]
fn source_built_table_roundtrips_text_format_crud() {
    let mut editor = PagesDocumentBuilder::new()
        .body_table("Text", 3, 3)
        .build()
        .unwrap();
    let model_id = editor.tables().unwrap()[0].model_object_id;
    editor
        .set_table_cell(model_id, 1, 1, PagesCellValue::Text("00123".to_owned()))
        .unwrap();
    editor.set_table_cell_text_format(model_id, 1, 1).unwrap();

    let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened.table_cell_text_format(model_id, 1, 1).unwrap(),
        Some(PagesTableCellTextFormat)
    );
    assert_eq!(
        reopened.table(model_id).unwrap().get_cell(1, 1),
        Some(&PagesCellValue::Text("00123".to_owned()))
    );
    assert!(
        reopened
            .reset_table_cell_text_format(model_id, 1, 1)
            .unwrap()
    );
    assert_eq!(
        reopened.table_cell_data_format(model_id, 1, 1).unwrap(),
        PagesTableCellDataFormat::Automatic
    );
}

#[test]
fn source_built_table_roundtrips_custom_format_crud() {
    let mut editor = PagesDocumentBuilder::new()
        .body_table("Custom Text", 3, 3)
        .build()
        .unwrap();
    let model_id = editor.tables().unwrap()[0].model_object_id;
    let format = PagesTableCellCustomFormat::Text(
        TableCellCustomTextFormat::try_new(
            TableCellCustomFormatName::try_new("Invoice Identifier").unwrap(),
            "",
            " ID",
        )
        .unwrap(),
    );
    editor
        .set_table_cell(model_id, 1, 1, PagesCellValue::Text("00123".to_owned()))
        .unwrap();
    editor
        .set_table_cell_custom_format(model_id, 1, 1, format.clone())
        .unwrap();

    let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened.table_cell_custom_format(model_id, 1, 1).unwrap(),
        Some(format)
    );
    assert!(
        reopened
            .reset_table_cell_custom_format(model_id, 1, 1)
            .unwrap()
    );
    assert_eq!(
        reopened.table_cell_data_format(model_id, 1, 1).unwrap(),
        PagesTableCellDataFormat::Automatic
    );
}

#[test]
fn source_built_table_roundtrips_pop_up_menu_format_crud() {
    let mut editor = PagesDocumentBuilder::new()
        .body_table("Menus", 3, 3)
        .build()
        .unwrap();
    let model_id = editor.tables().unwrap()[0].model_object_id;
    let format = PagesTableCellPopUpMenuFormat::try_new(["Draft", "Published"])
        .unwrap()
        .with_initial_selection(PagesTableCellPopUpMenuInitialSelection::Blank);
    editor
        .set_table_cell_pop_up_menu_format(model_id, 1, 1, format.clone())
        .unwrap();

    let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .table_cell_pop_up_menu_format(model_id, 1, 1)
            .unwrap(),
        Some(format)
    );
    assert_eq!(
        reopened.table(model_id).unwrap().get_cell(1, 1),
        Some(&PagesCellValue::Empty)
    );
    assert!(
        reopened
            .reset_table_cell_pop_up_menu_format(model_id, 1, 1)
            .unwrap()
    );
    assert_eq!(
        reopened.table_cell_data_format(model_id, 1, 1).unwrap(),
        PagesTableCellDataFormat::Automatic
    );
}

#[test]
fn source_built_table_roundtrips_cell_updates() {
    let mut editor = PagesDocumentBuilder::new()
        .body_text("Report 🙂\n")
        .body_table("Results", 3, 2)
        .build()
        .unwrap();
    let tables = editor.tables().unwrap();
    assert_eq!(tables.len(), 1);
    let info = &tables[0];
    assert_eq!(info.name, "Results");
    assert_eq!((info.rows, info.columns), (3, 2));
    assert_eq!(
        info.anchor_character_index,
        "Report 🙂\n".encode_utf16().count()
    );
    let model_id = info.model_object_id;
    assert_eq!(editor.table(model_id).unwrap().cell_count(), 0);

    assert_eq!(
        editor
            .set_table_cells(
                model_id,
                [
                    PagesTableCellUpdate::new(0, 0, PagesCellValue::Text("Header".to_owned()),),
                    PagesTableCellUpdate::new(
                        1,
                        1,
                        PagesCellValue::number(42.5).expect("finite test number"),
                    ),
                ],
            )
            .unwrap(),
        2
    );
    let before_invalid_batch = editor.to_bytes().unwrap();
    assert!(
        editor
            .set_table_cells(
                model_id,
                [
                    PagesTableCellUpdate::new(2, 0, PagesCellValue::Boolean(true)),
                    PagesTableCellUpdate::clear(2, 0),
                ],
            )
            .is_err()
    );
    assert_eq!(editor.to_bytes().unwrap(), before_invalid_batch);
    let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    let table = reopened.table(model_id).unwrap();
    assert_eq!(
        table.get_cell(0, 0),
        Some(&PagesCellValue::Text("Header".to_owned()))
    );
    assert_eq!(
        table.get_cell(1, 1),
        Some(&PagesCellValue::number(42.5).expect("finite test number"))
    );
    reopened.clear_table_cell(model_id, 0, 0).unwrap();
    assert!(reopened.table(model_id).unwrap().get_cell(0, 0).is_none());
}

#[test]
fn source_built_table_duplication_clones_formula_storage_and_attachment() {
    let mut editor = PagesDocumentBuilder::new()
        .body_text("Budget\n")
        .body_table("Budget", 3, 2)
        .build()
        .unwrap();
    let source = editor.tables().unwrap().remove(0);
    editor
        .set_table_cells(
            source.model_object_id,
            [
                PagesTableCellUpdate::new(0, 0, PagesCellValue::Text("Category".to_owned())),
                PagesTableCellUpdate::new(1, 0, PagesCellValue::Text("Travel".to_owned())),
                PagesTableCellUpdate::new(
                    1,
                    1,
                    PagesCellValue::number(125.0).expect("finite test number"),
                ),
            ],
        )
        .unwrap();
    editor
        .set_table_formula(
            source.model_object_id,
            2,
            1,
            PagesTableFormulaExpression::function(
                "SUM",
                [
                    PagesTableFormulaExpression::Number(100.0),
                    PagesTableFormulaExpression::Number(25.0),
                ],
            ),
            PagesTableFormulaCachedValue::number(125.0)
                .expect("finite cached formula number"),
        )
        .unwrap();

    let insertion_anchor = editor.body_text().unwrap().encode_utf16().count();
    let copied = editor
        .duplicate_table(source.model_object_id, insertion_anchor)
        .unwrap();
    assert_ne!(copied.drawable_object_id, source.drawable_object_id);
    assert_ne!(copied.model_object_id, source.model_object_id);
    assert_eq!(copied.name, "Budget copy");
    assert_eq!(copied.anchor_character_index, insertion_anchor);
    assert_eq!((copied.rows, copied.columns), (source.rows, source.columns));
    assert_eq!(
        editor.table(copied.model_object_id).unwrap().get_cell(1, 0),
        Some(&PagesCellValue::Text("Travel".to_owned()))
    );
    assert_eq!(
        editor
            .table_formula(copied.model_object_id, 2, 1)
            .unwrap()
            .as_deref(),
        Some("=SUM(100,25)")
    );

    editor
        .set_table_cell(
            copied.model_object_id,
            1,
            0,
            PagesCellValue::Text("Lodging".to_owned()),
        )
        .unwrap();
    assert_eq!(
        editor.table(source.model_object_id).unwrap().get_cell(1, 0),
        Some(&PagesCellValue::Text("Travel".to_owned()))
    );

    let second_anchor = editor.body_text().unwrap().encode_utf16().count();
    assert_eq!(
        editor
            .duplicate_table(source.model_object_id, second_anchor)
            .unwrap()
            .name,
        "Budget copy 2"
    );
    let before_invalid = editor.to_bytes().unwrap();
    assert!(
        editor
            .duplicate_table(source.model_object_id, usize::MAX)
            .is_err()
    );
    assert_eq!(editor.to_bytes().unwrap(), before_invalid);

    let reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(reopened.tables().unwrap().len(), 3);
    let mut reopened = reopened;
    reopened.remove_table(copied.model_object_id).unwrap();
    assert_eq!(reopened.tables().unwrap().len(), 2);
    assert_eq!(
        reopened
            .table(source.model_object_id)
            .unwrap()
            .get_cell(1, 0),
        Some(&PagesCellValue::Text("Travel".to_owned()))
    );
}

#[test]
fn source_built_table_roundtrips_full_table_sort_crud() {
    let mut editor = PagesDocumentBuilder::new()
        .body_table("Cities", 5, 2)
        .build()
        .unwrap();
    let model_id = editor.tables().unwrap()[0].model_object_id;
    editor
        .set_table_header_settings(
            model_id,
            HeaderSettings {
                header_rows: Some(HeaderCount::ONE),
                ..Default::default()
            },
        )
        .unwrap();
    editor
        .set_table_cells(
            model_id,
            [
                PagesTableCellUpdate::new(0, 0, PagesCellValue::Text("Name".to_owned())),
                PagesTableCellUpdate::new(0, 1, PagesCellValue::Text("Marker".to_owned())),
                PagesTableCellUpdate::new(1, 0, PagesCellValue::Text("zebra".to_owned())),
                PagesTableCellUpdate::new(1, 1, PagesCellValue::Text("last".to_owned())),
                PagesTableCellUpdate::new(2, 0, PagesCellValue::Text("apple".to_owned())),
                PagesTableCellUpdate::new(2, 1, PagesCellValue::Text("first apple".to_owned())),
                PagesTableCellUpdate::new(3, 0, PagesCellValue::Text("banana".to_owned())),
                PagesTableCellUpdate::new(3, 1, PagesCellValue::Text("middle".to_owned())),
                PagesTableCellUpdate::new(4, 0, PagesCellValue::Text("apple".to_owned())),
                PagesTableCellUpdate::new(4, 1, PagesCellValue::Text("second apple".to_owned())),
            ],
        )
        .unwrap();
    editor
        .set_table_cell_comment(model_id, 1, 1, "Zebra comment follows row")
        .unwrap();
    let reply_id = editor
        .add_table_cell_comment_reply(model_id, 1, 1, "Zebra reply follows row")
        .unwrap();
    let comment_id = editor
        .table_cell_comment(model_id, 1, 1)
        .unwrap()
        .unwrap()
        .storage_object_id;
    let hidden = PagesTableHiddenAxes::new([PagesTableAxisIndex::row(2)]).unwrap();
    editor.set_table_hidden_axes(model_id, &hidden).unwrap();
    let order = PagesTableSortOrder::new([PagesTableSortRule::new(
        PagesTableSortColumnIndex::new(0).unwrap(),
        PagesTableSortDirection::Ascending,
    )])
    .unwrap();

    assert_eq!(editor.table_sort_order(model_id).unwrap(), None);
    editor
        .set_table_sort_order(model_id, order.clone())
        .unwrap();
    assert_eq!(
        editor.table_sort_order(model_id).unwrap(),
        Some(order.clone())
    );
    assert!(editor.apply_table_sort_order(model_id).unwrap());
    assert!(!editor.apply_table_sort_order(model_id).unwrap());
    let table = editor.table(model_id).unwrap();
    assert_eq!(
        table.get_cell(0, 0),
        Some(&PagesCellValue::Text("Name".to_owned()))
    );
    assert_eq!(
        table.get_cell(1, 0),
        Some(&PagesCellValue::Text("apple".to_owned()))
    );
    assert_eq!(
        table.get_cell(1, 1),
        Some(&PagesCellValue::Text("first apple".to_owned()))
    );
    assert_eq!(
        table.get_cell(2, 1),
        Some(&PagesCellValue::Text("second apple".to_owned()))
    );
    assert_eq!(
        table.get_cell(3, 0),
        Some(&PagesCellValue::Text("banana".to_owned()))
    );
    assert_eq!(
        table.get_cell(4, 0),
        Some(&PagesCellValue::Text("zebra".to_owned()))
    );
    assert!(editor.table_cell_comment(model_id, 1, 1).unwrap().is_none());
    assert_eq!(
        editor
            .table_cell_comment(model_id, 4, 1)
            .unwrap()
            .unwrap()
            .storage_object_id,
        comment_id
    );
    assert_eq!(
        editor.table_cell_comment_replies(model_id, 4, 1).unwrap()[0].storage_object_id,
        reply_id
    );
    assert_eq!(editor.table_hidden_axes(model_id).unwrap(), hidden);

    let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened.table_sort_order(model_id).unwrap(),
        Some(order.clone())
    );
    assert_eq!(
        reopened
            .table_cell_comment(model_id, 4, 1)
            .unwrap()
            .unwrap()
            .storage_object_id,
        comment_id
    );
    assert_eq!(
        reopened.table_cell_comment_replies(model_id, 4, 1).unwrap()[0].storage_object_id,
        reply_id
    );
    assert_eq!(reopened.table_hidden_axes(model_id).unwrap(), hidden);
    let unchanged = reopened.to_bytes().unwrap();
    reopened.set_table_sort_order(model_id, order).unwrap();
    assert_eq!(reopened.to_bytes().unwrap(), unchanged);

    let invalid = PagesTableSortOrder::new([PagesTableSortRule::new(
        PagesTableSortColumnIndex::new(2).unwrap(),
        PagesTableSortDirection::Ascending,
    )])
    .unwrap();
    let before_invalid = reopened.to_bytes().unwrap();
    assert!(reopened.set_table_sort_order(model_id, invalid).is_err());
    assert_eq!(reopened.to_bytes().unwrap(), before_invalid);

    let selected_order = PagesTableSortOrder::selected_rows([PagesTableSortRule::new(
        PagesTableSortColumnIndex::new(0).unwrap(),
        PagesTableSortDirection::Descending,
    )])
    .unwrap();
    reopened
        .set_table_sort_order(model_id, selected_order.clone())
        .unwrap();
    assert_eq!(
        reopened.table_sort_order(model_id).unwrap(),
        Some(selected_order)
    );
    let before_wrong_executor = reopened.to_bytes().unwrap();
    assert!(reopened.apply_table_sort_order(model_id).is_err());
    assert_eq!(reopened.to_bytes().unwrap(), before_wrong_executor);
    assert!(
        reopened
            .apply_table_sort_order_to_rows(model_id, PagesTableSortRowRange::new(1, 4).unwrap(),)
            .unwrap()
    );
    let table = reopened.table(model_id).unwrap();
    assert_eq!(
        table.get_cell(1, 1),
        Some(&PagesCellValue::Text("first apple".to_owned()))
    );
    assert_eq!(
        table.get_cell(2, 0),
        Some(&PagesCellValue::Text("zebra".to_owned()))
    );
    assert_eq!(
        table.get_cell(3, 0),
        Some(&PagesCellValue::Text("banana".to_owned()))
    );
    assert_eq!(
        table.get_cell(4, 1),
        Some(&PagesCellValue::Text("second apple".to_owned()))
    );
    assert!(
        reopened
            .table_cell_comment(model_id, 4, 1)
            .unwrap()
            .is_none()
    );
    assert_eq!(
        reopened
            .table_cell_comment(model_id, 2, 1)
            .unwrap()
            .unwrap()
            .storage_object_id,
        comment_id
    );
    assert_eq!(
        reopened.table_cell_comment_replies(model_id, 2, 1).unwrap()[0].storage_object_id,
        reply_id
    );
    assert_eq!(reopened.table_hidden_axes(model_id).unwrap(), hidden);

    reopened.clear_table_sort_order(model_id).unwrap();
    assert_eq!(reopened.table_sort_order(model_id).unwrap(), None);
    let unchanged = reopened.to_bytes().unwrap();
    reopened.clear_table_sort_order(model_id).unwrap();
    assert_eq!(reopened.to_bytes().unwrap(), unchanged);
    assert!(reopened.apply_table_sort_order(model_id).is_err());
    assert_eq!(reopened.to_bytes().unwrap(), unchanged);
}

#[test]
fn source_built_table_roundtrips_formula_crud_transactionally() {
    let editor = PagesDocumentBuilder::new()
        .body_table("Formula", 3, 2)
        .build()
        .unwrap();
    let model_id = editor.tables().unwrap()[0].model_object_id;
    let mut package = editor.into_package();
    let engine = package.remove_entry("Index/CalculationEngine.iwa").unwrap();
    package
        .insert_entry("Index/CalculationEngine-82.iwa", engine)
        .unwrap();
    let mut editor = PagesEditor::from_package(package).unwrap();
    editor
        .set_table_formula(
            model_id,
            1,
            1,
            PagesTableFormulaExpression::function(
                "SUM",
                [
                    PagesTableFormulaExpression::Number(1.0),
                    PagesTableFormulaExpression::Number(2.0),
                ],
            ),
            PagesTableFormulaCachedValue::number(3.0)
                .expect("finite cached formula number"),
        )
        .unwrap();

    let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened.table_formula(model_id, 1, 1).unwrap().as_deref(),
        Some("=SUM(1,2)")
    );
    reopened
        .set_table_formula(
            model_id,
            1,
            1,
            PagesTableFormulaExpression::function(
                "SUM",
                [
                    PagesTableFormulaExpression::Number(3.0),
                    PagesTableFormulaExpression::Number(4.0),
                ],
            ),
            PagesTableFormulaCachedValue::number(7.0)
                .expect("finite cached formula number"),
        )
        .unwrap();
    assert_eq!(
        reopened.table_formula(model_id, 1, 1).unwrap().as_deref(),
        Some("=SUM(3,4)")
    );

    let before = reopened.to_bytes().unwrap();
    assert!(
        reopened
            .set_table_formula(
                model_id,
                usize::MAX,
                1,
                PagesTableFormulaExpression::Number(1.0),
                PagesTableFormulaCachedValue::number(1.0)
                    .expect("finite cached formula number"),
            )
            .is_err()
    );
    assert_eq!(reopened.to_bytes().unwrap(), before);
    assert_eq!(
        reopened.clear_table_formula(model_id, 1, 1).unwrap(),
        "=SUM(3,4)"
    );
    assert_eq!(reopened.table_formula(model_id, 1, 1).unwrap(), None);
    let cleared = reopened.to_bytes().unwrap();
    assert!(reopened.clear_table_formula(model_id, 1, 1).is_err());
    assert_eq!(reopened.to_bytes().unwrap(), cleared);
}

#[test]
fn source_built_table_roundtrips_section_relative_axis_crud_transactionally() {
    let mut editor = PagesDocumentBuilder::new()
        .body_table("Topology", 4, 4)
        .build()
        .unwrap();
    let model_id = editor.tables().unwrap()[0].model_object_id;
    let row_size = PagesTableDimensionSize::points(33.0).unwrap();
    let column_size = PagesTableDimensionSize::points(77.0).unwrap();
    editor
        .set_table_cell(model_id, 1, 1, PagesCellValue::Text("shift me".to_owned()))
        .unwrap();
    editor
        .set_table_formula(
            model_id,
            2,
            2,
            PagesTableFormulaExpression::cell(PagesTableFormulaCellReference::relative(1, 1)),
                PagesTableFormulaCachedValue::number(7.0)
                    .expect("finite cached formula number"),
        )
        .unwrap();
    editor.set_table_row_height(model_id, 1, row_size).unwrap();
    editor
        .set_table_column_width(model_id, 1, column_size)
        .unwrap();
    let baseline = editor.to_bytes().unwrap();

    editor
        .insert_table_row(model_id, PagesTableRowInsertion::body(1))
        .unwrap();
    editor
        .insert_table_column(model_id, PagesTableColumnInsertion::body(1))
        .unwrap();
    let reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    let table = reopened.table(model_id).unwrap();
    assert_eq!((table.info.rows, table.info.columns), (5, 5));
    assert_eq!(
        table.get_cell(1, 1),
        Some(&PagesCellValue::Text("shift me".to_owned()))
    );
    assert_eq!(
        table.get_cell(3, 3),
        Some(&PagesCellValue::Formula("=B2".to_owned()))
    );
    assert_eq!(reopened.table_row_height(model_id, 1).unwrap(), row_size);
    assert_eq!(
        reopened.table_column_width(model_id, 1).unwrap(),
        column_size
    );

    editor
        .remove_table_column(model_id, PagesTableColumnDeletion::body(1))
        .unwrap();
    editor
        .remove_table_row(model_id, PagesTableRowDeletion::body(1))
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    let before_error = editor.to_bytes().unwrap();
    assert!(
        editor
            .insert_table_row(model_id, PagesTableRowInsertion::body(usize::MAX))
            .is_err()
    );
    assert!(
        editor
            .remove_table_column(model_id, PagesTableColumnDeletion::body(usize::MAX))
            .is_err()
    );
    assert_eq!(editor.to_bytes().unwrap(), before_error);
}

#[test]
fn source_built_footer_formula_expands_and_contracts_with_body_rows() {
    let mut editor = PagesDocumentBuilder::new()
        .body_table("Footer aggregate", 4, 3)
        .build()
        .unwrap();
    let model_id = editor.tables().unwrap()[0].model_object_id;
    editor
        .set_table_header_settings(
            model_id,
            HeaderSettings {
                footer_rows: Some(HeaderCount::ONE),
                ..Default::default()
            },
        )
        .unwrap();
    editor
        .set_table_formula(
            model_id,
            3,
            1,
            PagesTableFormulaExpression::function(
                "SUM",
                [PagesTableFormulaExpression::range(
                    PagesTableFormulaCellReference::relative(1, 1),
                    PagesTableFormulaCellReference::relative(2, 1),
                )],
            ),
            PagesTableFormulaCachedValue::number(3.0)
                .expect("finite cached formula number"),
        )
        .unwrap();
    let mut package = editor.into_package();
    let engine = package.remove_entry("Index/CalculationEngine.iwa").unwrap();
    package
        .insert_entry("Index/CalculationEngine-42-2.iwa", engine)
        .unwrap();
    let mut editor = PagesEditor::from_package(package).unwrap();
    let baseline = editor.to_bytes().unwrap();

    editor
        .insert_table_row(model_id, PagesTableRowInsertion::body(3))
        .unwrap();
    assert_eq!(
        editor.table_formula(model_id, 4, 1).unwrap().as_deref(),
        Some("=SUM(B2:B4)")
    );
    editor
        .remove_table_row(model_id, PagesTableRowDeletion::body(3))
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);
}

#[test]
fn source_built_fixed_table_sections_roundtrip_full_axis_crud() {
    let mut editor = PagesDocumentBuilder::new()
        .body_table("Fixed sections", 4, 4)
        .build()
        .unwrap();
    let model_id = editor.tables().unwrap()[0].model_object_id;
    editor
        .set_table_header_settings(
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
        .insert_table_row(model_id, PagesTableRowInsertion::header(1))
        .unwrap();
    editor
        .insert_table_row(model_id, PagesTableRowInsertion::footer(0))
        .unwrap();
    editor
        .insert_table_column(model_id, PagesTableColumnInsertion::header(1))
        .unwrap();
    let settings = editor.table_header_settings(model_id).unwrap();
    assert_eq!(settings.header_row_count(), 2);
    assert_eq!(settings.footer_row_count(), 2);
    assert_eq!(settings.header_column_count(), 2);
    let table = editor.table(model_id).unwrap();
    assert_eq!((table.info.rows, table.info.columns), (6, 5));

    editor
        .remove_table_column(model_id, PagesTableColumnDeletion::header(1))
        .unwrap();
    editor
        .remove_table_row(model_id, PagesTableRowDeletion::footer(0))
        .unwrap();
    editor
        .remove_table_row(model_id, PagesTableRowDeletion::header(1))
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);

    editor
        .set_table_cell(model_id, 0, 0, PagesCellValue::Text("Header".to_owned()))
        .unwrap();
    editor
        .set_table_cell(model_id, 1, 1, PagesCellValue::Text("Body".to_owned()))
        .unwrap();
    editor
        .set_table_cell(model_id, 3, 2, PagesCellValue::Text("Footer".to_owned()))
        .unwrap();
    editor
        .remove_table_row(model_id, PagesTableRowDeletion::header(0))
        .unwrap();
    editor
        .remove_table_row(model_id, PagesTableRowDeletion::footer(0))
        .unwrap();
    editor
        .remove_table_column(model_id, PagesTableColumnDeletion::header(0))
        .unwrap();
    let settings = editor.table_header_settings(model_id).unwrap();
    assert_eq!(settings.header_row_count(), 0);
    assert_eq!(settings.footer_row_count(), 0);
    assert_eq!(settings.header_column_count(), 0);
    let table = editor.table(model_id).unwrap();
    assert_eq!((table.info.rows, table.info.columns), (2, 3));
    assert_eq!(
        table.get_cell(0, 0),
        Some(&PagesCellValue::Text("Body".to_owned()))
    );
    assert!(!table.iter_cells().any(|(_, value)| matches!(
        value,
        PagesCellValue::Text(text) if text == "Header" || text == "Footer"
    )));
}

#[test]
fn source_built_table_roundtrips_title_settings_transactionally() {
    let mut editor = PagesDocumentBuilder::new()
        .body_table("Revenue", 2, 2)
        .build()
        .unwrap();
    let model_id = editor.tables().unwrap()[0].model_object_id;
    let visible = PagesTableTitleSettings {
        visible: Some(true),
        outlined: Some(true),
    };
    let initially_hidden = PagesTableTitleSettings {
        visible: Some(false),
        outlined: None,
    };
    assert_eq!(
        editor.table_title_settings(model_id).unwrap(),
        initially_hidden
    );
    editor.set_table_title_settings(model_id, visible).unwrap();

    let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(reopened.table_title_settings(model_id).unwrap(), visible);
    let unchanged = reopened.to_bytes().unwrap();
    reopened
        .set_table_title_settings(model_id, visible)
        .unwrap();
    assert_eq!(reopened.to_bytes().unwrap(), unchanged);

    let explicit_hidden = PagesTableTitleSettings {
        visible: Some(false),
        outlined: Some(false),
    };
    reopened
        .set_table_title_settings(model_id, explicit_hidden)
        .unwrap();
    assert_eq!(
        reopened.table_title_settings(model_id).unwrap(),
        explicit_hidden
    );
    reopened
        .set_table_title_settings(model_id, PagesTableTitleSettings::default())
        .unwrap();
    assert_eq!(
        reopened.table_title_settings(model_id).unwrap(),
        PagesTableTitleSettings::default()
    );

    let before_error = reopened.to_bytes().unwrap();
    assert!(
        reopened
            .set_table_title_settings(u64::MAX, visible)
            .is_err()
    );
    assert_eq!(reopened.to_bytes().unwrap(), before_error);
}

#[test]
fn inserts_independent_table_and_shifts_existing_anchor() {
    let body = "Alpha 🙂\nBeta\n";
    let mut editor = PagesDocumentBuilder::new()
        .body_text(body)
        .body_table("Source", 3, 2)
        .build()
        .unwrap();
    let source = editor.tables().unwrap().remove(0);
    editor
        .set_table_cell(
            source.model_object_id,
            1,
            1,
            PagesCellValue::Text("source only".to_owned()),
        )
        .unwrap();
    let anchor = "Alpha 🙂\n".encode_utf16().count();
    let mut package = editor.into_package();
    let engine = package.remove_entry("Index/CalculationEngine.iwa").unwrap();
    package
        .insert_entry("Index/CalculationEngine-83.iwa", engine)
        .unwrap();
    let mut editor = PagesEditor::from_package(package).unwrap();

    let inserted = editor.add_table(anchor, "Inserted", 4, 3).unwrap();
    assert_eq!(inserted.anchor_character_index, anchor);
    assert_eq!((inserted.rows, inserted.columns), (4, 3));
    assert_eq!(
        editor.table(inserted.model_object_id).unwrap().cell_count(),
        0
    );
    let tables = editor.tables().unwrap();
    assert_eq!(tables.len(), 2);
    assert_eq!(tables[0], inserted);
    assert_eq!(
        tables[1].anchor_character_index,
        body.encode_utf16().count() + 1
    );
    assert_eq!(tables[1].model_object_id, source.model_object_id);

    editor
        .set_table_cell(
            inserted.model_object_id,
            0,
            0,
            PagesCellValue::Text("inserted only".to_owned()),
        )
        .unwrap();
    let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .table(source.model_object_id)
            .unwrap()
            .get_cell(1, 1),
        Some(&PagesCellValue::Text("source only".to_owned()))
    );
    assert_eq!(
        reopened
            .table(inserted.model_object_id)
            .unwrap()
            .get_cell(0, 0),
        Some(&PagesCellValue::Text("inserted only".to_owned()))
    );
    reopened.remove_table(inserted.model_object_id).unwrap();
    let retained = reopened.tables().unwrap();
    assert_eq!(retained.len(), 1);
    assert_eq!(retained[0].model_object_id, source.model_object_id);
    assert_eq!(
        retained[0].anchor_character_index,
        body.encode_utf16().count()
    );
    assert_eq!(
        reopened
            .table(source.model_object_id)
            .unwrap()
            .get_cell(1, 1),
        Some(&PagesCellValue::Text("source only".to_owned()))
    );
}

#[test]
fn inserts_and_removes_first_table_without_a_template() {
    let mut editor = PagesDocumentBuilder::new()
        .body_text("Before 🙂 after")
        .build()
        .unwrap();
    let anchor = "Before 🙂".encode_utf16().count();
    let created = editor.add_table(anchor, "First runtime", 2, 3).unwrap();
    assert_eq!(created.anchor_character_index, anchor);
    assert_eq!((created.rows, created.columns), (2, 3));
    editor
        .set_table_cell(
            created.model_object_id,
            1,
            2,
            PagesCellValue::Text("bootstrapped".to_owned()),
        )
        .unwrap();
    let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .table(created.model_object_id)
            .unwrap()
            .get_cell(1, 2),
        Some(&PagesCellValue::Text("bootstrapped".to_owned()))
    );
    reopened.remove_table(created.model_object_id).unwrap();
    assert!(reopened.tables().unwrap().is_empty());
    assert_eq!(reopened.body_text().unwrap(), "Before 🙂 after");
}

#[test]
fn first_table_bootstrap_rejects_reserved_id_collision_transactionally() {
    let mut package = PagesDocumentBuilder::new()
        .body_text("Collision")
        .build_package()
        .unwrap();
    package
        .update_archive("Index/Document.iwa", |archive| {
            archive.insert_object(ArchiveObject::new(
                SOURCE_BUILT_TABLE_INFO_OBJECT_ID,
                vec![RawMessage {
                    type_: u32::MAX,
                    data: Vec::new(),
                }],
            )?)
        })
        .unwrap();
    let mut editor = PagesEditor::from_package(package).unwrap();
    let before = editor.to_bytes().unwrap();
    assert!(editor.add_table(0, "Collision", 2, 2).is_err());
    assert_eq!(editor.to_bytes().unwrap(), before);
}

#[test]
fn invalid_insertion_anchor_is_transactional() {
    let mut editor = PagesDocumentBuilder::new()
        .body_text("Short")
        .body_table("Template", 2, 2)
        .build()
        .unwrap();
    let before = editor.to_bytes().unwrap();
    assert!(editor.add_table(usize::MAX, "Invalid", 2, 2).is_err());
    assert_eq!(editor.to_bytes().unwrap(), before);
}

#[test]
fn out_of_bounds_cell_update_is_transactional() {
    let mut editor = PagesDocumentBuilder::new()
        .body_table("Bounded", 2, 2)
        .build()
        .unwrap();
    let model_id = editor.tables().unwrap()[0].model_object_id;
    let before = editor.to_bytes().unwrap();
    assert!(
        editor
            .set_table_cell(model_id, 2, 0, PagesCellValue::Boolean(true))
            .is_err()
    );
    assert_eq!(editor.to_bytes().unwrap(), before);
}

#[test]
fn source_built_table_roundtrips_rename_and_resize() {
    let mut editor = PagesDocumentBuilder::new()
        .body_table("Original", 3, 2)
        .build()
        .unwrap();
    let model_id = editor.tables().unwrap()[0].model_object_id;
    editor
        .set_table_cell(model_id, 1, 1, PagesCellValue::Text("kept".to_owned()))
        .unwrap();

    editor.rename_table(model_id, "Renamed").unwrap();
    editor.resize_table(model_id, 5, 4).unwrap();
    let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    let info = reopened.tables().unwrap().remove(0);
    assert_eq!(info.name, "Renamed");
    assert_eq!((info.rows, info.columns), (5, 4));
    assert_eq!(
        reopened.table(model_id).unwrap().get_cell(1, 1),
        Some(&PagesCellValue::Text("kept".to_owned()))
    );

    reopened.resize_table(model_id, 2, 2).unwrap();
    let info = reopened.tables().unwrap().remove(0);
    assert_eq!((info.rows, info.columns), (2, 2));
}

#[test]
fn source_built_table_roundtrips_layout_crud_transactionally() {
    let mut editor = PagesDocumentBuilder::new()
        .body_table("Layout", 4, 3)
        .build()
        .unwrap();
    let model_id = editor.tables().unwrap()[0].model_object_id;
    let settings = HeaderSettings {
        header_rows: Some(HeaderCount::TWO),
        header_columns: Some(HeaderCount::ONE),
        footer_rows: Some(HeaderCount::ONE),
        ..Default::default()
    };

    editor
        .set_table_header_settings(model_id, settings)
        .unwrap();
    editor
        .set_table_column_width(model_id, 0, PagesTableDimensionSize::points(150.0).unwrap())
        .unwrap();
    editor
        .set_table_row_height(model_id, 2, PagesTableDimensionSize::points(42.0).unwrap())
        .unwrap();

    let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(reopened.table_header_settings(model_id).unwrap(), settings);
    assert_eq!(
        reopened.table_column_width(model_id, 0).unwrap(),
        PagesTableDimensionSize::points(150.0).unwrap()
    );
    assert_eq!(
        reopened.table_row_height(model_id, 2).unwrap(),
        PagesTableDimensionSize::points(42.0).unwrap()
    );
    assert_eq!(
        reopened.table_column_width(model_id, 1).unwrap(),
        PagesTableDimensionSize::Default
    );

    reopened
        .set_table_column_width(model_id, 0, PagesTableDimensionSize::Default)
        .unwrap();
    reopened
        .set_table_row_height(model_id, 2, PagesTableDimensionSize::Default)
        .unwrap();
    assert_eq!(
        reopened.table_column_width(model_id, 0).unwrap(),
        PagesTableDimensionSize::Default
    );
    assert_eq!(
        reopened.table_row_height(model_id, 2).unwrap(),
        PagesTableDimensionSize::Default
    );

    let before = reopened.to_bytes().unwrap();
    assert!(
        reopened
            .set_table_row_height(
                model_id,
                usize::MAX,
                PagesTableDimensionSize::points(20.0).unwrap(),
            )
            .is_err()
    );
    assert_eq!(reopened.to_bytes().unwrap(), before);
    assert!(
        reopened
            .set_table_header_settings(
                model_id,
                HeaderSettings {
                    header_rows: Some(HeaderCount::FOUR),
                    footer_rows: Some(HeaderCount::ONE),
                    ..Default::default()
                },
            )
            .is_err()
    );
    assert_eq!(reopened.to_bytes().unwrap(), before);
}

#[test]
fn table_rename_and_occupied_shrink_are_transactional() {
    let mut editor = PagesDocumentBuilder::new()
        .body_table("Protected", 3, 3)
        .build()
        .unwrap();
    let model_id = editor.tables().unwrap()[0].model_object_id;
    editor
        .set_table_cell(
            model_id,
            2,
            2,
            PagesCellValue::number(7.0).expect("finite test number"),
        )
        .unwrap();

    let before = editor.to_bytes().unwrap();
    assert!(editor.rename_table(model_id, "").is_err());
    assert_eq!(editor.to_bytes().unwrap(), before);
    assert!(editor.resize_table(model_id, 2, 2).is_err());
    assert_eq!(editor.to_bytes().unwrap(), before);
}

#[test]
fn source_built_table_deletion_removes_private_graph_and_anchor() {
    let body = "Before 🙂 after";
    let mut editor = PagesDocumentBuilder::new()
        .body_text(body)
        .body_table("Disposable", 3, 2)
        .build()
        .unwrap();
    let table = editor.tables().unwrap().remove(0);
    let owned = crate::numbers::editor::table_owned_object_ids_in_package(
        editor.package(),
        table.model_object_id,
    )
    .unwrap();
    editor
        .set_table_cell(
            table.model_object_id,
            1,
            1,
            PagesCellValue::Text("removed".to_owned()),
        )
        .unwrap();

    let removed = editor.remove_table(table.model_object_id).unwrap();
    assert_eq!(removed, table);
    assert!(editor.tables().unwrap().is_empty());
    assert_eq!(editor.body_text().unwrap(), body);
    let mut removed_ids = owned;
    removed_ids.extend([table.drawable_object_id, table.model_object_id]);
    for identifier in removed_ids {
        assert!(find_object_archive(editor.package(), identifier).is_err());
    }
    let reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert!(reopened.tables().unwrap().is_empty());
    assert_eq!(reopened.body_text().unwrap(), body);
}

#[test]
fn missing_table_deletion_is_transactional() {
    let mut editor = PagesDocumentBuilder::new()
        .body_table("Retained", 2, 2)
        .build()
        .unwrap();
    let before = editor.to_bytes().unwrap();
    assert!(editor.remove_table(u64::MAX).is_err());
    assert_eq!(editor.to_bytes().unwrap(), before);
}
