//! Create Pages, Numbers, and Keynote files with native table-cell formats.

use std::path::{Path, PathBuf};

use litchi_iwa::keynote::KeynoteDocumentBuilder;
use litchi_iwa::numbers::{NumbersDocumentBuilder, NumbersEditor};
use litchi_iwa::pages::PagesDocumentBuilder;
use litchi_iwa::shapes::{DrawablePoint, DrawableSize};
use litchi_numbers::cell::Value as CellValue;
use litchi_numbers::cell::data_format::duration::{Style as DurationStyle, UnitRange};
use litchi_numbers::cell::data_format::numeral_system::{
    Base, FixedPlaces, NegativeStyle as NumeralNegativeStyle, Places,
};
use litchi_numbers::cell::data_format::{
    self as numbers, Checkbox, Currency, CurrencyCode, CurrencyStyle, DateTime, DecimalPlaces,
    Duration, FixedDecimalPlaces, Fraction, FractionAccuracy, NegativeStyle, Number, NumeralSystem,
    Percentage, PopUpMenu, Scientific, Slider, StarRating, Stepper,
};

const ROW: usize = 1;
const NUMBER_COLUMN: usize = 1;
const PERCENTAGE_COLUMN: usize = 2;
const CURRENCY_COLUMN: usize = 3;
const SCIENTIFIC_COLUMN: usize = 4;
const FRACTION_COLUMN: usize = 5;
const NUMERAL_SYSTEM_COLUMN: usize = 6;
const DATE_TIME_COLUMN: usize = 7;
const DURATION_COLUMN: usize = 8;
const CHECKBOX_COLUMN: usize = 9;
const STAR_RATING_COLUMN: usize = 10;
const SLIDER_COLUMN: usize = 11;
const STEPPER_COLUMN: usize = 12;
const POP_UP_MENU_COLUMN: usize = 13;
const TEXT_COLUMN: usize = 14;
const CUSTOM_NUMBER_COLUMN: usize = 15;
const CUSTOM_DATE_TIME_COLUMN: usize = 16;
const CUSTOM_TEXT_COLUMN: usize = 17;
const NUMBER_VALUE: f64 = -1_234.5;
const PERCENTAGE_VALUE: f64 = -12.345;
const CURRENCY_VALUE: f64 = -1_234.5;
const SCIENTIFIC_VALUE: f64 = -1_234.5;
const FRACTION_VALUE: f64 = -12.375;
const NUMERAL_SYSTEM_VALUE: f64 = -1_234.5;
const DATE_TIME_VALUE: f64 = 789_332_889.0;
const DURATION_VALUE: f64 = 3_723.5;
const STAR_RATING_VALUE: f64 = 4.0;
const SLIDER_VALUE: f64 = 25.0;
const STEPPER_VALUE: f64 = 25.0;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let output = PathBuf::from(
        std::env::args()
            .nth(1)
            .ok_or("usage: create_iwork_table_number_formats <output-directory>")?,
    );
    std::fs::create_dir_all(&output)?;
    create_numbers(&output.join("table-number-formats.numbers"))?;
    create_pages(&output.join("table-number-formats.pages"))?;
    create_keynote(&output.join("table-number-formats.key"))?;
    Ok(())
}

fn semantic_format() -> Result<Number, Box<dyn std::error::Error>> {
    numbers_format()
}

fn numbers_format() -> Result<Number, Box<dyn std::error::Error>> {
    Ok(Number::new(
        DecimalPlaces::fixed(2)?,
        NegativeStyle::Parentheses,
        numbers::ThousandsSeparator::Shown,
    ))
}

fn numbers_percentage_format() -> Result<Percentage, Box<dyn std::error::Error>> {
    Ok(Percentage::new(
        DecimalPlaces::fixed(2)?,
        NegativeStyle::Parentheses,
        numbers::ThousandsSeparator::Shown,
    ))
}

fn numbers_currency_format() -> Result<Currency, Box<dyn std::error::Error>> {
    Ok(Currency::new(
        CurrencyCode::USD,
        DecimalPlaces::fixed(2)?,
        NegativeStyle::Parentheses,
        numbers::ThousandsSeparator::Shown,
        CurrencyStyle::Accounting,
    ))
}

fn numbers_scientific_format() -> Result<Scientific, Box<dyn std::error::Error>> {
    Ok(Scientific::new(FixedDecimalPlaces::new(5)?))
}

const fn numbers_fraction_format() -> Fraction {
    Fraction::new(FractionAccuracy::Eighths)
}

fn numbers_numeral_system_format() -> Result<NumeralSystem, Box<dyn std::error::Error>> {
    Ok(NumeralSystem::new(
        Base::HEXADECIMAL,
        Places::Fixed(FixedPlaces::EIGHT),
        NumeralNegativeStyle::TwosComplement,
    )?)
}

fn numbers_date_time_format() -> DateTime {
    DateTime::iso_date_time_24_hour_with_seconds()
}

const fn numbers_duration_format() -> Duration {
    Duration::custom(
        DurationStyle::Abbreviated,
        UnitRange::hours_to_milliseconds(),
    )
}

fn numbers_slider_format() -> Result<Slider, Box<dyn std::error::Error>> {
    Ok(Slider::new(
        numbers::control::Range::new(-10.0, 30.0, 0.5)?,
        numbers_format()?.into(),
    ))
}

fn numbers_stepper_format() -> Result<Stepper, Box<dyn std::error::Error>> {
    Ok(Stepper::new(
        numbers::control::Range::new(-10.0, 30.0, 0.5)?,
        numbers_format()?.into(),
    ))
}

fn numbers_pop_up_menu_format() -> Result<PopUpMenu, Box<dyn std::error::Error>> {
    Ok(PopUpMenu::new(["Low", "Medium", "High"])?)
}

fn numbers_custom_number_format() -> Result<numbers::Custom, Box<dyn std::error::Error>> {
    Ok(numbers::custom::Number::try_with_rules(
        numbers::custom::Name::try_new("Grouped Integer")?,
        numbers::custom::NumberPattern::try_new("#,###")?,
        [numbers::custom::NumberRule::new(
            numbers::custom::Condition::LessThan(numbers::custom::ConditionValue::try_new(0.0)?),
            numbers::custom::NumberPattern::try_new("(#,###)")?,
        )],
    )?
    .into())
}

fn numbers_custom_date_time_format() -> Result<numbers::Custom, Box<dyn std::error::Error>> {
    Ok(numbers::custom::DateTime::new(
        numbers::custom::Name::try_new("Month Day Year")?,
        numbers::custom::DateTimePattern::try_new("MMM d, y")?,
    )
    .into())
}

fn numbers_custom_text_format() -> Result<numbers::Custom, Box<dyn std::error::Error>> {
    Ok(numbers::custom::Text::try_new(
        numbers::custom::Name::try_new("Text With ID Suffix")?,
        "",
        "ID: ",
    )?
    .into())
}

fn create_numbers(output: &Path) -> Result<(), Box<dyn std::error::Error>> {
    let mut editor = NumbersDocumentBuilder::new()
        .table_name("Number Formats")
        .table_dimensions(3, 18)
        .build()?;
    let table_id = editor.tables()?.remove(0).id();
    set_numbers_cells(
        &mut editor,
        [
            litchi_numbers::cell::Update::new(ROW, NUMBER_COLUMN, CellValue::number(NUMBER_VALUE)?),
            litchi_numbers::cell::Update::new(
                ROW,
                PERCENTAGE_COLUMN,
                CellValue::number(PERCENTAGE_VALUE)?,
            ),
            litchi_numbers::cell::Update::new(
                ROW,
                CURRENCY_COLUMN,
                CellValue::number(CURRENCY_VALUE)?,
            ),
            litchi_numbers::cell::Update::new(
                ROW,
                SCIENTIFIC_COLUMN,
                CellValue::number(SCIENTIFIC_VALUE)?,
            ),
            litchi_numbers::cell::Update::new(
                ROW,
                FRACTION_COLUMN,
                CellValue::number(FRACTION_VALUE)?,
            ),
            litchi_numbers::cell::Update::new(
                ROW,
                NUMERAL_SYSTEM_COLUMN,
                CellValue::number(NUMERAL_SYSTEM_VALUE)?,
            ),
            litchi_numbers::cell::Update::new(
                ROW,
                DATE_TIME_COLUMN,
                CellValue::date(DATE_TIME_VALUE)?,
            ),
            litchi_numbers::cell::Update::new(
                ROW,
                DURATION_COLUMN,
                CellValue::duration(DURATION_VALUE)?,
            ),
            litchi_numbers::cell::Update::new(ROW, CHECKBOX_COLUMN, CellValue::Boolean(true)),
            litchi_numbers::cell::Update::new(
                ROW,
                STAR_RATING_COLUMN,
                CellValue::number(STAR_RATING_VALUE)?,
            ),
            litchi_numbers::cell::Update::new(ROW, SLIDER_COLUMN, CellValue::number(SLIDER_VALUE)?),
            litchi_numbers::cell::Update::new(
                ROW,
                STEPPER_COLUMN,
                CellValue::number(STEPPER_VALUE)?,
            ),
            litchi_numbers::cell::Update::new(
                ROW,
                TEXT_COLUMN,
                CellValue::Text("Invoice 001".to_owned()),
            ),
            litchi_numbers::cell::Update::new(
                ROW,
                CUSTOM_NUMBER_COLUMN,
                CellValue::number(NUMBER_VALUE)?,
            ),
            litchi_numbers::cell::Update::new(
                ROW,
                CUSTOM_DATE_TIME_COLUMN,
                CellValue::date(DATE_TIME_VALUE)?,
            ),
            litchi_numbers::cell::Update::new(
                ROW,
                CUSTOM_TEXT_COLUMN,
                CellValue::Text("Invoice 001".to_owned()),
            ),
        ],
    )?;
    editor.set_table_cell_number_format(table_id, ROW, NUMBER_COLUMN, numbers_format()?)?;
    editor.set_table_cell_percentage_format(
        table_id,
        ROW,
        PERCENTAGE_COLUMN,
        numbers_percentage_format()?,
    )?;
    editor.set_table_cell_currency_format(
        table_id,
        ROW,
        CURRENCY_COLUMN,
        numbers_currency_format()?,
    )?;
    editor.set_table_cell_scientific_format(
        table_id,
        ROW,
        SCIENTIFIC_COLUMN,
        numbers_scientific_format()?,
    )?;
    editor.set_table_cell_fraction_format(
        table_id,
        ROW,
        FRACTION_COLUMN,
        numbers_fraction_format(),
    )?;
    editor.set_table_cell_numeral_system_format(
        table_id,
        ROW,
        NUMERAL_SYSTEM_COLUMN,
        numbers_numeral_system_format()?,
    )?;
    editor.set_table_cell_date_time_format(
        table_id,
        ROW,
        DATE_TIME_COLUMN,
        numbers_date_time_format(),
    )?;
    editor.set_table_cell_duration_format(
        table_id,
        ROW,
        DURATION_COLUMN,
        numbers_duration_format(),
    )?;
    editor.set_table_cell_checkbox_format(table_id, ROW, CHECKBOX_COLUMN, Checkbox)?;
    editor.set_table_cell_star_rating_format(table_id, ROW, STAR_RATING_COLUMN, StarRating)?;
    editor.set_table_cell_slider_format(table_id, ROW, SLIDER_COLUMN, numbers_slider_format()?)?;
    editor.set_table_cell_stepper_format(
        table_id,
        ROW,
        STEPPER_COLUMN,
        numbers_stepper_format()?,
    )?;
    editor.set_table_cell_pop_up_menu_format(
        table_id,
        ROW,
        POP_UP_MENU_COLUMN,
        numbers_pop_up_menu_format()?,
    )?;
    editor.set_table_cell_text_format(table_id, ROW, TEXT_COLUMN)?;
    editor.set_table_cell_custom_format(
        table_id,
        ROW,
        CUSTOM_NUMBER_COLUMN,
        numbers_custom_number_format()?,
    )?;
    editor.set_table_cell_custom_format(
        table_id,
        ROW,
        CUSTOM_DATE_TIME_COLUMN,
        numbers_custom_date_time_format()?,
    )?;
    editor.set_table_cell_custom_format(
        table_id,
        ROW,
        CUSTOM_TEXT_COLUMN,
        numbers_custom_text_format()?,
    )?;
    editor.save(output)?;
    Ok(())
}

fn create_pages(output: &Path) -> Result<(), Box<dyn std::error::Error>> {
    let mut editor = PagesDocumentBuilder::new()
        .body_text("Created from scratch with native table-cell formats.\n")
        .body_table("Number Formats", 3, 18)
        .build()?;
    let table_id = editor.tables()?.remove(0).model_object_id;
    editor.set_table_cell(
        table_id,
        ROW,
        NUMBER_COLUMN,
        CellValue::number(NUMBER_VALUE)?,
    )?;
    editor.set_table_cell_number_format(table_id, ROW, NUMBER_COLUMN, semantic_format()?)?;
    editor.set_table_cell(
        table_id,
        ROW,
        PERCENTAGE_COLUMN,
        CellValue::number(PERCENTAGE_VALUE)?,
    )?;
    editor.set_table_cell_percentage_format(
        table_id,
        ROW,
        PERCENTAGE_COLUMN,
        numbers_percentage_format()?,
    )?;
    editor.set_table_cell(
        table_id,
        ROW,
        CURRENCY_COLUMN,
        CellValue::number(CURRENCY_VALUE)?,
    )?;
    editor.set_table_cell_currency_format(
        table_id,
        ROW,
        CURRENCY_COLUMN,
        numbers_currency_format()?,
    )?;
    editor.set_table_cell(
        table_id,
        ROW,
        SCIENTIFIC_COLUMN,
        CellValue::number(SCIENTIFIC_VALUE)?,
    )?;
    editor.set_table_cell_scientific_format(
        table_id,
        ROW,
        SCIENTIFIC_COLUMN,
        numbers_scientific_format()?,
    )?;
    editor.set_table_cell(
        table_id,
        ROW,
        FRACTION_COLUMN,
        CellValue::number(FRACTION_VALUE)?,
    )?;
    editor.set_table_cell_fraction_format(
        table_id,
        ROW,
        FRACTION_COLUMN,
        numbers_fraction_format(),
    )?;
    editor.set_table_cell(
        table_id,
        ROW,
        NUMERAL_SYSTEM_COLUMN,
        CellValue::number(NUMERAL_SYSTEM_VALUE)?,
    )?;
    editor.set_table_cell_numeral_system_format(
        table_id,
        ROW,
        NUMERAL_SYSTEM_COLUMN,
        numbers_numeral_system_format()?,
    )?;
    editor.set_table_cell(
        table_id,
        ROW,
        DATE_TIME_COLUMN,
        CellValue::date(DATE_TIME_VALUE)?,
    )?;
    editor.set_table_cell_date_time_format(
        table_id,
        ROW,
        DATE_TIME_COLUMN,
        numbers_date_time_format(),
    )?;
    editor.set_table_cell(
        table_id,
        ROW,
        DURATION_COLUMN,
        CellValue::duration(DURATION_VALUE)?,
    )?;
    editor.set_table_cell_duration_format(
        table_id,
        ROW,
        DURATION_COLUMN,
        numbers_duration_format(),
    )?;
    editor.set_table_cell(table_id, ROW, CHECKBOX_COLUMN, CellValue::Boolean(true))?;
    editor.set_table_cell_checkbox_format(table_id, ROW, CHECKBOX_COLUMN, Checkbox)?;
    editor.set_table_cell(
        table_id,
        ROW,
        STAR_RATING_COLUMN,
        CellValue::number(STAR_RATING_VALUE)?,
    )?;
    editor.set_table_cell_star_rating_format(table_id, ROW, STAR_RATING_COLUMN, StarRating)?;
    editor.set_table_cell(
        table_id,
        ROW,
        SLIDER_COLUMN,
        CellValue::number(SLIDER_VALUE)?,
    )?;
    editor.set_table_cell_slider_format(table_id, ROW, SLIDER_COLUMN, numbers_slider_format()?)?;
    editor.set_table_cell(
        table_id,
        ROW,
        STEPPER_COLUMN,
        CellValue::number(STEPPER_VALUE)?,
    )?;
    editor.set_table_cell_stepper_format(
        table_id,
        ROW,
        STEPPER_COLUMN,
        numbers_stepper_format()?,
    )?;
    editor.set_table_cell_pop_up_menu_format(
        table_id,
        ROW,
        POP_UP_MENU_COLUMN,
        numbers_pop_up_menu_format()?,
    )?;
    editor.set_table_cell(
        table_id,
        ROW,
        TEXT_COLUMN,
        CellValue::Text("Invoice 001".to_owned()),
    )?;
    editor.set_table_cell_text_format(table_id, ROW, TEXT_COLUMN)?;
    editor.set_table_cell(
        table_id,
        ROW,
        CUSTOM_NUMBER_COLUMN,
        CellValue::number(NUMBER_VALUE)?,
    )?;
    editor.set_table_cell_custom_format(
        table_id,
        ROW,
        CUSTOM_NUMBER_COLUMN,
        numbers_custom_number_format()?,
    )?;
    editor.set_table_cell(
        table_id,
        ROW,
        CUSTOM_DATE_TIME_COLUMN,
        CellValue::date(DATE_TIME_VALUE)?,
    )?;
    editor.set_table_cell_custom_format(
        table_id,
        ROW,
        CUSTOM_DATE_TIME_COLUMN,
        numbers_custom_date_time_format()?,
    )?;
    editor.set_table_cell(
        table_id,
        ROW,
        CUSTOM_TEXT_COLUMN,
        CellValue::Text("Invoice 001".to_owned()),
    )?;
    editor.set_table_cell_custom_format(
        table_id,
        ROW,
        CUSTOM_TEXT_COLUMN,
        numbers_custom_text_format()?,
    )?;
    editor.save(output)?;
    Ok(())
}

fn create_keynote(output: &Path) -> Result<(), Box<dyn std::error::Error>> {
    let mut editor = KeynoteDocumentBuilder::new()
        .title("Native table-cell formats")
        .build()?;
    let table = editor.add_slide_table(
        0,
        "Number Formats",
        3,
        18,
        DrawablePoint { x: 320.0, y: 360.0 },
        DrawableSize {
            width: 1_280.0,
            height: 480.0,
        },
    )?;
    editor.set_slide_table_cell(
        0,
        table.model_object_id,
        ROW,
        NUMBER_COLUMN,
        CellValue::number(NUMBER_VALUE)?,
    )?;
    editor.set_slide_table_cell_number_format(
        0,
        table.model_object_id,
        ROW,
        NUMBER_COLUMN,
        semantic_format()?,
    )?;
    editor.set_slide_table_cell(
        0,
        table.model_object_id,
        ROW,
        PERCENTAGE_COLUMN,
        CellValue::number(PERCENTAGE_VALUE)?,
    )?;
    editor.set_slide_table_cell_percentage_format(
        0,
        table.model_object_id,
        ROW,
        PERCENTAGE_COLUMN,
        numbers_percentage_format()?,
    )?;
    editor.set_slide_table_cell(
        0,
        table.model_object_id,
        ROW,
        CURRENCY_COLUMN,
        CellValue::number(CURRENCY_VALUE)?,
    )?;
    editor.set_slide_table_cell_currency_format(
        0,
        table.model_object_id,
        ROW,
        CURRENCY_COLUMN,
        numbers_currency_format()?,
    )?;
    editor.set_slide_table_cell(
        0,
        table.model_object_id,
        ROW,
        SCIENTIFIC_COLUMN,
        CellValue::number(SCIENTIFIC_VALUE)?,
    )?;
    editor.set_slide_table_cell_scientific_format(
        0,
        table.model_object_id,
        ROW,
        SCIENTIFIC_COLUMN,
        numbers_scientific_format()?,
    )?;
    editor.set_slide_table_cell(
        0,
        table.model_object_id,
        ROW,
        FRACTION_COLUMN,
        CellValue::number(FRACTION_VALUE)?,
    )?;
    editor.set_slide_table_cell_fraction_format(
        0,
        table.model_object_id,
        ROW,
        FRACTION_COLUMN,
        numbers_fraction_format(),
    )?;
    editor.set_slide_table_cell(
        0,
        table.model_object_id,
        ROW,
        NUMERAL_SYSTEM_COLUMN,
        CellValue::number(NUMERAL_SYSTEM_VALUE)?,
    )?;
    editor.set_slide_table_cell_numeral_system_format(
        0,
        table.model_object_id,
        ROW,
        NUMERAL_SYSTEM_COLUMN,
        numbers_numeral_system_format()?,
    )?;
    editor.set_slide_table_cell(
        0,
        table.model_object_id,
        ROW,
        DATE_TIME_COLUMN,
        CellValue::date(DATE_TIME_VALUE)?,
    )?;
    editor.set_slide_table_cell_date_time_format(
        0,
        table.model_object_id,
        ROW,
        DATE_TIME_COLUMN,
        numbers_date_time_format(),
    )?;
    editor.set_slide_table_cell(
        0,
        table.model_object_id,
        ROW,
        DURATION_COLUMN,
        CellValue::duration(DURATION_VALUE)?,
    )?;
    editor.set_slide_table_cell_duration_format(
        0,
        table.model_object_id,
        ROW,
        DURATION_COLUMN,
        numbers_duration_format(),
    )?;
    editor.set_slide_table_cell(
        0,
        table.model_object_id,
        ROW,
        CHECKBOX_COLUMN,
        CellValue::Boolean(true),
    )?;
    editor.set_slide_table_cell_checkbox_format(
        0,
        table.model_object_id,
        ROW,
        CHECKBOX_COLUMN,
        Checkbox,
    )?;
    editor.set_slide_table_cell(
        0,
        table.model_object_id,
        ROW,
        STAR_RATING_COLUMN,
        CellValue::number(STAR_RATING_VALUE)?,
    )?;
    editor.set_slide_table_cell_star_rating_format(
        0,
        table.model_object_id,
        ROW,
        STAR_RATING_COLUMN,
        StarRating,
    )?;
    editor.set_slide_table_cell(
        0,
        table.model_object_id,
        ROW,
        SLIDER_COLUMN,
        CellValue::number(SLIDER_VALUE)?,
    )?;
    editor.set_slide_table_cell_slider_format(
        0,
        table.model_object_id,
        ROW,
        SLIDER_COLUMN,
        numbers_slider_format()?,
    )?;
    editor.set_slide_table_cell(
        0,
        table.model_object_id,
        ROW,
        STEPPER_COLUMN,
        CellValue::number(STEPPER_VALUE)?,
    )?;
    editor.set_slide_table_cell_stepper_format(
        0,
        table.model_object_id,
        ROW,
        STEPPER_COLUMN,
        numbers_stepper_format()?,
    )?;
    editor.set_slide_table_cell_pop_up_menu_format(
        0,
        table.model_object_id,
        ROW,
        POP_UP_MENU_COLUMN,
        numbers_pop_up_menu_format()?,
    )?;
    editor.set_slide_table_cell(
        0,
        table.model_object_id,
        ROW,
        TEXT_COLUMN,
        CellValue::Text("Invoice 001".to_owned()),
    )?;
    editor.set_slide_table_cell_text_format(0, table.model_object_id, ROW, TEXT_COLUMN)?;
    for (column, value, format) in [
        (
            CUSTOM_NUMBER_COLUMN,
            CellValue::number(NUMBER_VALUE)?,
            numbers_custom_number_format()?,
        ),
        (
            CUSTOM_DATE_TIME_COLUMN,
            CellValue::date(DATE_TIME_VALUE)?,
            numbers_custom_date_time_format()?,
        ),
        (
            CUSTOM_TEXT_COLUMN,
            CellValue::Text("Invoice 001".to_owned()),
            numbers_custom_text_format()?,
        ),
    ] {
        editor.set_slide_table_cell(0, table.model_object_id, ROW, column, value)?;
        editor.set_slide_table_cell_custom_format(0, table.model_object_id, ROW, column, format)?;
    }
    editor.save(output)?;
    Ok(())
}

fn set_numbers_cells(
    editor: &mut NumbersEditor,
    updates: impl IntoIterator<Item = litchi_numbers::cell::Update>,
) -> Result<(), Box<dyn std::error::Error>> {
    let changes = updates
        .into_iter()
        .map(numbers_cell_change)
        .collect::<Result<Vec<_>, _>>()?;
    let package = litchi_numbers::Package::from_bytes(&editor.to_bytes()?)?;
    let commit = package
        .edit_table_cells(
            litchi_numbers::SheetSelector::index(0),
            litchi_numbers::TableSelector::index(0),
        )?
        .extend(changes)?
        .commit()?;
    let mut bytes = Vec::new();
    commit.package().write_to(&mut bytes)?;
    *editor = NumbersEditor::from_bytes(&bytes)?;
    Ok(())
}

fn numbers_cell_change(
    update: litchi_numbers::cell::Update,
) -> Result<litchi_numbers::table::cells::Change, Box<dyn std::error::Error>> {
    let position = litchi_numbers::CellPosition::try_from_usize(update.row, update.column)?;
    let change = match update.value {
        CellValue::Empty => litchi_numbers::table::cells::Change::clear(position),
        CellValue::Text(value) => litchi_numbers::table::cells::Change::set(
            position,
            litchi_numbers::table::cells::Input::text(value)?,
        ),
        CellValue::Number(value) => litchi_numbers::table::cells::Change::set(
            position,
            litchi_numbers::table::cells::Input::number(value.get())?,
        ),
        CellValue::Boolean(value) => litchi_numbers::table::cells::Change::set(
            position,
            litchi_numbers::table::cells::Input::boolean(value),
        ),
        CellValue::Date(value) => litchi_numbers::table::cells::Change::set(
            position,
            litchi_numbers::table::cells::Input::date(value.get())?,
        ),
        CellValue::Duration(value) => litchi_numbers::table::cells::Change::set(
            position,
            litchi_numbers::table::cells::Input::duration(value.get())?,
        ),
        CellValue::Formula(_) | CellValue::Error(_) => {
            return Err(std::io::Error::other("unsupported Numbers cell input").into());
        },
    };
    Ok(change)
}
