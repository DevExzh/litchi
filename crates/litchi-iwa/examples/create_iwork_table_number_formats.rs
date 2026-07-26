//! Create Pages, Numbers, and Keynote files with native table-cell formats.

use std::path::{Path, PathBuf};

use litchi_iwa::keynote::KeynoteDocumentBuilder;
use litchi_iwa::numbers::{CellValue, NumbersDocumentBuilder};
use litchi_iwa::pages::PagesDocumentBuilder;
use litchi_iwa::shapes::{DrawablePoint, DrawableSize};
use litchi_iwa::table_cell_data_format::{
    TableCellCheckboxFormat, TableCellCurrencyCode, TableCellCurrencyFormat,
    TableCellCurrencyStyle, TableCellDateTimeFormat, TableCellDecimalPlaces,
    TableCellDurationFormat, TableCellDurationStyle, TableCellDurationUnitRange,
    TableCellFixedDecimalPlaces, TableCellFractionAccuracy, TableCellFractionFormat,
    TableCellNegativeNumberStyle, TableCellNumberFormat, TableCellNumeralSystemBase,
    TableCellNumeralSystemFixedPlaces, TableCellNumeralSystemFormat,
    TableCellNumeralSystemNegativeStyle, TableCellNumeralSystemPlaces, TableCellPercentageFormat,
    TableCellScientificFormat, TableCellSliderFormat, TableCellSliderRange,
    TableCellStarRatingFormat, TableCellStepperFormat, TableCellStepperRange,
    TableCellThousandsSeparator,
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

fn format() -> Result<TableCellNumberFormat, litchi_iwa::Error> {
    Ok(TableCellNumberFormat::new(
        TableCellDecimalPlaces::fixed(2)?,
        TableCellNegativeNumberStyle::Parentheses,
        TableCellThousandsSeparator::Shown,
    ))
}

fn percentage_format() -> Result<TableCellPercentageFormat, litchi_iwa::Error> {
    Ok(TableCellPercentageFormat::new(
        TableCellDecimalPlaces::fixed(2)?,
        TableCellNegativeNumberStyle::Parentheses,
        TableCellThousandsSeparator::Shown,
    ))
}

fn currency_format() -> Result<TableCellCurrencyFormat, litchi_iwa::Error> {
    Ok(TableCellCurrencyFormat::new(
        TableCellCurrencyCode::USD,
        TableCellDecimalPlaces::fixed(2)?,
        TableCellNegativeNumberStyle::Parentheses,
        TableCellThousandsSeparator::Shown,
        TableCellCurrencyStyle::Accounting,
    ))
}

fn scientific_format() -> Result<TableCellScientificFormat, litchi_iwa::Error> {
    Ok(TableCellScientificFormat::new(
        TableCellFixedDecimalPlaces::new(5)?,
    ))
}

const fn fraction_format() -> TableCellFractionFormat {
    TableCellFractionFormat::new(TableCellFractionAccuracy::Eighths)
}

fn numeral_system_format() -> Result<TableCellNumeralSystemFormat, litchi_iwa::Error> {
    TableCellNumeralSystemFormat::new(
        TableCellNumeralSystemBase::HEXADECIMAL,
        TableCellNumeralSystemPlaces::Fixed(TableCellNumeralSystemFixedPlaces::EIGHT),
        TableCellNumeralSystemNegativeStyle::TwosComplement,
    )
}

fn date_time_format() -> TableCellDateTimeFormat {
    TableCellDateTimeFormat::iso_date_time_24_hour_with_seconds()
}

const fn duration_format() -> TableCellDurationFormat {
    TableCellDurationFormat::custom(
        TableCellDurationStyle::Abbreviated,
        TableCellDurationUnitRange::hours_to_milliseconds(),
    )
}

fn slider_format() -> Result<TableCellSliderFormat, litchi_iwa::Error> {
    Ok(TableCellSliderFormat::new(
        TableCellSliderRange::new(-10.0, 30.0, 0.5)?,
        TableCellNumberFormat::new(
            TableCellDecimalPlaces::fixed(2)?,
            TableCellNegativeNumberStyle::MinusSign,
            TableCellThousandsSeparator::Hidden,
        )
        .into(),
    ))
}

fn stepper_format() -> Result<TableCellStepperFormat, litchi_iwa::Error> {
    Ok(TableCellStepperFormat::new(
        TableCellStepperRange::new(-10.0, 30.0, 0.5)?,
        TableCellNumberFormat::new(
            TableCellDecimalPlaces::fixed(2)?,
            TableCellNegativeNumberStyle::MinusSign,
            TableCellThousandsSeparator::Hidden,
        )
        .into(),
    ))
}

fn create_numbers(output: &Path) -> Result<(), Box<dyn std::error::Error>> {
    let mut editor = NumbersDocumentBuilder::new()
        .table_name("Number Formats")
        .table_dimensions(3, 13)
        .build()?;
    let table_id = editor.tables()?.remove(0).object_id;
    editor.set_cell(
        table_id,
        ROW,
        NUMBER_COLUMN,
        CellValue::Number(NUMBER_VALUE),
    )?;
    editor.set_table_cell_number_format(table_id, ROW, NUMBER_COLUMN, format()?)?;
    editor.set_cell(
        table_id,
        ROW,
        PERCENTAGE_COLUMN,
        CellValue::Number(PERCENTAGE_VALUE),
    )?;
    editor.set_table_cell_percentage_format(
        table_id,
        ROW,
        PERCENTAGE_COLUMN,
        percentage_format()?,
    )?;
    editor.set_cell(
        table_id,
        ROW,
        CURRENCY_COLUMN,
        CellValue::Number(CURRENCY_VALUE),
    )?;
    editor.set_table_cell_currency_format(table_id, ROW, CURRENCY_COLUMN, currency_format()?)?;
    editor.set_cell(
        table_id,
        ROW,
        SCIENTIFIC_COLUMN,
        CellValue::Number(SCIENTIFIC_VALUE),
    )?;
    editor.set_table_cell_scientific_format(
        table_id,
        ROW,
        SCIENTIFIC_COLUMN,
        scientific_format()?,
    )?;
    editor.set_cell(
        table_id,
        ROW,
        FRACTION_COLUMN,
        CellValue::Number(FRACTION_VALUE),
    )?;
    editor.set_table_cell_fraction_format(table_id, ROW, FRACTION_COLUMN, fraction_format())?;
    editor.set_cell(
        table_id,
        ROW,
        NUMERAL_SYSTEM_COLUMN,
        CellValue::Number(NUMERAL_SYSTEM_VALUE),
    )?;
    editor.set_table_cell_numeral_system_format(
        table_id,
        ROW,
        NUMERAL_SYSTEM_COLUMN,
        numeral_system_format()?,
    )?;
    editor.set_cell(
        table_id,
        ROW,
        DATE_TIME_COLUMN,
        CellValue::Date(DATE_TIME_VALUE),
    )?;
    editor.set_table_cell_date_time_format(table_id, ROW, DATE_TIME_COLUMN, date_time_format())?;
    editor.set_cell(
        table_id,
        ROW,
        DURATION_COLUMN,
        CellValue::Duration(DURATION_VALUE),
    )?;
    editor.set_table_cell_duration_format(table_id, ROW, DURATION_COLUMN, duration_format())?;
    editor.set_cell(table_id, ROW, CHECKBOX_COLUMN, CellValue::Boolean(true))?;
    editor.set_table_cell_checkbox_format(
        table_id,
        ROW,
        CHECKBOX_COLUMN,
        TableCellCheckboxFormat,
    )?;
    editor.set_cell(
        table_id,
        ROW,
        STAR_RATING_COLUMN,
        CellValue::Number(STAR_RATING_VALUE),
    )?;
    editor.set_table_cell_star_rating_format(
        table_id,
        ROW,
        STAR_RATING_COLUMN,
        TableCellStarRatingFormat,
    )?;
    editor.set_cell(
        table_id,
        ROW,
        SLIDER_COLUMN,
        CellValue::Number(SLIDER_VALUE),
    )?;
    editor.set_table_cell_slider_format(table_id, ROW, SLIDER_COLUMN, slider_format()?)?;
    editor.set_cell(
        table_id,
        ROW,
        STEPPER_COLUMN,
        CellValue::Number(STEPPER_VALUE),
    )?;
    editor.set_table_cell_stepper_format(table_id, ROW, STEPPER_COLUMN, stepper_format()?)?;
    editor.save(output)?;
    Ok(())
}

fn create_pages(output: &Path) -> Result<(), Box<dyn std::error::Error>> {
    let mut editor = PagesDocumentBuilder::new()
        .body_text("Created from scratch with native table-cell formats.\n")
        .body_table("Number Formats", 3, 13)
        .build()?;
    let table_id = editor.tables()?.remove(0).model_object_id;
    editor.set_table_cell(
        table_id,
        ROW,
        NUMBER_COLUMN,
        CellValue::Number(NUMBER_VALUE),
    )?;
    editor.set_table_cell_number_format(table_id, ROW, NUMBER_COLUMN, format()?)?;
    editor.set_table_cell(
        table_id,
        ROW,
        PERCENTAGE_COLUMN,
        CellValue::Number(PERCENTAGE_VALUE),
    )?;
    editor.set_table_cell_percentage_format(
        table_id,
        ROW,
        PERCENTAGE_COLUMN,
        percentage_format()?,
    )?;
    editor.set_table_cell(
        table_id,
        ROW,
        CURRENCY_COLUMN,
        CellValue::Number(CURRENCY_VALUE),
    )?;
    editor.set_table_cell_currency_format(table_id, ROW, CURRENCY_COLUMN, currency_format()?)?;
    editor.set_table_cell(
        table_id,
        ROW,
        SCIENTIFIC_COLUMN,
        CellValue::Number(SCIENTIFIC_VALUE),
    )?;
    editor.set_table_cell_scientific_format(
        table_id,
        ROW,
        SCIENTIFIC_COLUMN,
        scientific_format()?,
    )?;
    editor.set_table_cell(
        table_id,
        ROW,
        FRACTION_COLUMN,
        CellValue::Number(FRACTION_VALUE),
    )?;
    editor.set_table_cell_fraction_format(table_id, ROW, FRACTION_COLUMN, fraction_format())?;
    editor.set_table_cell(
        table_id,
        ROW,
        NUMERAL_SYSTEM_COLUMN,
        CellValue::Number(NUMERAL_SYSTEM_VALUE),
    )?;
    editor.set_table_cell_numeral_system_format(
        table_id,
        ROW,
        NUMERAL_SYSTEM_COLUMN,
        numeral_system_format()?,
    )?;
    editor.set_table_cell(
        table_id,
        ROW,
        DATE_TIME_COLUMN,
        CellValue::Date(DATE_TIME_VALUE),
    )?;
    editor.set_table_cell_date_time_format(table_id, ROW, DATE_TIME_COLUMN, date_time_format())?;
    editor.set_table_cell(
        table_id,
        ROW,
        DURATION_COLUMN,
        CellValue::Duration(DURATION_VALUE),
    )?;
    editor.set_table_cell_duration_format(table_id, ROW, DURATION_COLUMN, duration_format())?;
    editor.set_table_cell(table_id, ROW, CHECKBOX_COLUMN, CellValue::Boolean(true))?;
    editor.set_table_cell_checkbox_format(
        table_id,
        ROW,
        CHECKBOX_COLUMN,
        TableCellCheckboxFormat,
    )?;
    editor.set_table_cell(
        table_id,
        ROW,
        STAR_RATING_COLUMN,
        CellValue::Number(STAR_RATING_VALUE),
    )?;
    editor.set_table_cell_star_rating_format(
        table_id,
        ROW,
        STAR_RATING_COLUMN,
        TableCellStarRatingFormat,
    )?;
    editor.set_table_cell(
        table_id,
        ROW,
        SLIDER_COLUMN,
        CellValue::Number(SLIDER_VALUE),
    )?;
    editor.set_table_cell_slider_format(table_id, ROW, SLIDER_COLUMN, slider_format()?)?;
    editor.set_table_cell(
        table_id,
        ROW,
        STEPPER_COLUMN,
        CellValue::Number(STEPPER_VALUE),
    )?;
    editor.set_table_cell_stepper_format(table_id, ROW, STEPPER_COLUMN, stepper_format()?)?;
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
        13,
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
        CellValue::Number(NUMBER_VALUE),
    )?;
    editor.set_slide_table_cell_number_format(
        0,
        table.model_object_id,
        ROW,
        NUMBER_COLUMN,
        format()?,
    )?;
    editor.set_slide_table_cell(
        0,
        table.model_object_id,
        ROW,
        PERCENTAGE_COLUMN,
        CellValue::Number(PERCENTAGE_VALUE),
    )?;
    editor.set_slide_table_cell_percentage_format(
        0,
        table.model_object_id,
        ROW,
        PERCENTAGE_COLUMN,
        percentage_format()?,
    )?;
    editor.set_slide_table_cell(
        0,
        table.model_object_id,
        ROW,
        CURRENCY_COLUMN,
        CellValue::Number(CURRENCY_VALUE),
    )?;
    editor.set_slide_table_cell_currency_format(
        0,
        table.model_object_id,
        ROW,
        CURRENCY_COLUMN,
        currency_format()?,
    )?;
    editor.set_slide_table_cell(
        0,
        table.model_object_id,
        ROW,
        SCIENTIFIC_COLUMN,
        CellValue::Number(SCIENTIFIC_VALUE),
    )?;
    editor.set_slide_table_cell_scientific_format(
        0,
        table.model_object_id,
        ROW,
        SCIENTIFIC_COLUMN,
        scientific_format()?,
    )?;
    editor.set_slide_table_cell(
        0,
        table.model_object_id,
        ROW,
        FRACTION_COLUMN,
        CellValue::Number(FRACTION_VALUE),
    )?;
    editor.set_slide_table_cell_fraction_format(
        0,
        table.model_object_id,
        ROW,
        FRACTION_COLUMN,
        fraction_format(),
    )?;
    editor.set_slide_table_cell(
        0,
        table.model_object_id,
        ROW,
        NUMERAL_SYSTEM_COLUMN,
        CellValue::Number(NUMERAL_SYSTEM_VALUE),
    )?;
    editor.set_slide_table_cell_numeral_system_format(
        0,
        table.model_object_id,
        ROW,
        NUMERAL_SYSTEM_COLUMN,
        numeral_system_format()?,
    )?;
    editor.set_slide_table_cell(
        0,
        table.model_object_id,
        ROW,
        DATE_TIME_COLUMN,
        CellValue::Date(DATE_TIME_VALUE),
    )?;
    editor.set_slide_table_cell_date_time_format(
        0,
        table.model_object_id,
        ROW,
        DATE_TIME_COLUMN,
        date_time_format(),
    )?;
    editor.set_slide_table_cell(
        0,
        table.model_object_id,
        ROW,
        DURATION_COLUMN,
        CellValue::Duration(DURATION_VALUE),
    )?;
    editor.set_slide_table_cell_duration_format(
        0,
        table.model_object_id,
        ROW,
        DURATION_COLUMN,
        duration_format(),
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
        TableCellCheckboxFormat,
    )?;
    editor.set_slide_table_cell(
        0,
        table.model_object_id,
        ROW,
        STAR_RATING_COLUMN,
        CellValue::Number(STAR_RATING_VALUE),
    )?;
    editor.set_slide_table_cell_star_rating_format(
        0,
        table.model_object_id,
        ROW,
        STAR_RATING_COLUMN,
        TableCellStarRatingFormat,
    )?;
    editor.set_slide_table_cell(
        0,
        table.model_object_id,
        ROW,
        SLIDER_COLUMN,
        CellValue::Number(SLIDER_VALUE),
    )?;
    editor.set_slide_table_cell_slider_format(
        0,
        table.model_object_id,
        ROW,
        SLIDER_COLUMN,
        slider_format()?,
    )?;
    editor.set_slide_table_cell(
        0,
        table.model_object_id,
        ROW,
        STEPPER_COLUMN,
        CellValue::Number(STEPPER_VALUE),
    )?;
    editor.set_slide_table_cell_stepper_format(
        0,
        table.model_object_id,
        ROW,
        STEPPER_COLUMN,
        stepper_format()?,
    )?;
    editor.save(output)?;
    Ok(())
}
