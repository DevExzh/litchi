//! Create Pages, Numbers, and Keynote files with native fixed-date highlighting.

use std::path::{Path, PathBuf};

use litchi_iwa::keynote::{KeynoteDocumentBuilder, KeynoteEditor};
use litchi_iwa::numbers::{NumbersDocumentBuilder, NumbersEditor};
use litchi_iwa::pages::{PagesDocumentBuilder, PagesEditor};
use litchi_iwa::shapes::{DrawablePoint, DrawableSize, RgbColorSpace, RgbaColor};
use litchi_iwa_common::table::cell::conditional_highlight::{
    Condition, Date,
    DateRange, Rule,
    Style,
};
use litchi_numbers::cell::Value as CellValue;

const DATE_ROW: usize = 1;
const FIRST_DATE_COLUMN: usize = 1;
const SECONDS_PER_DAY: f64 = 86_400.0;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let output = PathBuf::from(
        std::env::args()
            .nth(1)
            .ok_or("usage: create_iwork_fixed_date_conditional_highlights <output-directory>")?,
    );
    std::fs::create_dir_all(&output)?;
    create_numbers(&output.join("fixed-date-highlight.numbers"))?;
    create_pages(&output.join("fixed-date-highlight.pages"))?;
    create_keynote(&output.join("fixed-date-highlight.key"))?;
    Ok(())
}

fn date_cases()
-> Result<[(CellValue, Rule); 4], Box<dyn std::error::Error>> {
    let lower = Date::from_ymd(2026, 7, 26)?;
    let exact = Date::from_ymd(2026, 7, 27)?;
    let upper = Date::from_ymd(2026, 7, 28)?;
    let range = DateRange::new(lower, upper)?;
    let midday = SECONDS_PER_DAY / 2.0;
    let fill = RgbaColor::new(0.96, 0.22, 0.18, 1.0, RgbColorSpace::Srgb)?;
    let style = Style::new(Some(fill), None, true)?;
    Ok([
        (
            CellValue::Date(exact.apple_seconds() + midday),
            Rule::new(
                Condition::DateIs(exact),
                style,
            ),
        ),
        (
            CellValue::Date(lower.apple_seconds() + midday),
            Rule::new(
                Condition::DateIsBefore(exact),
                style,
            ),
        ),
        (
            CellValue::Date(upper.apple_seconds() + midday),
            Rule::new(
                Condition::DateIsAfter(exact),
                style,
            ),
        ),
        (
            CellValue::Date(exact.apple_seconds() + midday),
            Rule::new(
                Condition::DateIsBetween(range),
                style,
            ),
        ),
    ])
}

fn create_numbers(output: &Path) -> Result<(), Box<dyn std::error::Error>> {
    let mut editor = NumbersDocumentBuilder::new()
        .table_name("Fixed dates")
        .table_dimensions(2, 5)
        .build()?;
    let table_id = editor.tables()?.remove(0).object_id;
    let cases = date_cases()?;
    for (offset, (value, rule)) in cases.iter().enumerate() {
        let column = FIRST_DATE_COLUMN + offset;
        editor.set_cell(table_id, DATE_ROW, column, value.clone())?;
        editor.set_cell_conditional_highlighting(
            table_id,
            DATE_ROW,
            column,
            std::slice::from_ref(rule),
        )?;
    }
    editor.save(output)?;

    let reopened = NumbersEditor::open(output)?;
    for (offset, (_, rule)) in cases.iter().enumerate() {
        assert_eq!(
            reopened.cell_conditional_highlight_rules(
                table_id,
                DATE_ROW,
                FIRST_DATE_COLUMN + offset,
            )?,
            Some(vec![rule.clone()])
        );
    }
    Ok(())
}

fn create_pages(output: &Path) -> Result<(), Box<dyn std::error::Error>> {
    let mut editor = PagesDocumentBuilder::new()
        .body_text("Fixed-date conditional highlighting created from scratch.\n")
        .body_table("Fixed dates", 2, 5)
        .build()?;
    let table_id = editor.tables()?.remove(0).model_object_id;
    let cases = date_cases()?;
    for (offset, (value, rule)) in cases.iter().enumerate() {
        let column = FIRST_DATE_COLUMN + offset;
        editor.set_table_cell(table_id, DATE_ROW, column, value.clone())?;
        editor.set_table_cell_conditional_highlighting(
            table_id,
            DATE_ROW,
            column,
            std::slice::from_ref(rule),
        )?;
    }
    editor.save(output)?;

    let reopened = PagesEditor::open(output)?;
    for (offset, (_, rule)) in cases.iter().enumerate() {
        assert_eq!(
            reopened.table_cell_conditional_highlight_rules(
                table_id,
                DATE_ROW,
                FIRST_DATE_COLUMN + offset,
            )?,
            Some(vec![rule.clone()])
        );
    }
    Ok(())
}

fn create_keynote(output: &Path) -> Result<(), Box<dyn std::error::Error>> {
    let mut editor = KeynoteDocumentBuilder::new()
        .title("Fixed-date conditional highlighting")
        .build()?;
    let table = editor.add_slide_table(
        0,
        "Fixed dates",
        2,
        5,
        DrawablePoint { x: 320.0, y: 360.0 },
        DrawableSize {
            width: 1_280.0,
            height: 360.0,
        },
    )?;
    let cases = date_cases()?;
    for (offset, (value, rule)) in cases.iter().enumerate() {
        let column = FIRST_DATE_COLUMN + offset;
        editor.set_slide_table_cell(0, table.model_object_id, DATE_ROW, column, value.clone())?;
        editor.set_slide_table_cell_conditional_highlighting(
            0,
            table.model_object_id,
            DATE_ROW,
            column,
            std::slice::from_ref(rule),
        )?;
    }
    editor.save(output)?;

    let reopened = KeynoteEditor::open(output)?;
    for (offset, (_, rule)) in cases.iter().enumerate() {
        assert_eq!(
            reopened.slide_table_cell_conditional_highlight_rules(
                0,
                table.model_object_id,
                DATE_ROW,
                FIRST_DATE_COLUMN + offset,
            )?,
            Some(vec![rule.clone()])
        );
    }
    Ok(())
}
