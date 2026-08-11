//! Create Pages, Numbers, and Keynote files with native date-period highlighting.

use std::path::{Path, PathBuf};

use chrono::{Days, Local, Months, NaiveDate};
use litchi_iwa::keynote::{KeynoteDocumentBuilder, KeynoteEditor};
use litchi_iwa::numbers::{NumbersDocumentBuilder, NumbersEditor};
use litchi_iwa::pages::{PagesDocumentBuilder, PagesEditor};
use litchi_iwa::shapes::{DrawablePoint, DrawableSize, RgbColorSpace, RgbaColor};
use litchi_iwa_common::table::cell::conditional_highlight::{
    Condition, Offset, OffsetDirection, Period, PeriodUnit, Rule, Style,
};
use litchi_numbers::cell::Value as CellValue;

const DATE_ROW: usize = 1;
const FIRST_DATE_COLUMN: usize = 1;
const APPLE_EPOCH_YEAR: i32 = 2001;
const SECONDS_PER_DAY: f64 = 86_400.0;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let output = PathBuf::from(
        std::env::args()
            .nth(1)
            .ok_or("usage: create_iwork_date_period_conditional_highlights <output-directory>")?,
    );
    std::fs::create_dir_all(&output)?;
    create_numbers(&output.join("date-period-highlight.numbers"))?;
    create_pages(&output.join("date-period-highlight.pages"))?;
    create_keynote(&output.join("date-period-highlight.key"))?;
    Ok(())
}

fn date_cases() -> Result<[(CellValue, Rule); 4], Box<dyn std::error::Error>> {
    use OffsetDirection as Direction;
    use PeriodUnit as Unit;

    let today = Local::now().date_naive();
    let two_days = Period::new(2, Unit::Days)?;
    let two_weeks = Period::new(2, Unit::Weeks)?;
    let one_month = Period::new(1, Unit::Months)?;
    let one_quarter = Period::new(1, Unit::Quarters)?;
    let fill = RgbaColor::new(0.96, 0.22, 0.18, 1.0, RgbColorSpace::Srgb)?;
    let style = Style::new(Some(fill), None, true)?;
    let case = |date, condition| -> Result<_, litchi_numbers::cell::FiniteF64Error> {
        Ok((
            CellValue::date(date_in_apple_seconds(date) + SECONDS_PER_DAY / 2.0)?,
            Rule::new(condition, style),
        ))
    };
    Ok([
        case(
            today
                .checked_add_days(Days::new(2))
                .ok_or("date overflow")?,
            Condition::DateIsInNext(two_days),
        )?,
        case(
            today
                .checked_sub_days(Days::new(14))
                .ok_or("date overflow")?,
            Condition::DateIsInLast(two_weeks),
        )?,
        case(
            today
                .checked_add_months(Months::new(1))
                .ok_or("date overflow")?,
            Condition::DateIsOffsetFromToday(Offset::new(one_month, Direction::FromNow)),
        )?,
        case(
            today
                .checked_sub_months(Months::new(3))
                .ok_or("date overflow")?,
            Condition::DateIsOffsetFromToday(Offset::new(one_quarter, Direction::Ago)),
        )?,
    ])
}

fn date_in_apple_seconds(date: NaiveDate) -> f64 {
    let epoch = NaiveDate::from_ymd_opt(APPLE_EPOCH_YEAR, 1, 1)
        .expect("the Apple epoch is a valid calendar date");
    date.signed_duration_since(epoch).num_days() as f64 * SECONDS_PER_DAY
}

fn create_numbers(output: &Path) -> Result<(), Box<dyn std::error::Error>> {
    let mut editor = NumbersDocumentBuilder::new()
        .table_name("Date periods")
        .table_dimensions(2, 5)
        .build()?;
    let table_id = editor.tables()?.remove(0).id();
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
        .body_text("Date-period conditional highlighting created from scratch.\n")
        .body_table("Date periods", 2, 5)
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
        .title("Date-period conditional highlighting")
        .build()?;
    let table = editor.add_slide_table(
        0,
        "Date periods",
        2,
        5,
        DrawablePoint { x: 240.0, y: 360.0 },
        DrawableSize {
            width: 1_440.0,
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
