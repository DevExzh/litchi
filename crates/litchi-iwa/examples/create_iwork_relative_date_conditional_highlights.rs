//! Create Pages, Numbers, and Keynote files with native relative-date highlighting.

use std::path::{Path, PathBuf};

use chrono::{Local, NaiveDate};
use litchi_iwa::keynote::{KeynoteDocumentBuilder, KeynoteEditor};
use litchi_iwa::numbers::{NumbersDocumentBuilder, NumbersEditor};
use litchi_iwa::pages::{PagesDocumentBuilder, PagesEditor};
use litchi_iwa::shapes::{DrawablePoint, DrawableSize, RgbColorSpace, RgbaColor};
use litchi_iwa_common::table::cell::conditional_highlight::{Condition, Rule, Style};
use litchi_numbers::cell::Value as CellValue;

const DATE_ROW: usize = 1;
const FIRST_DATE_COLUMN: usize = 1;
const APPLE_EPOCH_YEAR: i32 = 2001;
const SECONDS_PER_DAY: f64 = 86_400.0;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let output =
        PathBuf::from(std::env::args().nth(1).ok_or(
            "usage: create_iwork_relative_date_conditional_highlights <output-directory>",
        )?);
    std::fs::create_dir_all(&output)?;
    create_numbers(&output.join("relative-date-highlight.numbers"))?;
    create_pages(&output.join("relative-date-highlight.pages"))?;
    create_keynote(&output.join("relative-date-highlight.key"))?;
    Ok(())
}

fn date_cases() -> Result<[(CellValue, Rule); 3], Box<dyn std::error::Error>> {
    let today = local_today_in_apple_seconds();
    let midday = SECONDS_PER_DAY / 2.0;
    let fill = RgbaColor::new(0.96, 0.22, 0.18, 1.0, RgbColorSpace::Srgb)?;
    let style = Style::new(Some(fill), None, true)?;
    Ok([
        (
            CellValue::date(today - SECONDS_PER_DAY + midday)?,
            Rule::new(Condition::DateIsYesterday, style),
        ),
        (
            CellValue::date(today + midday)?,
            Rule::new(Condition::DateIsToday, style),
        ),
        (
            CellValue::date(today + SECONDS_PER_DAY + midday)?,
            Rule::new(Condition::DateIsTomorrow, style),
        ),
    ])
}

fn local_today_in_apple_seconds() -> f64 {
    let epoch = NaiveDate::from_ymd_opt(APPLE_EPOCH_YEAR, 1, 1)
        .expect("the Apple epoch is a valid calendar date");
    Local::now()
        .date_naive()
        .signed_duration_since(epoch)
        .num_days() as f64
        * SECONDS_PER_DAY
}

fn create_numbers(output: &Path) -> Result<(), Box<dyn std::error::Error>> {
    let mut editor = NumbersDocumentBuilder::new()
        .table_name("Relative dates")
        .table_dimensions(2, 4)
        .build()?;
    let table_id = editor.tables()?.remove(0).id();
    let cases = date_cases()?;
    set_numbers_cells(
        &mut editor,
        cases.iter().enumerate().map(|(offset, (value, _))| {
            litchi_numbers::cell::Update::new(DATE_ROW, FIRST_DATE_COLUMN + offset, value.clone())
        }),
    )?;
    for (offset, (_, rule)) in cases.iter().enumerate() {
        let column = FIRST_DATE_COLUMN + offset;
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

fn create_pages(output: &Path) -> Result<(), Box<dyn std::error::Error>> {
    let mut editor = PagesDocumentBuilder::new()
        .body_text("Relative-date conditional highlighting created from scratch.\n")
        .body_table("Relative dates", 2, 4)
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
        .title("Relative-date conditional highlighting")
        .build()?;
    let table = editor.add_slide_table(
        0,
        "Relative dates",
        2,
        4,
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
