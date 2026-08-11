//! Create Pages, Numbers, and Keynote files with native conditional highlighting.

use std::path::{Path, PathBuf};

use litchi_iwa::keynote::{KeynoteDocumentBuilder, KeynoteEditor};
use litchi_iwa::numbers::{NumbersDocumentBuilder, NumbersEditor};
use litchi_iwa::pages::{PagesDocumentBuilder, PagesEditor};
use litchi_iwa::shapes::{DrawablePoint, DrawableSize, RgbColorSpace, RgbaColor};
use litchi_iwa_common::table::cell::conditional_highlight::{Condition, Rule, Style, Text};
use litchi_numbers::cell::Value as CellValue;

const HIGHLIGHT_ROW: usize = 1;
const HIGHLIGHT_COLUMN: usize = 1;
const SIGN_HIGHLIGHT_ROW: usize = 2;
const POSITIVE_HIGHLIGHT_COLUMN: usize = 1;
const NEGATIVE_HIGHLIGHT_COLUMN: usize = 2;
const POSITIVE_VALUE: f64 = 42.0;
const NEGATIVE_VALUE: f64 = -42.0;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let output = PathBuf::from(
        std::env::args()
            .nth(1)
            .ok_or("usage: create_iwork_conditional_highlights <output-directory>")?,
    );
    std::fs::create_dir_all(&output)?;
    create_numbers(&output.join("conditional-highlight.numbers"))?;
    create_pages(&output.join("conditional-highlight.pages"))?;
    create_keynote(&output.join("conditional-highlight.key"))?;
    Ok(())
}

fn highlight_rules() -> Result<Vec<Rule>, Box<dyn std::error::Error>> {
    let style = highlight_style()?;
    let text = |value| Text::new(value);
    Ok([
        Condition::CellIsBlank,
        Condition::CellIsNotBlank,
        Condition::TextEqualTo(text("organic grain")?),
        Condition::TextNotEqualTo(text("dairy")?),
        Condition::TextStartsWith(text("organic")?),
        Condition::TextDoesNotStartWith(text("dairy")?),
        Condition::TextEndsWith(text("grain")?),
        Condition::TextDoesNotEndWith(text("rice")?),
        Condition::TextContains(text("nic gr")?),
        Condition::TextDoesNotContain(text("rice")?),
    ]
    .into_iter()
    .map(|condition| Rule::new(condition, style))
    .collect())
}

fn highlight_style() -> Result<Style, Box<dyn std::error::Error>> {
    let red = RgbaColor::new(0.96, 0.22, 0.18, 1.0, RgbColorSpace::Srgb)?;
    Ok(Style::new(Some(red), None, true)?)
}

fn highlight_rule(condition: Condition) -> Result<Rule, Box<dyn std::error::Error>> {
    Ok(Rule::new(condition, highlight_style()?))
}

fn create_numbers(output: &Path) -> Result<(), Box<dyn std::error::Error>> {
    let mut editor = NumbersDocumentBuilder::new()
        .table_name("Conditional")
        .table_dimensions(3, 3)
        .build()?;
    let table_id = editor.tables()?.remove(0).id();
    set_numbers_cells(
        &mut editor,
        [
            litchi_numbers::cell::Update::new(
                HIGHLIGHT_ROW,
                HIGHLIGHT_COLUMN,
                CellValue::Text("Organic Grain".to_owned()),
            ),
            litchi_numbers::cell::Update::new(
                SIGN_HIGHLIGHT_ROW,
                POSITIVE_HIGHLIGHT_COLUMN,
                CellValue::number(POSITIVE_VALUE)?,
            ),
            litchi_numbers::cell::Update::new(
                SIGN_HIGHLIGHT_ROW,
                NEGATIVE_HIGHLIGHT_COLUMN,
                CellValue::number(NEGATIVE_VALUE)?,
            ),
        ],
    )?;
    let rules = highlight_rules()?;
    editor.set_cell_conditional_highlighting(table_id, HIGHLIGHT_ROW, HIGHLIGHT_COLUMN, &rules)?;
    let positive_rule = highlight_rule(Condition::NumberIsPositive)?;
    editor.set_cell_conditional_highlighting(
        table_id,
        SIGN_HIGHLIGHT_ROW,
        POSITIVE_HIGHLIGHT_COLUMN,
        std::slice::from_ref(&positive_rule),
    )?;
    let negative_rule = highlight_rule(Condition::NumberIsNegative)?;
    editor.set_cell_conditional_highlighting(
        table_id,
        SIGN_HIGHLIGHT_ROW,
        NEGATIVE_HIGHLIGHT_COLUMN,
        std::slice::from_ref(&negative_rule),
    )?;
    editor.save(output)?;
    assert_eq!(
        NumbersEditor::open(output)?.cell_conditional_highlight_rules(
            table_id,
            HIGHLIGHT_ROW,
            HIGHLIGHT_COLUMN,
        )?,
        Some(rules)
    );
    assert_eq!(
        NumbersEditor::open(output)?.cell_conditional_highlight_rules(
            table_id,
            SIGN_HIGHLIGHT_ROW,
            POSITIVE_HIGHLIGHT_COLUMN,
        )?,
        Some(vec![positive_rule])
    );
    assert_eq!(
        NumbersEditor::open(output)?.cell_conditional_highlight_rules(
            table_id,
            SIGN_HIGHLIGHT_ROW,
            NEGATIVE_HIGHLIGHT_COLUMN,
        )?,
        Some(vec![negative_rule])
    );
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
        .body_text("Conditional highlighting created from scratch.\n")
        .body_table("Conditional", 3, 3)
        .build()?;
    let table_id = editor.tables()?.remove(0).model_object_id;
    editor.set_table_cell(
        table_id,
        HIGHLIGHT_ROW,
        HIGHLIGHT_COLUMN,
        CellValue::Text("Organic Grain".to_owned()),
    )?;
    editor.set_table_cell(
        table_id,
        SIGN_HIGHLIGHT_ROW,
        POSITIVE_HIGHLIGHT_COLUMN,
        CellValue::number(POSITIVE_VALUE)?,
    )?;
    editor.set_table_cell(
        table_id,
        SIGN_HIGHLIGHT_ROW,
        NEGATIVE_HIGHLIGHT_COLUMN,
        CellValue::number(NEGATIVE_VALUE)?,
    )?;
    let rules = highlight_rules()?;
    editor.set_table_cell_conditional_highlighting(
        table_id,
        HIGHLIGHT_ROW,
        HIGHLIGHT_COLUMN,
        &rules,
    )?;
    let positive_rule = highlight_rule(Condition::NumberIsPositive)?;
    editor.set_table_cell_conditional_highlighting(
        table_id,
        SIGN_HIGHLIGHT_ROW,
        POSITIVE_HIGHLIGHT_COLUMN,
        std::slice::from_ref(&positive_rule),
    )?;
    let negative_rule = highlight_rule(Condition::NumberIsNegative)?;
    editor.set_table_cell_conditional_highlighting(
        table_id,
        SIGN_HIGHLIGHT_ROW,
        NEGATIVE_HIGHLIGHT_COLUMN,
        std::slice::from_ref(&negative_rule),
    )?;
    editor.save(output)?;
    assert_eq!(
        PagesEditor::open(output)?.table_cell_conditional_highlight_rules(
            table_id,
            HIGHLIGHT_ROW,
            HIGHLIGHT_COLUMN,
        )?,
        Some(rules)
    );
    assert_eq!(
        PagesEditor::open(output)?.table_cell_conditional_highlight_rules(
            table_id,
            SIGN_HIGHLIGHT_ROW,
            POSITIVE_HIGHLIGHT_COLUMN,
        )?,
        Some(vec![positive_rule])
    );
    assert_eq!(
        PagesEditor::open(output)?.table_cell_conditional_highlight_rules(
            table_id,
            SIGN_HIGHLIGHT_ROW,
            NEGATIVE_HIGHLIGHT_COLUMN,
        )?,
        Some(vec![negative_rule])
    );
    Ok(())
}

fn create_keynote(output: &Path) -> Result<(), Box<dyn std::error::Error>> {
    let mut editor = KeynoteDocumentBuilder::new()
        .title("Conditional highlighting")
        .build()?;
    let table = editor.add_slide_table(
        0,
        "Conditional",
        3,
        3,
        DrawablePoint { x: 320.0, y: 360.0 },
        DrawableSize {
            width: 1_280.0,
            height: 480.0,
        },
    )?;
    editor.set_slide_table_cell(
        0,
        table.model_object_id,
        HIGHLIGHT_ROW,
        HIGHLIGHT_COLUMN,
        CellValue::Text("Organic Grain".to_owned()),
    )?;
    editor.set_slide_table_cell(
        0,
        table.model_object_id,
        SIGN_HIGHLIGHT_ROW,
        POSITIVE_HIGHLIGHT_COLUMN,
        CellValue::number(POSITIVE_VALUE)?,
    )?;
    editor.set_slide_table_cell(
        0,
        table.model_object_id,
        SIGN_HIGHLIGHT_ROW,
        NEGATIVE_HIGHLIGHT_COLUMN,
        CellValue::number(NEGATIVE_VALUE)?,
    )?;
    let rules = highlight_rules()?;
    editor.set_slide_table_cell_conditional_highlighting(
        0,
        table.model_object_id,
        HIGHLIGHT_ROW,
        HIGHLIGHT_COLUMN,
        &rules,
    )?;
    let positive_rule = highlight_rule(Condition::NumberIsPositive)?;
    editor.set_slide_table_cell_conditional_highlighting(
        0,
        table.model_object_id,
        SIGN_HIGHLIGHT_ROW,
        POSITIVE_HIGHLIGHT_COLUMN,
        std::slice::from_ref(&positive_rule),
    )?;
    let negative_rule = highlight_rule(Condition::NumberIsNegative)?;
    editor.set_slide_table_cell_conditional_highlighting(
        0,
        table.model_object_id,
        SIGN_HIGHLIGHT_ROW,
        NEGATIVE_HIGHLIGHT_COLUMN,
        std::slice::from_ref(&negative_rule),
    )?;
    editor.save(output)?;
    assert_eq!(
        KeynoteEditor::open(output)?.slide_table_cell_conditional_highlight_rules(
            0,
            table.model_object_id,
            HIGHLIGHT_ROW,
            HIGHLIGHT_COLUMN,
        )?,
        Some(rules)
    );
    assert_eq!(
        KeynoteEditor::open(output)?.slide_table_cell_conditional_highlight_rules(
            0,
            table.model_object_id,
            SIGN_HIGHLIGHT_ROW,
            POSITIVE_HIGHLIGHT_COLUMN,
        )?,
        Some(vec![positive_rule])
    );
    assert_eq!(
        KeynoteEditor::open(output)?.slide_table_cell_conditional_highlight_rules(
            0,
            table.model_object_id,
            SIGN_HIGHLIGHT_ROW,
            NEGATIVE_HIGHLIGHT_COLUMN,
        )?,
        Some(vec![negative_rule])
    );
    Ok(())
}
