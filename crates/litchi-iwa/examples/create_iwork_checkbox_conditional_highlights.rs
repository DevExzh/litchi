//! Create Pages, Numbers, and Keynote files with native Checkbox highlighting.

use std::path::{Path, PathBuf};

use litchi_iwa::keynote::{KeynoteDocumentBuilder, KeynoteEditor};
use litchi_iwa::numbers::{NumbersDocumentBuilder, NumbersEditor};
use litchi_iwa::pages::{PagesDocumentBuilder, PagesEditor};
use litchi_iwa::shapes::{DrawablePoint, DrawableSize, RgbColorSpace, RgbaColor};
use litchi_iwa_common::table::cell::conditional_highlight::{Condition, Rule, Style};
use litchi_numbers::cell::Value as CellValue;
use litchi_numbers::cell::data_format::Checkbox;

const CHECKBOX_ROW: usize = 1;
const CHECKED_COLUMN: usize = 1;
const UNCHECKED_COLUMN: usize = 2;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let output = PathBuf::from(
        std::env::args()
            .nth(1)
            .ok_or("usage: create_iwork_checkbox_conditional_highlights <output-directory>")?,
    );
    std::fs::create_dir_all(&output)?;
    create_numbers(&output.join("checkbox-highlight.numbers"))?;
    create_pages(&output.join("checkbox-highlight.pages"))?;
    create_keynote(&output.join("checkbox-highlight.key"))?;
    Ok(())
}

fn checkbox_rules() -> Result<[Rule; 2], Box<dyn std::error::Error>> {
    let red = RgbaColor::new(0.96, 0.22, 0.18, 1.0, RgbColorSpace::Srgb)?;
    let style = Style::new(Some(red), None, true)?;
    Ok([
        Rule::new(Condition::CheckboxIsChecked, style),
        Rule::new(Condition::CheckboxIsNotChecked, style),
    ])
}

fn create_numbers(output: &Path) -> Result<(), Box<dyn std::error::Error>> {
    let mut editor = NumbersDocumentBuilder::new()
        .table_name("Checkboxes")
        .table_dimensions(2, 3)
        .build()?;
    let table_id = editor.tables()?.remove(0).object_id;
    for (column, value) in [(CHECKED_COLUMN, true), (UNCHECKED_COLUMN, false)] {
        editor.set_cell(table_id, CHECKBOX_ROW, column, CellValue::Boolean(value))?;
        editor.set_table_cell_checkbox_format(table_id, CHECKBOX_ROW, column, Checkbox)?;
    }
    let rules = checkbox_rules()?;
    for (column, rule) in [CHECKED_COLUMN, UNCHECKED_COLUMN].into_iter().zip(&rules) {
        editor.set_cell_conditional_highlighting(
            table_id,
            CHECKBOX_ROW,
            column,
            std::slice::from_ref(rule),
        )?;
    }
    editor.save(output)?;

    let reopened = NumbersEditor::open(output)?;
    for (column, rule) in [CHECKED_COLUMN, UNCHECKED_COLUMN].into_iter().zip(&rules) {
        assert_eq!(
            reopened.cell_conditional_highlight_rules(table_id, CHECKBOX_ROW, column)?,
            Some(vec![rule.clone()])
        );
    }
    Ok(())
}

fn create_pages(output: &Path) -> Result<(), Box<dyn std::error::Error>> {
    let mut editor = PagesDocumentBuilder::new()
        .body_text("Checkbox conditional highlighting created from scratch.\n")
        .body_table("Checkboxes", 2, 3)
        .build()?;
    let table_id = editor.tables()?.remove(0).model_object_id;
    for (column, value) in [(CHECKED_COLUMN, true), (UNCHECKED_COLUMN, false)] {
        editor.set_table_cell(table_id, CHECKBOX_ROW, column, CellValue::Boolean(value))?;
        editor.set_table_cell_checkbox_format(
            table_id,
            CHECKBOX_ROW,
            column,
            Checkbox,
        )?;
    }
    let rules = checkbox_rules()?;
    for (column, rule) in [CHECKED_COLUMN, UNCHECKED_COLUMN].into_iter().zip(&rules) {
        editor.set_table_cell_conditional_highlighting(
            table_id,
            CHECKBOX_ROW,
            column,
            std::slice::from_ref(rule),
        )?;
    }
    editor.save(output)?;

    let reopened = PagesEditor::open(output)?;
    for (column, rule) in [CHECKED_COLUMN, UNCHECKED_COLUMN].into_iter().zip(&rules) {
        assert_eq!(
            reopened.table_cell_conditional_highlight_rules(table_id, CHECKBOX_ROW, column)?,
            Some(vec![rule.clone()])
        );
    }
    Ok(())
}

fn create_keynote(output: &Path) -> Result<(), Box<dyn std::error::Error>> {
    let mut editor = KeynoteDocumentBuilder::new()
        .title("Checkbox conditional highlighting")
        .build()?;
    let table = editor.add_slide_table(
        0,
        "Checkboxes",
        2,
        3,
        DrawablePoint { x: 320.0, y: 360.0 },
        DrawableSize {
            width: 1_280.0,
            height: 360.0,
        },
    )?;
    for (column, value) in [(CHECKED_COLUMN, true), (UNCHECKED_COLUMN, false)] {
        editor.set_slide_table_cell(
            0,
            table.model_object_id,
            CHECKBOX_ROW,
            column,
            CellValue::Boolean(value),
        )?;
        editor.set_slide_table_cell_checkbox_format(
            0,
            table.model_object_id,
            CHECKBOX_ROW,
            column,
            Checkbox,
        )?;
    }
    let rules = checkbox_rules()?;
    for (column, rule) in [CHECKED_COLUMN, UNCHECKED_COLUMN].into_iter().zip(&rules) {
        editor.set_slide_table_cell_conditional_highlighting(
            0,
            table.model_object_id,
            CHECKBOX_ROW,
            column,
            std::slice::from_ref(rule),
        )?;
    }
    editor.save(output)?;

    let reopened = KeynoteEditor::open(output)?;
    for (column, rule) in [CHECKED_COLUMN, UNCHECKED_COLUMN].into_iter().zip(&rules) {
        assert_eq!(
            reopened.slide_table_cell_conditional_highlight_rules(
                0,
                table.model_object_id,
                CHECKBOX_ROW,
                column,
            )?,
            Some(vec![rule.clone()])
        );
    }
    Ok(())
}
