//! Print the sheet and table hierarchy resolved from a Numbers root document.

use std::env;

use litchi_iwa::numbers::NumbersEditor;
use litchi_numbers::{Document, cell::Value as CellValue};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let path = arguments
        .next()
        .ok_or("usage: inspect_numbers_document <file.numbers> [--cells]")?;
    let show_cells = arguments.any(|argument| argument == "--cells");
    let document = Document::open(&path)?;
    for sheet in document.sheets() {
        println!("sheet {}: {:?}", sheet.index(), sheet.name());
        for table in sheet.tables() {
            println!(
                "  table {:?}: {} rows x {} columns",
                table.name(),
                table.row_count(),
                table.column_count()
            );
            if show_cells {
                for cell in table.iter_cells() {
                    let position = cell.position();
                    let value = cell.value();
                    if !matches!(value, CellValue::Empty) {
                        println!("    ({}, {}) {value:?}", position.row(), position.column());
                    }
                }
                continue;
            }
            let formulas = table.iter_cells().filter_map(|cell| match cell.value() {
                CellValue::Formula(formula) => Some((cell.position(), formula)),
                _ => None,
            });
            for (position, formula) in formulas {
                println!("    ({}, {}) {formula}", position.row(), position.column());
            }
        }
    }
    let editor = NumbersEditor::open(path)?;
    for table in editor.tables()? {
        println!(
            "table id={} name={:?} dimensions={}x{}",
            table.id(),
            table.name,
            table.rows,
            table.columns
        );
    }
    for category in editor.pivot_categories()? {
        println!(
            "pivot category label={:?} group_by={:?} column={:?} group={:?} aggregate_type={} level={}",
            category.label,
            category.reference.group_by_uid,
            category.reference.column_uid,
            category.reference.group_uid,
            category.reference.aggregate_type,
            category.reference.group_level,
        );
    }
    Ok(())
}
