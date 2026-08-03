//! Print the sheet and table hierarchy resolved from a Numbers root document.

use std::env;

use litchi_iwa::numbers::{CellValue, NumbersDocument, NumbersEditor};
use prost::Message;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let path = arguments
        .next()
        .ok_or("usage: inspect_numbers_document <file.numbers> [--cells]")?;
    let show_cells = arguments.any(|argument| argument == "--cells");
    let document = NumbersDocument::open(&path)?;
    if let Some(root) = document
        .bundle()
        .get_archive("Index/Document.iwa")
        .and_then(|archive| archive.object(1))
        .and_then(|object| object.messages.first())
        .and_then(|message| {
            litchi_iwa::protobuf::tn::DocumentArchive::decode(message.data.as_slice()).ok()
        })
    {
        println!(
            "document sheets: {:?}",
            root.sheets
                .iter()
                .map(|reference| reference.identifier)
                .collect::<Vec<_>>()
        );
    }
    for sheet in document.sheets()? {
        println!("sheet {}: {:?}", sheet.index, sheet.name);
        for table in sheet.tables {
            println!(
                "  table {:?}: {} rows x {} columns",
                table.name(),
                table.row_count(),
                table.column_count()
            );
            let mut comments = table.iter_comments().collect::<Vec<_>>();
            comments.sort_by_key(|((row, column), _)| (*row, *column));
            for ((row, column), comment) in comments {
                println!("    comment ({row}, {column}) {:?}", comment.text);
            }
            if show_cells {
                let mut cells = table.iter_cells().collect::<Vec<_>>();
                cells.sort_by_key(|((row, column), _)| (*row, *column));
                for ((row, column), value) in cells {
                    if !matches!(value, CellValue::Empty) {
                        println!("    ({row}, {column}) {value:?}");
                    }
                }
                continue;
            }
            let mut formulas = table
                .iter_cells()
                .filter_map(|((row, column), value)| match value {
                    CellValue::Formula(formula) => Some((row, column, formula)),
                    _ => None,
                })
                .collect::<Vec<_>>();
            formulas.sort_by_key(|(row, column, _)| (*row, *column));
            for (row, column, formula) in formulas {
                println!("    ({row}, {column}) {formula}");
            }
        }
    }
    let editor = NumbersEditor::open(path)?;
    for table in editor.tables()? {
        println!(
            "table id={} name={:?} dimensions={}x{}",
            table.object_id, table.name, table.rows, table.columns
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
