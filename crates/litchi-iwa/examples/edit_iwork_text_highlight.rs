//! List, create, move, or delete native plain-text highlights in any iWork package.

use std::env;

use litchi_iwa::text::{IWorkTextEditor, TextHighlightId, TextRange};

const USAGE: &str = "usage:\n  edit_iwork_text_highlight list <input> <storage-id>\n  edit_iwork_text_highlight add <input> <output> <storage-id> <start> <end>\n  edit_iwork_text_highlight update <input> <output> <storage-id> <highlight-id> <start> <end>\n  edit_iwork_text_highlight remove <input> <output> <storage-id> <highlight-id>";

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let command = arguments.next().ok_or(USAGE)?;
    match command.as_str() {
        "list" => {
            let input = arguments.next().ok_or(USAGE)?;
            let storage_id = parse_u64(arguments.next(), "storage-id")?;
            require_end(arguments)?;
            let editor = IWorkTextEditor::open(input)?;
            for highlight in editor.text_highlights(storage_id)? {
                println!(
                    "id={} range={}..{}",
                    highlight.id.object_id(),
                    highlight.range.start().utf16_index(),
                    highlight.range.end().utf16_index(),
                );
            }
        },
        "add" => {
            let input = arguments.next().ok_or(USAGE)?;
            let output = arguments.next().ok_or(USAGE)?;
            let storage_id = parse_u64(arguments.next(), "storage-id")?;
            let range = parse_range(&mut arguments)?;
            require_end(arguments)?;
            let mut editor = IWorkTextEditor::open(input)?;
            let highlight = editor.add_text_highlight(storage_id, range)?;
            editor.save(output)?;
            println!("created highlight {}", highlight.id.object_id());
        },
        "update" => {
            let input = arguments.next().ok_or(USAGE)?;
            let output = arguments.next().ok_or(USAGE)?;
            let storage_id = parse_u64(arguments.next(), "storage-id")?;
            let highlight_id =
                TextHighlightId::from_object_id(parse_u64(arguments.next(), "highlight-id")?)?;
            let range = parse_range(&mut arguments)?;
            require_end(arguments)?;
            let mut editor = IWorkTextEditor::open(input)?;
            editor.update_text_highlight(storage_id, highlight_id, range)?;
            editor.save(output)?;
            println!("updated highlight {}", highlight_id.object_id());
        },
        "remove" => {
            let input = arguments.next().ok_or(USAGE)?;
            let output = arguments.next().ok_or(USAGE)?;
            let storage_id = parse_u64(arguments.next(), "storage-id")?;
            let highlight_id =
                TextHighlightId::from_object_id(parse_u64(arguments.next(), "highlight-id")?)?;
            require_end(arguments)?;
            let mut editor = IWorkTextEditor::open(input)?;
            editor.remove_text_highlight(storage_id, highlight_id)?;
            editor.save(output)?;
            println!("removed highlight {}", highlight_id.object_id());
        },
        _ => return Err(USAGE.into()),
    }
    Ok(())
}

fn parse_range(
    arguments: &mut impl Iterator<Item = String>,
) -> Result<TextRange, Box<dyn std::error::Error>> {
    let start = parse_usize(arguments.next(), "start")?;
    let end = parse_usize(arguments.next(), "end")?;
    Ok(TextRange::from_utf16_indexes(start, end)?)
}

fn parse_u64(value: Option<String>, label: &str) -> Result<u64, Box<dyn std::error::Error>> {
    value
        .ok_or_else(|| USAGE.into())
        .and_then(|value| value.parse().map_err(|_| format!("invalid {label}").into()))
}

fn parse_usize(value: Option<String>, label: &str) -> Result<usize, Box<dyn std::error::Error>> {
    value
        .ok_or_else(|| USAGE.into())
        .and_then(|value| value.parse().map_err(|_| format!("invalid {label}").into()))
}

fn require_end(
    mut arguments: impl Iterator<Item = String>,
) -> Result<(), Box<dyn std::error::Error>> {
    if arguments.next().is_some() {
        return Err("unexpected extra arguments".into());
    }
    Ok(())
}
