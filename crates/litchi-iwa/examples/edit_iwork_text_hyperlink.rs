//! List, create, update, or delete native text hyperlinks in any iWork package.

use std::env;

use litchi_iwa::text::{IWorkTextEditor, TextHyperlinkTarget, TextRange};
use litchi_iwa_text::hyperlink::raw::{from_object_id, object_id as native_object_id};

const USAGE: &str = "usage:\n  edit_iwork_text_hyperlink list <input> <storage-id>\n  edit_iwork_text_hyperlink add <input> <output> <storage-id> <start> <end> <target>\n  edit_iwork_text_hyperlink update <input> <output> <storage-id> <hyperlink-id> <start> <end> <target>\n  edit_iwork_text_hyperlink remove <input> <output> <storage-id> <hyperlink-id>";

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let command = arguments.next().ok_or(USAGE)?;
    match command.as_str() {
        "list" => {
            let input = arguments.next().ok_or(USAGE)?;
            let storage_id = parse_u64(arguments.next(), "storage-id")?;
            require_end(arguments)?;
            let editor = IWorkTextEditor::open(input)?;
            for hyperlink in editor.text_hyperlinks(storage_id)? {
                println!(
                    "id={} range={}..{} target={}",
                    native_object_id(hyperlink.id),
                    hyperlink.range.start().utf16_index(),
                    hyperlink.range.end().utf16_index(),
                    hyperlink.target.as_str(),
                );
            }
        },
        "add" => {
            let input = arguments.next().ok_or(USAGE)?;
            let output = arguments.next().ok_or(USAGE)?;
            let storage_id = parse_u64(arguments.next(), "storage-id")?;
            let range = parse_range(&mut arguments)?;
            let target = TextHyperlinkTarget::try_from(arguments.next().ok_or(USAGE)?)?;
            require_end(arguments)?;
            let mut editor = IWorkTextEditor::open(input)?;
            let hyperlink = editor.add_text_hyperlink(storage_id, range, target)?;
            editor.save(output)?;
            println!("created hyperlink {}", native_object_id(hyperlink.id));
        },
        "update" => {
            let input = arguments.next().ok_or(USAGE)?;
            let output = arguments.next().ok_or(USAGE)?;
            let storage_id = parse_u64(arguments.next(), "storage-id")?;
            let hyperlink_id = from_object_id(parse_u64(arguments.next(), "hyperlink-id")?)?;
            let range = parse_range(&mut arguments)?;
            let target = TextHyperlinkTarget::try_from(arguments.next().ok_or(USAGE)?)?;
            require_end(arguments)?;
            let mut editor = IWorkTextEditor::open(input)?;
            editor.update_text_hyperlink(storage_id, hyperlink_id, range, target)?;
            editor.save(output)?;
            println!("updated hyperlink {}", native_object_id(hyperlink_id));
        },
        "remove" => {
            let input = arguments.next().ok_or(USAGE)?;
            let output = arguments.next().ok_or(USAGE)?;
            let storage_id = parse_u64(arguments.next(), "storage-id")?;
            let hyperlink_id = from_object_id(parse_u64(arguments.next(), "hyperlink-id")?)?;
            require_end(arguments)?;
            let mut editor = IWorkTextEditor::open(input)?;
            editor.remove_text_hyperlink(storage_id, hyperlink_id)?;
            editor.save(output)?;
            println!("removed hyperlink {}", native_object_id(hyperlink_id));
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
