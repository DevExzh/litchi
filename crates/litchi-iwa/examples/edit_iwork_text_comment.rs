//! List, create, update, or delete native ranged comments in any iWork package.

use std::env;

use litchi_iwa::text::{IWorkTextEditor, TextCommentBody, TextCommentId, TextRange};

const USAGE: &str = "usage:\n  edit_iwork_text_comment list <input> <storage-id>\n  edit_iwork_text_comment add <input> <output> <storage-id> <start> <end> <body>\n  edit_iwork_text_comment update <input> <output> <storage-id> <comment-id> <start> <end> <body>\n  edit_iwork_text_comment remove <input> <output> <storage-id> <comment-id>";

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let command = arguments.next().ok_or(USAGE)?;
    match command.as_str() {
        "list" => {
            let input = arguments.next().ok_or(USAGE)?;
            let storage_id = parse_u64(arguments.next(), "storage-id")?;
            require_end(arguments)?;
            let editor = IWorkTextEditor::open(input)?;
            for comment in editor.text_comments(storage_id)? {
                println!(
                    "id={} range={}..{} replies={} body={:?}",
                    comment.id.object_id(),
                    comment.range.start().utf16_index(),
                    comment.range.end().utf16_index(),
                    comment.reply_count,
                    comment.body.as_str(),
                );
            }
        },
        "add" => {
            let input = arguments.next().ok_or(USAGE)?;
            let output = arguments.next().ok_or(USAGE)?;
            let storage_id = parse_u64(arguments.next(), "storage-id")?;
            let range = parse_range(&mut arguments)?;
            let body = TextCommentBody::new(arguments.next().ok_or(USAGE)?)?;
            require_end(arguments)?;
            let mut editor = IWorkTextEditor::open(input)?;
            let comment = editor.add_text_comment(storage_id, range, body)?;
            editor.save(output)?;
            println!("created comment {}", comment.id.object_id());
        },
        "update" => {
            let input = arguments.next().ok_or(USAGE)?;
            let output = arguments.next().ok_or(USAGE)?;
            let storage_id = parse_u64(arguments.next(), "storage-id")?;
            let comment_id =
                TextCommentId::from_object_id(parse_u64(arguments.next(), "comment-id")?)?;
            let range = parse_range(&mut arguments)?;
            let body = TextCommentBody::new(arguments.next().ok_or(USAGE)?)?;
            require_end(arguments)?;
            let mut editor = IWorkTextEditor::open(input)?;
            editor.update_text_comment(storage_id, comment_id, range, body)?;
            editor.save(output)?;
            println!("updated comment {}", comment_id.object_id());
        },
        "remove" => {
            let input = arguments.next().ok_or(USAGE)?;
            let output = arguments.next().ok_or(USAGE)?;
            let storage_id = parse_u64(arguments.next(), "storage-id")?;
            let comment_id =
                TextCommentId::from_object_id(parse_u64(arguments.next(), "comment-id")?)?;
            require_end(arguments)?;
            let mut editor = IWorkTextEditor::open(input)?;
            editor.remove_text_comment(storage_id, comment_id)?;
            editor.save(output)?;
            println!("removed comment {}", comment_id.object_id());
        },
        _ => return Err(USAGE.into()),
    }
    Ok(())
}

fn parse_range(
    arguments: &mut impl Iterator<Item = String>,
) -> Result<TextRange, Box<dyn std::error::Error>> {
    Ok(TextRange::from_utf16_indexes(
        parse_usize(arguments.next(), "start")?,
        parse_usize(arguments.next(), "end")?,
    )?)
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
