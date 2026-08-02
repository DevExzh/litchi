//! Create, update, or delete a direct reply in a Numbers cell-comment thread.

use std::env;

use litchi_iwa::numbers::NumbersEditor;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments.next().ok_or(
        "usage: edit_numbers_comment_reply <input.numbers> <output.numbers> <table-id-or-name> <row> <column> <add|set|remove> [reply-id] [text]",
    )?;
    let output = arguments.next().ok_or("missing output path")?;
    let table_selector = arguments.next().ok_or("missing table ID or name")?;
    let row = arguments
        .next()
        .ok_or("missing zero-based row")?
        .parse::<usize>()?;
    let column = arguments
        .next()
        .ok_or("missing zero-based column")?
        .parse::<usize>()?;
    let operation = arguments.next().ok_or("missing add, set, or remove")?;

    let mut editor = NumbersEditor::open(&input)?;
    let table_id = table_selector.parse::<u64>().ok().or_else(|| {
        editor
            .tables()
            .ok()?
            .into_iter()
            .find(|table| table.name == table_selector)
            .map(|table| table.object_id)
    });
    let table_id = table_id.ok_or("table selector did not match a Numbers table")?;

    match operation.as_str() {
        "add" => {
            let text = arguments.collect::<Vec<_>>().join(" ");
            if text.is_empty() {
                return Err("missing reply text".into());
            }
            let reply_id = editor.add_cell_comment_reply(table_id, row, column, text)?;
            println!("added reply={reply_id}");
        },
        "set" => {
            let reply_id = arguments
                .next()
                .ok_or("missing reply object identifier")?
                .parse::<u64>()?;
            let text = arguments.collect::<Vec<_>>().join(" ");
            if text.is_empty() {
                return Err("missing replacement text".into());
            }
            let current_id =
                editor.set_cell_comment_reply(table_id, row, column, reply_id, text)?;
            println!("updated reply={reply_id} current_reply={current_id}");
        },
        "remove" => {
            let reply_id = arguments
                .next()
                .ok_or("missing reply object identifier")?
                .parse::<u64>()?;
            if arguments.next().is_some() {
                return Err("unexpected extra arguments".into());
            }
            editor.remove_cell_comment_reply(table_id, row, column, reply_id)?;
            println!("removed reply={reply_id}");
        },
        _ => return Err("operation must be add, set, or remove".into()),
    }
    editor.save(&output)?;
    for reply in editor.cell_comment_replies(table_id, row, column)? {
        println!(
            "reply={} author={:?} text={:?}",
            reply.storage_object_id, reply.comment.author_object_id, reply.comment.text
        );
    }
    Ok(())
}
