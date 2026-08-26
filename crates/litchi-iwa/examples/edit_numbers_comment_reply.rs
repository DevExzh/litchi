//! Create, update, or delete a direct reply in a Numbers cell-comment thread.

use std::env;

use litchi_iwa::numbers::NumbersEditor;
use litchi_numbers::cell::comment::CommentReplyIndex;
use litchi_numbers::{CellPosition, Package, SheetSelector, TableSelector};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments.next().ok_or(
        "usage: edit_numbers_comment_reply <input.numbers> <output.numbers> <table-id-or-name> <row> <column> <add|set|remove> [reply-index] [text]",
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
    let tables = editor.tables()?;
    let table = table_selector
        .parse::<u64>()
        .ok()
        .and_then(|id| tables.iter().find(|table| table.id() == id))
        .or_else(|| tables.iter().find(|table| table.name == table_selector));
    let table_id = table
        .ok_or("table selector did not match a Numbers table")?
        .id();
    let source = Package::from_bytes(&editor.to_bytes()?)?;
    let table_name = tables
        .iter()
        .find(|candidate| candidate.id() == table_id)
        .ok_or("selected Numbers table disappeared")?
        .name
        .clone();
    let mut table_matches = source
        .sheets()
        .iter()
        .enumerate()
        .flat_map(|(sheet_index, sheet)| {
            sheet
                .tables()
                .enumerate()
                .map(move |(table_index, table)| (sheet_index, table_index, table))
        })
        .filter(|(_, _, table)| table.name() == table_name);
    let (sheet_index, table_index, _) = table_matches
        .next()
        .ok_or("selected Numbers table is not rooted in the semantic package")?;
    if table_matches.next().is_some() {
        return Err("selected Numbers table name is ambiguous".into());
    }
    let sheet = SheetSelector::index(sheet_index);
    let table = TableSelector::index(table_index);
    let position = CellPosition::try_from_usize(row, column)?;

    match operation.as_str() {
        "add" => {
            let text = arguments.collect::<Vec<_>>().join(" ");
            if text.is_empty() {
                return Err("missing reply text".into());
            }
            let commit = source.add_table_cell_comment_reply(sheet, table, position, text)?;
            publish(&mut editor, commit)?;
            println!("added reply");
        },
        "set" => {
            let reply_index = arguments
                .next()
                .ok_or("missing reply index")?
                .parse::<u32>()?;
            let text = arguments.collect::<Vec<_>>().join(" ");
            if text.is_empty() {
                return Err("missing replacement text".into());
            }
            let commit = source.set_table_cell_comment_reply(
                sheet,
                table,
                position,
                CommentReplyIndex::new(reply_index),
                text,
            )?;
            publish(&mut editor, commit)?;
            println!("updated reply-index={reply_index}");
        },
        "remove" => {
            let reply_index = arguments
                .next()
                .ok_or("missing reply index")?
                .parse::<u32>()?;
            if arguments.next().is_some() {
                return Err("unexpected extra arguments".into());
            }
            let commit = source.remove_table_cell_comment_reply(
                sheet,
                table,
                position,
                CommentReplyIndex::new(reply_index),
            )?;
            publish(&mut editor, commit)?;
            println!("removed reply-index={reply_index}");
        },
        _ => return Err("operation must be add, set, or remove".into()),
    }
    editor.save(&output)?;
    let final_package = Package::from_bytes(&editor.to_bytes()?)?;
    for (index, reply) in final_package
        .table_cell_comment_replies(sheet, table, position)?
        .iter()
        .enumerate()
    {
        println!("reply-index={index} text={:?}", reply.text());
    }
    Ok(())
}

fn publish(
    editor: &mut NumbersEditor,
    commit: litchi_numbers::cell::comment::transaction::Commit,
) -> Result<(), Box<dyn std::error::Error>> {
    let mut bytes = Vec::new();
    commit.package().write_to(&mut bytes)?;
    *editor = NumbersEditor::from_bytes(&bytes)?;
    Ok(())
}
