//! List or mutate ordered Keynote soundtrack audio items.

use std::env;
use std::fs;
use std::path::Path;

use litchi_iwa::keynote::KeynoteEditor;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let operation = arguments.next().ok_or(USAGE)?;
    let input = arguments.next().ok_or(USAGE)?;
    let mut editor = KeynoteEditor::open(input)?;

    if operation == "list" {
        require_exhausted(arguments)?;
        print_items(&editor)?;
        return Ok(());
    }

    let output = arguments.next().ok_or(USAGE)?;
    match operation.as_str() {
        "add" => {
            let audio = arguments.next().ok_or(USAGE)?;
            require_exhausted(arguments)?;
            let (filename, data) = read_audio(&audio)?;
            editor.add_soundtrack_item(&filename, &data)?;
        },
        "insert" => {
            let index = parse_index(arguments.next())?;
            let audio = arguments.next().ok_or(USAGE)?;
            require_exhausted(arguments)?;
            let (filename, data) = read_audio(&audio)?;
            editor.insert_soundtrack_item(index, &filename, &data)?;
        },
        "replace" => {
            let index = parse_index(arguments.next())?;
            let audio = arguments.next().ok_or(USAGE)?;
            require_exhausted(arguments)?;
            let (filename, data) = read_audio(&audio)?;
            editor.replace_soundtrack_item(index, &filename, &data)?;
        },
        "move" => {
            let from = parse_index(arguments.next())?;
            let to = parse_index(arguments.next())?;
            require_exhausted(arguments)?;
            editor.move_soundtrack_item(from, to)?;
        },
        "remove" => {
            let index = parse_index(arguments.next())?;
            require_exhausted(arguments)?;
            editor.remove_soundtrack_item(index)?;
        },
        _ => return Err(USAGE.into()),
    }
    editor.save(output)?;
    print_items(&editor)?;
    Ok(())
}

fn print_items(editor: &KeynoteEditor) -> Result<(), Box<dyn std::error::Error>> {
    for item in editor.soundtrack_items()? {
        println!(
            "{}: data={} file={} size={}",
            item.index,
            item.asset.data_identifier,
            item.asset.preferred_filename,
            item.asset.size.unwrap_or(0)
        );
    }
    Ok(())
}

fn parse_index(value: Option<String>) -> Result<usize, Box<dyn std::error::Error>> {
    Ok(value.ok_or(USAGE)?.parse()?)
}

fn read_audio(path: &str) -> Result<(String, Vec<u8>), Box<dyn std::error::Error>> {
    let filename = Path::new(path)
        .file_name()
        .and_then(|filename| filename.to_str())
        .ok_or("audio path must end in a UTF-8 filename")?
        .to_owned();
    Ok((filename, fs::read(path)?))
}

fn require_exhausted(mut arguments: impl Iterator<Item = String>) -> Result<(), &'static str> {
    if arguments.next().is_some() {
        return Err(USAGE);
    }
    Ok(())
}

const USAGE: &str = "usage:\n  edit_keynote_soundtrack_items list <input.key>\n  edit_keynote_soundtrack_items add <input.key> <output.key> <audio>\n  edit_keynote_soundtrack_items insert <input.key> <output.key> <index> <audio>\n  edit_keynote_soundtrack_items replace <input.key> <output.key> <index> <audio>\n  edit_keynote_soundtrack_items move <input.key> <output.key> <from-index> <to-index>\n  edit_keynote_soundtrack_items remove <input.key> <output.key> <index>";
