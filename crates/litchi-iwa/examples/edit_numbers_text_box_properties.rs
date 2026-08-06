//! Update shared drawable properties on an ordinary Numbers text box.

use std::env;

use litchi_iwa::numbers::NumbersEditor;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments.next().ok_or(
        "usage: edit_numbers_text_box_properties <input.numbers> <output.numbers> <sheet-index> <drawable-id> <locked:true|false|none> <aspect-ratio-locked:true|false|none> <hyperlink|none> <accessibility-description|none>",
    )?;
    let output = arguments.next().ok_or("missing output path")?;
    let sheet_index: usize = arguments.next().ok_or("missing sheet index")?.parse()?;
    let drawable_id: u64 = arguments.next().ok_or("missing drawable ID")?.parse()?;
    let locked = parse_optional_bool(arguments.next().ok_or("missing locked value")?)?;
    let aspect_ratio_locked = parse_optional_bool(
        arguments
            .next()
            .ok_or("missing aspect-ratio-locked value")?,
    )?;
    let hyperlink_url = parse_optional_string(arguments.next().ok_or("missing hyperlink")?);
    let accessibility_description = parse_optional_string(
        arguments
            .next()
            .ok_or("missing accessibility description")?,
    );
    if arguments.next().is_some() {
        return Err("unexpected extra arguments".into());
    }

    let mut editor = NumbersEditor::open(input)?;
    let sheet = editor
        .sheets()?
        .into_iter()
        .nth(sheet_index)
        .ok_or("sheet index out of range")?;
    let mut properties = editor.sheet_text_box_properties(sheet.id(), drawable_id)?;
    properties.locked = locked;
    properties.aspect_ratio_locked = aspect_ratio_locked;
    properties.hyperlink_url = hyperlink_url;
    properties.accessibility_description = accessibility_description;
    editor.set_sheet_text_box_properties(sheet.id(), drawable_id, properties.clone())?;
    editor.save(output)?;
    println!("sheet={sheet_index} drawable={drawable_id} properties={properties:?}");
    Ok(())
}

fn parse_optional_bool(value: String) -> Result<Option<bool>, Box<dyn std::error::Error>> {
    match value.as_str() {
        "true" => Ok(Some(true)),
        "false" => Ok(Some(false)),
        "none" => Ok(None),
        _ => Err("boolean property must be true, false, or none".into()),
    }
}

fn parse_optional_string(value: String) -> Option<String> {
    (value != "none").then_some(value)
}
