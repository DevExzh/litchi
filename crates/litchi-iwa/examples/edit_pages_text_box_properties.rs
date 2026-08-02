//! Update shared drawable properties on an ordinary Pages text box.

use std::env;

use litchi_iwa::pages::PagesEditor;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments.next().ok_or(
        "usage: edit_pages_text_box_properties <input.pages> <output.pages> <drawable-id> <locked:true|false|none> <aspect-ratio-locked:true|false|none> <hyperlink|none> <accessibility-description|none>",
    )?;
    let output = arguments.next().ok_or("missing output path")?;
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

    let mut editor = PagesEditor::open(input)?;
    let mut properties = editor.text_box_properties(drawable_id)?;
    properties.locked = locked;
    properties.aspect_ratio_locked = aspect_ratio_locked;
    properties.hyperlink_url = hyperlink_url;
    properties.accessibility_description = accessibility_description;
    editor.set_text_box_properties(drawable_id, properties.clone())?;
    editor.save(output)?;
    println!("drawable={drawable_id} properties={properties:?}");
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
