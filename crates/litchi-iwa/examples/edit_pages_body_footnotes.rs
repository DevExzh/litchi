use std::env;

use litchi_iwa::pages::PagesEditor;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments.next().ok_or(
        "usage: edit_pages_body_footnotes <input.pages> <output.pages> <set|remove> <index> [text]",
    )?;
    let output = arguments
        .next()
        .ok_or("missing output Pages document path")?;
    let operation = arguments.next().ok_or("missing footnote operation")?;
    let index = arguments
        .next()
        .ok_or("missing footnote index")?
        .parse::<usize>()?;

    let mut pages = PagesEditor::open(input)?;
    let footnote = pages
        .body_footnotes()?
        .get(index)
        .cloned()
        .ok_or("footnote index is out of bounds")?;
    match operation.as_str() {
        "set" => {
            let text = arguments
                .next()
                .ok_or("missing replacement footnote text")?;
            if arguments.next().is_some() {
                return Err("replacement footnote text must be one argument".into());
            }
            pages.set_body_footnote_text(footnote.id, text)?;
        },
        "remove" => {
            if arguments.next().is_some() {
                return Err("remove does not accept replacement text".into());
            }
            pages.remove_body_footnote(footnote.id)?;
        },
        _ => return Err("footnote operation must be set or remove".into()),
    }
    pages.save(output)?;
    Ok(())
}
