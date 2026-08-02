//! Update native Pages section settings while preserving every other field.

use std::env;

use litchi_iwa::pages::PagesEditor;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments.next().ok_or(
        "usage: set_pages_section_settings <input.pages> <output.pages> <section-id> \
         <match-previous:true|false> <hide-first-page:true|false>",
    )?;
    let output = arguments.next().ok_or("missing output path")?;
    let section_id = arguments
        .next()
        .ok_or("missing section ID")?
        .parse::<u64>()?;
    let match_previous = arguments
        .next()
        .ok_or("missing match-previous value")?
        .parse::<bool>()?;
    let hide_first_page = arguments
        .next()
        .ok_or("missing hide-first-page value")?
        .parse::<bool>()?;

    let mut editor = PagesEditor::open(input)?;
    let mut settings = editor.section_settings(section_id)?;
    settings.inherit_previous_header_footer = Some(match_previous);
    settings.first_page_hides_header_footer = Some(hide_first_page);
    editor.set_section_settings(section_id, settings)?;
    editor.save(output)?;
    Ok(())
}
