//! Update a Pages section's typed page-start and numbering settings.

use std::env;

use litchi_iwa::pages::PagesEditor;
use litchi_pages::section::{PageNumber, PageNumbering, Settings, Start};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments.next().ok_or(
        "usage: edit_pages_section_pagination <input.pages> <output.pages> <section-id> \
         <next|right|left> <continue|restart> <starting-page-number>",
    )?;
    let output = arguments.next().ok_or("missing output path")?;
    let section_id = arguments
        .next()
        .ok_or("missing section ID")?
        .parse::<u64>()?;
    let start = match arguments.next().ok_or("missing section start")?.as_str() {
        "next" => Start::NextPage,
        "right" => Start::RightPage,
        "left" => Start::LeftPage,
        _ => return Err("section start must be next, right, or left".into()),
    };
    let numbering = match arguments
        .next()
        .ok_or("missing page numbering behavior")?
        .as_str()
    {
        "continue" => PageNumbering::ContinueFromPrevious,
        "restart" => PageNumbering::Restart,
        _ => return Err("page numbering must be continue or restart".into()),
    };
    let starting_page_number = PageNumber::new(
        arguments
            .next()
            .ok_or("missing starting page number")?
            .parse::<u32>()?,
    )?;
    if arguments.next().is_some() {
        return Err("unexpected extra argument".into());
    }

    let mut editor = PagesEditor::open(input)?;
    let mut settings: Settings = editor.section_settings(section_id)?;
    settings.start = Some(start);
    settings.page_numbering = Some(numbering);
    settings.starting_page_number = Some(starting_page_number);
    editor.set_section_settings(section_id, settings)?;
    editor.save(output)?;
    Ok(())
}
