//! Edit one Pages section's pagination through a semantic selector.

use std::error::Error;
use std::fs::OpenOptions;
use std::io::Write as _;
use std::path::{Path, PathBuf};

use litchi_pages::{
    Package, SectionSelector,
    section::{PageNumber, PageNumbering, Start},
};

fn main() -> Result<(), Box<dyn Error>> {
    let mut arguments = std::env::args_os().skip(1);
    let input = PathBuf::from(arguments.next().ok_or(
        "usage: edit_section_pagination <input.pages> <output.pages> <section-index> \
         <next|right|left|absent> <continue|restart|absent> <page|absent> \
         [inverse-output.pages]",
    )?);
    let output = PathBuf::from(arguments.next().ok_or("missing output Pages path")?);
    let position =
        text_argument(arguments.next(), "missing semantic section index")?.parse::<usize>()?;
    let start = match text_argument(arguments.next(), "missing section start")?.as_str() {
        "next" => Some(Start::NextPage),
        "right" => Some(Start::RightPage),
        "left" => Some(Start::LeftPage),
        "absent" => None,
        _ => return Err("section start must be next, right, left, or absent".into()),
    };
    let numbering = match text_argument(arguments.next(), "missing page numbering")?.as_str() {
        "continue" => Some(PageNumbering::ContinueFromPrevious),
        "restart" => Some(PageNumbering::Restart),
        "absent" => None,
        _ => return Err("page numbering must be continue, restart, or absent".into()),
    };
    let page = match text_argument(arguments.next(), "missing starting page number")?.as_str() {
        "absent" => None,
        value => Some(PageNumber::new(value.parse::<u32>()?)?),
    };
    let inverse_output = arguments.next().map(PathBuf::from);
    if arguments.next().is_some() {
        return Err("unexpected trailing arguments".into());
    }

    let package = Package::open(input)?;
    let mut edit = package.edit_section_pagination(SectionSelector::index(position))?;
    edit.set_start(start)?;
    edit.set_page_numbering(numbering)?;
    edit.set_starting_page_number(page);
    let commit = edit.commit()?;
    write_new(&output, commit.package().source_bytes())?;

    if let Some(inverse_path) = inverse_output {
        let inverse = commit.patch().inverse();
        let restored = commit.package().apply_section_pagination(&inverse)?;
        write_new(&inverse_path, restored.package().source_bytes())?;
    }
    Ok(())
}

fn text_argument(
    argument: Option<std::ffi::OsString>,
    missing: &'static str,
) -> Result<String, Box<dyn Error>> {
    argument.ok_or_else(|| missing.into()).and_then(|value| {
        value
            .into_string()
            .map_err(|_value| "argument is not valid UTF-8".into())
    })
}

fn write_new(path: &Path, bytes: &[u8]) -> Result<(), Box<dyn Error>> {
    let mut destination = OpenOptions::new().write(true).create_new(true).open(path)?;
    destination.write_all(bytes)?;
    destination.sync_all()?;
    Ok(())
}
