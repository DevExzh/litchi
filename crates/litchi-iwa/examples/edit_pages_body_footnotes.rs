use std::env;
use std::path::{Path, PathBuf};

use litchi_pages::Package;
use litchi_pages::footnote::body::Selector;
use tempfile::NamedTempFile;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = PathBuf::from(arguments.next().ok_or(
        "usage: edit_pages_body_footnotes <input.pages> <output.pages> <set|remove> <index> [text]",
    )?);
    let output = PathBuf::from(
        arguments
            .next()
            .ok_or("missing output Pages document path")?,
    );
    let operation = arguments.next().ok_or("missing footnote operation")?;
    let index = arguments
        .next()
        .ok_or("missing footnote index")?
        .parse::<usize>()?;

    let selector = Selector::Index(index);
    let package = Package::open(&input)?;
    match operation.as_str() {
        "set" => {
            let text = arguments
                .next()
                .ok_or("missing replacement footnote text")?;
            if arguments.next().is_some() {
                return Err("replacement footnote text must be one argument".into());
            }
            let mut edit = package.edit_body_footnote(selector)?;
            edit.set(&text)?;
            let commit = edit.commit()?;
            save_new(&output, commit.package())?;
        },
        "remove" => {
            if arguments.next().is_some() {
                return Err("remove does not accept replacement text".into());
            }
            let mut edit = package.edit_body_footnote(selector)?;
            edit.clear();
            let commit = edit.commit()?;
            save_new(&output, commit.package())?;
        },
        _ => return Err("footnote operation must be set or remove".into()),
    }
    Ok(())
}

/// Publishes the focused package through a synchronized sibling temporary
/// file without overwriting an existing target.
fn save_new(path: &Path, package: &Package) -> Result<(), Box<dyn std::error::Error>> {
    let parent = path
        .parent()
        .filter(|parent| !parent.as_os_str().is_empty())
        .unwrap_or_else(|| Path::new("."));
    let mut temporary = NamedTempFile::new_in(parent)?;
    package.write_to(temporary.as_file_mut())?;
    temporary.as_file().sync_all()?;
    temporary
        .persist_noclobber(path)
        .map_err(|error| error.error)?;
    Ok(())
}
