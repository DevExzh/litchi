use std::env;
use std::path::{Path, PathBuf};

use litchi_iwa::pages::PagesEditor;
use litchi_pages::Package;
use litchi_pages::footnote::body::{Position, Selector};
use tempfile::NamedTempFile;

const BODY: &str = "Alpha Beta";
const FOOTNOTE_POSITION: usize = 6;
const INITIAL_FOOTNOTE_TEXT: &str = "Initial note from the host graph lifecycle.";
const FOOTNOTE_TEXT: &str = "Created from scratch with litchi-iwa.";

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let output = PathBuf::from(
        env::args()
            .nth(1)
            .ok_or("usage: create_pages_footnotes <output.pages>")?,
    );
    let pages = PagesEditor::create_with_text(BODY)?;
    let package = Package::from_bytes(&pages.to_bytes()?)?;
    let position = Position::from_utf16_index(FOOTNOTE_POSITION)?;
    let inserted = Package::insert_body_footnote(&package, position, INITIAL_FOOTNOTE_TEXT, None)?;
    let footnote = inserted
        .package()
        .body_footnotes()?
        .into_iter()
        .find(|footnote| footnote.position == position)
        .ok_or("inserted footnote was not found")?;
    let mut edit = inserted
        .package()
        .edit_body_footnote(Selector::At(footnote.position))?;
    edit.set(FOOTNOTE_TEXT)?;
    let commit = edit.commit()?;
    save_new(&output, commit.package())?;
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
