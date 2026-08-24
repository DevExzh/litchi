use std::env;
use std::path::{Path, PathBuf};

use litchi_iwa::pages::{PagesEditor, Position};
use litchi_pages::Package;
use litchi_pages::footnote::body::Selector;
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
    let mut pages = PagesEditor::create_with_text(BODY)?;
    let footnote = pages.insert_body_footnote(
        Position::from_utf16_index(FOOTNOTE_POSITION)?,
        INITIAL_FOOTNOTE_TEXT,
    )?;
    let package = Package::from_bytes(&pages.to_bytes()?)?;
    let mut edit = package.edit_body_footnote_text(Selector::At(footnote.position))?;
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
