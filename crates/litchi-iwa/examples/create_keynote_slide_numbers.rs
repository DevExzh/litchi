//! Create a Keynote presentation with native slide numbers and no input file.

use std::path::{Path, PathBuf};

use litchi_iwa::keynote::KeynoteDocumentBuilder;
use litchi_keynote::{Package, SlideSelector, SlideTextRole};
use tempfile::NamedTempFile;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let output = PathBuf::from(
        std::env::args()
            .nth(1)
            .ok_or("usage: create_keynote_slide_numbers <output.key>")?,
    );

    let mut keynote = KeynoteDocumentBuilder::new()
        .title("Native slide numbers")
        .subtitle("Created entirely by litchi-iwa")
        .slide_number_visible(true)
        .build()?;
    let layout = keynote.default_slide_layout()?;
    keynote.add_slide(layout)?;
    keynote.set_slide_number_visible(1, true)?;

    let package = Package::from_bytes(&keynote.to_bytes()?)
        .map_err(|error| std::io::Error::other(format!("focused reopen failed: {error:?}")))?;
    let mut title = package
        .edit_slide_text(SlideSelector::index(1), SlideTextRole::Title)
        .map_err(|error| std::io::Error::other(format!("title preflight failed: {error:?}")))?;
    title.set("Second slide")?;
    let title = title
        .commit()
        .map_err(|error| std::io::Error::other(format!("title commit failed: {error:?}")))?;
    write_new(&output, title.package())?;
    Ok(())
}

/// Publishes through a sibling temporary file without replacing an existing target.
///
/// This example does not provide the library's durable atomic-save contract.
fn write_new(path: &Path, package: &Package) -> Result<(), Box<dyn std::error::Error>> {
    let parent = path
        .parent()
        .filter(|parent| !parent.as_os_str().is_empty())
        .unwrap_or_else(|| Path::new("."));
    let mut temporary = NamedTempFile::new_in(parent)?;
    package.write_to(&mut temporary)?;
    temporary.as_file().sync_all()?;
    temporary
        .persist_noclobber(path)
        .map_err(|error| Box::new(error.error) as Box<dyn std::error::Error>)?;
    Ok(())
}
