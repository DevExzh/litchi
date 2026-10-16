//! PPTX structural edits on producer decks through the unified API
//! (record 0303): slides that own Parts, p14 sections, and an OLE frame
//! under markup compatibility.

#![cfg(feature = "pptx")]

use std::path::{Path, PathBuf};

use litchi_ops::{File, Placement, open, open_bytes};

fn path(relative: &str) -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data")
        .join(relative)
}

fn input(relative: &str) -> (File, Vec<u8>) {
    let path = path(relative);
    (open(&path).unwrap(), std::fs::read(&path).unwrap())
}

fn slides(file: &File) -> usize {
    file.get("/")
        .unwrap()
        .children()
        .filter_map(Result::ok)
        .filter(|child| child.kind() == "slide")
        .count()
}

fn slide_text(file: &File, slide: usize) -> Option<String> {
    let node = file.get(format!("/slide[{slide}]").as_str()).ok()?;
    node.view().text().ok()?.map(std::borrow::Cow::into_owned)
}

#[test]
fn a_slide_with_notes_comments_and_an_ole_object_is_removed() {
    let (file, _) = input("poi/test-data/slideshow/45545_Comment.pptx");
    let second = slide_text(&file, 2);
    let removed = file.remove("/slide[1]").unwrap();
    let reopened = open_bytes(removed.to_bytes().unwrap()).unwrap();
    assert_eq!(slides(&reopened), 10);
    assert_eq!(slide_text(&reopened, 1), second);
}

#[test]
fn a_sectioned_deck_moves_and_removes_slides() {
    let (file, bytes) = input("libreoffice-core/sd/qa/unit/data/pptx/slide-sections.pptx");
    let moved = file
        .move_to("/slide[1]", Placement::after("/slide[2]").unwrap())
        .unwrap();
    let back = moved.move_to("/slide[2]", Placement::Start).unwrap();
    assert_eq!(back.to_bytes().unwrap(), bytes);

    let removed = file.remove("/slide[4]").unwrap();
    let reopened = open_bytes(removed.to_bytes().unwrap()).unwrap();
    assert_eq!(slides(&reopened), 6);
}

#[test]
fn an_ole_frame_under_markup_compatibility_is_removed() {
    let (file, _) = input("libreoffice-core/sd/qa/unit/data/pptx/ole.pptx");
    let ole = file
        .query("ole")
        .unwrap()
        .next()
        .expect("the deck has an OLE object")
        .unwrap()
        .path()
        .to_string();
    let removed = file.remove(ole.as_str()).unwrap();
    let reopened = open_bytes(removed.to_bytes().unwrap()).unwrap();
    assert!(reopened.query("ole").unwrap().next().is_none());
    // The rectangle beside it stays. Editing it on the slide that holds the
    // markup compatibility stays refused: only whole-shape removal is
    // proven on the source bytes.
    assert!(reopened.get("/slide[1]/shape[1]").is_ok());
    let error = file
        .set(
            "/slide[1]/shape[1]",
            litchi_ops::props! { "text" => "changed" },
        )
        .map(drop)
        .unwrap_err();
    assert_eq!(error.code().as_str(), "unsupported_capability");
}
