//! Structural edits of LibreOffice-authored drawings publish as exact-source
//! `content.xml` splices (record 0288): the producer's prolog and every byte
//! outside the inserted, removed or moved elements are kept.

use litchi_odg::Package;
use litchi_odg::page::Page;
use litchi_odg::shape::{Shape, ShapeKind};

fn fixture(name: &str) -> Package {
    let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/libreoffice-core")
        .join(name);
    Package::from_bytes(&std::fs::read(path).unwrap()).unwrap()
}

const THREE_PAGES: &str = "desktop/qa/data/3page.odg";
const FONT_WORK: &str = "svx/qa/unit/data/FontWork.odg";

fn names(package: &Package) -> Vec<String> {
    package
        .drawing()
        .pages()
        .iter()
        .map(|page| page.name().unwrap_or_default().to_owned())
        .collect()
}

/// The common prefix and suffix of two parts cover all but `changed` bytes
/// of the source.
fn assert_local(before: &str, after: &str, changed: usize) {
    let prefix = before
        .bytes()
        .zip(after.bytes())
        .take_while(|(left, right)| left == right)
        .count();
    let suffix = before
        .bytes()
        .rev()
        .zip(after.bytes().rev())
        .take_while(|(left, right)| left == right)
        .count();
    assert!(
        prefix + suffix + changed >= before.len(),
        "the edit rewrote more than {changed} source bytes"
    );
}

#[test]
fn producer_sources_are_not_compact() {
    // The fixtures carry LibreOffice's newline after the XML declaration,
    // which a whole-part rewrite refuses.
    let package = fixture(THREE_PAGES);
    assert!(
        package
            .content_xml()
            .starts_with("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n")
    );
}

#[test]
fn inserts_pages_and_shapes_into_a_producer_drawing() {
    let package = fixture(THREE_PAGES);
    let source = package.content_xml().to_owned();
    let mut edit = package.edit();
    edit.add_page(Page::new("added")).unwrap();
    edit.add_page(Page::new("second")).unwrap();
    edit.add_shape(
        0,
        Shape::new(ShapeKind::Rectangle)
            .with_geometry("1cm", "1cm", "4cm", "2cm")
            .with_text("Box"),
    )
    .unwrap();
    let next = edit.commit().unwrap().into_package();
    assert_eq!(names(&next), ["page1", "page2", "page3", "added", "second"]);
    assert_eq!(
        next.drawing().pages()[0].shapes().last().unwrap().text(),
        "Box"
    );
    assert!(
        next.content_xml()
            .starts_with("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n")
    );
    assert!(next.content_xml().len() > source.len());
    // Untouched members keep their bytes.
    assert_eq!(next.styles_xml(), package.styles_xml());
    let reopened = Package::from_bytes(next.as_bytes()).unwrap();
    assert_eq!(names(&reopened), names(&next));
}

#[test]
fn removes_and_moves_producer_pages() {
    let package = fixture(THREE_PAGES);
    let source = package.content_xml().to_owned();
    let mut edit = package.edit();
    edit.remove_page(1usize).unwrap();
    let removed = edit.commit().unwrap().into_package();
    assert_eq!(names(&removed), ["page1", "page3"]);
    assert_local(
        &source,
        removed.content_xml(),
        source.len() - removed.content_xml().len(),
    );

    let mut edit = package.edit();
    edit.move_page(0, 2).unwrap();
    let moved = edit.commit().unwrap().into_package();
    assert_eq!(names(&moved), ["page2", "page3", "page1"]);
    assert_eq!(moved.content_xml().len(), source.len());

    let mut edit = package.edit();
    edit.move_page(2, 0).unwrap();
    let moved = edit.commit().unwrap().into_package();
    assert_eq!(names(&moved), ["page3", "page1", "page2"]);
    // The inverse patch restores the source.
    let mut edit = moved.edit();
    edit.move_page(0, 2).unwrap();
    let back = edit.commit().unwrap().into_package();
    assert_eq!(back.content_xml(), source);
}

#[test]
fn moves_and_removes_producer_shapes() {
    let package = fixture(FONT_WORK);
    let drawing = package.drawing();
    let page = &drawing.pages()[0];
    if page.shapes().len() < 2 {
        let mut edit = package.edit();
        edit.remove_shape(0, 0).unwrap();
        let next = edit.commit().unwrap().into_package();
        assert_eq!(
            next.drawing().pages()[0].shapes().len(),
            page.shapes().len() - 1
        );
        return;
    }
    let first = page.shapes()[0].name().map(str::to_owned);
    let mut edit = package.edit();
    edit.move_shape(0, 0, 1).unwrap();
    let next = edit.commit().unwrap().into_package();
    assert_eq!(
        next.drawing().pages()[0].shapes()[1]
            .name()
            .map(str::to_owned),
        first
    );
}

#[test]
fn a_removal_next_to_an_insertion_joins_one_splice() {
    let package = fixture(THREE_PAGES);
    let mut edit = package.edit();
    edit.remove_page(2usize).unwrap();
    edit.add_page(Page::new("tail")).unwrap();
    edit.insert_page(0, Page::new("head")).unwrap();
    let next = edit.commit().unwrap().into_package();
    assert_eq!(names(&next), ["head", "page1", "page2", "tail"]);
}
