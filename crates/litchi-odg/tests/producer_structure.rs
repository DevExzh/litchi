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

/// Shape counts per page.
fn shape_counts(package: &Package) -> Vec<usize> {
    package
        .drawing()
        .pages()
        .iter()
        .map(|page| page.shapes().len())
        .collect()
}

#[test]
fn moves_a_shape_into_another_producer_page_and_back() {
    let package = fixture(THREE_PAGES);
    let mut edit = package.edit();
    edit.add_shape(
        0,
        Shape::new(ShapeKind::Rectangle)
            .with_geometry("1cm", "1cm", "4cm", "2cm")
            .with_text("Travel"),
    )
    .unwrap();
    let start = edit.commit().unwrap().into_package();
    let counts = shape_counts(&start);
    let from = counts[0] - 1;
    let source = start.content_xml().to_owned();

    // Append to the last page (record 0292).
    let mut edit = start.edit();
    edit.move_shape_to_page(0, from, 2, counts[2]).unwrap();
    let moved = edit.commit().unwrap().into_package();
    let after = shape_counts(&moved);
    assert_eq!(after[0], counts[0] - 1);
    assert_eq!(after[2], counts[2] + 1);
    assert_eq!(
        moved.drawing().pages()[2].shapes()[counts[2]].text(),
        "Travel"
    );
    assert_eq!(moved.styles_xml(), start.styles_xml());
    let reopened = Package::from_bytes(moved.as_bytes()).unwrap();
    assert_eq!(shape_counts(&reopened), after);

    // Moving it back restores the source part byte for byte when the
    // destination page held shapes before; an expanded empty page keeps its
    // open/close form, so compare semantically otherwise.
    let mut edit = moved.edit();
    edit.move_shape_to_page(2, counts[2], 0, from).unwrap();
    let back = edit.commit().unwrap().into_package();
    assert_eq!(shape_counts(&back), counts);
    if counts[2] > 0 {
        assert_eq!(back.content_xml(), source);
    }

    // Out-of-range destinations are refused.
    let mut edit = start.edit();
    assert!(edit.move_shape_to_page(0, from, 1, counts[1] + 1).is_err());
    assert!(edit.move_shape_to_page(0, counts[0], 1, 0).is_err());
}

#[test]
fn a_cross_page_move_refuses_identity_references() {
    let content = r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" office:version="1.3"><office:body><office:drawing><draw:page draw:name="a" draw:master-page-name="Default"><draw:rect draw:id="r1" xml:id="r1" svg:x="0cm" svg:y="0cm" svg:width="1cm" svg:height="1cm"/><draw:rect svg:x="2cm" svg:y="0cm" svg:width="1cm" svg:height="1cm"/><draw:connector draw:start-shape="r1" svg:x1="0cm" svg:y1="0cm" svg:x2="1cm" svg:y2="1cm"/></draw:page><draw:page draw:name="b" draw:master-page-name="Default"/></office:drawing></office:body></office:document-content>"#;
    let package = litchi_odg::Builder::new()
        .content_xml(content)
        .build()
        .unwrap();
    let mut edit = package.edit();
    // The rectangle a connector glues to stays on its page.
    assert!(edit.move_shape_to_page(0, 0, 1, 0).is_err());
    // The connector glued to it too.
    assert!(edit.move_shape_to_page(0, 2, 1, 0).is_err());
    // A free shape moves into the empty page.
    edit.move_shape_to_page(0, 1, 1, 0).unwrap();
    let next = edit.commit().unwrap().into_package();
    assert_eq!(shape_counts(&next), [2, 1]);
}
