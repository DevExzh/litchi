use litchi_odt::Document;
use litchi_odt::font_face::{Declarations, Face};
mod support;

const CONTENT: &str = include_str!("../../../test-data/odf/odt/font-face-declarations-content.xml");
const STYLES: &str = include_str!("../../../test-data/odf/odt/font-face-declarations-styles.xml");
const FLAT: &str = include_str!("../../../test-data/odf/odt/font-face-declarations-flat.fodt");
const MIMETYPE: &str = "application/vnd.oasis.opendocument.text";

fn declarations(name: &str) -> Declarations {
    Declarations {
        faces: vec![Face {
            name: name.to_string(),
            family: Some(format!("'{name}'")),
            ..Default::default()
        }],
    }
}

fn document() -> Document {
    Document::from_bytes(support::package(
        MIMETYPE,
        &[
            ("content.xml", CONTENT.as_bytes()),
            ("styles.xml", STYLES.as_bytes()),
        ],
    ))
    .unwrap()
}

#[test]
fn document_and_generic_package_keep_content_and_styles_declarations_separate() {
    let content = declarations("Content Body");
    let styles = declarations("Styles Body");
    let source = document();

    assert_eq!(
        source.content_font_face_declarations().unwrap(),
        Some(content.clone())
    );
    assert_eq!(
        source.styles_font_face_declarations().unwrap(),
        Some(styles.clone())
    );

    let package =
        litchi_odt::generic::OpenDocumentPackage::from_bytes(source.to_bytes().unwrap()).unwrap();
    assert_eq!(
        package.content_font_face_declarations().unwrap(),
        Some(content)
    );
    assert_eq!(
        package.styles_font_face_declarations().unwrap(),
        Some(styles)
    );
}

#[test]
fn mutable_document_replaces_and_clears_each_font_face_part_without_rewriting_neighbors() {
    let original_content = declarations("Content Body");
    let original_styles = declarations("Styles Body");
    let replacement_content = declarations("Content Replacement");
    let replacement_styles = declarations("Styles Replacement");
    let mut mutable = litchi_odt::mutable::MutableDocument::from_document(document()).unwrap();

    assert_eq!(
        mutable
            .set_content_font_face_declarations(&replacement_content)
            .unwrap(),
        Some(original_content)
    );
    assert_eq!(
        mutable
            .set_styles_font_face_declarations(&replacement_styles)
            .unwrap(),
        Some(original_styles)
    );

    let reopened = Document::from_bytes(mutable.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened.content_font_face_declarations().unwrap(),
        Some(replacement_content.clone())
    );
    assert_eq!(
        reopened.styles_font_face_declarations().unwrap(),
        Some(replacement_styles.clone())
    );

    let content_xml = String::from_utf8(reopened.get_file("content.xml").unwrap()).unwrap();
    assert!(content_xml.contains("<!--content-keep-->"));
    assert!(
        content_xml.find("<office:scripts/>").unwrap()
            < content_xml.find("<office:font-face-decls").unwrap()
    );
    assert!(
        content_xml.find("<office:font-face-decls").unwrap()
            < content_xml.find("<office:automatic-styles/>").unwrap()
    );
    let styles_xml = String::from_utf8(reopened.get_file("styles.xml").unwrap()).unwrap();
    assert!(styles_xml.contains("<!--styles-keep-->"));
    assert!(styles_xml.contains(r#"<style:style style:name="Keep" style:family="paragraph"/>"#));

    assert_eq!(
        mutable.clear_content_font_face_declarations().unwrap(),
        Some(replacement_content.clone())
    );
    assert_eq!(
        mutable.clear_styles_font_face_declarations().unwrap(),
        Some(replacement_styles.clone())
    );
    assert_eq!(mutable.content_font_face_declarations().unwrap(), None);
    assert_eq!(mutable.styles_font_face_declarations().unwrap(), None);

    assert_eq!(
        mutable
            .set_content_font_face_declarations(&replacement_content)
            .unwrap(),
        None
    );
    assert_eq!(
        mutable
            .set_styles_font_face_declarations(&replacement_styles)
            .unwrap(),
        None
    );
    assert_eq!(
        mutable.content_font_face_declarations().unwrap(),
        Some(replacement_content)
    );
    assert_eq!(
        mutable.styles_font_face_declarations().unwrap(),
        Some(replacement_styles)
    );

    let reinserted = Document::from_bytes(mutable.to_bytes().unwrap()).unwrap();
    let content_xml = String::from_utf8(reinserted.get_file("content.xml").unwrap()).unwrap();
    assert!(
        content_xml.find("<office:scripts/>").unwrap()
            < content_xml.find("<office:font-face-decls").unwrap()
    );
    assert!(
        content_xml.find("<office:font-face-decls").unwrap()
            < content_xml.find("<office:automatic-styles/>").unwrap()
    );
    let styles_xml = String::from_utf8(reinserted.get_file("styles.xml").unwrap()).unwrap();
    assert!(
        styles_xml.find("<office:font-face-decls").unwrap()
            < styles_xml.find("<office:styles>").unwrap()
    );
}

#[test]
fn flat_document_exposes_its_single_font_face_declarations_part() {
    let document =
        litchi_odt::generic::FlatOpenDocument::from_bytes(FLAT.as_bytes().to_vec()).unwrap();

    assert_eq!(
        document.font_face_declarations().unwrap(),
        Some(declarations("Flat Body"))
    );
}
