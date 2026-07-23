use litchi_odf::{
    Document, FlatOpenDocument, MutableDocument, OdfFontFace, OdfFontFaceDeclarations,
    OpenDocumentPackage,
};
use std::io::{Cursor, Write};

const CONTENT: &str = include_str!("../../../test-data/odf/odt/font-face-declarations-content.xml");
const STYLES: &str = include_str!("../../../test-data/odf/odt/font-face-declarations-styles.xml");
const FLAT: &str = include_str!("../../../test-data/odf/odt/font-face-declarations-flat.fodt");
const MIMETYPE: &str = "application/vnd.oasis.opendocument.text";

fn declarations(name: &str) -> OdfFontFaceDeclarations {
    OdfFontFaceDeclarations {
        faces: vec![OdfFontFace {
            name: name.to_string(),
            family: Some(format!("'{name}'")),
            ..Default::default()
        }],
    }
}

fn document() -> Document {
    let mut output = Cursor::new(Vec::new());
    let mut zip = zip::ZipWriter::new(&mut output);
    let stored =
        zip::write::SimpleFileOptions::default().compression_method(zip::CompressionMethod::Stored);
    let deflated = zip::write::SimpleFileOptions::default()
        .compression_method(zip::CompressionMethod::Deflated);
    zip.start_file("mimetype", stored).unwrap();
    zip.write_all(MIMETYPE.as_bytes()).unwrap();
    zip.start_file("content.xml", deflated).unwrap();
    zip.write_all(CONTENT.as_bytes()).unwrap();
    zip.start_file("styles.xml", deflated).unwrap();
    zip.write_all(STYLES.as_bytes()).unwrap();
    zip.start_file("META-INF/manifest.xml", deflated).unwrap();
    write!(
        zip,
        r#"<m:manifest xmlns:m="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" m:version="1.3"><m:file-entry m:full-path="/" m:media-type="{MIMETYPE}"/><m:file-entry m:full-path="content.xml" m:media-type="text/xml"/><m:file-entry m:full-path="styles.xml" m:media-type="text/xml"/></m:manifest>"#
    )
    .unwrap();
    zip.finish().unwrap();
    Document::from_bytes(output.into_inner()).unwrap()
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

    let package = OpenDocumentPackage::from_bytes(source.to_bytes().unwrap()).unwrap();
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
    let mut mutable = MutableDocument::from_document(document()).unwrap();

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
    let document = FlatOpenDocument::from_bytes(FLAT.as_bytes().to_vec()).unwrap();

    assert_eq!(
        document.font_face_declarations().unwrap(),
        Some(declarations("Flat Body"))
    );
}
