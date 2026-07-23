use litchi_odf::{Document, MutableDocument};
use std::io::{Cursor, Write};

const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const TEXT: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const MIMETYPE: &str = "application/vnd.oasis.opendocument.text";
const STYLES: &str = include_str!("../../../test-data/odf/odt/outline-style-styles.xml");

fn document(styles: &str) -> Document {
    let content = format!(
        r#"<o:document-content xmlns:o="{OFFICE}" xmlns:t="{TEXT}" o:version="1.3"><o:body><o:text><t:h t:outline-level="1">Heading</t:h></o:text></o:body></o:document-content>"#
    );
    let mut output = Cursor::new(Vec::new());
    let mut zip = zip::ZipWriter::new(&mut output);
    let stored =
        zip::write::SimpleFileOptions::default().compression_method(zip::CompressionMethod::Stored);
    let deflated = zip::write::SimpleFileOptions::default()
        .compression_method(zip::CompressionMethod::Deflated);
    zip.start_file("mimetype", stored).unwrap();
    zip.write_all(MIMETYPE.as_bytes()).unwrap();
    zip.start_file("content.xml", deflated).unwrap();
    zip.write_all(content.as_bytes()).unwrap();
    zip.start_file("styles.xml", deflated).unwrap();
    zip.write_all(styles.as_bytes()).unwrap();
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
fn document_reads_outline_styles_from_styles_metadata() {
    let source = document(STYLES);
    let styles = source.outline_styles().unwrap();
    let outline = styles.get("Outline").unwrap();

    assert_eq!(styles.styles.len(), 1);
    assert_eq!(outline.levels.len(), 1);
    assert_eq!(
        outline
            .level(1)
            .unwrap()
            .number_format
            .as_ref()
            .unwrap()
            .as_str(),
        "1"
    );
}

#[test]
fn mutable_document_manages_outline_styles_without_rewriting_other_styles() {
    let source = document(STYLES);
    let original = source
        .outline_styles()
        .unwrap()
        .get("Outline")
        .unwrap()
        .clone();
    let mut replacement = original.clone();
    replacement.levels[0].number_prefix = Some("(".to_string());
    replacement.levels[0].number_suffix = Some(")".to_string());
    let mut mutable = MutableDocument::from_document(source).unwrap();

    assert_eq!(
        mutable.set_outline_style(&replacement).unwrap(),
        Some(original.clone())
    );
    let styles = mutable.outline_styles().unwrap();
    assert_eq!(styles.get("Outline"), Some(&replacement));

    let reopened = Document::from_bytes(mutable.to_bytes().unwrap()).unwrap();
    let reopened_styles = reopened.outline_styles().unwrap();
    assert_eq!(reopened_styles.get("Outline"), Some(&replacement));
    let styles_xml = String::from_utf8(reopened.get_file("styles.xml").unwrap()).unwrap();
    assert!(styles_xml.contains("<!--keep-->"));
    assert!(styles_xml.contains(r#"<style:style style:name="Keep" style:family="paragraph"/>"#));

    assert_eq!(
        mutable.remove_outline_style("Outline").unwrap(),
        Some(replacement)
    );
    assert!(mutable.outline_styles().unwrap().get("Outline").is_none());
    assert_eq!(mutable.set_outline_style(&original).unwrap(), None);
    assert_eq!(
        mutable.outline_styles().unwrap().get("Outline"),
        Some(&original)
    );
}
