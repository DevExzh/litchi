use litchi_odf::{
    Document, MutableDocument, OdfBibliographyConfiguration, OdfBibliographyField,
    OdfBibliographySortKey, OpenDocumentPackage,
};
use std::io::{Cursor, Write};

const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const TEXT: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const MIMETYPE: &str = "application/vnd.oasis.opendocument.text";
const STYLES: &str =
    include_str!("../../../test-data/odf/odt/bibliography-configuration-styles.xml");

fn document(styles: &str) -> Document {
    let content = format!(
        r#"<o:document-content xmlns:o="{OFFICE}" xmlns:t="{TEXT}" o:version="1.3"><o:body><o:text><t:p>Body</t:p></o:text></o:body></o:document-content>"#
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

fn fixture_configuration() -> OdfBibliographyConfiguration {
    OdfBibliographyConfiguration {
        prefix: Some("[".to_string()),
        suffix: Some("]".to_string()),
        numbered_entries: Some(true),
        sort_by_position: Some(false),
        sort_algorithm: Some("unicode".to_string()),
        language: Some("en".to_string()),
        country: Some("US".to_string()),
        script: Some("Latn".to_string()),
        rfc_language_tag: Some("en-US".to_string()),
        sort_keys: vec![
            OdfBibliographySortKey {
                field: OdfBibliographyField::Author,
                ascending: Some(true),
            },
            OdfBibliographySortKey {
                field: OdfBibliographyField::Year,
                ascending: Some(false),
            },
            OdfBibliographySortKey {
                field: OdfBibliographyField::Isbn,
                ascending: None,
            },
        ],
    }
}

fn replacement_configuration() -> OdfBibliographyConfiguration {
    OdfBibliographyConfiguration {
        prefix: Some("(".to_string()),
        suffix: Some(")".to_string()),
        numbered_entries: Some(false),
        language: Some("de".to_string()),
        country: Some("DE".to_string()),
        rfc_language_tag: Some("de-DE".to_string()),
        sort_keys: vec![OdfBibliographySortKey {
            field: OdfBibliographyField::Title,
            ascending: Some(false),
        }],
        ..Default::default()
    }
}

#[test]
fn document_reads_styles_bibliography_configuration() {
    let expected = fixture_configuration();
    let source = document(STYLES);

    assert_eq!(
        source.bibliography_configuration().unwrap(),
        Some(expected.clone())
    );
    assert_eq!(
        source
            .variable_declarations()
            .unwrap()
            .bibliography_configuration,
        Some(expected.clone())
    );
    let generic = OpenDocumentPackage::from_bytes(source.to_bytes().unwrap()).unwrap();
    assert_eq!(
        generic.bibliography_configuration().unwrap(),
        Some(expected)
    );
}

#[test]
fn mutable_document_replaces_removes_and_inserts_bibliography_configuration() {
    let original = fixture_configuration();
    let replacement = replacement_configuration();
    let mut mutable = MutableDocument::from_document(document(STYLES)).unwrap();

    assert_eq!(
        mutable
            .set_bibliography_configuration(&replacement)
            .unwrap(),
        Some(original)
    );
    assert_eq!(
        mutable.bibliography_configuration().unwrap(),
        Some(replacement.clone())
    );

    let reopened = Document::from_bytes(mutable.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened.bibliography_configuration().unwrap(),
        Some(replacement.clone())
    );
    let styles = String::from_utf8(reopened.get_file("styles.xml").unwrap()).unwrap();
    assert!(styles.contains(r#"<style:style style:name="Keep" style:family="paragraph"/>"#));

    assert_eq!(
        mutable.clear_bibliography_configuration().unwrap(),
        Some(replacement.clone())
    );
    assert_eq!(mutable.bibliography_configuration().unwrap(), None);
    assert_eq!(
        mutable
            .set_bibliography_configuration(&replacement)
            .unwrap(),
        None
    );
    assert_eq!(
        mutable.bibliography_configuration().unwrap(),
        Some(replacement)
    );
}
