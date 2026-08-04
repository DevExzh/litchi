use litchi_odt::{
    BibliographyConfiguration, BibliographyField, BibliographySortKey, Document, MutableDocument,
    OpenDocumentPackage,
};
mod support;

const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const TEXT: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const MIMETYPE: &str = "application/vnd.oasis.opendocument.text";
const STYLES: &str =
    include_str!("../../../test-data/odf/odt/bibliography-configuration-styles.xml");

fn document(styles: &str) -> Document {
    let content = format!(
        r#"<o:document-content xmlns:o="{OFFICE}" xmlns:t="{TEXT}" o:version="1.3"><o:body><o:text><t:p>Body</t:p></o:text></o:body></o:document-content>"#
    );
    Document::from_bytes(support::package(
        MIMETYPE,
        &[
            ("content.xml", content.as_bytes()),
            ("styles.xml", styles.as_bytes()),
        ],
    ))
    .unwrap()
}

fn fixture_configuration() -> BibliographyConfiguration {
    BibliographyConfiguration {
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
            BibliographySortKey {
                field: BibliographyField::Author,
                ascending: Some(true),
            },
            BibliographySortKey {
                field: BibliographyField::Year,
                ascending: Some(false),
            },
            BibliographySortKey {
                field: BibliographyField::Isbn,
                ascending: None,
            },
        ],
    }
}

fn replacement_configuration() -> BibliographyConfiguration {
    BibliographyConfiguration {
        prefix: Some("(".to_string()),
        suffix: Some(")".to_string()),
        numbered_entries: Some(false),
        language: Some("de".to_string()),
        country: Some("DE".to_string()),
        rfc_language_tag: Some("de-DE".to_string()),
        sort_keys: vec![BibliographySortKey {
            field: BibliographyField::Title,
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
