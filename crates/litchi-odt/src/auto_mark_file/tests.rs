//! Regression tests for the auto-mark-file model and codec.

use super::{AlphabeticalIndexAutoMarkFile, parse_auto_mark_file_parts};
use crate::variable_declaration::{Body, Part, Scope};
use litchi_core::Result;

const CONTENT_PREFIX: &str = r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:xlink="http://www.w3.org/1999/xlink"><o:body><o:text>"#;
const CONTENT_SUFFIX: &str = r"</o:text></o:body></o:document-content>";

fn parse(content: &str) -> Result<Vec<AlphabeticalIndexAutoMarkFile>> {
    let xml = format!("{CONTENT_PREFIX}{content}{CONTENT_SUFFIX}");
    parse_auto_mark_file_parts(&[(xml.as_str(), Part::Content)])
}

#[test]
fn parses_inert_auto_mark_file_reference() {
    let references =
        parse(r#"<t:alphabetical-index-auto-mark-file xlink:type="simple" xlink:href="concordance.sdi"/><t:p>Text</t:p>"#).unwrap();
    assert_eq!(references.len(), 1);
    let reference = &references[0];
    assert_eq!(reference.part, Part::Content);
    assert_eq!(reference.scope, Scope::Body(Body::Text));
    assert_eq!(reference.href, "concordance.sdi");
}

#[test]
fn accepts_explicit_empty_element_pair() {
    let references =
        parse(r#"<t:alphabetical-index-auto-mark-file xlink:type="simple" xlink:href="a.sdi"></t:alphabetical-index-auto-mark-file>"#)
            .unwrap();
    assert_eq!(references[0].href, "a.sdi");
}

#[test]
fn absent_reference_yields_empty_result() {
    assert!(parse("<t:p>No reference</t:p>").unwrap().is_empty());
}

#[test]
fn rejects_missing_or_invalid_xlink_metadata() {
    assert!(parse(r#"<t:alphabetical-index-auto-mark-file xlink:type="simple"/>"#).is_err());
    assert!(
        parse(r#"<t:alphabetical-index-auto-mark-file xlink:type="simple" xlink:href=""/>"#)
            .is_err()
    );
    assert!(parse(r#"<t:alphabetical-index-auto-mark-file xlink:href="a.sdi"/>"#).is_err());
    assert!(
        parse(r#"<t:alphabetical-index-auto-mark-file xlink:type="extended" xlink:href="a.sdi"/>"#)
            .is_err()
    );
    assert!(
        parse(r#"<t:alphabetical-index-auto-mark-file xlink:type="simple" xlink:href="a.sdi" o:track-changes="true"/>"#)
            .is_err()
    );
}

#[test]
fn rejects_duplicates_content_and_children() {
    let reference =
        r#"<t:alphabetical-index-auto-mark-file xlink:type="simple" xlink:href="a.sdi"/>"#;
    assert!(parse(&format!("{reference}{reference}")).is_err());
    assert!(parse(&format!("<t:p>Content</t:p>{reference}")).is_err());
    assert!(
        parse(r#"<t:alphabetical-index-auto-mark-file xlink:type="simple" xlink:href="a.sdi">text</t:alphabetical-index-auto-mark-file>"#)
            .is_err()
    );
    assert!(
        parse(r#"<t:alphabetical-index-auto-mark-file xlink:type="simple" xlink:href="a.sdi"><t:p/></t:alphabetical-index-auto-mark-file>"#)
            .is_err()
    );
}

#[test]
fn rejects_misplaced_or_spoofed_elements() {
    // Outside office:text.
    let floating = r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:xlink="http://www.w3.org/1999/xlink"><o:body><t:alphabetical-index-auto-mark-file xlink:type="simple" xlink:href="a.sdi"/></o:body></o:document-content>"#;
    assert!(parse_auto_mark_file_parts(&[(floating, Part::Content)]).is_err());
    // Nested inside document content.
    assert!(
        parse(r#"<t:p><t:alphabetical-index-auto-mark-file xlink:type="simple" xlink:href="a.sdi"/></t:p>"#)
            .is_err()
    );
    // Wrong namespace spelling of the element.
    assert!(
        parse(r#"<o:alphabetical-index-auto-mark-file xlink:type="simple" xlink:href="a.sdi"/>"#)
            .is_err()
    );
}
