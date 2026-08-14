//! Compact corpus and malformed-input hardening for the ODT parser facade.

use litchi_core::Error;
use litchi_odt::{Document, core::PackageWriter};
mod support;

const CONTENT_HEAD: &str = r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:form="urn:oasis:names:tc:opendocument:xmlns:form:1.0" xmlns:dc="http://purl.org/dc/elements/1.1/" office:version="1.3"><office:body><office:text>"#;
const CONTENT_TAIL: &str = "</office:text></office:body></office:document-content>";
const FEATURE_BODY: &str = r#"<text:page-sequence><text:page text:master-page-name="Standard"/></text:page-sequence><text:tracked-changes><text:changed-region text:id="c1"><text:insertion><office:change-info><dc:date>2024-01-01T00:00:00</dc:date></office:change-info><text:p>added</text:p></text:insertion></text:changed-region></text:tracked-changes><office:forms><form:form form:name="f"><form:text form:name="Name" xml:id="name_field" form:current-value="Ada"/></form:form></office:forms><text:h text:outline-level="1">Title</text:h><text:p>body<text:bookmark text:name="mark"/><text:note text:note-class="footnote"><text:note-citation>1</text:note-citation><text:note-body><text:p>footnote</text:p></text:note-body></text:note></text:p><text:section text:name="s1" text:protected="true" text:protection-key="aGk="><text:p>secret</text:p></text:section><text:p><text:ruby><text:ruby-base>base</text:ruby-base><text:ruby-text>reading</text:ruby-text></text:ruby><text:conditional-text text:condition="of:=1=1" text:string-value-if-true="yes" text:string-value-if-false="no">yes</text:conditional-text></text:p><table:table table:name="T"><table:table-row><table:table-cell office:value-type="string"><text:p>cell</text:p></table:table-cell></table:table-row></table:table><text:table-of-content text:name="toc"><text:table-of-content-source/><text:index-body><text:p>entry</text:p></text:index-body></text:table-of-content>"#;
const STYLES: &str = r#"<?xml version="1.0" encoding="UTF-8"?><office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" office:version="1.3"><office:automatic-styles><style:page-layout style:name="pm1"/></office:automatic-styles><office:master-styles><style:master-page style:name="Standard" style:page-layout-name="pm1"><style:header><text:p>Header</text:p></style:header></style:master-page></office:master-styles></office:document-styles>"#;
const META: &str = r#"<?xml version="1.0" encoding="UTF-8"?><office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:dc="http://purl.org/dc/elements/1.1/" office:version="1.3"><office:meta><dc:title>Compact corpus</dc:title></office:meta></office:document-meta>"#;

fn content(body: &str) -> Vec<u8> {
    format!("{CONTENT_HEAD}{body}{CONTENT_TAIL}").into_bytes()
}

fn package(
    content_xml: &[u8],
    styles_xml: Option<&str>,
    meta_xml: Option<&str>,
) -> Result<Vec<u8>, Error> {
    let mut writer = PackageWriter::new();
    writer.set_mimetype("application/vnd.oasis.opendocument.text")?;
    writer.add_file("content.xml", content_xml)?;
    if let Some(styles_xml) = styles_xml {
        writer.add_file("styles.xml", styles_xml.as_bytes())?;
    }
    if let Some(meta_xml) = meta_xml {
        writer.add_file("meta.xml", meta_xml.as_bytes())?;
    }
    writer.finish_to_bytes()
}

fn exercise_odt(content_xml: &[u8]) -> Option<Error> {
    let bytes = match package(content_xml, None, None) {
        Ok(bytes) => bytes,
        Err(error) => return Some(error),
    };
    let document = match Document::from_bytes(bytes) {
        Ok(document) => document,
        Err(error) => return Some(error),
    };
    for result in [
        document.text().map(|_| ()),
        document.paragraphs().map(|_| ()),
        document.tables().map(|_| ()),
        document.sections().map(|_| ()),
        document.forms().map(|_| ()),
        document.page_sequence().map(|_| ()),
        document.tracked_changes().map(|_| ()),
        document.dynamic_text_fields().map(|_| ()),
        document.text_indexes().map(|_| ()),
        document.master_pages().map(|_| ()),
        document.ruby_annotations().map(|_| ()),
    ] {
        if let Err(error) = result {
            return Some(error);
        }
    }
    None
}

fn assert_typed_error(content_xml: &[u8], case: &str) {
    match exercise_odt(content_xml) {
        Some(Error::InvalidFormat(_)) => {},
        Some(error) => panic!("{case} produced a non-format error: {error:?}"),
        None => panic!("{case} unexpectedly parsed"),
    }
}

#[test]
fn corpus_text_documents_parse_and_extract() -> Result<(), Error> {
    let bytes = package(&content(FEATURE_BODY), Some(STYLES), Some(META))?;
    let document = Document::from_bytes(bytes)?;
    assert!(!document.text()?.trim().is_empty());
    assert!(document.paragraphs()?.len() >= 5);
    assert_eq!(document.tables()?.len(), 1);
    assert_eq!(document.text_indexes()?.len(), 1);
    assert!(!document.master_pages()?.is_empty());
    assert_eq!(document.forms()?.groups.len(), 1);
    assert_eq!(document.tracked_changes()?.changes.len(), 1);
    assert_eq!(document.sections()?.len(), 1);
    assert_eq!(document.ruby_annotations()?.annotations.len(), 1);
    drop(document.metadata()?);
    Ok(())
}

#[test]
fn corpus_xxe_probe_is_rejected_without_resolution() -> Result<(), Error> {
    let xml = format!(
        r#"<?xml version="1.0" encoding="UTF-8"?><!DOCTYPE office [<!ENTITY nasty SYSTEM "externalcontent.txt">]>{CONTENT_HEAD}<text:p>&nasty;</text:p>{CONTENT_TAIL}"#
    );
    assert!(package(xml.as_bytes(), None, Some(META)).is_err());
    let raw = support::package(
        "application/vnd.oasis.opendocument.text",
        &[
            ("content.xml", xml.as_bytes()),
            ("meta.xml", META.as_bytes()),
        ],
    );
    let document = match Document::from_bytes(raw) {
        Ok(document) => document,
        Err(Error::InvalidFormat(_)) => return Ok(()),
        Err(error) => panic!("open produced a non-format error: {error:?}"),
    };
    match document.text() {
        Err(Error::InvalidFormat(_)) => {},
        Err(error) => panic!("text produced a non-format error: {error:?}"),
        Ok(_) => panic!("XXE entity unexpectedly resolved in text"),
    }
    match document.paragraphs() {
        Err(Error::InvalidFormat(_)) => {},
        Err(error) => panic!("paragraphs produced a non-format error: {error:?}"),
        Ok(_) => panic!("XXE entity unexpectedly resolved in paragraphs"),
    }
    match document.forms() {
        Err(Error::InvalidFormat(_)) => {},
        Err(error) => panic!("forms produced a non-format error: {error:?}"),
        Ok(_) => panic!("DOCTYPE unexpectedly parsed in forms"),
    }
    drop(document.metadata()?);
    Ok(())
}

#[test]
fn odt_truncation_sweep_never_panics() {
    let seed = content(FEATURE_BODY);
    for end in 0..seed.len() {
        drop(exercise_odt(&seed[..end]));
    }
    assert!(exercise_odt(&seed).is_none());
}

#[test]
fn odt_byte_mutation_sweep_never_panics() {
    let seed = content(FEATURE_BODY);
    for position in 0..seed.len() {
        let mut mutated = seed.clone();
        mutated[position] ^= 1;
        drop(exercise_odt(&mutated));
    }
    assert!(exercise_odt(&seed).is_none());
}

#[test]
fn odt_malformed_inputs_yield_typed_errors() {
    assert_typed_error(
        &content(r#"<text:section text:name="s"><text:p>open"#),
        "unterminated section",
    );
    assert_typed_error(
        &content(
            r#"<text:p><text:page-sequence><text:page text:master-page-name="A"/></text:page-sequence></text:p>"#,
        ),
        "misplaced page sequence",
    );
    assert_typed_error(
        &content(
            r#"<text:page-sequence text:style-name="x"><text:page text:master-page-name="A"/></text:page-sequence>"#,
        ),
        "page sequence attributes",
    );
    assert_typed_error(
        &content(
            r#"<office:forms><form:form form:name="f"><form:text form:name="a" xml:id="dup"/><form:button form:name="b" xml:id="dup"/></form:form></office:forms>"#,
        ),
        "duplicate control id",
    );
    assert_typed_error(
        &content(
            r#"<text:section text:name="s" text:protected="maybe"><text:p>x</text:p></text:section>"#,
        ),
        "invalid section boolean",
    );
    assert!(exercise_odt(&content(r#"<text:section text:name="s" text:protection-key="aGk="><text:p>x</text:p></text:section>"#)).is_none());
    assert_typed_error(
        &content(r#"<text:section text:name="s"><text:p>x</text:p></text:p>"#),
        "mismatched section close",
    );
}

#[test]
fn odt_misplaced_but_wellformed_inputs_do_not_panic() {
    let nested = format!(
        "{}{}",
        "<text:span>".repeat(200),
        "</text:span>".repeat(200)
    );
    drop(exercise_odt(&content(&nested)));
    drop(exercise_odt(&content(
        r#"<text:p><foo:bar xmlns:foo="urn:example"/>text</text:p>"#,
    )));
    drop(exercise_odt(&content("")));
    let mut invalid_utf8 = content("<text:p>x</text:p>");
    invalid_utf8.insert(invalid_utf8.len() / 2, 0xfe);
    assert_typed_error(&invalid_utf8, "invalid UTF-8");
}
