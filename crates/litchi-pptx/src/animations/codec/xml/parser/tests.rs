#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

use super::{parse_processed_timing, parse_recursive_timing_tree};

const PRESENTATIONML_NS: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
const STRICT_PRESENTATIONML_NS: &str = "http://purl.oclc.org/ooxml/presentationml/main";

#[test]
fn processed_parser_accepts_transitional_namespace_without_timing() {
    let xml = format!(r#"<p:sld xmlns:p="{PRESENTATIONML_NS}"><p:cSld/></p:sld>"#);
    let sequence = parse_processed_timing(xml.as_bytes(), true).expect("parse slide");
    assert!(sequence.is_empty());
}

#[test]
fn recursive_parser_preserves_strict_timing_snapshot() {
    let xml =
        format!(r#"<p:sld xmlns:p="{STRICT_PRESENTATIONML_NS}"><p:timing></p:timing></p:sld>"#);
    let tree = parse_recursive_timing_tree(&xml).expect("parse timing tree");
    assert!(tree.roots.is_empty());
    assert_eq!(tree.source_xml.as_deref(), Some("<p:timing></p:timing>"));
}

#[test]
fn processed_parser_rejects_doctype() {
    let xml = format!(r#"<!DOCTYPE sld><p:sld xmlns:p="{PRESENTATIONML_NS}"/>"#);
    let error = parse_processed_timing(xml.as_bytes(), false).expect_err("reject doctype");
    assert!(error.to_string().contains("DOCTYPE"));
}
