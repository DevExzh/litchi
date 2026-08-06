//! Prefix-free annotation serialization.

use super::model::Annotation;
use litchi_core::{Result, xml::escape_xml};

const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const TEXT: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const TABLE: &str = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
const DRAW: &str = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
const SVG: &str = "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0";
const DC: &str = "http://purl.org/dc/elements/1.1/";
const META: &str = "urn:oasis:names:tc:opendocument:xmlns:meta:1.0";
const XLINK: &str = "http://www.w3.org/1999/xlink";
const LOEXT: &str = "urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0";

pub(crate) fn serialize(annotation: &Annotation) -> Result<String> {
    let mut annotation = annotation.clone();
    for (prefix, uri) in [
        ("office", OFFICE),
        ("text", TEXT),
        ("table", TABLE),
        ("draw", DRAW),
        ("svg", SVG),
        ("dc", DC),
        ("meta", META),
        ("xlink", XLINK),
        ("loext", LOEXT),
    ] {
        annotation.set_namespace(prefix, uri)?;
    }
    annotation.validate()?;
    let mut output = String::new();
    annotation.write_xml(&mut output);
    Ok(output)
}

pub(crate) fn end_marker(name: &str) -> String {
    format!(
        "<office:annotation-end xmlns:office=\"{OFFICE}\" office:name=\"{}\"/>",
        escape_xml(name)
    )
}
