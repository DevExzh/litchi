//! Demonstrates compile-time XML minification of an inline string literal.
//!
//! Run with:
//! ```sh
//! cargo run -p xml-minifier --example minify_inline
//! ```
//!
//! The input is a compact raw string literal containing an authoring comment
//! and an explicit empty-element pair. The macro normalizes those constructs
//! at compile time without deleting character data.

#![allow(
    clippy::print_stdout,
    reason = "the command-line example exists to display generated XML"
)]

use xml_minifier::minified_xml_str;

const INPUT: &str = r#"<?xml version="1.0" encoding="UTF-8"?><!-- A sample document --><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><empty></empty></Types>"#;

const MINIFIED: &str = minified_xml_str!(
    r#"<?xml version="1.0" encoding="UTF-8"?><!-- A sample document --><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><empty></empty></Types>"#
);

fn main() {
    println!("Input length:    {} bytes", INPUT.len());
    println!("Minified length: {} bytes", MINIFIED.len());
    println!("Minified XML:");
    println!("{MINIFIED}");
}
