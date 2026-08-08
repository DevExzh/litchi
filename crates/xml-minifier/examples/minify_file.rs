//! Demonstrates compile-time XML minification of a file.
//!
//! Run with:
//! ```sh
//! cargo run -p xml-minifier --example minify_file
//! ```
//!
//! The XML file (`sample.xml`) sits next to this source file and is read,
//! minified, and embedded as a `&'static str` at compile time.

#![allow(
    clippy::print_stdout,
    reason = "the command-line example exists to display generated XML"
)]

use xml_minifier::minified_xml;

const MINIFIED: &str = minified_xml!("sample.xml");

fn main() {
    println!("Minified XML ({} bytes):", MINIFIED.len());
    println!("{MINIFIED}");
}
