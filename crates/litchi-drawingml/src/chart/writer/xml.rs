//! Small XML wire primitives shared by chart record-family writers.

use litchi_core::xml::escape_xml;
use std::io::Write;

/// Write a chart boolean element using the schema's `0`/`1` lexical form.
#[inline]
pub(super) fn write_bool_element<W: Write>(
    writer: &mut W,
    name: &str,
    value: bool,
) -> std::io::Result<()> {
    write!(
        writer,
        r#"<c:{name} val="{}"/>"#,
        if value { "1" } else { "0" }
    )
}

/// Write a chart text element with XML escaping.
#[inline]
pub(super) fn write_text_element<W: Write>(
    writer: &mut W,
    name: &str,
    value: &str,
) -> std::io::Result<()> {
    write!(writer, "<c:{name}>{}</c:{name}>", escape_xml(value))
}

/// Append an already validated `DrawingML` fragment without reformatting it.
#[inline]
pub(super) fn write_fragment<W: Write>(writer: &mut W, fragment: &[u8]) -> std::io::Result<()> {
    writer.write_all(fragment)
}
