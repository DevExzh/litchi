//! Small XML wire primitives shared by chart record-family writers.

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
    write!(writer, "<c:{name}>")?;
    write_escaped_xml(writer, value)?;
    write!(writer, "</c:{name}>")
}

/// Stream an XML-escaped value into the sink without first allocating a
/// second copy of the value.  Chart text and relationship identifiers are
/// caller-owned strings, so materializing their escaped form before the
/// bounded writer sees it would let a large value bypass the selected output
/// and semantic budgets.
#[inline]
pub(super) fn write_escaped_xml<W: Write>(writer: &mut W, value: &str) -> std::io::Result<()> {
    let bytes = value.as_bytes();
    let mut cursor = 0usize;
    for (index, byte) in bytes.iter().enumerate() {
        let replacement = match byte {
            b'&' => "&amp;",
            b'<' => "&lt;",
            b'>' => "&gt;",
            b'"' => "&quot;",
            b'\'' => "&apos;",
            _ => continue,
        };
        writer.write_all(&bytes[cursor..index])?;
        writer.write_all(replacement.as_bytes())?;
        cursor = index + 1;
    }
    writer.write_all(&bytes[cursor..])
}

/// Append an already validated `DrawingML` fragment without reformatting it.
#[inline]
pub(super) fn write_fragment<W: Write>(writer: &mut W, fragment: &[u8]) -> std::io::Result<()> {
    writer.write_all(fragment)
}
