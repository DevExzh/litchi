//! DrawingML transform primitives.

use std::{fmt, fmt::Write as _};

/// Write a transform containing an offset and extent.
pub fn write_offset_extent(
    xml: &mut String,
    x: i64,
    y: i64,
    width: i64,
    height: i64,
) -> fmt::Result {
    xml.push_str("<a:xfrm>");
    write!(xml, r#"<a:off x="{x}" y="{y}"/>"#)?;
    write!(xml, r#"<a:ext cx="{width}" cy="{height}"/>"#)?;
    xml.push_str("</a:xfrm>");
    Ok(())
}
