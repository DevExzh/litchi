//! DrawingML fill primitives.

/// Write a stretch fill covering the complete rectangle.
#[inline]
pub fn write_stretch_rect(xml: &mut String) {
    xml.push_str("<a:stretch><a:fillRect/></a:stretch>");
}
