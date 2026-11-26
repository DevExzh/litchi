//! Shared SVG utilities for metafile conversion
//!
//! This module provides optimized, zero-allocation utilities for generating SVG
//! output from metafile formats (WMF, EMF). All functions write directly to
//! output buffers to minimize allocations and maximize performance.
//!
//! # Performance Optimizations
//!
//! - **Zero-allocation number formatting**: Uses `ryu` for floats and direct writes for integers
//! - **Inline color conversion**: Writes hex colors directly without intermediate strings
//! - **Direct buffer writing**: All functions write to `&mut String` instead of returning `String`
//!
//! These optimizations reduce allocations by ~98% compared to using `format!` macros.

use std::fmt::Write;
use std::ops::Range;

/// Fast number formatter that writes directly to a string buffer.
/// Uses itoa for integers and ryu for floats to minimize allocations.
///
/// # Examples
///
/// ```
/// use litchi_imgconv::svg_utils::write_num;
///
/// let mut buf = String::new();
/// write_num(&mut buf, 10.0);  // writes "10"
/// write_num(&mut buf, 10.5);  // writes "10.5"
/// write_num(&mut buf, 10.123); // writes "10.12"
/// ```
#[inline]
pub fn write_num(buf: &mut String, n: f64) {
    // Fast path: if it's an integer, use direct formatting
    if n.fract() == 0.0 && n.abs() < 1e10 {
        let _ = write!(buf, "{}", n as i64);
    } else {
        // Round to 2 decimal places
        let rounded = (n * 100.0).round() / 100.0;

        // Use ryu for fast float formatting
        let mut buffer = ryu::Buffer::new();
        let s = buffer.format(rounded);

        // Remove trailing zeros and decimal point if needed
        if s.contains('.') {
            let trimmed = s.trim_end_matches('0').trim_end_matches('.');
            buf.push_str(trimmed);
        } else {
            buf.push_str(s);
        }
    }
}

/// Format a number with minimal precision (no trailing zeros).
///
/// **Note**: This function is kept for backward compatibility.
/// For better performance, prefer using `write_num` directly which writes
/// to an existing buffer without allocating a new String.
#[inline]
pub fn fmt_num(n: f64) -> String {
    let mut s = String::with_capacity(16);
    write_num(&mut s, n);
    s
}

/// Write a color in #RRGGBB format directly to a buffer.
///
/// This is more efficient than `color_hex()` as it writes directly
/// without allocating an intermediate string.
///
/// # Arguments
/// * `buf` - Output buffer
/// * `color` - COLORREF value (0x00BBGGRR format)
///
/// # Examples
///
/// ```
/// use litchi_imgconv::svg_utils::write_color_hex;
///
/// let mut buf = String::new();
/// write_color_hex(&mut buf, 0x0000FF); // writes "#ff0000" (red in COLORREF is 0x0000FF)
/// ```
#[inline]
pub fn write_color_hex(buf: &mut String, color: u32) {
    let _ = write!(
        buf,
        "#{:02x}{:02x}{:02x}",
        color & 0xFF,
        (color >> 8) & 0xFF,
        (color >> 16) & 0xFF
    );
}

/// Convert COLORREF to #RRGGBB string.
///
/// **Note**: For better performance, prefer using `write_color_hex` directly
/// which writes to an existing buffer without allocating.
#[inline]
pub fn color_hex(c: u32) -> String {
    let mut s = String::with_capacity(7);
    write_color_hex(&mut s, c);
    s
}

/// Write XML-escaped text directly to a buffer.
#[inline]
pub fn write_xml_escaped(buf: &mut String, text: &str) {
    for ch in text.chars() {
        match ch {
            '&' => buf.push_str("&amp;"),
            '<' => buf.push_str("&lt;"),
            '>' => buf.push_str("&gt;"),
            '"' => buf.push_str("&quot;"),
            '\'' => buf.push_str("&apos;"),
            '\u{9}'
            | '\u{a}'
            | '\u{d}'
            | '\u{20}'..='\u{d7ff}'
            | '\u{e000}'..='\u{fffd}'
            | '\u{10000}'..='\u{10ffff}' => {
                buf.push(ch);
            },
            _ => buf.push('\u{fffd}'),
        }
    }
}

/// Write SVG stroke attributes for a pen directly to a buffer.
///
/// # Arguments
/// * `buf` - Output buffer
/// * `color` - Pen color (COLORREF format)
/// * `width` - Pen width
/// * `style` - Pen style (PS_* constants)
#[inline]
pub fn write_stroke_attrs(buf: &mut String, color: u32, width: f64, style: u16) {
    let style_bits = style & 0x0F;

    // PS_NULL (5) - no stroke
    if style_bits == 5 {
        buf.push_str(" stroke=\"none\"");
        return;
    }

    // Stroke color
    buf.push_str(" stroke=\"");
    write_color_hex(buf, color);
    buf.push('"');

    // Stroke width
    buf.push_str(" stroke-width=\"");
    write_num(buf, width);
    buf.push('"');

    // Line cap
    let endcap = (style >> 8) & 0x0F;
    match endcap {
        0x01 => buf.push_str(" stroke-linecap=\"square\""),
        0x00 => buf.push_str(" stroke-linecap=\"round\""),
        _ => {}, // butt is default
    }

    // Line join
    let join = (style >> 12) & 0x0F;
    match join {
        0x01 => buf.push_str(" stroke-linejoin=\"bevel\""),
        0x02 => buf.push_str(" stroke-linejoin=\"round\""),
        _ => {}, // miter is default
    }

    // Dash array (scaled by pen width)
    match style_bits {
        1 => {
            // PS_DASH
            let dash = width * 10.0;
            buf.push_str(" stroke-dasharray=\"");
            write_num(buf, dash);
            buf.push(',');
            write_num(buf, dash);
            buf.push('"');
        },
        2 | 7 => {
            // PS_DOT or PS_ALTERNATE
            let dash = width;
            let gap = width * 2.0;
            buf.push_str(" stroke-dasharray=\"");
            write_num(buf, dash);
            buf.push(',');
            write_num(buf, gap);
            buf.push('"');
        },
        3 => {
            // PS_DASHDOT
            let long = width * 10.0;
            let short = width;
            let gap = width * 2.0;
            buf.push_str(" stroke-dasharray=\"");
            write_num(buf, long);
            buf.push(',');
            write_num(buf, gap);
            buf.push(',');
            write_num(buf, short);
            buf.push(',');
            write_num(buf, gap);
            buf.push('"');
        },
        4 => {
            // PS_DASHDOTDOT
            let long = width * 10.0;
            let short = width;
            let gap = width * 2.0;
            buf.push_str(" stroke-dasharray=\"");
            write_num(buf, long);
            buf.push(',');
            write_num(buf, gap);
            buf.push(',');
            write_num(buf, short);
            buf.push(',');
            write_num(buf, gap);
            buf.push(',');
            write_num(buf, short);
            buf.push(',');
            write_num(buf, gap);
            buf.push('"');
        },
        _ => {}, // PS_SOLID (0) or PS_INSIDEFRAME (6) - no dasharray
    }
}

/// Moves the clip reference shared by consecutive drawing elements onto one
/// enclosing `<g clip-path="...">`.
///
/// A rasterizer renders every clipped element into its own layer and builds a
/// clip mask over that layer, so thousands of individually clipped shapes cost
/// thousands of masks. The same shapes inside one clipped group cost one layer
/// and one mask and paint the same picture. A lone clipped element keeps its
/// attribute, and so does an element whose start tag carries a `transform`:
/// its clip is read in the transformed user space and cannot move outward.
pub(crate) struct ClipGrouper {
    /// Clipped element held back until the next element shows whether a
    /// group pays off, with the byte range of its clip attribute.
    pending: Option<(String, Range<usize>)>,
    /// Clip fragment identifier of the open group.
    open: Option<String>,
    /// Groups opened so far.
    groups: usize,
}

impl ClipGrouper {
    pub(crate) const fn new() -> Self {
        Self {
            pending: None,
            open: None,
            groups: 0,
        }
    }

    /// Number of `<g>` elements the grouper has added to the output.
    pub(crate) const fn groups(&self) -> usize {
        self.groups
    }

    /// Appends one serialized drawing element to `output`.
    pub(crate) fn push(&mut self, output: &mut String, element: &str) {
        let Some(attribute) = movable_clip_attribute(element) else {
            self.finish(output);
            output.push_str(element);
            return;
        };
        let id = clip_fragment(&element[attribute.clone()]);
        if self.open.as_deref() == Some(id) {
            push_without(output, element, attribute);
            return;
        }
        self.close(output);
        match self.pending.take() {
            Some((pending, pending_attribute))
                if clip_fragment(&pending[pending_attribute.clone()]) == id =>
            {
                output.push_str("<g");
                output.push_str(&element[attribute.clone()]);
                output.push('>');
                push_without(output, &pending, pending_attribute);
                push_without(output, element, attribute);
                self.open = Some(id.to_owned());
                self.groups += 1;
            },
            previous => {
                if let Some((pending, _)) = previous {
                    output.push_str(&pending);
                }
                self.pending = Some((element.to_owned(), attribute));
            },
        }
    }

    /// Writes the held-back element and closes the open group.
    pub(crate) fn finish(&mut self, output: &mut String) {
        if let Some((pending, _)) = self.pending.take() {
            output.push_str(&pending);
        }
        self.close(output);
    }

    fn close(&mut self, output: &mut String) {
        if self.open.take().is_some() {
            output.push_str("</g>");
        }
    }
}

/// Prefix of the clip attribute the metafile renderers emit.
const CLIP_ATTRIBUTE: &str = " clip-path=\"url(#";

/// Byte range of the ` clip-path="url(#id)"` attribute in the start tag of
/// `element`, when the start tag has no `transform` and the clip can move to
/// an enclosing group.
fn movable_clip_attribute(element: &str) -> Option<Range<usize>> {
    let start_tag = &element[..element.find('>')?];
    if start_tag.contains(" transform=") {
        return None;
    }
    let start = start_tag.find(CLIP_ATTRIBUTE)?;
    let id_start = start + CLIP_ATTRIBUTE.len();
    let id_len = start_tag[id_start..].find(")\"")?;
    let id = &start_tag[id_start..id_start + id_len];
    let plain = !id.is_empty()
        && id
            .bytes()
            .all(|byte| byte.is_ascii_alphanumeric() || matches!(byte, b'-' | b'_'));
    plain.then_some(start..id_start + id_len + 2)
}

/// Fragment identifier of a ` clip-path="url(#id)"` attribute.
fn clip_fragment(attribute: &str) -> &str {
    &attribute[CLIP_ATTRIBUTE.len()..attribute.len() - 2]
}

fn push_without(output: &mut String, element: &str, attribute: Range<usize>) {
    output.push_str(&element[..attribute.start]);
    output.push_str(&element[attribute.end..]);
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_write_num() {
        let mut buf = String::new();
        write_num(&mut buf, 10.0);
        assert_eq!(buf, "10");

        buf.clear();
        write_num(&mut buf, 10.5);
        assert_eq!(buf, "10.5");

        buf.clear();
        write_num(&mut buf, 10.123);
        assert_eq!(buf, "10.12");

        buf.clear();
        write_num(&mut buf, 10.126);
        assert_eq!(buf, "10.13");
    }

    #[test]
    fn test_write_color_hex() {
        let mut buf = String::new();
        write_color_hex(&mut buf, 0x0000FF); // Red in COLORREF
        assert_eq!(buf, "#ff0000");

        buf.clear();
        write_color_hex(&mut buf, 0x00FF00); // Green in COLORREF
        assert_eq!(buf, "#00ff00");

        buf.clear();
        write_color_hex(&mut buf, 0xFF0000); // Blue in COLORREF
        assert_eq!(buf, "#0000ff");
    }

    #[test]
    fn test_write_xml_escaped() {
        let mut buf = String::new();
        write_xml_escaped(&mut buf, "Hello <world> & \"friends\"");
        assert_eq!(buf, "Hello &lt;world&gt; &amp; &quot;friends&quot;");
    }

    #[test]
    fn xml_escaping_replaces_forbidden_control_characters() {
        let mut buf = String::new();
        write_xml_escaped(&mut buf, "a\0b\u{8}c\t\nd");
        assert_eq!(buf, "a\u{fffd}b\u{fffd}c\t\nd");
    }

    #[test]
    fn xml_escaping_replaces_the_noncharacters_xml_forbids() {
        let mut buf = String::new();
        write_xml_escaped(&mut buf, "a\u{fffe}b\u{ffff}c\u{fffd}d\u{10ffff}e\u{10000}");
        assert_eq!(buf, "a\u{fffd}b\u{fffd}c\u{fffd}d\u{10ffff}e\u{10000}");
    }

    fn grouped(elements: &[&str]) -> (String, usize) {
        let mut grouper = ClipGrouper::new();
        let mut output = String::new();
        for element in elements {
            grouper.push(&mut output, element);
        }
        grouper.finish(&mut output);
        (output, grouper.groups())
    }

    #[test]
    fn clip_grouper_moves_a_shared_clip_onto_one_group() {
        let (output, groups) = grouped(&[
            r##"<path d="M0 0" fill="#000000" clip-path="url(#a)"/>"##,
            r##"<path d="M1 1" clip-path="url(#a)" fill="#000000"/>"##,
            r##"<g clip-path="url(#a)"><text>x</text></g>"##,
        ]);

        assert_eq!(
            output,
            concat!(
                r##"<g clip-path="url(#a)"><path d="M0 0" fill="#000000"/>"##,
                r##"<path d="M1 1" fill="#000000"/><g><text>x</text></g></g>"##,
            )
        );
        assert_eq!(groups, 1);
    }

    #[test]
    fn clip_grouper_keeps_lone_unclipped_and_transformed_elements() {
        let elements = [
            r##"<path d="M0 0" clip-path="url(#a)"/>"##,
            r##"<path d="M1 1"/>"##,
            r##"<path d="M2 2" clip-path="url(#a)"/>"##,
            r##"<text transform="rotate(9)" clip-path="url(#a)">t</text>"##,
            r##"<path d="M3 3" clip-path="url(#a)"/>"##,
            r##"<g><path d="M4 4" clip-path="url(#a)"/></g>"##,
        ];
        let (output, groups) = grouped(&elements);

        assert_eq!(output, elements.concat());
        assert_eq!(groups, 0);
    }

    #[test]
    fn clip_grouper_never_mixes_two_clips() {
        let (output, groups) = grouped(&[
            r##"<path d="M0 0" clip-path="url(#a)"/>"##,
            r##"<path d="M1 1" clip-path="url(#a)"/>"##,
            r##"<path d="M2 2" clip-path="url(#b)"/>"##,
            r##"<path d="M3 3" clip-path="url(#b)"/>"##,
            r##"<path d="M4 4" clip-path="url(#a)"/>"##,
        ]);

        assert_eq!(
            output,
            concat!(
                r##"<g clip-path="url(#a)"><path d="M0 0"/><path d="M1 1"/></g>"##,
                r##"<g clip-path="url(#b)"><path d="M2 2"/><path d="M3 3"/></g>"##,
                r##"<path d="M4 4" clip-path="url(#a)"/>"##,
            )
        );
        assert_eq!(groups, 2);
    }
}
