//! The legacy VML note shapes Excel draws cell notes from, shared by the
//! XLSX and XLSB note editors (unified-ops records 0320, 0322 and 0325).
//!
//! VML parts are not always well-formed XML, so [`shapes`] scans them
//! bytewise; [`note_shape`] and [`new_drawing`] author the hidden note box
//! Excel itself writes.

use std::ops::Range;

use crate::{Error, LimitKind, Result};

/// The largest VML shape member retained by a scanner.
pub(crate) const MAX_VML_BYTES: usize = crate::objects::MAX_DRAWING_BYTES;
/// The largest VML shape inventory retained by a scanner.
pub(crate) const MAX_SHAPES: usize = 100_000;

/// Caller-selected bounds for a raw VML shape scan.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[must_use]
pub struct Limits {
    /// Maximum VML member bytes inspected.
    pub max_xml_bytes: usize,
    /// Maximum shapes retained by one scan.
    pub max_shapes: usize,
}

impl Limits {
    /// The established finite VML scanner policy.
    pub const DEFAULT: Self = Self {
        max_xml_bytes: MAX_VML_BYTES,
        max_shapes: MAX_SHAPES,
    };

    /// Construct explicit VML scanner bounds.
    #[must_use = "use the selected VML limits"]
    pub const fn new(max_xml_bytes: usize, max_shapes: usize) -> Self {
        Self {
            max_xml_bytes,
            max_shapes,
        }
    }

    /// Derive VML bounds from an OPC source owner without widening the hard
    /// scanner ceilings.
    #[must_use = "use the limits derived from the retained source policy"]
    pub fn from_source_limits(source: litchi_opc::SourceLimits) -> Self {
        Self {
            max_xml_bytes: usize::try_from(source.max_part_bytes())
                .unwrap_or(usize::MAX)
                .clamp(1, MAX_VML_BYTES),
            max_shapes: source.max_events().clamp(1, MAX_SHAPES),
        }
    }

    /// Validate the policy against the hard scanner ceilings.
    pub fn validate(self) -> Result<Self> {
        for (kind, value) in [
            (LimitKind::DrawingBytes, self.max_xml_bytes),
            (LimitKind::VmlShapes, self.max_shapes),
        ] {
            let maximum = kind.maximum();
            if value == 0 || value > maximum {
                return Err(Error::invalid_limit(kind, value, maximum));
            }
        }
        Ok(self)
    }
}

impl Default for Limits {
    fn default() -> Self {
        Self::DEFAULT
    }
}

fn invalid(message: &str) -> Error {
    Error::Invalid(message.to_owned())
}

/// The `_x0000_t202` text-box shape type every note shape refers to.
pub const SHAPE_TYPE: &str = "<v:shapetype id=\"_x0000_t202\" coordsize=\"21600,21600\" o:spt=\"202\" path=\"m,l,21600r21600,l21600,xe\"><v:stroke joinstyle=\"miter\"/><v:path gradientshapeok=\"t\" o:connecttype=\"rect\"/></v:shapetype>";

fn first_shape_id(sheet: usize) -> u64 {
    (sheet as u64 + 1) * 1024 + 1
}

/// One VML shape: its byte span and, for a note, its zero-based cell.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Shape {
    /// The shape's bytes in the part.
    pub span: Range<usize>,
    /// The note's zero-based row and column, for a note shape.
    pub note: Option<(u32, u32)>,
    id: Option<u64>,
}

/// Whether `needle` occurs in `haystack`.
#[must_use]
pub fn contains(haystack: &[u8], needle: &[u8]) -> bool {
    haystack
        .windows(needle.len())
        .any(|window| window == needle)
}

fn find(haystack: &[u8], needle: &[u8], from: usize) -> Option<usize> {
    haystack
        .get(from..)?
        .windows(needle.len())
        .position(|window| window == needle)
        .map(|at| at + from)
}

/// The last position of `needle` in `haystack`.
#[must_use]
pub fn rfind(haystack: &[u8], needle: &[u8]) -> Option<usize> {
    haystack
        .windows(needle.len())
        .rposition(|window| window == needle)
}

fn element_text(xml: &[u8], tag: &[u8]) -> Option<u32> {
    let open = [b"<x:", tag, b">"].concat();
    let close = [b"</x:", tag, b">"].concat();
    let start = find(xml, &open, 0)? + open.len();
    let end = find(xml, &close, start)?;
    std::str::from_utf8(&xml[start..end])
        .ok()?
        .trim()
        .parse()
        .ok()
}

/// The `v:shape` elements of a VML drawing, scanned bytewise: VML parts
/// are not always well-formed XML.
///
/// # Errors
///
/// Returns [`Error::Invalid`] for an unterminated shape.
pub fn shapes(xml: &[u8], limits: &Limits) -> Result<Vec<Shape>> {
    let limits = limits.validate()?;
    if xml.len() > limits.max_xml_bytes {
        return Err(Error::limit(
            LimitKind::DrawingBytes,
            xml.len(),
            limits.max_xml_bytes,
        ));
    }
    let mut shapes = Vec::new();
    let mut cursor = 0;
    while let Some(start) = find(xml, b"<v:shape", cursor) {
        let after = xml.get(start + 8).copied();
        if !matches!(after, Some(b' ' | b'>' | b'\t' | b'\r' | b'\n')) {
            cursor = start + 8;
            continue;
        }
        let tag_end = find(xml, b">", start).ok_or_else(|| invalid("unterminated VML shape"))?;
        let end = if xml[tag_end - 1] == b'/' {
            tag_end + 1
        } else {
            find(xml, b"</v:shape>", tag_end).ok_or_else(|| invalid("unterminated VML shape"))?
                + b"</v:shape>".len()
        };
        let body = &xml[start..end];
        let note = (contains(body, b"ObjectType=\"Note\"") || contains(body, b"ObjectType='Note'"))
            .then(|| element_text(body, b"Row").zip(element_text(body, b"Column")))
            .flatten();
        let id = find(body, b"id=\"_x0000_s", 0).and_then(|at| {
            let digits = &body[at + 11..];
            let len = digits
                .iter()
                .take_while(|byte| byte.is_ascii_digit())
                .count();
            std::str::from_utf8(&digits[..len]).ok()?.parse().ok()
        });
        let next = shapes
            .len()
            .checked_add(1)
            .ok_or_else(|| Error::limit(LimitKind::VmlShapes, usize::MAX, limits.max_shapes))?;
        if next > limits.max_shapes {
            return Err(Error::limit(LimitKind::VmlShapes, next, limits.max_shapes));
        }
        shapes.push(Shape {
            span: start..end,
            note,
            id,
        });
        cursor = end;
    }
    Ok(shapes)
}

/// The id a new shape of the VML part of sheet `sheet` (zero-based)
/// takes: past every `_x0000_sN` in use, in the part's id block.
#[must_use]
pub fn next_shape_id(xml: &[u8], shapes: &[Shape], sheet: usize) -> u64 {
    let base = find(xml, b"data=\"", 0)
        .and_then(|at| {
            let digits = &xml[at + 6..];
            let len = digits
                .iter()
                .take_while(|byte| byte.is_ascii_digit())
                .count();
            std::str::from_utf8(&digits[..len])
                .ok()?
                .parse::<u64>()
                .ok()
        })
        .map_or_else(|| first_shape_id(sheet), |block| block * 1024 + 1);
    shapes
        .iter()
        .filter_map(|shape| shape.id)
        .max()
        .map_or(base, |id| id.max(base - 1) + 1)
}

/// A hidden note box anchored beside the cell, as Excel draws one.
#[must_use]
pub fn note_shape(id: u64, row: u32, column: u32) -> String {
    let anchor_col = column + 1;
    let anchor_row = row.saturating_sub(1);
    format!(
        "<v:shape id=\"_x0000_s{id}\" type=\"#_x0000_t202\" style=\"position:absolute;margin-left:59.25pt;margin-top:1.5pt;width:108pt;height:59.25pt;z-index:1;visibility:hidden\" fillcolor=\"#ffffe1\" o:insetmode=\"auto\"><v:fill color2=\"#ffffe1\"/><v:shadow on=\"t\" color=\"black\" obscured=\"t\"/><v:path o:connecttype=\"none\"/><v:textbox style=\"mso-direction-alt:auto\"><div style=\"text-align:left\"></div></v:textbox><x:ClientData ObjectType=\"Note\"><x:MoveWithCells/><x:SizeWithCells/><x:Anchor>{anchor_col}, 15, {anchor_row}, 2, {}, 15, {}, 16</x:Anchor><x:AutoFill>False</x:AutoFill><x:Row>{row}</x:Row><x:Column>{column}</x:Column></x:ClientData></v:shape>",
        anchor_col + 2,
        anchor_row + 4
    )
}

/// A new VML drawing holding one note shape.
#[must_use]
pub fn new_drawing(sheet: usize, row: u32, column: u32) -> String {
    format!(
        "<xml xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:x=\"urn:schemas-microsoft-com:office:excel\"><o:shapelayout v:ext=\"edit\"><o:idmap v:ext=\"edit\" data=\"{}\"/></o:shapelayout>{SHAPE_TYPE}{}</xml>",
        sheet + 1,
        note_shape(first_shape_id(sheet), row, column)
    )
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn shape_scan_admits_the_selected_shape_count_before_retaining() {
        let xml = br#"<xml><v:shape id="_x0000_s1"/><v:shape id="_x0000_s2"/></xml>"#;
        let limits = Limits::new(xml.len(), 1);
        assert!(matches!(
            shapes(xml, &limits),
            Err(Error::Limit {
                kind: LimitKind::VmlShapes,
                observed: 2,
                maximum: 1,
            })
        ));
    }

    #[test]
    fn shape_scan_rejects_oversized_input_before_scanning() {
        let xml = br#"<xml/>"#;
        let limits = Limits::new(xml.len() - 1, MAX_SHAPES);
        assert!(matches!(
            shapes(xml, &limits),
            Err(Error::Limit {
                kind: LimitKind::DrawingBytes,
                observed,
                maximum: 5,
            }) if observed == xml.len()
        ));
    }
}
