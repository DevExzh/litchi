//! Cell XF formatting for the opened-workbook edit.
//!
//! A cell's appearance is the workbook-global extended format (XF) its
//! record names (`[MS-XLS]` 2.4.353).  A format change reads the cell's
//! effective 20-byte XF payload, rewrites the alignment, border-style and
//! fill fields it names, and reuses an identical effective XF or authors a
//! new one.  The XF's six-byte font, number-format and protection prefix is
//! kept, so the existing resource-dependency rule admits the new XF; font,
//! number-format and protection changes are therefore outside this edit.
//! A changed XF drops its `fHasXFExt` flag: the authored resource carries no
//! `XFExt` extension, so its palette fields are authoritative.

use super::{Error, Result};
use crate::{BorderStyle, Color, FillPattern, HorizontalAlignment, Palette, VerticalAlignment};
use litchi_biff::Records;
use litchi_core::binary;

/// Requested XF field changes for one cell; `None` keeps the current value.
///
/// ```
/// use litchi_xls::cell_values::FormatChange;
/// use litchi_xls::{BorderStyle, HorizontalAlignment};
/// let mut change = FormatChange::default();
/// change.horizontal = Some(HorizontalAlignment::Center);
/// change.border_bottom = Some(BorderStyle::Thin);
/// assert!(!change.is_empty());
/// ```
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
#[non_exhaustive]
pub struct FormatChange {
    /// Horizontal alignment (`alc`).
    pub horizontal: Option<HorizontalAlignment>,
    /// Vertical alignment (`alcV`).
    pub vertical: Option<VerticalAlignment>,
    /// Text wrapping (`fWrap`).
    pub wrap_text: Option<bool>,
    /// Indent level 0 through 15 (`cIndent`).
    pub indent: Option<u8>,
    /// Shrink-to-fit (`fShrinkToFit`).
    pub shrink_to_fit: Option<bool>,
    /// Left border line style (`dgLeft`).
    pub border_left: Option<BorderStyle>,
    /// Right border line style (`dgRight`).
    pub border_right: Option<BorderStyle>,
    /// Top border line style (`dgTop`).
    pub border_top: Option<BorderStyle>,
    /// Bottom border line style (`dgBottom`).
    pub border_bottom: Option<BorderStyle>,
    /// Fill pattern (`fls`).
    pub fill_pattern: Option<FillPattern>,
    /// Pattern foreground color; must equal a workbook palette entry.
    pub fill_foreground: Option<Color>,
    /// Pattern background color; must equal a workbook palette entry.
    pub fill_background: Option<Color>,
}

impl FormatChange {
    /// Whether this change names no field.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        *self == Self::default()
    }
}

const PALETTE: u16 = 0x0092;
const EOF: u16 = 0x000a;
const FIRST_PALETTE_INDEX: u16 = 0x08;
const LAST_PALETTE_INDEX: u16 = 0x3f;
/// System window-text ("automatic") color index.
const AUTOMATIC: u32 = 0x40;
const HAS_XF_EXT: u32 = 1 << 25;
const ATTRIBUTE_ALIGNMENT: u8 = 1 << 4;
const ATTRIBUTE_BORDER: u8 = 1 << 5;
const ATTRIBUTE_PATTERN: u8 = 1 << 6;

/// The workbook palette from the globals' `Palette` record, or the BIFF8
/// default palette.
pub(super) fn workbook_palette(workbook: &[u8]) -> Result<Palette> {
    for item in Records::new(workbook, litchi_biff::Limits::default())? {
        let record = item?;
        match record.kind().get() {
            PALETTE => {
                let mut report = crate::ToleranceReport::default();
                let offset = u32::try_from(record.offset()).unwrap_or(u32::MAX);
                return Palette::parse_tolerant(record.payload(), offset, &mut report);
            },
            EOF => break,
            _ => {},
        }
    }
    Ok(Palette::default())
}

fn palette_index(palette: &Palette, color: Color) -> Result<u32> {
    (FIRST_PALETTE_INDEX..=LAST_PALETTE_INDEX)
        .find(|index| palette.color(*index) == Some(color))
        .map(u32::from)
        .ok_or_else(|| {
            Error::UnsupportedFeature(
                "the fill color is not an entry of the workbook's 56-color palette".into(),
            )
        })
}

fn set_bits(value: u32, shift: u32, width: u32, field: u32) -> u32 {
    let mask = ((1_u32 << width) - 1) << shift;
    (value & !mask) | ((field << shift) & mask)
}

/// Apply `change` to a 20-byte XF payload.
pub(super) fn apply(payload: &[u8], change: &FormatChange, palette: &Palette) -> Result<Vec<u8>> {
    if payload.len() != 20 {
        return Err(Error::InvalidData(
            "BIFF8 XF payload must be exactly 20 bytes".into(),
        ));
    }
    if change.indent.is_some_and(|indent| indent > 15) {
        return Err(Error::InvalidData(
            "cell indent must be 0 through 15".into(),
        ));
    }
    let mut output = payload.to_vec();
    let mut attributes = output[9];
    let mut alignment = output[6];
    if let Some(horizontal) = change.horizontal {
        alignment = (alignment & !0x07) | horizontal.code();
    }
    if let Some(wrap) = change.wrap_text {
        alignment = (alignment & !0x08) | (u8::from(wrap) << 3);
    }
    if let Some(vertical) = change.vertical {
        alignment = (alignment & !0x70) | (vertical.code() << 4);
    }
    let mut indent = output[8];
    if let Some(value) = change.indent {
        indent = (indent & !0x0f) | value;
    }
    if let Some(shrink) = change.shrink_to_fit {
        indent = (indent & !0x10) | (u8::from(shrink) << 4);
    }
    if alignment != output[6] || indent != output[8] {
        attributes |= ATTRIBUTE_ALIGNMENT;
    }
    output[6] = alignment;
    output[8] = indent;

    let mut borders = binary::read_u32_le_at(&output, 10)?;
    let mut lower = binary::read_u32_le_at(&output, 14)?;
    let edges = [
        (change.border_left, 0, Some(16)),
        (change.border_right, 4, Some(23)),
        (change.border_top, 8, None),
        (change.border_bottom, 12, None),
    ];
    let source_borders = borders;
    let source_lower = lower;
    for (index, (style, shift, color_shift)) in edges.into_iter().enumerate() {
        let Some(style) = style else { continue };
        borders = set_bits(borders, shift, 4, u32::from(style.code()));
        // A newly drawn edge with the black built-in index uses the
        // automatic color, as Excel writes it.
        if style != BorderStyle::None {
            match color_shift {
                Some(color_shift) if (borders >> color_shift) & 0x7f == 0 => {
                    borders = set_bits(borders, color_shift, 7, AUTOMATIC);
                },
                None => {
                    let color_shift = if index == 2 { 0 } else { 7 };
                    if (lower >> color_shift) & 0x7f == 0 {
                        lower = set_bits(lower, color_shift, 7, AUTOMATIC);
                    }
                },
                Some(_) => {},
            }
        }
    }
    if borders != source_borders || lower != source_lower {
        attributes |= ATTRIBUTE_BORDER;
    }
    let mut fill = binary::read_u16_le_at(&output, 18)?;
    let source_fill = fill;
    let source_pattern = (lower >> 26) & 0x3f;
    if let Some(pattern) = change.fill_pattern {
        lower = set_bits(lower, 26, 6, u32::from(pattern.code()));
    }
    if let Some(color) = change.fill_foreground {
        fill = (fill & !0x007f) | u16::try_from(palette_index(palette, color)?).unwrap_or(0);
    }
    if let Some(color) = change.fill_background {
        fill = (fill & !0x3f80) | (u16::try_from(palette_index(palette, color)?).unwrap_or(0) << 7);
    }
    if fill != source_fill || (lower >> 26) & 0x3f != source_pattern {
        attributes |= ATTRIBUTE_PATTERN;
    }
    output[10..14].copy_from_slice(&borders.to_le_bytes());
    output[18..20].copy_from_slice(&fill.to_le_bytes());
    output[9] = attributes;
    if output != payload {
        // The authored resource has no XFExt extension.
        lower &= !HAS_XF_EXT;
    }
    output[14..18].copy_from_slice(&lower.to_le_bytes());
    Ok(output)
}
