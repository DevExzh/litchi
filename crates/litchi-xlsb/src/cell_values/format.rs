//! Partial cell-format changes for package-owned cell-value edits.
//!
//! A [`Format`] names only the attributes it changes. Applying it to a cell
//! starts from the cell's current `BrtXF` and its font, fill and border
//! records, patches the named fields in copies of those records, and interns
//! the results in the styles part (an equal existing record is reused). Every
//! attribute a format does not name keeps its exact bytes, including the
//! parent cell style, protection, the font family, charset and scheme, and
//! gradient or theme-colored resources the change does not touch.
//!
//! [`super::structure::AuthoredStyle`] remains the complete-style input for
//! newly authored cells; a `Format` is the partial change of an existing or
//! virtual (blank) cell, and shares the same record encoders.

#![deny(
    clippy::cast_lossless,
    clippy::cast_possible_truncation,
    clippy::cast_possible_wrap,
    clippy::cast_precision_loss,
    clippy::cast_sign_loss,
    clippy::checked_conversions,
    clippy::expect_used,
    clippy::float_cmp,
    clippy::let_underscore_must_use,
    clippy::unnecessary_unwrap,
    reason = "format edits use checked BIFF12 indexes and field widths"
)]

use crate::package::error::{Error, Result};
use crate::raw::{Encoder, kind};
use crate::styles::{BorderStyle, HorizontalAlignment, VerticalAlignment};

/// One border side as a format sets it: a line style and an optional direct
/// ARGB color (`None` is the automatic color).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct BorderLine {
    /// The line style; [`BorderStyle::None`] removes the line.
    pub style: BorderStyle,
    /// The direct ARGB color, or `None` for the automatic color.
    pub color: Option<u32>,
}

/// A partial change of a cell's formatting.
///
/// `None` leaves an attribute unchanged. Colors are direct ARGB values
/// (`0xAARRGGBB`); `Some(None)` on a color, fill or border side resets it to
/// the automatic color, no fill or no line.
///
/// ```
/// use litchi_xlsb::cell_values::Format;
///
/// let mut format = Format::default();
/// format.bold = Some(true);
/// format.fill = Some(Some(0xFF1F_4E79));
/// format.number_format = Some("0.00%".to_owned());
/// assert!(!format.is_empty());
/// ```
#[derive(Debug, Clone, Default, PartialEq)]
#[non_exhaustive]
pub struct Format {
    /// Bold weight.
    pub bold: Option<bool>,
    /// Italic.
    pub italic: Option<bool>,
    /// Strikethrough.
    pub strike: Option<bool>,
    /// Single underline (`true`) or none (`false`).
    pub underline: Option<bool>,
    /// Font size in points, stored in twentieths of a point (1–409 pt).
    pub font_size: Option<f64>,
    /// Font name (1–31 UTF-16 units).
    pub font_name: Option<String>,
    /// Font color: direct ARGB, or `Some(None)` for automatic.
    pub font_color: Option<Option<u32>>,
    /// Solid fill color, or `Some(None)` for no fill.
    pub fill: Option<Option<u32>>,
    /// Top border line, or `Some(None)` for no line.
    pub border_top: Option<Option<BorderLine>>,
    /// Bottom border line, or `Some(None)` for no line.
    pub border_bottom: Option<Option<BorderLine>>,
    /// Left border line, or `Some(None)` for no line.
    pub border_left: Option<Option<BorderLine>>,
    /// Right border line, or `Some(None)` for no line.
    pub border_right: Option<Option<BorderLine>>,
    /// Number-format code (1–255 UTF-16 units). Built-in codes such as
    /// `General` or `0.00` select their built-in identifier.
    pub number_format: Option<String>,
    /// Horizontal alignment.
    pub horizontal: Option<HorizontalAlignment>,
    /// Vertical alignment.
    pub vertical: Option<VerticalAlignment>,
    /// Text wrapping.
    pub wrap_text: Option<bool>,
    /// Text rotation: 0–180 degrees, or 255 for vertical text.
    pub rotation: Option<u8>,
    /// Indent level (0–250).
    pub indent: Option<u8>,
    /// Shrink to fit.
    pub shrink_to_fit: Option<bool>,
    /// Reading order: 0 context, 1 left-to-right, 2 right-to-left.
    pub reading_order: Option<u8>,
}

impl Format {
    /// Whether the format changes nothing.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        *self == Self::default()
    }

    fn changes_font(&self) -> bool {
        self.bold.is_some()
            || self.italic.is_some()
            || self.strike.is_some()
            || self.underline.is_some()
            || self.font_size.is_some()
            || self.font_name.is_some()
            || self.font_color.is_some()
    }

    fn changes_border(&self) -> bool {
        self.border_top.is_some()
            || self.border_bottom.is_some()
            || self.border_left.is_some()
            || self.border_right.is_some()
    }

    fn changes_alignment(&self) -> bool {
        self.horizontal.is_some()
            || self.vertical.is_some()
            || self.wrap_text.is_some()
            || self.rotation.is_some()
            || self.indent.is_some()
            || self.shrink_to_fit.is_some()
            || self.reading_order.is_some()
    }

    /// Check every value against its BIFF12 field before anything is staged.
    ///
    /// # Errors
    ///
    /// Returns [`Error::InvalidFormat`] naming the first value outside its
    /// field.
    pub fn validate(&self) -> Result<()> {
        if let Some(size) = self.font_size {
            font_height(size)?;
        }
        if let Some(name) = &self.font_name {
            let units = name.encode_utf16().count();
            if !(1..=31).contains(&units) || name.contains('\0') {
                return Err(Error::InvalidFormat(
                    "a font name must hold 1..=31 UTF-16 units and no NUL",
                ));
            }
        }
        if let Some(code) = &self.number_format {
            let units = code.encode_utf16().count();
            if !(1..=255).contains(&units) || code.contains('\0') {
                return Err(Error::InvalidFormat(
                    "a number format must hold 1..=255 UTF-16 units and no NUL",
                ));
            }
        }
        if self
            .rotation
            .is_some_and(|value| value > 180 && value != 255)
        {
            return Err(Error::InvalidFormat("text rotation must be 0..=180 or 255"));
        }
        if self.indent.is_some_and(|value| value > 250) {
            return Err(Error::InvalidFormat("indent must be 0250"));
        }
        if self.reading_order.is_some_and(|value| value > 2) {
            return Err(Error::InvalidFormat("reading order must be 0, 1 or 2"));
        }
        Ok(())
    }
}

/// `BrtXF` field offsets (MS-XLSB 2.4.812).
const XF_LEN: usize = 16;
const XF_NUMBER_FORMAT: usize = 2;
const XF_FONT: usize = 4;
const XF_FILL: usize = 6;
const XF_BORDER: usize = 8;
const XF_ROTATION: usize = 10;
const XF_INDENT: usize = 11;
const XF_ALIGNMENT: usize = 12;
const XF_PROPERTIES: usize = 13;
const XF_APPLIED: usize = 14;

/// `xfGrbitAtr` bits: the attribute groups a cell format sets itself.
const APPLY_NUMBER_FORMAT: u8 = 0x01;
const APPLY_FONT: u8 = 0x02;
const APPLY_ALIGNMENT: u8 = 0x04;
const APPLY_BORDER: u8 = 0x08;
const APPLY_FILL: u8 = 0x10;

/// `BrtFont` field offsets (MS-XLSB 2.4.661).
const FONT_NAME: usize = 21;
const FONT_COLOR: usize = 12;

/// `BrtBorder` side offsets (MS-XLSB 2.4.43): top, bottom, left, right.
const BORDER_LEN: usize = 51;

fn font_height(points: f64) -> Result<u16> {
    let twips = (points * 20.0).round();
    if !points.is_finite() || !(20.0..=8180.0).contains(&twips) {
        return Err(Error::InvalidFormat("a font size must be 1..=409 points"));
    }
    #[allow(
        clippy::cast_possible_truncation,
        clippy::cast_sign_loss,
        reason = "`twips` is an integral value in 20..=8180"
    )]
    let height = twips as u16;
    Ok(height)
}

fn color_bytes(color: Option<u32>) -> [u8; 8] {
    match color {
        Some(argb) => {
            let [alpha, red, green, blue] = argb.to_be_bytes();
            // fValidRGB = 1, xColorType = 2 (direct RGB).
            [0x05, 0, 0, 0, red, green, blue, alpha]
        },
        None => [0; 8],
    }
}

fn field<'a>(
    payload: &'a mut [u8],
    range: std::ops::Range<usize>,
    _what: &str,
) -> Result<&'a mut [u8]> {
    payload
        .get_mut(range)
        .ok_or(Error::InvalidFormat("record is truncated"))
}

fn read_index(payload: &[u8], offset: usize) -> Result<usize> {
    let bytes = payload
        .get(offset..offset + 2)
        .ok_or(Error::InvalidFormat("BrtXF record is truncated"))?;
    Ok(usize::from(u16::from_le_bytes([bytes[0], bytes[1]])))
}

fn patch_font(font: &mut Vec<u8>, format: &Format, limits: crate::raw::Limits) -> Result<()> {
    if font.len() < FONT_NAME {
        return Err(Error::InvalidFormat("BrtFont record is truncated"));
    }
    if let Some(size) = format.font_size {
        field(font, 0..2, "BrtFont")?.copy_from_slice(&font_height(size)?.to_le_bytes());
    }
    let mut flags = u16::from_le_bytes([font[2], font[3]]);
    if let Some(italic) = format.italic {
        flags = if italic {
            flags | 0x0002
        } else {
            flags & !0x0002
        };
    }
    if let Some(strike) = format.strike {
        flags = if strike {
            flags | 0x0008
        } else {
            flags & !0x0008
        };
    }
    field(font, 2..4, "BrtFont")?.copy_from_slice(&flags.to_le_bytes());
    if let Some(bold) = format.bold {
        let weight: u16 = if bold { 0x02BC } else { 0x0190 };
        field(font, 4..6, "BrtFont")?.copy_from_slice(&weight.to_le_bytes());
    }
    if let Some(underline) = format.underline {
        field(font, 8..9, "BrtFont")?[0] = u8::from(underline);
    }
    if let Some(color) = format.font_color {
        field(font, FONT_COLOR..FONT_COLOR + 8, "BrtFont")?.copy_from_slice(&color_bytes(color));
    }
    if let Some(name) = &format.font_name {
        font.truncate(FONT_NAME);
        Encoder::new(&mut *font, limits).write_wide_string(name)?;
    }
    Ok(())
}

fn patch_border(border: &mut [u8], format: &Format) -> Result<()> {
    if border.len() < BORDER_LEN {
        return Err(Error::InvalidFormat("BrtBorder record is truncated"));
    }
    for (offset, change) in [
        (1, format.border_top),
        (11, format.border_bottom),
        (21, format.border_left),
        (31, format.border_right),
    ] {
        let Some(line) = change else {
            continue;
        };
        let side = field(border, offset..offset + 10, "BrtBorder")?;
        match line {
            Some(line) if line.style != BorderStyle::None => {
                side[0] = super::resources::border_style_bits(line.style);
                side[1] = 0;
                side[2..10].copy_from_slice(&color_bytes(line.color));
            },
            Some(_) | None => side.copy_from_slice(&[0; 10]),
        }
    }
    Ok(())
}

fn patch_alignment(xf: &mut [u8], format: &Format) -> Result<()> {
    if let Some(rotation) = format.rotation {
        field(xf, XF_ROTATION..XF_ROTATION + 1, "BrtXF")?[0] = rotation;
    }
    if let Some(indent) = format.indent {
        field(xf, XF_INDENT..XF_INDENT + 1, "BrtXF")?[0] = indent;
    }
    let mut alignment = xf[XF_ALIGNMENT];
    if let Some(horizontal) = format.horizontal {
        alignment = (alignment & !0x07) | super::resources::horizontal_bits(horizontal);
    }
    if let Some(vertical) = format.vertical {
        alignment = (alignment & !0x38) | (super::resources::vertical_bits(vertical) << 3);
    }
    if let Some(wrap) = format.wrap_text {
        alignment = if wrap {
            alignment | 0x40
        } else {
            alignment & !0x40
        };
    }
    xf[XF_ALIGNMENT] = alignment;
    let mut properties = xf[XF_PROPERTIES];
    if let Some(shrink) = format.shrink_to_fit {
        properties = if shrink {
            properties | 0x01
        } else {
            properties & !0x01
        };
    }
    if let Some(order) = format.reading_order {
        properties = (properties & !0x0C) | (order << 2);
    }
    xf[XF_PROPERTIES] = properties;
    Ok(())
}

/// Built-in number-format identifier of `code`, when it is one.
fn builtin_number_format(code: &str) -> Option<u16> {
    crate::styles::Table::default()
        .num_fmts
        .into_iter()
        .filter(|(_, builtin)| builtin == code)
        .filter_map(|(id, _)| u16::try_from(id).ok())
        .min()
}

/// Apply `format` to the cell XF `old` of the styles stream `styles`,
/// interning every changed record. Returns the resulting XF index.
pub(super) fn restyle(
    styles: &mut Vec<u8>,
    old: u32,
    format: &Format,
    limits: crate::raw::Limits,
) -> Result<u32> {
    let xfs = super::resources::collection(
        styles,
        kind::BEGIN_CELL_XFS,
        kind::END_CELL_XFS,
        kind::XF,
        limits,
    )?;
    let old_index = usize::try_from(old)
        .map_err(|_error| Error::InvalidFormat("cell XF index exceeds this platform"))?;
    let mut xf = xfs
        .get(old_index)
        .cloned()
        .ok_or(Error::InvalidFormat("cell XF is not in the styles part"))?;
    if xf.len() < XF_LEN {
        return Err(Error::InvalidFormat("BrtXF record is truncated"));
    }
    let mut applied = 0_u8;
    if format.changes_font() {
        let fonts = super::resources::collection(
            styles,
            kind::BEGIN_FONTS,
            kind::END_FONTS,
            kind::FONT,
            limits,
        )?;
        let mut font = fonts
            .get(read_index(&xf, XF_FONT)?)
            .cloned()
            .ok_or(Error::InvalidFormat("the cell XF names a missing font"))?;
        patch_font(&mut font, format, limits)?;
        let index = super::resources::intern_payload(
            styles,
            kind::BEGIN_FONTS,
            kind::END_FONTS,
            kind::FONT,
            &font,
            limits,
        )?;
        if usize::from(index) != read_index(&xf, XF_FONT)? {
            xf[XF_FONT..XF_FONT + 2].copy_from_slice(&index.to_le_bytes());
            applied |= APPLY_FONT;
        }
    }
    if let Some(fill) = format.fill {
        let payload = crate::writer::StylesWriter::encode_fill_payload(&crate::styles::Fill {
            pattern_type: u32::from(fill.is_some()),
            fg_color: fill,
            bg_color: None,
        })?;
        let index = super::resources::intern_payload(
            styles,
            kind::BEGIN_FILLS,
            kind::END_FILLS,
            kind::FILL,
            &payload,
            limits,
        )?;
        if usize::from(index) != read_index(&xf, XF_FILL)? {
            xf[XF_FILL..XF_FILL + 2].copy_from_slice(&index.to_le_bytes());
            applied |= APPLY_FILL;
        }
    }
    if format.changes_border() {
        let borders = super::resources::collection(
            styles,
            kind::BEGIN_BORDERS,
            kind::END_BORDERS,
            kind::BORDER,
            limits,
        )?;
        let mut border = borders
            .get(read_index(&xf, XF_BORDER)?)
            .cloned()
            .ok_or(Error::InvalidFormat("the cell XF names a missing border"))?;
        patch_border(&mut border, format)?;
        let index = super::resources::intern_payload(
            styles,
            kind::BEGIN_BORDERS,
            kind::END_BORDERS,
            kind::BORDER,
            &border,
            limits,
        )?;
        if usize::from(index) != read_index(&xf, XF_BORDER)? {
            xf[XF_BORDER..XF_BORDER + 2].copy_from_slice(&index.to_le_bytes());
            applied |= APPLY_BORDER;
        }
    }
    if let Some(code) = &format.number_format {
        let id = match builtin_number_format(code) {
            Some(id) => id,
            None => super::resources::intern_number_format(styles, code, limits)?,
        };
        if usize::from(id) != read_index(&xf, XF_NUMBER_FORMAT)? {
            xf[XF_NUMBER_FORMAT..XF_NUMBER_FORMAT + 2].copy_from_slice(&id.to_le_bytes());
            applied |= APPLY_NUMBER_FORMAT;
        }
    }
    if format.changes_alignment() {
        let before = xf[XF_ROTATION..=XF_PROPERTIES].to_vec();
        patch_alignment(&mut xf, format)?;
        if xf[XF_ROTATION..=XF_PROPERTIES] != before[..] {
            applied |= APPLY_ALIGNMENT;
        }
    }
    xf[XF_APPLIED] |= applied;
    if xf == xfs[old_index] {
        return Ok(old);
    }
    let index = super::resources::intern_payload(
        styles,
        kind::BEGIN_CELL_XFS,
        kind::END_CELL_XFS,
        kind::XF,
        &xf,
        limits,
    )?;
    Ok(u32::from(index))
}
