//! Semantic image inputs for the DOC writer.

use super::super::core::WriteError;
use super::codec::{MAX_PICF_DIMENSION_TWIPS, SHAPE_TYPE_PICTURE_FRAME};
use super::validation::{intrinsic_dimensions_twips, sniff_kind, validate_kind};
use crate::parts::spa::{
    ShapeHorizontalOrigin, ShapeTextWrap, ShapeVerticalOrigin, ShapeWrapSide, Spa,
};
use litchi_odraw::image::Kind as BlipKind;

/// An inline picture to be embedded in a DOC document.
///
/// Encoded bytes are stored as-is except that a 14-byte BMP file header is
/// removed to obtain the DIB payload required by `OfficeArt`.
#[derive(Debug, Clone)]
pub struct Picture {
    /// Raw BLIP file data.
    pub(super) data: Vec<u8>,
    /// Detected native `OfficeArt` kind.
    pub(super) kind: BlipKind,
    /// Display width in twips.
    pub(super) width_twips: u32,
    /// Display height in twips.
    pub(super) height_twips: u32,
}

impl Picture {
    /// Create a picture from raw image bytes.
    ///
    /// The format is sniffed from the byte signature and the display
    /// dimensions are derived from bitmap pixels or metafile bounds. Returns
    /// an error for unsupported formats or when the dimensions cannot be
    /// determined; use [`Self::from_parts`] to supply dimensions explicitly.
    pub fn new(data: Vec<u8>) -> Result<Self, WriteError> {
        let kind = sniff_kind(&data)?;
        let dimensions = intrinsic_dimensions_twips(kind, &data).ok_or_else(|| {
            WriteError::InvalidData(
                "DOC picture dimensions are unreadable; use Picture::from_parts".to_string(),
            )
        })?;
        Self::from_parts_as(data, kind, dimensions.0, dimensions.1)
    }

    /// Create a picture from raw image bytes and explicit display dimensions
    /// in twips (1/1440 inch).
    pub fn from_parts(
        data: Vec<u8>,
        width_twips: u32,
        height_twips: u32,
    ) -> Result<Self, WriteError> {
        let kind = sniff_kind(&data)?;
        Self::from_parts_as(data, kind, width_twips, height_twips)
    }

    /// Create a picture with an explicit native `OfficeArt` format and display
    /// dimensions. This is useful for headerless DIB and PICT data whose
    /// format cannot always be inferred unambiguously.
    pub fn from_parts_as(
        mut data: Vec<u8>,
        kind: BlipKind,
        width_twips: u32,
        height_twips: u32,
    ) -> Result<Self, WriteError> {
        validate_kind(kind, &data)?;
        if kind == BlipKind::Dib && data.starts_with(b"BM") {
            data.drain(..14);
        }
        let picture = Self {
            data,
            kind,
            width_twips: 0,
            height_twips: 0,
        };
        picture.with_dimensions_twips(width_twips, height_twips)
    }

    /// Override the display dimensions in twips (1/1440 inch).
    ///
    /// Dimensions must fit the signed 16-bit PICF goal fields, i.e. they must
    /// be in `1..=32767` (about 22.7 inches at 100% scale).
    pub fn with_dimensions_twips(
        mut self,
        width_twips: u32,
        height_twips: u32,
    ) -> Result<Self, WriteError> {
        for dimension in [width_twips, height_twips] {
            if !(1..=MAX_PICF_DIMENSION_TWIPS).contains(&dimension) {
                return Err(WriteError::InvalidData(format!(
                    "DOC picture dimension {dimension} twips is outside 1..={MAX_PICF_DIMENSION_TWIPS}"
                )));
            }
        }
        self.width_twips = width_twips;
        self.height_twips = height_twips;
        Ok(self)
    }

    /// Raw image bytes.
    #[must_use]
    pub fn data(&self) -> &[u8] {
        &self.data
    }

    /// Detected native `OfficeArt` kind.
    #[must_use]
    pub const fn kind(&self) -> BlipKind {
        self.kind
    }

    /// Display width in twips.
    #[must_use]
    pub fn width_twips(&self) -> u32 {
        self.width_twips
    }

    /// Display height in twips.
    #[must_use]
    pub fn height_twips(&self) -> u32 {
        self.height_twips
    }
}

/// Position and wrapping of a floating picture.
///
/// The position is the top-left corner of the picture in twips, relative to
/// the origins selected by [`ShapeHorizontalOrigin`] and
/// [`ShapeVerticalOrigin`]. The size comes from the [`Picture`] display
/// dimensions. Defaults match a typical Word floating picture: page-relative
/// offsets, square wrapping on both sides, in front of the text.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct FloatingPosition {
    /// Left offset in twips relative to the horizontal origin.
    pub(super) left_twips: i32,
    /// Top offset in twips relative to the vertical origin.
    pub(super) top_twips: i32,
    /// Horizontal position origin (Spa `bx`).
    pub(super) horizontal_origin: ShapeHorizontalOrigin,
    /// Vertical position origin (Spa `by`).
    pub(super) vertical_origin: ShapeVerticalOrigin,
    /// Text-wrapping style (Spa `wr`).
    pub(super) wrap: ShapeTextWrap,
    /// Wrap side restriction (Spa `wrk`).
    pub(super) wrap_side: ShapeWrapSide,
    /// Whether the picture appears behind the text (Spa `fBelowText`).
    pub(super) behind_text: bool,
    /// Whether the anchor is locked to its paragraph (Spa `fAnchorLock`).
    pub(super) anchor_locked: bool,
}

impl FloatingPosition {
    /// Create a position from offsets in twips, defaulting to page-relative
    /// origins and square wrapping in front of the text.
    #[must_use]
    pub fn new(left_twips: i32, top_twips: i32) -> Self {
        Self {
            left_twips,
            top_twips,
            horizontal_origin: ShapeHorizontalOrigin::Page,
            vertical_origin: ShapeVerticalOrigin::Page,
            wrap: ShapeTextWrap::Square,
            wrap_side: ShapeWrapSide::Both,
            behind_text: false,
            anchor_locked: false,
        }
    }

    /// Set the horizontal and vertical position origins.
    #[must_use]
    pub fn with_origins(
        mut self,
        horizontal: ShapeHorizontalOrigin,
        vertical: ShapeVerticalOrigin,
    ) -> Self {
        self.horizontal_origin = horizontal;
        self.vertical_origin = vertical;
        self
    }

    /// Set the text-wrapping style.
    #[must_use]
    pub fn with_text_wrap(mut self, wrap: ShapeTextWrap) -> Self {
        self.wrap = wrap;
        self
    }

    /// Set the wrap side restriction.
    #[must_use]
    pub fn with_wrap_side(mut self, wrap_side: ShapeWrapSide) -> Self {
        self.wrap_side = wrap_side;
        self
    }

    /// Place the picture behind (or in front of) the text.
    #[must_use]
    pub fn behind_text(mut self, behind_text: bool) -> Self {
        self.behind_text = behind_text;
        self
    }

    /// Lock the anchor to its paragraph.
    #[must_use]
    pub fn lock_anchor(mut self, anchor_locked: bool) -> Self {
        self.anchor_locked = anchor_locked;
        self
    }
}

/// The visual content of a floating shape in the drawing layer.
pub(crate) enum FloatingShapeContent<'a> {
    /// A picture frame whose BLIP is stored in the blip store.
    Picture(&'a Picture),
    /// A primitive preset-geometry shape (rectangle, ellipse, ...).
    Primitive(&'a super::super::shapes::Shape),
}

/// Everything the table-stream builders need to know about one floating
/// picture or primitive shape anchored in the Main Document.
pub(crate) struct FloatingShapeInfo<'a> {
    /// Character position of the 0x0008 anchor character (Main Document CP).
    pub anchor_cp: u32,
    /// Shape id, shared with the picture's Data-stream block when present.
    pub shape_id: u32,
    /// What the shape renders.
    pub content: FloatingShapeContent<'a>,
    /// Display width in twips.
    pub width_twips: u32,
    /// Display height in twips.
    pub height_twips: u32,
    /// Position and wrapping.
    pub position: &'a FloatingPosition,
    /// Textbox story text when the shape is a text box.
    pub text: Option<&'a str>,
}

impl FloatingShapeInfo<'_> {
    /// The MSOSPT shape type for the `OfficeArtFSP` record instance.
    pub(super) fn shape_type(&self) -> u16 {
        if self.text.is_some() {
            return super::super::shapes::MSOSPT_TEXT_BOX;
        }
        match &self.content {
            FloatingShapeContent::Picture(_) => SHAPE_TYPE_PICTURE_FRAME,
            FloatingShapeContent::Primitive(shape) => shape.kind().shape_type(),
        }
    }

    /// Build the Spa record for this shape.
    pub(super) fn spa(&self) -> Spa {
        let left = self.position.left_twips;
        let top = self.position.top_twips;
        Spa {
            shape_id: self.shape_id,
            left,
            top,
            right: left + self.width_twips as i32,
            bottom: top + self.height_twips as i32,
            horizontal_origin: self.position.horizontal_origin,
            vertical_origin: self.position.vertical_origin,
            wrap: self.position.wrap,
            wrap_side: self.position.wrap_side,
            below_text: self.position.behind_text,
            anchor_locked: self.position.anchor_locked,
        }
    }
}
