//! Normalized, checked bitmap operations shared by EMF and WMF renderers.

#![allow(
    clippy::missing_errors_doc,
    reason = "fallible public APIs uniformly report malformed records, limits, or unsupported rendering"
)]
#![allow(
    clippy::module_name_repetitions,
    reason = "public names remain unambiguous when imported beside EMF and WMF record types"
)]

use std::{
    fmt::Write as _,
    io::{Cursor, Seek, SeekFrom, Write},
};

use base64::{Engine as _, engine::general_purpose::STANDARD};
use image::{DynamicImage, ImageFormat, RgbaImage};
use litchi_core::error::{Error, Result};

use crate::dib::{AlphaInterpretation, ColorUsage, Dib, DibLimits};

const EMF_HEADER_LEN: usize = 8;
const SRCCOPY: u32 = 0x00cc_0020;
const NOTSRCCOPY: u32 = 0x0033_0008;

/// Bitmap-record kind after container-specific parsing.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum BitmapRecordKind {
    EmfBitBlt,
    EmfStretchBlt,
    EmfMaskBlt,
    EmfPlgBlt,
    EmfSetDibitsToDevice,
    EmfStretchDibits,
    EmfAlphaBlend,
    EmfTransparentBlt,
    WmfBitBlt,
    WmfDibBitBlt,
    WmfDibStretchBlt,
    WmfSetDibToDevice,
    WmfStretchBlt,
    WmfStretchDib,
}

/// Signed rectangle preserving metafile mirroring semantics.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SignedRect {
    pub x: i32,
    pub y: i32,
    pub width: i32,
    pub height: i32,
}

/// A signed logical point.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Point {
    pub x: i32,
    pub y: i32,
}

/// Source world-to-page transform used by EMF bitmap records.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct AffineTransform {
    pub m11: f32,
    pub m12: f32,
    pub m21: f32,
    pub m22: f32,
    pub dx: f32,
    pub dy: f32,
}

impl AffineTransform {
    pub const IDENTITY: Self = Self {
        m11: 1.0,
        m12: 0.0,
        m21: 0.0,
        m22: 1.0,
        dx: 0.0,
        dy: 0.0,
    };

    #[must_use]
    pub fn is_identity(self) -> bool {
        self == Self::IDENTITY
    }
}

/// Current stretch policy supplied by the containing playback state.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum StretchPolicy {
    BlackOnWhite,
    WhiteOnBlack,
    ColorOnColor,
    #[default]
    Halftone,
}

/// AlphaBlend parameters.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct AlphaBlend {
    pub constant_alpha: u8,
    pub source_alpha: bool,
}

/// Scan subset used by SetDIBitsToDevice records.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ScanRange {
    pub start: u32,
    pub count: u32,
}

/// Validated WMF device-dependent bitmap, retained for future raster playback.
#[derive(Debug, Clone, Copy)]
pub struct Bitmap16<'a> {
    pub width: u16,
    pub height: u16,
    pub stride: usize,
    pub bit_count: u8,
    pub bits: &'a [u8],
}

/// Normalized source bitmap representation.
#[derive(Debug, Clone, Copy)]
pub enum BitmapSource<'a> {
    Dib(Dib<'a>),
    DeviceDependent(Bitmap16<'a>),
}

/// A fully checked bitmap operation independent of its metafile container.
#[derive(Debug, Clone, Copy)]
pub struct BitmapOp<'a> {
    pub kind: BitmapRecordKind,
    pub source: Option<BitmapSource<'a>>,
    pub mask: Option<Dib<'a>>,
    pub source_rect: SignedRect,
    pub destination: SignedRect,
    pub parallelogram: Option<[Point; 3]>,
    pub source_transform: AffineTransform,
    pub rop3: Rop3,
    pub background_rop3: Option<Rop3>,
    pub scan_range: Option<ScanRange>,
    pub alpha: Option<AlphaBlend>,
    /// COLORREF encoded as `0x00bbggrr`.
    pub transparent_color: Option<u32>,
    pub stretch: StretchPolicy,
    limits: DibLimits,
}

/// Renderable self-contained SVG image fragment.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SvgImage {
    pub element: String,
}

/// The eight-bit ROP3 truth table stored in bits 16..23 of a raster-op code.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Rop3(u8);

impl Rop3 {
    #[must_use]
    pub const fn from_code(code: u32) -> Self {
        Self(code.to_le_bytes()[2])
    }

    #[must_use]
    pub const fn table(self) -> u8 {
        self.0
    }

    /// Evaluates one Boolean bit. The truth-table index is `(P,S,D)`.
    #[must_use]
    pub fn eval_bit(self, pattern: bool, source: bool, destination: bool) -> bool {
        let index = (u8::from(pattern) << 2) | (u8::from(source) << 1) | u8::from(destination);
        self.0 & (1 << index) != 0
    }

    /// Applies the truth table independently to every bit of a byte.
    #[must_use]
    pub fn eval_byte(self, pattern: u8, source: u8, destination: u8) -> u8 {
        let mut output = 0u8;
        for bit in 0..8 {
            let flag = 1u8 << bit;
            if self.eval_bit(
                pattern & flag != 0,
                source & flag != 0,
                destination & flag != 0,
            ) {
                output |= flag;
            }
        }
        output
    }

    /// Composites equal-length byte planes in place for future raster playback.
    pub fn composite_bytes(
        self,
        pattern: &[u8],
        source: &[u8],
        destination: &mut [u8],
    ) -> Result<()> {
        if pattern.len() != source.len() || source.len() != destination.len() {
            return Err(parse("ROP3 compositor planes have different lengths"));
        }
        for ((pattern, source), destination) in
            pattern.iter().zip(source).zip(destination.iter_mut())
        {
            *destination = self.eval_byte(*pattern, *source, *destination);
        }
        Ok(())
    }

    #[must_use]
    pub fn depends_on_destination(self) -> bool {
        (0..=1).any(|pattern| {
            (0..=1).any(|source| {
                self.eval_bit(pattern != 0, source != 0, false)
                    != self.eval_bit(pattern != 0, source != 0, true)
            })
        })
    }

    #[must_use]
    pub fn depends_on_pattern(self) -> bool {
        (0..=1).any(|source| {
            (0..=1).any(|destination| {
                self.eval_bit(false, source != 0, destination != 0)
                    != self.eval_bit(true, source != 0, destination != 0)
            })
        })
    }
}

impl<'a> BitmapOp<'a> {
    /// Parses an EMF payload excluding its eight-byte Type/Size header.
    pub fn parse_emf(
        record_type: u32,
        payload: &'a [u8],
        stretch: StretchPolicy,
        limits: DibLimits,
    ) -> Result<Self> {
        let operation = match record_type {
            76 => parse_emf_bitblt(payload, false, stretch, limits),
            77 => parse_emf_bitblt(payload, true, stretch, limits),
            78 => parse_emf_maskblt(payload, stretch, limits),
            79 => parse_emf_plgblt(payload, stretch, limits),
            80 => parse_emf_setdib(payload, stretch, limits),
            81 => parse_emf_stretchdib(payload, stretch, limits),
            114 => parse_emf_alpha(payload, stretch, limits),
            116 => parse_emf_transparent(payload, stretch, limits),
            _ => {
                return Err(unsupported(format!(
                    "EMF record {record_type} is not a bitmap record"
                )));
            },
        }?;
        validate_operation(operation)
    }

    /// Parses WMF parameters excluding the six-byte Size/Function header.
    pub fn parse_wmf(
        record_size_words: u32,
        function: u16,
        params: &'a [u8],
        stretch: StretchPolicy,
        limits: DibLimits,
    ) -> Result<Self> {
        let expected = words_to_bytes(record_size_words)?
            .checked_sub(6)
            .ok_or_else(|| parse("WMF bitmap record is smaller than its header"))?;
        if expected != params.len() {
            return Err(parse("WMF parameter length does not match RecordSize"));
        }
        let operation = match function & 0x00ff {
            0x22 => parse_wmf_blt(
                record_size_words,
                function,
                params,
                false,
                false,
                stretch,
                limits,
            ),
            0x40 => parse_wmf_blt(
                record_size_words,
                function,
                params,
                false,
                true,
                stretch,
                limits,
            ),
            0x23 => parse_wmf_blt(
                record_size_words,
                function,
                params,
                true,
                false,
                stretch,
                limits,
            ),
            0x41 => parse_wmf_blt(
                record_size_words,
                function,
                params,
                true,
                true,
                stretch,
                limits,
            ),
            0x33 => parse_wmf_setdib(params, stretch, limits),
            0x43 => parse_wmf_stretchdib(params, stretch, limits),
            _ => {
                return Err(unsupported(format!(
                    "WMF function 0x{function:04x} is not a bitmap record"
                )));
            },
        }?;
        validate_operation(operation)
    }

    /// Emits a self-contained SVG `<image>` for source-copy, alpha, or color-key operations.
    pub fn to_svg_image(&self) -> Result<SvgImage> {
        self.to_svg_image_at(None)
    }

    /// Emits the image at an already transformed destination rectangle.
    pub fn to_svg_image_at(&self, destination_override: Option<[f64; 4]>) -> Result<SvgImage> {
        if self.mask.is_some() {
            return Err(unsupported(
                "masked bitmap operations require raster compositing",
            ));
        }
        if !self.source_transform.is_identity() {
            return Err(unsupported(
                "non-identity EMF source transforms require raster playback",
            ));
        }
        if self.rop3.depends_on_destination() {
            return Err(unsupported(format!(
                "ROP3 0x{:02x} depends on destination pixels and cannot be represented by SVG image embedding",
                self.rop3.table()
            )));
        }
        if self.rop3.depends_on_pattern() {
            return Err(unsupported(format!(
                "ROP3 0x{:02x} depends on the selected brush pattern",
                self.rop3.table()
            )));
        }
        if !matches!(self.rop3.table(), 0xcc | 0x33) {
            return Err(unsupported(format!(
                "source-only ROP3 0x{:02x} has no SVG image mapping",
                self.rop3.table()
            )));
        }
        let source = match self.source {
            Some(BitmapSource::Dib(dib)) => dib,
            Some(BitmapSource::DeviceDependent(_)) => {
                return Err(unsupported(
                    "WMF Bitmap16 sources require device palette playback",
                ));
            },
            None => return Err(parse("bitmap operation has no embedded source")),
        };
        let alpha_mode = if self.alpha.is_some_and(|blend| blend.source_alpha) {
            AlphaInterpretation::Premultiplied
        } else {
            AlphaInterpretation::Ignore
        };
        let decoded = source.to_dynamic_image_with_alpha(alpha_mode)?;
        let mut image = crop_signed(&decoded, self.source_rect)?;
        if self.source_rect.width < 0 {
            image = image.fliph();
        }
        if self.source_rect.height < 0 {
            image = image.flipv();
        }
        if self.destination.width < 0 {
            image = image.fliph();
        }
        if self.destination.height < 0 {
            image = image.flipv();
        }
        let mut rgba = image.to_rgba8();
        if self.rop3.table() == Rop3::from_code(NOTSRCCOPY).table() {
            for pixel in rgba.pixels_mut() {
                pixel.0[0] = !pixel.0[0];
                pixel.0[1] = !pixel.0[1];
                pixel.0[2] = !pixel.0[2];
            }
        }
        if let Some(color) = self.transparent_color {
            let bytes = color.to_le_bytes();
            let key = [bytes[0], bytes[1], bytes[2]];
            for pixel in rgba.pixels_mut() {
                if pixel.0[..3] == key {
                    pixel.0[3] = 0;
                }
            }
        }
        let opacity = self
            .alpha
            .map_or(1.0, |blend| f64::from(blend.constant_alpha) / 255.0);
        let png = encode_png(&rgba, self.limits.max_output_bytes)?;
        let encoded_len = png
            .len()
            .checked_add(2)
            .and_then(|length| length.checked_div(3))
            .and_then(|groups| groups.checked_mul(4))
            .and_then(|length| length.checked_add(256))
            .ok_or_else(|| parse("SVG bitmap data URL size overflow"))?;
        if encoded_len > self.limits.max_output_bytes {
            return Err(limit("SVG bitmap data URL", self.limits.max_output_bytes));
        }
        let destination = normalized(self.destination)?;
        let mut element = String::with_capacity(encoded_len);
        if let Some(destination) = destination_override {
            let [x, y, width, height] = destination;
            let _ = write!(
                &mut element,
                "<image x=\"{x}\" y=\"{y}\" width=\"{width}\" height=\"{height}\" opacity=\"{opacity}\" preserveAspectRatio=\"none\" href=\"data:image/png;base64,"
            );
        } else if let Some(points) = self.parallelogram {
            let width = f64::from(image.width());
            let height = f64::from(image.height());
            let a = (f64::from(points[1].x) - f64::from(points[0].x)) / width;
            let b = (f64::from(points[1].y) - f64::from(points[0].y)) / width;
            let c = (f64::from(points[2].x) - f64::from(points[0].x)) / height;
            let d = (f64::from(points[2].y) - f64::from(points[0].y)) / height;
            let _ = write!(
                &mut element,
                "<image width=\"{}\" height=\"{}\" opacity=\"{}\" preserveAspectRatio=\"none\" transform=\"matrix({a} {b} {c} {d} {} {})\" href=\"data:image/png;base64,",
                image.width(),
                image.height(),
                opacity,
                points[0].x,
                points[0].y
            );
        } else {
            let _ = write!(
                &mut element,
                "<image x=\"{}\" y=\"{}\" width=\"{}\" height=\"{}\" opacity=\"{}\" preserveAspectRatio=\"none\" href=\"data:image/png;base64,",
                destination.x, destination.y, destination.width, destination.height, opacity
            );
        }
        STANDARD.encode_string(png, &mut element);
        element.push_str("\"/>");
        Ok(SvgImage { element })
    }
}

fn parse_emf_bitblt<'a>(
    payload: &'a [u8],
    stretched: bool,
    stretch: StretchPolicy,
    limits: DibLimits,
) -> Result<BitmapOp<'a>> {
    let fixed = if stretched { 100 } else { 92 };
    require_len(payload, fixed, "EMF BitBlt fixed fields")?;
    let source = emf_dib(payload, 76, 80, 84, 88, 72, limits)?;
    Ok(BitmapOp {
        kind: if stretched {
            BitmapRecordKind::EmfStretchBlt
        } else {
            BitmapRecordKind::EmfBitBlt
        },
        source,
        mask: None,
        source_rect: SignedRect {
            x: i32_at(payload, 36)?,
            y: i32_at(payload, 40)?,
            width: if stretched {
                i32_at(payload, 92)?
            } else {
                i32_at(payload, 24)?
            },
            height: if stretched {
                i32_at(payload, 96)?
            } else {
                i32_at(payload, 28)?
            },
        },
        destination: rect_at(payload, 16)?,
        parallelogram: None,
        source_transform: xform_at(payload, 44)?,
        rop3: Rop3::from_code(u32_at(payload, 32)?),
        background_rop3: None,
        scan_range: None,
        alpha: None,
        transparent_color: None,
        stretch,
        limits,
    })
}

fn parse_emf_maskblt(
    payload: &[u8],
    stretch: StretchPolicy,
    limits: DibLimits,
) -> Result<BitmapOp<'_>> {
    require_len(payload, 120, "EMF MaskBlt fixed fields")?;
    let rop4 = u32_at(payload, 32)?;
    Ok(BitmapOp {
        kind: BitmapRecordKind::EmfMaskBlt,
        source: emf_dib(payload, 76, 80, 84, 88, 72, limits)?,
        mask: emf_dib(payload, 104, 108, 112, 116, 100, limits)?.and_then(|source| match source {
            BitmapSource::Dib(dib) => Some(dib),
            BitmapSource::DeviceDependent(_) => None,
        }),
        source_rect: SignedRect {
            x: i32_at(payload, 36)?,
            y: i32_at(payload, 40)?,
            width: i32_at(payload, 24)?,
            height: i32_at(payload, 28)?,
        },
        destination: rect_at(payload, 16)?,
        parallelogram: None,
        source_transform: xform_at(payload, 44)?,
        rop3: Rop3(rop4.to_le_bytes()[3]),
        background_rop3: Some(Rop3(rop4.to_le_bytes()[2])),
        scan_range: None,
        alpha: None,
        transparent_color: None,
        stretch,
        limits,
    })
}

fn parse_emf_plgblt(
    payload: &[u8],
    stretch: StretchPolicy,
    limits: DibLimits,
) -> Result<BitmapOp<'_>> {
    require_len(payload, 132, "EMF PlgBlt fixed fields")?;
    let points = [
        point_at(payload, 16)?,
        point_at(payload, 24)?,
        point_at(payload, 32)?,
    ];
    Ok(BitmapOp {
        kind: BitmapRecordKind::EmfPlgBlt,
        source: emf_dib(payload, 88, 92, 96, 100, 84, limits)?,
        mask: emf_dib(payload, 116, 120, 124, 128, 112, limits)?.and_then(|source| match source {
            BitmapSource::Dib(dib) => Some(dib),
            BitmapSource::DeviceDependent(_) => None,
        }),
        source_rect: SignedRect {
            x: i32_at(payload, 40)?,
            y: i32_at(payload, 44)?,
            width: i32_at(payload, 48)?,
            height: i32_at(payload, 52)?,
        },
        destination: bounds_of_points(points)?,
        parallelogram: Some(points),
        source_transform: xform_at(payload, 56)?,
        rop3: Rop3::from_code(SRCCOPY),
        background_rop3: None,
        scan_range: None,
        alpha: None,
        transparent_color: None,
        stretch,
        limits,
    })
}

fn parse_emf_setdib(
    payload: &[u8],
    stretch: StretchPolicy,
    limits: DibLimits,
) -> Result<BitmapOp<'_>> {
    require_len(payload, 68, "EMF SetDIBitsToDevice fixed fields")?;
    Ok(BitmapOp {
        kind: BitmapRecordKind::EmfSetDibitsToDevice,
        source: emf_dib(payload, 40, 44, 48, 52, 56, limits)?,
        mask: None,
        source_rect: SignedRect {
            x: i32_at(payload, 24)?,
            y: i32_at(payload, 28)?,
            width: i32_at(payload, 32)?,
            height: i32_at(payload, 36)?,
        },
        destination: SignedRect {
            x: i32_at(payload, 16)?,
            y: i32_at(payload, 20)?,
            width: i32_at(payload, 32)?,
            height: i32_at(payload, 36)?,
        },
        parallelogram: None,
        source_transform: AffineTransform::IDENTITY,
        rop3: Rop3::from_code(SRCCOPY),
        background_rop3: None,
        scan_range: Some(ScanRange {
            start: u32_at(payload, 60)?,
            count: u32_at(payload, 64)?,
        }),
        alpha: None,
        transparent_color: None,
        stretch,
        limits,
    })
}

fn parse_emf_stretchdib(
    payload: &[u8],
    stretch: StretchPolicy,
    limits: DibLimits,
) -> Result<BitmapOp<'_>> {
    require_len(payload, 72, "EMF StretchDIBits fixed fields")?;
    Ok(BitmapOp {
        kind: BitmapRecordKind::EmfStretchDibits,
        source: emf_dib(payload, 40, 44, 48, 52, 56, limits)?,
        mask: None,
        source_rect: SignedRect {
            x: i32_at(payload, 24)?,
            y: i32_at(payload, 28)?,
            width: i32_at(payload, 32)?,
            height: i32_at(payload, 36)?,
        },
        destination: SignedRect {
            x: i32_at(payload, 16)?,
            y: i32_at(payload, 20)?,
            width: i32_at(payload, 64)?,
            height: i32_at(payload, 68)?,
        },
        parallelogram: None,
        source_transform: AffineTransform::IDENTITY,
        rop3: Rop3::from_code(u32_at(payload, 60)?),
        background_rop3: None,
        scan_range: None,
        alpha: None,
        transparent_color: None,
        stretch,
        limits,
    })
}

fn parse_emf_alpha(
    payload: &[u8],
    stretch: StretchPolicy,
    limits: DibLimits,
) -> Result<BitmapOp<'_>> {
    require_len(payload, 100, "EMF AlphaBlend fixed fields")?;
    if payload[32] != 0 || payload[33] != 0 || payload[35] > 1 {
        return Err(parse("invalid EMF BlendFunction"));
    }
    Ok(BitmapOp {
        kind: BitmapRecordKind::EmfAlphaBlend,
        source: emf_dib(payload, 76, 80, 84, 88, 72, limits)?,
        mask: None,
        source_rect: SignedRect {
            x: i32_at(payload, 36)?,
            y: i32_at(payload, 40)?,
            width: i32_at(payload, 92)?,
            height: i32_at(payload, 96)?,
        },
        destination: rect_at(payload, 16)?,
        parallelogram: None,
        source_transform: xform_at(payload, 44)?,
        rop3: Rop3::from_code(SRCCOPY),
        background_rop3: None,
        scan_range: None,
        alpha: Some(AlphaBlend {
            constant_alpha: payload[34],
            source_alpha: payload[35] == 1,
        }),
        transparent_color: None,
        stretch,
        limits,
    })
}

fn parse_emf_transparent(
    payload: &[u8],
    stretch: StretchPolicy,
    limits: DibLimits,
) -> Result<BitmapOp<'_>> {
    require_len(payload, 100, "EMF TransparentBlt fixed fields")?;
    Ok(BitmapOp {
        kind: BitmapRecordKind::EmfTransparentBlt,
        source: emf_dib(payload, 76, 80, 84, 88, 72, limits)?,
        mask: None,
        source_rect: SignedRect {
            x: i32_at(payload, 36)?,
            y: i32_at(payload, 40)?,
            width: i32_at(payload, 92)?,
            height: i32_at(payload, 96)?,
        },
        destination: rect_at(payload, 16)?,
        parallelogram: None,
        source_transform: xform_at(payload, 44)?,
        rop3: Rop3::from_code(SRCCOPY),
        background_rop3: None,
        scan_range: None,
        alpha: None,
        transparent_color: Some(u32_at(payload, 32)?),
        stretch,
        limits,
    })
}

fn parse_wmf_blt<'a>(
    size: u32,
    function: u16,
    params: &'a [u8],
    stretched: bool,
    dib_source: bool,
    stretch: StretchPolicy,
    limits: DibLimits,
) -> Result<BitmapOp<'a>> {
    let common = if stretched { 20 } else { 16 };
    require_len(params, common, "WMF blit fixed fields")?;
    let no_source = size == u32::from(function >> 8) + 3;
    let source = if no_source {
        None
    } else if dib_source {
        Some(BitmapSource::Dib(Dib::parse(&params[common..], limits)?))
    } else {
        Some(BitmapSource::DeviceDependent(parse_bitmap16(
            &params[common..],
            limits,
        )?))
    };
    let (sh, sw, sy, sx, dh, dw, dy, dx) = if stretched {
        (
            i16_at(params, 4)?,
            i16_at(params, 6)?,
            i16_at(params, 8)?,
            i16_at(params, 10)?,
            i16_at(params, 12)?,
            i16_at(params, 14)?,
            i16_at(params, 16)?,
            i16_at(params, 18)?,
        )
    } else {
        let h = i16_at(params, 8)?;
        let w = i16_at(params, 10)?;
        (
            h,
            w,
            i16_at(params, 4)?,
            i16_at(params, 6)?,
            h,
            w,
            i16_at(params, 12)?,
            i16_at(params, 14)?,
        )
    };
    let kind = match (stretched, dib_source) {
        (false, false) => BitmapRecordKind::WmfBitBlt,
        (false, true) => BitmapRecordKind::WmfDibBitBlt,
        (true, false) => BitmapRecordKind::WmfStretchBlt,
        (true, true) => BitmapRecordKind::WmfDibStretchBlt,
    };
    Ok(BitmapOp {
        kind,
        source,
        mask: None,
        source_rect: SignedRect {
            x: i32::from(sx),
            y: i32::from(sy),
            width: i32::from(sw),
            height: i32::from(sh),
        },
        destination: SignedRect {
            x: i32::from(dx),
            y: i32::from(dy),
            width: i32::from(dw),
            height: i32::from(dh),
        },
        parallelogram: None,
        source_transform: AffineTransform::IDENTITY,
        rop3: Rop3::from_code(u32_at(params, 0)?),
        background_rop3: None,
        scan_range: None,
        alpha: None,
        transparent_color: None,
        stretch,
        limits,
    })
}

fn parse_wmf_setdib(
    params: &[u8],
    stretch: StretchPolicy,
    limits: DibLimits,
) -> Result<BitmapOp<'_>> {
    require_len(params, 18, "WMF SetDIBToDev fields")?;
    let usage = ColorUsage::from_raw(u32::from(u16_at(params, 0)?))?;
    let dib = Dib::parse_with_usage(&params[18..], usage, limits)?;
    Ok(BitmapOp {
        kind: BitmapRecordKind::WmfSetDibToDevice,
        source: Some(BitmapSource::Dib(dib)),
        mask: None,
        source_rect: SignedRect {
            x: i32::from(u16_at(params, 8)?),
            y: i32::from(u16_at(params, 6)?),
            width: i32::from(u16_at(params, 12)?),
            height: i32::from(u16_at(params, 10)?),
        },
        destination: SignedRect {
            x: i32::from(u16_at(params, 16)?),
            y: i32::from(u16_at(params, 14)?),
            width: i32::from(u16_at(params, 12)?),
            height: i32::from(u16_at(params, 10)?),
        },
        parallelogram: None,
        source_transform: AffineTransform::IDENTITY,
        rop3: Rop3::from_code(SRCCOPY),
        background_rop3: None,
        scan_range: Some(ScanRange {
            start: u32::from(u16_at(params, 4)?),
            count: u32::from(u16_at(params, 2)?),
        }),
        alpha: None,
        transparent_color: None,
        stretch,
        limits,
    })
}

fn parse_wmf_stretchdib(
    params: &[u8],
    stretch: StretchPolicy,
    limits: DibLimits,
) -> Result<BitmapOp<'_>> {
    require_len(params, 22, "WMF StretchDIB fields")?;
    let usage = ColorUsage::from_raw(u32::from(u16_at(params, 4)?))?;
    let dib = Dib::parse_with_usage(&params[22..], usage, limits)?;
    Ok(BitmapOp {
        kind: BitmapRecordKind::WmfStretchDib,
        source: Some(BitmapSource::Dib(dib)),
        mask: None,
        source_rect: SignedRect {
            x: i32::from(i16_at(params, 12)?),
            y: i32::from(i16_at(params, 10)?),
            width: i32::from(i16_at(params, 8)?),
            height: i32::from(i16_at(params, 6)?),
        },
        destination: SignedRect {
            x: i32::from(i16_at(params, 20)?),
            y: i32::from(i16_at(params, 18)?),
            width: i32::from(i16_at(params, 16)?),
            height: i32::from(i16_at(params, 14)?),
        },
        parallelogram: None,
        source_transform: AffineTransform::IDENTITY,
        rop3: Rop3::from_code(u32_at(params, 0)?),
        background_rop3: None,
        scan_range: None,
        alpha: None,
        transparent_color: None,
        stretch,
        limits,
    })
}

fn validate_operation(operation: BitmapOp<'_>) -> Result<BitmapOp<'_>> {
    normalized(operation.destination)?;
    if operation.source.is_some() {
        normalized(operation.source_rect)?;
    }
    if let Some(range) = operation.scan_range {
        if range.count == 0 {
            return Err(parse("bitmap scan count must be nonzero"));
        }
        let end = range
            .start
            .checked_add(range.count)
            .ok_or_else(|| parse("bitmap scan range overflows"))?;
        if let Some(BitmapSource::Dib(dib)) = operation.source
            && end > dib.info().height
        {
            return Err(parse("bitmap scan range exceeds the DIB height"));
        }
    }
    Ok(operation)
}

fn emf_dib<'a>(
    payload: &'a [u8],
    off_bmi: usize,
    cb_bmi: usize,
    off_bits: usize,
    cb_bits: usize,
    usage_offset: usize,
    limits: DibLimits,
) -> Result<Option<BitmapSource<'a>>> {
    let bmi_len = usize_at_u32(payload, cb_bmi)?;
    let bits_len = usize_at_u32(payload, cb_bits)?;
    if bmi_len == 0 && bits_len == 0 {
        return Ok(None);
    }
    if bmi_len == 0 || bits_len == 0 {
        return Err(parse("EMF bitmap header and bits must both be present"));
    }
    let bmi = emf_range(payload, u32_at(payload, off_bmi)?, bmi_len)?;
    let bits = emf_range(payload, u32_at(payload, off_bits)?, bits_len)?;
    let usage = ColorUsage::from_raw(u32_at(payload, usage_offset)?)?;
    Ok(Some(BitmapSource::Dib(Dib::parse_parts(
        bmi, bits, usage, limits,
    )?)))
}

fn emf_range(data: &[u8], record_offset: u32, len: usize) -> Result<&[u8]> {
    let start = usize::try_from(record_offset)
        .map_err(|_error| parse("EMF bitmap offset does not fit usize"))?
        .checked_sub(EMF_HEADER_LEN)
        .ok_or_else(|| parse("EMF bitmap offset points into record header"))?;
    let end = start
        .checked_add(len)
        .ok_or_else(|| parse("EMF bitmap range overflows"))?;
    data.get(start..end)
        .ok_or_else(|| parse("EMF bitmap range extends past payload"))
}

fn parse_bitmap16(data: &[u8], limits: DibLimits) -> Result<Bitmap16<'_>> {
    require_len(data, 10, "WMF Bitmap16 header")?;
    let width = i16_at(data, 2)?;
    let height = i16_at(data, 4)?;
    let declared = i16_at(data, 6)?;
    if width <= 0 || height <= 0 || declared <= 0 || data[8] != 1 || data[9] == 0 {
        return Err(parse("invalid WMF Bitmap16 geometry"));
    }
    let stride = usize::try_from(declared).map_err(|_error| parse("Bitmap16 stride is invalid"))?;
    let len = stride
        .checked_mul(usize::try_from(height).map_err(|_error| parse("Bitmap16 height is invalid"))?)
        .ok_or_else(|| parse("Bitmap16 size overflows"))?;
    if len > limits.max_input_bytes {
        return Err(limit("WMF Bitmap16", limits.max_input_bytes));
    }
    let end = 10usize
        .checked_add(len)
        .ok_or_else(|| parse("WMF Bitmap16 extent overflows"))?;
    let bits = data
        .get(10..end)
        .ok_or_else(|| parse("WMF Bitmap16 bits are truncated"))?;
    Ok(Bitmap16 {
        width: u16::try_from(width).map_err(|_error| parse("Bitmap16 width is invalid"))?,
        height: u16::try_from(height).map_err(|_error| parse("Bitmap16 height is invalid"))?,
        stride,
        bit_count: data[9],
        bits,
    })
}

fn crop_signed(image: &DynamicImage, rect: SignedRect) -> Result<DynamicImage> {
    let normalized = normalized(rect)?;
    let x = u32::try_from(normalized.x).map_err(|_error| parse("source x is negative"))?;
    let y = u32::try_from(normalized.y).map_err(|_error| parse("source y is negative"))?;
    let width =
        u32::try_from(normalized.width).map_err(|_error| parse("source width is invalid"))?;
    let height =
        u32::try_from(normalized.height).map_err(|_error| parse("source height is invalid"))?;
    if x.checked_add(width).is_none_or(|end| end > image.width())
        || y.checked_add(height).is_none_or(|end| end > image.height())
    {
        return Err(parse("source rectangle extends outside decoded DIB"));
    }
    Ok(image.crop_imm(x, y, width, height))
}
fn normalized(rect: SignedRect) -> Result<SignedRect> {
    if rect.width == 0 || rect.height == 0 {
        return Err(parse("bitmap rectangle dimensions must be nonzero"));
    }
    let x = if rect.width < 0 {
        rect.x
            .checked_add(rect.width)
            .ok_or_else(|| parse("rectangle x overflows"))?
    } else {
        rect.x
    };
    let y = if rect.height < 0 {
        rect.y
            .checked_add(rect.height)
            .ok_or_else(|| parse("rectangle y overflows"))?
    } else {
        rect.y
    };
    Ok(SignedRect {
        x,
        y,
        width: rect
            .width
            .checked_abs()
            .ok_or_else(|| parse("rectangle width overflows"))?,
        height: rect
            .height
            .checked_abs()
            .ok_or_else(|| parse("rectangle height overflows"))?,
    })
}
fn bounds_of_points(points: [Point; 3]) -> Result<SignedRect> {
    let min_x = points
        .iter()
        .map(|p| p.x)
        .min()
        .ok_or_else(|| parse("missing parallelogram"))?;
    let max_x = points
        .iter()
        .map(|p| p.x)
        .max()
        .ok_or_else(|| parse("missing parallelogram"))?;
    let min_y = points
        .iter()
        .map(|p| p.y)
        .min()
        .ok_or_else(|| parse("missing parallelogram"))?;
    let max_y = points
        .iter()
        .map(|p| p.y)
        .max()
        .ok_or_else(|| parse("missing parallelogram"))?;
    Ok(SignedRect {
        x: min_x,
        y: min_y,
        width: max_x
            .checked_sub(min_x)
            .ok_or_else(|| parse("parallelogram width overflows"))?,
        height: max_y
            .checked_sub(min_y)
            .ok_or_else(|| parse("parallelogram height overflows"))?,
    })
}
fn rect_at(data: &[u8], offset: usize) -> Result<SignedRect> {
    Ok(SignedRect {
        x: i32_at(data, offset)?,
        y: i32_at(data, offset + 4)?,
        width: i32_at(data, offset + 8)?,
        height: i32_at(data, offset + 12)?,
    })
}
fn point_at(data: &[u8], offset: usize) -> Result<Point> {
    Ok(Point {
        x: i32_at(data, offset)?,
        y: i32_at(data, offset + 4)?,
    })
}
fn xform_at(data: &[u8], offset: usize) -> Result<AffineTransform> {
    let transform = AffineTransform {
        m11: f32::from_bits(u32_at(data, offset)?),
        m12: f32::from_bits(u32_at(data, offset + 4)?),
        m21: f32::from_bits(u32_at(data, offset + 8)?),
        m22: f32::from_bits(u32_at(data, offset + 12)?),
        dx: f32::from_bits(u32_at(data, offset + 16)?),
        dy: f32::from_bits(u32_at(data, offset + 20)?),
    };
    if [
        transform.m11,
        transform.m12,
        transform.m21,
        transform.m22,
        transform.dx,
        transform.dy,
    ]
    .into_iter()
    .any(|value| !value.is_finite())
    {
        return Err(parse("EMF bitmap transform contains a non-finite value"));
    }
    Ok(transform)
}
fn require_len(data: &[u8], minimum: usize, field: &str) -> Result<()> {
    if data.len() < minimum {
        Err(parse(format!("{field} is truncated")))
    } else {
        Ok(())
    }
}
fn words_to_bytes(words: u32) -> Result<usize> {
    usize::try_from(words)
        .map_err(|_error| parse("WMF size does not fit usize"))?
        .checked_mul(2)
        .ok_or_else(|| parse("WMF byte size overflows"))
}
fn usize_at_u32(data: &[u8], offset: usize) -> Result<usize> {
    usize::try_from(u32_at(data, offset)?)
        .map_err(|_error| parse("record length does not fit usize"))
}
fn u16_at(data: &[u8], offset: usize) -> Result<u16> {
    Ok(u16::from_le_bytes(array_at(data, offset)?))
}
fn i16_at(data: &[u8], offset: usize) -> Result<i16> {
    Ok(i16::from_le_bytes(array_at(data, offset)?))
}
fn u32_at(data: &[u8], offset: usize) -> Result<u32> {
    Ok(u32::from_le_bytes(array_at(data, offset)?))
}
fn i32_at(data: &[u8], offset: usize) -> Result<i32> {
    Ok(i32::from_le_bytes(array_at(data, offset)?))
}
fn array_at<const N: usize>(data: &[u8], offset: usize) -> Result<[u8; N]> {
    let end = offset
        .checked_add(N)
        .ok_or_else(|| parse("record field offset overflows"))?;
    data.get(offset..end)
        .ok_or_else(|| parse("record field is truncated"))?
        .try_into()
        .map_err(|_error| parse("record field length is invalid"))
}

fn encode_png(image: &RgbaImage, maximum: usize) -> Result<Vec<u8>> {
    let mut writer = LimitedWriter::new(maximum)?;
    DynamicImage::ImageRgba8(image.clone())
        .write_to(&mut writer, ImageFormat::Png)
        .map_err(|error| parse(format!("PNG encoding failed: {error}")))?;
    Ok(writer.into_inner())
}
struct LimitedWriter {
    inner: Cursor<Vec<u8>>,
    maximum: u64,
}
impl LimitedWriter {
    fn new(maximum: usize) -> Result<Self> {
        Ok(Self {
            inner: Cursor::new(Vec::new()),
            maximum: u64::try_from(maximum)
                .map_err(|_error| parse("output limit does not fit u64"))?,
        })
    }
    fn into_inner(self) -> Vec<u8> {
        self.inner.into_inner()
    }
}
impl Write for LimitedWriter {
    fn write(&mut self, buf: &[u8]) -> std::io::Result<usize> {
        let len = u64::try_from(buf.len())
            .map_err(|_error| std::io::Error::other("output length overflow"))?;
        if self
            .inner
            .position()
            .checked_add(len)
            .is_none_or(|end| end > self.maximum)
        {
            return Err(std::io::Error::other("output limit exceeded"));
        }
        self.inner.write(buf)
    }
    fn flush(&mut self) -> std::io::Result<()> {
        self.inner.flush()
    }
}
impl Seek for LimitedWriter {
    fn seek(&mut self, pos: SeekFrom) -> std::io::Result<u64> {
        let target = match pos {
            SeekFrom::Start(value) => Some(value),
            SeekFrom::Current(delta) => signed_add(self.inner.position(), delta),
            SeekFrom::End(delta) => signed_add(
                u64::try_from(self.inner.get_ref().len())
                    .map_err(|_error| std::io::Error::other("output length overflow"))?,
                delta,
            ),
        }
        .ok_or_else(|| std::io::Error::other("invalid output seek"))?;
        if target > self.maximum {
            return Err(std::io::Error::other("output limit exceeded"));
        }
        self.inner.seek(SeekFrom::Start(target))
    }
}
fn signed_add(base: u64, delta: i64) -> Option<u64> {
    if delta >= 0 {
        base.checked_add(delta.cast_unsigned())
    } else {
        base.checked_sub(delta.unsigned_abs())
    }
}
fn parse(message: impl Into<String>) -> Error {
    Error::ParseError(message.into())
}
fn unsupported(message: impl Into<String>) -> Error {
    Error::Unsupported(message.into())
}
fn limit(resource: &str, maximum: usize) -> Error {
    parse(format!("{resource} exceeds configured limit {maximum}"))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn dib() -> Vec<u8> {
        let mut d = vec![0; 40];
        d[0..4].copy_from_slice(&40u32.to_le_bytes());
        d[4..8].copy_from_slice(&1i32.to_le_bytes());
        d[8..12].copy_from_slice(&1i32.to_le_bytes());
        d[12..14].copy_from_slice(&1u16.to_le_bytes());
        d[14..16].copy_from_slice(&24u16.to_le_bytes());
        d.extend_from_slice(&[1, 2, 3, 0]);
        d
    }
    fn put32(data: &mut [u8], offset: usize, value: u32) {
        data[offset..offset + 4].copy_from_slice(&value.to_le_bytes())
    }
    fn puti(data: &mut [u8], offset: usize, value: i32) {
        data[offset..offset + 4].copy_from_slice(&value.to_le_bytes())
    }
    fn identity(data: &mut [u8], offset: usize) {
        put32(data, offset, 1f32.to_bits());
        put32(data, offset + 12, 1f32.to_bits())
    }
    fn emf_payload(len: usize, bmi_off: usize, bits_off: usize) -> Vec<u8> {
        let d = dib();
        let mut p = vec![0; len + d.len()];
        puti(&mut p, 24, 1);
        puti(&mut p, 28, 1);
        identity(&mut p, 44);
        put32(&mut p, bmi_off, (len + 8) as u32);
        put32(&mut p, bmi_off + 4, 40);
        put32(&mut p, bits_off, (len + 48) as u32);
        put32(&mut p, bits_off + 4, 4);
        p[len..len + d.len()].copy_from_slice(&d);
        p
    }

    #[test]
    fn rop_truth_table() {
        let copy = Rop3::from_code(SRCCOPY);
        assert_eq!(copy.eval_byte(0x55, 0xa6, 0x39), 0xa6);
        assert!(!copy.depends_on_destination());
        assert!(!copy.depends_on_pattern());
        let xor = Rop3::from_code(0x0066_0046);
        assert_eq!(xor.eval_byte(0, 0xa5, 0x0f), 0xaa);
        assert!(xor.depends_on_destination());
        let mut destination = [0x0f, 0xf0];
        xor.composite_bytes(&[0; 2], &[0xa5, 0x5a], &mut destination)
            .unwrap();
        assert_eq!(destination, [0xaa, 0xaa]);
    }
    #[test]
    fn all_emf_kinds_reject_truncation() {
        for kind in [76, 77, 78, 79, 80, 81, 114, 116] {
            assert!(
                BitmapOp::parse_emf(kind, &[], StretchPolicy::default(), DibLimits::default())
                    .is_err()
            );
        }
    }
    #[test]
    fn parses_emf_bitblt_and_adjusts_record_offsets() {
        let mut p = emf_payload(92, 76, 84);
        put32(&mut p, 32, SRCCOPY);
        let op =
            BitmapOp::parse_emf(76, &p, StretchPolicy::Halftone, DibLimits::default()).unwrap();
        assert_eq!(op.kind, BitmapRecordKind::EmfBitBlt);
        assert!(matches!(op.source, Some(BitmapSource::Dib(_))));
    }
    #[test]
    fn malformed_emf_offsets_are_rejected() {
        let mut p = emf_payload(92, 76, 84);
        put32(&mut p, 76, 4);
        assert!(
            BitmapOp::parse_emf(76, &p, StretchPolicy::default(), DibLimits::default()).is_err()
        );
        put32(&mut p, 76, 100);
        put32(&mut p, 44, f32::NAN.to_bits());
        assert!(
            BitmapOp::parse_emf(76, &p, StretchPolicy::default(), DibLimits::default()).is_err()
        );
        put32(&mut p, 76, u32::MAX);
        assert!(
            BitmapOp::parse_emf(76, &p, StretchPolicy::default(), DibLimits::default()).is_err()
        );
    }
    #[test]
    fn parses_wmf_stretchdib() {
        let d = dib();
        let mut params = vec![0; 22 + d.len()];
        put32(&mut params, 0, SRCCOPY);
        params[6..8].copy_from_slice(&1i16.to_le_bytes());
        params[8..10].copy_from_slice(&1i16.to_le_bytes());
        params[12..14].copy_from_slice(&1i16.to_le_bytes());
        params[14..16].copy_from_slice(&1i16.to_le_bytes());
        params[16..18].copy_from_slice(&1i16.to_le_bytes());
        params[22..].copy_from_slice(&d);
        let words = ((params.len() + 6) / 2) as u32;
        let op = BitmapOp::parse_wmf(
            words,
            0x0f43,
            &params,
            StretchPolicy::default(),
            DibLimits::default(),
        )
        .unwrap();
        assert_eq!(op.kind, BitmapRecordKind::WmfStretchDib);
    }
    #[test]
    fn wmf_record_size_mismatch_is_rejected() {
        assert!(
            BitmapOp::parse_wmf(
                20,
                0x0f43,
                &[0; 22],
                StretchPolicy::default(),
                DibLimits::default()
            )
            .is_err()
        );
    }
    #[test]
    fn svg_copy_embeds_cropped_png() {
        let d = dib();
        let source = Dib::parse(&d, DibLimits::default()).unwrap();
        let op = BitmapOp {
            kind: BitmapRecordKind::WmfStretchDib,
            source: Some(BitmapSource::Dib(source)),
            mask: None,
            source_rect: SignedRect {
                x: 0,
                y: 0,
                width: 1,
                height: 1,
            },
            destination: SignedRect {
                x: 3,
                y: 4,
                width: 5,
                height: 6,
            },
            parallelogram: None,
            source_transform: AffineTransform::IDENTITY,
            rop3: Rop3::from_code(SRCCOPY),
            background_rop3: None,
            scan_range: None,
            alpha: None,
            transparent_color: None,
            stretch: StretchPolicy::Halftone,
            limits: DibLimits::default(),
        };
        let svg = op.to_svg_image().unwrap();
        assert!(svg.element.contains("data:image/png;base64,"));
        assert!(svg.element.contains("x=\"3\""));
    }
    #[test]
    fn svg_rejects_destination_dependent_rop() {
        let d = dib();
        let source = Dib::parse(&d, DibLimits::default()).unwrap();
        let mut op = BitmapOp {
            kind: BitmapRecordKind::EmfBitBlt,
            source: Some(BitmapSource::Dib(source)),
            mask: None,
            source_rect: SignedRect {
                x: 0,
                y: 0,
                width: 1,
                height: 1,
            },
            destination: SignedRect {
                x: 0,
                y: 0,
                width: 1,
                height: 1,
            },
            parallelogram: None,
            source_transform: AffineTransform::IDENTITY,
            rop3: Rop3::from_code(SRCCOPY),
            background_rop3: None,
            scan_range: None,
            alpha: None,
            transparent_color: None,
            stretch: StretchPolicy::default(),
            limits: DibLimits::default(),
        };
        op.rop3 = Rop3::from_code(0x0066_0046);
        assert!(matches!(op.to_svg_image(), Err(Error::Unsupported(_))));
    }
}
