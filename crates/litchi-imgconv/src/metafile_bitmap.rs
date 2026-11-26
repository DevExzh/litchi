//! Normalized, checked bitmap operations shared by EMF and WMF renderers.

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

/// A brush pattern bitmap, tiled through an SVG `<pattern>`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct BrushPattern {
    width: u32,
    height: u32,
    pixels: PatternPixels,
}

#[derive(Debug, Clone, PartialEq, Eq)]
enum PatternPixels {
    /// Bits of a monochrome device-dependent bitmap, row by row from the top.
    ///
    /// GDI paints the 0 bits of such a pattern with the text colour and the 1
    /// bits with the background colour of the device context that uses it.
    Monochrome(Vec<bool>),
    /// Pixels of a pattern that carries its own colours.
    Color(RgbaImage),
}

impl BrushPattern {
    /// Decodes a DIB pattern, whose colours come from its own color table.
    pub fn from_dib(dib: &Dib<'_>) -> Result<Self> {
        let image = dib.to_dynamic_image()?.to_rgba8();
        let (width, height) = image.dimensions();
        Ok(Self {
            width,
            height,
            pixels: PatternPixels::Color(image),
        })
    }

    /// Decodes the pattern of a monochrome brush from a 1-bpp DIB.
    ///
    /// Only the bit values matter: the device context supplies both colours.
    pub fn from_monochrome_dib(dib: &Dib<'_>) -> Result<Self> {
        let info = dib.info();
        if info.bit_count != 1 || info.compression != crate::dib::Compression::Rgb {
            return Self::from_dib(dib);
        }
        let stride = info
            .stride
            .ok_or_else(|| parse("monochrome pattern stride is unavailable"))?;
        Self::from_monochrome_bits(
            info.width,
            info.height,
            stride,
            dib.bitmap_bits(),
            info.top_down,
        )
    }

    /// Copies the bits of a monochrome bitmap whose rows are `stride` bytes
    /// apart, listed from the top row unless `top_down` is false.
    pub fn from_monochrome_bits(
        width: u32,
        height: u32,
        stride: usize,
        bits: &[u8],
        top_down: bool,
    ) -> Result<Self> {
        let columns = usize::try_from(width).map_err(|_error| parse("pattern width is invalid"))?;
        let rows = usize::try_from(height).map_err(|_error| parse("pattern height is invalid"))?;
        let count = columns
            .checked_mul(rows)
            .ok_or_else(|| parse("pattern pixel count overflows"))?;
        if columns == 0 || rows == 0 || stride.checked_mul(8).is_none_or(|bits| bits < columns) {
            return Err(parse("monochrome pattern geometry is invalid"));
        }
        let needed = stride
            .checked_mul(rows)
            .ok_or_else(|| parse("pattern bitmap size overflows"))?;
        if bits.len() < needed {
            return Err(parse("brush pattern bits are truncated"));
        }
        let mut mask = Vec::new();
        mask.try_reserve_exact(count)
            .map_err(|source| Error::Allocation {
                resource: "brush pattern",
                source,
            })?;
        for row in 0..rows {
            let source_row = if top_down { row } else { rows - 1 - row };
            let line = &bits[source_row * stride..(source_row + 1) * stride];
            for column in 0..columns {
                mask.push(line[column / 8] & (0x80 >> (column % 8)) != 0);
            }
        }
        Ok(Self {
            width,
            height,
            pixels: PatternPixels::Monochrome(mask),
        })
    }

    /// Pattern width in device pixels.
    #[must_use]
    pub const fn width(&self) -> u32 {
        self.width
    }

    /// Pattern height in device pixels.
    #[must_use]
    pub const fn height(&self) -> u32 {
        self.height
    }

    /// Whether the device context supplies the colours of this pattern.
    #[must_use]
    pub const fn is_monochrome(&self) -> bool {
        matches!(self.pixels, PatternPixels::Monochrome(_))
    }

    /// Approximate retained heap size, for object-table accounting.
    #[must_use]
    pub fn retained_bytes(&self) -> usize {
        match &self.pixels {
            PatternPixels::Monochrome(mask) => mask.len(),
            PatternPixels::Color(image) => image.as_raw().len(),
        }
    }

    /// Tile pixels for the given text and background `COLORREF`s.
    fn tile(&self, text_color: u32, background_color: u32) -> RgbaImage {
        match &self.pixels {
            PatternPixels::Color(image) => image.clone(),
            PatternPixels::Monochrome(mask) => {
                let text = colorref_rgba(text_color);
                let background = colorref_rgba(background_color);
                let mut image = RgbaImage::new(self.width, self.height);
                for (pixel, bit) in image.pixels_mut().zip(mask) {
                    pixel.0 = if *bit { background } else { text };
                }
                image
            },
        }
    }

    /// Encodes one tile for the given text and background `COLORREF`s as
    /// PNG, bounded by `maximum` bytes.
    pub fn tile_png(
        &self,
        text_color: u32,
        background_color: u32,
        maximum: usize,
    ) -> Result<Vec<u8>> {
        encode_png(&self.tile(text_color, background_color), maximum)
    }

    /// The pattern whose pixel colours, resolved with the text and
    /// background `COLORREF`s, go through `map`: `Some(Some(color))` paints a
    /// colour and `Some(None)` leaves the pixel transparent. Returns `None`
    /// when `map` rejects any pixel.
    #[must_use]
    pub fn mapped(
        &self,
        text_color: u32,
        background_color: u32,
        map: impl Fn(u32) -> Option<Option<u32>>,
    ) -> Option<Self> {
        let mut tile = self.tile(text_color, background_color);
        for pixel in tile.pixels_mut() {
            pixel.0 = match map(rgba_colorref(pixel.0))? {
                Some(color) => colorref_rgba(color),
                None => [0, 0, 0, 0],
            };
        }
        Some(Self {
            width: self.width,
            height: self.height,
            pixels: PatternPixels::Color(tile),
        })
    }

    /// The pattern as a brush paints it under `rop2`: every pixel either
    /// keeps the destination or becomes a colour. Returns `None` when a pixel
    /// combines with the destination.
    #[must_use]
    pub fn under_rop2(&self, rop2: Rop2, text_color: u32, background_color: u32) -> Option<Self> {
        self.mapped(text_color, background_color, |color| {
            match rop2.paint(color) {
                Rop2Paint::Unchanged => Some(None),
                Rop2Paint::Color(color) => Some(Some(color)),
                Rop2Paint::DestinationDependent => None,
            }
        })
    }

    /// Emits an SVG pattern definition that tiles a PNG tile.
    ///
    /// `tile` is the size of one pattern repetition and `origin` the brush
    /// origin, both in SVG user units.
    pub fn svg_definition(
        id: &str,
        png: &[u8],
        tile: (f64, f64),
        origin: (f64, f64),
    ) -> Result<String> {
        if !tile.0.is_finite() || !tile.1.is_finite() || tile.0 <= 0.0 || tile.1 <= 0.0 {
            return Err(parse("brush pattern tile size is invalid"));
        }
        if !origin.0.is_finite() || !origin.1.is_finite() {
            return Err(parse("brush pattern origin is invalid"));
        }
        let mut definition = String::with_capacity(png.len() * 4 / 3 + 256);
        definition.push_str("<pattern id=\"");
        definition.push_str(id);
        definition.push_str("\" patternUnits=\"userSpaceOnUse\" x=\"");
        crate::svg_utils::write_num(&mut definition, origin.0);
        definition.push_str("\" y=\"");
        crate::svg_utils::write_num(&mut definition, origin.1);
        definition.push_str("\" width=\"");
        crate::svg_utils::write_num(&mut definition, tile.0);
        definition.push_str("\" height=\"");
        crate::svg_utils::write_num(&mut definition, tile.1);
        definition.push_str("\"><image width=\"");
        crate::svg_utils::write_num(&mut definition, tile.0);
        definition.push_str("\" height=\"");
        crate::svg_utils::write_num(&mut definition, tile.1);
        definition.push_str(
            "\" preserveAspectRatio=\"none\" image-rendering=\"optimizeSpeed\" href=\"data:image/png;base64,",
        );
        STANDARD.encode_string(png, &mut definition);
        definition.push_str("\"/></pattern>");
        Ok(definition)
    }
}

/// Converts a `COLORREF` (`0x00bbggrr`) to opaque RGBA bytes.
fn colorref_rgba(color: u32) -> [u8; 4] {
    let [red, green, blue, _] = color.to_le_bytes();
    [red, green, blue, 255]
}

/// Converts RGBA bytes to a `COLORREF` (`0x00bbggrr`), ignoring alpha.
fn rgba_colorref(pixel: [u8; 4]) -> u32 {
    u32::from_le_bytes([pixel[0], pixel[1], pixel[2], 0])
}

/// The colour bits a `COLORREF` contributes to a raster operation.
const COLOR_BITS: u32 = 0x00ff_ffff;

/// A binary raster operation ([MS-WMF] 2.1.1.6), the mix mode with which
/// pens and brushes combine with the destination.
///
/// The truth table is indexed by `(P << 1) | D`, so the R2_* mode value is
/// the table plus one.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Rop2(u8);

/// What a pen or brush colour paints under a binary raster operation.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Rop2Paint {
    /// The destination keeps every pixel.
    Unchanged,
    /// The destination becomes this `COLORREF`, whatever it was.
    Color(u32),
    /// The result combines the colour with destination pixels.
    DestinationDependent,
}

impl Rop2 {
    /// R2_COPYPEN, the default mix mode.
    pub const COPY_PEN: Self = Self(12);

    /// Maps an R2_* mode value, or `None` for values outside 1..=16.
    #[must_use]
    pub fn from_mode(mode: u32) -> Option<Self> {
        (1..=16)
            .contains(&mode)
            .then(|| Self(u8::try_from(mode - 1).unwrap_or(12)))
    }

    /// The R2_* mode value.
    #[must_use]
    pub const fn mode(self) -> u32 {
        self.0 as u32 + 1
    }

    /// Resolves what painting `color` does to the destination.
    #[must_use]
    pub fn paint(self, color: u32) -> Rop2Paint {
        let ones = color & COLOR_BITS;
        let zeros = !color & COLOR_BITS;
        // Two-bit truth tables `f(D=1) f(D=0)` for pen bits 0 and 1.
        let for_zero = self.0 & 0b11;
        let for_one = (self.0 >> 2) & 0b11;
        let keep = |table: u8, bits: u32| bits == 0 || table == 0b10;
        let constant = |table: u8, bits: u32| bits == 0 || table == 0b00 || table == 0b11;
        if keep(for_zero, zeros) && keep(for_one, ones) {
            return Rop2Paint::Unchanged;
        }
        if constant(for_zero, zeros) && constant(for_one, ones) {
            let mut value = 0;
            if for_zero == 0b11 {
                value |= zeros;
            }
            if for_one == 0b11 {
                value |= ones;
            }
            return Rop2Paint::Color(value);
        }
        Rop2Paint::DestinationDependent
    }

    /// Applies the mix mode to every bit of a pen and a destination word.
    #[must_use]
    pub const fn apply(self, pen: u32, destination: u32) -> u32 {
        let mut output = 0;
        let mut index = 0;
        while index < 4 {
            if self.0 & (1 << index) != 0 {
                let pen = if index & 2 != 0 { pen } else { !pen };
                let destination = if index & 1 != 0 {
                    destination
                } else {
                    !destination
                };
                output |= pen & destination;
            }
            index += 1;
        }
        output
    }

    /// Whether painting twice with one colour paints the same as once.
    #[must_use]
    pub const fn is_idempotent(self) -> bool {
        // A two-bit truth table that inverts the destination bit is the
        // only one that changes when applied twice.
        self.0 & 0b11 != 0b01 && (self.0 >> 2) & 0b11 != 0b01
    }

    /// Whether the result ignores the pen or brush colour.
    #[must_use]
    pub const fn ignores_color(self) -> bool {
        self.0 & 0b11 == (self.0 >> 2) & 0b11
    }
}

/// The eight-bit ROP3 truth table stored in bits 16..23 of a raster-op code.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Rop3(u8);

/// How a bitmap record paints its destination rectangle.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum BitmapPaint {
    /// The record leaves every destination pixel unchanged.
    Nothing,
    /// Fills the destination rectangle with one solid `COLORREF` (`0x00bbggrr`).
    Solid(u32),
    /// Fills the destination rectangle with the selected brush.
    Brush,
    /// Draws the source bitmap, or its inverse for NOTSRCCOPY.
    Image,
    /// Combines source or brush pixels with destination pixels, which needs
    /// raster composition with the operations around it.
    Composite,
}

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

    /// Applies the truth table independently to every bit of three words.
    #[must_use]
    pub const fn apply(self, pattern: u32, source: u32, destination: u32) -> u32 {
        let mut output = 0;
        let mut index = 0;
        while index < 8 {
            if self.0 & (1 << index) != 0 {
                let pattern = if index & 4 != 0 { pattern } else { !pattern };
                let source = if index & 2 != 0 { source } else { !source };
                let destination = if index & 1 != 0 {
                    destination
                } else {
                    !destination
                };
                output |= pattern & source & destination;
            }
            index += 1;
        }
        output
    }

    /// Composites equal-length byte planes in place.
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
    pub const fn depends_on_destination(self) -> bool {
        self.0 & 0xaa != (self.0 & 0x55) << 1
    }

    #[must_use]
    pub const fn depends_on_pattern(self) -> bool {
        self.0 >> 4 != self.0 & 0x0f
    }

    #[must_use]
    pub const fn depends_on_source(self) -> bool {
        self.0 & 0xcc != (self.0 & 0x33) << 2
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

    /// Resolves how this record paints its destination rectangle.
    ///
    /// A ROP3 is reproduced directly when its output is a constant, the
    /// selected brush, or the source bitmap; every other truth table needs
    /// [`RasterComposite`].
    pub fn paint(&self) -> Result<BitmapPaint> {
        if rect_extent(self.destination).is_none() {
            return Ok(BitmapPaint::Nothing);
        }
        let table = self.rop3.table();
        if table == 0xaa {
            return Ok(BitmapPaint::Nothing);
        }
        let paint = if self.rop3.depends_on_destination() {
            BitmapPaint::Composite
        } else if !self.rop3.depends_on_source() {
            match table {
                0x00 => BitmapPaint::Solid(0),
                0xff => BitmapPaint::Solid(COLOR_BITS),
                0xf0 => BitmapPaint::Brush,
                _ => BitmapPaint::Composite,
            }
        } else if !self.rop3.depends_on_pattern() && matches!(table, 0xcc | 0x33) {
            BitmapPaint::Image
        } else {
            BitmapPaint::Composite
        };
        if self.rop3.depends_on_source() {
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
            if self.source.is_none() {
                return Err(parse("bitmap operation has no embedded source"));
            }
        }
        Ok(paint)
    }

    /// Whether [`RasterComposite`] can reproduce this record: a plain ROP3
    /// blit without alpha, colour key, parallelogram or mask.
    #[must_use]
    pub fn is_plain_blit(&self) -> bool {
        self.alpha.is_none()
            && self.transparent_color.is_none()
            && self.parallelogram.is_none()
            && self.mask.is_none()
            && self.source_transform.is_identity()
    }

    /// Destination rectangle with positive extents.
    pub fn destination_rect(&self) -> Result<SignedRect> {
        normalized(self.destination)
    }

    /// Source pixels the record reads and the part of the destination they
    /// paint, without decoding the bitmap. `Ok(None)` means no source pixel
    /// lies inside the bitmap, so the record paints nothing.
    pub fn blit_window(&self) -> Result<Option<BlitWindow>> {
        match self.source {
            Some(BitmapSource::Dib(dib)) => {
                self.clip_to_bitmap(dib.info().width, dib.info().height)
            },
            Some(BitmapSource::DeviceDependent(_)) => Err(unsupported(
                "WMF Bitmap16 sources require device palette playback",
            )),
            None => Err(parse("bitmap operation has no embedded source")),
        }
    }

    /// Decodes the source pixels the record reads, oriented like the
    /// destination.
    pub fn blit_pixels(&self) -> Result<Option<BlitPixels>> {
        let Some(window) = self.blit_window()? else {
            return Ok(None);
        };
        let Some(BitmapSource::Dib(source)) = self.source else {
            return Err(parse("bitmap operation has no embedded source"));
        };
        let alpha_mode = if self.alpha.is_some_and(|blend| blend.source_alpha) {
            AlphaInterpretation::Premultiplied
        } else {
            AlphaInterpretation::Ignore
        };
        let decoded = source.to_dynamic_image_with_alpha(alpha_mode)?;
        if decoded.width() != source.info().width || decoded.height() != source.info().height {
            return Err(parse("decoded DIB dimensions differ from its header"));
        }
        let crop = |value: i32| {
            u32::try_from(value).map_err(|_error| parse("clipped source rectangle is invalid"))
        };
        let image = decoded.crop_imm(
            crop(window.source.x)?,
            crop(window.source.y)?,
            crop(window.source.width)?,
            crop(window.source.height)?,
        );
        let image = if window.flip_x { image.fliph() } else { image };
        let image = if window.flip_y { image.flipv() } else { image };
        Ok(Some(BlitPixels {
            image: image.to_rgba8(),
            window,
        }))
    }

    /// Emits a self-contained SVG `<image>` for source-copy, alpha, or
    /// color-key operations at the logical destination rectangle.
    ///
    /// Returns `Ok(None)` when the record paints nothing.
    pub fn to_svg_image(&self) -> Result<Option<SvgImage>> {
        self.to_svg_image_at(None)
    }

    /// Emits the image at an already transformed destination rectangle.
    ///
    /// `destination_override` replaces the logical destination rectangle; the
    /// part of the source outside the bitmap is clipped away from it
    /// proportionally, as GDI does. Returns `Ok(None)` when the record paints
    /// nothing.
    pub fn to_svg_image_at(
        &self,
        destination_override: Option<[f64; 4]>,
    ) -> Result<Option<SvgImage>> {
        if self.paint()? != BitmapPaint::Image {
            return Err(unsupported(
                "bitmap operation does not draw its source image",
            ));
        }
        let Some(BlitPixels { image, window }) = self.blit_pixels()? else {
            return Ok(None);
        };
        let mut rgba = image;
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
        let placement = match (self.parallelogram, destination_override) {
            (Some(points), None) => {
                let [x0, y0, x1, y1] = window.fraction;
                let width = f64::from(rgba.width());
                let height = f64::from(rgba.height());
                let u = (
                    f64::from(points[1].x) - f64::from(points[0].x),
                    f64::from(points[1].y) - f64::from(points[0].y),
                );
                let v = (
                    f64::from(points[2].x) - f64::from(points[0].x),
                    f64::from(points[2].y) - f64::from(points[0].y),
                );
                ImagePlacement::Matrix([
                    u.0 * (x1 - x0) / width,
                    u.1 * (x1 - x0) / width,
                    v.0 * (y1 - y0) / height,
                    v.1 * (y1 - y0) / height,
                    f64::from(points[0].x) + x0 * u.0 + y0 * v.0,
                    f64::from(points[0].y) + x0 * u.1 + y0 * v.1,
                ])
            },
            (_, Some(destination)) => ImagePlacement::Rect(window.painted(destination)),
            (None, None) => {
                let destination = normalized(self.destination)?;
                ImagePlacement::Rect(window.painted([
                    f64::from(destination.x),
                    f64::from(destination.y),
                    f64::from(destination.width),
                    f64::from(destination.height),
                ]))
            },
        };
        image_element(&rgba, placement, opacity, self.limits.max_output_bytes).map(Some)
    }

    /// Intersects the source rectangle with a `width` x `height` bitmap.
    ///
    /// GDI clips a blit whose source rectangle extends past the bitmap edges
    /// and paints only the proportionally corresponding destination part.
    fn clip_to_bitmap(&self, width: u32, height: u32) -> Result<Option<BlitWindow>> {
        if rect_extent(self.destination).is_none() || rect_extent(self.source_rect).is_none() {
            return Ok(None);
        }
        let source = normalized(self.source_rect)?;
        let x0 = i64::from(source.x);
        let y0 = i64::from(source.y);
        let x1 = x0 + i64::from(source.width);
        let y1 = y0 + i64::from(source.height);
        let left = x0.max(0);
        let top = y0.max(0);
        let right = x1.min(i64::from(width));
        let bottom = y1.min(i64::from(height));
        if right <= left || bottom <= top {
            return Ok(None);
        }
        let span = |start: i64, end: i64, origin: i64, extent: i32| {
            let extent = f64::from(extent);
            #[allow(
                clippy::cast_precision_loss,
                reason = "the values are bounded by i32 bitmap coordinates"
            )]
            [
                (start - origin) as f64 / extent,
                (end - origin) as f64 / extent,
            ]
        };
        let mut horizontal = span(left, right, x0, source.width);
        let mut vertical = span(top, bottom, y0, source.height);
        let flip_x = (self.source_rect.width < 0) != (self.destination.width < 0);
        let flip_y = (self.source_rect.height < 0) != (self.destination.height < 0);
        if flip_x {
            horizontal = [1.0 - horizontal[1], 1.0 - horizontal[0]];
        }
        if flip_y {
            vertical = [1.0 - vertical[1], 1.0 - vertical[0]];
        }
        let narrow = |value: i64| {
            i32::try_from(value).map_err(|_error| parse("clipped source rectangle overflows"))
        };
        Ok(Some(BlitWindow {
            source: SignedRect {
                x: narrow(left)?,
                y: narrow(top)?,
                width: narrow(right - left)?,
                height: narrow(bottom - top)?,
            },
            fraction: [horizontal[0], vertical[0], horizontal[1], vertical[1]],
            flip_x,
            flip_y,
        }))
    }
}

/// Source pixels a blit reads and the part of its destination they paint.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct BlitWindow {
    /// Source rectangle inside the bitmap, in top-down bitmap pixels.
    pub source: SignedRect,
    /// Painted part of the destination rectangle as fractions
    /// `[x0, y0, x1, y1]` of its width and height.
    pub fraction: [f64; 4],
    /// The source is mirrored horizontally relative to the destination.
    pub flip_x: bool,
    /// The source is mirrored vertically relative to the destination.
    pub flip_y: bool,
}

impl BlitWindow {
    /// The painted part of `destination` (`[x, y, width, height]`).
    #[must_use]
    pub fn painted(&self, destination: [f64; 4]) -> [f64; 4] {
        let [x, y, width, height] = destination;
        let [x0, y0, x1, y1] = self.fraction;
        [
            x + x0 * width,
            y + y0 * height,
            (x1 - x0) * width,
            (y1 - y0) * height,
        ]
    }

    /// Number of source pixels across and down.
    #[must_use]
    pub fn pixels(&self) -> (u32, u32) {
        (
            u32::try_from(self.source.width).unwrap_or(0),
            u32::try_from(self.source.height).unwrap_or(0),
        )
    }
}

/// Decoded source pixels of a blit, oriented like its destination.
#[derive(Debug, Clone, PartialEq)]
pub struct BlitPixels {
    pub image: RgbaImage,
    pub window: BlitWindow,
}

/// Where an `<image>` element is placed.
#[derive(Debug, Clone, Copy, PartialEq)]
enum ImagePlacement {
    /// `[x, y, width, height]` in user units.
    Rect([f64; 4]),
    /// Affine matrix applied to the image in pixel units.
    Matrix([f64; 6]),
}

/// Emits a PNG-backed `<image>` element.
fn image_element(
    image: &RgbaImage,
    placement: ImagePlacement,
    opacity: f64,
    maximum: usize,
) -> Result<SvgImage> {
    let png = encode_png(image, maximum)?;
    let encoded_len = png
        .len()
        .checked_add(2)
        .and_then(|length| length.checked_div(3))
        .and_then(|groups| groups.checked_mul(4))
        .and_then(|length| length.checked_add(256))
        .ok_or_else(|| parse("SVG bitmap data URL size overflow"))?;
    if encoded_len > maximum {
        return Err(limit("SVG bitmap data URL", maximum));
    }
    let mut element = String::with_capacity(encoded_len);
    match placement {
        ImagePlacement::Rect([x, y, width, height]) => {
            let _ = write!(
                &mut element,
                "<image x=\"{x}\" y=\"{y}\" width=\"{width}\" height=\"{height}\" opacity=\"{opacity}\" preserveAspectRatio=\"none\" href=\"data:image/png;base64,"
            );
        },
        ImagePlacement::Matrix([a, b, c, d, e, f]) => {
            let _ = write!(
                &mut element,
                "<image width=\"{}\" height=\"{}\" opacity=\"{opacity}\" preserveAspectRatio=\"none\" transform=\"matrix({a} {b} {c} {d} {e} {f})\" href=\"data:image/png;base64,",
                image.width(),
                image.height(),
            );
        },
    }
    STANDARD.encode_string(png, &mut element);
    element.push_str("\"/>");
    Ok(SvgImage { element })
}

/// What a pixel becomes when two equal XOR passes of `xor` bracket a fill
/// of `color` under `rop2`: every bit turns into `m(D ^ xor) ^ xor`.
///
/// Returns `Some(None)` when the pixel keeps the destination, `Some(Some(c))`
/// when it becomes the `COLORREF` `c`, and `None` when it still depends on
/// the destination.
#[must_use]
pub fn xor_bracket_paint(rop2: Rop2, color: u32, xor: u32) -> Option<Option<u32>> {
    let low = (rop2.apply(color, xor) ^ xor) & COLOR_BITS;
    let high = (rop2.apply(color, !xor) ^ xor) & COLOR_BITS;
    if low == 0 && high == COLOR_BITS {
        Some(None)
    } else if low == high {
        Some(Some(low))
    } else {
        None
    }
}

/// Whether a polygon outline is convex.
#[must_use]
pub fn is_convex(outline: &[(i64, i64)]) -> bool {
    let mut sign = 0i64;
    for index in 0..outline.len() {
        let a = outline[index];
        let b = outline[(index + 1) % outline.len()];
        let c = outline[(index + 2) % outline.len()];
        let cross = (b.0 - a.0) * (c.1 - b.1) - (b.1 - a.1) * (c.0 - b.0);
        if cross != 0 {
            if sign != 0 && cross.signum() != sign {
                return false;
            }
            sign = cross.signum();
        }
    }
    true
}

/// Whether the interiors of any two convex outlines overlap; touching
/// edges do not count. Exact for integer coordinates.
#[must_use]
pub fn convex_outlines_overlap(outlines: &[&[(i64, i64)]]) -> bool {
    let bounds = |outline: &[(i64, i64)]| {
        outline.iter().fold(
            (i64::MAX, i64::MAX, i64::MIN, i64::MIN),
            |(x0, y0, x1, y1), &(x, y)| (x0.min(x), y0.min(y), x1.max(x), y1.max(y)),
        )
    };
    let boxes = outlines
        .iter()
        .map(|outline| bounds(outline))
        .collect::<Vec<_>>();
    for first in 0..outlines.len() {
        for second in first + 1..outlines.len() {
            let (a, b) = (boxes[first], boxes[second]);
            if a.2 <= b.0 || b.2 <= a.0 || a.3 <= b.1 || b.3 <= a.1 {
                continue;
            }
            if convex_interiors_overlap(outlines[first], outlines[second]) {
                return true;
            }
        }
    }
    false
}

/// Separating-axis test for two convex outlines; touching edges do not
/// count as overlap.
fn convex_interiors_overlap(first: &[(i64, i64)], second: &[(i64, i64)]) -> bool {
    for outline in [first, second] {
        for index in 0..outline.len() {
            let (x0, y0) = outline[index];
            let (x1, y1) = outline[(index + 1) % outline.len()];
            let axis = (y0 - y1, x1 - x0);
            if axis == (0, 0) {
                continue;
            }
            let project = |points: &[(i64, i64)]| {
                points
                    .iter()
                    .fold((i64::MAX, i64::MIN), |(low, high), &(x, y)| {
                        let value = axis.0 * x + axis.1 * y;
                        (low.min(value), high.max(value))
                    })
            };
            let (a_low, a_high) = project(first);
            let (b_low, b_high) = project(second);
            if a_high <= b_low || b_high <= a_low {
                return false;
            }
        }
    }
    true
}

/// Largest distance from a whole cell, in cells, that still counts as
/// aligned. It absorbs floating-point error from coordinate transforms.
const GRID_TOLERANCE: f64 = 1e-3;

/// The pixel grid on which consecutive raster operations are composed.
///
/// Operations whose painted rectangles fall on whole cells of one grid touch
/// the same destination pixels, so their raster operations can be composed
/// cell by cell.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct BlitGrid {
    origin: (f64, f64),
    cell: (f64, f64),
    /// Covered cells `[column0, row0, column1, row1)` relative to `origin`.
    bounds: [i64; 4],
}

impl BlitGrid {
    /// The grid of `columns` x `rows` source pixels painted into `painted`
    /// (`[x, y, width, height]`).
    #[must_use]
    pub fn for_pixels(painted: [f64; 4], columns: u32, rows: u32) -> Option<Self> {
        let [x, y, width, height] = painted;
        if ![x, y, width, height].iter().all(|value| value.is_finite())
            || width <= 0.0
            || height <= 0.0
            || columns == 0
            || rows == 0
        {
            return None;
        }
        Some(Self {
            origin: (x, y),
            cell: (width / f64::from(columns), height / f64::from(rows)),
            bounds: [0, 0, i64::from(columns), i64::from(rows)],
        })
    }

    /// A one-cell grid covering `area`, for operations without a source.
    #[must_use]
    pub fn for_area(area: [f64; 4]) -> Option<Self> {
        Self::for_pixels(area, 1, 1)
    }

    /// The cells `painted` covers, if it lies on this grid. With `pixels`,
    /// the rectangle must hold exactly that many cells across and down.
    fn cells(&self, painted: [f64; 4], pixels: Option<(u32, u32)>) -> Option<[i64; 4]> {
        let [x, y, width, height] = painted;
        let snap = |value: f64| {
            let rounded = value.round();
            ((value - rounded).abs() <= GRID_TOLERANCE && rounded.abs() < 1e12).then_some(rounded)
        };
        let column0 = snap((x - self.origin.0) / self.cell.0)?;
        let row0 = snap((y - self.origin.1) / self.cell.1)?;
        let column1 = snap((x + width - self.origin.0) / self.cell.0)?;
        let row1 = snap((y + height - self.origin.1) / self.cell.1)?;
        #[allow(
            clippy::cast_possible_truncation,
            reason = "snap() bounds the values far inside i64"
        )]
        let cells = [column0 as i64, row0 as i64, column1 as i64, row1 as i64];
        if cells[2] <= cells[0] || cells[3] <= cells[1] {
            return None;
        }
        if let Some((columns, rows)) = pixels
            && (cells[2] - cells[0] != i64::from(columns) || cells[3] - cells[1] != i64::from(rows))
        {
            return None;
        }
        Some(cells)
    }

    /// Extends the grid to `painted` when it lies on whole cells and the
    /// grid stays within `max_cells`.
    pub fn extend(
        &mut self,
        painted: [f64; 4],
        pixels: Option<(u32, u32)>,
        max_cells: usize,
    ) -> bool {
        let Some(cells) = self.cells(painted, pixels) else {
            return false;
        };
        let bounds = [
            self.bounds[0].min(cells[0]),
            self.bounds[1].min(cells[1]),
            self.bounds[2].max(cells[2]),
            self.bounds[3].max(cells[3]),
        ];
        let count = (bounds[2] - bounds[0]).checked_mul(bounds[3] - bounds[1]);
        if count.is_none_or(|count| usize::try_from(count).map_or(true, |count| count > max_cells))
        {
            return false;
        }
        self.bounds = bounds;
        true
    }

    fn columns(&self) -> usize {
        usize::try_from(self.bounds[2] - self.bounds[0]).unwrap_or(0)
    }

    fn rows(&self) -> usize {
        usize::try_from(self.bounds[3] - self.bounds[1]).unwrap_or(0)
    }

    /// The covered area `[x, y, width, height]`.
    #[must_use]
    pub fn area(&self) -> [f64; 4] {
        #[allow(
            clippy::cast_precision_loss,
            reason = "cell counts are bounded by the configured pixel limit"
        )]
        let [column0, row0, column1, row1] = self.bounds.map(|value| value as f64);
        [
            self.origin.0 + column0 * self.cell.0,
            self.origin.1 + row0 * self.cell.1,
            (column1 - column0) * self.cell.0,
            (row1 - row0) * self.cell.1,
        ]
    }
}

/// Destination pixels of raster operations composed on one [`BlitGrid`].
///
/// Every colour bit of a destination pixel goes through the operations as a
/// Boolean function of its original value: `low` holds the result for an
/// original 0 bit and `high` for an original 1 bit. A cell whose functions
/// are the identity keeps its destination pixel; a cell whose functions are
/// constant paints that colour. Anything else depends on the destination.
#[derive(Debug, Clone, PartialEq)]
pub struct RasterComposite {
    grid: BlitGrid,
    low: Vec<u32>,
    high: Vec<u32>,
}

/// Outcome of rendering a [`RasterComposite`].
#[derive(Debug, Clone, PartialEq)]
pub struct CompositeImage {
    pub image: SvgImage,
    /// Some cells depend on the destination; they were painted as if the
    /// destination were white.
    pub approximated: bool,
}

impl RasterComposite {
    /// Starts with every destination pixel of `grid` unchanged.
    pub fn new(grid: BlitGrid) -> Result<Self> {
        let count = grid
            .columns()
            .checked_mul(grid.rows())
            .ok_or_else(|| parse("raster composite size overflows"))?;
        let mut low = Vec::new();
        let mut high = Vec::new();
        for plane in [&mut low, &mut high] {
            plane
                .try_reserve_exact(count)
                .map_err(|source| Error::Allocation {
                    resource: "raster composite",
                    source,
                })?;
        }
        low.resize(count, 0);
        high.resize(count, COLOR_BITS);
        Ok(Self { grid, low, high })
    }

    /// Applies one ROP3 to the cells under `painted`.
    ///
    /// `source` holds the pixels painted there, or `None` for operations
    /// without a source; `pattern` is the solid brush `COLORREF`.
    pub fn apply(
        &mut self,
        rop3: Rop3,
        painted: [f64; 4],
        source: Option<&RgbaImage>,
        pattern: u32,
    ) -> Result<()> {
        let cells = self
            .grid
            .cells(painted, source.map(RgbaImage::dimensions))
            .filter(|cells| {
                cells[0] >= self.grid.bounds[0]
                    && cells[1] >= self.grid.bounds[1]
                    && cells[2] <= self.grid.bounds[2]
                    && cells[3] <= self.grid.bounds[3]
            })
            .ok_or_else(|| parse("raster operation lies outside its composite grid"))?;
        let columns = self.grid.columns();
        let index = |value: i64, origin: i64| usize::try_from(value - origin).unwrap_or(0);
        for row in cells[1]..cells[3] {
            for column in cells[0]..cells[2] {
                let cell =
                    index(row, self.grid.bounds[1]) * columns + index(column, self.grid.bounds[0]);
                let source = source.map_or(0, |image| {
                    rgba_colorref(
                        image
                            .get_pixel(
                                u32::try_from(column - cells[0]).unwrap_or(0),
                                u32::try_from(row - cells[1]).unwrap_or(0),
                            )
                            .0,
                    )
                });
                self.low[cell] = rop3.apply(pattern, source, self.low[cell]) & COLOR_BITS;
                self.high[cell] = rop3.apply(pattern, source, self.high[cell]) & COLOR_BITS;
            }
        }
        Ok(())
    }

    /// Emits the composed cells as one image over the grid area.
    ///
    /// Unchanged cells are transparent and constant cells opaque. Cells that
    /// still depend on the destination are painted as on a white
    /// destination, and [`CompositeImage::approximated`] reports them.
    pub fn to_svg_image(&self, maximum: usize) -> Result<CompositeImage> {
        let width = u32::try_from(self.grid.columns())
            .map_err(|_error| parse("raster composite width overflows"))?;
        let height = u32::try_from(self.grid.rows())
            .map_err(|_error| parse("raster composite height overflows"))?;
        let mut image = RgbaImage::new(width, height);
        let mut approximated = false;
        for ((pixel, low), high) in image.pixels_mut().zip(&self.low).zip(&self.high) {
            pixel.0 = if *low == 0 && *high == COLOR_BITS {
                [0, 0, 0, 0]
            } else if low == high {
                colorref_rgba(*low)
            } else {
                approximated = true;
                colorref_rgba(*high)
            };
        }
        let image = image_element(&image, ImagePlacement::Rect(self.grid.area()), 1.0, maximum)?;
        Ok(CompositeImage {
            image,
            approximated,
        })
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
        // MAKEROP4 keeps the foreground ROP3 in the low three bytes and the
        // background ROP3 table in the high byte.
        rop3: Rop3::from_code(rop4),
        background_rop3: Some(Rop3(rop4.to_le_bytes()[3])),
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
    let no_source = size == u32::from(function >> 8) + 3;
    // [MS-WMF] 2.3.1: records without a source bitmap carry a reserved
    // field between the source origin and the destination extent.
    let reserved = if no_source { 2 } else { 0 };
    let source_extent = if stretched { 4 } else { 0 };
    let common = 16 + source_extent + reserved;
    require_len(params, common, "WMF blit fixed fields")?;
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
    let origin = 4 + source_extent;
    let destination = origin + 4 + reserved;
    let (dh, dw, dy, dx) = (
        i16_at(params, destination)?,
        i16_at(params, destination + 2)?,
        i16_at(params, destination + 4)?,
        i16_at(params, destination + 6)?,
    );
    let (sh, sw) = if stretched {
        (i16_at(params, 4)?, i16_at(params, 6)?)
    } else {
        (dh, dw)
    };
    let (sy, sx) = (i16_at(params, origin)?, i16_at(params, origin + 2)?);
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
    // GDI paints nothing for an empty rectangle, so only the extents of
    // nonempty rectangles must be representable.
    if rect_extent(operation.destination).is_some() {
        normalized(operation.destination)?;
    }
    if operation.source.is_some() && rect_extent(operation.source_rect).is_some() {
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

/// Parses the DIB an EMF object record (such as EMR_CREATEDIBPATTERNBRUSHPT)
/// references through record-relative offsets. `payload` excludes the
/// eight-byte record header.
pub fn parse_emf_object_dib(
    payload: &[u8],
    offsets: [u32; 4],
    usage: ColorUsage,
    limits: DibLimits,
) -> Result<Dib<'_>> {
    let [off_bmi, cb_bmi, off_bits, cb_bits] = offsets;
    let bmi_len =
        usize::try_from(cb_bmi).map_err(|_error| parse("bitmap-info length is invalid"))?;
    let bits_len =
        usize::try_from(cb_bits).map_err(|_error| parse("bitmap-bits length is invalid"))?;
    if bmi_len == 0 || bits_len == 0 {
        return Err(parse("EMF bitmap header and bits must both be present"));
    }
    Dib::parse_parts(
        emf_range(payload, off_bmi, bmi_len)?,
        emf_range(payload, off_bits, bits_len)?,
        usage,
        limits,
    )
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

/// Returns the positive extents of a rectangle that covers at least one pixel.
fn rect_extent(rect: SignedRect) -> Option<(u32, u32)> {
    let width = rect.width.checked_abs()?;
    let height = rect.height.checked_abs()?;
    (width != 0 && height != 0).then_some((u32::try_from(width).ok()?, u32::try_from(height).ok()?))
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
    fn rop2_paint_resolves_modes_that_do_not_read_the_destination() {
        let mode = |value| Rop2::from_mode(value).unwrap();
        assert_eq!(mode(13).paint(0x0012_3456), Rop2Paint::Color(0x0012_3456));
        assert_eq!(mode(11).paint(0x0012_3456), Rop2Paint::Unchanged);
        assert_eq!(mode(1).paint(0x0012_3456), Rop2Paint::Color(0));
        assert_eq!(mode(16).paint(0x0012_3456), Rop2Paint::Color(0x00ff_ffff));
        assert_eq!(mode(4).paint(0x0012_3456), Rop2Paint::Color(0x00ed_cba9));
        // R2_XORPEN with black and R2_MASKPEN with white keep every bit.
        assert_eq!(mode(7).paint(0), Rop2Paint::Unchanged);
        assert_eq!(mode(9).paint(0x00ff_ffff), Rop2Paint::Unchanged);
        assert_eq!(mode(9).paint(0), Rop2Paint::Color(0));
        assert_eq!(mode(10).paint(0x00ff_ffff), Rop2Paint::Unchanged);
        assert_eq!(mode(7).paint(0x0012_3456), Rop2Paint::DestinationDependent);
        assert_eq!(mode(6).paint(0), Rop2Paint::DestinationDependent);
        // Palette flags in the high byte are not colour bits.
        assert_eq!(mode(7).paint(0x0200_0000), Rop2Paint::Unchanged);
        assert!(Rop2::from_mode(0).is_none() && Rop2::from_mode(17).is_none());
        assert!(mode(11).ignores_color() && mode(16).ignores_color());
        assert!(!mode(13).ignores_color());
    }

    #[test]
    fn rop3_word_evaluation_matches_the_truth_table() {
        for table in [0x88u32, 0xee, 0x66, 0xb8, 0xe2, 0xa0, 0x5a, 0x33] {
            let rop3 = Rop3::from_code(table << 16);
            let (pattern, source, destination) =
                (0x0055_aa0f_u32, 0x0033_cc3c_u32, 0x000f_f0a5_u32);
            let expected = u32::from_le_bytes([
                rop3.eval_byte(pattern as u8, source as u8, destination as u8),
                rop3.eval_byte(
                    (pattern >> 8) as u8,
                    (source >> 8) as u8,
                    (destination >> 8) as u8,
                ),
                rop3.eval_byte(
                    (pattern >> 16) as u8,
                    (source >> 16) as u8,
                    (destination >> 16) as u8,
                ),
                0,
            ]);
            assert_eq!(
                rop3.apply(pattern, source, destination) & COLOR_BITS,
                expected
            );
        }
        assert!(Rop3::from_code(0x0088_00c6).depends_on_destination());
        assert!(!Rop3::from_code(0x00f0_0021).depends_on_source());
        assert!(Rop3::from_code(0x00f0_0021).depends_on_pattern());
        assert!(!Rop3::from_code(SRCCOPY).depends_on_pattern());
    }

    #[test]
    fn empty_and_unchanged_operations_paint_nothing() {
        let d = dib();
        let source = Dib::parse(&d, DibLimits::default()).unwrap();
        let mut op = BitmapOp {
            kind: BitmapRecordKind::EmfBitBlt,
            source: Some(BitmapSource::Dib(source)),
            mask: None,
            source_rect: SignedRect {
                x: 0,
                y: 0,
                width: 0,
                height: 0,
            },
            destination: SignedRect {
                x: 0,
                y: 0,
                width: 10,
                height: 10,
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
        // An empty source rectangle selects no pixels.
        assert_eq!(op.paint().unwrap(), BitmapPaint::Image);
        assert!(op.to_svg_image().unwrap().is_none());

        op.source_rect = SignedRect {
            x: 0,
            y: 0,
            width: 1,
            height: 1,
        };
        op.destination.width = 0;
        assert_eq!(op.paint().unwrap(), BitmapPaint::Nothing);
        assert!(matches!(op.to_svg_image(), Err(Error::Unsupported(_))));

        op.destination.width = 10;
        op.rop3 = Rop3::from_code(0x00aa_0066);
        assert_eq!(op.paint().unwrap(), BitmapPaint::Nothing);

        for code in [
            0x0088_00c6,
            0x00ee_0086,
            0x00a0_00c9,
            0x0066_0046,
            0x000f_0001,
        ] {
            op.rop3 = Rop3::from_code(code);
            assert_eq!(op.paint().unwrap(), BitmapPaint::Composite);
        }
        op.rop3 = Rop3::from_code(NOTSRCCOPY);
        assert_eq!(op.paint().unwrap(), BitmapPaint::Image);

        op.rop3 = Rop3::from_code(0x00f0_0021);
        assert_eq!(op.paint().unwrap(), BitmapPaint::Brush);

        op.rop3 = Rop3::from_code(0x00ff_0062);
        assert_eq!(op.paint().unwrap(), BitmapPaint::Solid(0x00ff_ffff));

        op.rop3 = Rop3::from_code(0x0000_0042);
        assert_eq!(op.paint().unwrap(), BitmapPaint::Solid(0));
    }

    #[test]
    fn out_of_range_source_is_clipped_proportionally() {
        // A 2x2 bitmap stretched to a 10x10 destination with a source
        // rectangle that starts one pixel above and left of the bitmap.
        let mut d = vec![0; 40];
        d[0..4].copy_from_slice(&40u32.to_le_bytes());
        d[4..8].copy_from_slice(&2i32.to_le_bytes());
        d[8..12].copy_from_slice(&2i32.to_le_bytes());
        d[12..14].copy_from_slice(&1u16.to_le_bytes());
        d[14..16].copy_from_slice(&24u16.to_le_bytes());
        d.extend_from_slice(&[0; 16]);
        let source = Dib::parse(&d, DibLimits::default()).unwrap();
        let op = BitmapOp {
            kind: BitmapRecordKind::EmfStretchBlt,
            source: Some(BitmapSource::Dib(source)),
            mask: None,
            source_rect: SignedRect {
                x: -1,
                y: -1,
                width: 4,
                height: 4,
            },
            destination: SignedRect {
                x: 0,
                y: 0,
                width: 16,
                height: 16,
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
        let svg = op.to_svg_image().unwrap().unwrap();
        assert!(svg.element.contains(r#"x="4" y="4" width="8" height="8""#));

        // Mirroring maps the clipped source edge to the other destination
        // side: source columns -1..5 keep 0..2, which land in 3/6..5/6 of
        // the destination when mirrored.
        let mirrored = BitmapOp {
            source_rect: SignedRect {
                x: 5,
                y: 0,
                width: -6,
                height: 2,
            },
            destination: SignedRect {
                x: 0,
                y: 0,
                width: 12,
                height: 4,
            },
            ..op
        };
        let window = mirrored.blit_window().unwrap().unwrap();
        assert!(window.flip_x);
        assert_eq!(window.source.width, 2);
        assert_eq!(window.painted([0.0, 0.0, 12.0, 4.0]), [6.0, 0.0, 4.0, 4.0]);

        // A source rectangle entirely outside the bitmap paints nothing.
        let op = BitmapOp {
            source_rect: SignedRect {
                x: 8,
                y: 8,
                width: 2,
                height: 2,
            },
            ..op
        };
        assert!(op.to_svg_image().unwrap().is_none());
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
    fn wmf_blit_without_source_reads_past_the_reserved_field() {
        // DIBBitBlt: RasterOperation, YSrc, XSrc, Reserved, Height, Width,
        // YDest, XDest.
        let mut params = 0x00f0_0021u32.to_le_bytes().to_vec();
        for value in [0i16, 0, 0x7777, 10, 20, 30, 40] {
            params.extend_from_slice(&value.to_le_bytes());
        }
        let words = u32::try_from((params.len() + 6) / 2).unwrap();
        let op = BitmapOp::parse_wmf(
            words,
            0x0940,
            &params,
            StretchPolicy::default(),
            DibLimits::default(),
        )
        .unwrap();
        assert!(op.source.is_none());
        assert_eq!(
            op.destination,
            SignedRect {
                x: 40,
                y: 30,
                width: 20,
                height: 10,
            }
        );
        assert_eq!(op.paint().unwrap(), BitmapPaint::Brush);
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
        let svg = op.to_svg_image().unwrap().unwrap();
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

    /// A bottom-up 24-bpp DIB whose rows are listed from the top.
    fn rgb_dib(width: u32, rows: &[&[u32]]) -> Vec<u8> {
        let stride = (width as usize * 3).div_ceil(4) * 4;
        let mut d = vec![0; 40];
        d[0..4].copy_from_slice(&40u32.to_le_bytes());
        d[4..8].copy_from_slice(&(width as i32).to_le_bytes());
        d[8..12].copy_from_slice(&(rows.len() as i32).to_le_bytes());
        d[12..14].copy_from_slice(&1u16.to_le_bytes());
        d[14..16].copy_from_slice(&24u16.to_le_bytes());
        for row in rows.iter().rev() {
            let mut line = vec![0; stride];
            for (index, color) in row.iter().enumerate() {
                let [red, green, blue, _] = color.to_le_bytes();
                line[index * 3..index * 3 + 3].copy_from_slice(&[blue, green, red]);
            }
            d.extend_from_slice(&line);
        }
        d
    }

    fn blit<'a>(dib: &'a [u8], rop3: u32) -> BitmapOp<'a> {
        let source = Dib::parse(dib, DibLimits::default()).unwrap();
        BitmapOp {
            kind: BitmapRecordKind::WmfStretchDib,
            source: Some(BitmapSource::Dib(source)),
            mask: None,
            source_rect: SignedRect {
                x: 0,
                y: 0,
                width: i32::try_from(source.info().width).unwrap(),
                height: i32::try_from(source.info().height).unwrap(),
            },
            destination: SignedRect {
                x: 10,
                y: 20,
                width: 4,
                height: 2,
            },
            parallelogram: None,
            source_transform: AffineTransform::IDENTITY,
            rop3: Rop3::from_code(rop3),
            background_rop3: None,
            scan_range: None,
            alpha: None,
            transparent_color: None,
            stretch: StretchPolicy::default(),
            limits: DibLimits::default(),
        }
    }

    fn compose(operations: &[BitmapOp<'_>], pattern: u32) -> (RgbaImage, bool) {
        let destination = [10.0, 20.0, 4.0, 2.0];
        let first = operations[0].blit_window().unwrap().unwrap();
        let (columns, rows) = first.pixels();
        let mut grid = BlitGrid::for_pixels(first.painted(destination), columns, rows).unwrap();
        for operation in operations {
            let window = operation.blit_window().unwrap().unwrap();
            assert!(grid.extend(window.painted(destination), Some(window.pixels()), 64));
        }
        let mut composite = RasterComposite::new(grid).unwrap();
        for operation in operations {
            let pixels = operation.blit_pixels().unwrap().unwrap();
            composite
                .apply(
                    operation.rop3,
                    pixels.window.painted(destination),
                    Some(&pixels.image),
                    pattern,
                )
                .unwrap();
        }
        let image = composite.to_svg_image(1 << 20).unwrap();
        assert!(
            image
                .image
                .element
                .contains(r#"x="10" y="20" width="4" height="2""#)
        );
        let png = image
            .image
            .element
            .split("base64,")
            .nth(1)
            .unwrap()
            .trim_end_matches("\"/>");
        let decoded = image::load_from_memory(&STANDARD.decode(png).unwrap())
            .unwrap()
            .to_rgba8();
        (decoded, image.approximated)
    }

    #[test]
    fn srcand_mask_and_srcpaint_image_compose_one_transparent_image() {
        const WHITE: u32 = 0x00ff_ffff;
        const RED: u32 = 0x0000_00ff;
        // The mask keeps the destination where it is white; the image is
        // black there.
        let mask = rgb_dib(2, &[&[0, WHITE], &[WHITE, 0]]);
        let image = rgb_dib(2, &[&[RED, 0], &[0, RED]]);
        let (pixels, approximated) =
            compose(&[blit(&mask, 0x0088_00c6), blit(&image, 0x00ee_0086)], 0);
        assert!(!approximated);
        assert_eq!(pixels.get_pixel(0, 0).0, [255, 0, 0, 255]);
        assert_eq!(pixels.get_pixel(1, 0).0[3], 0);
        assert_eq!(pixels.get_pixel(0, 1).0[3], 0);
        assert_eq!(pixels.get_pixel(1, 1).0, [255, 0, 0, 255]);

        // SRCINVERT, SRCAND mask, SRCINVERT keeps the destination wherever
        // the mask is white and paints the image elsewhere.
        let (pixels, approximated) = compose(
            &[
                blit(&image, 0x0066_0046),
                blit(&mask, 0x0088_00c6),
                blit(&image, 0x0066_0046),
            ],
            0,
        );
        assert!(!approximated);
        assert_eq!(pixels.get_pixel(0, 0).0, [255, 0, 0, 255]);
        assert_eq!(pixels.get_pixel(1, 0).0[3], 0);
    }

    #[test]
    fn destination_dependent_pixels_are_reported_as_approximated() {
        // SRCAND with a coloured source mixes source and destination bits.
        let image = rgb_dib(1, &[&[0x0000_00ff]]);
        let (pixels, approximated) = compose(&[blit(&image, 0x0088_00c6)], 0);
        assert!(approximated);
        // A white destination ANDed with red stays red.
        assert_eq!(pixels.get_pixel(0, 0).0, [255, 0, 0, 255]);
        // A lone SRCAND of black and white is exact.
        let mask = rgb_dib(2, &[&[0, 0x00ff_ffff]]);
        let (pixels, approximated) = compose(&[blit(&mask, 0x0088_00c6)], 0);
        assert!(!approximated);
        assert_eq!(pixels.get_pixel(0, 0).0, [0, 0, 0, 255]);
        assert_eq!(pixels.get_pixel(1, 0).0[3], 0);
    }

    #[test]
    fn blit_grid_accepts_only_rectangles_on_whole_cells() {
        let mut grid = BlitGrid::for_pixels([0.0, 0.0, 8.0, 4.0], 4, 2).unwrap();
        assert!(grid.extend([2.0, 2.0, 4.0, 2.0], Some((2, 1)), 64));
        assert_eq!(grid.area(), [0.0, 0.0, 8.0, 4.0]);
        // Another cell size or a fractional offset is a different grid.
        assert!(!grid.extend([0.0, 0.0, 8.0, 4.0], Some((8, 2)), 64));
        assert!(!grid.extend([1.0, 0.0, 4.0, 2.0], Some((2, 1)), 64));
        // Growing past the cell budget is refused.
        assert!(!grid.extend([8.0, 0.0, 200.0, 4.0], Some((100, 2)), 64));
        assert!(grid.extend([8.0, 0.0, 2.0, 2.0], None, 64));
        assert_eq!(grid.area(), [0.0, 0.0, 10.0, 4.0]);
    }

    #[test]
    fn monochrome_pattern_paints_zero_bits_with_the_text_colour() {
        // Two rows of 8 bits, listed from the top.
        let pattern =
            BrushPattern::from_monochrome_bits(8, 2, 2, &[0x80, 0, 0x01, 0], true).unwrap();
        assert!(pattern.is_monochrome());
        let tile = pattern.tile(0x0000_00ff, 0x00ff_0000);
        assert_eq!(tile.get_pixel(0, 0).0, [0, 0, 255, 255]);
        assert_eq!(tile.get_pixel(1, 0).0, [255, 0, 0, 255]);
        assert_eq!(tile.get_pixel(7, 1).0, [0, 0, 255, 255]);
        let png = pattern.tile_png(0, 0x00ff_ffff, 1 << 16).unwrap();
        let definition =
            BrushPattern::svg_definition("brush", &png, (8.0, 2.0), (1.0, 0.0)).unwrap();
        assert!(definition.starts_with(
            r#"<pattern id="brush" patternUnits="userSpaceOnUse" x="1" y="0" width="8" height="2">"#
        ));
        assert!(BrushPattern::from_monochrome_bits(8, 2, 1, &[0], true).is_err());
    }
}
