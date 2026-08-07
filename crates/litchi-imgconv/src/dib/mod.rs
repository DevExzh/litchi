//! Safe, bounded parsing and decoding of Windows device-independent bitmaps.
//!
//! Both packed DIBs (as commonly found in WMF) and the separate bitmap-info /
//! bitmap-bits representation used by EMF records are supported.  Parsing is
//! deliberately strict: only the CORE, INFO, V4, and V5 headers described by
//! MS-WMF are accepted.

#![allow(
    clippy::missing_errors_doc,
    reason = "all public fallible APIs report malformed input, resource limits, or codec failure"
)]

use std::io::{Cursor, Seek, SeekFrom, Write};

use image::{DynamicImage, ImageFormat, ImageReader};
use litchi_core::error::{Error, Result};

const BITMAP_CORE_HEADER: u32 = 12;
const BITMAP_INFO_HEADER: u32 = 40;
const BITMAP_V4_HEADER: u32 = 108;
const BITMAP_V5_HEADER: u32 = 124;

/// Explicit resource ceilings applied before allocation or codec dispatch.
#[allow(
    clippy::module_name_repetitions,
    reason = "DibLimits is clearer at public call sites than the ambiguous name Limits"
)]
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DibLimits {
    /// Maximum combined bitmap-info and bitmap-bits input length.
    pub max_input_bytes: usize,
    /// Maximum logical image width.
    pub max_width: u32,
    /// Maximum logical image height.
    pub max_height: u32,
    /// Maximum logical pixel count.
    pub max_pixels: u64,
    /// Maximum number of color-table entries.
    pub max_palette_entries: u32,
    /// Maximum decoded image-buffer length.
    pub max_decoded_bytes: usize,
    /// Maximum BMP or PNG adapter output length.
    pub max_output_bytes: usize,
}

impl Default for DibLimits {
    fn default() -> Self {
        Self {
            max_input_bytes: 256 * 1024 * 1024,
            max_width: 8192,
            max_height: 8192,
            max_pixels: 32 * 1024 * 1024,
            max_palette_entries: 4096,
            max_decoded_bytes: 256 * 1024 * 1024,
            max_output_bytes: 256 * 1024 * 1024,
        }
    }
}

/// The exact DIB header variant.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum HeaderKind {
    Core,
    Info,
    V4,
    V5,
}

/// The values accepted in a `BitmapInfoHeader.Compression` field.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(u32)]
pub enum Compression {
    Rgb = 0,
    Rle8 = 1,
    Rle4 = 2,
    Bitfields = 3,
    Jpeg = 4,
    Png = 5,
    Cmyk = 0x0b,
    CmykRle8 = 0x0c,
    CmykRle4 = 0x0d,
}

impl Compression {
    fn parse(value: u32) -> Result<Self> {
        match value {
            0 => Ok(Self::Rgb),
            1 => Ok(Self::Rle8),
            2 => Ok(Self::Rle4),
            3 => Ok(Self::Bitfields),
            4 => Ok(Self::Jpeg),
            5 => Ok(Self::Png),
            0x0b => Ok(Self::Cmyk),
            0x0c => Ok(Self::CmykRle8),
            0x0d => Ok(Self::CmykRle4),
            _ => Err(parse(format!("unsupported DIB compression value {value}"))),
        }
    }

    const fn is_embedded(self) -> bool {
        matches!(self, Self::Jpeg | Self::Png)
    }

    const fn is_scanline_based(self) -> bool {
        matches!(self, Self::Rgb | Self::Bitfields | Self::Cmyk)
    }

    const fn is_cmyk(self) -> bool {
        matches!(self, Self::Cmyk | Self::CmykRle8 | Self::CmykRle4)
    }
}

/// How the otherwise-reserved fourth byte of a 32-bpp `BI_RGB` pixel is handled.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum AlphaInterpretation {
    /// Ignore byte 3, as required for ordinary `BI_RGB` rendering.
    #[default]
    Ignore,
    /// Treat byte 3 as straight alpha.
    Straight,
    /// Treat byte 3 as alpha and un-premultiply RGB, suitable for `AC_SRC_ALPHA` input.
    Premultiplied,
}

/// Interpretation of the color table requested by the containing EMF/WMF record.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum ColorUsage {
    /// Entries are RGB triples/quads (`DIB_RGB_COLORS`).
    #[default]
    Rgb,
    /// Color-table entries are 16-bit logical-palette indexes (`DIB_PAL_COLORS`).
    PaletteEntries,
    /// No table is present; pixels index the logical palette (`DIB_PAL_INDICES`).
    PaletteIndices,
}

impl ColorUsage {
    /// Converts an MS-EMF/MS-WMF `DIBColors`/`ColorUsage` numeric value.
    pub fn from_raw(value: u32) -> Result<Self> {
        match value {
            0 => Ok(Self::Rgb),
            1 => Ok(Self::PaletteEntries),
            2 => Ok(Self::PaletteIndices),
            _ => Err(parse(format!("unsupported DIB color-usage value {value}"))),
        }
    }
}

/// Validated, allocation-free DIB metadata.
#[allow(
    clippy::module_name_repetitions,
    reason = "DibInfo distinguishes validated DIB metadata from container record metadata"
)]
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DibInfo {
    pub header: HeaderKind,
    pub width: u32,
    pub height: u32,
    pub top_down: bool,
    pub bit_count: u16,
    pub compression: Compression,
    pub palette_entries: u32,
    /// Bytes per DWORD-aligned row for uncompressed scanline formats.
    pub stride: Option<usize>,
    /// Number of bytes belonging to the bitmap bits or embedded image.
    pub pixel_data_len: usize,
    /// Red, green, blue, and alpha masks. Alpha is zero when absent.
    pub masks: [u32; 4],
}

/// A parsed DIB borrowing its bitmap-info and bitmap-bits buffers.
#[derive(Debug, Clone, Copy)]
pub struct Dib<'a> {
    info: DibInfo,
    bitmap_info: &'a [u8],
    bitmap_bits: &'a [u8],
    usage: ColorUsage,
    limits: DibLimits,
}

/// A complete BMP byte stream assembled from an already validated DIB.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ValidatedBmp {
    bytes: Vec<u8>,
    pixel_offset: u32,
}

impl ValidatedBmp {
    #[must_use]
    pub fn as_bytes(&self) -> &[u8] {
        &self.bytes
    }

    #[must_use]
    pub fn into_bytes(self) -> Vec<u8> {
        self.bytes
    }

    #[must_use]
    pub const fn pixel_offset(&self) -> u32 {
        self.pixel_offset
    }
}

impl AsRef<[u8]> for ValidatedBmp {
    fn as_ref(&self) -> &[u8] {
        self.as_bytes()
    }
}

impl<'a> Dib<'a> {
    /// Parses a packed DIB whose pixels immediately follow its header, masks,
    /// color table, and (when it occurs there) V5 profile.
    pub fn parse(data: &'a [u8], limits: DibLimits) -> Result<Self> {
        Self::parse_with_usage(data, ColorUsage::Rgb, limits)
    }

    /// Parses a packed DIB with an explicit color-table interpretation.
    pub fn parse_with_usage(data: &'a [u8], usage: ColorUsage, limits: DibLimits) -> Result<Self> {
        check_input_len(data.len(), limits)?;
        let mut layout = parse_layout(data, usage, limits)?;
        let required = required_pixel_len(&layout)?;

        // A V5 profile can precede the pixels in a packed DIB.  A profile at a
        // later offset is permitted only after the pixel payload, never inside it.
        if let Some((start, end)) = layout.profile_range {
            if end > data.len() {
                return Err(parse("V5 color profile extends past packed DIB"));
            }
            if start == layout.bitmap_info_len {
                layout.bitmap_info_len = align4(end)?;
            } else {
                let pixel_end = layout
                    .bitmap_info_len
                    .checked_add(required)
                    .ok_or_else(|| parse("DIB pixel extent overflows"))?;
                if start < pixel_end && end > layout.bitmap_info_len {
                    return Err(parse("V5 color profile overlaps bitmap pixels"));
                }
            }
        }

        let end = layout
            .bitmap_info_len
            .checked_add(required)
            .ok_or_else(|| parse("DIB pixel extent overflows"))?;
        let bits = data
            .get(layout.bitmap_info_len..end)
            .ok_or_else(|| parse("DIB bitmap bits are truncated"))?;
        let bitmap_info = data
            .get(..layout.bitmap_info_len)
            .ok_or_else(|| parse("DIB bitmap-info extent is invalid"))?;
        Ok(Self::from_layout(layout, bitmap_info, bits, usage, limits))
    }

    /// Parses the split bitmap-info and bitmap-bits representation used by EMF.
    pub fn parse_parts(
        bitmap_info: &'a [u8],
        bitmap_bits: &'a [u8],
        usage: ColorUsage,
        limits: DibLimits,
    ) -> Result<Self> {
        let input_len = bitmap_info
            .len()
            .checked_add(bitmap_bits.len())
            .ok_or_else(|| parse("combined DIB input length overflows"))?;
        check_input_len(input_len, limits)?;
        let layout = parse_layout(bitmap_info, usage, limits)?;
        if bitmap_info.len() < layout.bitmap_info_len {
            return Err(parse("DIB bitmap info or color table is truncated"));
        }
        #[allow(
            clippy::collapsible_if,
            reason = "nested form remains compatible with direct rustfmt file invocation"
        )]
        if let Some((_, end)) = layout.profile_range {
            if end > bitmap_info.len() {
                return Err(parse("V5 color profile extends past bitmap info"));
            }
        }
        let required = required_pixel_len(&layout)?;
        let bits = bitmap_bits
            .get(..required)
            .ok_or_else(|| parse("DIB bitmap bits are truncated"))?;
        Ok(Self::from_layout(layout, bitmap_info, bits, usage, limits))
    }

    fn from_layout(
        layout: Layout,
        bitmap_info: &'a [u8],
        bitmap_bits: &'a [u8],
        usage: ColorUsage,
        limits: DibLimits,
    ) -> Self {
        Self {
            info: DibInfo {
                header: layout.header,
                width: layout.width,
                height: layout.height,
                top_down: layout.top_down,
                bit_count: layout.bit_count,
                compression: layout.compression,
                palette_entries: layout.palette_entries,
                stride: layout.stride,
                pixel_data_len: bitmap_bits.len(),
                masks: layout.masks,
            },
            bitmap_info,
            bitmap_bits,
            usage,
            limits,
        }
    }

    #[must_use]
    pub const fn info(&self) -> &DibInfo {
        &self.info
    }

    #[must_use]
    pub const fn bitmap_info(&self) -> &'a [u8] {
        self.bitmap_info
    }

    #[must_use]
    pub const fn bitmap_bits(&self) -> &'a [u8] {
        self.bitmap_bits
    }

    /// Produces a BMP file wrapper for non-embedded DIB encodings.
    pub fn to_bmp(&self) -> Result<ValidatedBmp> {
        if self.usage != ColorUsage::Rgb {
            return Err(unsupported(
                "logical-palette DIB usage requires the containing metafile's palette",
            ));
        }
        if self.info.compression.is_embedded() {
            return Err(unsupported(
                "embedded JPEG/PNG DIBs do not have a portable BMP wrapper",
            ));
        }

        let pixel_offset = 14usize
            .checked_add(self.bitmap_info.len())
            .ok_or_else(|| parse("BMP pixel offset overflows"))?;
        let file_len = pixel_offset
            .checked_add(self.bitmap_bits.len())
            .ok_or_else(|| parse("BMP file length overflows"))?;
        if file_len > self.limits.max_output_bytes {
            return Err(limit("BMP output", self.limits.max_output_bytes));
        }
        let file_len_u32 =
            u32::try_from(file_len).map_err(|_error| parse("BMP file length exceeds u32"))?;
        let pixel_offset_u32 =
            u32::try_from(pixel_offset).map_err(|_error| parse("BMP pixel offset exceeds u32"))?;

        let mut bytes = Vec::new();
        bytes
            .try_reserve_exact(file_len)
            .map_err(|source| Error::Allocation {
                resource: "validated BMP",
                source,
            })?;
        bytes.extend_from_slice(b"BM");
        bytes.extend_from_slice(&file_len_u32.to_le_bytes());
        bytes.extend_from_slice(&0u32.to_le_bytes());
        bytes.extend_from_slice(&pixel_offset_u32.to_le_bytes());
        bytes.extend_from_slice(self.bitmap_info);
        bytes.extend_from_slice(self.bitmap_bits);
        Ok(ValidatedBmp {
            bytes,
            pixel_offset: pixel_offset_u32,
        })
    }

    /// Decodes the DIB, dispatching embedded JPEG/PNG payloads directly.
    pub fn to_dynamic_image(&self) -> Result<DynamicImage> {
        self.to_dynamic_image_with_alpha(AlphaInterpretation::Ignore)
    }

    /// Decodes with an explicit policy for the reserved byte in 32-bpp `BI_RGB`.
    /// Bitfield alpha masks remain authoritative and do not use this policy.
    pub fn to_dynamic_image_with_alpha(&self, alpha: AlphaInterpretation) -> Result<DynamicImage> {
        if self.usage != ColorUsage::Rgb {
            return Err(unsupported(
                "logical-palette DIB usage requires the containing metafile's palette",
            ));
        }
        if self.info.compression.is_cmyk() {
            return Err(unsupported(format!(
                "recognized {:?} DIB compression is not decodable",
                self.info.compression
            )));
        }
        let (source, format) = match self.info.compression {
            Compression::Jpeg => (self.bitmap_bits, ImageFormat::Jpeg),
            Compression::Png => (self.bitmap_bits, ImageFormat::Png),
            Compression::Rgb
            | Compression::Rle8
            | Compression::Rle4
            | Compression::Bitfields
            | Compression::Cmyk
            | Compression::CmykRle8
            | Compression::CmykRle4 => {
                let bmp = self.to_bmp()?;
                let image = decode_image(bmp.as_bytes(), ImageFormat::Bmp, self.info, self.limits)?;
                return self.apply_rgb32_alpha(image, alpha);
            },
        };
        decode_image(source, format, self.info, self.limits)
    }

    fn apply_rgb32_alpha(
        &self,
        image: DynamicImage,
        alpha: AlphaInterpretation,
    ) -> Result<DynamicImage> {
        if alpha == AlphaInterpretation::Ignore
            || self.info.compression != Compression::Rgb
            || self.info.bit_count != 32
        {
            return Ok(image);
        }
        let stride = self
            .info
            .stride
            .ok_or_else(|| parse("32-bpp BI_RGB stride is unavailable"))?;
        let rgba_len = usize_from_u32(self.info.width)?
            .checked_mul(usize_from_u32(self.info.height)?)
            .and_then(|value| value.checked_mul(4))
            .ok_or_else(|| parse("DIB RGBA byte length overflows"))?;
        if rgba_len > self.limits.max_decoded_bytes {
            return Err(limit("decoded DIB", self.limits.max_decoded_bytes));
        }
        let mut rgba = image.to_rgba8();
        for y in 0..self.info.height {
            let source_y = if self.info.top_down {
                y
            } else {
                self.info.height - 1 - y
            };
            let row = usize_from_u32(source_y)?
                .checked_mul(stride)
                .ok_or_else(|| parse("DIB alpha row offset overflows"))?;
            for x in 0..self.info.width {
                let pixel_offset = usize_from_u32(x)?
                    .checked_mul(4)
                    .ok_or_else(|| parse("DIB alpha pixel offset overflows"))?;
                let offset = row
                    .checked_add(pixel_offset)
                    .and_then(|value| value.checked_add(3))
                    .ok_or_else(|| parse("DIB alpha byte offset overflows"))?;
                let value = *self
                    .bitmap_bits
                    .get(offset)
                    .ok_or_else(|| parse("DIB alpha byte is truncated"))?;
                let pixel = rgba.get_pixel_mut(x, y);
                if alpha == AlphaInterpretation::Premultiplied {
                    if value == 0 {
                        pixel.0[..3].fill(0);
                    } else {
                        for component in &mut pixel.0[..3] {
                            let scaled = (u32::from(*component) * 255 / u32::from(value)).min(255);
                            *component = u8::try_from(scaled)
                                .map_err(|_error| parse("un-premultiplied channel exceeds u8"))?;
                        }
                    }
                }
                pixel.0[3] = value;
            }
        }
        Ok(DynamicImage::ImageRgba8(rgba))
    }

    /// Decodes and encodes the image as PNG under the configured output limit.
    pub fn to_png(&self) -> Result<Vec<u8>> {
        let image = self.to_dynamic_image()?;
        let mut output = BoundedCursor::new(self.limits.max_output_bytes)?;
        image
            .write_to(&mut output, ImageFormat::Png)
            .map_err(|error| parse(format!("failed to encode DIB as PNG: {error}")))?;
        Ok(output.into_inner())
    }
}

#[derive(Debug, Clone, Copy)]
struct Layout {
    header: HeaderKind,
    width: u32,
    height: u32,
    top_down: bool,
    bit_count: u16,
    compression: Compression,
    image_size: u32,
    palette_entries: u32,
    stride: Option<usize>,
    masks: [u32; 4],
    bitmap_info_len: usize,
    profile_range: Option<(usize, usize)>,
}

fn parse_layout(data: &[u8], usage: ColorUsage, limits: DibLimits) -> Result<Layout> {
    let header_size = read_u32(data, 0, "DIB header size")?;
    let header = match header_size {
        BITMAP_CORE_HEADER => HeaderKind::Core,
        BITMAP_INFO_HEADER => HeaderKind::Info,
        BITMAP_V4_HEADER => HeaderKind::V4,
        BITMAP_V5_HEADER => HeaderKind::V5,
        _ => {
            return Err(parse(format!(
                "unsupported DIB header size {header_size}; expected 12, 40, 108, or 124"
            )));
        },
    };
    let header_len = usize::try_from(header_size)
        .map_err(|_error| parse("DIB header size does not fit this platform"))?;
    if data.len() < header_len {
        return Err(parse("DIB header is truncated"));
    }

    let (
        width,
        height,
        top_down,
        planes,
        bit_count,
        compression,
        image_size,
        colors_used,
        colors_important,
    ) = if header == HeaderKind::Core {
        (
            u32::from(read_u16(data, 4, "DIB width")?),
            u32::from(read_u16(data, 6, "DIB height")?),
            false,
            read_u16(data, 8, "DIB planes")?,
            read_u16(data, 10, "DIB bit count")?,
            Compression::Rgb,
            0,
            0,
            0,
        )
    } else {
        let signed_width = read_i32(data, 4, "DIB width")?;
        if signed_width <= 0 {
            return Err(parse("DIB width must be positive"));
        }
        let signed_height = read_i32(data, 8, "DIB height")?;
        if signed_height == 0 {
            return Err(parse("DIB height must not be zero"));
        }
        let height = signed_height
            .checked_abs()
            .ok_or_else(|| parse("DIB height magnitude overflows i32"))?;
        (
            u32::try_from(signed_width).map_err(|_error| parse("DIB width is invalid"))?,
            u32::try_from(height).map_err(|_error| parse("DIB height is invalid"))?,
            signed_height < 0,
            read_u16(data, 12, "DIB planes")?,
            read_u16(data, 14, "DIB bit count")?,
            Compression::parse(read_u32(data, 16, "DIB compression")?)?,
            read_u32(data, 20, "DIB image size")?,
            read_u32(data, 32, "DIB colors used")?,
            read_u32(data, 36, "DIB important colors")?,
        )
    };

    if width == 0 || height == 0 {
        return Err(parse("DIB dimensions must be nonzero"));
    }
    if planes != 1 {
        return Err(parse(format!("DIB planes must be 1, found {planes}")));
    }
    check_dimensions(width, height, limits)?;
    validate_format(header, bit_count, compression, top_down)?;

    let maximum_colors = if bit_count <= 8 && bit_count != 0 {
        1u32.checked_shl(u32::from(bit_count))
            .ok_or_else(|| parse("DIB palette count overflows"))?
    } else {
        0
    };
    let palette_entries = if usage == ColorUsage::PaletteIndices
        || compression.is_embedded()
        || compression.is_cmyk()
    {
        0
    } else if colors_used == 0 {
        maximum_colors
    } else if maximum_colors != 0 {
        colors_used.min(maximum_colors)
    } else {
        colors_used
    };
    if (compression.is_embedded() || compression.is_cmyk()) && colors_used != 0 {
        return Err(parse("embedded or CMYK DIB must not have a color table"));
    }
    if palette_entries > limits.max_palette_entries {
        return Err(limit(
            "DIB color-table entries",
            usize_from_u32(limits.max_palette_entries)?,
        ));
    }
    if colors_important != 0 && colors_important > palette_entries {
        return Err(parse(
            "DIB important-color count exceeds the color-table size",
        ));
    }

    let external_mask_count =
        usize::from(header == HeaderKind::Info && compression == Compression::Bitfields) * 3;
    let external_mask_len = external_mask_count
        .checked_mul(4)
        .ok_or_else(|| parse("DIB mask length overflows"))?;
    let palette_entry_len = match usage {
        ColorUsage::Rgb if header == HeaderKind::Core => 3,
        ColorUsage::Rgb => 4,
        ColorUsage::PaletteEntries => 2,
        ColorUsage::PaletteIndices => 0,
    };
    let palette_len = usize_from_u32(palette_entries)?
        .checked_mul(palette_entry_len)
        .ok_or_else(|| parse("DIB color-table length overflows"))?;
    let bitmap_info_len = header_len
        .checked_add(external_mask_len)
        .and_then(|value| value.checked_add(palette_len))
        .ok_or_else(|| parse("DIB bitmap-info length overflows"))?;

    let masks = if compression == Compression::Bitfields {
        let mask_offset = if header == HeaderKind::Info {
            header_len
        } else {
            40
        };
        let result = [
            read_u32(data, mask_offset, "DIB red mask")?,
            read_u32(data, mask_offset + 4, "DIB green mask")?,
            read_u32(data, mask_offset + 8, "DIB blue mask")?,
            if header == HeaderKind::V4 || header == HeaderKind::V5 {
                read_u32(data, mask_offset + 12, "DIB alpha mask")?
            } else {
                0
            },
        ];
        validate_masks(result, bit_count)?;
        result
    } else {
        [0; 4]
    };

    let stride = if compression.is_scanline_based() {
        let row_bits = usize_from_u32(width)?
            .checked_mul(usize::from(bit_count))
            .ok_or_else(|| parse("DIB row bit count overflows"))?;
        Some(
            row_bits
                .checked_add(31)
                .ok_or_else(|| parse("DIB stride rounding overflows"))?
                / 32
                * 4,
        )
    } else {
        None
    };

    let profile_range = if header == HeaderKind::V5 {
        let offset = read_u32(data, 112, "DIB V5 profile offset")?;
        let size = read_u32(data, 116, "DIB V5 profile size")?;
        match (offset, size) {
            (0, 0) => None,
            (0, _) | (_, 0) => {
                return Err(parse(
                    "DIB V5 profile offset and size must both be zero or nonzero",
                ));
            },
            _ => {
                let start = usize_from_u32(offset)?;
                let end = start
                    .checked_add(usize_from_u32(size)?)
                    .ok_or_else(|| parse("DIB V5 profile extent overflows"))?;
                if start < bitmap_info_len {
                    return Err(parse(
                        "DIB V5 profile overlaps header, masks, or color table",
                    ));
                }
                Some((start, end))
            },
        }
    } else {
        None
    };

    Ok(Layout {
        header,
        width,
        height,
        top_down,
        bit_count,
        compression,
        image_size,
        palette_entries,
        stride,
        masks,
        bitmap_info_len,
        profile_range,
    })
}

fn validate_format(
    header: HeaderKind,
    bit_count: u16,
    compression: Compression,
    top_down: bool,
) -> Result<()> {
    let valid = match compression {
        Compression::Rgb => matches!(bit_count, 1 | 4 | 8 | 16 | 24 | 32),
        Compression::Rle8 | Compression::CmykRle8 => bit_count == 8,
        Compression::Rle4 | Compression::CmykRle4 => bit_count == 4,
        Compression::Bitfields => matches!(bit_count, 16 | 32),
        Compression::Jpeg | Compression::Png => bit_count == 0,
        Compression::Cmyk => bit_count == 32,
    };
    if !valid {
        return Err(parse(format!(
            "DIB bit count {bit_count} is invalid for {compression:?} compression"
        )));
    }
    if header == HeaderKind::Core && compression != Compression::Rgb {
        return Err(parse("BITMAPCOREHEADER supports only BI_RGB"));
    }
    if top_down && compression != Compression::Rgb && compression != Compression::Bitfields {
        return Err(parse("top-down DIBs do not support this compression"));
    }
    Ok(())
}

fn validate_masks(masks: [u32; 4], bit_count: u16) -> Result<()> {
    if masks[..3].contains(&0) {
        return Err(parse("DIB bitfield RGB masks must be nonzero"));
    }
    let valid_bits = if bit_count == 32 {
        u32::MAX
    } else {
        (1u32 << bit_count) - 1
    };
    let mut used = 0u32;
    for mask in masks.into_iter().filter(|mask| *mask != 0) {
        if mask & !valid_bits != 0 {
            return Err(parse("DIB bitfield mask exceeds the pixel bit count"));
        }
        let shifted = mask >> mask.trailing_zeros();
        if shifted & shifted.wrapping_add(1) != 0 {
            return Err(parse("DIB bitfield masks must contain contiguous bits"));
        }
        if used & mask != 0 {
            return Err(parse("DIB bitfield masks must not overlap"));
        }
        used |= mask;
    }
    Ok(())
}

fn required_pixel_len(layout: &Layout) -> Result<usize> {
    if let Some(stride) = layout.stride {
        let expected = stride
            .checked_mul(usize_from_u32(layout.height)?)
            .ok_or_else(|| parse("DIB pixel byte length overflows"))?;
        if layout.image_size != 0 && usize_from_u32(layout.image_size)? < expected {
            return Err(parse("DIB image-size field is smaller than its scanlines"));
        }
        Ok(expected)
    } else {
        if layout.image_size == 0 {
            return Err(parse("compressed DIB image-size field must be nonzero"));
        }
        usize_from_u32(layout.image_size)
    }
}

fn decode_image(
    source: &[u8],
    format: ImageFormat,
    expected: DibInfo,
    limits: DibLimits,
) -> Result<DynamicImage> {
    let max_alloc = u64::try_from(limits.max_decoded_bytes)
        .map_err(|_error| parse("DIB decoded-byte limit does not fit u64"))?;
    let mut reader = ImageReader::with_format(Cursor::new(source), format);
    let mut image_limits = image::Limits::default();
    image_limits.max_image_width = Some(limits.max_width);
    image_limits.max_image_height = Some(limits.max_height);
    image_limits.max_alloc = Some(max_alloc);
    reader.limits(image_limits);
    let image = reader
        .decode()
        .map_err(|error| parse(format!("failed to decode DIB: {error}")))?;
    if image.width() != expected.width || image.height() != expected.height {
        return Err(parse(format!(
            "decoded DIB dimensions {}x{} do not match header {}x{}",
            image.width(),
            image.height(),
            expected.width,
            expected.height
        )));
    }
    check_dimensions(image.width(), image.height(), limits)?;
    if image.as_bytes().len() > limits.max_decoded_bytes {
        return Err(limit("decoded DIB", limits.max_decoded_bytes));
    }
    Ok(image)
}

fn check_dimensions(width: u32, height: u32, limits: DibLimits) -> Result<()> {
    if width > limits.max_width {
        return Err(limit("DIB width", usize_from_u32(limits.max_width)?));
    }
    if height > limits.max_height {
        return Err(limit("DIB height", usize_from_u32(limits.max_height)?));
    }
    let pixels = u64::from(width)
        .checked_mul(u64::from(height))
        .ok_or_else(|| parse("DIB pixel count overflows"))?;
    if pixels > limits.max_pixels {
        return Err(limit_u64("DIB pixel count", limits.max_pixels));
    }
    Ok(())
}

fn check_input_len(actual: usize, limits: DibLimits) -> Result<()> {
    if actual > limits.max_input_bytes {
        return Err(limit("DIB input", limits.max_input_bytes));
    }
    Ok(())
}

fn align4(value: usize) -> Result<usize> {
    value
        .checked_add(3)
        .map(|rounded| rounded & !3)
        .ok_or_else(|| parse("DIB DWORD alignment overflows"))
}

fn usize_from_u32(value: u32) -> Result<usize> {
    usize::try_from(value).map_err(|_error| parse("DIB size does not fit this platform"))
}

fn read_u16(data: &[u8], offset: usize, field: &str) -> Result<u16> {
    let bytes = read_array::<2>(data, offset, field)?;
    Ok(u16::from_le_bytes(bytes))
}

fn read_u32(data: &[u8], offset: usize, field: &str) -> Result<u32> {
    let bytes = read_array::<4>(data, offset, field)?;
    Ok(u32::from_le_bytes(bytes))
}

fn read_i32(data: &[u8], offset: usize, field: &str) -> Result<i32> {
    let bytes = read_array::<4>(data, offset, field)?;
    Ok(i32::from_le_bytes(bytes))
}

fn read_array<const N: usize>(data: &[u8], offset: usize, field: &str) -> Result<[u8; N]> {
    let end = offset
        .checked_add(N)
        .ok_or_else(|| parse(format!("{field} extent overflows")))?;
    data.get(offset..end)
        .ok_or_else(|| parse(format!("{field} is truncated")))?
        .try_into()
        .map_err(|_error| parse(format!("{field} has an invalid length")))
}

fn parse(message: impl Into<String>) -> Error {
    Error::ParseError(message.into())
}

fn unsupported(message: impl Into<String>) -> Error {
    Error::Unsupported(message.into())
}

fn limit(resource: &str, maximum: usize) -> Error {
    parse(format!(
        "{resource} exceeds the configured limit of {maximum}"
    ))
}

fn limit_u64(resource: &str, maximum: u64) -> Error {
    parse(format!(
        "{resource} exceeds the configured limit of {maximum}"
    ))
}

#[allow(
    clippy::arbitrary_source_item_ordering,
    reason = "the bounded writer is grouped directly with its trait implementations"
)]
struct BoundedCursor {
    inner: Cursor<Vec<u8>>,
    maximum: u64,
}

impl BoundedCursor {
    fn new(maximum: usize) -> Result<Self> {
        Ok(Self {
            inner: Cursor::new(Vec::new()),
            maximum: u64::try_from(maximum)
                .map_err(|_error| parse("DIB output limit does not fit u64"))?,
        })
    }

    fn into_inner(self) -> Vec<u8> {
        self.inner.into_inner()
    }
}

impl Write for BoundedCursor {
    fn write(&mut self, buffer: &[u8]) -> std::io::Result<usize> {
        let end = self
            .inner
            .position()
            .checked_add(u64::try_from(buffer.len()).unwrap_or(u64::MAX))
            .ok_or_else(output_limit_error)?;
        if end > self.maximum {
            return Err(output_limit_error());
        }
        self.inner.write(buffer)
    }

    fn flush(&mut self) -> std::io::Result<()> {
        self.inner.flush()
    }
}

impl Seek for BoundedCursor {
    fn seek(&mut self, position: SeekFrom) -> std::io::Result<u64> {
        let target = match position {
            SeekFrom::Start(value) => Some(value),
            SeekFrom::End(delta) => {
                let length = u64::try_from(self.inner.get_ref().len())
                    .map_err(|_error| output_limit_error())?;
                add_signed(length, delta)
            },
            SeekFrom::Current(delta) => add_signed(self.inner.position(), delta),
        }
        .ok_or_else(output_limit_error)?;
        if target > self.maximum {
            return Err(output_limit_error());
        }
        self.inner.seek(SeekFrom::Start(target))
    }
}

fn add_signed(base: u64, delta: i64) -> Option<u64> {
    if delta >= 0 {
        base.checked_add(delta.cast_unsigned())
    } else {
        base.checked_sub(delta.unsigned_abs())
    }
}

fn output_limit_error() -> std::io::Error {
    std::io::Error::other("DIB output exceeds configured limit")
}

#[cfg(test)]
mod tests {
    use super::*;
    use image::{GenericImageView, Rgba, RgbaImage};

    fn info_header(width: i32, height: i32, bpp: u16, compression: u32, size: u32) -> Vec<u8> {
        let mut dib = vec![0; 40];
        dib[0..4].copy_from_slice(&40u32.to_le_bytes());
        dib[4..8].copy_from_slice(&width.to_le_bytes());
        dib[8..12].copy_from_slice(&height.to_le_bytes());
        dib[12..14].copy_from_slice(&1u16.to_le_bytes());
        dib[14..16].copy_from_slice(&bpp.to_le_bytes());
        dib[16..20].copy_from_slice(&compression.to_le_bytes());
        dib[20..24].copy_from_slice(&size.to_le_bytes());
        dib
    }

    fn tiny_rgb24() -> Vec<u8> {
        let mut dib = info_header(2, 1, 24, 0, 0);
        dib.extend_from_slice(&[0, 0, 255, 0, 255, 0, 0, 0]);
        dib
    }

    fn encode(image: &DynamicImage, format: ImageFormat) -> Vec<u8> {
        let mut bytes = Cursor::new(Vec::new());
        image.write_to(&mut bytes, format).unwrap();
        bytes.into_inner()
    }

    #[test]
    fn parses_and_decodes_bottom_up_info_rgb24() {
        let bytes = tiny_rgb24();
        let dib = Dib::parse(&bytes, DibLimits::default()).unwrap();
        assert_eq!(dib.info().header, HeaderKind::Info);
        assert_eq!(dib.info().stride, Some(8));
        assert_eq!(dib.info().pixel_data_len, 8);
        let image = dib.to_dynamic_image().unwrap().to_rgba8();
        assert_eq!(image.dimensions(), (2, 1));
        assert_eq!(image.get_pixel(0, 0), &Rgba([255, 0, 0, 255]));
        assert_eq!(image.get_pixel(1, 0), &Rgba([0, 255, 0, 255]));
    }

    #[test]
    fn rgb32_alpha_is_ignored_by_default_and_explicit_for_alpha_blend() {
        let mut bytes = info_header(1, 1, 32, 0, 4);
        bytes.extend_from_slice(&[10, 20, 30, 128]);
        let dib = Dib::parse(&bytes, DibLimits::default()).unwrap();
        assert_eq!(
            dib.to_dynamic_image().unwrap().get_pixel(0, 0),
            Rgba([30, 20, 10, 255])
        );
        assert_eq!(
            dib.to_dynamic_image_with_alpha(AlphaInterpretation::Straight)
                .unwrap()
                .get_pixel(0, 0),
            Rgba([30, 20, 10, 128])
        );
        assert_eq!(
            dib.to_dynamic_image_with_alpha(AlphaInterpretation::Premultiplied)
                .unwrap()
                .get_pixel(0, 0),
            Rgba([59, 39, 19, 128])
        );
    }

    #[test]
    fn parses_core_palette_and_checked_stride() {
        let mut dib = Vec::new();
        dib.extend_from_slice(&12u32.to_le_bytes());
        dib.extend_from_slice(&2u16.to_le_bytes());
        dib.extend_from_slice(&1u16.to_le_bytes());
        dib.extend_from_slice(&1u16.to_le_bytes());
        dib.extend_from_slice(&1u16.to_le_bytes());
        dib.extend_from_slice(&[0, 0, 0, 255, 255, 255]);
        dib.extend_from_slice(&[0b0100_0000, 0, 0, 0]);
        let parsed = Dib::parse(&dib, DibLimits::default()).unwrap();
        assert_eq!(parsed.info().header, HeaderKind::Core);
        assert_eq!(parsed.info().palette_entries, 2);
        assert_eq!(parsed.info().stride, Some(4));
        assert_eq!(parsed.to_dynamic_image().unwrap().dimensions(), (2, 1));
    }

    #[test]
    fn split_info_and_bits_do_not_require_packed_offsets() {
        let info = info_header(2, -1, 24, 0, 8);
        let bits = [255, 0, 0, 255, 255, 255, 0, 0];
        let dib = Dib::parse_parts(&info, &bits, ColorUsage::Rgb, DibLimits::default()).unwrap();
        assert!(dib.info().top_down);
        assert_eq!(dib.bitmap_bits(), &bits);
        assert_eq!(dib.to_bmp().unwrap().pixel_offset(), 54);
    }

    #[test]
    fn parses_v4_bitfields_and_masks() {
        let mut header = vec![0; 108];
        header[0..4].copy_from_slice(&108u32.to_le_bytes());
        header[4..8].copy_from_slice(&1i32.to_le_bytes());
        header[8..12].copy_from_slice(&1i32.to_le_bytes());
        header[12..14].copy_from_slice(&1u16.to_le_bytes());
        header[14..16].copy_from_slice(&32u16.to_le_bytes());
        header[16..20].copy_from_slice(&3u32.to_le_bytes());
        header[20..24].copy_from_slice(&4u32.to_le_bytes());
        header[40..44].copy_from_slice(&0x00ff_0000u32.to_le_bytes());
        header[44..48].copy_from_slice(&0x0000_ff00u32.to_le_bytes());
        header[48..52].copy_from_slice(&0x0000_00ffu32.to_le_bytes());
        header[52..56].copy_from_slice(&0xff00_0000u32.to_le_bytes());
        header.extend_from_slice(&0xff11_2233u32.to_le_bytes());
        let parsed = Dib::parse(&header, DibLimits::default()).unwrap();
        assert_eq!(parsed.info().header, HeaderKind::V4);
        assert_eq!(
            parsed.info().masks,
            [0x00ff_0000, 0x0000_ff00, 0x0000_00ff, 0xff00_0000]
        );
    }

    #[test]
    fn parses_and_decodes_info_bitfields() {
        let mut dib = info_header(1, 1, 16, 3, 4);
        dib.extend_from_slice(&0x7c00u32.to_le_bytes());
        dib.extend_from_slice(&0x03e0u32.to_le_bytes());
        dib.extend_from_slice(&0x001fu32.to_le_bytes());
        dib.extend_from_slice(&[0x00, 0x7c, 0, 0]);
        let parsed = Dib::parse(&dib, DibLimits::default()).unwrap();
        assert_eq!(parsed.info().masks, [0x7c00, 0x03e0, 0x001f, 0]);
        assert_eq!(
            parsed.to_dynamic_image().unwrap().get_pixel(0, 0),
            Rgba([255, 0, 0, 255])
        );
    }

    #[test]
    fn parses_and_decodes_rle8() {
        let mut dib = info_header(2, 1, 8, 1, 6);
        dib[32..36].copy_from_slice(&2u32.to_le_bytes());
        dib.extend_from_slice(&[0, 0, 0, 0, 0, 0, 255, 0]);
        dib.extend_from_slice(&[2, 1, 0, 0, 0, 1]);
        let parsed = Dib::parse(&dib, DibLimits::default()).unwrap();
        assert_eq!(parsed.info().compression, Compression::Rle8);
        assert_eq!(
            parsed.to_dynamic_image().unwrap().get_pixel(0, 0),
            Rgba([255, 0, 0, 255])
        );
    }

    #[test]
    fn parses_v5_with_profile_before_pixels() {
        let mut dib = vec![0; 124];
        dib[0..4].copy_from_slice(&124u32.to_le_bytes());
        dib[4..8].copy_from_slice(&1i32.to_le_bytes());
        dib[8..12].copy_from_slice(&1i32.to_le_bytes());
        dib[12..14].copy_from_slice(&1u16.to_le_bytes());
        dib[14..16].copy_from_slice(&24u16.to_le_bytes());
        dib[112..116].copy_from_slice(&124u32.to_le_bytes());
        dib[116..120].copy_from_slice(&3u32.to_le_bytes());
        dib.extend_from_slice(&[1, 2, 3, 0]);
        dib.extend_from_slice(&[10, 20, 30, 0]);
        let parsed = Dib::parse(&dib, DibLimits::default()).unwrap();
        assert_eq!(parsed.info().header, HeaderKind::V5);
        assert_eq!(parsed.bitmap_info().len(), 128);
        assert_eq!(parsed.bitmap_bits(), &[10, 20, 30, 0]);
    }

    #[test]
    fn embedded_png_is_dispatched_directly_and_can_be_reencoded() {
        let source = DynamicImage::ImageRgba8(RgbaImage::from_pixel(1, 1, Rgba([7, 8, 9, 10])));
        let png = encode(&source, ImageFormat::Png);
        let mut header = info_header(1, 1, 0, 5, png.len() as u32);
        header.extend_from_slice(&png);
        let dib = Dib::parse(&header, DibLimits::default()).unwrap();
        assert_eq!(dib.info().compression, Compression::Png);
        assert_eq!(
            dib.to_dynamic_image().unwrap().get_pixel(0, 0),
            Rgba([7, 8, 9, 10])
        );
        assert!(dib.to_png().unwrap().starts_with(b"\x89PNG\r\n\x1a\n"));
        assert!(dib.to_bmp().is_err());
    }

    #[test]
    fn embedded_jpeg_is_dispatched_directly() {
        let source = DynamicImage::ImageRgb8(image::RgbImage::from_pixel(
            1,
            1,
            image::Rgb([100, 110, 120]),
        ));
        let jpeg = encode(&source, ImageFormat::Jpeg);
        let mut header = info_header(1, 1, 0, 4, jpeg.len() as u32);
        header.extend_from_slice(&jpeg);
        let dib = Dib::parse(&header, DibLimits::default()).unwrap();
        assert_eq!(dib.to_dynamic_image().unwrap().dimensions(), (1, 1));
    }

    #[test]
    fn logical_palette_entries_parse_but_require_external_palette_to_decode() {
        let mut info = info_header(1, 1, 8, 0, 4);
        info[32..36].copy_from_slice(&1u32.to_le_bytes());
        info.extend_from_slice(&5u16.to_le_bytes());
        let bits = [0, 0, 0, 0];
        let dib = Dib::parse_parts(
            &info,
            &bits,
            ColorUsage::PaletteEntries,
            DibLimits::default(),
        )
        .unwrap();
        assert_eq!(dib.info().palette_entries, 1);
        assert!(matches!(dib.to_dynamic_image(), Err(Error::Unsupported(_))));
    }

    #[test]
    fn palette_indices_have_no_color_table() {
        let info = info_header(1, 1, 8, 0, 4);
        let bits = [3, 0, 0, 0];
        let dib = Dib::parse_parts(
            &info,
            &bits,
            ColorUsage::PaletteIndices,
            DibLimits::default(),
        )
        .unwrap();
        assert_eq!(dib.info().palette_entries, 0);
        assert!(dib.to_dynamic_image().is_err());
        assert_eq!(ColorUsage::from_raw(1).unwrap(), ColorUsage::PaletteEntries);
        assert_eq!(ColorUsage::from_raw(2).unwrap(), ColorUsage::PaletteIndices);
    }

    #[test]
    fn rejects_truncation_dimensions_planes_and_bit_depth() {
        assert!(Dib::parse(&[40, 0, 0], DibLimits::default()).is_err());

        let mut zero_width = info_header(0, 1, 24, 0, 0);
        zero_width.extend_from_slice(&[0; 4]);
        assert!(Dib::parse(&zero_width, DibLimits::default()).is_err());

        let mut planes = info_header(1, 1, 24, 0, 0);
        planes[12..14].copy_from_slice(&2u16.to_le_bytes());
        planes.extend_from_slice(&[0; 4]);
        assert!(Dib::parse(&planes, DibLimits::default()).is_err());

        let bad_bpp = info_header(1, 1, 12, 0, 0);
        assert!(Dib::parse(&bad_bpp, DibLimits::default()).is_err());
    }

    #[test]
    fn rejects_compression_mismatches_and_top_down_compression() {
        assert!(Dib::parse(&info_header(1, 1, 4, 1, 1), DibLimits::default()).is_err());
        assert!(Dib::parse(&info_header(1, -1, 8, 1, 1), DibLimits::default()).is_err());
        assert!(Dib::parse(&info_header(1, 1, 24, 5, 1), DibLimits::default()).is_err());
        assert!(Dib::parse(&info_header(1, 1, 0, 5, 0), DibLimits::default()).is_err());
    }

    #[test]
    fn recognizes_cmyk_family_but_reports_decode_as_unsupported() {
        for (compression, bit_count, size, bits) in [
            (0x0b, 32, 4, &[0, 0, 0, 0][..]),
            (0x0c, 8, 2, &[0, 1][..]),
            (0x0d, 4, 2, &[0, 1][..]),
        ] {
            let mut dib = info_header(1, 1, bit_count, compression, size);
            dib.extend_from_slice(bits);
            let parsed = Dib::parse(&dib, DibLimits::default()).unwrap();
            assert!(matches!(
                parsed.to_dynamic_image(),
                Err(Error::Unsupported(_))
            ));
        }
    }

    #[test]
    fn rejects_bad_masks() {
        fn bitfields(red: u32, green: u32, blue: u32) -> Vec<u8> {
            let mut dib = info_header(1, 1, 16, 3, 4);
            dib.extend_from_slice(&red.to_le_bytes());
            dib.extend_from_slice(&green.to_le_bytes());
            dib.extend_from_slice(&blue.to_le_bytes());
            dib.extend_from_slice(&[0; 4]);
            dib
        }
        assert!(Dib::parse(&bitfields(0, 0x03e0, 0x001f), DibLimits::default()).is_err());
        assert!(Dib::parse(&bitfields(0x7c00, 0x03e0, 0x03e0), DibLimits::default()).is_err());
        assert!(Dib::parse(&bitfields(0x6c00, 0x03e0, 0x001f), DibLimits::default()).is_err());
        assert!(Dib::parse(&bitfields(0x1_0000, 0x03e0, 0x001f), DibLimits::default()).is_err());
    }

    #[test]
    fn rejects_short_palette_and_pixels() {
        let mut palette = info_header(1, 1, 8, 0, 0);
        palette[32..36].copy_from_slice(&2u32.to_le_bytes());
        palette.extend_from_slice(&[0; 7]);
        assert!(Dib::parse(&palette, DibLimits::default()).is_err());

        let mut pixels = info_header(2, 2, 24, 0, 0);
        pixels.extend_from_slice(&[0; 15]);
        assert!(Dib::parse(&pixels, DibLimits::default()).is_err());
    }

    #[test]
    fn enforces_all_relevant_limits() {
        let data = tiny_rgb24();
        let mut limits = DibLimits::default();
        limits.max_input_bytes = data.len() - 1;
        assert!(Dib::parse(&data, limits).is_err());

        let mut limits = DibLimits::default();
        limits.max_width = 1;
        assert!(Dib::parse(&data, limits).is_err());

        let mut limits = DibLimits::default();
        limits.max_pixels = 1;
        assert!(Dib::parse(&data, limits).is_err());

        let mut limits = DibLimits::default();
        limits.max_output_bytes = 10;
        let parsed = Dib::parse(&data, limits).unwrap();
        assert!(parsed.to_bmp().is_err());
    }

    #[test]
    fn decoded_dimensions_must_match_embedded_header() {
        let source = DynamicImage::new_rgb8(2, 1);
        let png = encode(&source, ImageFormat::Png);
        let mut dib = info_header(1, 1, 0, 5, png.len() as u32);
        dib.extend_from_slice(&png);
        assert!(
            Dib::parse(&dib, DibLimits::default())
                .unwrap()
                .to_dynamic_image()
                .is_err()
        );
    }

    #[test]
    fn rejects_overlapping_or_truncated_v5_profile() {
        let mut header = vec![0; 124];
        header[0..4].copy_from_slice(&124u32.to_le_bytes());
        header[4..8].copy_from_slice(&1i32.to_le_bytes());
        header[8..12].copy_from_slice(&1i32.to_le_bytes());
        header[12..14].copy_from_slice(&1u16.to_le_bytes());
        header[14..16].copy_from_slice(&24u16.to_le_bytes());
        header[112..116].copy_from_slice(&40u32.to_le_bytes());
        header[116..120].copy_from_slice(&4u32.to_le_bytes());
        header.extend_from_slice(&[0; 4]);
        assert!(Dib::parse(&header, DibLimits::default()).is_err());

        header[112..116].copy_from_slice(&124u32.to_le_bytes());
        header[116..120].copy_from_slice(&8u32.to_le_bytes());
        assert!(Dib::parse_parts(&header, &[0; 4], ColorUsage::Rgb, DibLimits::default()).is_err());
    }
}
