//! QuickTime image records embedded in PICT pictures.
//!
//! `CompressedQuickTime` (opcode `0x8200`) carries a transformation matrix,
//! an optional matte and mask region, an `ImageDescription` and the codec
//! data (*Inside Macintosh: QuickTime*, "Image Compression Manager",
//! compressed-picture opcode and `ImageDescription` record). Codecs whose
//! data is an ordinary still-image file are decoded here; video codecs are
//! reported to the caller as unsupported.

use std::io::Cursor;

use image::{DynamicImage, ImageFormat, ImageReader, RgbaImage};

use super::bitmap::default_palette;
use super::types::{PictRect, read_fixed, read_i16, read_u16, read_u32};
use crate::Limits;

/// Fixed fields of a `CompressedQuickTime` record before the matte data.
const HEADER_BYTES: usize = 72;
/// Fixed size of an `ImageDescription` without extensions.
const DESCRIPTION_BYTES: usize = 86;

/// A parsed `CompressedQuickTime` record.
#[derive(Debug, Clone)]
pub(super) struct CompressedImage<'a> {
    /// The matrix `a b c d tx ty`, mapping image coordinates to picture
    /// coordinates as `x' = a*x + c*y + tx`, `y' = b*x + d*y + ty`.
    pub(super) matrix: [f64; 6],
    /// Whether the record carries a matte (an alpha image).
    pub(super) has_matte: bool,
    /// Transfer mode.
    pub(super) mode: u16,
    /// Mask region limiting the drawn area, in picture coordinates.
    pub(super) mask_region: Option<&'a [u8]>,
    pub(super) description: ImageDescription<'a>,
    /// Compressed image data.
    pub(super) data: &'a [u8],
}

/// The fields of an `ImageDescription` used for decoding.
#[derive(Debug, Clone)]
pub(super) struct ImageDescription<'a> {
    pub(super) codec: [u8; 4],
    pub(super) width: u16,
    pub(super) height: u16,
    pub(super) depth: u16,
    pub(super) clut_id: i16,
    /// Bytes after the fixed 86-byte description (colour table, atoms).
    pub(super) extension: &'a [u8],
}

impl ImageDescription<'_> {
    /// The codec identifier as text.
    pub(super) fn codec_name(&self) -> String {
        self.codec.iter().map(|&byte| char::from(byte)).collect()
    }
}

/// Parses a `CompressedQuickTime` record, starting with its 32-bit length.
pub(super) fn parse_compressed(record: &[u8]) -> Option<CompressedImage<'_>> {
    let length = usize::try_from(read_u32(record, 0)?).ok()?;
    let body = record.get(4..4usize.checked_add(length)?)?;
    let mut matrix = [0.0; 6];
    for (slot, index) in matrix.iter_mut().zip([0usize, 1, 3, 4, 6, 7]) {
        *slot = read_fixed(body, 2 + index * 4)?;
    }
    let matte_size = usize::try_from(read_u32(body, 38)?).ok()?;
    let mode = read_u16(body, 50)?;
    let mask_size = usize::try_from(read_u32(body, 64)?).ok()?;
    let mask_start = (HEADER_BYTES - 4).checked_add(matte_size)?;
    let description_start = mask_start.checked_add(mask_size)?;
    let mask_region = (mask_size > 0)
        .then(|| body.get(mask_start..description_start))
        .flatten();
    let description_size = usize::try_from(read_u32(body, description_start)?).ok()?;
    if description_size < DESCRIPTION_BYTES {
        return None;
    }
    let description_end = description_start.checked_add(description_size)?;
    let description = body.get(description_start..description_end)?;
    let codec: [u8; 4] = description.get(4..8)?.try_into().ok()?;
    let data_size = usize::try_from(read_u32(description, 44)?).ok()?;
    let data = body.get(description_end..)?;
    // Some writers record a zero data size; the codec data then runs to the
    // end of the record.
    let data = match data_size {
        0 => data,
        size => data.get(..size.min(data.len()))?,
    };
    Some(CompressedImage {
        matrix,
        has_matte: matte_size > 0,
        mode,
        mask_region,
        description: ImageDescription {
            codec,
            width: read_u16(description, 32)?,
            height: read_u16(description, 34)?,
            depth: read_u16(description, 82)?,
            clut_id: read_i16(description, 84)?,
            extension: &description[DESCRIPTION_BYTES..],
        },
        data,
    })
}

/// A decoded QuickTime image ready to embed in SVG.
#[derive(Debug, Clone)]
pub(super) struct EmbeddedImage {
    pub(super) mime: &'static str,
    pub(super) bytes: Vec<u8>,
    pub(super) width: u32,
    pub(super) height: u32,
    /// Lossy diagnostics of a nested metafile conversion.
    pub(super) diagnostics: Vec<crate::ConversionDiagnostic>,
}

/// Outcome of decoding a QuickTime image.
#[derive(Debug, Clone)]
pub(super) enum Decoded {
    Image(EmbeddedImage),
    /// The codec is not one this crate decodes (video codecs).
    Unsupported,
    /// The codec is supported but its data could not be decoded.
    Failed(String),
}

/// Decodes the codec data of a QuickTime image.
pub(super) fn decode(image: &CompressedImage<'_>, limits: &Limits) -> Decoded {
    let description = &image.description;
    let result = match &description.codec {
        b"jpeg" => decode_jpeg(image.data, limits),
        b"png " => decode_file(image.data, Some(ImageFormat::Png), limits),
        b"gif " => decode_file(image.data, Some(ImageFormat::Gif), limits),
        b"tiff" => decode_tiff(image.data, limits),
        b"BMP " => decode_file(image.data, Some(ImageFormat::Bmp), limits),
        b"raw " => decode_raw(description, image.data, limits),
        b"WRAW" | b"WRLE" => decode_windows_dib(description, image.data, limits),
        b"WMF " | b"WMFC" => decode_windows_metafile(description, image.data, limits),
        _ => return Decoded::Unsupported,
    };
    match result {
        Ok(image) => Decoded::Image(image),
        Err(reason) => Decoded::Failed(reason),
    }
}

fn image_limits(limits: &Limits) -> image::Limits {
    let mut image_limits = image::Limits::default();
    image_limits.max_image_width = Some(limits.max_width);
    image_limits.max_image_height = Some(limits.max_height);
    image_limits.max_alloc = limits.max_pixels.checked_mul(8);
    image_limits
}

fn check_size(width: u32, height: u32, limits: &Limits) -> Result<(), String> {
    if width == 0 || height == 0 {
        return Err("image has no pixels".into());
    }
    if width > limits.max_width
        || height > limits.max_height
        || u64::from(width) * u64::from(height) > limits.max_pixels
    {
        return Err(format!("{width}x{height} image exceeds the pixel limits"));
    }
    Ok(())
}

fn read_image(
    data: &[u8],
    format: Option<ImageFormat>,
    limits: &Limits,
) -> Result<DynamicImage, String> {
    if data.len() > limits.max_encoded_bytes {
        return Err("image data exceeds the encoded-size limit".into());
    }
    let reader = |data| -> Result<ImageReader<Cursor<&[u8]>>, String> {
        let reader = ImageReader::new(Cursor::new(data));
        match format {
            Some(format) => {
                let mut reader = reader;
                reader.set_format(format);
                Ok(reader)
            },
            None => reader
                .with_guessed_format()
                .map_err(|error| error.to_string()),
        }
    };
    let (width, height) = reader(data)?
        .into_dimensions()
        .map_err(|error| error.to_string())?;
    check_size(width, height, limits)?;
    let mut decoder = reader(data)?;
    decoder.limits(image_limits(limits));
    decoder.decode().map_err(|error| error.to_string())
}

fn encode_png(image: &DynamicImage, limits: &Limits) -> Result<EmbeddedImage, String> {
    let bytes = crate::codec::encode(image, ImageFormat::Png, limits.max_output_bytes)
        .map_err(|error| error.to_string())?;
    Ok(EmbeddedImage {
        mime: "image/png",
        bytes,
        width: image.width(),
        height: image.height(),
        diagnostics: Vec::new(),
    })
}

/// JPEG data stays JPEG when it is an ordinary RGB or greyscale image; the
/// decode verifies it before embedding.
fn decode_jpeg(data: &[u8], limits: &Limits) -> Result<EmbeddedImage, String> {
    let image = read_image(data, Some(ImageFormat::Jpeg), limits)?;
    if matches!(
        image,
        DynamicImage::ImageRgb8(_) | DynamicImage::ImageLuma8(_)
    ) {
        return Ok(EmbeddedImage {
            mime: "image/jpeg",
            bytes: data.to_vec(),
            width: image.width(),
            height: image.height(),
            diagnostics: Vec::new(),
        });
    }
    encode_png(&image, limits)
}

fn decode_file(
    data: &[u8],
    format: Option<ImageFormat>,
    limits: &Limits,
) -> Result<EmbeddedImage, String> {
    let image = read_image(data, format, limits)?;
    encode_png(&image, limits)
}

/// Decodes TIFF data. Palette-colour TIFFs, which the TIFF decoder does not
/// expand, are decoded as greyscale indices and mapped through their
/// `ColorMap` (TIFF 6.0, section 5, "Palette-color Images").
fn decode_tiff(data: &[u8], limits: &Limits) -> Result<EmbeddedImage, String> {
    let Some(palette) = TiffPalette::find(data) else {
        return decode_file(data, Some(ImageFormat::Tiff), limits);
    };
    let mut patched = data.to_vec();
    palette.mark_greyscale(&mut patched);
    let indices = read_image(&patched, Some(ImageFormat::Tiff), limits)?;
    let DynamicImage::ImageLuma8(indices) = indices else {
        return Err("palette TIFF indices did not decode as 8-bit samples".into());
    };
    let colors = palette.colors(data).ok_or("truncated TIFF colour map")?;
    // Samples narrower than eight bits are scaled to the 8-bit range.
    let levels = (1u32 << palette.bits) - 1;
    let mut image = RgbaImage::new(indices.width(), indices.height());
    for (target, source) in image.pixels_mut().zip(indices.pixels()) {
        let value = u32::from(source.0[0]);
        let index = if palette.bits < 8 {
            (value * levels + 127) / 255
        } else {
            value
        };
        let color = colors.get(index as usize).copied().unwrap_or([0, 0, 0]);
        *target = image::Rgba([color[0], color[1], color[2], 255]);
    }
    encode_png(&DynamicImage::ImageRgba8(image), limits)
}

/// The palette fields of a TIFF's first image directory.
struct TiffPalette {
    big_endian: bool,
    /// Offset of the `PhotometricInterpretation` value.
    photometric: usize,
    /// Offset and entry count of the `ColorMap` values.
    color_map: (usize, usize),
    bits: u32,
}

impl TiffPalette {
    fn find(data: &[u8]) -> Option<Self> {
        let big_endian = match data.get(..4)? {
            b"MM\0*" => true,
            b"II*\0" => false,
            _ => return None,
        };
        let u16_at = |at: usize| {
            let bytes = [*data.get(at)?, *data.get(at + 1)?];
            Some(if big_endian {
                u16::from_be_bytes(bytes)
            } else {
                u16::from_le_bytes(bytes)
            })
        };
        let u32_at = |at: usize| {
            let bytes: [u8; 4] = data.get(at..at + 4)?.try_into().ok()?;
            Some(if big_endian {
                u32::from_be_bytes(bytes)
            } else {
                u32::from_le_bytes(bytes)
            })
        };
        let directory = usize::try_from(u32_at(4)?).ok()?;
        let entries = usize::from(u16_at(directory)?);
        let mut photometric = None;
        let mut color_map = None;
        let mut bits = 8;
        for index in 0..entries {
            let entry = directory + 2 + index * 12;
            let (tag, kind) = (u16_at(entry)?, u16_at(entry + 2)?);
            let count = usize::try_from(u32_at(entry + 4)?).ok()?;
            match tag {
                // PhotometricInterpretation 3 is palette colour.
                262 if kind == 3 && count == 1 && u16_at(entry + 8)? == 3 => {
                    photometric = Some(entry + 8)
                },
                258 if kind == 3 && count == 1 => bits = u32::from(u16_at(entry + 8)?),
                320 if kind == 3 => {
                    color_map = Some((usize::try_from(u32_at(entry + 8)?).ok()?, count))
                },
                _ => {},
            }
        }
        if !matches!(bits, 1 | 2 | 4 | 8) {
            return None;
        }
        Some(Self {
            big_endian,
            photometric: photometric?,
            color_map: color_map?,
            bits,
        })
    }

    /// Relabels the image as `BlackIsZero` greyscale so that its samples
    /// decode as palette indices.
    fn mark_greyscale(&self, data: &mut [u8]) {
        let value = if self.big_endian { [0, 1] } else { [1, 0] };
        data[self.photometric..self.photometric + 2].copy_from_slice(&value);
    }

    /// The colour map: all red values, then green, then blue, 16 bits each.
    fn colors(&self, data: &[u8]) -> Option<Vec<[u8; 3]>> {
        let (offset, count) = self.color_map;
        let entries = count / 3;
        let high_byte = |index: usize| {
            let at = offset.checked_add(index.checked_mul(2)?)?;
            let bytes = data.get(at..at + 2)?;
            Some(if self.big_endian { bytes[0] } else { bytes[1] })
        };
        (0..entries)
            .map(|index| {
                Some([
                    high_byte(index)?,
                    high_byte(entries + index)?,
                    high_byte(2 * entries + index)?,
                ])
            })
            .collect()
    }
}

/// Decodes the `raw ` codec: uncompressed rows of the description's depth.
/// Depths 33-40 are greyscale versions of depths 1-8.
fn decode_raw(
    description: &ImageDescription<'_>,
    data: &[u8],
    limits: &Limits,
) -> Result<EmbeddedImage, String> {
    let width = u32::from(description.width);
    let height = u32::from(description.height);
    check_size(width, height, limits)?;
    let (bits, gray) = match description.depth {
        depth @ (1 | 2 | 4 | 8 | 16 | 24 | 32) => (depth, false),
        depth @ (33 | 34 | 36 | 40) => (depth - 32, true),
        depth => return Err(format!("raw images of depth {depth} are not defined")),
    };
    let columns = usize::from(description.width);
    let rows = usize::from(description.height);
    let minimum = (columns * usize::from(bits)).div_ceil(8);
    let row_bytes = if data.len() % rows == 0 && data.len() / rows >= minimum {
        data.len() / rows
    } else {
        minimum
    };
    if row_bytes
        .checked_mul(rows)
        .is_none_or(|needed| needed > data.len())
    {
        return Err("raw image data is shorter than its dimensions".into());
    }
    let palette = if gray {
        let levels = (1usize << bits) - 1;
        (0..=levels)
            .map(|level| {
                let value = u8::try_from(255 - level * 255 / levels).unwrap_or(0);
                [value; 3]
            })
            .collect()
    } else if bits <= 8 {
        description_palette(description).unwrap_or_else(|| default_palette(bits))
    } else {
        Vec::new()
    };
    let mut image = RgbaImage::new(width, height);
    for (y, row) in data.chunks(row_bytes).take(rows).enumerate() {
        for x in 0..columns {
            let color = match bits {
                1 | 2 | 4 | 8 => {
                    let bit = x * usize::from(bits);
                    let shift = 8 - usize::from(bits) - bit % 8;
                    let index = usize::from((row[bit / 8] >> shift) & ((1u16 << bits) - 1) as u8);
                    let color = palette.get(index).copied().unwrap_or([0, 0, 0]);
                    [color[0], color[1], color[2]]
                },
                16 => {
                    let value = u16::from_be_bytes([row[x * 2], row[x * 2 + 1]]);
                    let expand = |component: u16| {
                        let component = (component & 0x1F) as u8;
                        (component << 3) | (component >> 2)
                    };
                    [expand(value >> 10), expand(value >> 5), expand(value)]
                },
                24 => [row[x * 3], row[x * 3 + 1], row[x * 3 + 2]],
                _ => [row[x * 4 + 1], row[x * 4 + 2], row[x * 4 + 3]],
            };
            image.put_pixel(
                x as u32,
                y as u32,
                image::Rgba([color[0], color[1], color[2], 255]),
            );
        }
    }
    encode_png(&DynamicImage::ImageRgba8(image), limits)
}

/// Reads the colour table that follows an indexed image description whose
/// `clutID` is 0.
fn description_palette(description: &ImageDescription<'_>) -> Option<Vec<[u8; 3]>> {
    if description.clut_id != 0 {
        return None;
    }
    let table = description.extension;
    let size = usize::from(read_u16(table, 6)?);
    let mut palette = vec![[0u8; 3]; 256];
    for (index, slot) in palette.iter_mut().enumerate().take(size.min(255) + 1) {
        let entry = table.get(8 + index * 8..16 + index * 8)?;
        *slot = [entry[2], entry[4], entry[6]];
    }
    Some(palette)
}

/// Decodes Windows DIB pixels (`WRAW` uncompressed, `WRLE` RLE) whose
/// `BITMAPINFOHEADER` is stored in a `strf` atom of the description and whose
/// colour table, for indexed depths, precedes it as a QuickDraw colour table.
fn decode_windows_dib(
    description: &ImageDescription<'_>,
    data: &[u8],
    limits: &Limits,
) -> Result<EmbeddedImage, String> {
    let extension = description.extension;
    let mut atoms = 0usize;
    let mut palette = Vec::new();
    if description.depth <= 8 && description.clut_id == 0 {
        let size = read_u16(extension, 6).ok_or("truncated colour table")?;
        let entries = usize::from(size) + 1;
        for index in 0..entries {
            let entry = extension
                .get(8 + index * 8..16 + index * 8)
                .ok_or("truncated colour table")?;
            palette.extend_from_slice(&[entry[6], entry[4], entry[2], 0]);
        }
        atoms = 8 + entries * 8;
    }
    let mut header = None;
    while let Some(size) = read_u32(extension, atoms) {
        let size = usize::try_from(size).map_err(|_| "atom size overflows")?;
        if size < 8 {
            break;
        }
        let end = atoms.checked_add(size).ok_or("atom size overflows")?;
        if extension.get(atoms + 4..atoms + 8) == Some(b"strf") {
            header = extension.get(atoms + 8..end);
            break;
        }
        atoms = end;
    }
    let header = header.ok_or("no Windows bitmap header in the image description")?;
    let header_size = header
        .get(..4)
        .map(|bytes| u32::from_le_bytes([bytes[0], bytes[1], bytes[2], bytes[3]]))
        .ok_or("truncated Windows bitmap header")?;
    let header_size = usize::try_from(header_size).map_err(|_| "bitmap header size overflows")?;
    let mut info = header
        .get(..header_size)
        .ok_or("truncated Windows bitmap header")?
        .to_vec();
    if !palette.is_empty() && info.len() >= 36 {
        // biClrUsed: the number of colour-table entries that follow.
        let used = u32::try_from(palette.len() / 4).unwrap_or(0);
        info[32..36].copy_from_slice(&used.to_le_bytes());
        info.extend_from_slice(&palette);
    }
    let dib_limits = crate::dib::DibLimits {
        max_input_bytes: limits.max_uncompressed_bytes,
        max_width: limits.max_width,
        max_height: limits.max_height,
        max_pixels: limits.max_pixels,
        max_palette_entries: 256,
        max_decoded_bytes: limits.max_uncompressed_bytes,
        max_output_bytes: limits.max_output_bytes,
    };
    let dib = crate::dib::Dib::parse_parts(&info, data, crate::dib::ColorUsage::Rgb, dib_limits)
        .map_err(|error| error.to_string())?;
    let image = dib.to_dynamic_image().map_err(|error| error.to_string())?;
    encode_png(&image, limits)
}

/// Renders a wrapped Windows metafile (`WMF `, `WMFC`) to PNG at the image
/// description's size through the crate's WMF converter.
fn decode_windows_metafile(
    description: &ImageDescription<'_>,
    data: &[u8],
    limits: &Limits,
) -> Result<EmbeddedImage, String> {
    let width = u32::from(description.width);
    let height = u32::from(description.height);
    check_size(width, height, limits)?;
    let options = crate::Options {
        width: Some(width),
        height: Some(height),
        limits: *limits,
    };
    let converted = crate::convert_metafile(
        data,
        crate::InputFormat::Wmf,
        crate::OutputFormat::Png,
        options,
    )
    .map_err(|error| error.to_string())?;
    let diagnostics = converted
        .report
        .diagnostics
        .into_iter()
        .filter(|diagnostic| crate::codec::is_lossy_diagnostic(diagnostic.code))
        .collect();
    Ok(EmbeddedImage {
        mime: "image/png",
        bytes: converted.bytes,
        width,
        height,
        diagnostics,
    })
}

/// The picture-space rectangle of an image's `(0, 0, width, height)` bounds
/// after the matrix, when the matrix only scales and translates.
pub(super) fn axis_aligned_destination(image: &CompressedImage<'_>) -> Option<PictRect> {
    let [a, b, c, d, tx, ty] = image.matrix;
    if b != 0.0 || c != 0.0 {
        return None;
    }
    let width = f64::from(image.description.width);
    let height = f64::from(image.description.height);
    let clamp = |value: f64| {
        value
            .round()
            .clamp(f64::from(i16::MIN), f64::from(i16::MAX)) as i16
    };
    let (x0, x1) = (tx, a * width + tx);
    let (y0, y1) = (ty, d * height + ty);
    Some(PictRect {
        top: clamp(y0.min(y1)),
        left: clamp(x0.min(x1)),
        bottom: clamp(y0.max(y1)),
        right: clamp(x0.max(x1)),
    })
}

#[cfg(test)]
pub(super) mod tests {
    use super::*;

    /// Builds a `CompressedQuickTime` record (length included) that places a
    /// `width` x `height` image with codec `codec` at `(tx, ty)`.
    pub(in crate::pict) fn compressed_record(
        codec: &[u8; 4],
        width: u16,
        height: u16,
        depth: u16,
        (tx, ty): (i16, i16),
        data: &[u8],
    ) -> Vec<u8> {
        let mut body = Vec::new();
        body.extend_from_slice(&0u16.to_be_bytes()); // version
        let fixed = |value: i16| (i32::from(value) << 16).to_be_bytes();
        for value in [
            fixed(1),
            fixed(0),
            0i32.to_be_bytes(),
            fixed(0),
            fixed(1),
            0i32.to_be_bytes(),
            fixed(tx),
            fixed(ty),
            (1i32 << 30).to_be_bytes(),
        ] {
            body.extend_from_slice(&value);
        }
        body.extend_from_slice(&0u32.to_be_bytes()); // matteSize
        body.extend_from_slice(&[0; 8]); // matteRect
        body.extend_from_slice(&0u16.to_be_bytes()); // mode
        for value in [0i16, 0, height as i16, width as i16] {
            body.extend_from_slice(&value.to_be_bytes());
        }
        body.extend_from_slice(&0u32.to_be_bytes()); // accuracy
        body.extend_from_slice(&0u32.to_be_bytes()); // maskSize
        let mut description = Vec::new();
        description.extend_from_slice(&86u32.to_be_bytes());
        description.extend_from_slice(codec);
        description.extend_from_slice(&[0; 24]);
        description.extend_from_slice(&width.to_be_bytes());
        description.extend_from_slice(&height.to_be_bytes());
        description.extend_from_slice(&0x0048_0000u32.to_be_bytes());
        description.extend_from_slice(&0x0048_0000u32.to_be_bytes());
        description.extend_from_slice(&(data.len() as u32).to_be_bytes());
        description.extend_from_slice(&1u16.to_be_bytes());
        description.extend_from_slice(&[0; 32]);
        description.extend_from_slice(&depth.to_be_bytes());
        description.extend_from_slice(&(-1i16).to_be_bytes());
        assert_eq!(description.len(), DESCRIPTION_BYTES);
        body.extend_from_slice(&description);
        body.extend_from_slice(data);
        let mut record = (body.len() as u32).to_be_bytes().to_vec();
        record.extend_from_slice(&body);
        record
    }

    #[test]
    fn compressed_record_fields_are_located() {
        let record = compressed_record(b"jpeg", 4, 3, 24, (10, 20), &[1, 2, 3]);
        let image = parse_compressed(&record).unwrap();
        assert_eq!(image.matrix, [1.0, 0.0, 0.0, 1.0, 10.0, 20.0]);
        assert_eq!(&image.description.codec, b"jpeg");
        assert_eq!((image.description.width, image.description.height), (4, 3));
        assert_eq!(image.data, [1, 2, 3]);
        assert!(image.mask_region.is_none());
        assert_eq!(
            axis_aligned_destination(&image),
            Some(PictRect {
                top: 20,
                left: 10,
                bottom: 23,
                right: 14
            })
        );
    }

    #[test]
    fn jpeg_codec_data_is_embedded_after_verification() {
        let source =
            DynamicImage::ImageRgb8(image::RgbImage::from_pixel(4, 3, image::Rgb([200, 10, 10])));
        let mut jpeg = Cursor::new(Vec::new());
        source.write_to(&mut jpeg, ImageFormat::Jpeg).unwrap();
        let jpeg = jpeg.into_inner();
        let record = compressed_record(b"jpeg", 4, 3, 24, (0, 0), &jpeg);
        let image = parse_compressed(&record).unwrap();
        let Decoded::Image(embedded) = decode(&image, &Limits::default()) else {
            panic!("JPEG should decode");
        };
        assert_eq!(embedded.mime, "image/jpeg");
        assert_eq!((embedded.width, embedded.height), (4, 3));
        assert_eq!(embedded.bytes, jpeg);

        let small = Limits {
            max_pixels: 4,
            ..Limits::default()
        };
        assert!(matches!(decode(&image, &small), Decoded::Failed(_)));
    }

    #[test]
    fn raw_codec_decodes_32_bit_pixels() {
        let data = [0, 255, 0, 0, 0, 0, 0, 255];
        let record = compressed_record(b"raw ", 2, 1, 32, (0, 0), &data);
        let image = parse_compressed(&record).unwrap();
        let Decoded::Image(embedded) = decode(&image, &Limits::default()) else {
            panic!("raw should decode");
        };
        let decoded = image::load_from_memory(&embedded.bytes).unwrap().to_rgba8();
        assert_eq!(decoded.get_pixel(0, 0).0, [255, 0, 0, 255]);
        assert_eq!(decoded.get_pixel(1, 0).0, [0, 0, 255, 255]);
    }

    /// A big-endian, uncompressed, 8-bit palette TIFF of two pixels.
    fn palette_tiff() -> Vec<u8> {
        let mut tiff = b"MM\0*".to_vec();
        tiff.extend_from_slice(&8u32.to_be_bytes());
        let entries: [(u16, u16, u32, u32); 9] = [
            (256, 3, 1, 2 << 16), // ImageWidth
            (257, 3, 1, 1 << 16), // ImageLength
            (258, 3, 1, 8 << 16), // BitsPerSample
            (259, 3, 1, 1 << 16), // Compression: none
            (262, 3, 1, 3 << 16), // PhotometricInterpretation: palette
            (273, 4, 1, 122),     // StripOffsets
            (278, 3, 1, 1 << 16), // RowsPerStrip
            (279, 4, 1, 2),       // StripByteCounts
            (320, 3, 768, 124),   // ColorMap
        ];
        tiff.extend_from_slice(&(entries.len() as u16).to_be_bytes());
        for (tag, kind, count, value) in entries {
            tiff.extend_from_slice(&tag.to_be_bytes());
            tiff.extend_from_slice(&kind.to_be_bytes());
            tiff.extend_from_slice(&count.to_be_bytes());
            tiff.extend_from_slice(&value.to_be_bytes());
        }
        tiff.extend_from_slice(&0u32.to_be_bytes());
        assert_eq!(tiff.len(), 122);
        tiff.extend_from_slice(&[0, 1]); // pixel indices
        let mut map = vec![0u16; 768];
        map[1] = 0xFF00; // index 1: red
        map[256 + 1] = 0x8000; // half green
        for value in map {
            tiff.extend_from_slice(&value.to_be_bytes());
        }
        tiff
    }

    #[test]
    fn palette_tiffs_are_expanded_through_their_color_map() {
        let embedded = decode_tiff(&palette_tiff(), &Limits::default()).unwrap();
        let image = image::load_from_memory(&embedded.bytes).unwrap().to_rgba8();
        assert_eq!(image.get_pixel(0, 0).0, [0, 0, 0, 255]);
        assert_eq!(image.get_pixel(1, 0).0, [0xFF, 0x80, 0, 255]);
    }

    #[test]
    fn video_codecs_are_unsupported() {
        let record = compressed_record(b"cvid", 2, 2, 24, (0, 0), &[0; 16]);
        let image = parse_compressed(&record).unwrap();
        assert!(matches!(
            decode(&image, &Limits::default()),
            Decoded::Unsupported
        ));
    }
}
