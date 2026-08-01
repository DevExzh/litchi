use std::{
    borrow::Cow,
    io::{Cursor, Read, Seek, SeekFrom, Write},
};

use image::{DynamicImage, ImageFormat, ImageReader};
use litchi_core::error::{Error, Result};
use litchi_odraw::image::{Blip, Compression, Meta};

/// Explicit resource ceilings for Office image decoding and rendering.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Limits {
    /// Maximum stored BLIP file-data bytes.
    pub max_encoded_bytes: usize,
    /// Maximum decompressed metafile bytes.
    pub max_uncompressed_bytes: usize,
    /// Maximum decoded or rendered width.
    pub max_width: u32,
    /// Maximum decoded or rendered height.
    pub max_height: u32,
    /// Maximum decoded or rendered pixel count.
    pub max_pixels: u64,
    /// Maximum encoded output bytes.
    pub max_output_bytes: usize,
}

impl Default for Limits {
    fn default() -> Self {
        Self {
            max_encoded_bytes: 256 * 1024 * 1024,
            max_uncompressed_bytes: 256 * 1024 * 1024,
            max_width: 8192,
            max_height: 8192,
            max_pixels: 32 * 1024 * 1024,
            max_output_bytes: 256 * 1024 * 1024,
        }
    }
}

/// Output sizing and resource limits for a conversion.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct Options {
    /// Optional target width.
    pub width: Option<u32>,
    /// Optional target height.
    pub height: Option<u32>,
    /// Decode and output ceilings.
    pub limits: Limits,
}

impl Options {
    /// Sets the target width while retaining the source aspect ratio when
    /// height is absent.
    pub const fn width(mut self, width: u32) -> Self {
        self.width = Some(width);
        self
    }

    /// Sets the target height while retaining the source aspect ratio when
    /// width is absent.
    pub const fn height(mut self, height: u32) -> Self {
        self.height = Some(height);
        self
    }

    /// Replaces all resource ceilings.
    pub const fn limits(mut self, limits: Limits) -> Self {
        self.limits = limits;
        self
    }
}

/// Returns uncompressed BLIP file data without rendering it.
///
/// Bitmap data remains borrowed. Uncompressed metafiles remain borrowed;
/// RFC1950-compressed metafiles allocate exactly one bounded output buffer.
pub fn decode_data<'data>(blip: &Blip<'data>, limits: &Limits) -> Result<Cow<'data, [u8]>> {
    check_encoded(blip.data().len(), limits)?;
    match blip {
        Blip::Emf(meta) | Blip::Wmf(meta) | Blip::Pict(meta) => inflate(meta, limits),
        Blip::Jpeg(bitmap) | Blip::Png(bitmap) | Blip::Dib(bitmap) | Blip::Tiff(bitmap) => {
            Ok(Cow::Borrowed(bitmap.data()))
        },
        Blip::Opaque(_) => Err(parse("cannot decode an unknown OfficeArt BLIP kind")),
    }
}

/// Decodes or renders a BLIP and encodes it in a requested raster format.
pub fn convert(blip: &Blip<'_>, format: ImageFormat, options: Options) -> Result<Vec<u8>> {
    let image = match blip {
        Blip::Emf(meta) => render_emf(meta, options)?,
        Blip::Wmf(meta) => render_wmf(meta, options)?,
        Blip::Pict(meta) => render_pict(meta, options)?,
        Blip::Jpeg(_) | Blip::Png(_) | Blip::Dib(_) | Blip::Tiff(_) => {
            decode_bitmap(blip, options)?
        },
        Blip::Opaque(_) => return Err(parse("cannot render an unknown OfficeArt BLIP kind")),
    };
    encode(&image, format, options.limits.max_output_bytes)
}

/// Converts a BLIP to PNG under explicit options.
pub fn to_png(blip: &Blip<'_>, options: Options) -> Result<Vec<u8>> {
    convert(blip, ImageFormat::Png, options)
}

/// Converts a BLIP to JPEG under explicit options.
pub fn to_jpeg(blip: &Blip<'_>, options: Options) -> Result<Vec<u8>> {
    convert(blip, ImageFormat::Jpeg, options)
}

/// Converts a BLIP to WebP under explicit options.
pub fn to_webp(blip: &Blip<'_>, options: Options) -> Result<Vec<u8>> {
    convert(blip, ImageFormat::WebP, options)
}

/// Converts an EMF or WMF BLIP to SVG under explicit input/output limits.
pub fn to_svg(blip: &Blip<'_>, options: Options) -> Result<String> {
    let data = decode_data(blip, &options.limits)?;
    let svg = match blip {
        Blip::Emf(_) => crate::emf::convert_emf_to_svg(&data)?,
        Blip::Wmf(meta) => {
            let data = wmf_with_header(meta, data, &options.limits)?;
            crate::wmf::convert_wmf_to_svg(&data)?
        },
        Blip::Pict(_) => return Err(parse("PICT to SVG rendering is not implemented")),
        _ => return Err(parse("SVG conversion requires an EMF or WMF BLIP")),
    };
    if svg.len() > options.limits.max_output_bytes {
        return Err(limit("SVG output", options.limits.max_output_bytes));
    }
    Ok(svg)
}

fn check_encoded(actual: usize, limits: &Limits) -> Result<()> {
    if actual > limits.max_encoded_bytes {
        return Err(limit("encoded BLIP", limits.max_encoded_bytes));
    }
    Ok(())
}

fn inflate<'data>(meta: &Meta<'data>, limits: &Limits) -> Result<Cow<'data, [u8]>> {
    let header = meta.header();
    let expected = usize::try_from(header.size)
        .map_err(|_| parse("metafile cbSize does not fit this platform"))?;
    if expected > limits.max_uncompressed_bytes {
        return Err(limit(
            "decompressed metafile",
            limits.max_uncompressed_bytes,
        ));
    }
    if header.compression == Compression::None {
        if meta.data().len() != expected {
            return Err(parse("uncompressed metafile length does not match cbSize"));
        }
        return Ok(Cow::Borrowed(meta.data()));
    }

    validate_zlib_header(meta.data())?;
    let mut decoder = flate2::read::ZlibDecoder::new(meta.data());
    let maximum = limits
        .max_uncompressed_bytes
        .checked_add(1)
        .ok_or_else(|| parse("decompression limit cannot be represented"))?;
    let mut output = Vec::with_capacity(expected.min(limits.max_uncompressed_bytes));
    {
        let maximum =
            u64::try_from(maximum).map_err(|_| parse("decompression limit does not fit u64"))?;
        let mut bounded = (&mut decoder).take(maximum);
        bounded
            .read_to_end(&mut output)
            .map_err(|error| parse(format!("RFC1950 decompression failed: {error}")))?;
    }
    if output.len() > limits.max_uncompressed_bytes {
        return Err(limit(
            "decompressed metafile",
            limits.max_uncompressed_bytes,
        ));
    }
    if output.len() != expected {
        return Err(parse(format!(
            "metafile cbSize is {expected}, but decompression produced {} bytes",
            output.len()
        )));
    }
    let encoded_len = u64::try_from(meta.data().len())
        .map_err(|_| parse("compressed metafile length does not fit u64"))?;
    if decoder.total_in() != encoded_len {
        return Err(parse(
            "compressed metafile has trailing bytes after RFC1950 data",
        ));
    }
    Ok(Cow::Owned(output))
}

fn validate_zlib_header(data: &[u8]) -> Result<()> {
    let header = data
        .get(..2)
        .ok_or_else(|| parse("compressed metafile has no RFC1950 header"))?;
    let cmf = header[0];
    let flags = header[1];
    let check = u16::from_be_bytes([cmf, flags]);
    if cmf & 0x0F != 8 || cmf >> 4 > 7 || check % 31 != 0 {
        return Err(parse("compressed metafile is not RFC1950-wrapped DEFLATE"));
    }
    if flags & 0x20 != 0 {
        return Err(parse("RFC1950 preset dictionaries are not supported"));
    }
    Ok(())
}

fn decode_bitmap(blip: &Blip<'_>, options: Options) -> Result<DynamicImage> {
    let data = decode_data(blip, &options.limits)?;
    let adapted = if matches!(blip, Blip::Dib(_)) {
        dib_to_bmp(&data, &options.limits)?
    } else {
        data
    };
    let probe = ImageReader::new(Cursor::new(adapted.as_ref()))
        .with_guessed_format()
        .map_err(|error| parse(format!("failed to identify bitmap: {error}")))?;
    let (source_width, source_height) = probe
        .into_dimensions()
        .map_err(|error| parse(format!("failed to read bitmap dimensions: {error}")))?;
    check_dimensions(source_width, source_height, &options.limits)?;
    let cursor = Cursor::new(adapted.as_ref());
    let mut reader = ImageReader::new(cursor)
        .with_guessed_format()
        .map_err(|error| parse(format!("failed to identify bitmap: {error}")))?;
    let mut image_limits = image::Limits::default();
    image_limits.max_image_width = Some(options.limits.max_width);
    image_limits.max_image_height = Some(options.limits.max_height);
    image_limits.max_alloc = Some(
        options
            .limits
            .max_pixels
            .checked_mul(8)
            .ok_or_else(|| parse("pixel allocation limit overflows"))?,
    );
    reader.limits(image_limits);
    let image = reader
        .decode()
        .map_err(|error| parse(format!("failed to decode bitmap: {error}")))?;
    let (width, height) = target_dimensions(
        image.width(),
        image.height(),
        options.width,
        options.height,
        &options.limits,
    )?;
    if image.width() == width && image.height() == height {
        Ok(image)
    } else {
        Ok(DynamicImage::ImageRgba8(image::imageops::resize(
            &image,
            width,
            height,
            image::imageops::FilterType::Lanczos3,
        )))
    }
}

fn render_emf(meta: &Meta<'_>, options: Options) -> Result<DynamicImage> {
    let data = inflate_checked(meta, &options.limits)?;
    let parser = crate::emf::EmfParser::new(&data)?;
    let (width, height) = target_dimensions(
        positive(parser.width()),
        positive(parser.height()),
        options.width,
        options.height,
        &options.limits,
    )?;
    crate::emf::EmfConverter::new(
        parser,
        crate::emf::EmfToRasterOptions {
            width: Some(width),
            height: Some(height),
            background_color: image::Rgba([255, 255, 255, 255]),
        },
    )
    .convert_to_image()
}

fn render_wmf(meta: &Meta<'_>, options: Options) -> Result<DynamicImage> {
    let data = inflate_checked(meta, &options.limits)?;
    let data = wmf_with_header(meta, data, &options.limits)?;
    let parser = crate::wmf::WmfParser::new(&data)?;
    let (width, height) = target_dimensions(
        positive(parser.width()),
        positive(parser.height()),
        options.width,
        options.height,
        &options.limits,
    )?;
    crate::wmf::WmfConverter::new(
        parser,
        crate::wmf::WmfToRasterOptions {
            width: Some(width),
            height: Some(height),
            background_color: image::Rgba([255, 255, 255, 255]),
        },
    )
    .convert_to_image()
}

fn render_pict(meta: &Meta<'_>, options: Options) -> Result<DynamicImage> {
    let data = inflate_checked(meta, &options.limits)?;
    let parser = crate::pict::PictParser::new(&data)?;
    let (width, height) = target_dimensions(
        positive(parser.width()),
        positive(parser.height()),
        options.width,
        options.height,
        &options.limits,
    )?;
    crate::pict::PictConverter::new(
        parser,
        crate::pict::PictToRasterOptions {
            width: Some(width),
            height: Some(height),
            background_color: image::Rgba([255, 255, 255, 255]),
        },
    )
    .convert_to_image()
}

fn inflate_checked<'data>(meta: &Meta<'data>, limits: &Limits) -> Result<Cow<'data, [u8]>> {
    check_encoded(meta.data().len(), limits)?;
    inflate(meta, limits)
}

fn positive(value: i32) -> u32 {
    u32::try_from(value)
        .ok()
        .filter(|value| *value != 0)
        .unwrap_or(1)
}

fn target_dimensions(
    source_width: u32,
    source_height: u32,
    width: Option<u32>,
    height: Option<u32>,
    limits: &Limits,
) -> Result<(u32, u32)> {
    let source_width = source_width.max(1);
    let source_height = source_height.max(1);
    let (width, height) = match (width, height) {
        (Some(0), _) | (_, Some(0)) => return Err(parse("target dimensions must be nonzero")),
        (Some(width), Some(height)) => (width, height),
        (Some(width), None) => {
            let height = proportional(source_height, width, source_width)?;
            (width, height)
        },
        (None, Some(height)) => {
            let width = proportional(source_width, height, source_height)?;
            (width, height)
        },
        (None, None) => (source_width, source_height),
    };
    check_dimensions(width, height, limits)?;
    Ok((width, height))
}

fn proportional(source_other: u32, target: u32, source_axis: u32) -> Result<u32> {
    let numerator = u64::from(source_other)
        .checked_mul(u64::from(target))
        .ok_or_else(|| parse("aspect-ratio calculation overflows"))?;
    let rounded = numerator
        .checked_add(u64::from(source_axis / 2))
        .ok_or_else(|| parse("aspect-ratio rounding overflows"))?
        / u64::from(source_axis);
    u32::try_from(rounded.max(1)).map_err(|_| parse("target dimension exceeds u32"))
}

fn check_dimensions(width: u32, height: u32, limits: &Limits) -> Result<()> {
    if width > limits.max_width {
        return Err(bound("image width", width, limits.max_width));
    }
    if height > limits.max_height {
        return Err(bound("image height", height, limits.max_height));
    }
    let pixels = u64::from(width)
        .checked_mul(u64::from(height))
        .ok_or_else(|| parse("image pixel count overflows"))?;
    if pixels > limits.max_pixels {
        return Err(parse(format!(
            "image has {pixels} pixels; limit is {}",
            limits.max_pixels
        )));
    }
    Ok(())
}

fn wmf_with_header<'data>(
    meta: &Meta<'data>,
    data: Cow<'data, [u8]>,
    limits: &Limits,
) -> Result<Cow<'data, [u8]>> {
    if data.starts_with(&0x9AC6_CDD7u32.to_le_bytes()) {
        return Ok(data);
    }
    let total = data
        .len()
        .checked_add(22)
        .ok_or_else(|| parse("WMF placeable-header length overflows"))?;
    if total > limits.max_uncompressed_bytes {
        return Err(limit(
            "WMF with placeable header",
            limits.max_uncompressed_bytes,
        ));
    }
    let header = meta.header();
    let bounds = if header.bounds == litchi_odraw::image::Rect::default() {
        let width = emu_to_twips(header.extent.x)?;
        let height = emu_to_twips(header.extent.y)?;
        (0, 0, width, height)
    } else {
        (
            checked_i16(header.bounds.left, "WMF left bound")?,
            checked_i16(header.bounds.top, "WMF top bound")?,
            checked_i16(header.bounds.right, "WMF right bound")?,
            checked_i16(header.bounds.bottom, "WMF bottom bound")?,
        )
    };
    let mut output = Vec::with_capacity(total);
    output.extend_from_slice(&0x9AC6_CDD7u32.to_le_bytes());
    output.extend_from_slice(&0u16.to_le_bytes());
    output.extend_from_slice(&bounds.0.to_le_bytes());
    output.extend_from_slice(&bounds.1.to_le_bytes());
    output.extend_from_slice(&bounds.2.to_le_bytes());
    output.extend_from_slice(&bounds.3.to_le_bytes());
    output.extend_from_slice(&1440u16.to_le_bytes());
    output.extend_from_slice(&0u32.to_le_bytes());
    let checksum = output[..20].chunks_exact(2).fold(0u16, |sum, bytes| {
        sum ^ u16::from_le_bytes([bytes[0], bytes[1]])
    });
    output.extend_from_slice(&checksum.to_le_bytes());
    output.extend_from_slice(&data);
    Ok(Cow::Owned(output))
}

fn emu_to_twips(value: i32) -> Result<i16> {
    let scaled = i64::from(value)
        .checked_mul(1440)
        .ok_or_else(|| parse("EMU-to-twips conversion overflows"))?
        / 914_400;
    i16::try_from(scaled).map_err(|_| parse("WMF extent does not fit a placeable header"))
}

fn checked_i16(value: i32, field: &str) -> Result<i16> {
    i16::try_from(value).map_err(|_| parse(format!("{field} does not fit i16")))
}

fn dib_to_bmp<'data>(data: &'data [u8], limits: &Limits) -> Result<Cow<'data, [u8]>> {
    if data.starts_with(b"BM") {
        return Ok(Cow::Borrowed(data));
    }
    let header_size = read_u32(data, 0, "DIB header size")?;
    let header_size_usize =
        usize::try_from(header_size).map_err(|_| parse("DIB header size does not fit usize"))?;
    if header_size_usize < 12 || header_size_usize > data.len() {
        return Err(parse("DIB header size is invalid"));
    }
    let (bit_count, colors_used, palette_entry, masks) = if header_size == 12 {
        (
            u32::from(read_u16(data, 10, "DIB bit count")?),
            0,
            3usize,
            0usize,
        )
    } else if header_size >= 40 {
        let bit_count = u32::from(read_u16(data, 14, "DIB bit count")?);
        let compression = read_u32(data, 16, "DIB compression")?;
        let colors_used = read_u32(data, 32, "DIB palette size")?;
        let masks = if header_size == 40 {
            match compression {
                3 => 12,
                6 => 16,
                _ => 0,
            }
        } else {
            0
        };
        (bit_count, colors_used, 4usize, masks)
    } else {
        return Err(parse("unsupported DIB header layout"));
    };
    let palette_colors = if colors_used != 0 {
        colors_used
    } else if bit_count <= 8 {
        1u32.checked_shl(bit_count)
            .ok_or_else(|| parse("DIB palette size overflows"))?
    } else {
        0
    };
    let palette = usize::try_from(palette_colors)
        .ok()
        .and_then(|count| count.checked_mul(palette_entry))
        .ok_or_else(|| parse("DIB palette byte length overflows"))?;
    let pixel_offset = 14usize
        .checked_add(header_size_usize)
        .and_then(|value| value.checked_add(masks))
        .and_then(|value| value.checked_add(palette))
        .ok_or_else(|| parse("BMP pixel offset overflows"))?;
    let total = data
        .len()
        .checked_add(14)
        .ok_or_else(|| parse("BMP file size overflows"))?;
    if total > limits.max_encoded_bytes.saturating_add(14) {
        return Err(limit("adapted BMP", limits.max_encoded_bytes));
    }
    if pixel_offset > total {
        return Err(parse("DIB palette extends past the image data"));
    }
    let total_u32 = u32::try_from(total).map_err(|_| parse("BMP file size exceeds u32"))?;
    let pixel_u32 =
        u32::try_from(pixel_offset).map_err(|_| parse("BMP pixel offset exceeds u32"))?;
    let mut bmp = Vec::with_capacity(total);
    bmp.extend_from_slice(b"BM");
    bmp.extend_from_slice(&total_u32.to_le_bytes());
    bmp.extend_from_slice(&0u32.to_le_bytes());
    bmp.extend_from_slice(&pixel_u32.to_le_bytes());
    bmp.extend_from_slice(data);
    Ok(Cow::Owned(bmp))
}

fn read_u16(data: &[u8], offset: usize, field: &str) -> Result<u16> {
    let end = offset
        .checked_add(2)
        .ok_or_else(|| parse(format!("{field} extent overflows")))?;
    let bytes = data
        .get(offset..end)
        .ok_or_else(|| parse(format!("{field} is truncated")))?;
    Ok(u16::from_le_bytes([bytes[0], bytes[1]]))
}

fn read_u32(data: &[u8], offset: usize, field: &str) -> Result<u32> {
    let end = offset
        .checked_add(4)
        .ok_or_else(|| parse(format!("{field} extent overflows")))?;
    let bytes = data
        .get(offset..end)
        .ok_or_else(|| parse(format!("{field} is truncated")))?;
    Ok(u32::from_le_bytes([bytes[0], bytes[1], bytes[2], bytes[3]]))
}

fn encode(image: &DynamicImage, format: ImageFormat, maximum: usize) -> Result<Vec<u8>> {
    let mut output = LimitedCursor::new(maximum)?;
    image
        .write_to(&mut output, format)
        .map_err(|error| parse(format!("failed to encode image: {error}")))?;
    Ok(output.into_inner())
}

struct LimitedCursor {
    inner: Cursor<Vec<u8>>,
    maximum: u64,
}

impl LimitedCursor {
    fn new(maximum: usize) -> Result<Self> {
        Ok(Self {
            inner: Cursor::new(Vec::new()),
            maximum: u64::try_from(maximum)
                .map_err(|_| parse("encoded output limit does not fit u64"))?,
        })
    }

    fn into_inner(self) -> Vec<u8> {
        self.inner.into_inner()
    }
}

impl Write for LimitedCursor {
    fn write(&mut self, buffer: &[u8]) -> std::io::Result<usize> {
        let length = u64::try_from(buffer.len())
            .map_err(|_| std::io::Error::other("encoded write length does not fit u64"))?;
        let end = self
            .inner
            .position()
            .checked_add(length)
            .ok_or_else(|| std::io::Error::other("encoded output position overflows"))?;
        if end > self.maximum {
            return Err(std::io::Error::other("encoded output limit exceeded"));
        }
        self.inner.write(buffer)
    }

    fn flush(&mut self) -> std::io::Result<()> {
        self.inner.flush()
    }
}

impl Seek for LimitedCursor {
    fn seek(&mut self, position: SeekFrom) -> std::io::Result<u64> {
        let previous = self.inner.position();
        let next = self.inner.seek(position)?;
        if next > self.maximum {
            self.inner.set_position(previous);
            return Err(std::io::Error::other("encoded output seek exceeds limit"));
        }
        Ok(next)
    }
}

fn parse(error: impl Into<String>) -> Error {
    Error::ParseError(error.into())
}

fn limit(resource: &str, maximum: usize) -> Error {
    parse(format!("{resource} exceeds the {maximum}-byte limit"))
}

fn bound(resource: &str, actual: u32, maximum: u32) -> Error {
    parse(format!("{resource} is {actual}; limit is {maximum}"))
}

#[cfg(test)]
mod tests {
    use super::*;
    use flate2::{
        Compression as Level,
        write::{DeflateEncoder, ZlibEncoder},
    };
    use litchi_odraw::image::Blip;

    fn record(instance: u16, kind: u16, body: &[u8]) -> Vec<u8> {
        let mut bytes = Vec::new();
        bytes.extend_from_slice(&(instance << 4).to_le_bytes());
        bytes.extend_from_slice(&kind.to_le_bytes());
        bytes.extend_from_slice(&(body.len() as u32).to_le_bytes());
        bytes.extend_from_slice(body);
        bytes
    }

    fn meta(compressed: &[u8], uncompressed_size: u32) -> Vec<u8> {
        let mut body = vec![0; 16];
        body.extend_from_slice(&uncompressed_size.to_le_bytes());
        body.extend_from_slice(&[0; 24]);
        body.extend_from_slice(&(compressed.len() as u32).to_le_bytes());
        body.extend_from_slice(&[0x00, 0xFE]);
        body.extend_from_slice(compressed);
        record(0x3D4, 0xF01A, &body)
    }

    #[test]
    fn accepts_only_rfc1950_and_requires_exact_cb_size() {
        let mut zlib = ZlibEncoder::new(Vec::new(), Level::default());
        zlib.write_all(b"metafile").unwrap();
        let encoded = zlib.finish().unwrap();
        let bytes = meta(&encoded, 8);
        let blip = Blip::parse(&bytes).unwrap();
        assert_eq!(
            decode_data(&blip, &Limits::default()).unwrap(),
            b"metafile".as_slice()
        );

        let wrong = meta(&encoded, 9);
        let wrong = Blip::parse(&wrong).unwrap();
        assert!(decode_data(&wrong, &Limits::default()).is_err());

        let mut raw = DeflateEncoder::new(Vec::new(), Level::default());
        raw.write_all(b"metafile").unwrap();
        let raw = raw.finish().unwrap();
        let raw = meta(&raw, 8);
        let raw = Blip::parse(&raw).unwrap();
        assert!(decode_data(&raw, &Limits::default()).is_err());
    }

    #[test]
    fn enforces_decompression_limit_before_allocation() {
        let mut zlib = ZlibEncoder::new(Vec::new(), Level::default());
        zlib.write_all(&[0; 1024]).unwrap();
        let encoded = zlib.finish().unwrap();
        let bytes = meta(&encoded, 1024);
        let blip = Blip::parse(&bytes).unwrap();
        let limits = Limits {
            max_uncompressed_bytes: 32,
            ..Limits::default()
        };
        assert!(decode_data(&blip, &limits).is_err());
    }

    #[test]
    fn bitmap_decode_honors_pixel_and_output_limits() {
        let image = DynamicImage::new_rgba8(4, 4);
        let mut png = Cursor::new(Vec::new());
        image.write_to(&mut png, ImageFormat::Png).unwrap();
        let png = png.into_inner();
        let mut body = vec![0; 16];
        body.push(0xFF);
        body.extend_from_slice(&png);
        let bytes = record(0x6E0, 0xF01E, &body);
        let blip = Blip::parse(&bytes).unwrap();
        let limits = Limits {
            max_pixels: 15,
            ..Limits::default()
        };
        assert!(to_png(&blip, Options::default().limits(limits)).is_err());

        let limits = Limits {
            max_output_bytes: 4,
            ..Limits::default()
        };
        assert!(to_png(&blip, Options::default().limits(limits)).is_err());
    }
}
