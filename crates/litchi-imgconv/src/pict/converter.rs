// PICT to raster image converter
//
// Converts Macintosh PICT files to modern raster formats (PNG, JPEG, WebP).

use super::data::{get_bitmap_pixel, unpack_bits};
use super::parser::PictParser;
use super::types::{PictBitmap, PictRect};
use image::{DynamicImage, ImageBuffer, ImageFormat, Rgba, RgbaImage};
use litchi_core::error::{Error, Result};
use std::io::Cursor;

const PICT_BITMAP_HEADER_BYTES: usize = 28;

/// Options for PICT to raster conversion
#[derive(Debug, Clone)]
pub struct PictToRasterOptions {
    /// Target width (None = use source dimensions)
    pub width: Option<u32>,
    /// Target height (None = use source dimensions)
    pub height: Option<u32>,
    /// Background color for rendering
    pub background_color: Rgba<u8>,
    /// Maximum width of an embedded bitmap.
    pub max_width: u32,
    /// Maximum height of an embedded bitmap.
    pub max_height: u32,
    /// Maximum pixels decoded from an embedded bitmap.
    pub max_pixels: u64,
    /// Maximum encoded output bytes.
    pub max_output_bytes: usize,
}

impl Default for PictToRasterOptions {
    fn default() -> Self {
        Self {
            width: None,
            height: None,
            background_color: Rgba([255, 255, 255, 255]),
            max_width: 8192,
            max_height: 8192,
            max_pixels: 32 * 1024 * 1024,
            max_output_bytes: 256 * 1024 * 1024,
        }
    }
}

/// Convert Rect fields from big-endian to native endianness
#[inline]
pub fn rect_to_native(rect: &mut PictRect) {
    rect.top = i16::from_be(rect.top);
    rect.left = i16::from_be(rect.left);
    rect.bottom = i16::from_be(rect.bottom);
    rect.right = i16::from_be(rect.right);
}

/// Convert Bitmap fields from big-endian to native endianness
#[inline]
pub fn bitmap_to_native(bitmap: &mut PictBitmap) {
    bitmap.row_bytes = i16::from_be(bitmap.row_bytes);
    rect_to_native(&mut bitmap.bounds);
    rect_to_native(&mut bitmap.src_rect);
    rect_to_native(&mut bitmap.dst_rect);
    bitmap.mode = i16::from_be(bitmap.mode);
}

/// Convert Region fields from big-endian to native endianness
#[inline]
pub fn region_to_native(region: &mut super::types::PictRegion) {
    region.region_size = i16::from_be(region.region_size);
    rect_to_native(&mut region.rect);
}

/// PICT to raster converter
pub struct PictConverter {
    parser: PictParser,
    options: PictToRasterOptions,
}

impl PictConverter {
    /// Create a new PICT converter
    pub fn new(parser: PictParser, options: PictToRasterOptions) -> Self {
        Self { parser, options }
    }

    /// Calculate output dimensions maintaining aspect ratio
    fn calculate_dimensions(&self) -> (u32, u32) {
        let src_width = self.parser.width().max(1) as u32;
        let src_height = self.parser.height().max(1) as u32;

        match (self.options.width, self.options.height) {
            (Some(w), Some(h)) => (w, h),
            (Some(w), None) => {
                let aspect = src_height as f64 / src_width as f64;
                let h = (w as f64 * aspect) as u32;
                (w, h)
            },
            (None, Some(h)) => {
                let aspect = src_width as f64 / src_height as f64;
                let w = (h as f64 * aspect) as u32;
                (w, h)
            },
            (None, None) => {
                let max_dim = 4096;
                if src_width > max_dim || src_height > max_dim {
                    let scale = (max_dim as f64) / src_width.max(src_height) as f64;
                    (
                        (src_width as f64 * scale) as u32,
                        (src_height as f64 * scale) as u32,
                    )
                } else {
                    (src_width, src_height)
                }
            },
        }
    }

    /// Try to extract embedded bitmap from PICT records
    ///
    /// PICT files can contain bitmap data in:
    /// - DirectBitsRect (0x009A)
    /// - PackedDirectBitsRect (0x009B)
    /// - CompressedQuickTime (0x8200)
    fn extract_embedded_bitmap(&self) -> Result<Option<DynamicImage>> {
        for record in &self.parser.records {
            match record.opcode {
                0x009A | 0x009B => {
                    // DirectBitsRect or PackedDirectBitsRect
                    if let Some(img) = self.parse_direct_bits(&record.data)? {
                        return Ok(Some(img));
                    }
                },
                0x8200 => {
                    // CompressedQuickTime - contains JPEG or other compressed data
                    if let Some(img) = self.parse_compressed_quicktime(&record.data)? {
                        return Ok(Some(img));
                    }
                },
                _ => {},
            }
        }
        Ok(None)
    }

    /// Parse DirectBitsRect data
    ///
    /// Handles PackBitsRect (0x0098) and PackedDirectBitsRect (0x009B) opcodes.
    /// These contain compressed bitmap data that needs to be decompressed and rendered.
    fn parse_direct_bits(&self, data: &[u8]) -> Result<Option<DynamicImage>> {
        // The parser retains the opcode's two-byte size prefix. Decode every
        // fixed-width field explicitly; a Rust type containing `Vec` must
        // never be reconstructed by transmuting attacker-controlled bytes.
        let header_start = 2usize;
        let header_end = header_start + PICT_BITMAP_HEADER_BYTES;
        if data.len() < header_end {
            return Ok(None);
        }
        let mut offset = header_start;
        let row_bytes = read_i16_be(data, &mut offset)?;
        let bounds = read_rect_be(data, &mut offset)?;
        let src_rect = read_rect_be(data, &mut offset)?;
        let dst_rect = read_rect_be(data, &mut offset)?;
        let mode = read_i16_be(data, &mut offset)?;
        let bitmap = PictBitmap {
            row_bytes,
            bounds,
            src_rect,
            dst_rect,
            mode,
            data: Vec::new(),
        };

        // Calculate dimensions
        let width = u32::try_from(i32::from(bitmap.bounds.right) - i32::from(bitmap.bounds.left))
            .map_err(|_| {
            Error::ParseError("PICT bitmap has reversed horizontal bounds".into())
        })?;
        let height = u32::try_from(i32::from(bitmap.bounds.bottom) - i32::from(bitmap.bounds.top))
            .map_err(|_| Error::ParseError("PICT bitmap has reversed vertical bounds".into()))?;
        let pixels = u64::from(width)
            .checked_mul(u64::from(height))
            .ok_or_else(|| Error::ParseError("PICT bitmap pixel count overflow".into()))?;

        if width == 0
            || height == 0
            || width > self.options.max_width
            || height > self.options.max_height
            || pixels > self.options.max_pixels
        {
            return Err(Error::ParseError(
                "PICT embedded bitmap exceeds configured dimensions".into(),
            ));
        }

        // Calculate bitmap data size
        let _row_bytes = bitmap.row_bytes as usize;
        let bitmap_data_start = header_end;
        let bitmap_data_end = data.len();

        if bitmap_data_start >= bitmap_data_end {
            return Ok(None);
        }

        let compressed_data = &data[bitmap_data_start..];

        // Create output image
        let mut img = ImageBuffer::new(width, height);

        // Decompress and render each row
        let mut data_offset = 0;
        let expected_row_size = (width as usize).div_ceil(8); // Round up for byte alignment

        for y in 0..height as usize {
            if data_offset >= compressed_data.len() {
                return Err(Error::ParseError("truncated PICT bitmap rows".into()));
            }

            // Read the byte count for this row
            let byte_count = compressed_data[data_offset] as usize;
            data_offset += 1;

            // Skip the compressed data for this row (we'll decompress it)
            let row_compressed_start = data_offset;
            let row_compressed_end = data_offset
                .checked_add(byte_count)
                .ok_or_else(|| Error::ParseError("PICT bitmap row range overflow".into()))?;
            if row_compressed_end > compressed_data.len() {
                return Err(Error::ParseError("truncated PICT bitmap row".into()));
            }
            data_offset = row_compressed_end;

            if byte_count == 0 {
                return Err(Error::ParseError("empty compressed PICT bitmap row".into()));
            }

            // Decompress this row
            let row_compressed = &compressed_data[row_compressed_start..row_compressed_end];
            let unpacked_row = unpack_bits(row_compressed, expected_row_size)?;
            self.render_bitmap_row(&unpacked_row, &bitmap, y as i32, &mut img);
        }

        Ok(Some(DynamicImage::ImageRgba8(img)))
    }

    /// Render a single decompressed bitmap row to the image
    fn render_bitmap_row(
        &self,
        unpacked_row: &[u8],
        bitmap: &PictBitmap,
        y: i32,
        img: &mut ImageBuffer<Rgba<u8>, Vec<u8>>,
    ) {
        let width = i32::from(bitmap.bounds.right) - i32::from(bitmap.bounds.left);

        // Calculate source and destination rectangles relative to image bounds
        let src_left = i32::from(bitmap.src_rect.left) - i32::from(bitmap.bounds.left);
        let src_top = i32::from(bitmap.src_rect.top) - i32::from(bitmap.bounds.top);
        let src_width = i32::from(bitmap.src_rect.right) - i32::from(bitmap.src_rect.left);
        let src_height = i32::from(bitmap.src_rect.bottom) - i32::from(bitmap.src_rect.top);
        let dst_left = i32::from(bitmap.dst_rect.left) - i32::from(self.parser.header.frame.1);
        let dst_top = i32::from(bitmap.dst_rect.top) - i32::from(self.parser.header.frame.0);
        let dst_width = i32::from(bitmap.dst_rect.right) - i32::from(bitmap.dst_rect.left);
        let dst_height = i32::from(bitmap.dst_rect.bottom) - i32::from(bitmap.dst_rect.top);
        if dst_width == 0 || dst_height == 0 {
            return;
        }

        // Render each pixel in the destination row
        for x in 0..width {
            let src_x = i64::from(src_left)
                + (i64::from(x) - i64::from(dst_left)) * i64::from(src_width)
                    / i64::from(dst_width);
            let src_y = i64::from(src_top)
                + (i64::from(y) - i64::from(dst_top)) * i64::from(src_height)
                    / i64::from(dst_height);
            let src_x = i32::try_from(src_x).unwrap_or(if src_x < 0 { i32::MIN } else { i32::MAX });
            let src_y = i32::try_from(src_y).unwrap_or(if src_y < 0 { i32::MIN } else { i32::MAX });

            let color_u32 = get_bitmap_pixel(unpacked_row, &bitmap.bounds, src_x, src_y);
            let color = Rgba([
                ((color_u32 >> 16) & 0xFF) as u8, // R
                ((color_u32 >> 8) & 0xFF) as u8,  // G
                (color_u32 & 0xFF) as u8,         // B
                ((color_u32 >> 24) & 0xFF) as u8, // A
            ]);

            if x < img.width() as i32 && y < img.height() as i32 {
                img.put_pixel(x as u32, y as u32, color);
            }
        }
    }

    /// Parse CompressedQuickTime data
    ///
    /// QuickTime compressed images are often JPEG
    fn parse_compressed_quicktime(&self, data: &[u8]) -> Result<Option<DynamicImage>> {
        // QuickTime compressed data may contain JPEG or other formats
        // Try to detect and decode

        // Look for JPEG markers
        if data.len() > 2 {
            for i in 0..data.len() - 2 {
                if data[i] == 0xFF && data[i + 1] == 0xD8 {
                    // Found JPEG SOI marker
                    let jpeg = &data[i..];
                    let probe =
                        image::ImageReader::with_format(Cursor::new(jpeg), ImageFormat::Jpeg);
                    let (width, height) = probe.into_dimensions().map_err(|error| {
                        Error::ParseError(format!("invalid embedded PICT JPEG: {error}"))
                    })?;
                    let pixels =
                        u64::from(width)
                            .checked_mul(u64::from(height))
                            .ok_or_else(|| {
                                Error::ParseError("PICT JPEG pixel count overflow".into())
                            })?;
                    if width > self.options.max_width
                        || height > self.options.max_height
                        || pixels > self.options.max_pixels
                    {
                        return Err(Error::ParseError(
                            "embedded PICT JPEG exceeds configured dimensions".into(),
                        ));
                    }
                    let mut reader =
                        image::ImageReader::with_format(Cursor::new(jpeg), ImageFormat::Jpeg);
                    let mut image_limits = image::Limits::default();
                    image_limits.max_image_width = Some(self.options.max_width);
                    image_limits.max_image_height = Some(self.options.max_height);
                    image_limits.max_alloc = Some(self.options.max_pixels.saturating_mul(8));
                    reader.limits(image_limits);
                    let image = reader.decode().map_err(|error| {
                        Error::ParseError(format!("failed to decode embedded PICT JPEG: {error}"))
                    })?;
                    return Ok(Some(image));
                }
            }
        }

        Ok(None)
    }

    /// Create a placeholder image
    fn create_placeholder(&self, width: u32, height: u32) -> RgbaImage {
        let mut img = ImageBuffer::from_pixel(width, height, self.options.background_color);

        let border_color = Rgba([128, 128, 128, 255]);

        // Draw border
        for x in 0..width {
            if x < height {
                img.put_pixel(x, 0, border_color);
                img.put_pixel(x, height - 1, border_color);
            }
        }
        for y in 0..height {
            if y < width {
                img.put_pixel(0, y, border_color);
                img.put_pixel(width - 1, y, border_color);
            }
        }

        // Draw diagonals
        let min_dim = width.min(height);
        for i in 0..min_dim {
            img.put_pixel(i, i, border_color);
            if height > i {
                img.put_pixel(i, height - 1 - i, border_color);
            }
        }

        img
    }

    /// Convert PICT to a raster image
    pub fn convert_to_image(&self) -> Result<DynamicImage> {
        let (target_width, target_height) = self.calculate_dimensions();
        let target_pixels = u64::from(target_width)
            .checked_mul(u64::from(target_height))
            .ok_or_else(|| Error::ParseError("PICT target pixel count overflow".into()))?;
        if target_width == 0
            || target_height == 0
            || target_width > self.options.max_width
            || target_height > self.options.max_height
            || target_pixels > self.options.max_pixels
        {
            return Err(Error::ParseError(
                "PICT target image exceeds configured dimensions".into(),
            ));
        }

        // Try to extract embedded bitmap first
        if let Some(embedded) = self.extract_embedded_bitmap()? {
            if embedded.width() != target_width || embedded.height() != target_height {
                return Ok(DynamicImage::ImageRgba8(image::imageops::resize(
                    &embedded,
                    target_width,
                    target_height,
                    image::imageops::FilterType::Lanczos3,
                )));
            }
            return Ok(embedded);
        }

        // Create placeholder
        let placeholder = self.create_placeholder(target_width, target_height);
        Ok(DynamicImage::ImageRgba8(placeholder))
    }

    /// Convert PICT to specified image format
    pub fn convert_to_format(&self, format: ImageFormat) -> Result<Vec<u8>> {
        let image = self.convert_to_image()?;
        crate::codec::encode(&image, format, self.options.max_output_bytes)
    }

    /// Convert PICT to PNG bytes
    pub fn convert_to_png(&self) -> Result<Vec<u8>> {
        self.convert_to_format(ImageFormat::Png)
    }

    /// Convert PICT to JPEG bytes
    pub fn convert_to_jpeg(&self) -> Result<Vec<u8>> {
        self.convert_to_format(ImageFormat::Jpeg)
    }

    /// Convert PICT to WebP bytes
    pub fn convert_to_webp(&self) -> Result<Vec<u8>> {
        self.convert_to_format(ImageFormat::WebP)
    }
}

fn read_i16_be(data: &[u8], offset: &mut usize) -> Result<i16> {
    let end = offset
        .checked_add(2)
        .ok_or_else(|| Error::ParseError("PICT bitmap field offset overflow".into()))?;
    let bytes = data
        .get(*offset..end)
        .ok_or_else(|| Error::ParseError("truncated PICT bitmap field".into()))?;
    *offset = end;
    Ok(i16::from_be_bytes([bytes[0], bytes[1]]))
}

fn read_rect_be(data: &[u8], offset: &mut usize) -> Result<PictRect> {
    Ok(PictRect {
        top: read_i16_be(data, offset)?,
        left: read_i16_be(data, offset)?,
        bottom: read_i16_be(data, offset)?,
        right: read_i16_be(data, offset)?,
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    fn empty_parser() -> PictParser {
        let mut pict = vec![0; 10];
        pict.extend_from_slice(&0x00ff_u16.to_be_bytes());
        PictParser::new(&pict).unwrap()
    }

    #[test]
    fn direct_bits_uses_field_parsing_and_checks_pixels_before_allocation() {
        let mut data = vec![0, PICT_BITMAP_HEADER_BYTES as u8];
        data.extend_from_slice(&1_i16.to_be_bytes());
        for rect in [[0_i16, 0, 2, 2]; 3] {
            for value in rect {
                data.extend_from_slice(&value.to_be_bytes());
            }
        }
        data.extend_from_slice(&0_i16.to_be_bytes());
        let converter = PictConverter::new(
            empty_parser(),
            PictToRasterOptions {
                max_pixels: 1,
                ..PictToRasterOptions::default()
            },
        );
        assert!(converter.parse_direct_bits(&data).is_err());
    }

    #[test]
    fn embedded_jpeg_dimensions_are_checked_before_decode() {
        let image = DynamicImage::ImageRgb8(image::RgbImage::new(2, 2));
        let mut jpeg = Cursor::new(Vec::new());
        image.write_to(&mut jpeg, ImageFormat::Jpeg).unwrap();
        let converter = PictConverter::new(
            empty_parser(),
            PictToRasterOptions {
                max_pixels: 1,
                ..PictToRasterOptions::default()
            },
        );
        assert!(
            converter
                .parse_compressed_quicktime(&jpeg.into_inner())
                .is_err()
        );
    }
}
