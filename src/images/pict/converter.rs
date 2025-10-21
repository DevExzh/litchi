// PICT to raster image converter
//
// Converts Macintosh PICT files to modern raster formats (PNG, JPEG, WebP).

use super::data::{get_bitmap_pixel, stretch_coordinates, unpack_bits};
use super::parser::PictParser;
use super::types::{PictBitmap, PictRect};
use crate::common::error::{Error, Result};
use image::{DynamicImage, ImageBuffer, ImageFormat, Rgba, RgbaImage};
use std::io::Cursor;

/// Options for PICT to raster conversion
#[derive(Debug, Clone)]
pub struct PictToRasterOptions {
    /// Target width (None = use source dimensions)
    pub width: Option<u32>,
    /// Target height (None = use source dimensions)
    pub height: Option<u32>,
    /// Background color for rendering
    pub background_color: Rgba<u8>,
}

impl Default for PictToRasterOptions {
    fn default() -> Self {
        Self {
            width: None,
            height: None,
            background_color: Rgba([255, 255, 255, 255]),
        }
    }
}

/// Convert Rect fields from big-endian to native endianness
#[inline]
pub fn rect_to_native(rect: &mut super::types::PictRect) {
    rect.top = i16::from_be(rect.top);
    rect.left = i16::from_be(rect.left);
    rect.bottom = i16::from_be(rect.bottom);
    rect.right = i16::from_be(rect.right);
}

/// Convert Bitmap fields from big-endian to native endianness
#[inline]
pub fn bitmap_to_native(bitmap: &mut super::types::PictBitmap) {
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
    fn extract_embedded_bitmap(&self) -> Option<DynamicImage> {
        for record in &self.parser.records {
            match record.opcode {
                0x009A | 0x009B => {
                    // DirectBitsRect or PackedDirectBitsRect
                    if let Some(img) = self.parse_direct_bits(&record.data) {
                        return Some(img);
                    }
                },
                0x8200 => {
                    // CompressedQuickTime - contains JPEG or other compressed data
                    if let Some(img) = self.parse_compressed_quicktime(&record.data) {
                        return Some(img);
                    }
                },
                _ => {},
            }
        }
        None
    }

    /// Parse DirectBitsRect data
    ///
    /// Handles PackBitsRect (0x0098) and PackedDirectBitsRect (0x009B) opcodes.
    /// These contain compressed bitmap data that needs to be decompressed and rendered.
    fn parse_direct_bits(&self, data: &[u8]) -> Option<DynamicImage> {
        if data.len() < std::mem::size_of::<PictBitmap>() {
            return None;
        }

        // Parse the bitmap header (big-endian format)
        let mut bitmap: PictBitmap =
            unsafe { std::ptr::read_unaligned(data.as_ptr() as *const PictBitmap) };

        // Convert from big-endian to native endianness
        bitmap_to_native(&mut bitmap);

        // Calculate dimensions
        let width = (bitmap.bounds.right - bitmap.bounds.left) as u32;
        let height = (bitmap.bounds.bottom - bitmap.bounds.top) as u32;

        if width == 0 || height == 0 || width > 8192 || height > 8192 {
            return None;
        }

        // Calculate bitmap data size
        let _row_bytes = bitmap.row_bytes as usize;
        let bitmap_data_start = std::mem::size_of::<PictBitmap>();
        let bitmap_data_end = data.len();

        if bitmap_data_start >= bitmap_data_end {
            return None;
        }

        let compressed_data = &data[bitmap_data_start..];

        // Create output image
        let mut img = ImageBuffer::new(width, height);

        // Decompress and render each row
        let mut data_offset = 0;
        let expected_row_size = (width as usize).div_ceil(8); // Round up for byte alignment

        for y in 0..height as usize {
            if data_offset >= compressed_data.len() {
                break;
            }

            // Read the byte count for this row
            if data_offset + 1 >= compressed_data.len() {
                break;
            }
            let byte_count = compressed_data[data_offset] as usize;
            data_offset += 1;

            // Skip the compressed data for this row (we'll decompress it)
            let row_compressed_start = data_offset;
            let row_compressed_end = std::cmp::min(data_offset + byte_count, compressed_data.len());
            data_offset = row_compressed_end;

            if byte_count == 0 {
                continue;
            }

            // Decompress this row
            let row_compressed = &compressed_data[row_compressed_start..row_compressed_end];
            match unpack_bits(row_compressed, expected_row_size) {
                Ok(unpacked_row) => {
                    // Render the unpacked row to the image
                    self.render_bitmap_row(&unpacked_row, &bitmap, y as i32, &mut img);
                },
                Err(_) => {
                    // If decompression fails, skip this row
                    continue;
                },
            }
        }

        Some(DynamicImage::ImageRgba8(img))
    }

    /// Render a single decompressed bitmap row to the image
    fn render_bitmap_row(
        &self,
        unpacked_row: &[u8],
        bitmap: &PictBitmap,
        y: i32,
        img: &mut ImageBuffer<Rgba<u8>, Vec<u8>>,
    ) {
        let width = (bitmap.bounds.right - bitmap.bounds.left) as u32;
        let _height = (bitmap.bounds.bottom - bitmap.bounds.top) as u32;

        // Calculate source and destination rectangles relative to image bounds
        let src_width = bitmap.src_rect.right - bitmap.src_rect.left;
        let src_height = bitmap.src_rect.bottom - bitmap.src_rect.top;
        let dst_width = bitmap.dst_rect.right - bitmap.dst_rect.left;
        let dst_height = bitmap.dst_rect.bottom - bitmap.dst_rect.top;

        // Create adjusted rectangles relative to bitmap bounds
        let src_rect = PictRect {
            left: bitmap.src_rect.left - bitmap.bounds.left,
            top: bitmap.src_rect.top - bitmap.bounds.top,
            right: bitmap.src_rect.left - bitmap.bounds.left + src_width,
            bottom: bitmap.src_rect.top - bitmap.bounds.top + src_height,
        };

        let dst_rect = PictRect {
            left: bitmap.dst_rect.left - self.parser.header.frame.1,
            top: bitmap.dst_rect.top - self.parser.header.frame.0,
            right: bitmap.dst_rect.left - self.parser.header.frame.1 + dst_width,
            bottom: bitmap.dst_rect.top - self.parser.header.frame.0 + dst_height,
        };

        // Render each pixel in the destination row
        for x in 0..width as i32 {
            let mut src_x = 0;
            let mut src_y = 0;

            stretch_coordinates(&dst_rect, &src_rect, x, y, &mut src_x, &mut src_y);

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
    fn parse_compressed_quicktime(&self, data: &[u8]) -> Option<DynamicImage> {
        // QuickTime compressed data may contain JPEG or other formats
        // Try to detect and decode

        // Look for JPEG markers
        if data.len() > 2 {
            for i in 0..data.len() - 2 {
                if data[i] == 0xFF && data[i + 1] == 0xD8 {
                    // Found JPEG SOI marker
                    if let Ok(img) = image::load_from_memory(&data[i..]) {
                        return Some(img);
                    }
                }
            }
        }

        None
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

        // Try to extract embedded bitmap first
        if let Some(embedded) = self.extract_embedded_bitmap() {
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

        let mut buffer = Cursor::new(Vec::new());
        image
            .write_to(&mut buffer, format)
            .map_err(|e| Error::ParseError(format!("Failed to encode image: {}", e)))?;

        Ok(buffer.into_inner())
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
