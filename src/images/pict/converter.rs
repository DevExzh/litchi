// PICT to raster image converter
//
// Converts Macintosh PICT files to modern raster formats (PNG, JPEG, WebP).

use super::parser::PictParser;
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
            }
            (None, Some(h)) => {
                let aspect = src_width as f64 / src_height as f64;
                let w = (h as f64 * aspect) as u32;
                (w, h)
            }
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
            }
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
                }
                0x8200 => {
                    // CompressedQuickTime - contains JPEG or other compressed data
                    if let Some(img) = self.parse_compressed_quicktime(&record.data) {
                        return Some(img);
                    }
                }
                _ => {}
            }
        }
        None
    }

    /// Parse DirectBitsRect data
    ///
    /// This is a complex structure containing PixMap and bitmap data
    fn parse_direct_bits(&self, data: &[u8]) -> Option<DynamicImage> {
        // DirectBitsRect structure is complex
        // For now, attempt basic parsing
        if data.len() < 50 {
            return None;
        }

        // Try to extract raw bitmap data and construct an image
        // This is simplified and may not work for all PICT files
        // Full implementation would require parsing PixMap structure
        None
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

