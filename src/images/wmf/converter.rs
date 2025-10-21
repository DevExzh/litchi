// WMF to raster image converter
//
// Converts WMF metafiles to modern raster formats (PNG, JPEG, WebP).
//
// Similar to EMF, full WMF rendering requires implementing a complete GDI rendering engine.
// This implementation provides extraction and placeholder generation.

use super::parser::WmfParser;
use crate::common::error::{Error, Result};
use image::{DynamicImage, ImageBuffer, ImageFormat, Rgba, RgbaImage};
use std::io::Cursor;

/// Options for WMF to raster conversion
#[derive(Debug, Clone)]
pub struct WmfToRasterOptions {
    /// Target width (None = use source dimensions)
    pub width: Option<u32>,
    /// Target height (None = use source dimensions)
    pub height: Option<u32>,
    /// Background color for rendering
    pub background_color: Rgba<u8>,
}

impl Default for WmfToRasterOptions {
    fn default() -> Self {
        Self {
            width: None,
            height: None,
            background_color: Rgba([255, 255, 255, 255]),
        }
    }
}

/// WMF to raster converter
pub struct WmfConverter {
    parser: WmfParser,
    options: WmfToRasterOptions,
}

impl WmfConverter {
    /// Create a new WMF converter
    pub fn new(parser: WmfParser, options: WmfToRasterOptions) -> Self {
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
                // Use source dimensions, but cap at reasonable size
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

    /// Try to extract embedded bitmap from WMF records
    ///
    /// WMF files can contain embedded bitmaps via:
    /// - META_DIBSTRETCHBLT (0x0B41)
    /// - META_STRETCHDIB (0x0F43)
    /// - META_DIBBITBLT (0x0940)
    fn extract_embedded_bitmap(&self) -> Option<DynamicImage> {
        for record in &self.parser.records {
            match record.function {
                0x0B41 | 0x0F43 | 0x0940 => {
                    // These records may contain DIB data
                    if let Some(img) = self.parse_dib_from_record(&record.params) {
                        return Some(img);
                    }
                },
                _ => {},
            }
        }
        None
    }

    /// Parse DIB data from WMF record parameters
    fn parse_dib_from_record(&self, data: &[u8]) -> Option<DynamicImage> {
        if data.len() < 40 {
            return None;
        }

        // Try to construct BMP from DIB
        if let Ok(img) = self.construct_bmp_from_dib(data) {
            return Some(img);
        }

        None
    }

    /// Construct a BMP image from DIB data
    fn construct_bmp_from_dib(&self, dib_data: &[u8]) -> Result<DynamicImage> {
        let file_size = 14u32 + dib_data.len() as u32;
        let pixel_data_offset = 14u32 + 40u32;

        let mut bmp_data = Vec::with_capacity(file_size as usize);

        // BMP file header
        bmp_data.extend_from_slice(b"BM");
        bmp_data.extend_from_slice(&file_size.to_le_bytes());
        bmp_data.extend_from_slice(&[0u8; 4]);
        bmp_data.extend_from_slice(&pixel_data_offset.to_le_bytes());
        bmp_data.extend_from_slice(dib_data);

        match image::load_from_memory(&bmp_data) {
            Ok(img) => Ok(img),
            Err(_) => Err(Error::ParseError("Failed to load DIB as BMP".into())),
        }
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

    /// Convert WMF to a raster image
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

    /// Convert WMF to specified image format
    pub fn convert_to_format(&self, format: ImageFormat) -> Result<Vec<u8>> {
        let image = self.convert_to_image()?;

        let mut buffer = Cursor::new(Vec::new());
        image
            .write_to(&mut buffer, format)
            .map_err(|e| Error::ParseError(format!("Failed to encode image: {}", e)))?;

        Ok(buffer.into_inner())
    }

    /// Convert WMF to PNG bytes
    pub fn convert_to_png(&self) -> Result<Vec<u8>> {
        self.convert_to_format(ImageFormat::Png)
    }

    /// Convert WMF to JPEG bytes
    pub fn convert_to_jpeg(&self) -> Result<Vec<u8>> {
        self.convert_to_format(ImageFormat::Jpeg)
    }

    /// Convert WMF to WebP bytes
    pub fn convert_to_webp(&self) -> Result<Vec<u8>> {
        self.convert_to_format(ImageFormat::WebP)
    }
}
