// Macintosh PICT format parser and converter
//
// This module provides functionality to parse PICT data and convert it to
// modern image formats (PNG, JPEG, WebP).
//
// PICT is the native graphics metafile format for Mac OS Classic. There are
// two versions: PICT 1 (original) and PICT 2 (extended).
//
// References:
// - Inside Macintosh: Imaging With QuickDraw
// - Apple Technical Note TN1023: Understanding the PICT Format

pub mod converter;
pub mod parser;

/// Common data types for PICT format
mod types;

/// Data manipulation and compression utilities
mod data;

pub use converter::{PictConverter, PictToRasterOptions};
pub use parser::{PictParser, PictVersion};

use crate::common::error::Result;
use image::ImageFormat;

/// Convert PICT data to a raster image in the specified format
///
/// # Arguments
/// * `pict_data` - Raw PICT file data
/// * `format` - Target image format (PNG, JPEG, WebP)
/// * `width` - Optional output width (maintains aspect ratio if only one dimension specified)
/// * `height` - Optional output height
///
/// # Returns
/// Encoded image bytes in the target format
///
/// # Example
/// ```no_run
/// use litchi::images::pict::convert_pict;
/// use image::ImageFormat;
///
/// let pict_data = std::fs::read("image.pict")?;
/// let png_data = convert_pict(&pict_data, ImageFormat::Png, Some(800), None)?;
/// std::fs::write("output.png", png_data)?;
/// # Ok::<(), Box<dyn std::error::Error>>(())
/// ```
pub fn convert_pict(
    pict_data: &[u8],
    format: ImageFormat,
    width: Option<u32>,
    height: Option<u32>,
) -> Result<Vec<u8>> {
    let parser = PictParser::new(pict_data)?;
    let options = PictToRasterOptions {
        width,
        height,
        background_color: image::Rgba([255, 255, 255, 255]),
    };

    let converter = PictConverter::new(parser, options);
    converter.convert_to_format(format)
}

/// Convert PICT data to PNG format
pub fn convert_pict_to_png(
    pict_data: &[u8],
    width: Option<u32>,
    height: Option<u32>,
) -> Result<Vec<u8>> {
    convert_pict(pict_data, ImageFormat::Png, width, height)
}

/// Convert PICT data to JPEG format
pub fn convert_pict_to_jpeg(
    pict_data: &[u8],
    width: Option<u32>,
    height: Option<u32>,
) -> Result<Vec<u8>> {
    convert_pict(pict_data, ImageFormat::Jpeg, width, height)
}

/// Convert PICT data to WebP format
pub fn convert_pict_to_webp(
    pict_data: &[u8],
    width: Option<u32>,
    height: Option<u32>,
) -> Result<Vec<u8>> {
    convert_pict(pict_data, ImageFormat::WebP, width, height)
}
