// Windows Metafile (WMF) format parser and converter
//
// This module provides functionality to parse WMF data and convert it to
// modern image formats (PNG, JPEG, WebP).
//
// WMF is a 16-bit vector graphics format for Windows, introduced in Windows 3.0.
// It's the predecessor to EMF (Enhanced Metafile).
//
// References:
// - [MS-WMF]: Windows Metafile Format Specification
// - https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-wmf/

pub mod converter;
pub mod parser;
pub mod svg_converter;

pub use converter::{WmfConverter, WmfToRasterOptions};
pub use parser::WmfParser;
pub use svg_converter::WmfSvgConverter;

use crate::common::error::Result;
use image::ImageFormat;

/// Convert WMF data to a raster image in the specified format
///
/// # Arguments
/// * `wmf_data` - Raw WMF file data
/// * `format` - Target image format (PNG, JPEG, WebP)
/// * `width` - Optional output width (maintains aspect ratio if only one dimension specified)
/// * `height` - Optional output height
///
/// # Returns
/// Encoded image bytes in the target format
///
/// # Example
/// ```no_run
/// use litchi::images::wmf::convert_wmf;
/// use image::ImageFormat;
///
/// let wmf_data = std::fs::read("image.wmf")?;
/// let png_data = convert_wmf(&wmf_data, ImageFormat::Png, Some(800), None)?;
/// std::fs::write("output.png", png_data)?;
/// # Ok::<(), Box<dyn std::error::Error>>(())
/// ```
pub fn convert_wmf(
    wmf_data: &[u8],
    format: ImageFormat,
    width: Option<u32>,
    height: Option<u32>,
) -> Result<Vec<u8>> {
    let parser = WmfParser::new(wmf_data)?;
    let options = WmfToRasterOptions {
        width,
        height,
        background_color: image::Rgba([255, 255, 255, 255]),
    };

    let converter = WmfConverter::new(parser, options);
    converter.convert_to_format(format)
}

/// Convert WMF data to PNG format
pub fn convert_wmf_to_png(
    wmf_data: &[u8],
    width: Option<u32>,
    height: Option<u32>,
) -> Result<Vec<u8>> {
    convert_wmf(wmf_data, ImageFormat::Png, width, height)
}

/// Convert WMF data to JPEG format
pub fn convert_wmf_to_jpeg(
    wmf_data: &[u8],
    width: Option<u32>,
    height: Option<u32>,
) -> Result<Vec<u8>> {
    convert_wmf(wmf_data, ImageFormat::Jpeg, width, height)
}

/// Convert WMF data to WebP format
pub fn convert_wmf_to_webp(
    wmf_data: &[u8],
    width: Option<u32>,
    height: Option<u32>,
) -> Result<Vec<u8>> {
    convert_wmf(wmf_data, ImageFormat::WebP, width, height)
}

/// Convert WMF data to SVG format
///
/// This converts vector graphics to minimal SVG while extracting embedded
/// raster images as PNG data URLs. Uses parallel processing for performance.
///
/// # Arguments
/// * `wmf_data` - Raw WMF file data
///
/// # Returns
/// SVG document as string
///
/// # Example
/// ```no_run
/// use litchi::images::wmf::convert_wmf_to_svg;
///
/// let wmf_data = std::fs::read("image.wmf")?;
/// let svg = convert_wmf_to_svg(&wmf_data)?;
/// std::fs::write("output.svg", svg)?;
/// # Ok::<(), Box<dyn std::error::Error>>(())
/// ```
pub fn convert_wmf_to_svg(wmf_data: &[u8]) -> Result<String> {
    let parser = WmfParser::new(wmf_data)?;
    let converter = WmfSvgConverter::new(parser);
    converter.convert_to_svg()
}

/// Convert WMF data to SVG bytes
pub fn convert_wmf_to_svg_bytes(wmf_data: &[u8]) -> Result<Vec<u8>> {
    Ok(convert_wmf_to_svg(wmf_data)?.into_bytes())
}
