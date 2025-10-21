// Enhanced Metafile (EMF) format parser and converter
//
// This module provides functionality to parse EMF data and convert it to
// modern image formats (PNG, JPEG, WebP).
//
// EMF is a 32-bit vector graphics format for Windows, introduced in Windows NT 3.1.
// It's an improved version of WMF with better support for modern graphics features.
//
// References:
// - [MS-EMF]: Enhanced Metafile Format Specification
// - https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-emf/

pub mod converter;
pub mod parser;
pub mod svg_converter;

pub use converter::{EmfConverter, EmfToRasterOptions};
pub use parser::EmfParser;
pub use svg_converter::EmfSvgConverter;

use crate::common::error::Result;
use image::ImageFormat;

/// Convert EMF data to a raster image in the specified format
///
/// # Arguments
/// * `emf_data` - Raw EMF file data
/// * `format` - Target image format (PNG, JPEG, WebP)
/// * `width` - Optional output width (maintains aspect ratio if only one dimension specified)
/// * `height` - Optional output height
///
/// # Returns
/// Encoded image bytes in the target format
///
/// # Example
/// ```no_run
/// use litchi::images::emf::convert_emf;
/// use image::ImageFormat;
///
/// let emf_data = std::fs::read("image.emf")?;
/// let png_data = convert_emf(&emf_data, ImageFormat::Png, Some(800), None)?;
/// std::fs::write("output.png", png_data)?;
/// # Ok::<(), Box<dyn std::error::Error>>(())
/// ```
pub fn convert_emf(
    emf_data: &[u8],
    format: ImageFormat,
    width: Option<u32>,
    height: Option<u32>,
) -> Result<Vec<u8>> {
    let parser = EmfParser::new(emf_data)?;
    let options = EmfToRasterOptions {
        width,
        height,
        background_color: image::Rgba([255, 255, 255, 255]),
    };

    let converter = EmfConverter::new(parser, options);
    converter.convert_to_format(format)
}

/// Convert EMF data to PNG format
///
/// # Arguments
/// * `emf_data` - Raw EMF file data
/// * `width` - Optional output width
/// * `height` - Optional output height
///
/// # Returns
/// PNG-encoded image bytes
pub fn convert_emf_to_png(
    emf_data: &[u8],
    width: Option<u32>,
    height: Option<u32>,
) -> Result<Vec<u8>> {
    convert_emf(emf_data, ImageFormat::Png, width, height)
}

/// Convert EMF data to JPEG format
///
/// # Arguments
/// * `emf_data` - Raw EMF file data
/// * `width` - Optional output width
/// * `height` - Optional output height
///
/// # Returns
/// JPEG-encoded image bytes
pub fn convert_emf_to_jpeg(
    emf_data: &[u8],
    width: Option<u32>,
    height: Option<u32>,
) -> Result<Vec<u8>> {
    convert_emf(emf_data, ImageFormat::Jpeg, width, height)
}

/// Convert EMF data to WebP format
///
/// # Arguments
/// * `emf_data` - Raw EMF file data
/// * `width` - Optional output width
/// * `height` - Optional output height
///
/// # Returns
/// WebP-encoded image bytes
pub fn convert_emf_to_webp(
    emf_data: &[u8],
    width: Option<u32>,
    height: Option<u32>,
) -> Result<Vec<u8>> {
    convert_emf(emf_data, ImageFormat::WebP, width, height)
}

/// Convert EMF data to SVG format
///
/// This converts vector graphics to minimal SVG while extracting embedded
/// raster images as PNG data URLs. Uses parallel processing for performance.
///
/// # Arguments
/// * `emf_data` - Raw EMF file data
///
/// # Returns
/// SVG document as string
///
/// # Example
/// ```no_run
/// use litchi::images::emf::convert_emf_to_svg;
///
/// let emf_data = std::fs::read("image.emf")?;
/// let svg = convert_emf_to_svg(&emf_data)?;
/// std::fs::write("output.svg", svg)?;
/// # Ok::<(), Box<dyn std::error::Error>>(())
/// ```
pub fn convert_emf_to_svg(emf_data: &[u8]) -> Result<String> {
    let parser = EmfParser::new(emf_data)?;
    let converter = EmfSvgConverter::new(parser);
    converter.convert_to_svg()
}

/// Convert EMF data to SVG bytes
pub fn convert_emf_to_svg_bytes(emf_data: &[u8]) -> Result<Vec<u8>> {
    Ok(convert_emf_to_svg(emf_data)?.into_bytes())
}
