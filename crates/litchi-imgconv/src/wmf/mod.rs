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

mod constants;
pub mod converter;
pub mod parser;
mod svg;

pub use constants::*;

pub use converter::{WmfConverter, WmfToRasterOptions};
pub use parser::WmfParser;
pub use svg::WmfConverter as WmfSvgConverter;

use image::ImageFormat;
use litchi_core::error::Result;

use crate::{InputFormat, Options, convert_metafile};

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
/// use litchi_imgconv::wmf::convert_wmf;
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
    convert_wmf_with_options(
        wmf_data,
        format,
        Options {
            width,
            height,
            ..Options::default()
        },
    )
}

/// Converts WMF bytes to a bounded raster format under explicit options.
pub fn convert_wmf_with_options(
    wmf_data: &[u8],
    format: ImageFormat,
    options: Options,
) -> Result<Vec<u8>> {
    crate::codec::rasterize_raw_metafile(wmf_data, InputFormat::Wmf, format, options)
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
///
/// # Arguments
/// * `wmf_data` - Raw WMF file data
///
/// # Returns
/// SVG document as string
///
/// # Example
/// ```no_run
/// use litchi_imgconv::wmf::convert_wmf_to_svg;
///
/// let wmf_data = std::fs::read("image.wmf")?;
/// let svg = convert_wmf_to_svg(&wmf_data)?;
/// std::fs::write("output.svg", svg)?;
/// # Ok::<(), Box<dyn std::error::Error>>(())
/// ```
pub fn convert_wmf_to_svg(wmf_data: &[u8]) -> Result<String> {
    convert_wmf_to_svg_with_options(wmf_data, Options::default())
}

/// Converts WMF bytes to bounded SVG under explicit options.
pub fn convert_wmf_to_svg_with_options(wmf_data: &[u8], options: Options) -> Result<String> {
    let converted = convert_metafile(
        wmf_data,
        InputFormat::Wmf,
        crate::OutputFormat::Svg,
        options,
    )?;
    crate::codec::reject_lossy_diagnostics(&converted.report)?;
    String::from_utf8(converted.bytes)
        .map_err(|_| litchi_core::error::Error::ParseError("SVG output was not UTF-8".to_string()))
}
/// Convert WMF data to SVG bytes
pub fn convert_wmf_to_svg_bytes(wmf_data: &[u8]) -> Result<Vec<u8>> {
    Ok(convert_wmf_to_svg(wmf_data)?.into_bytes())
}
