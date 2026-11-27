// Macintosh PICT reader and renderer
//
// PICT is the QuickDraw picture format of the classic Mac OS. Version 1
// pictures use one-byte opcodes; version 2 and extended version 2 pictures
// use two-byte opcodes and add colour, pixel maps and a native resolution.
// Pictures are played back to SVG (shapes, text and patterns as vector
// elements, bitmaps and QuickTime images as embedded raster images), and
// raster output rasterizes that SVG.
//
// References:
// - Inside Macintosh: Imaging With QuickDraw (chapters 3, 4 and 7, and
//   Appendix A, "Picture Opcodes")
// - Inside Macintosh: QuickTime (Image Compression Manager, compressed
//   picture opcodes and the ImageDescription record)

pub mod converter;
pub mod parser;

/// Bitmap, pixel map and pixel pattern records
mod bitmap;
/// PackBits row decompression
mod data;
/// QuickTime image records
mod quicktime;
/// QuickDraw regions
mod region;
/// Playback to SVG
mod render;
/// Shape geometry
mod shape;
/// Common data types for PICT format
mod types;

pub use converter::{PictConverter, PictToRasterOptions};
pub use parser::{PictParser, PictVersion};
pub(crate) use render::render;

use image::ImageFormat;
use litchi_core::error::Result;

use crate::{InputFormat, Options, OutputFormat, convert_metafile};

/// Convert PICT data to a raster image in the specified format
///
/// The conversion is strict: a picture that can only be drawn approximately
/// (for example one whose QuickTime image uses a video codec) is rejected.
/// Use [`crate::convert_metafile`] with [`InputFormat::Pict`] to obtain the
/// approximation together with its diagnostics.
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
/// use litchi_imgconv::pict::convert_pict;
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
    convert_pict_with_options(
        pict_data,
        format,
        Options {
            width,
            height,
            ..Options::default()
        },
    )
}

/// Converts PICT bytes to a bounded raster format under explicit options.
pub fn convert_pict_with_options(
    pict_data: &[u8],
    format: ImageFormat,
    options: Options,
) -> Result<Vec<u8>> {
    crate::codec::rasterize_raw_metafile(pict_data, InputFormat::Pict, format, options)
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

/// Converts PICT data to SVG, rejecting pictures that need approximation.
pub fn convert_pict_to_svg(pict_data: &[u8]) -> Result<String> {
    convert_pict_to_svg_with_options(pict_data, Options::default())
}

/// Converts PICT bytes to bounded SVG under explicit options.
pub fn convert_pict_to_svg_with_options(pict_data: &[u8], options: Options) -> Result<String> {
    let converted = convert_metafile(pict_data, InputFormat::Pict, OutputFormat::Svg, options)?;
    crate::codec::reject_lossy_diagnostics(&converted.report)?;
    String::from_utf8(converted.bytes)
        .map_err(|_| litchi_core::error::Error::ParseError("SVG output was not UTF-8".to_string()))
}
