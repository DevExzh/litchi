// Image processing and conversion module
//
// This module provides functionality to parse and convert Office Drawing formats
// (EMF, WMF, PICT) to modern image standards (PNG, JPEG, WebP).
//
// # Architecture
//
// - `blip`: Core BLIP (Binary Large Image or Picture) record parsing
// - `emf`: Enhanced Metafile (EMF) format support
// - `wmf`: Windows Metafile (WMF) format support
// - `pict`: Macintosh PICT format support
//
// # Example: Converting a BLIP record
//
// ```no_run
// use litchi::images::blip::Blip;
// use litchi::images::convert_blip_to_png;
//
/// let blip_data = vec![/* BLIP record bytes */];
/// let blip = Blip::parse(&blip_data)?;
/// let png_bytes = convert_blip_to_png(&blip, Some(800), None)?;
/// # Ok::<(), litchi::common::error::Error>(())
/// ```
///
/// # Example: Converting EMF to PNG
///
/// ```no_run
/// use litchi::images::emf::convert_emf_to_png;
///
/// let emf_data = std::fs::read("image.emf")?;
/// let png_data = convert_emf_to_png(&emf_data, Some(800), None)?;
/// std::fs::write("output.png", png_data)?;
/// # Ok::<(), Box<dyn std::error::Error>>(())
/// ```
pub mod blip;
pub mod emf;
pub mod pict;
pub mod svg;
pub mod wmf;

use crate::common::error::Result;
pub use blip::{BitmapBlip, Blip, BlipType, MetafileBlip};
use image::ImageFormat;

/// Convert a BLIP record to a raster image format
///
/// This is a high-level convenience function that handles all BLIP types
/// (EMF, WMF, PICT, JPEG, PNG, DIB, TIFF) and converts them to the specified format.
///
/// # Arguments
/// * `blip` - Parsed BLIP record
/// * `format` - Target image format
/// * `width` - Optional output width
/// * `height` - Optional output height
///
/// # Returns
/// Encoded image bytes in the target format
pub fn convert_blip_to_format(
    blip: &Blip,
    format: ImageFormat,
    width: Option<u32>,
    height: Option<u32>,
) -> Result<Vec<u8>> {
    match blip {
        Blip::Metafile(metafile) => {
            let data = metafile.decompress()?;
            match metafile.blip_type() {
                Some(BlipType::Emf) => emf::convert_emf(&data, format, width, height),
                Some(BlipType::Wmf) => wmf::convert_wmf(&data, format, width, height),
                Some(BlipType::Pict) => pict::convert_pict(&data, format, width, height),
                _ => Err(crate::common::error::Error::ParseError(
                    "Unknown metafile BLIP type".into(),
                )),
            }
        },
        Blip::Bitmap(bitmap) => {
            // For bitmap formats that are already in a modern format, we may just need
            // to re-encode or pass through
            let img = image::load_from_memory(&bitmap.picture_data).map_err(|e| {
                crate::common::error::Error::ParseError(format!("Failed to load bitmap: {}", e))
            })?;

            // Resize if requested
            let img = match (width, height) {
                (Some(w), Some(h)) if img.width() != w || img.height() != h => {
                    image::DynamicImage::ImageRgba8(image::imageops::resize(
                        &img,
                        w,
                        h,
                        image::imageops::FilterType::Lanczos3,
                    ))
                },
                (Some(w), None) => {
                    let aspect = img.height() as f64 / img.width() as f64;
                    let h = (w as f64 * aspect) as u32;
                    image::DynamicImage::ImageRgba8(image::imageops::resize(
                        &img,
                        w,
                        h,
                        image::imageops::FilterType::Lanczos3,
                    ))
                },
                (None, Some(h)) => {
                    let aspect = img.width() as f64 / img.height() as f64;
                    let w = (h as f64 * aspect) as u32;
                    image::DynamicImage::ImageRgba8(image::imageops::resize(
                        &img,
                        w,
                        h,
                        image::imageops::FilterType::Lanczos3,
                    ))
                },
                _ => img,
            };

            // Encode to target format
            let mut buffer = std::io::Cursor::new(Vec::new());
            img.write_to(&mut buffer, format).map_err(|e| {
                crate::common::error::Error::ParseError(format!("Failed to encode image: {}", e))
            })?;

            Ok(buffer.into_inner())
        },
    }
}

/// Convert a BLIP record to PNG format
///
/// # Arguments
/// * `blip` - Parsed BLIP record
/// * `width` - Optional output width
/// * `height` - Optional output height
///
/// # Returns
/// PNG-encoded image bytes
///
/// # Example
/// ```no_run
/// use litchi::images::{blip::Blip, convert_blip_to_png};
///
/// let blip_data = vec![/* BLIP record bytes */];
/// let blip = Blip::parse(&blip_data)?;
/// let png = convert_blip_to_png(&blip, Some(800), None)?;
/// # Ok::<(), litchi::common::error::Error>(())
/// ```
pub fn convert_blip_to_png(
    blip: &Blip,
    width: Option<u32>,
    height: Option<u32>,
) -> Result<Vec<u8>> {
    convert_blip_to_format(blip, ImageFormat::Png, width, height)
}

/// Convert a BLIP record to JPEG format
pub fn convert_blip_to_jpeg(
    blip: &Blip,
    width: Option<u32>,
    height: Option<u32>,
) -> Result<Vec<u8>> {
    convert_blip_to_format(blip, ImageFormat::Jpeg, width, height)
}

/// Convert a BLIP record to WebP format
pub fn convert_blip_to_webp(
    blip: &Blip,
    width: Option<u32>,
    height: Option<u32>,
) -> Result<Vec<u8>> {
    convert_blip_to_format(blip, ImageFormat::WebP, width, height)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_blip_type_extensions() {
        assert_eq!(BlipType::Emf.extension(), "emf");
        assert_eq!(BlipType::Png.extension(), "png");
        assert_eq!(BlipType::Jpeg.extension(), "jpg");
    }

    #[test]
    fn test_blip_type_classification() {
        assert!(BlipType::Emf.is_metafile());
        assert!(BlipType::Wmf.is_metafile());
        assert!(BlipType::Pict.is_metafile());
        assert!(!BlipType::Jpeg.is_metafile());
        assert!(!BlipType::Png.is_metafile());
    }
}
