// Image processing and conversion module
//
// This module provides functionality to parse and convert Office Drawing formats
// (EMF, WMF, PICT) to modern image standards (PNG, JPEG, WebP).
//
// # Features
//
// - **Format Auto-detection**: Automatically detects image format from BLIP records
// - **Robust Parsing**: Handles malformed records and compression issues gracefully
// - **High Performance**: Zero-copy parsing where possible, minimal allocations
// - **Wide Format Support**: EMF, WMF, PICT, PNG, JPEG, DIB, TIFF
// - **Batch Extraction**: Extract all images from DOC/PPT files at once
//
// # Architecture
//
// - `blip`: Core BLIP (Binary Large Image or Picture) record parsing
// - `bse`: BLIP Store Entry (BSE) metadata parsing
// - `emf`: Enhanced Metafile (EMF) format support
// - `wmf`: Windows Metafile (WMF) format support
// - `pict`: Macintosh PICT format support
// - `extractor`: High-level image extraction from Office files
// - `svg`: SVG conversion utilities
//
// # Quick Start: Extract Images from Office Files
//
// ```no_run
// use litchi::images::{extract_images_from_doc, extract_images_from_ppt};
//
// // Extract from Word document
// let images = extract_images_from_doc("document.doc")?;
// for (i, img) in images.iter().enumerate() {
///     let png = img.to_png(Some(800), None)?;
///     std::fs::write(format!("image_{}.png", i), png)?;
/// }
///
/// // Extract from PowerPoint presentation
/// let images = extract_images_from_ppt("presentation.ppt")?;
/// for img in images {
///     let filename = img.suggested_filename();
///     let png = img.to_png(None, None)?;
///     std::fs::write(filename, png)?;
/// }
/// # Ok::<(), litchi::common::error::Error>(())
/// ```
///
/// # Example: Converting a BLIP record
///
/// ```no_run
/// use litchi::images::blip::Blip;
/// use litchi::images::convert_blip_to_png;
///
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
///
/// # Supported Formats
///
/// | Format | Type | Compression | Notes |
/// |--------|------|-------------|-------|
/// | EMF | Metafile | Optional | Enhanced Metafile |
/// | WMF | Metafile | Optional | Windows Metafile |
/// | PICT | Metafile | Optional | Macintosh PICT |
/// | PNG | Bitmap | N/A | Direct passthrough |
/// | JPEG | Bitmap | N/A | Direct passthrough |
/// | DIB | Bitmap | N/A | Device Independent Bitmap |
/// | TIFF | Bitmap | N/A | Tagged Image File Format |
///
/// # Performance
///
/// - **Zero-copy parsing**: Uses `Cow<'data, [u8]>` to avoid unnecessary allocations
/// - **Fast scanning**: O(n) single-pass search for BLIPs in data streams
/// - **Efficient conversion**: Direct raster rendering for metafiles
///
/// # Error Handling
///
/// The module gracefully handles common real-world issues:
/// - Incorrect compression flags (attempts raw data if decompression fails)
/// - Misaligned BLIP records (manual parsing instead of zerocopy)
/// - Invalid length fields (lenient bounds checking)
/// - Corrupted metafile data (returns error but continues extraction)
pub mod blip;
pub mod bse;
pub mod emf;
pub mod extractor;
pub mod pict;
pub mod svg;
pub mod svg_utils;
pub mod wmf;

use crate::common::error::Result;
pub use blip::{BitmapBlip, Blip, BlipType, MetafileBlip, RecordHeader};
pub use bse::{BlipStore, BlipStoreEntry};
pub use extractor::{ExtractedImage, ImageExtractor};
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
pub fn convert_blip_to_format<'data>(
    blip: &Blip<'data>,
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
            let img = image::load_from_memory(&bitmap.picture_data[..]).map_err(|e| {
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
pub fn convert_blip_to_png<'data>(
    blip: &Blip<'data>,
    width: Option<u32>,
    height: Option<u32>,
) -> Result<Vec<u8>> {
    convert_blip_to_format(blip, ImageFormat::Png, width, height)
}

/// Convert a BLIP record to JPEG format
pub fn convert_blip_to_jpeg<'data>(
    blip: &Blip<'data>,
    width: Option<u32>,
    height: Option<u32>,
) -> Result<Vec<u8>> {
    convert_blip_to_format(blip, ImageFormat::Jpeg, width, height)
}

/// Convert a BLIP record to WebP format
pub fn convert_blip_to_webp<'data>(
    blip: &Blip<'data>,
    width: Option<u32>,
    height: Option<u32>,
) -> Result<Vec<u8>> {
    convert_blip_to_format(blip, ImageFormat::WebP, width, height)
}

// User-friendly convenience functions for extracting images from Office files

/// Extract all images from a PPT presentation file
///
/// # Arguments
/// * `path` - Path to the .ppt file
///
/// # Returns
/// Vector of extracted images with metadata
///
/// # Example
/// ```no_run
/// use litchi::images::extract_images_from_ppt;
///
/// let images = extract_images_from_ppt("presentation.ppt")?;
/// for (i, img) in images.iter().enumerate() {
///     let png_data = img.to_png(None, None)?;
///     std::fs::write(format!("image_{}.png", i), png_data)?;
/// }
/// # Ok::<(), Box<dyn std::error::Error>>(())
/// ```
#[cfg(feature = "ole")]
pub fn extract_images_from_ppt<P: AsRef<std::path::Path>>(
    path: P,
) -> Result<Vec<ExtractedImage<'static>>> {
    use crate::ole::OleFile;
    use std::fs::File;

    let file = File::open(path).map_err(crate::common::error::Error::Io)?;
    let mut ole = OleFile::open(file).map_err(|e| {
        crate::common::error::Error::ParseError(format!("Failed to open OLE file: {}", e))
    })?;

    ImageExtractor::extract_from_ppt(&mut ole)
}

/// Extract all images from a DOC document file
///
/// # Arguments
/// * `path` - Path to the .doc file
///
/// # Returns
/// Vector of extracted images with metadata
///
/// # Example
/// ```no_run
/// use litchi::images::extract_images_from_doc;
///
/// let images = extract_images_from_doc("document.doc")?;
/// for img in images {
///     let filename = img.suggested_filename();
///     let data = img.decompressed_data()?;
///     std::fs::write(filename, &*data)?;
/// }
/// # Ok::<(), Box<dyn std::error::Error>>(())
/// ```
#[cfg(feature = "ole")]
pub fn extract_images_from_doc<P: AsRef<std::path::Path>>(
    path: P,
) -> Result<Vec<ExtractedImage<'static>>> {
    use crate::ole::OleFile;
    use std::fs::File;

    let file = File::open(path).map_err(crate::common::error::Error::Io)?;
    let mut ole = OleFile::open(file).map_err(|e| {
        crate::common::error::Error::ParseError(format!("Failed to open OLE file: {}", e))
    })?;

    ImageExtractor::extract_from_doc(&mut ole)
}

/// Extract images from raw Escher drawing data
///
/// This is a lower-level function useful when you already have Escher data
/// extracted from a document.
///
/// # Arguments
/// * `escher_data` - Raw Escher drawing layer data
///
/// # Returns
/// Vector of extracted images
pub fn extract_images_from_escher(escher_data: &[u8]) -> Result<Vec<ExtractedImage<'static>>> {
    ImageExtractor::extract_blips(escher_data)
}

/// Parse a BLIP store (BSE index) from Escher data
///
/// The BLIP store provides metadata about all images in a document.
///
/// # Arguments
/// * `escher_data` - Raw Escher drawing layer data
///
/// # Returns
/// BlipStore with all BSE entries
pub fn parse_blip_store(escher_data: &[u8]) -> Result<BlipStore<'_>> {
    ImageExtractor::extract_blip_store(escher_data)
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
