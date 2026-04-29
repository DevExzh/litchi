//! Image processing and conversion module.
//!
//! Re-exports the [`litchi-imgconv`](../../litchi_imgconv/index.html) crate
//! (pure BLIP / EMF / WMF / PICT decoders and converters) plus an
//! integration helper, [`extractor`], that bridges OLE Escher records to
//! `litchi-imgconv` types. The integration helper stays in the umbrella
//! because it depends on `crate::ole` and therefore cannot live in a leaf
//! crate.
//!
//! # Quick Start: Extract Images from Office Files
//!
//! ```no_run
//! use litchi::images::{extract_images_from_doc, extract_images_from_ppt};
//!
//! # fn main() -> Result<(), litchi::Error> {
//! // Extract from Word document
//! let images = extract_images_from_doc("document.doc")?;
//! for (i, img) in images.iter().enumerate() {
//!     let png = img.to_png(Some(800), None)?;
//!     std::fs::write(format!("image_{}.png", i), png)?;
//! }
//!
//! // Extract from PowerPoint presentation
//! let images = extract_images_from_ppt("presentation.ppt")?;
//! for img in images {
//!     let filename = img.suggested_filename();
//!     let png = img.to_png(None, None)?;
//!     std::fs::write(filename, png)?;
//! }
//! # Ok::<(), litchi::Error>(())
//! # }
//! ```
//!
//! # Example: Converting a BLIP record
//!
//! ```no_run
//! use litchi::images::blip::Blip;
//! use litchi::images::convert_blip_to_png;
//!
//! let blip_data = vec![/* BLIP record bytes */];
//! let blip = Blip::parse(&blip_data)?;
//! let png_bytes = convert_blip_to_png(&blip, Some(800), None)?;
//! # Ok::<(), litchi::Error>(())
//! ```

pub use litchi_imgconv::*;

// Integration glue between `litchi_ole::escher` and `litchi_imgconv` lives
// in the `litchi-ole` crate (relocated from the umbrella in P4c) because it
// reaches into private `litchi-ole` Escher types. The umbrella re-exports
// the public surface here so callers using `litchi::images::ImageExtractor`
// keep resolving.
#[cfg(all(feature = "ole", feature = "imgconv"))]
pub use litchi_ole::extractor::{ExtractedImage, ImageExtractor};

#[cfg(all(feature = "ole", feature = "imgconv"))]
use litchi_core::error::Result;

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
/// # Ok::<(), litchi::Error>(())
/// ```
#[cfg(all(feature = "ole", feature = "imgconv"))]
pub fn extract_images_from_ppt<P: AsRef<std::path::Path>>(
    path: P,
) -> Result<Vec<ExtractedImage<'static>>> {
    use crate::ole::OleFile;
    use std::fs::File;

    let file = File::open(path).map_err(litchi_core::error::Error::Io)?;
    let mut ole = OleFile::open(file).map_err(|e| {
        litchi_core::error::Error::ParseError(format!("Failed to open OLE file: {}", e))
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
/// # Ok::<(), litchi::Error>(())
/// ```
#[cfg(all(feature = "ole", feature = "imgconv"))]
pub fn extract_images_from_doc<P: AsRef<std::path::Path>>(
    path: P,
) -> Result<Vec<ExtractedImage<'static>>> {
    use crate::ole::OleFile;
    use std::fs::File;

    let file = File::open(path).map_err(litchi_core::error::Error::Io)?;
    let mut ole = OleFile::open(file).map_err(|e| {
        litchi_core::error::Error::ParseError(format!("Failed to open OLE file: {}", e))
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
#[cfg(all(feature = "ole", feature = "imgconv"))]
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
#[cfg(all(feature = "ole", feature = "imgconv"))]
pub fn parse_blip_store(escher_data: &[u8]) -> Result<BlipStore<'_>> {
    ImageExtractor::extract_blip_store(escher_data)
}
