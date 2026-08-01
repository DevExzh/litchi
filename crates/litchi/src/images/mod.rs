//! Office image grammar, codecs, and legacy-container discovery.
//!
//! The layers remain explicit:
//!
//! - [`art`] owns borrowed OfficeArt BLIP and BStore grammar;
//! - [`codec`] owns bounded decoding, rendering, and conversion;
//! - [`ImageExtractor`] maps DOC and PPT storage topology onto those views.
//!
//! ```no_run
//! use litchi::images::{Options, doc, ppt};
//!
//! # fn main() -> Result<(), litchi::Error> {
//! for image in doc("document.doc")? {
//!     std::fs::write(image.suggested_filename(), image.extract(Options::default())?)?;
//! }
//! for image in ppt("presentation.ppt")? {
//!     std::fs::write(image.suggested_filename(), image.extract(Options::default())?)?;
//! }
//! # Ok(())
//! # }
//! ```

/// Borrowed OfficeArt image records, checked identifiers, stores, and writers.
pub mod art {
    pub use litchi_odraw::image::*;
}

/// Bounded native-image decoding, rendering, and format conversion.
pub mod codec {
    pub use litchi_imgconv::*;
}

pub use codec::{Limits, Options, convert, decode_data, to_jpeg, to_png, to_svg, to_webp};

#[cfg(feature = "ole")]
pub use litchi_ole::extractor::{ExtractedImage, ImageExtractor};

#[cfg(feature = "ole")]
use litchi_core::error::Result;

/// Extracts images from a legacy PowerPoint presentation.
///
/// Returned records own their OfficeArt framing because the compound file is
/// closed before this function returns. Native image data is not decoded.
#[cfg(feature = "ole")]
pub fn ppt(path: impl AsRef<std::path::Path>) -> Result<Vec<ExtractedImage<'static>>> {
    use litchi_ole::OleFile;
    use std::fs::File;

    let file = File::open(path).map_err(litchi_core::error::Error::Io)?;
    let mut ole = OleFile::open(file).map_err(|error| {
        litchi_core::error::Error::ParseError(format!("failed to open OLE file: {error}"))
    })?;
    ImageExtractor::from_ppt(&mut ole)
}

/// Extracts images from a legacy Word document.
///
/// Returned records own their OfficeArt framing because the compound file is
/// closed before this function returns. Native image data is not decoded.
#[cfg(feature = "ole")]
pub fn doc(path: impl AsRef<std::path::Path>) -> Result<Vec<ExtractedImage<'static>>> {
    use litchi_ole::OleFile;
    use std::fs::File;

    let file = File::open(path).map_err(litchi_core::error::Error::Io)?;
    let mut ole = OleFile::open(file).map_err(|error| {
        litchi_core::error::Error::ParseError(format!("failed to open OLE file: {error}"))
    })?;
    ImageExtractor::from_doc(&mut ole)
}

/// Discovers borrowed images in an OfficeArt record sequence.
#[cfg(feature = "ole")]
pub fn escher(data: &[u8]) -> Result<Vec<ExtractedImage<'_>>> {
    ImageExtractor::blips(data)
}

/// Finds the unique borrowed BStore in an OfficeArt drawing sequence.
#[cfg(feature = "ole")]
pub fn store(data: &[u8]) -> Result<Option<art::Store<'_>>> {
    ImageExtractor::store(data)
}
