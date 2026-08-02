//! Office image grammar, codecs, and legacy-container discovery.
//!
//! The layers remain explicit:
//!
//! - [`art`](crate::images::art) owns borrowed OfficeArt BLIP and BStore grammar;
//! - [`codec`](crate::images::codec) owns bounded decoding, rendering, and conversion;
//! - concrete DOC and PPT packages map their storage topology onto those views.
//!
//! ```no_run
//! use litchi::images::{Convert, Options, doc, ppt};
//!
//! # fn main() -> Result<(), litchi::Error> {
//! for image in doc("document.doc")? {
//!     std::fs::write(image.out_name(), image.extract(Options::default())?)?;
//! }
//! for image in ppt("presentation.ppt")? {
//!     std::fs::write(image.out_name(), image.extract(Options::default())?)?;
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

pub use art::File;
pub use codec::{Convert, Limits, Options, convert, decode_data, to_jpeg, to_png, to_svg, to_webp};

use litchi_core::error::{Error, Result};

fn office_art(error: litchi_odraw::Error) -> Error {
    Error::ParseError(error.to_string())
}

/// Extracts images from a legacy PowerPoint presentation.
///
/// Returned records own their OfficeArt framing because the compound file is
/// closed before this function returns. Native image data is not decoded.
#[cfg(feature = "ppt")]
pub fn ppt(path: impl AsRef<std::path::Path>) -> Result<Vec<File<'static>>> {
    let mut package = litchi_ppt::Package::open(path).map_err(Error::from)?;
    let presentation = package.presentation().map_err(Error::from)?;
    presentation
        .images()
        .map_err(Error::from)?
        .into_iter()
        .map(|image| image.into_owned().map_err(office_art))
        .collect()
}

/// Extracts images from a legacy Word document.
///
/// Returned records own their OfficeArt framing because the compound file is
/// closed before this function returns. Native image data is not decoded.
#[cfg(feature = "doc")]
pub fn doc(path: impl AsRef<std::path::Path>) -> Result<Vec<File<'static>>> {
    let mut package = litchi_doc::Package::open(path).map_err(Error::from)?;
    let document = package.document().map_err(Error::from)?;
    let mut seen = std::collections::HashSet::new();
    let mut images = Vec::new();

    for paragraph in document.paragraphs().map_err(Error::from)? {
        for run in paragraph.runs().map_err(Error::from)? {
            let Some(image) = run.image() else {
                continue;
            };
            if !seen.insert(image.pic_offset()) {
                continue;
            }
            let file = document
                .image_data(image)
                .map_err(|error| Error::ParseError(error.to_string()))?;
            let index = images.len();
            images.push(file.with_index(index).into_owned().map_err(office_art)?);
        }
    }

    Ok(images)
}

/// Discovers borrowed images in an OfficeArt record sequence.
pub fn scan(data: &[u8]) -> Result<Vec<File<'_>>> {
    art::scan(data).map_err(office_art)
}

/// Finds the unique borrowed BStore in an OfficeArt drawing sequence.
pub fn store(data: &[u8]) -> Result<Option<art::Store<'_>>> {
    art::store(data).map_err(office_art)
}
