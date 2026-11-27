//! Office image grammar, codecs, and legacy-container discovery.
//!
//! The layers remain explicit:
//!
//! - [`art`](crate::images::art) owns borrowed OfficeArt BLIP and BStore grammar;
//! - [`codec`](crate::images::codec) owns bounded decoding, rendering, and conversion;
//! - concrete DOC and PPT packages map their storage topology onto those views.
//!
//! ```no_run
//! # fn main() -> Result<(), litchi::common::Error> {
//! #[cfg(feature = "doc")]
//! {
//!     use litchi::images::{Convert, Options, doc};
//!     for image in doc("document.doc")? {
//!         std::fs::write(image.out_name(), image.extract(Options::default())?)?;
//!     }
//! }
//! #[cfg(feature = "ppt")]
//! {
//!     use litchi::images::{Convert, Options, ppt};
//!     for image in ppt("presentation.ppt")? {
//!         std::fs::write(image.out_name(), image.extract(Options::default())?)?;
//!     }
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
    let package = litchi_ppt::Package::open(path).map_err(Error::from)?;
    let presentation = package.presentation().map_err(Error::from)?;
    presentation
        .images()
        .map_err(Error::from)?
        .into_iter()
        .map(|image| image.into_owned().map_err(office_art))
        .collect()
}

/// Most native image bytes [`doc`] returns in total.
///
/// Every returned picture is copied out of the document, which is closed
/// before [`doc`] returns.
#[cfg(feature = "doc")]
pub const MAX_DOC_IMAGE_BYTES: u64 = 1 << 30;

/// Extracts images from a legacy Word document.
///
/// Returned records own their OfficeArt framing because the compound file is
/// closed before this function returns. Native image data is not decoded.
///
/// Each inline picture (one Data-stream PICF) is returned once, and so is each
/// picture of the drawing's image store that floating pictures show (one
/// OfficeArt shape, including each picture inside an anchored group), however
/// many shapes show it. Picture frames that store no picture, such as linked
/// pictures saved without their data, are skipped. The document is read with
/// [`litchi_doc::Leniency::TolerateStylesheetDefects`].
///
/// # Errors
///
/// Returns an error when the document cannot be read, when it holds more
/// pictures than the default [`art::Limits::max_files`], or when their image
/// data exceeds [`MAX_DOC_IMAGE_BYTES`] in total.
#[cfg(feature = "doc")]
pub fn doc(path: impl AsRef<std::path::Path>) -> Result<Vec<File<'static>>> {
    // Style names only label styles; a duplicated one, which Word opens
    // without complaint, cannot change where a picture is stored.
    let options = litchi_doc::OpenOptions::default()
        .leniency(litchi_doc::Leniency::TolerateStylesheetDefects);
    let package = litchi_doc::Package::open_with_options(path, options).map_err(Error::from)?;
    let max_files = usize::try_from(art::Limits::default().max_files).unwrap_or(usize::MAX);
    doc_images(&package.document(), max_files, MAX_DOC_IMAGE_BYTES)
}

/// The pictures of `document`, failing past `max_files` pictures or
/// `max_bytes` bytes of image data.
#[cfg(feature = "doc")]
fn doc_images(
    document: &litchi_doc::Document,
    max_files: usize,
    max_bytes: u64,
) -> Result<Vec<File<'static>>> {
    let mut inline_seen = std::collections::HashSet::new();
    let mut shapes_seen = std::collections::HashSet::new();
    let mut entries_seen = std::collections::HashSet::new();
    let mut retained = 0u64;
    let mut images = Vec::new();

    for paragraph in document.paragraphs().map_err(Error::from)? {
        for run in paragraph.runs().map_err(Error::from)? {
            let pictures = match run.image() {
                Some(image) => vec![*image],
                None => anchored_pictures(document, &run)?,
            };
            for image in &pictures {
                // A floating picture's CPicLocation is only a legacy hint that
                // other pictures can share; its shape identifies it.
                let first = match image.drawing_shape_id() {
                    Some(shape_id) => shapes_seen.insert(shape_id),
                    None => inline_seen.insert(image.pic_offset()),
                };
                if !first {
                    continue;
                }
                let file = match document.image_data(image) {
                    Ok(file) => file,
                    Err(litchi_doc::ImageError::NoPicture) => continue,
                    Err(error) => return Err(Error::ParseError(error.to_string())),
                };
                // A floating picture's file is indexed by its image-store
                // entry, which other shapes can show as well.
                if image.drawing_shape_id().is_some() && !entries_seen.insert(file.index()) {
                    continue;
                }
                if images.len() >= max_files {
                    return Err(office_art(litchi_odraw::Error::ImageLimitExceeded {
                        limit: art::LimitKind::Files,
                        observed: u64::try_from(images.len() + 1).unwrap_or(u64::MAX),
                        maximum: u64::try_from(max_files).unwrap_or(u64::MAX),
                    }));
                }
                let size = file.data().map_err(office_art)?.len();
                retained = retained.saturating_add(u64::try_from(size).unwrap_or(u64::MAX));
                if retained > max_bytes {
                    return Err(Error::ParseError(format!(
                        "document pictures hold more than {max_bytes} bytes of image data"
                    )));
                }
                let index = images.len();
                images.push(file.with_index(index).into_owned().map_err(office_art)?);
            }
        }
    }

    Ok(images)
}

/// The floating pictures of a run that `Run::image` does not describe: a run
/// of several anchor characters, or an anchor of a group of pictures.
#[cfg(feature = "doc")]
fn anchored_pictures(
    document: &litchi_doc::Document,
    run: &litchi_doc::Run,
) -> Result<Vec<litchi_doc::Image>> {
    let properties = run.properties();
    if !properties.is_spec || properties.is_obj || properties.is_ole2 || properties.is_data {
        return Ok(Vec::new());
    }
    let (Some(range), Ok(text)) = (run.cp_range(), run.text()) else {
        return Ok(Vec::new());
    };
    let mut pictures = Vec::new();
    for (offset, unit) in text.encode_utf16().enumerate() {
        let cp = u32::try_from(offset)
            .ok()
            .and_then(|offset| range.start().checked_add(offset));
        if let (0x0008, Some(cp)) = (unit, cp) {
            pictures.extend(document.floating_images_at(cp).map_err(Error::from)?);
        }
    }
    Ok(pictures)
}

/// Discovers borrowed images in an OfficeArt record sequence.
pub fn scan(data: &[u8], limits: art::Limits) -> Result<Vec<File<'_>>> {
    art::scan(data, limits).map_err(office_art)
}

/// Finds the unique borrowed BStore in an OfficeArt drawing sequence.
pub fn store(data: &[u8], limits: art::Limits) -> Result<Option<art::Store<'_>>> {
    art::store(data, limits).map_err(office_art)
}

#[cfg(all(test, feature = "doc"))]
mod tests {
    use super::*;

    fn floating_pictures(max_files: usize, max_bytes: u64) -> Result<Vec<File<'static>>> {
        let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/ole/doc/FloatingPictures.doc");
        let options = litchi_doc::OpenOptions::default()
            .leniency(litchi_doc::Leniency::TolerateStylesheetDefects);
        let package = litchi_doc::Package::open_with_options(path, options).map_err(Error::from)?;
        doc_images(&package.document(), max_files, max_bytes)
    }

    #[test]
    fn doc_images_bound_the_pictures_and_bytes_they_return() {
        // Seventeen pictures: sixteen inline ones, and one image-store entry
        // that four floating shapes show.
        let all = floating_pictures(17, u64::MAX).unwrap();
        assert_eq!(all.len(), 17);
        let bytes = all
            .iter()
            .map(|file| file.data().unwrap().len() as u64)
            .sum::<u64>();
        assert!(floating_pictures(16, u64::MAX).is_err());
        assert!(floating_pictures(17, bytes - 1).is_err());
        assert!(floating_pictures(17, bytes).is_ok());
    }
}
