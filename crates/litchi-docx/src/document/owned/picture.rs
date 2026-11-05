//! Inline picture insertion staged with a main-document edit (record 0251d).
//!
//! The image becomes a new media part related from the main part (a graph
//! change, 0251b), and the picture a new run holding an inline drawing that
//! references it, so one commit publishes both.

use litchi_core::Position;

use super::super::transaction::{PictureMarkup, PictureRewrite};
use super::super::{ElementPath, Error, Refusal, Result};
use super::Edit;
use crate::Package;
use crate::format::ImageFormat;

const IMAGE_RELATIONSHIP: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
const STRICT_IMAGE_RELATIONSHIP: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/image";
/// EMUs per pixel at 96 DPI.
const EMUS_PER_PIXEL: i64 = 9_525;
/// The extent of a picture whose size is neither given nor readable: one
/// inch square.
const DEFAULT_EXTENT: i64 = 914_400;
/// The largest accepted extent: 22 inches, Word's page-size limit.
const MAX_EXTENT: i64 = 22 * 914_400;

/// A picture to insert: the image bytes, an optional extent in EMUs, and
/// its name and description (alternative text).
///
/// Without an extent, the image's pixel size at 96 DPI is used for PNG,
/// JPEG, GIF and BMP images; when only one dimension is given, the other
/// follows the image's aspect ratio.
///
/// ```
/// use litchi_docx::document::NewPicture;
///
/// let picture = NewPicture::new(vec![0x89, b'P', b'N', b'G', b'\r', b'\n', 0x1a, b'\n'])
///     .with_size(Some(914_400), None)
///     .with_description("A logo");
/// assert_eq!(picture.description(), "A logo");
/// ```
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct NewPicture {
    data: Vec<u8>,
    width: Option<i64>,
    height: Option<i64>,
    name: Option<String>,
    description: String,
}

impl NewPicture {
    /// A picture of `data` (PNG, JPEG, GIF, BMP, TIFF, EMF or WMF).
    #[must_use]
    pub fn new(data: Vec<u8>) -> Self {
        Self {
            data,
            width: None,
            height: None,
            name: None,
            description: String::new(),
        }
    }

    /// The extent in EMUs; either dimension may be left to the image.
    #[must_use]
    pub const fn with_size(mut self, width: Option<i64>, height: Option<i64>) -> Self {
        self.width = width;
        self.height = height;
        self
    }

    /// The drawing object's name (`Picture N` by default).
    #[must_use]
    pub fn with_name(mut self, name: impl Into<String>) -> Self {
        self.name = Some(name.into());
        self
    }

    /// The alternative text.
    #[must_use]
    pub fn with_description(mut self, description: impl Into<String>) -> Self {
        self.description = description.into();
        self
    }

    /// The image bytes.
    #[must_use]
    pub fn data(&self) -> &[u8] {
        &self.data
    }

    /// The alternative text.
    #[must_use]
    pub fn description(&self) -> &str {
        &self.description
    }
}

impl Edit<Package> {
    /// Insert `picture` as a new run of the paragraph at `paragraph`,
    /// before its direct run `at` (after its last child for `None`), and
    /// return its position among the paragraph's inline pictures (the order
    /// of [`crate::Paragraph::images`]).
    ///
    /// The image is staged as a new part under `media/` with an `image`
    /// relationship of the main part.
    ///
    /// ```
    /// use litchi_core::Position;
    /// use litchi_docx::Package;
    /// use litchi_docx::document::{ElementPath, ElementStep, NewPicture};
    ///
    /// // A 1x1 PNG.
    /// let png = [
    ///     0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 0, 0, 13, b'I', b'H', b'D', b'R',
    ///     0, 0, 0, 1, 0, 0, 0, 1, 8, 6, 0, 0, 0, 0x1f, 0x15, 0xc4, 0x89,
    /// ];
    /// let mut builder = litchi_docx::Builder::new()?;
    /// builder.document()?.add_paragraph_with_text("Logo: ");
    /// let package = builder.build()?;
    /// let mut edit = package.edit()?;
    /// let paragraph = ElementPath::new([ElementStep::Paragraph(Position::new(0))]);
    /// edit.insert_picture_at(&paragraph, None, &NewPicture::new(png.to_vec()))?;
    /// let mut next = edit.commit()?.into_package();
    /// let reopened = Package::from_bytes(&next.to_bytes()?)?;
    /// let images = reopened.document()?.paragraphs()?[0].images()?;
    /// assert_eq!(images[0].width_emu(), 9_525);
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    ///
    /// # Errors
    ///
    /// Returns [`Refusal::InvalidPropertyValue`] for an unrecognized image
    /// or an extent outside 1 EMU to 22 inches, the refusals of
    /// [`Self::insert_run_at`], or a resource-limit error.
    pub fn insert_picture_at(
        &mut self,
        paragraph: &ElementPath,
        at: Option<Position>,
        picture: &NewPicture,
    ) -> Result<Position> {
        let invalid = || Error::Refused {
            position: 0,
            reason: Refusal::InvalidPropertyValue,
        };
        let format = ImageFormat::detect_from_bytes(&picture.data).ok_or_else(invalid)?;
        let (width, height) = extent(picture, format);
        if !(1..=MAX_EXTENT).contains(&width) || !(1..=MAX_EXTENT).contains(&height) {
            return Err(invalid());
        }
        let id = self
            .inner_mut()?
            .max_drawing_id()?
            .map_or(Some(1), |maximum| maximum.checked_add(1))
            .ok_or_else(invalid)?;
        let relationship = if self.is_strict()? {
            STRICT_IMAGE_RELATIONSHIP
        } else {
            IMAGE_RELATIONSHIP
        };
        let (_, relationship_id) = self.add_related_part(
            relationship,
            "media/",
            "image",
            &format!(".{}", format.extension()),
            format.mime_type(),
            picture.data.clone(),
        )?;
        let name = picture
            .name
            .clone()
            .unwrap_or_else(|| format!("Picture {id}"));
        self.inner_mut()?.insert_picture_markup_at(
            paragraph,
            at,
            &relationship_id,
            &PictureMarkup {
                width,
                height,
                id,
                name: &name,
                description: &picture.description,
            },
        )
    }
}

/// A change of an existing picture (record 0335): its extent, its name and
/// description (alternative text), and its image. Each part left unset
/// keeps the picture's current value.
///
/// ```
/// use litchi_docx::document::PictureChange;
///
/// let change = PictureChange::new()
///     .with_size(914_400, 457_200)
///     .with_description("The new logo");
/// assert!(!change.is_empty());
/// assert!(PictureChange::new().is_empty());
/// ```
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct PictureChange {
    extent: Option<(i64, i64)>,
    name: Option<String>,
    description: Option<String>,
    image: Option<Vec<u8>>,
}

impl PictureChange {
    /// A change that changes nothing.
    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }

    /// The new extent in EMUs (the drawing's `wp:extent` and the shape's
    /// own `a:ext`).
    #[must_use]
    pub const fn with_size(mut self, width: i64, height: i64) -> Self {
        self.extent = Some((width, height));
        self
    }

    /// The new drawing object name (`wp:docPr/@name`).
    #[must_use]
    pub fn with_name(mut self, name: impl Into<String>) -> Self {
        self.name = Some(name.into());
        self
    }

    /// The new alternative text (`wp:docPr/@descr`).
    #[must_use]
    pub fn with_description(mut self, description: impl Into<String>) -> Self {
        self.description = Some(description.into());
        self
    }

    /// The new image (PNG, JPEG, GIF, BMP, TIFF, EMF or WMF), staged as a
    /// new media part the picture then names.
    #[must_use]
    pub fn with_image(mut self, data: Vec<u8>) -> Self {
        self.image = Some(data);
        self
    }

    /// Whether the change sets nothing.
    #[must_use]
    pub const fn is_empty(&self) -> bool {
        self.extent.is_none()
            && self.name.is_none()
            && self.description.is_none()
            && self.image.is_none()
    }
}

impl Edit<Package> {
    /// Change the picture at `picture` among the paragraph's inline
    /// pictures (the order of [`crate::Paragraph::images`]) and return
    /// whether anything changed.
    ///
    /// A new image is staged as a new part under `media/` with an `image`
    /// relationship, and the picture's `a:blip` names it; the previous
    /// image part stays in the package, as a removed picture's does.
    ///
    /// ```
    /// use litchi_core::Position;
    /// use litchi_docx::Package;
    /// use litchi_docx::document::{ElementPath, ElementStep, NewPicture, PictureChange};
    ///
    /// let png = [
    ///     0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 0, 0, 13, b'I', b'H', b'D', b'R',
    ///     0, 0, 0, 1, 0, 0, 0, 1, 8, 6, 0, 0, 0, 0x1f, 0x15, 0xc4, 0x89,
    /// ];
    /// let mut builder = litchi_docx::Builder::new()?;
    /// builder.document()?.add_paragraph_with_text("Logo: ");
    /// let package = builder.build()?;
    /// let paragraph = ElementPath::new([ElementStep::Paragraph(Position::new(0))]);
    /// let mut edit = package.edit()?;
    /// edit.insert_picture_at(&paragraph, None, &NewPicture::new(png.to_vec()))?;
    /// let package = edit.commit()?.into_package();
    /// let mut edit = package.edit()?;
    /// let change = PictureChange::new().with_size(914_400, 914_400).with_description("Logo");
    /// assert!(edit.set_picture_at(&paragraph, Position::new(0), &change)?);
    /// let images = edit.commit()?.into_package().document()?.paragraphs()?[0].images()?;
    /// assert_eq!((images[0].width_emu(), images[0].description()), (914_400, "Logo"));
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    ///
    /// # Errors
    ///
    /// Returns [`Refusal::InvalidPropertyValue`] for an extent outside 1 EMU
    /// to 22 inches, a name or description holding a control character, or
    /// an unrecognized image; [`Refusal::RunNotFound`] for a missing
    /// picture; or a resource-limit error.
    pub fn set_picture_at(
        &mut self,
        paragraph: &ElementPath,
        picture: Position,
        change: &PictureChange,
    ) -> Result<bool> {
        let invalid = || Error::Refused {
            position: 0,
            reason: Refusal::InvalidPropertyValue,
        };
        if let Some((width, height)) = change.extent
            && (!(1..=MAX_EXTENT).contains(&width) || !(1..=MAX_EXTENT).contains(&height))
        {
            return Err(invalid());
        }
        let controls = |text: &str| text.chars().any(char::is_control);
        if change.name.as_deref().is_some_and(controls)
            || change.description.as_deref().is_some_and(controls)
        {
            return Err(invalid());
        }
        let format = match &change.image {
            Some(data) => Some(ImageFormat::detect_from_bytes(data).ok_or_else(invalid)?),
            None => None,
        };
        // The drawing first, so a missing picture stages no part.
        let mut changed = false;
        if change.extent.is_some() || change.name.is_some() || change.description.is_some() {
            changed = self.inner_mut()?.rewrite_picture_at(
                paragraph,
                picture,
                Some(&PictureRewrite {
                    extent: change.extent,
                    name: change.name.as_deref(),
                    description: change.description.as_deref(),
                    embed: None,
                }),
            )?;
        }
        if let (Some(data), Some(format)) = (&change.image, format) {
            // Probe the picture before the part is staged.
            self.inner_mut()?.rewrite_picture_at(
                paragraph,
                picture,
                Some(&PictureRewrite::default()),
            )?;
            let relationship = if self.is_strict()? {
                STRICT_IMAGE_RELATIONSHIP
            } else {
                IMAGE_RELATIONSHIP
            };
            let (_, relationship_id) = self.add_related_part(
                relationship,
                "media/",
                "image",
                &format!(".{}", format.extension()),
                format.mime_type(),
                data.clone(),
            )?;
            self.inner_mut()?.rewrite_picture_at(
                paragraph,
                picture,
                Some(&PictureRewrite {
                    embed: Some(&relationship_id),
                    ..PictureRewrite::default()
                }),
            )?;
            changed = true;
        }
        Ok(changed)
    }

    /// Remove the picture at `picture` among the paragraph's inline
    /// pictures: its drawing, and its run when the run holds nothing else.
    /// The image part and its relationship stay in the package,
    /// unreferenced (the policy of a removed chart).
    ///
    /// # Errors
    ///
    /// Returns [`Refusal::RunNotFound`] for a missing picture.
    pub fn remove_picture_at(&mut self, paragraph: &ElementPath, picture: Position) -> Result<()> {
        self.inner_mut()?
            .rewrite_picture_at(paragraph, picture, None)?;
        Ok(())
    }
}

/// The picture's extent in EMUs: the given dimensions, completed from the
/// image's pixel size and aspect ratio.
fn extent(picture: &NewPicture, format: ImageFormat) -> (i64, i64) {
    let pixels = pixel_size(&picture.data, format)
        .map(|(width, height)| (i64::from(width), i64::from(height)));
    match (picture.width, picture.height, pixels) {
        (Some(width), Some(height), _) => (width, height),
        (Some(width), None, Some((w, h))) if w > 0 => (width, width.saturating_mul(h) / w),
        (None, Some(height), Some((w, h))) if h > 0 => (height.saturating_mul(w) / h, height),
        (None, None, Some((w, h))) => (
            w.saturating_mul(EMUS_PER_PIXEL),
            h.saturating_mul(EMUS_PER_PIXEL),
        ),
        (width, height, _) => (
            width.unwrap_or(DEFAULT_EXTENT),
            height.unwrap_or(DEFAULT_EXTENT),
        ),
    }
}

/// The pixel size encoded in a PNG, GIF, BMP or JPEG header.
fn pixel_size(data: &[u8], format: ImageFormat) -> Option<(u32, u32)> {
    let be32 = |at: usize| Some(u32::from_be_bytes(data.get(at..at + 4)?.try_into().ok()?));
    let le16 = |at: usize| {
        Some(u32::from(u16::from_le_bytes(
            data.get(at..at + 2)?.try_into().ok()?,
        )))
    };
    let le32 = |at: usize| Some(i32::from_le_bytes(data.get(at..at + 4)?.try_into().ok()?));
    match format {
        ImageFormat::Png if data.get(12..16) == Some(b"IHDR") => Some((be32(16)?, be32(20)?)),
        ImageFormat::Gif => Some((le16(6)?, le16(8)?)),
        ImageFormat::Bmp => Some((le32(18)?.unsigned_abs(), le32(22)?.unsigned_abs())),
        ImageFormat::Jpeg => jpeg_size(data),
        _ => None,
    }
}

/// The pixel size of a JPEG's first start-of-frame segment.
fn jpeg_size(data: &[u8]) -> Option<(u32, u32)> {
    let mut at = 2usize;
    while at + 4 <= data.len() {
        if data[at] != 0xFF {
            return None;
        }
        let marker = data[at + 1];
        if marker == 0xFF {
            at += 1;
            continue;
        }
        let length = usize::from(u16::from_be_bytes([data[at + 2], data[at + 3]]));
        let frame = matches!(marker, 0xC0..=0xCF) && !matches!(marker, 0xC4 | 0xC8 | 0xCC);
        if frame {
            let height = u16::from_be_bytes([*data.get(at + 5)?, *data.get(at + 6)?]);
            let width = u16::from_be_bytes([*data.get(at + 7)?, *data.get(at + 8)?]);
            return Some((u32::from(width), u32::from(height)));
        }
        at = at.checked_add(2 + length)?;
    }
    None
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn extents_follow_the_image_and_its_aspect_ratio() {
        let mut png = vec![0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 0, 0, 13];
        png.extend_from_slice(b"IHDR");
        png.extend_from_slice(&200u32.to_be_bytes());
        png.extend_from_slice(&100u32.to_be_bytes());
        let picture = NewPicture::new(png.clone());
        assert_eq!(
            extent(&picture, ImageFormat::Png),
            (200 * EMUS_PER_PIXEL, 100 * EMUS_PER_PIXEL)
        );
        let picture = NewPicture::new(png).with_size(Some(914_400), None);
        assert_eq!(extent(&picture, ImageFormat::Png), (914_400, 457_200));
        let gif = b"GIF89a\x10\x00\x20\x00".to_vec();
        assert_eq!(pixel_size(&gif, ImageFormat::Gif), Some((16, 32)));
        let jpeg = [
            0xFF, 0xD8, 0xFF, 0xE0, 0x00, 0x04, 0x00, 0x00, 0xFF, 0xC0, 0x00, 0x0B, 0x08, 0x00,
            0x30, 0x00, 0x40, 0x01, 0x01, 0x11, 0x00,
        ];
        assert_eq!(jpeg_size(&jpeg), Some((64, 48)));
    }
}
