//! RTF picture/image extraction and processing.
//!
//! This module handles extraction of embedded pictures from RTF documents.
//! RTF supports several image formats:
//! - Windows Metafile (WMF)
//! - Enhanced Metafile (EMF)
//! - PNG
//! - JPEG
//! - DIB (Device Independent Bitmap)
//! - BMP

use crate::{RtfError, RtfResult};
use std::borrow::Cow;

pub(crate) const MAX_PICTURE_WRITE_BYTES: usize = 64 * 1_048_576;
/// Maximum number of scalar `OfficeArt` properties on one inline picture.
pub const MAX_PICTURE_SHAPE_PROPERTIES: usize = 4_096;
/// Maximum aggregate name/value bytes on one inline picture.
pub const MAX_PICTURE_SHAPE_PROPERTY_BYTES: usize = 1024 * 1024;

#[allow(
    clippy::module_name_repetitions,
    reason = "names mirror the RTF specification vocabulary"
)]
/// Ordered `OfficeArt` properties from a picture's starred `picprop` destination.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct PictureShapeProperties<'a> {
    /// Optional producer shape identifier (`shplid`)
    pub shape_id: Option<i32>,
    /// Ordered scalar or binary `sp` records, including optional theme metadata
    pub properties: Vec<crate::ShapeProperty<'a>>,
}

impl PictureShapeProperties<'_> {
    /// Validate the normative nonempty property list and resource bounds.
    ///
    /// # Errors
    /// Returns an error when the input is malformed or a configured limit is exceeded.
    pub fn validate(&self) -> RtfResult<()> {
        if self.properties.is_empty() || self.properties.len() > MAX_PICTURE_SHAPE_PROPERTIES {
            return Err(RtfError::MalformedDocument(
                "RTF picprop must contain a bounded nonempty property list".to_string(),
            ));
        }
        let mut bytes = 0usize;
        for property in &self.properties {
            property.validate()?;
            bytes = bytes
                .checked_add(property.name.len())
                .and_then(|size| size.checked_add(property.value.len()))
                .and_then(|size| {
                    size.checked_add(
                        property
                            .binary_value
                            .as_ref()
                            .map_or(0, |value| value.len()),
                    )
                })
                .ok_or_else(|| {
                    RtfError::MalformedDocument(
                        "RTF picture shape-property size overflow".to_string(),
                    )
                })?;
            if bytes > MAX_PICTURE_SHAPE_PROPERTY_BYTES {
                return Err(RtfError::MalformedDocument(
                    "RTF picture shape properties exceed the text safety limit".to_string(),
                ));
            }
        }
        Ok(())
    }

    fn into_owned(self) -> PictureShapeProperties<'static> {
        PictureShapeProperties {
            shape_id: self.shape_id,
            properties: self
                .properties
                .into_iter()
                .map(|property| crate::ShapeProperty {
                    name: Cow::Owned(property.name.into_owned()),
                    value: Cow::Owned(property.value.into_owned()),
                    binary_value: property
                        .binary_value
                        .map(|value| Cow::Owned(value.into_owned())),
                    theme_value: property.theme_value,
                    hyperlink: property.hyperlink.map(crate::ShapeHyperlink::into_owned),
                })
                .collect(),
        }
    }
}

#[allow(
    clippy::module_name_repetitions,
    reason = "names mirror the RTF specification vocabulary"
)]
/// Inert identity metadata attached to one RTF picture payload.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct PictureIdentity<'a> {
    /// Producer-defined signed cache tag.
    pub tag: Option<i32>,
    /// Source bitmap resolution in units per inch.
    pub units_per_inch: Option<u16>,
    /// Producer-defined 16-byte image identifier. Some producers emit an empty value.
    pub uid: Option<Cow<'a, [u8]>>,
}

impl PictureIdentity<'_> {
    ///
    /// # Errors
    /// Returns an error when the input is malformed or a configured limit is exceeded.
    pub fn validate(&self) -> RtfResult<()> {
        if self.units_per_inch == Some(0) {
            return Err(RtfError::MalformedDocument(
                "RTF blipupi must be positive".to_string(),
            ));
        }
        if self
            .uid
            .as_ref()
            .is_some_and(|uid| !uid.is_empty() && uid.len() != 16)
        {
            return Err(RtfError::MalformedDocument(
                "RTF blipuid must contain exactly 16 bytes or be empty".to_string(),
            ));
        }
        Ok(())
    }

    pub(crate) fn into_owned(self) -> PictureIdentity<'static> {
        PictureIdentity {
            tag: self.tag,
            units_per_inch: self.units_per_inch,
            uid: self.uid.map(|uid| Cow::Owned(uid.into_owned())),
        }
    }
}

/// Image type in RTF documents.
///
/// Note: This enum is specific to RTF parsing. For general image processing,
/// see the `images` module which has comprehensive format support.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ImageType {
    /// Enhanced Metafile
    Emf,
    /// Windows Metafile
    Wmf,
    /// PNG image
    Png,
    /// JPEG image
    Jpeg,
    /// DIB (Device Independent Bitmap)
    Dib,
    /// Mac PICT format
    Pict,
    /// Unknown or unsupported format
    Unknown,
}

#[allow(
    clippy::module_name_repetitions,
    reason = "names mirror the RTF specification vocabulary"
)]
/// Passive crop distances in twips from the four source-picture edges.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct PictureCrop {
    pub left: Option<i32>,
    pub right: Option<i32>,
    pub top: Option<i32>,
    pub bottom: Option<i32>,
}

#[allow(
    clippy::module_name_repetitions,
    reason = "names mirror the RTF specification vocabulary"
)]
/// Passive legacy bitmap header controls carried by a `pict` destination.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct PictureBitmapMetadata {
    /// Source selected `wbitmap` rather than `dibitmap`.
    pub windows_bitmap: bool,
    /// Source explicitly used the legacy `picbmp` flag.
    pub bitmap_source: bool,
    /// Source bits per pixel from `picbpp`.
    pub bits_per_pixel: Option<u16>,
    /// Windows bitmap bits per pixel from `wbmbitspixel`.
    pub windows_bits_per_pixel: Option<u16>,
    /// Windows bitmap plane count from `wbmplanes`.
    pub planes: Option<u16>,
    /// Bytes per source bitmap scan line from `wbmwidthbytes`.
    pub width_bytes: Option<u32>,
}

impl PictureBitmapMetadata {
    fn validate(&self, image_type: ImageType) -> RtfResult<()> {
        let present = self.windows_bitmap
            || self.bitmap_source
            || self.bits_per_pixel.is_some()
            || self.windows_bits_per_pixel.is_some()
            || self.planes.is_some()
            || self.width_bytes.is_some();
        if present && image_type != ImageType::Dib {
            return Err(RtfError::MalformedDocument(
                "RTF bitmap header controls require a DIB or Windows bitmap picture".to_string(),
            ));
        }
        for (name, value) in [
            ("picbpp", self.bits_per_pixel),
            ("wbmbitspixel", self.windows_bits_per_pixel),
            ("wbmplanes", self.planes),
        ] {
            if value == Some(0) {
                return Err(RtfError::MalformedDocument(format!(
                    "RTF {name} must be positive"
                )));
            }
        }
        if self.width_bytes == Some(0) {
            return Err(RtfError::MalformedDocument(
                "RTF wbmwidthbytes must be positive".to_string(),
            ));
        }
        if self
            .width_bytes
            .is_some_and(|value| value > i32::MAX as u32)
        {
            return Err(RtfError::MalformedDocument(
                "RTF wbmwidthbytes exceeds the signed control-word range".to_string(),
            ));
        }
        Ok(())
    }
}

/// Extracted picture from RTF document.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Picture<'a> {
    /// Image type
    pub image_type: ImageType,
    /// Image data (hex-encoded in RTF, decoded here)
    pub data: Cow<'a, [u8]>,
    /// Optional cache/source identity metadata.
    pub identity: Option<PictureIdentity<'a>>,
    /// Optional inline `OfficeArt` property destination (`picprop`)
    pub shape_properties: Option<PictureShapeProperties<'a>>,
    /// Picture width (in twips, 1/1440 inch)
    pub width: Option<i32>,
    /// Picture height (in twips)
    pub height: Option<i32>,
    /// Goal width (desired width in twips)
    pub goal_width: Option<i32>,
    /// Goal height (desired height in twips)
    pub goal_height: Option<i32>,
    /// Horizontal scaling percentage
    pub scale_x: Option<i32>,
    /// Vertical scaling percentage
    pub scale_y: Option<i32>,
    /// Whether the source supplied the `picscaled` flag.
    pub scaled: bool,
    /// Passive source crop distances.
    pub crop: PictureCrop,
    /// Passive legacy bitmap header metadata.
    pub bitmap: PictureBitmapMetadata,
}

impl<'a> Picture<'a> {
    /// Create a new picture with minimal information.
    #[inline]
    #[must_use]
    pub fn new(image_type: ImageType, data: Cow<'a, [u8]>) -> Self {
        Self {
            image_type,
            data,
            identity: None,
            shape_properties: None,
            width: None,
            height: None,
            goal_width: None,
            goal_height: None,
            scale_x: None,
            scale_y: None,
            scaled: false,
            crop: PictureCrop::default(),
            bitmap: PictureBitmapMetadata::default(),
        }
    }

    /// Get the image data as a byte slice.
    #[inline]
    #[must_use]
    pub fn data(&self) -> &[u8] {
        &self.data
    }

    ///
    /// # Errors
    /// Returns an error when the input is malformed or a configured limit is exceeded.
    pub fn validate(&self) -> RtfResult<()> {
        if self.data.is_empty() || self.data.len() > MAX_PICTURE_WRITE_BYTES {
            return Err(RtfError::MalformedDocument(
                "RTF picture payload is empty or exceeds the writing limit".to_string(),
            ));
        }
        if let Some(identity) = &self.identity {
            identity.validate()?;
        }
        if let Some(properties) = &self.shape_properties {
            properties.validate()?;
        }
        self.bitmap.validate(self.image_type)?;
        Ok(())
    }

    pub(crate) fn into_owned(self) -> Picture<'static> {
        Picture {
            image_type: self.image_type,
            data: Cow::Owned(self.data.into_owned()),
            identity: self.identity.map(PictureIdentity::into_owned),
            shape_properties: self
                .shape_properties
                .map(PictureShapeProperties::into_owned),
            width: self.width,
            height: self.height,
            goal_width: self.goal_width,
            goal_height: self.goal_height,
            scale_x: self.scale_x,
            scale_y: self.scale_y,
            scaled: self.scaled,
            crop: self.crop,
            bitmap: self.bitmap,
        }
    }

    /// Get the computed width in twips, considering scaling.
    #[inline]
    #[must_use]
    pub fn computed_width(&self) -> Option<i32> {
        self.goal_width.or(self.width).map(|w| match self.scale_x {
            Some(scale) => (w * scale) / 100,
            None => w,
        })
    }

    /// Get the computed height in twips, considering scaling.
    #[inline]
    #[must_use]
    pub fn computed_height(&self) -> Option<i32> {
        self.goal_height
            .or(self.height)
            .map(|h| match self.scale_y {
                Some(scale) => (h * scale) / 100,
                None => h,
            })
    }

    /// Convert width from twips to pixels at given DPI.
    ///
    /// # Arguments
    ///
    /// * `dpi` - Dots per inch (typically 96 for screen, 72 for print)
    #[inline]
    #[must_use]
    pub fn width_pixels(&self, dpi: u32) -> Option<u32> {
        self.computed_width()
            .map(|tw| (tw.cast_unsigned() * dpi) / 1440)
    }

    /// Convert height from twips to pixels at given DPI.
    ///
    /// # Arguments
    ///
    /// * `dpi` - Dots per inch (typically 96 for screen, 72 for print)
    #[inline]
    #[must_use]
    pub fn height_pixels(&self, dpi: u32) -> Option<u32> {
        self.computed_height()
            .map(|tw| (tw.cast_unsigned() * dpi) / 1440)
    }
}

/// Detect image type from binary signature.
///
/// # Arguments
///
/// * `data` - Binary image data
///
/// # Returns
///
/// Detected image type or Unknown
#[must_use]
pub fn detect_image_type(data: &[u8]) -> ImageType {
    if data.is_empty() {
        return ImageType::Unknown;
    }

    // Check JPEG signature (starts with FFD8)
    if data.starts_with(&[0xFF, 0xD8]) {
        return ImageType::Jpeg;
    }

    // Check PNG signature
    if data.len() >= 8 && data.starts_with(&[0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A]) {
        return ImageType::Png;
    }

    // Check EMF signature (0x01 0x00 0x00 0x00)
    if data.starts_with(&[0x01, 0x00, 0x00, 0x00])
        && data.get(40..44) == Some(&[0x20, 0x45, 0x4D, 0x46])
    {
        return ImageType::Emf;
    }

    // Check WMF signature (0xD7, 0xCD, 0xC6, 0x9A) - Aldus Placeable Metafile
    if data.starts_with(&[0xD7, 0xCD, 0xC6, 0x9A]) {
        return ImageType::Wmf;
    }

    // Check DIB/BMP signature
    if data.starts_with(&[0x42, 0x4D]) {
        // "BM"
        return ImageType::Dib;
    }

    ImageType::Unknown
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_detect_png() {
        let png_sig = vec![0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A];
        assert_eq!(detect_image_type(&png_sig), ImageType::Png);
    }

    #[test]
    fn test_detect_jpeg() {
        let jpeg_sig = vec![0xFF, 0xD8, 0xFF, 0xE0];
        assert_eq!(detect_image_type(&jpeg_sig), ImageType::Jpeg);
    }

    #[test]
    fn detects_metafile_and_bitmap_signatures() {
        let mut emf = [0_u8; 44];
        emf[..4].copy_from_slice(&[0x01, 0x00, 0x00, 0x00]);
        emf[40..].copy_from_slice(&[0x20, 0x45, 0x4D, 0x46]);
        assert_eq!(detect_image_type(&emf), ImageType::Emf);
        assert_eq!(detect_image_type(&[0xD7, 0xCD, 0xC6, 0x9A]), ImageType::Wmf);
        assert_eq!(detect_image_type(b"BM"), ImageType::Dib);
    }

    #[test]
    fn truncated_and_near_miss_emf_headers_are_unknown() {
        let mut emf = [0_u8; 44];
        emf[..4].copy_from_slice(&[0x01, 0x00, 0x00, 0x00]);
        emf[40..].copy_from_slice(&[0x20, 0x45, 0x4D, 0x46]);
        for length in 0..emf.len() {
            assert_eq!(detect_image_type(&emf[..length]), ImageType::Unknown);
        }
        emf[43] = 0;
        assert_eq!(detect_image_type(&emf), ImageType::Unknown);
    }

    #[test]
    fn test_picture_dimensions() {
        let pic = Picture {
            image_type: ImageType::Png,
            data: Cow::Borrowed(&[]),
            identity: None,
            shape_properties: None,
            width: Some(1440), // 1 inch
            height: Some(1440),
            goal_width: None,
            goal_height: None,
            scale_x: Some(200), // 200% scale
            scale_y: Some(200),
            scaled: false,
            crop: PictureCrop::default(),
            bitmap: PictureBitmapMetadata::default(),
        };

        assert_eq!(pic.computed_width(), Some(2880)); // 2 inches
        assert_eq!(pic.width_pixels(96), Some(192)); // 2 inches at 96 DPI
    }
}
