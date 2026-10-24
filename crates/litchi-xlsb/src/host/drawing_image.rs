#![allow(
    clippy::wildcard_enum_match_arm,
    reason = "legacy module confines an intentional opaque or future-variant fallback to this codec boundary"
)]

//! Typed worksheet image payloads used by XLSB Drawings parts.

use crate::builder::chart::Anchor;
use crate::package::error::{Error, Result};
use litchi_opc::constants::content_type as ct;
use std::sync::Arc;

const PNG_SIGNATURE: &[u8] = b"\x89PNG\r\n\x1a\n";
const JPEG_SIGNATURE: &[u8] = &[0xFF, 0xD8, 0xFF];
const GIF87A_SIGNATURE: &[u8] = b"GIF87a";
const GIF89A_SIGNATURE: &[u8] = b"GIF89a";
const BMP_SIGNATURE: &[u8] = b"BM";
const TIFF_LE_SIGNATURE: &[u8] = b"II\x2A\0";
const TIFF_BE_SIGNATURE: &[u8] = b"MM\0\x2A";
const WDP_LE_SIGNATURE: &[u8] = b"II\xBC\x01";
const WDP_BE_SIGNATURE: &[u8] = b"MM\x01\xBC";
const PLACEABLE_WMF_SIGNATURE: &[u8] = &[0xD7, 0xCD, 0xC6, 0x9A];
const STANDARD_WMF_SIGNATURE: &[u8] = &[0x01, 0x00, 0x09, 0x00];
const STANDARD_MEMORY_WMF_SIGNATURE: &[u8] = &[0x02, 0x00, 0x09, 0x00];
const EMF_SIGNATURE_OFFSET: usize = 40;
const EMF_SIGNATURE: &[u8] = b" EMF";
const MAX_SVG_XML_DEPTH: usize = 256;

/// Maximum encoded bytes accepted for one worksheet image.
pub const MAX_XLSB_WORKSHEET_IMAGE_BYTES: usize = 32 * 1024 * 1024;
/// Maximum number of resolved or authored images in one worksheet drawing.
pub const MAX_XLSB_WORKSHEET_IMAGES: usize = 4_096;
/// Maximum combined encoded image bytes in one worksheet drawing.
pub const MAX_XLSB_WORKSHEET_IMAGE_TOTAL_BYTES: usize = 256 * 1024 * 1024;
/// Maximum UTF-8 bytes accepted for picture alternative text.
pub const MAX_XLSB_IMAGE_DESCRIPTION_BYTES: usize = 32 * 1024;

/// Image formats that can be embedded in an XLSB worksheet drawing.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum ImageFormat {
    /// Windows bitmap.
    Bmp,
    /// Graphics Interchange Format.
    Gif,
    /// Joint Photographic Experts Group image.
    Jpeg,
    /// Portable Network Graphics image.
    Png,
    /// Scalable Vector Graphics image.
    Svg,
    /// Tagged Image File Format image.
    Tiff,
    /// Enhanced Metafile vector image.
    Emf,
    /// Windows Metafile vector image.
    Wmf,
    /// JPEG XR / Windows Media Photo image.
    Wdp,
}

impl ImageFormat {
    /// Canonical package filename extension.
    pub const fn extension(self) -> &'static str {
        match self {
            Self::Bmp => "bmp",
            Self::Gif => "gif",
            Self::Jpeg => "jpeg",
            Self::Png => "png",
            Self::Svg => "svg",
            Self::Tiff => "tiff",
            Self::Emf => "emf",
            Self::Wmf => "wmf",
            Self::Wdp => "wdp",
        }
    }

    /// OPC content type for the image part.
    pub const fn content_type(self) -> &'static str {
        match self {
            Self::Bmp => ct::BMP,
            Self::Gif => ct::GIF,
            Self::Jpeg => ct::JPEG,
            Self::Png => ct::PNG,
            Self::Svg => "image/svg+xml",
            Self::Tiff => ct::TIFF,
            Self::Emf => ct::X_EMF,
            Self::Wmf => ct::X_WMF,
            Self::Wdp => ct::MS_PHOTO,
        }
    }

    /// Map an OPC image content type to a supported format.
    pub fn from_content_type(content_type: &str) -> Option<Self> {
        match content_type {
            ct::BMP => Some(Self::Bmp),
            ct::GIF => Some(Self::Gif),
            ct::JPEG => Some(Self::Jpeg),
            ct::PNG => Some(Self::Png),
            "image/svg+xml" => Some(Self::Svg),
            ct::TIFF => Some(Self::Tiff),
            ct::X_EMF | "image/emf" => Some(Self::Emf),
            ct::X_WMF | "image/wmf" => Some(Self::Wmf),
            ct::MS_PHOTO => Some(Self::Wdp),
            _ => None,
        }
    }

    fn has_matching_signature(self, data: &[u8]) -> bool {
        match self {
            Self::Bmp => data.starts_with(BMP_SIGNATURE),
            Self::Gif => data.starts_with(GIF87A_SIGNATURE) || data.starts_with(GIF89A_SIGNATURE),
            Self::Jpeg => data.starts_with(JPEG_SIGNATURE),
            Self::Png => data.starts_with(PNG_SIGNATURE),
            Self::Svg => validate_svg(data),
            Self::Tiff => {
                data.starts_with(TIFF_LE_SIGNATURE) || data.starts_with(TIFF_BE_SIGNATURE)
            },
            Self::Emf => {
                data.get(EMF_SIGNATURE_OFFSET..EMF_SIGNATURE_OFFSET + EMF_SIGNATURE.len())
                    == Some(EMF_SIGNATURE)
            },
            Self::Wmf => {
                data.starts_with(PLACEABLE_WMF_SIGNATURE)
                    || data.starts_with(STANDARD_WMF_SIGNATURE)
                    || data.starts_with(STANDARD_MEMORY_WMF_SIGNATURE)
            },
            Self::Wdp => data.starts_with(WDP_LE_SIGNATURE) || data.starts_with(WDP_BE_SIGNATURE),
        }
    }

    pub(crate) fn validate_payload(self, data: &[u8]) -> Result<()> {
        if data.is_empty() {
            return Err(Error::InvalidFormula(
                "worksheet image payload cannot be empty",
            ));
        }
        if data.len() > MAX_XLSB_WORKSHEET_IMAGE_BYTES {
            return Err(Error::InvalidLength {
                expected: MAX_XLSB_WORKSHEET_IMAGE_BYTES,
                found: data.len(),
            });
        }
        if !self.has_matching_signature(data) {
            return Err(Error::Unrecognized {
                typ: "worksheet image payload",
                val: "bytes do not match declared format",
            });
        }
        Ok(())
    }
}

/// A bounded dimension of worksheet-image authoring.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Encoded bytes of one image ([`MAX_XLSB_WORKSHEET_IMAGE_BYTES`]).
    ImageBytes,
    /// UTF-8 bytes of one image description
    /// ([`MAX_XLSB_IMAGE_DESCRIPTION_BYTES`]).
    DescriptionBytes,
}

impl std::fmt::Display for LimitKind {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.write_str(match self {
            Self::ImageBytes => "worksheet image bytes",
            Self::DescriptionBytes => "worksheet image description bytes",
        })
    }
}

impl From<LimitKind> for crate::package::error::LimitKind {
    fn from(kind: LimitKind) -> Self {
        Self::Image(kind)
    }
}

/// Caller limits for worksheet-image authoring.
///
/// The defaults are the format ceilings; a setter accepts any smaller value
/// and refuses a larger one with [`Error::InvalidLimit`].
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Limits {
    image_bytes: usize,
    description_bytes: usize,
}

impl Limits {
    /// The format ceilings.
    pub const DEFAULT: Self = Self {
        image_bytes: MAX_XLSB_WORKSHEET_IMAGE_BYTES,
        description_bytes: MAX_XLSB_IMAGE_DESCRIPTION_BYTES,
    };

    /// Bound the encoded bytes of one image.
    ///
    /// # Errors
    ///
    /// Returns [`Error::InvalidLimit`] above [`MAX_XLSB_WORKSHEET_IMAGE_BYTES`].
    pub fn max_image_bytes(mut self, value: usize) -> Result<Self> {
        self.image_bytes = checked(LimitKind::ImageBytes, value, MAX_XLSB_WORKSHEET_IMAGE_BYTES)?;
        Ok(self)
    }

    /// Bound the UTF-8 bytes of one image description.
    ///
    /// # Errors
    ///
    /// Returns [`Error::InvalidLimit`] above [`MAX_XLSB_IMAGE_DESCRIPTION_BYTES`].
    pub fn max_description_bytes(mut self, value: usize) -> Result<Self> {
        self.description_bytes = checked(
            LimitKind::DescriptionBytes,
            value,
            MAX_XLSB_IMAGE_DESCRIPTION_BYTES,
        )?;
        Ok(self)
    }

    /// The configured bound of one image's encoded bytes.
    #[must_use]
    pub const fn image_bytes(&self) -> usize {
        self.image_bytes
    }

    /// The configured bound of one image description's UTF-8 bytes.
    #[must_use]
    pub const fn description_bytes(&self) -> usize {
        self.description_bytes
    }
}

impl Default for Limits {
    fn default() -> Self {
        Self::DEFAULT
    }
}

fn checked(kind: LimitKind, value: usize, maximum: usize) -> Result<usize> {
    if value > maximum {
        return Err(Error::InvalidLimit {
            kind: kind.into(),
            value,
            maximum,
        });
    }
    Ok(value)
}

/// One image and its two-cell worksheet anchor.
#[derive(Debug, Clone)]
pub struct Image {
    data: Arc<[u8]>,
    format: ImageFormat,
    anchor: Anchor,
    description: Option<String>,
    limits: Limits,
}

impl Image {
    /// Create and validate an embedded worksheet image under `limits`.
    ///
    /// The encoded length is admitted against `limits` before the bytes are
    /// taken into shared ownership; `limits` is retained for later
    /// [`Self::with_description`] checks.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Limit`] for bytes above the image ceiling, and the
    /// anchor, signature and format errors of the payload.
    pub fn new<D>(data: D, format: ImageFormat, anchor: Anchor, limits: &Limits) -> Result<Self>
    where
        D: AsRef<[u8]> + Into<Arc<[u8]>>,
    {
        let length = data.as_ref().len();
        if length > limits.image_bytes {
            return Err(Error::limit(
                LimitKind::ImageBytes,
                length,
                limits.image_bytes,
            ));
        }
        let image = Self {
            data: data.into(),
            format,
            anchor,
            description: None,
            limits: *limits,
        };
        image.validate()?;
        Ok(image)
    }

    /// Attach alternative text used by assistive technology.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Limit`] above the retained description ceiling and an
    /// encoding error for characters XML cannot carry.
    pub fn with_description(mut self, description: impl Into<String>) -> Result<Self> {
        let description = description.into();
        if description.len() > self.limits.description_bytes {
            return Err(Error::limit(
                LimitKind::DescriptionBytes,
                description.len(),
                self.limits.description_bytes,
            ));
        }
        self.description = Some(description);
        self.validate()?;
        Ok(self)
    }

    /// The limits this image was admitted under.
    #[must_use]
    pub const fn limits(&self) -> &Limits {
        &self.limits
    }

    /// Encoded image bytes.
    pub fn data(&self) -> &[u8] {
        &self.data
    }

    /// Declared encoded image format.
    pub const fn format(&self) -> ImageFormat {
        self.format
    }

    /// Two-cell image anchor, using zero-based rows and columns and EMU offsets.
    pub const fn anchor(&self) -> &Anchor {
        &self.anchor
    }

    /// Optional image alternative text.
    pub fn description(&self) -> Option<&str> {
        self.description.as_deref()
    }

    pub(crate) fn validate(&self) -> Result<()> {
        crate::chart::anchor::validate(&self.anchor)?;
        self.format.validate_payload(&self.data)?;
        if let Some(description) = &self.description {
            if description.len() > MAX_XLSB_IMAGE_DESCRIPTION_BYTES {
                return Err(Error::InvalidLength {
                    expected: MAX_XLSB_IMAGE_DESCRIPTION_BYTES,
                    found: description.len(),
                });
            }
            if description.chars().any(|character| {
                !matches!(
                    character as u32,
                    0x9 | 0xA | 0xD | 0x20..=0xD7FF | 0xE000..=0xFFFD | 0x10000..=0x10FFFF
                )
            }) {
                return Err(Error::Encoding(
                    "worksheet image description contains an invalid XML character",
                ));
            }
        }
        Ok(())
    }
}

fn validate_svg(data: &[u8]) -> bool {
    use quick_xml::events::Event;
    use quick_xml::name::{Namespace, ResolveResult};
    use quick_xml::reader::NsReader;

    const SVG_NAMESPACE: Namespace<'_> = Namespace(b"http://www.w3.org/2000/svg");
    let mut reader = NsReader::from_reader(data);
    reader.config_mut().trim_text(true);
    let mut saw_svg_root = false;
    let mut root_closed = false;
    let mut depth = 0usize;
    loop {
        let (namespace, event) = match reader.read_resolved_event() {
            Ok(value) => value,
            Err(_) => return false,
        };
        match event {
            Event::Start(element) if !saw_svg_root => {
                if element.local_name().as_ref() != b"svg"
                    || namespace != ResolveResult::Bound(SVG_NAMESPACE)
                {
                    return false;
                }
                saw_svg_root = true;
                depth = 1;
            },
            Event::Empty(element) if !saw_svg_root => {
                if element.local_name().as_ref() != b"svg"
                    || namespace != ResolveResult::Bound(SVG_NAMESPACE)
                {
                    return false;
                }
                saw_svg_root = true;
                root_closed = true;
            },
            Event::Start(_) if saw_svg_root && !root_closed => {
                let Some(next_depth) = depth.checked_add(1) else {
                    return false;
                };
                if next_depth > MAX_SVG_XML_DEPTH {
                    return false;
                }
                depth = next_depth;
            },
            Event::Empty(_) if root_closed => return false,
            Event::End(_) if !root_closed => {
                let Some(next_depth) = depth.checked_sub(1) else {
                    return false;
                };
                depth = next_depth;
                root_closed = depth == 0;
            },
            Event::Start(_) | Event::End(_) if root_closed => return false,
            Event::Text(text) if (!saw_svg_root || root_closed) && !text.is_empty() => {
                return false;
            },
            Event::CData(text) if (!saw_svg_root || root_closed) && !text.is_empty() => {
                return false;
            },
            Event::Eof => return saw_svg_root && root_closed,
            _ => {},
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn anchor() -> Anchor {
        Anchor::new(0, 0, 1, 1)
    }

    #[test]
    fn accepts_every_typed_image_signature() {
        let mut emf = vec![0; EMF_SIGNATURE_OFFSET];
        emf.extend_from_slice(EMF_SIGNATURE);
        for (format, data) in [
            (ImageFormat::Bmp, BMP_SIGNATURE),
            (ImageFormat::Gif, GIF89A_SIGNATURE),
            (ImageFormat::Jpeg, JPEG_SIGNATURE),
            (ImageFormat::Png, PNG_SIGNATURE),
            (
                ImageFormat::Svg,
                br#"<svg xmlns="http://www.w3.org/2000/svg"/>"#,
            ),
            (ImageFormat::Tiff, TIFF_LE_SIGNATURE),
            (ImageFormat::Wmf, PLACEABLE_WMF_SIGNATURE),
            (ImageFormat::Wdp, WDP_LE_SIGNATURE),
        ] {
            assert_eq!(
                ImageFormat::from_content_type(format.content_type()),
                Some(format)
            );
            Image::new(data.to_vec(), format, anchor(), &Limits::DEFAULT).unwrap();
        }
        assert_eq!(
            ImageFormat::from_content_type(ImageFormat::Emf.content_type()),
            Some(ImageFormat::Emf)
        );
        Image::new(emf, ImageFormat::Emf, anchor(), &Limits::DEFAULT).unwrap();
    }

    #[test]
    fn svg_requires_well_formed_namespaced_svg_xml() {
        for invalid in [
            b"<svg/>".as_slice(),
            br#"<svg xmlns="urn:not-svg"/>"#,
            br#"<svg xmlns="http://www.w3.org/2000/svg">"#,
            br#"<html xmlns="http://www.w3.org/2000/svg"/>"#,
        ] {
            assert!(
                Image::new(
                    invalid.to_vec(),
                    ImageFormat::Svg,
                    anchor(),
                    &Limits::DEFAULT,
                )
                .is_err()
            );
        }
    }

    #[test]
    fn image_limits_admit_bytes_before_ownership_and_bound_descriptions() {
        let png = PNG_SIGNATURE.to_vec();
        let exact = Limits::DEFAULT.max_image_bytes(png.len()).unwrap();
        assert!(Image::new(png.clone(), ImageFormat::Png, anchor(), &exact).is_ok());
        let tight = Limits::DEFAULT.max_image_bytes(png.len() - 1).unwrap();
        assert!(matches!(
            Image::new(png.clone(), ImageFormat::Png, anchor(), &tight),
            Err(Error::Limit {
                kind: crate::package::LimitKind::Image(LimitKind::ImageBytes),
                ..
            })
        ));
        assert!(matches!(
            Limits::DEFAULT.max_image_bytes(MAX_XLSB_WORKSHEET_IMAGE_BYTES + 1),
            Err(Error::InvalidLimit { .. })
        ));
        let short = Limits::DEFAULT.max_description_bytes(3).unwrap();
        let image = Image::new(png, ImageFormat::Png, anchor(), &short).unwrap();
        assert!(image.clone().with_description("abc").is_ok());
        assert!(matches!(
            image.with_description("abcd"),
            Err(Error::Limit {
                kind: crate::package::LimitKind::Image(LimitKind::DescriptionBytes),
                ..
            })
        ));
    }
}
