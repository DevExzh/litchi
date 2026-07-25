//! Authoring of inert `draw:frame` image and text-box frames for ODT.
//!
//! Frames are emitted as DrawingML-shaped ODF markup: a `draw:frame` carrying
//! identity (`draw:name`), anchor (`text:anchor-type`), and geometry
//! (`svg:width`/`svg:height`), wrapping either a package-linked `draw:image`
//! or a `draw:text-box` story. Image payloads are stored verbatim under
//! `Pictures/` — nothing is re-encoded, fetched, or laid out.

use crate::elements::element::{Element, ElementBase};
use litchi_core::{Error, Result};

/// Maximum accepted image payload size.
const MAX_IMAGE_BYTES: usize = 64 * 1024 * 1024;
/// Maximum text-box story size.
const MAX_TEXT_BOX_BYTES: usize = 1024 * 1024;
/// Maximum frame name length.
const MAX_FRAME_NAME_CHARS: usize = 256;
/// Directory that stores authored picture payloads.
const PICTURES_DIRECTORY: &str = "Pictures/";

fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}

/// Anchor behavior of a `draw:frame` (`text:anchor-type`).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum OdfFrameAnchor {
    /// Anchored to a paragraph, floating beside it (`paragraph`).
    Paragraph,
    /// Anchored to a character position (`char`).
    Char,
    /// Flows inline like a character (`as-char`).
    AsChar,
    /// Anchored to the page (`page`).
    Page,
    /// Anchored inside another frame (`frame`).
    Frame,
}

impl OdfFrameAnchor {
    /// The ODF attribute spelling.
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Paragraph => "paragraph",
            Self::Char => "char",
            Self::AsChar => "as-char",
            Self::Page => "page",
            Self::Frame => "frame",
        }
    }

    /// Parse an ODF `text:anchor-type` value.
    pub fn parse(value: &str) -> Option<Self> {
        match value {
            "paragraph" => Some(Self::Paragraph),
            "char" => Some(Self::Char),
            "as-char" => Some(Self::AsChar),
            "page" => Some(Self::Page),
            "frame" => Some(Self::Frame),
            _ => None,
        }
    }
}

/// An ODF length value such as `5cm` (`svg:width`/`svg:height`).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct OdfLength(String);

impl OdfLength {
    /// A length in centimeters.
    pub fn centimeters(value: f64) -> Self {
        Self::format(value, "cm")
    }

    /// A length in millimeters.
    pub fn millimeters(value: f64) -> Self {
        Self::format(value, "mm")
    }

    /// A length in inches.
    pub fn inches(value: f64) -> Self {
        Self::format(value, "in")
    }

    /// A length in points.
    pub fn points(value: f64) -> Self {
        Self::format(value, "pt")
    }

    /// A length in picas.
    pub fn picas(value: f64) -> Self {
        Self::format(value, "pc")
    }

    /// A length in pixels.
    pub fn pixels(value: f64) -> Self {
        Self::format(value, "px")
    }

    fn format(value: f64, unit: &str) -> Self {
        debug_assert!(value.is_finite() && value >= 0.0);
        let number = if value.fract() == 0.0 {
            format!("{}", value as i64)
        } else {
            format!("{value:.2}")
        };
        Self(format!("{number}{unit}"))
    }

    /// Parse and validate an ODF length (`<number><unit>`).
    pub fn parse(value: &str) -> Result<Self> {
        let split = value
            .find(|character: char| character.is_ascii_alphabetic() || character == '%')
            .ok_or_else(|| invalid(format!("ODF length '{value}' lacks a unit")))?;
        let (number, unit) = value.split_at(split);
        if number.is_empty()
            || !number
                .chars()
                .all(|character| character.is_ascii_digit() || matches!(character, '.' | '-' | '+'))
            || number.parse::<f64>().is_err()
        {
            return Err(invalid(format!("invalid ODF length number '{number}'")));
        }
        if !matches!(unit, "cm" | "mm" | "in" | "pt" | "pc" | "px" | "em" | "%") {
            return Err(invalid(format!("unsupported ODF length unit '{unit}'")));
        }
        Ok(Self(value.to_string()))
    }

    /// The ODF attribute spelling.
    pub fn as_str(&self) -> &str {
        &self.0
    }
}

/// A sniffed raster image format accepted for frame authoring.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum OdfImageFormat {
    /// Portable Network Graphics.
    Png,
    /// JPEG.
    Jpeg,
    /// Graphics Interchange Format.
    Gif,
}

impl OdfImageFormat {
    /// File extension used for the package part.
    pub const fn extension(self) -> &'static str {
        match self {
            Self::Png => "png",
            Self::Jpeg => "jpg",
            Self::Gif => "gif",
        }
    }

    /// IANA media type used for the package manifest entry.
    pub const fn media_type(self) -> &'static str {
        match self {
            Self::Png => "image/png",
            Self::Jpeg => "image/jpeg",
            Self::Gif => "image/gif",
        }
    }

    /// Detect the format from magic bytes.
    pub fn sniff(bytes: &[u8]) -> Option<Self> {
        const PNG_MAGIC: &[u8] = b"\x89PNG\r\n\x1a\n";
        const JPEG_MAGIC: &[u8] = b"\xff\xd8\xff";
        const GIF87_MAGIC: &[u8] = b"GIF87a";
        const GIF89_MAGIC: &[u8] = b"GIF89a";
        if bytes.starts_with(PNG_MAGIC) {
            Some(Self::Png)
        } else if bytes.starts_with(JPEG_MAGIC) {
            Some(Self::Jpeg)
        } else if bytes.starts_with(GIF87_MAGIC) || bytes.starts_with(GIF89_MAGIC) {
            Some(Self::Gif)
        } else {
            None
        }
    }
}

/// An authored picture payload awaiting package insertion.
#[derive(Debug, Clone)]
pub(crate) struct PendingImage {
    /// Package part path (`Pictures/imageN.<ext>`).
    pub path: String,
    /// Verbatim payload bytes.
    pub bytes: Vec<u8>,
}

/// Validate a frame name (non-empty, XML-attribute safe, bounded).
fn validate_frame_name(name: &str) -> Result<()> {
    if name.is_empty() || name.chars().count() > MAX_FRAME_NAME_CHARS {
        return Err(invalid("ODF frame name is empty or exceeds the limit"));
    }
    if name
        .chars()
        .any(|character| matches!(character, '<' | '>' | '&' | '"' | '\''))
    {
        return Err(invalid("ODF frame name contains markup characters"));
    }
    Ok(())
}

/// Allocate a non-colliding `Pictures/imageN.<ext>` part path.
pub(crate) fn allocate_picture_path(
    extension: &str,
    mut is_taken: impl FnMut(&str) -> bool,
) -> Result<String> {
    const MAX_PICTURE_INDEX: u32 = 1_000_000;
    for index in 1..MAX_PICTURE_INDEX {
        let path = format!("{PICTURES_DIRECTORY}image{index}.{extension}");
        if !is_taken(&path) {
            return Ok(path);
        }
    }
    Err(invalid("ODF picture part namespace is exhausted"))
}

/// Build the `draw:frame` element for a package-linked image.
pub(crate) fn image_frame_element(
    name: &str,
    width: &OdfLength,
    height: &OdfLength,
    anchor: OdfFrameAnchor,
    href: &str,
) -> Result<Element> {
    validate_frame_name(name)?;
    let mut frame = frame_shell(name, width, height, anchor);
    let mut image = Element::new("draw:image");
    image.set_attribute("xlink:href", href);
    image.set_attribute("xlink:type", "simple");
    image.set_attribute("xlink:show", "embed");
    image.set_attribute("xlink:actuate", "onLoad");
    frame.add_child(image);
    Ok(frame)
}

/// Build the `draw:frame` element for a plain-text text box.
pub(crate) fn text_box_frame_element(
    name: &str,
    width: &OdfLength,
    height: &OdfLength,
    anchor: OdfFrameAnchor,
    text: &str,
) -> Result<Element> {
    validate_frame_name(name)?;
    if text.len() > MAX_TEXT_BOX_BYTES {
        return Err(invalid("ODF text-box story exceeds the size limit"));
    }
    let mut frame = frame_shell(name, width, height, anchor);
    let mut text_box = Element::new("draw:text-box");
    if text.is_empty() {
        text_box.add_child(Element::new("text:p"));
    } else {
        for line in text.split('\n') {
            let mut paragraph = Element::new("text:p");
            if !line.is_empty() {
                paragraph.set_text(line);
            }
            text_box.add_child(paragraph);
        }
    }
    frame.add_child(text_box);
    Ok(frame)
}

fn frame_shell(
    name: &str,
    width: &OdfLength,
    height: &OdfLength,
    anchor: OdfFrameAnchor,
) -> Element {
    let mut frame = Element::new("draw:frame");
    frame.set_attribute("draw:name", name);
    frame.set_attribute("text:anchor-type", anchor.as_str());
    frame.set_attribute("svg:width", width.as_str());
    frame.set_attribute("svg:height", height.as_str());
    frame
}

/// Validate an image payload and bound its size.
pub(crate) fn validate_image_payload(bytes: &[u8]) -> Result<OdfImageFormat> {
    if bytes.len() > MAX_IMAGE_BYTES {
        return Err(invalid("ODF image payload exceeds the size limit"));
    }
    OdfImageFormat::sniff(bytes)
        .ok_or_else(|| invalid("unsupported image format: PNG, JPEG, and GIF are accepted"))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn length_construction_and_parsing() {
        assert_eq!(OdfLength::centimeters(5.0).as_str(), "5cm");
        assert_eq!(OdfLength::inches(2.5).as_str(), "2.50in");
        assert_eq!(OdfLength::points(12.0).as_str(), "12pt");
        assert_eq!(OdfLength::parse("10mm").unwrap().as_str(), "10mm");
        assert_eq!(OdfLength::parse("3.25px").unwrap().as_str(), "3.25px");
        assert!(OdfLength::parse("10").is_err());
        assert!(OdfLength::parse("cm").is_err());
        assert!(OdfLength::parse("10furlongs").is_err());
        assert!(OdfLength::parse("xcm").is_err());
    }

    #[test]
    fn anchor_round_trip() {
        for anchor in [
            OdfFrameAnchor::Paragraph,
            OdfFrameAnchor::Char,
            OdfFrameAnchor::AsChar,
            OdfFrameAnchor::Page,
            OdfFrameAnchor::Frame,
        ] {
            assert_eq!(OdfFrameAnchor::parse(anchor.as_str()), Some(anchor));
        }
        assert_eq!(OdfFrameAnchor::parse("marginalia"), None);
    }

    #[test]
    fn sniffs_supported_formats_and_rejects_others() {
        assert_eq!(
            OdfImageFormat::sniff(b"\x89PNG\r\n\x1a\nrest"),
            Some(OdfImageFormat::Png)
        );
        assert_eq!(
            OdfImageFormat::sniff(b"\xff\xd8\xff\xe0rest"),
            Some(OdfImageFormat::Jpeg)
        );
        assert_eq!(
            OdfImageFormat::sniff(b"GIF89a-rest"),
            Some(OdfImageFormat::Gif)
        );
        assert_eq!(OdfImageFormat::sniff(b"BM bitmap"), None);
        assert_eq!(OdfImageFormat::sniff(b""), None);
        assert!(validate_image_payload(b"\x89PNG\r\n\x1a\nrest").is_ok());
        assert!(validate_image_payload(b"tiff?").is_err());
    }

    #[test]
    fn allocates_first_free_picture_path() {
        let path = allocate_picture_path("png", |_| false).unwrap();
        assert_eq!(path, "Pictures/image1.png");
        let path = allocate_picture_path("jpg", |candidate| candidate == "Pictures/image1.jpg")
            .unwrap();
        assert_eq!(path, "Pictures/image2.jpg");
        assert!(allocate_picture_path("png", |_| true).is_err());
    }

    #[test]
    fn image_frame_carries_identity_anchor_geometry_and_link() {
        let frame = image_frame_element(
            "Chart 1",
            &OdfLength::centimeters(10.0),
            &OdfLength::centimeters(4.0),
            OdfFrameAnchor::AsChar,
            "Pictures/image1.png",
        )
        .unwrap();
        assert_eq!(frame.get_attribute("draw:name"), Some("Chart 1"));
        assert_eq!(frame.get_attribute("text:anchor-type"), Some("as-char"));
        assert_eq!(frame.get_attribute("svg:width"), Some("10cm"));
        let image = &frame.get_children()[0];
        assert_eq!(image.tag_name(), "draw:image");
        assert_eq!(image.get_attribute("xlink:href"), Some("Pictures/image1.png"));
        let xml = frame.to_xml_string();
        assert!(xml.contains("xlink:actuate=\"onLoad\""));
    }

    #[test]
    fn text_box_frame_splits_lines_and_escapes_text() {
        let frame = text_box_frame_element(
            "Box",
            &OdfLength::inches(2.0),
            &OdfLength::inches(1.0),
            OdfFrameAnchor::Paragraph,
            "a < b\nsecond & line",
        )
        .unwrap();
        let text_box = &frame.get_children()[0];
        assert_eq!(text_box.tag_name(), "draw:text-box");
        assert_eq!(text_box.get_children().len(), 2);
        let xml = frame.to_xml_string();
        assert!(xml.contains("a &lt; b"));
        assert!(xml.contains("second &amp; line"));
        assert!(text_box_frame_element("B", &OdfLength::points(1.0), &OdfLength::points(1.0), OdfFrameAnchor::Page, "").is_ok());
    }

    #[test]
    fn rejects_bad_frame_names() {
        for name in ["", "a<b", "a\"b", "x".repeat(300).as_str()] {
            assert!(
                image_frame_element(
                    name,
                    &OdfLength::points(1.0),
                    &OdfLength::points(1.0),
                    OdfFrameAnchor::Page,
                    "Pictures/image1.png",
                )
                .is_err(),
                "accepted frame name {name:?}"
            );
        }
    }
}
