use litchi_core::{Error, Result};

const MAX_FRAME_NAME_CHARS: usize = 256;
const MAX_TEXT_BOX_BYTES: usize = 1024 * 1024;

/// Anchor behavior of an ODF `draw:frame` (`text:anchor-type`).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Anchor {
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

impl Anchor {
    /// The ODF attribute spelling.
    #[must_use]
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
    #[must_use]
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
pub struct Length(String);

impl Length {
    /// A length in centimeters.
    #[must_use]
    pub fn centimeters(value: f64) -> Self {
        Self::format(value, "cm")
    }

    /// A length in millimeters.
    #[must_use]
    pub fn millimeters(value: f64) -> Self {
        Self::format(value, "mm")
    }

    /// A length in inches.
    #[must_use]
    pub fn inches(value: f64) -> Self {
        Self::format(value, "in")
    }

    /// A length in points.
    #[must_use]
    pub fn points(value: f64) -> Self {
        Self::format(value, "pt")
    }

    /// A length in picas.
    #[must_use]
    pub fn picas(value: f64) -> Self {
        Self::format(value, "pc")
    }

    /// A length in pixels.
    #[must_use]
    pub fn pixels(value: f64) -> Self {
        Self::format(value, "px")
    }

    fn format(value: f64, unit: &str) -> Self {
        debug_assert!(value.is_finite() && value >= 0.0);
        let number = if value.fract() == 0.0 {
            format!("{value:.0}")
        } else {
            let mut number = format!("{value:.2}");
            while number.ends_with('0') {
                number.pop();
            }
            if number.ends_with('.') {
                number.pop();
            }
            number
        };
        Self(format!("{number}{unit}"))
    }

    /// Parse and validate an ODF length (`<number><unit>`).
    ///
    /// # Errors
    ///
    /// Returns an error if the lexical value has no supported unit or has an
    /// invalid numeric component.
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
    #[must_use]
    pub fn as_str(&self) -> &str {
        &self.0
    }
}

/// Validated geometry and placement for one authored ODF frame.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Frame {
    name: String,
    width: Length,
    height: Length,
    anchor: Anchor,
}

impl Frame {
    /// Construct a frame shell shared by image, text-box, and future drawing
    /// hosts.
    ///
    /// # Errors
    ///
    /// Returns an error if `name` is empty, too long, or contains markup.
    pub fn new(name: &str, width: Length, height: Length, anchor: Anchor) -> Result<Self> {
        validate_name(name)?;
        Ok(Self {
            name: name.to_string(),
            width,
            height,
            anchor,
        })
    }

    /// The authored `draw:name`, if any.
    #[must_use]
    pub fn name(&self) -> &str {
        &self.name
    }

    /// The authored `svg:width`.
    #[must_use]
    pub fn width(&self) -> &Length {
        &self.width
    }

    /// The authored `svg:height`.
    #[must_use]
    pub fn height(&self) -> &Length {
        &self.height
    }

    /// The authored `text:anchor-type`.
    #[must_use]
    pub const fn anchor(&self) -> Anchor {
        self.anchor
    }
}

fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}

/// Validate a frame name before it is emitted as an XML attribute.
fn validate_name(name: &str) -> Result<()> {
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

/// Validate a plain-text `draw:text-box` story against the common authoring
/// allocation bound.
///
/// # Errors
///
/// Returns an error if `text` exceeds the authoring size limit.
pub fn validate_text_box(text: &str) -> Result<()> {
    if text.len() > MAX_TEXT_BOX_BYTES {
        return Err(invalid("ODF text-box story exceeds the size limit"));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn length_construction_and_parsing() -> Result<()> {
        assert_eq!(Length::centimeters(5.0).as_str(), "5cm");
        assert_eq!(Length::inches(2.5).as_str(), "2.5in");
        assert_eq!(Length::points(12.0).as_str(), "12pt");
        assert_eq!(Length::parse("10mm")?.as_str(), "10mm");
        assert_eq!(Length::parse("3.25px")?.as_str(), "3.25px");
        assert!(Length::parse("10").is_err());
        assert!(Length::parse("cm").is_err());
        assert!(Length::parse("10furlongs").is_err());
        assert!(Length::parse("xcm").is_err());
        Ok(())
    }

    #[test]
    fn anchor_round_trip() {
        for anchor in [
            Anchor::Paragraph,
            Anchor::Char,
            Anchor::AsChar,
            Anchor::Page,
            Anchor::Frame,
        ] {
            assert_eq!(Anchor::parse(anchor.as_str()), Some(anchor));
        }
        assert_eq!(Anchor::parse("marginalia"), None);
    }

    #[test]
    fn frame_validates_identity_and_retains_geometry() -> Result<()> {
        let frame = Frame::new(
            "Photo",
            Length::centimeters(10.0),
            Length::centimeters(4.0),
            Anchor::AsChar,
        )?;
        assert_eq!(frame.name(), "Photo");
        assert_eq!(frame.width().as_str(), "10cm");
        assert_eq!(frame.height().as_str(), "4cm");
        assert_eq!(frame.anchor(), Anchor::AsChar);

        for name in ["", "a<b", "a\"b", "x".repeat(300).as_str()] {
            assert!(
                Frame::new(name, Length::points(1.0), Length::points(1.0), Anchor::Page,).is_err(),
                "accepted frame name {name:?}"
            );
        }
        Ok(())
    }

    #[test]
    fn text_box_bound_is_shared() {
        assert!(validate_text_box("plain text").is_ok());
    }
}
