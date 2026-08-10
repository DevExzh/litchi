//! Typed ODT header/footer models and lexical invariants.

use litchi_core::Result;

pub use litchi_odf_common::style::master::Region as MasterRegion;
pub use litchi_odf_common::style::master::content::{
    Block, Column, Field, FieldKind, Inline, SenderKind,
};
pub use litchi_odf_common::style::master::{Child, ChildKind, Kind, Master};

use super::{MAX_VALUE, bad};

const PHYSICAL_UNITS: [&str; 6] = ["cm", "mm", "in", "pt", "pc", "px"];

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Length(String);

impl Length {
    pub fn new(value: impl Into<String>) -> Result<Self> {
        let value = value.into();
        if value.is_empty() || value.len() > MAX_VALUE || value.chars().any(char::is_control) {
            return Err(bad("invalid header/footer length"));
        }
        if value == "0" || value == "+0" || value == "-0" {
            return Ok(Self(value));
        }
        let unit = PHYSICAL_UNITS
            .into_iter()
            .find(|unit| value.ends_with(unit))
            .ok_or_else(|| bad("header/footer length requires a physical unit"))?;
        let number = &value[..value.len() - unit.len()];
        if number.is_empty()
            || number.contains(['e', 'E'])
            || number
                .parse::<f64>()
                .map_or(true, |number| !number.is_finite())
        {
            return Err(bad("invalid header/footer length number"));
        }
        Ok(Self(value))
    }

    pub fn as_str(&self) -> &str {
        &self.0
    }

    pub(super) fn nonnegative(value: String, name: &str) -> Result<Self> {
        let parsed = Self::new(value)?;
        if parsed.0.starts_with('-') && parsed.0 != "-0" {
            return Err(bad(format!("{name} must be nonnegative")));
        }
        Ok(parsed)
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Color {
    Transparent,
    Rgb(u8, u8, u8),
}

impl Color {
    pub(super) fn parse(value: &str, transparent: bool) -> Result<Self> {
        if transparent && value == "transparent" {
            return Ok(Self::Transparent);
        }
        if value.len() != 7 || !value.starts_with('#') {
            return Err(bad("color must be #RRGGBB"));
        }
        let rgb = u32::from_str_radix(&value[1..], 16).map_err(|_error| bad("invalid color"))?;
        Ok(Self::Rgb((rgb >> 16) as u8, (rgb >> 8) as u8, rgb as u8))
    }

    pub(super) fn xml(self) -> String {
        match self {
            Self::Transparent => "transparent".into(),
            Self::Rgb(r, g, b) => format!("#{r:02X}{g:02X}{b:02X}"),
        }
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum BorderStyle {
    Hidden,
    Dotted,
    Dashed,
    Solid,
    Double,
    Groove,
    Ridge,
    Inset,
    Outset,
}

impl BorderStyle {
    pub(super) fn parse(value: &str) -> Result<Self> {
        match value {
            "hidden" => Ok(Self::Hidden),
            "dotted" => Ok(Self::Dotted),
            "dashed" => Ok(Self::Dashed),
            "solid" => Ok(Self::Solid),
            "double" => Ok(Self::Double),
            "groove" => Ok(Self::Groove),
            "ridge" => Ok(Self::Ridge),
            "inset" => Ok(Self::Inset),
            "outset" => Ok(Self::Outset),
            _ => Err(bad("invalid header/footer border style")),
        }
    }

    pub(super) fn xml(self) -> &'static str {
        match self {
            Self::Hidden => "hidden",
            Self::Dotted => "dotted",
            Self::Dashed => "dashed",
            Self::Solid => "solid",
            Self::Double => "double",
            Self::Groove => "groove",
            Self::Ridge => "ridge",
            Self::Inset => "inset",
            Self::Outset => "outset",
        }
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub enum Border {
    None,
    Line {
        width: Length,
        style: BorderStyle,
        color: Color,
    },
}

impl Border {
    pub(super) fn parse(value: &str) -> Result<Self> {
        if value == "none" {
            return Ok(Self::None);
        }
        let parts: Vec<_> = value.split_ascii_whitespace().collect();
        if parts.len() != 3 {
            return Err(bad("border must be none or width style color"));
        }
        Ok(Self::Line {
            width: Length::nonnegative(parts[0].into(), "border width")?,
            style: BorderStyle::parse(parts[1])?,
            color: Color::parse(parts[2], false)?,
        })
    }

    pub(super) fn xml(&self) -> String {
        match self {
            Self::None => "none".into(),
            Self::Line {
                width,
                style,
                color,
            } => format!("{} {} {}", width.as_str(), style.xml(), color.xml()),
        }
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct BorderLineWidth {
    pub inner: Length,
    pub spacing: Length,
    pub outer: Length,
}

impl BorderLineWidth {
    pub(super) fn parse(value: &str) -> Result<Self> {
        let parts: Vec<_> = value.split_ascii_whitespace().collect();
        if parts.len() != 3 {
            return Err(bad("border-line-width requires three lengths"));
        }
        Ok(Self {
            inner: Length::nonnegative(parts[0].into(), "inner border width")?,
            spacing: Length::nonnegative(parts[1].into(), "border spacing")?,
            outer: Length::nonnegative(parts[2].into(), "outer border width")?,
        })
    }

    pub(super) fn xml(&self) -> String {
        format!(
            "{} {} {}",
            self.inner.as_str(),
            self.spacing.as_str(),
            self.outer.as_str()
        )
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub enum Shadow {
    None,
    Drop {
        color: Color,
        offset_x: Length,
        offset_y: Length,
    },
}

impl Shadow {
    pub(super) fn parse(value: &str) -> Result<Self> {
        if value == "none" {
            return Ok(Self::None);
        }
        let parts: Vec<_> = value.split_ascii_whitespace().collect();
        if parts.len() != 3 {
            return Err(bad("shadow must be none or color x-offset y-offset"));
        }
        Ok(Self::Drop {
            color: Color::parse(parts[0], false)?,
            offset_x: Length::new(parts[1])?,
            offset_y: Length::new(parts[2])?,
        })
    }

    pub(super) fn xml(&self) -> String {
        match self {
            Self::None => "none".into(),
            Self::Drop {
                color,
                offset_x,
                offset_y,
            } => format!(
                "{} {} {}",
                color.xml(),
                offset_x.as_str(),
                offset_y.as_str()
            ),
        }
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Edges<T> {
    pub all: Option<T>,
    pub left: Option<T>,
    pub right: Option<T>,
    pub top: Option<T>,
    pub bottom: Option<T>,
}

impl<T> Default for Edges<T> {
    fn default() -> Self {
        Self {
            all: None,
            left: None,
            right: None,
            top: None,
            bottom: None,
        }
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Region {
    Header,
    Footer,
}

impl Region {
    pub(super) fn wrapper(self) -> &'static str {
        match self {
            Self::Header => "header-style",
            Self::Footer => "footer-style",
        }
    }
}

#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct StyleProperties {
    pub height: Option<Length>,
    pub min_height: Option<Length>,
    pub margins: Edges<Length>,
    pub borders: Edges<Border>,
    pub border_line_widths: Edges<BorderLineWidth>,
    pub padding: Edges<Length>,
    pub background_color: Option<Color>,
    pub shadow: Option<Shadow>,
    pub dynamic_spacing: Option<bool>,
    pub background_image: Option<crate::SectionBackgroundImage>,
}

impl StyleProperties {
    pub fn validate(&self) -> Result<()> {
        if let Some(image) = &self.background_image {
            image.validate()?;
        }
        Ok(())
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Properties {
    pub page_layout_name: String,
    pub region: Region,
    pub properties: StyleProperties,
}
