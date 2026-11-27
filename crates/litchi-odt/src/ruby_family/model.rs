//! Typed semantic models for ODF ruby styles and inline annotations.

use litchi_core::{Error, Result, xml::escape_xml};
use std::collections::HashSet;

use super::{MAX_BASE, MAX_STYLES, MAX_VALUE};

pub(super) fn bad(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}
pub(super) fn validate_text(value: &str, context: &str, empty: bool) -> Result<()> {
    if (!empty && value.is_empty()) || value.len() > MAX_VALUE || value.chars().any(|c| matches!(c, '\0'..='\u{8}' | '\u{b}' | '\u{c}' | '\u{e}'..='\u{1f}' | '\u{fffe}' | '\u{ffff}')) {
        return Err(bad(format!("invalid {context}")));
    }
    Ok(())
}
fn ncname_start(c: char) -> bool {
    c == '_' || c.is_alphabetic()
}
fn ncname_continue(c: char) -> bool {
    ncname_start(c)
        || c.is_alphanumeric()
        || matches!(c, '-' | '.' | '\u{b7}' | '\u{300}'..='\u{36f}' | '\u{203f}'..='\u{2040}')
}
pub(super) fn validate_style_name(value: &str, context: &str) -> Result<()> {
    validate_text(value, context, false)?;
    let mut chars = value.chars();
    if !chars.next().is_some_and(ncname_start) || !chars.all(ncname_continue) {
        return Err(bad(format!("{context} is not an NCName")));
    }
    Ok(())
}
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Position {
    Above,
    Below,
}
impl Position {
    pub const ALL: [Self; 2] = [Self::Above, Self::Below];
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Above => "above",
            Self::Below => "below",
        }
    }
    pub(super) fn parse(value: &str) -> Result<Self> {
        match value {
            "above" => Ok(Self::Above),
            "below" => Ok(Self::Below),
            _ => Err(bad("invalid style:ruby-position")),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Alignment {
    Left,
    Center,
    Right,
    DistributeLetter,
    DistributeSpace,
}
impl Alignment {
    pub const ALL: [Self; 5] = [
        Self::Left,
        Self::Center,
        Self::Right,
        Self::DistributeLetter,
        Self::DistributeSpace,
    ];
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Left => "left",
            Self::Center => "center",
            Self::Right => "right",
            Self::DistributeLetter => "distribute-letter",
            Self::DistributeSpace => "distribute-space",
        }
    }
    pub(super) fn parse(value: &str) -> Result<Self> {
        match value {
            "left" => Ok(Self::Left),
            "center" => Ok(Self::Center),
            "right" => Ok(Self::Right),
            "distribute-letter" => Ok(Self::DistributeLetter),
            "distribute-space" => Ok(Self::DistributeSpace),
            _ => Err(bad("invalid style:ruby-align")),
        }
    }
}

#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Properties {
    pub position: Option<Position>,
    pub alignment: Option<Alignment>,
}
impl Properties {
    pub fn to_xml_fragment(&self) -> String {
        let mut xml = String::from(
            r#"<style:ruby-properties xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0""#,
        );
        if let Some(value) = self.position {
            xml.push_str(&format!(r#" style:ruby-position="{}""#, value.as_str()));
        }
        if let Some(value) = self.alignment {
            xml.push_str(&format!(r#" style:ruby-align="{}""#, value.as_str()));
        }
        xml.push_str("/>");
        xml
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Style {
    pub name: String,
    pub display_name: Option<String>,
    pub parent_style_name: Option<String>,
    pub properties: Option<Properties>,
}
impl Style {
    pub fn new(name: impl Into<String>, properties: Option<Properties>) -> Result<Self> {
        let value = Self {
            name: name.into(),
            display_name: None,
            parent_style_name: None,
            properties,
        };
        value.validate()?;
        Ok(value)
    }
    pub fn validate(&self) -> Result<()> {
        validate_style_name(&self.name, "ruby style name")?;
        if let Some(value) = &self.display_name {
            validate_text(value, "ruby style display name", true)?;
        }
        if let Some(value) = &self.parent_style_name {
            validate_style_name(value, "ruby parent style name")?;
        }
        Ok(())
    }
    pub fn to_xml_fragment(&self) -> Result<String> {
        self.validate()?;
        let mut xml = format!(
            r#"<style:style xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" style:name="{}" style:family="ruby""#,
            escape_xml(&self.name)
        );
        if let Some(value) = &self.display_name {
            xml.push_str(&format!(r#" style:display-name="{}""#, escape_xml(value)));
        }
        if let Some(value) = &self.parent_style_name {
            xml.push_str(&format!(
                r#" style:parent-style-name="{}""#,
                escape_xml(value)
            ));
        }
        if let Some(value) = &self.properties {
            xml.push('>');
            xml.push_str(&value.to_xml_fragment());
            xml.push_str("</style:style>");
        } else {
            xml.push_str("/>");
        }
        Ok(xml)
    }
}

#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Styles {
    pub styles: Vec<Style>,
}
impl Styles {
    pub fn get(&self, name: &str) -> Option<&Style> {
        self.styles.iter().find(|style| style.name == name)
    }
    pub fn validate(&self) -> Result<()> {
        if self.styles.len() > MAX_STYLES {
            return Err(bad("too many ruby styles"));
        }
        let mut names = HashSet::new();
        for style in &self.styles {
            style.validate()?;
            if !names.insert(style.name.as_str()) {
                return Err(bad("duplicate ruby style name"));
            }
        }
        Ok(())
    }
    pub fn to_xml(&self) -> Result<String> {
        self.validate()?;
        let mut xml = String::from(
            r#"<office:styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0">"#,
        );
        for style in &self.styles {
            xml.push_str(&style.to_xml_fragment()?);
        }
        xml.push_str("</office:styles>");
        Ok(xml)
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Base {
    pub(super) xml: String,
}
impl Base {
    pub fn from_text(text: &str) -> Result<Self> {
        validate_text(text, "ruby base text", true)?;
        if text.len() > MAX_BASE {
            return Err(bad("ruby base is too large"));
        }
        Ok(Self {
            xml: escape_xml(text),
        })
    }
    pub fn from_xml_fragment(fragment: &str) -> Result<Self> {
        if fragment.len() > MAX_BASE {
            return Err(bad("ruby base is too large"));
        }
        let ruby = Annotation::new(
            None,
            Self {
                xml: fragment.to_owned(),
            },
            "",
            None,
        )?;
        Annotation::from_xml_fragment(&ruby.to_xml_fragment()?).map(|value| value.base)
    }
    pub fn xml(&self) -> &str {
        &self.xml
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Annotation {
    pub style_name: Option<String>,
    pub base: Base,
    pub text: String,
    pub text_style_name: Option<String>,
}
impl Annotation {
    pub fn new(
        style_name: Option<String>,
        base: Base,
        text: impl Into<String>,
        text_style_name: Option<String>,
    ) -> Result<Self> {
        let value = Self {
            style_name,
            base,
            text: text.into(),
            text_style_name,
        };
        value.validate()?;
        Ok(value)
    }
    pub fn validate(&self) -> Result<()> {
        if let Some(value) = &self.style_name {
            validate_style_name(value, "ruby style reference")?;
        }
        if let Some(value) = &self.text_style_name {
            validate_style_name(value, "ruby text style reference")?;
        }
        validate_text(&self.text, "ruby pronunciation", true)?;
        if self.base.xml.len() > MAX_BASE {
            return Err(bad("ruby base is too large"));
        }
        Ok(())
    }
    pub fn to_xml_fragment(&self) -> Result<String> {
        self.validate()?;
        let mut xml = String::from(
            r#"<text:ruby xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:dr3d="urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0" xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0""#,
        );
        if let Some(value) = &self.style_name {
            xml.push_str(&format!(r#" text:style-name="{}""#, escape_xml(value)));
        }
        xml.push_str("><text:ruby-base>");
        xml.push_str(&self.base.xml);
        xml.push_str("</text:ruby-base><text:ruby-text");
        if let Some(value) = &self.text_style_name {
            xml.push_str(&format!(r#" text:style-name="{}""#, escape_xml(value)));
        }
        xml.push('>');
        xml.push_str(&escape_xml(&self.text));
        xml.push_str("</text:ruby-text></text:ruby>");
        Ok(xml)
    }
    pub fn from_xml_fragment(fragment: &str) -> Result<Self> {
        let xml = format!(
            r#"<text:p xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">{fragment}</text:p>"#
        );
        let mut entries = super::codec::parse_ruby_entries(&xml)?;
        entries.sort_by_key(|entry| entry.span.start);
        entries
            .into_iter()
            .next()
            .map(|entry| entry.value)
            .ok_or_else(|| bad("fragment contains no text:ruby"))
    }
}

#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Annotations {
    pub annotations: Vec<Annotation>,
}

/// Source span of a validated ruby annotation in an XML document.
#[derive(Clone)]
pub(super) struct Span {
    pub(super) start: usize,
    pub(super) end: usize,
}

/// Parsed annotation plus its lossless source span.
pub(super) struct Entry {
    pub(super) value: Annotation,
    pub(super) span: Span,
}
