//! Typed ODF outline-numbering models and lexical invariants.

use std::collections::HashSet;

use litchi_core::Result;

use crate::list_label_alignment::{Alignment, Length};

use super::{MAX_OUTLINE_LEVELS, MAX_VALUE_BYTES, invalid, namespace_kind};

/// A positive XML Schema integer retained without narrowing its value.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct PositiveInteger(String);

impl PositiveInteger {
    pub fn new(value: impl Into<String>) -> Result<Self> {
        let value = value.into();
        if value.is_empty()
            || value.len() > MAX_VALUE_BYTES
            || !value.bytes().all(|byte| byte.is_ascii_digit())
            || value.bytes().all(|byte| byte == b'0')
        {
            return invalid("outline integer must be a positive XML Schema integer");
        }
        Ok(Self(value))
    }

    pub fn as_str(&self) -> &str {
        &self.0
    }
}

/// Number-format lexical value. ODF explicitly permits the empty string.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct NumberFormat(String);

impl NumberFormat {
    pub fn new(value: impl Into<String>) -> Result<Self> {
        let value = value.into();
        validate_text(&value, "style:num-format", true)?;
        Ok(Self(value))
    }

    pub fn as_str(&self) -> &str {
        &self.0
    }
}

/// Namespace-resolved formatting or producer-extension attribute.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct Attribute {
    pub(super) namespace_uri: String,
    pub(super) local_name: String,
    pub(super) value: String,
}

impl Attribute {
    pub fn new(
        namespace_uri: impl Into<String>,
        local_name: impl Into<String>,
        value: impl Into<String>,
    ) -> Result<Self> {
        let attribute = Self {
            namespace_uri: namespace_uri.into(),
            local_name: local_name.into(),
            value: value.into(),
        };
        attribute.validate()?;
        Ok(attribute)
    }

    pub fn namespace_uri(&self) -> &str {
        &self.namespace_uri
    }

    pub fn local_name(&self) -> &str {
        &self.local_name
    }

    pub fn value(&self) -> &str {
        &self.value
    }

    pub(super) fn validate(&self) -> Result<()> {
        if self.namespace_uri.is_empty()
            || self.namespace_uri.len() > MAX_VALUE_BYTES
            || !is_xml_local_name(&self.local_name)
        {
            return invalid("invalid outline extension attribute identity");
        }
        validate_text(&self.value, "outline extension attribute", true)
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum TextAlign {
    Start,
    End,
    Left,
    Right,
    Center,
    Justify,
}

impl TextAlign {
    pub(super) fn parse(value: &str) -> Result<Self> {
        match value {
            "start" => Ok(Self::Start),
            "end" => Ok(Self::End),
            "left" => Ok(Self::Left),
            "right" => Ok(Self::Right),
            "center" => Ok(Self::Center),
            "justify" => Ok(Self::Justify),
            _ => invalid("unsupported fo:text-align in outline properties"),
        }
    }

    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Start => "start",
            Self::End => "end",
            Self::Left => "left",
            Self::Right => "right",
            Self::Center => "center",
            Self::Justify => "justify",
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum PositionMode {
    LabelWidthAndPosition,
    LabelAlignment,
}

impl PositionMode {
    pub(super) fn parse(value: &str) -> Result<Self> {
        match value {
            "label-width-and-position" => Ok(Self::LabelWidthAndPosition),
            "label-alignment" => Ok(Self::LabelAlignment),
            _ => invalid("unsupported text:list-level-position-and-space-mode"),
        }
    }

    pub const fn as_str(self) -> &'static str {
        match self {
            Self::LabelWidthAndPosition => "label-width-and-position",
            Self::LabelAlignment => "label-alignment",
        }
    }
}

/// Complete positioning metadata below one outline level.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct ListProperties {
    pub text_align: Option<TextAlign>,
    pub space_before: Option<Length>,
    pub minimum_label_width: Option<Length>,
    pub minimum_label_distance: Option<Length>,
    pub font_name: Option<String>,
    pub width: Option<Length>,
    pub height: Option<Length>,
    pub vertical_relation: Option<String>,
    pub vertical_position: Option<String>,
    pub position_mode: Option<PositionMode>,
    pub label_alignment: Option<Alignment>,
    pub extensions: Vec<Attribute>,
}

/// Namespace-resolved `style:text-properties` attributes for one level.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct TextProperties {
    pub attributes: Vec<Attribute>,
}

/// One `text:outline-level-style` declaration.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct LevelStyle {
    pub level: u16,
    pub text_style_name: Option<String>,
    pub number_format: Option<NumberFormat>,
    pub number_prefix: Option<String>,
    pub number_suffix: Option<String>,
    pub letter_sync: Option<bool>,
    pub display_levels: Option<PositiveInteger>,
    pub start_value: Option<PositiveInteger>,
    pub list_level_properties: Option<ListProperties>,
    pub text_properties: Option<TextProperties>,
    pub extensions: Vec<Attribute>,
}

/// One named outline numbering style.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Style {
    pub name: String,
    pub levels: Vec<LevelStyle>,
    pub extensions: Vec<Attribute>,
}

impl Style {
    pub fn level(&self, level: u16) -> Option<&LevelStyle> {
        self.levels
            .iter()
            .find(|candidate| candidate.level == level)
    }

    pub fn validate(&self) -> Result<()> {
        validate_text(&self.name, "style:name", false)?;
        if self.levels.is_empty() || self.levels.len() > usize::from(MAX_OUTLINE_LEVELS) {
            return invalid("outline style has an invalid number of levels");
        }
        validate_extension_attributes(&self.extensions, true)?;
        let mut levels = HashSet::new();
        for level in &self.levels {
            if !(1..=MAX_OUTLINE_LEVELS).contains(&level.level) || !levels.insert(level.level) {
                return invalid("outline style has an invalid or duplicate level");
            }
            level.validate()?;
        }
        Ok(())
    }
}

impl LevelStyle {
    pub fn validate(&self) -> Result<()> {
        if !(1..=MAX_OUTLINE_LEVELS).contains(&self.level) {
            return invalid("outline level is outside the supported range");
        }
        if let Some(value) = self.text_style_name.as_deref() {
            validate_text(value, "text:style-name", false)?;
        }
        for (value, name) in [
            (self.number_prefix.as_deref(), "style:num-prefix"),
            (self.number_suffix.as_deref(), "style:num-suffix"),
        ] {
            if let Some(value) = value {
                validate_text(value, name, true)?;
            }
        }
        if self.letter_sync.is_some()
            && !self
                .number_format
                .as_ref()
                .is_some_and(|format| matches!(format.as_str(), "a" | "A"))
        {
            return invalid("style:num-letter-sync requires style:num-format 'a' or 'A'");
        }
        validate_extension_attributes(&self.extensions, true)?;
        if let Some(properties) = self.list_level_properties.as_ref() {
            properties.validate()?;
        }
        if let Some(properties) = self.text_properties.as_ref() {
            validate_extension_attributes(&properties.attributes, false)?;
        }
        Ok(())
    }
}

impl ListProperties {
    pub fn validate(&self) -> Result<()> {
        if let Some(value) = self.font_name.as_deref() {
            validate_text(value, "style:font-name", true)?;
        }
        for (value, name) in [
            (self.vertical_relation.as_deref(), "style:vertical-rel"),
            (self.vertical_position.as_deref(), "style:vertical-pos"),
        ] {
            if let Some(value) = value {
                validate_text(value, name, false)?;
            }
        }
        if self.position_mode == Some(PositionMode::LabelAlignment)
            && self.label_alignment.is_none()
        {
            return invalid("label-alignment mode requires style:list-level-label-alignment");
        }
        if let Some(alignment) = self.label_alignment.as_ref() {
            alignment.validate()?;
        }
        validate_extension_attributes(&self.extensions, true)
    }
}

/// The named outline styles stored in one ODF style part.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Styles {
    pub styles: Vec<Style>,
}

impl Styles {
    pub fn get(&self, name: &str) -> Option<&Style> {
        self.styles.iter().find(|style| style.name == name)
    }
}

pub(super) fn validate_extension_attributes(
    attributes: &[Attribute],
    foreign_only: bool,
) -> Result<()> {
    let mut seen = HashSet::new();
    for attribute in attributes {
        attribute.validate()?;
        if foreign_only
            && namespace_kind(attribute.namespace_uri.as_bytes()) != super::NamespaceKind::Other
        {
            return invalid("standard outline attributes must use typed fields");
        }
        if !seen.insert((&attribute.namespace_uri, &attribute.local_name)) {
            return invalid("duplicate expanded outline formatting attribute");
        }
    }
    Ok(())
}

pub(super) fn validate_text(value: &str, name: &str, allow_empty: bool) -> Result<()> {
    if value.len() > MAX_VALUE_BYTES
        || (!allow_empty && value.is_empty())
        || value.chars().any(char::is_control)
    {
        return invalid(format!("invalid {name}"));
    }
    Ok(())
}

pub(super) fn is_xml_local_name(value: &str) -> bool {
    let mut characters = value.chars();
    let Some(first) = characters.next() else {
        return false;
    };
    (first == '_' || first.is_ascii_alphabetic())
        && characters.all(|character| {
            character == '_'
                || character == '-'
                || character == '.'
                || character.is_ascii_alphanumeric()
        })
}
