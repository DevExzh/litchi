//! Typed ODF list-level label-alignment models.

use super::{FO_S, MAX_LEVEL, MAX_VALUE, STYLE_S, TEXT_S, bad};
use litchi_core::{Result, xml::escape_xml};

/// What separates a list label from the following text.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FollowedBy {
    ListTab,
    Space,
    Nothing,
}

impl FollowedBy {
    pub(super) fn parse(value: &str) -> Result<Self> {
        match value {
            "listtab" => Ok(Self::ListTab),
            "space" => Ok(Self::Space),
            "nothing" => Ok(Self::Nothing),
            _ => Err(bad("invalid text:label-followed-by")),
        }
    }

    fn xml(self) -> &'static str {
        match self {
            Self::ListTab => "listtab",
            Self::Space => "space",
            Self::Nothing => "nothing",
        }
    }
}

/// A validated ODF length lexical value.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Length(String);

impl Length {
    pub fn new(value: impl Into<String>) -> Result<Self> {
        let value = value.into();
        if value.len() > MAX_VALUE || !is_length(&value) {
            return Err(bad("invalid ODF list label length"));
        }
        Ok(Self(value))
    }

    pub fn as_str(&self) -> &str {
        &self.0
    }
}

fn is_length(value: &str) -> bool {
    let Some(number) = ["cm", "mm", "in", "pt", "pc", "px"]
        .iter()
        .find_map(|unit| value.strip_suffix(unit))
    else {
        return false;
    };
    let number = number.strip_prefix('-').unwrap_or(number);
    let mut parts = number.split('.');
    let integer = parts.next().unwrap_or("");
    let fraction = parts.next();
    if parts.next().is_some() {
        return false;
    }
    let digits = |part: &str| part.bytes().all(|byte| byte.is_ascii_digit());
    match fraction {
        None => !integer.is_empty() && digits(integer),
        Some(fraction) => {
            digits(integer) && digits(fraction) && (!integer.is_empty() || !fraction.is_empty())
        },
    }
}

/// Label alignment properties attached to one list level.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Alignment {
    pub label_followed_by: FollowedBy,
    pub list_tab_stop_position: Option<Length>,
    pub text_indent: Option<Length>,
    pub margin_left: Option<Length>,
}

impl Alignment {
    pub fn new(label_followed_by: FollowedBy) -> Self {
        Self {
            label_followed_by,
            list_tab_stop_position: None,
            text_indent: None,
            margin_left: None,
        }
    }

    pub fn validate(&self) -> Result<()> {
        for value in [
            &self.list_tab_stop_position,
            &self.text_indent,
            &self.margin_left,
        ]
        .into_iter()
        .flatten()
        {
            if !is_length(value.as_str()) {
                return Err(bad("invalid ODF list label length"));
            }
        }
        Ok(())
    }

    pub fn to_xml_fragment(&self) -> Result<String> {
        self.validate()?;
        let mut xml = format!(
            r#"<style:list-level-label-alignment xmlns:style="{STYLE_S}" xmlns:text="{TEXT_S}" xmlns:fo="{FO_S}" text:label-followed-by="{}""#,
            self.label_followed_by.xml()
        );
        if let Some(value) = &self.list_tab_stop_position {
            xml.push_str(&format!(
                r#" text:list-tab-stop-position="{}""#,
                escape_xml(value.as_str())
            ));
        }
        if let Some(value) = &self.text_indent {
            xml.push_str(&format!(
                r#" fo:text-indent="{}""#,
                escape_xml(value.as_str())
            ));
        }
        if let Some(value) = &self.margin_left {
            xml.push_str(&format!(
                r#" fo:margin-left="{}""#,
                escape_xml(value.as_str())
            ));
        }
        xml.push_str("/>");
        Ok(xml)
    }
}

/// ODF list-style family containing a list-level alignment.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Kind {
    List,
    Outline,
}

/// Alignment properties for one named style and level.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Style {
    pub list_style_kind: Kind,
    pub list_style_name: String,
    pub level: u16,
    pub alignment: Alignment,
}

impl Style {
    pub fn new(name: impl Into<String>, level: u16, alignment: Alignment) -> Result<Self> {
        Self::new_in(Kind::List, name, level, alignment)
    }

    pub fn new_in(
        list_style_kind: Kind,
        name: impl Into<String>,
        level: u16,
        alignment: Alignment,
    ) -> Result<Self> {
        let style = Self {
            list_style_kind,
            list_style_name: name.into(),
            level,
            alignment,
        };
        style.validate()?;
        Ok(style)
    }

    pub fn validate(&self) -> Result<()> {
        if self.list_style_name.is_empty()
            || self.list_style_name.len() > MAX_VALUE
            || self.list_style_name.chars().any(char::is_control)
        {
            return Err(bad("invalid list style name"));
        }
        if !(1..=MAX_LEVEL).contains(&self.level) {
            return Err(bad("list level outside supported range"));
        }
        self.alignment.validate()
    }
}

/// All parsed list-level alignments from one ODF XML part.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Styles {
    pub levels: Vec<Style>,
}

impl Styles {
    pub fn get(&self, name: &str, level: u16) -> Option<&Style> {
        self.get_in(Kind::List, name, level)
    }

    pub fn get_in(&self, kind: Kind, name: &str, level: u16) -> Option<&Style> {
        self.levels.iter().find(|style| {
            style.list_style_kind == kind && style.list_style_name == name && style.level == level
        })
    }
}
