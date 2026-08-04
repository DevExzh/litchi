//! Typed ODF graphic-property values, records, and lexical validation.

use litchi_core::{Error, Result, xml::escape_xml};
use std::collections::BTreeMap;

use super::{DR3D_NS, DRAW_NS, FO_NS, MAX_VALUE, OFFICE_NS, STYLE_NS, SVG_NS, TEXT_NS, XLINK_NS};

pub(super) fn bad(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}
pub(super) fn safe(value: &str, name: &str, empty: bool) -> Result<()> {
    if (!empty && value.is_empty())
        || value.len() > MAX_VALUE
        || value.chars().any(
            |c| matches!(c,'\0'..='\u{8}'|'\u{b}'|'\u{c}'|'\u{e}'..='\u{1f}'|'\u{fffe}'|'\u{ffff}'),
        )
    {
        return Err(bad(format!("invalid {name}")));
    }
    Ok(())
}
fn decimal(value: &str, signed: bool) -> bool {
    let value = if signed {
        value.strip_prefix('-').unwrap_or(value)
    } else {
        value
    };
    if value.is_empty() {
        return false;
    }
    let mut parts = value.split('.');
    let left = parts.next().unwrap_or_default();
    let right = parts.next();
    if parts.next().is_some() {
        return false;
    }
    match right {
        None => !left.is_empty() && left.bytes().all(|b| b.is_ascii_digit()),
        Some(right) => {
            (!left.is_empty() || !right.is_empty())
                && left.bytes().all(|b| b.is_ascii_digit())
                && right.bytes().all(|b| b.is_ascii_digit())
        },
    }
}
fn length(value: &str, signed: bool, positive: bool, pixel_only: bool) -> bool {
    let units = if pixel_only {
        &["px"][..]
    } else {
        &["cm", "mm", "in", "pt", "pc", "px"][..]
    };
    units.iter().any(|unit| {
        value.strip_suffix(unit).is_some_and(|number| {
            decimal(number, signed)
                && (!positive || number.bytes().any(|b| b.is_ascii_digit() && b != b'0'))
        })
    })
}
fn percent(value: &str, signed: bool, ranged: bool) -> bool {
    let Some(number) = value.strip_suffix('%') else {
        return false;
    };
    if !decimal(number, signed) {
        return false;
    }
    if !ranged {
        return true;
    }
    number
        .trim_start_matches('-')
        .parse::<f64>()
        .is_ok_and(|number| number <= 100.0)
}
fn integer(value: &str, positive: bool) -> bool {
    let value = value.strip_prefix('+').unwrap_or(value);
    if positive {
        !value.is_empty()
            && value.bytes().all(|b| b.is_ascii_digit())
            && value.bytes().any(|b| b != b'0')
    } else {
        !value.is_empty() && value.bytes().all(|b| b.is_ascii_digit())
    }
}
fn duration(value: &str) -> bool {
    let value = value.strip_prefix('-').unwrap_or(value);
    value.starts_with('P')
        && value.len() > 1
        && value.bytes().any(|b| b.is_ascii_digit())
        && value.bytes().all(|b| {
            b.is_ascii_digit() || matches!(b, b'P' | b'T' | b'Y' | b'M' | b'D' | b'H' | b'S' | b'.')
        })
}
pub(super) fn ncname(value: &str, empty: bool) -> bool {
    if value.is_empty() {
        return empty;
    }
    let mut chars = value.chars();
    let Some(first) = chars.next() else {
        return false;
    };
    (first == '_' || first.is_alphabetic())
        && chars.all(|c| c == '_' || c == '-' || c == '.' || c.is_alphanumeric())
}
fn clip(value: &str) -> bool {
    let Some(inner) = value
        .strip_prefix("rect(")
        .and_then(|v| v.strip_suffix(')'))
    else {
        return false;
    };
    let values: Vec<_> = inner.split(',').map(str::trim).collect();
    values.len() == 4
        && values.iter().all(|value| {
            *value == "auto"
                || ["cm", "mm", "in", "pt", "pc"].iter().any(|unit| {
                    value
                        .strip_suffix(unit)
                        .is_some_and(|number| decimal(number, true))
                })
        })
}

/// Closed namespaces used by the 174 graphic property names.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub enum Namespace {
    Dr3d,
    Draw,
    Fo,
    Style,
    Svg,
    Text,
}

/// A lexically validated value whose variant identifies its normative datatype.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum Value {
    Boolean(bool),
    Color(String),
    Length(String),
    NonNegativeLength(String),
    PositiveLength(String),
    NonNegativePixelLength(String),
    Percent(String),
    ZeroToHundredPercent(String),
    SignedZeroToHundredPercent(String),
    NonNegativeInteger(String),
    PositiveInteger(String),
    Duration(String),
    Clip(String),
    StyleNameRef(String),
    StyleNameRefs(Vec<String>),
    Keyword(String),
    Text(String),
    Compound(String),
}
impl Value {
    pub fn lexical(&self) -> String {
        match self {
            Self::Boolean(value) => value.to_string(),
            Self::Color(value)
            | Self::Length(value)
            | Self::NonNegativeLength(value)
            | Self::PositiveLength(value)
            | Self::NonNegativePixelLength(value)
            | Self::Percent(value)
            | Self::ZeroToHundredPercent(value)
            | Self::SignedZeroToHundredPercent(value)
            | Self::NonNegativeInteger(value)
            | Self::PositiveInteger(value)
            | Self::Duration(value)
            | Self::Clip(value)
            | Self::StyleNameRef(value)
            | Self::Keyword(value)
            | Self::Text(value)
            | Self::Compound(value) => value.clone(),
            Self::StyleNameRefs(values) => values.join(" "),
        }
    }
}

// The generated table is the normative 174-property ODF grammar.
include!("../graphic_property_specs.rs");

fn validate_ref(reference: &str, value: &str) -> Option<Value> {
    match reference {
        "boolean" => match value {
            "true" => Some(Value::Boolean(true)),
            "false" => Some(Value::Boolean(false)),
            _ => None,
        },
        "color" => (value.len() == 7
            && value.starts_with('#')
            && value[1..].bytes().all(|b| b.is_ascii_hexdigit()))
        .then(|| Value::Color(value.to_owned())),
        "length" | "coordinate" | "distance" => {
            length(value, true, false, false).then(|| Value::Length(value.to_owned()))
        },
        "nonNegativeLength" => {
            length(value, false, false, false).then(|| Value::NonNegativeLength(value.to_owned()))
        },
        "positiveLength" => {
            length(value, false, true, false).then(|| Value::PositiveLength(value.to_owned()))
        },
        "nonNegativePixelLength" => length(value, false, false, true)
            .then(|| Value::NonNegativePixelLength(value.to_owned())),
        "percent" => percent(value, true, false).then(|| Value::Percent(value.to_owned())),
        "zeroToHundredPercent" => {
            percent(value, false, true).then(|| Value::ZeroToHundredPercent(value.to_owned()))
        },
        "signedZeroToHundredPercent" => {
            percent(value, true, true).then(|| Value::SignedZeroToHundredPercent(value.to_owned()))
        },
        "nonNegativeInteger" => {
            integer(value, false).then(|| Value::NonNegativeInteger(value.to_owned()))
        },
        "positiveInteger" => integer(value, true).then(|| Value::PositiveInteger(value.to_owned())),
        "duration" => duration(value).then(|| Value::Duration(value.to_owned())),
        "clipShape" => clip(value).then(|| Value::Clip(value.to_owned())),
        "styleNameRef" => ncname(value, true).then(|| Value::StyleNameRef(value.to_owned())),
        "styleNameRefs" => {
            let values: Vec<_> = value.split_ascii_whitespace().map(str::to_owned).collect();
            values
                .iter()
                .all(|value| ncname(value, false))
                .then_some(Value::StyleNameRefs(values))
        },
        "horizontal-mirror" => matches!(
            value,
            "horizontal" | "horizontal-on-odd" | "horizontal-on-even"
        )
        .then(|| Value::Keyword(value.to_owned())),
        "borderWidths" => {
            let values: Vec<_> = value.split_ascii_whitespace().collect();
            (values.len() == 3 && values.iter().all(|value| length(value, false, true, false)))
                .then(|| Value::Compound(value.to_owned()))
        },
        "angle" | "string" | "shadowType" => Some(Value::Text(value.to_owned())),
        _ => None,
    }
}
fn validate_spec(
    value: &str,
    keywords: &[&str],
    references: &[&str],
    list: bool,
    kind: Kind,
) -> Result<Value> {
    safe(value, "graphic property value", true)?;
    if kind == Kind::DrawTileRepeatOffset {
        let parts: Vec<_> = value.split_ascii_whitespace().collect();
        if parts.len() == 2
            && percent(parts[0], false, true)
            && matches!(parts[1], "horizontal" | "vertical")
        {
            return Ok(Value::Compound(value.to_owned()));
        }
        return Err(bad("invalid draw:tile-repeat-offset"));
    }
    if list {
        let parts: Vec<_> = value.split_ascii_whitespace().collect();
        if parts.iter().all(|part| {
            keywords.contains(part)
                || references
                    .iter()
                    .any(|reference| validate_ref(reference, part).is_some())
        }) {
            return Ok(Value::Compound(value.to_owned()));
        }
        return Err(bad(format!(
            "invalid {}:{} list",
            kind.namespace().prefix(),
            kind.local_name()
        )));
    }
    if keywords.contains(&value) {
        return Ok(Value::Keyword(value.to_owned()));
    }
    for reference in references {
        if let Some(value) = validate_ref(reference, value) {
            return Ok(value);
        }
    }
    Err(bad(format!(
        "invalid {}:{} value",
        kind.namespace().prefix(),
        kind.local_name()
    )))
}

impl Namespace {
    fn prefix(self) -> &'static str {
        match self {
            Self::Dr3d => "dr3d",
            Self::Draw => "draw",
            Self::Fo => "fo",
            Self::Style => "style",
            Self::Svg => "svg",
            Self::Text => "text",
        }
    }
}

/// One closed-name typed graphic property.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Property {
    kind: Kind,
    value: Value,
}
impl Property {
    pub fn new(kind: Kind, lexical: &str) -> Result<Self> {
        Ok(Self {
            kind,
            value: kind.parse_value(lexical)?,
        })
    }
    pub fn kind(&self) -> Kind {
        self.kind
    }
    pub fn value(&self) -> &Value {
        &self.value
    }
    pub fn lexical(&self) -> String {
        self.value.lexical()
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub enum ChildKind {
    ListStyle,
    BackgroundImage,
    Columns,
}
impl ChildKind {
    pub(super) fn namespace(self) -> Namespace {
        match self {
            Self::ListStyle => Namespace::Text,
            Self::BackgroundImage | Self::Columns => Namespace::Style,
        }
    }
    pub(super) fn local(self) -> &'static str {
        match self {
            Self::ListStyle => "list-style",
            Self::BackgroundImage => "background-image",
            Self::Columns => "columns",
        }
    }
}
/// Bounded inert XML for an immediate normative child. It is never executed, rendered, or fetched.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Child {
    kind: ChildKind,
    xml: String,
}
impl Child {
    pub fn new(kind: ChildKind, xml: impl Into<String>) -> Result<Self> {
        let xml = xml.into();
        super::codec::validate_child(kind, &xml)?;
        Ok(Self { kind, xml })
    }
    pub fn kind(&self) -> ChildKind {
        self.kind
    }
    pub fn xml(&self) -> &str {
        &self.xml
    }
}
/// Complete typed `style:graphic-properties` value.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Properties {
    pub(super) properties: BTreeMap<Kind, Value>,
    pub(super) children: BTreeMap<ChildKind, Child>,
}
impl Properties {
    pub fn set(&mut self, property: Property) -> Option<Value> {
        self.properties.insert(property.kind, property.value)
    }
    pub fn set_lexical(&mut self, kind: Kind, value: &str) -> Result<Option<Value>> {
        Ok(self.set(Property::new(kind, value)?))
    }
    pub fn get(&self, kind: Kind) -> Option<&Value> {
        self.properties.get(&kind)
    }
    pub fn remove(&mut self, kind: Kind) -> Option<Value> {
        self.properties.remove(&kind)
    }
    pub fn iter(&self) -> impl Iterator<Item = (Kind, &Value)> {
        self.properties.iter().map(|(kind, value)| (*kind, value))
    }
    pub fn set_child(&mut self, child: Child) -> Option<Child> {
        self.children.insert(child.kind, child)
    }
    pub fn child(&self, kind: ChildKind) -> Option<&Child> {
        self.children.get(&kind)
    }
    pub fn remove_child(&mut self, kind: ChildKind) -> Option<Child> {
        self.children.remove(&kind)
    }
    pub fn from_xml_fragment(fragment: &str) -> Result<Self> {
        let xml = format!(
            r#"<office:document xmlns:office="{OFFICE_NS}" xmlns:style="{STYLE_NS}"><office:styles><style:style style:name="fragment" style:family="graphic">{fragment}</style:style></office:styles></office:document>"#
        );
        let mut set = super::codec::parse_graphic_style_properties(&xml)?;
        set.styles
            .pop()
            .and_then(|style| style.properties)
            .ok_or_else(|| bad("fragment does not contain style:graphic-properties"))
    }
    pub fn to_xml_fragment(&self) -> Result<String> {
        let mut xml = format!(
            r#"<style:graphic-properties xmlns:office="{OFFICE_NS}" xmlns:style="{STYLE_NS}" xmlns:dr3d="{DR3D_NS}" xmlns:draw="{DRAW_NS}" xmlns:fo="{FO_NS}" xmlns:svg="{SVG_NS}" xmlns:text="{TEXT_NS}" xmlns:xlink="{XLINK_NS}""#
        );
        for (kind, value) in &self.properties {
            xml.push(' ');
            xml.push_str(kind.namespace().prefix());
            xml.push(':');
            xml.push_str(kind.local_name());
            xml.push_str("=\"");
            xml.push_str(&escape_xml(&value.lexical()));
            xml.push('"')
        }
        if self.children.is_empty() {
            xml.push_str("/>")
        } else {
            xml.push('>');
            for child in self.children.values() {
                super::codec::validate_child(child.kind, &child.xml)?;
                xml.push_str(&child.xml)
            }
            xml.push_str("</style:graphic-properties>")
        }
        Ok(xml)
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Style {
    pub name: Option<String>,
    pub parent_style_name: Option<String>,
    pub is_default_style: bool,
    pub properties: Option<Properties>,
}
impl Style {
    pub fn named(name: impl Into<String>, properties: Option<Properties>) -> Result<Self> {
        let value = Self {
            name: Some(name.into()),
            parent_style_name: None,
            is_default_style: false,
            properties,
        };
        value.validate()?;
        Ok(value)
    }
    pub fn default_style(properties: Option<Properties>) -> Self {
        Self {
            name: None,
            parent_style_name: None,
            is_default_style: true,
            properties,
        }
    }
    pub fn validate(&self) -> Result<()> {
        match (&self.name, self.is_default_style) {
            (Some(value), false) if ncname(value, false) => {},
            (None, true) => {},
            _ => return Err(bad("invalid graphic style identity")),
        }
        if let Some(value) = &self.parent_style_name
            && (self.is_default_style || !ncname(value, false))
        {
            return Err(bad("invalid parent graphic style name"));
        }
        Ok(())
    }
    pub fn to_xml_fragment(&self) -> Result<String> {
        self.validate()?;
        let tag = if self.is_default_style {
            "default-style"
        } else {
            "style"
        };
        let mut xml = format!(r#"<style:{tag} xmlns:style="{STYLE_NS}" style:family="graphic""#);
        if let Some(value) = &self.name {
            xml.push_str(&format!(r#" style:name="{}""#, escape_xml(value)))
        }
        if let Some(value) = &self.parent_style_name {
            xml.push_str(&format!(
                r#" style:parent-style-name="{}""#,
                escape_xml(value)
            ))
        }
        if let Some(value) = &self.properties {
            xml.push('>');
            xml.push_str(&value.to_xml_fragment()?);
            xml.push_str(&format!("</style:{tag}>"))
        } else {
            xml.push_str("/>")
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
        self.styles
            .iter()
            .find(|style| style.name.as_deref() == Some(name))
    }
    pub fn default_style(&self) -> Option<&Style> {
        self.styles.iter().find(|style| style.is_default_style)
    }
}
