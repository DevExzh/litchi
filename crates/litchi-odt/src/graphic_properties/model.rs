//! Typed ODF graphic-property values, records, and lexical validation.

use litchi_core::{Result, xml::escape_xml};
use std::collections::BTreeMap;

use super::{DR3D_NS, DRAW_NS, FO_NS, OFFICE_NS, STYLE_NS, SVG_NS, TEXT_NS, XLINK_NS};
pub(crate) use crate::graphic_property_specs::codec::{bad, ncname, safe};
pub use crate::graphic_property_specs::model::{Kind, Namespace, Value};

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
