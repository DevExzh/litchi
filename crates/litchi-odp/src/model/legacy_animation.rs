//! Legacy `presentation:animations` effect trees.

use super::{Attribute, Namespace};
use litchi_core::{Error, Result, xml::escape_xml};
use std::collections::{BTreeMap, BTreeSet};

/// A schema-defined legacy ODF presentation effect element.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Kind {
    /// The page-level `presentation:animations` container.
    Animations,
    /// A group of effects.
    Group,
    /// Dim a shape to a color.
    Dim,
    /// Hide an entire shape.
    HideShape,
    /// Hide a shape's text.
    HideText,
    /// Play a media shape.
    Play,
    /// Show an entire shape.
    ShowShape,
    /// Show a shape's text.
    ShowText,
    /// Play a referenced sound with an effect.
    Sound,
}

impl Kind {
    pub(crate) fn from_local_name(name: &[u8]) -> Option<Self> {
        match name {
            b"animations" => Some(Self::Animations),
            b"animation-group" => Some(Self::Group),
            b"dim" => Some(Self::Dim),
            b"hide-shape" => Some(Self::HideShape),
            b"hide-text" => Some(Self::HideText),
            b"play" => Some(Self::Play),
            b"show-shape" => Some(Self::ShowShape),
            b"show-text" => Some(Self::ShowText),
            b"sound" => Some(Self::Sound),
            _ => None,
        }
    }

    pub(crate) const fn local_name(self) -> &'static str {
        match self {
            Self::Animations => "animations",
            Self::Group => "animation-group",
            Self::Dim => "dim",
            Self::HideShape => "hide-shape",
            Self::HideText => "hide-text",
            Self::Play => "play",
            Self::ShowShape => "show-shape",
            Self::ShowText => "show-text",
            Self::Sound => "sound",
        }
    }

    pub(crate) const fn allows_child(self, child: Self) -> bool {
        match self {
            Self::Animations => !matches!(child, Self::Animations | Self::Sound),
            Self::Group => !matches!(child, Self::Animations | Self::Group | Self::Sound),
            Self::Dim | Self::HideShape | Self::HideText | Self::ShowShape | Self::ShowText => {
                matches!(child, Self::Sound)
            },
            Self::Play | Self::Sound => false,
        }
    }
}

/// A bounded, inert node in the legacy ODF presentation effect tree.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Node {
    kind: Kind,
    attributes: Vec<Attribute>,
    children: Vec<Node>,
}

impl Node {
    /// Create an empty legacy effect node.
    #[must_use]
    pub fn new(kind: Kind) -> Self {
        Self {
            kind,
            attributes: Vec::new(),
            children: Vec::new(),
        }
    }

    /// Return the element kind.
    #[must_use]
    pub fn kind(&self) -> Kind {
        self.kind
    }

    /// Return the expanded-name attributes.
    #[must_use]
    pub fn attributes(&self) -> &[Attribute] {
        &self.attributes
    }

    /// Return mutable attributes.
    pub fn attributes_mut(&mut self) -> &mut Vec<Attribute> {
        &mut self.attributes
    }

    /// Add or replace an expanded-name attribute.
    pub fn set_attribute(&mut self, attribute: Attribute) {
        if let Some(existing) = self.attributes.iter_mut().find(|existing| {
            existing.namespace() == attribute.namespace()
                && existing.local_name() == attribute.local_name()
        }) {
            *existing = attribute;
        } else {
            self.attributes.push(attribute);
        }
    }

    /// Find an attribute by expanded name.
    pub fn attribute(&self, namespace: &Namespace, local_name: &str) -> Option<&str> {
        self.attributes
            .iter()
            .find(|attribute| {
                attribute.namespace() == namespace && attribute.local_name() == local_name
            })
            .map(Attribute::value)
    }

    /// Return child effects.
    #[must_use]
    pub fn children(&self) -> &[Node] {
        &self.children
    }

    /// Return mutable child effects.
    pub fn children_mut(&mut self) -> &mut Vec<Node> {
        &mut self.children
    }

    /// Add a schema-valid child effect.
    ///
    /// # Errors
    /// Returns an error when the input is malformed or a configured limit is exceeded.
    pub fn add_child(&mut self, child: Node) -> Result<()> {
        if !self.kind.allows_child(child.kind) {
            return Err(invalid_child(self.kind, child.kind));
        }
        self.children.push(child);
        Ok(())
    }

    pub(crate) fn from_parsed(kind: Kind, attributes: Vec<Attribute>, children: Vec<Self>) -> Self {
        Self {
            kind,
            attributes,
            children,
        }
    }

    pub(crate) fn collect_extension_namespaces(&self, uris: &mut BTreeSet<String>) {
        for attribute in &self.attributes {
            if let Namespace::Other(uri) = attribute.namespace() {
                uris.insert(uri.clone());
            }
        }
        for child in &self.children {
            child.collect_extension_namespaces(uris);
        }
    }

    pub(crate) fn validate(&self, depth: usize, count: &mut usize) -> Result<()> {
        if depth > 128 {
            return Err(Error::InvalidFormat(
                "legacy ODP animation nesting exceeds 128 levels".to_string(),
            ));
        }
        *count = count.checked_add(1).ok_or_else(|| {
            Error::InvalidFormat("legacy ODP animation node count overflow".to_string())
        })?;
        if *count > 65_536 || self.attributes.len() > 256 {
            return Err(Error::InvalidFormat(
                "legacy ODP animation exceeds resource limits".to_string(),
            ));
        }
        let mut names = BTreeSet::new();
        for attribute in &self.attributes {
            let name = (attribute.namespace(), attribute.local_name());
            if !names.insert(name) {
                return Err(Error::InvalidFormat(format!(
                    "duplicate legacy animation attribute '{}'",
                    attribute.local_name()
                )));
            }
        }
        self.validate_required_attributes()?;
        for child in &self.children {
            if !self.kind.allows_child(child.kind) {
                return Err(invalid_child(self.kind, child.kind));
            }
            child.validate(depth + 1, count)?;
        }
        Ok(())
    }

    fn validate_required_attributes(&self) -> Result<()> {
        let draw = &Namespace::Draw;
        match self.kind {
            Kind::Dim => {
                self.require(draw, "shape-id")?;
                self.require(draw, "color")?;
            },
            Kind::HideShape | Kind::HideText | Kind::Play | Kind::ShowShape | Kind::ShowText => {
                self.require(draw, "shape-id")?;
            },
            Kind::Sound => {
                self.require(&Namespace::Xlink, "href")?;
                let link_type = self.require(&Namespace::Xlink, "type")?;
                if link_type != "simple" {
                    return Err(Error::InvalidFormat(
                        "presentation:sound xlink:type must be 'simple'".to_string(),
                    ));
                }
            },
            Kind::Animations | Kind::Group => {},
        }
        Ok(())
    }

    fn require(&self, namespace: &Namespace, local_name: &str) -> Result<&str> {
        self.attribute(namespace, local_name).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "presentation:{} is missing required attribute '{local_name}'",
                self.kind.local_name()
            ))
        })
    }

    pub(crate) fn write_xml(
        &self,
        output: &mut String,
        extensions: &BTreeMap<String, String>,
    ) -> Result<()> {
        output.push_str("<presentation:");
        output.push_str(self.kind.local_name());
        for attribute in &self.attributes {
            output.push(' ');
            if let Some(prefix) = attribute.namespace().prefix(extensions)? {
                output.push_str(prefix);
                output.push(':');
            }
            output.push_str(attribute.local_name());
            output.push_str("=\"");
            output.push_str(&escape_xml(attribute.value()));
            output.push('"');
        }
        if self.children.is_empty() {
            output.push_str("/>");
            return Ok(());
        }
        output.push('>');
        for child in &self.children {
            child.write_xml(output, extensions)?;
        }
        output.push_str("</presentation:");
        output.push_str(self.kind.local_name());
        output.push('>');
        Ok(())
    }
}

pub(crate) fn validate_legacy_animation_root(root: &Node) -> Result<()> {
    if root.kind != Kind::Animations {
        return Err(Error::InvalidFormat(
            "legacy ODP animation root must be presentation:animations".to_string(),
        ));
    }
    let mut node_count = 0;
    root.validate(1, &mut node_count)
}

fn invalid_child(parent: Kind, child: Kind) -> Error {
    Error::InvalidFormat(format!(
        "presentation:{} cannot contain presentation:{}",
        parent.local_name(),
        child.local_name()
    ))
}
