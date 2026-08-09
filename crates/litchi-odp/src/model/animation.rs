//! Inert ODF presentation animation trees.

use litchi_core::{Error, Result, xml::escape_xml};
use std::collections::{BTreeMap, BTreeSet};

pub(crate) const ANIMATION_NAMESPACE: &str = "urn:oasis:names:tc:opendocument:xmlns:animation:1.0";
pub(crate) const SMIL_NAMESPACE: &str = "urn:oasis:names:tc:opendocument:xmlns:smil-compatible:1.0";
pub(crate) const PRESENTATION_NAMESPACE: &str =
    "urn:oasis:names:tc:opendocument:xmlns:presentation:1.0";
pub(crate) const DRAW_NAMESPACE: &str = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
pub(crate) const SVG_NAMESPACE: &str = "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0";
pub(crate) const XLINK_NAMESPACE: &str = "http://www.w3.org/1999/xlink";
pub(crate) const XML_NAMESPACE: &str = "http://www.w3.org/XML/1998/namespace";

/// One of the animation elements defined by ODF 1.3, Part 3, section 10.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Kind {
    /// Animate an attribute value.
    Animate,
    /// Animate a color value.
    AnimateColor,
    /// Animate an object along a motion path.
    AnimateMotion,
    /// Animate a transformation.
    AnimateTransform,
    /// Play referenced audio as part of a timing tree.
    Audio,
    /// Store an application command. Litchi never executes this command.
    Command,
    /// Apply child animations to parts of a target in turn.
    Iterate,
    /// Play child animations in parallel.
    Parallel,
    /// A name/value parameter belonging to a command.
    Parameter,
    /// Play child animations sequentially.
    Sequence,
    /// Set an attribute value at a point in the timing tree.
    Set,
    /// Apply a transition filter.
    TransitionFilter,
}

impl Kind {
    pub(crate) fn from_local_name(name: &[u8]) -> Option<Self> {
        match name {
            b"animate" => Some(Self::Animate),
            b"animateColor" => Some(Self::AnimateColor),
            b"animateMotion" => Some(Self::AnimateMotion),
            b"animateTransform" => Some(Self::AnimateTransform),
            b"audio" => Some(Self::Audio),
            b"command" => Some(Self::Command),
            b"iterate" => Some(Self::Iterate),
            b"par" => Some(Self::Parallel),
            b"param" => Some(Self::Parameter),
            b"seq" => Some(Self::Sequence),
            b"set" => Some(Self::Set),
            b"transitionFilter" => Some(Self::TransitionFilter),
            _ => None,
        }
    }

    pub(crate) const fn local_name(self) -> &'static str {
        match self {
            Self::Animate => "animate",
            Self::AnimateColor => "animateColor",
            Self::AnimateMotion => "animateMotion",
            Self::AnimateTransform => "animateTransform",
            Self::Audio => "audio",
            Self::Command => "command",
            Self::Iterate => "iterate",
            Self::Parallel => "par",
            Self::Parameter => "param",
            Self::Sequence => "seq",
            Self::Set => "set",
            Self::TransitionFilter => "transitionFilter",
        }
    }

    pub(crate) const fn allows_child(self, child: Self) -> bool {
        match self {
            Self::Command => matches!(child, Self::Parameter),
            Self::Iterate | Self::Parallel | Self::Sequence => !matches!(child, Self::Parameter),
            Self::Animate
            | Self::AnimateColor
            | Self::AnimateMotion
            | Self::AnimateTransform
            | Self::Audio
            | Self::Parameter
            | Self::Set
            | Self::TransitionFilter => false,
        }
    }

    pub(crate) const fn allowed_at_page_root(self) -> bool {
        !matches!(self, Self::Parameter)
    }
}

/// Namespace of an animation attribute.
///
/// ODF animation attributes primarily use `anim`, `smil`, `presentation`,
/// `svg`, `xlink`, and `xml`. Foreign namespaces are retained by URI and are
/// assigned deterministic prefixes when a mutable presentation is saved.
#[derive(Debug, Clone, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub enum Namespace {
    /// No namespace.
    None,
    /// The ODF animation namespace.
    Animation,
    /// The ODF SMIL-compatible namespace.
    Smil,
    /// The ODF presentation namespace.
    Presentation,
    /// The ODF drawing namespace.
    Draw,
    /// The ODF SVG-compatible namespace.
    Svg,
    /// The W3C `XLink` namespace.
    Xlink,
    /// The reserved XML namespace.
    Xml,
    /// A foreign extension namespace, identified by its URI.
    Other(String),
}

impl Namespace {
    pub(crate) fn from_uri(uri: Option<&str>) -> Self {
        match uri {
            None => Self::None,
            Some(ANIMATION_NAMESPACE) => Self::Animation,
            Some(SMIL_NAMESPACE) => Self::Smil,
            Some(PRESENTATION_NAMESPACE) => Self::Presentation,
            Some(DRAW_NAMESPACE) => Self::Draw,
            Some(SVG_NAMESPACE) => Self::Svg,
            Some(XLINK_NAMESPACE) => Self::Xlink,
            Some(XML_NAMESPACE) => Self::Xml,
            Some(extension_uri) => Self::Other(extension_uri.to_string()),
        }
    }

    pub(crate) fn prefix<'a>(
        &'a self,
        extensions: &'a BTreeMap<String, String>,
    ) -> Result<Option<&'a str>> {
        match self {
            Self::None => Ok(None),
            Self::Animation => Ok(Some("anim")),
            Self::Smil => Ok(Some("smil")),
            Self::Presentation => Ok(Some("presentation")),
            Self::Draw => Ok(Some("draw")),
            Self::Svg => Ok(Some("svg")),
            Self::Xlink => Ok(Some("xlink")),
            Self::Xml => Ok(Some("xml")),
            Self::Other(uri) => extensions
                .get(uri)
                .map(String::as_str)
                .map(Some)
                .ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "animation extension namespace '{uri}' was not declared"
                    ))
                }),
        }
    }
}

/// An expanded-name attribute attached to an [`Node`].
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Attribute {
    namespace: Namespace,
    local_name: String,
    value: String,
}

impl Attribute {
    /// Create an animation attribute.
    ///
    /// # Errors
    /// Returns an error when the input is malformed or a configured limit is exceeded.
    pub fn new(
        namespace: Namespace,
        local_name: impl Into<String>,
        value: impl Into<String>,
    ) -> Result<Self> {
        let local_name_string = local_name.into();
        let value_string = value.into();
        validate_ncname(&local_name_string)?;
        validate_attribute_namespace(&namespace)?;
        validate_attribute_value(&value_string)?;
        Ok(Self {
            namespace,
            local_name: local_name_string,
            value: value_string,
        })
    }

    /// Return the attribute namespace.
    #[must_use]
    pub fn namespace(&self) -> &Namespace {
        &self.namespace
    }

    /// Return the attribute local name.
    #[must_use]
    pub fn local_name(&self) -> &str {
        &self.local_name
    }

    /// Return the unescaped attribute value.
    #[must_use]
    pub fn value(&self) -> &str {
        &self.value
    }

    /// Replace the attribute value.
    ///
    /// # Errors
    /// Returns an error when the input is malformed or a configured limit is exceeded.
    pub fn set_value(&mut self, value: impl Into<String>) -> Result<()> {
        let replacement = value.into();
        validate_attribute_value(&replacement)?;
        self.value = replacement;
        Ok(())
    }

    pub(crate) fn from_parsed(
        namespace: Namespace,
        local_name: String,
        value: String,
    ) -> Result<Self> {
        Self::new(namespace, local_name, value)
    }
}

/// A bounded, inert node in an ODF presentation timing tree.
///
/// The tree preserves timing, target, motion, audio-reference, transition, and
/// command metadata. It deliberately provides no playback or command-execution
/// behavior.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Node {
    kind: Kind,
    attributes: Vec<Attribute>,
    children: Vec<Node>,
}

impl Node {
    /// Create an empty animation node.
    #[must_use]
    pub fn new(kind: Kind) -> Self {
        Self {
            kind,
            attributes: Vec::new(),
            children: Vec::new(),
        }
    }

    /// Return the schema-defined node kind.
    #[must_use]
    pub fn kind(&self) -> Kind {
        self.kind
    }

    /// Return the node's expanded-name attributes.
    #[must_use]
    pub fn attributes(&self) -> &[Attribute] {
        &self.attributes
    }

    /// Return mutable attributes.
    pub fn attributes_mut(&mut self) -> &mut Vec<Attribute> {
        &mut self.attributes
    }

    /// Add or replace an attribute with the same expanded name.
    pub fn set_attribute(&mut self, attribute: Attribute) {
        if let Some(existing) = self.attributes.iter_mut().find(|existing| {
            existing.namespace == attribute.namespace && existing.local_name == attribute.local_name
        }) {
            *existing = attribute;
        } else {
            self.attributes.push(attribute);
        }
    }

    /// Find an attribute by expanded name.
    #[must_use]
    pub fn attribute(&self, namespace: &Namespace, local_name: &str) -> Option<&str> {
        self.attributes
            .iter()
            .find(|attribute| {
                &attribute.namespace == namespace && attribute.local_name == local_name
            })
            .map(|attribute| attribute.value.as_str())
    }

    /// Return the schema-defined child nodes.
    #[must_use]
    pub fn children(&self) -> &[Node] {
        &self.children
    }

    /// Return mutable child nodes.
    ///
    /// The tree is validated before serialization, so invalid child relations
    /// introduced through this method cause saving to return an error.
    pub fn children_mut(&mut self) -> &mut Vec<Node> {
        &mut self.children
    }

    /// Add a child if this node kind permits it under the ODF schema.
    ///
    /// # Errors
    /// Returns an error when the input is malformed or a configured limit is exceeded.
    pub fn add_child(&mut self, child: Node) -> Result<()> {
        if !self.kind.allows_child(child.kind) {
            return Err(Error::InvalidFormat(format!(
                "anim:{} cannot contain anim:{}",
                self.kind.local_name(),
                child.kind.local_name()
            )));
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

    fn validate(&self, depth: usize, node_count: &mut usize) -> Result<()> {
        if depth > 128 {
            return Err(Error::InvalidFormat(
                "ODP animation nesting exceeds 128 levels".to_string(),
            ));
        }
        *node_count = node_count
            .checked_add(1)
            .ok_or_else(|| Error::InvalidFormat("ODP animation node count overflow".to_string()))?;
        if *node_count > 65_536 {
            return Err(Error::InvalidFormat(
                "ODP animation tree exceeds 65536 nodes".to_string(),
            ));
        }
        if self.attributes.len() > 256 {
            return Err(Error::InvalidFormat(
                "ODP animation node exceeds 256 attributes".to_string(),
            ));
        }
        let mut names = BTreeSet::new();
        for attribute in &self.attributes {
            validate_ncname(&attribute.local_name)?;
            validate_attribute_namespace(&attribute.namespace)?;
            validate_attribute_value(&attribute.value)?;
            let name = (&attribute.namespace, &attribute.local_name);
            if !names.insert(name) {
                return Err(Error::InvalidFormat(format!(
                    "duplicate animation attribute '{}'",
                    attribute.local_name
                )));
            }
        }
        for child in &self.children {
            if !self.kind.allows_child(child.kind) {
                return Err(Error::InvalidFormat(format!(
                    "anim:{} cannot contain anim:{}",
                    self.kind.local_name(),
                    child.kind.local_name()
                )));
            }
            child.validate(depth + 1, node_count)?;
        }
        Ok(())
    }

    pub(crate) fn collect_extension_namespaces(&self, uris: &mut BTreeSet<String>) {
        for attribute in &self.attributes {
            if let Namespace::Other(uri) = &attribute.namespace {
                uris.insert(uri.clone());
            }
        }
        for child in &self.children {
            child.collect_extension_namespaces(uris);
        }
    }

    pub(crate) fn write_xml(
        &self,
        output: &mut String,
        extensions: &BTreeMap<String, String>,
    ) -> Result<()> {
        output.push_str("<anim:");
        output.push_str(self.kind.local_name());
        for attribute in &self.attributes {
            output.push(' ');
            if let Some(prefix) = attribute.namespace.prefix(extensions)? {
                output.push_str(prefix);
                output.push(':');
            }
            output.push_str(&attribute.local_name);
            output.push_str("=\"");
            output.push_str(&escape_xml(&attribute.value));
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
        output.push_str("</anim:");
        output.push_str(self.kind.local_name());
        output.push('>');
        Ok(())
    }
}

pub(crate) fn validate_animation_roots(roots: &[Node]) -> Result<()> {
    let mut node_count = 0;
    for root in roots {
        if !root.kind.allowed_at_page_root() {
            return Err(Error::InvalidFormat(
                "anim:param is only valid below anim:command".to_string(),
            ));
        }
        root.validate(1, &mut node_count)?;
    }
    Ok(())
}

fn validate_ncname(name: &str) -> Result<()> {
    let mut chars = name.chars();
    let Some(first) = chars.next() else {
        return Err(Error::InvalidFormat(
            "animation attribute local name cannot be empty".to_string(),
        ));
    };
    if !(first == '_' || first.is_alphabetic())
        || !chars.all(|character| {
            character == '_'
                || character == '-'
                || character == '.'
                || character.is_alphanumeric()
                || character == '\u{00B7}'
                || ('\u{0300}'..='\u{036F}').contains(&character)
                || ('\u{203F}'..='\u{2040}').contains(&character)
        })
    {
        return Err(Error::InvalidFormat(format!(
            "invalid animation attribute local name '{name}'"
        )));
    }
    Ok(())
}

fn validate_attribute_namespace(namespace: &Namespace) -> Result<()> {
    if let Namespace::Other(uri) = namespace {
        if uri.is_empty() {
            return Err(Error::InvalidFormat(
                "animation extension namespace URI cannot be empty".to_string(),
            ));
        }
        if matches!(
            uri.as_str(),
            ANIMATION_NAMESPACE
                | SMIL_NAMESPACE
                | PRESENTATION_NAMESPACE
                | DRAW_NAMESPACE
                | SVG_NAMESPACE
                | XLINK_NAMESPACE
                | XML_NAMESPACE
                | "http://www.w3.org/2000/xmlns/"
        ) {
            return Err(Error::InvalidFormat(format!(
                "reserved or standard namespace URI '{uri}' cannot be stored as an animation extension namespace"
            )));
        }
        validate_xml_text(uri, "animation extension namespace URI")?;
    }
    Ok(())
}

fn validate_attribute_value(value: &str) -> Result<()> {
    if value.len() > 1_048_576 {
        return Err(Error::InvalidFormat(
            "ODP animation attribute exceeds 1 MiB".to_string(),
        ));
    }
    validate_xml_text(value, "animation attribute value")
}

fn validate_xml_text(value: &str, description: &str) -> Result<()> {
    if value.chars().any(|character| {
        !matches!(
            character,
            '\u{0009}' | '\u{000A}' | '\u{000D}' | '\u{0020}'..='\u{D7FF}' | '\u{E000}'..='\u{FFFD}' | '\u{10000}'..='\u{10FFFF}'
        )
    }) {
        return Err(Error::InvalidFormat(format!(
            "{description} contains a character forbidden by XML 1.0"
        )));
    }
    Ok(())
}
