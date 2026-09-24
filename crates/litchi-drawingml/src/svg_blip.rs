//! Source-backed support for the `[MS-ODRAWXML]` SVG blip extension.
//!
//! `asvg:svgBlip` is deliberately small in the standard: it is a
//! relationship-bearing element with the shared `AG_Blob` attributes.  The
//! public value therefore keeps the relationship metadata typed while
//! retaining unknown attributes and children for forward compatibility.  A
//! parsed value keeps its source bytes, so an untouched fragment can be
//! written without changing prefixes, attribute order, or lexical forms.

use std::{fmt, io::Write, sync::Arc};

use litchi_ooxml_common::{relationships, xml::is_ncname};
use quick_xml::{
    XmlVersion,
    events::{BytesStart, Event},
    name::{Namespace as XmlNamespace, Prefix, QName, ResolveResult},
    reader::{NsReader, Reader},
};
use thiserror::Error as ThisError;

use crate::{Error, Result};

/// Transitional `[MS-ODRAWXML]` SVG namespace.
pub const NAMESPACE: &str = "http://schemas.microsoft.com/office/drawing/2016/SVG/main";
/// The relationship namespace used by transitional and strict OOXML.
pub const RELATIONSHIP_NAMESPACE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
/// The strict OOXML relationship namespace.
pub const RELATIONSHIP_NAMESPACE_STRICT: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships";
/// The XML namespace reserved for the `xml` prefix.
pub const XML_NAMESPACE: &str = "http://www.w3.org/XML/1998/namespace";
/// The namespace reserved for namespace declaration machinery.
pub const XMLNS_NAMESPACE: &str = "http://www.w3.org/2000/xmlns/";

/// Maximum source fragment size accepted by this bounded codec.
pub const MAX_XML_BYTES: usize = 16 * 1024 * 1024;
/// Maximum relationship identifier size accepted by this codec.
pub const MAX_RELATIONSHIP_ID_BYTES: usize = 256;
/// Maximum namespace URI or prefix text retained by one fragment.
pub const MAX_NAMESPACE_BYTES: usize = 4_096;
/// Maximum decoded unknown attribute value retained by one fragment.
pub const MAX_ATTRIBUTE_VALUE_BYTES: usize = 1_048_576;
/// Maximum unknown qualified attribute name retained by one fragment.
pub const MAX_ATTRIBUTE_NAME_BYTES: usize = 4_096;
/// Maximum retained namespace declarations on one element.
pub const MAX_NAMESPACE_DECLARATIONS: usize = 256;
/// Maximum retained unknown attributes on one element.
pub const MAX_ATTRIBUTES: usize = 512;
/// Maximum retained unknown child elements on one element.
pub const MAX_CHILDREN: usize = 512;
/// Maximum nested XML depth while validating unknown children.
pub const MAX_DEPTH: usize = 128;
/// Maximum element count while validating unknown children.
pub const MAX_NODES: usize = 100_000;

/// A checked `ST_RelationshipId` value.
#[derive(Debug, Clone, PartialEq, Eq, Hash, PartialOrd, Ord)]
#[must_use]
pub struct RelationshipId(Box<str>);

impl RelationshipId {
    /// Construct a bounded XML `NCName` relationship identifier.
    ///
    /// # Errors
    ///
    /// Returns [`ValueError::RelationshipId`] when the value is empty, too
    /// long, or is not an XML `NCName`.
    pub fn new(value: impl AsRef<str>) -> std::result::Result<Self, ValueError> {
        let value = value.as_ref();
        if value.len() > MAX_RELATIONSHIP_ID_BYTES {
            return Err(ValueError::TooLong {
                field: "relationship ID",
                limit: MAX_RELATIONSHIP_ID_BYTES,
            });
        }
        if value.is_empty() || !is_ncname(value) {
            return Err(ValueError::RelationshipId {
                value: value.to_owned(),
            });
        }
        Ok(Self(value.into()))
    }

    /// Borrow the exact lexical relationship identifier.
    #[must_use]
    pub fn as_str(&self) -> &str {
        &self.0
    }
}

impl AsRef<str> for RelationshipId {
    fn as_ref(&self) -> &str {
        self.as_str()
    }
}

impl fmt::Display for RelationshipId {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(self.as_str())
    }
}

impl TryFrom<&str> for RelationshipId {
    type Error = ValueError;

    fn try_from(value: &str) -> std::result::Result<Self, Self::Error> {
        Self::new(value)
    }
}

impl TryFrom<String> for RelationshipId {
    type Error = ValueError;

    fn try_from(value: String) -> std::result::Result<Self, Self::Error> {
        Self::new(value)
    }
}

impl From<RelationshipId> for String {
    fn from(value: RelationshipId) -> Self {
        value.0.into()
    }
}

/// The two independent alternatives of the DrawingML `AG_Blob` attribute
/// group.  Most Office files use `embedded`; linked SVG resources are retained
/// as a first-class value because dropping them would change document meaning.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
#[must_use]
pub struct Reference {
    /// Package-local SVG relationship (`r:embed`).
    pub embedded: Option<RelationshipId>,
    /// External SVG relationship (`r:link`).
    pub linked: Option<RelationshipId>,
}

impl Reference {
    /// Construct an empty relationship reference.
    #[must_use]
    pub const fn none() -> Self {
        Self {
            embedded: None,
            linked: None,
        }
    }

    /// Construct an embedded relationship reference.
    ///
    /// # Errors
    ///
    /// Returns an error when `value` is not a bounded XML relationship ID.
    pub fn embedded(value: impl AsRef<str>) -> std::result::Result<Self, ValueError> {
        Ok(Self {
            embedded: Some(RelationshipId::new(value)?),
            linked: None,
        })
    }

    /// Construct a linked relationship reference.
    ///
    /// # Errors
    ///
    /// Returns an error when `value` is not a bounded XML relationship ID.
    pub fn linked(value: impl AsRef<str>) -> std::result::Result<Self, ValueError> {
        Ok(Self {
            embedded: None,
            linked: Some(RelationshipId::new(value)?),
        })
    }

    /// Return whether neither relationship attribute is present.
    #[must_use]
    pub const fn is_empty(&self) -> bool {
        self.embedded.is_none() && self.linked.is_none()
    }
}

/// A namespace declaration retained from a parsed `svgBlip` element.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
#[must_use]
pub struct Namespace {
    prefix: Option<Box<str>>,
    uri: Box<str>,
}

impl Namespace {
    /// Construct a namespace declaration; `None` denotes the default prefix.
    /// An empty URI clears the default namespace, as in `xmlns=""`.
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid prefix, an empty prefixed URI, or an
    /// overlong URI.
    pub fn new(
        prefix: Option<&str>,
        uri: impl AsRef<str>,
    ) -> std::result::Result<Self, ValueError> {
        if let Some(prefix) = prefix
            && (!prefix.is_empty() && (prefix.len() > MAX_NAMESPACE_BYTES || !is_ncname(prefix)))
        {
            if prefix.len() > MAX_NAMESPACE_BYTES {
                return Err(ValueError::TooLong {
                    field: "namespace prefix",
                    limit: MAX_NAMESPACE_BYTES,
                });
            }
            return Err(ValueError::NamespacePrefix {
                value: prefix.to_owned(),
            });
        }
        if prefix == Some("xmlns") {
            return Err(ValueError::NamespaceBinding);
        }
        let uri = uri.as_ref();
        if uri.is_empty() && prefix.is_some_and(|value| !value.is_empty()) {
            return Err(ValueError::EmptyNamespace);
        }
        if uri.len() > MAX_NAMESPACE_BYTES {
            return Err(ValueError::TooLong {
                field: "namespace URI",
                limit: MAX_NAMESPACE_BYTES,
            });
        }
        if uri == XMLNS_NAMESPACE
            || (prefix == Some("xml") && uri != XML_NAMESPACE)
            || (prefix != Some("xml") && uri == XML_NAMESPACE)
        {
            return Err(ValueError::NamespaceBinding);
        }
        Ok(Self {
            prefix: prefix.filter(|value| !value.is_empty()).map(Into::into),
            uri: uri.into(),
        })
    }

    /// Borrow the prefix; `None` denotes the default namespace.
    #[must_use]
    pub fn prefix(&self) -> Option<&str> {
        self.prefix.as_deref()
    }

    /// Borrow the namespace URI.
    #[must_use]
    pub fn uri(&self) -> &str {
        &self.uri
    }
}

/// Maximum number of active bindings retained by one contextual fragment.
///
/// A context is persistent: extending it stores only the declarations on the
/// new XML element and points at the previous context.  The bound therefore
/// limits both resolver work and the amount of namespace state that an export
/// may have to visit without requiring each fragment to own a flattened copy.
pub const MAX_CONTEXT_BINDINGS: usize = 16 * 1024;
/// Maximum number of XML scope frames retained by one contextual handle.
///
/// Empty declaration frames have no observable namespace effect and are
/// collapsed by [`NamespaceContext::child`], but a host can still construct
/// a deeply nested chain by adding one binding at a time.  Keep that walk
/// independently bounded even when most bindings are later shadowed.
pub const MAX_CONTEXT_DEPTH: usize = 1_024;

/// An immutable, persistent namespace scope for a source-backed fragment.
///
/// [`codec::read`] remains the complete-document codec and retains its exact
/// source fast path.  Hosts that index a nested element can instead retain a
/// `NamespaceContext` handle and call [`codec::read_contextual`].  Cloning this
/// value is cheap: declarations are stored only on the scope that introduced
/// them and each descendant shares its ancestors through `Arc`.
#[derive(Clone)]
#[must_use]
pub struct NamespaceContext {
    node: Arc<NamespaceContextNode>,
}

#[derive(Clone)]
struct NamespaceContextNode {
    parent: Option<Arc<NamespaceContextNode>>,
    declarations: Arc<[Namespace]>,
    bindings: usize,
    depth: usize,
}

impl fmt::Debug for NamespaceContext {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("NamespaceContext")
            .field("bindings", &self.binding_count())
            .field("depth", &self.depth())
            .finish()
    }
}

impl PartialEq for NamespaceContext {
    fn eq(&self, other: &Self) -> bool {
        if Arc::ptr_eq(&self.node, &other.node) {
            return true;
        }
        context_nodes_equal(&self.node, &other.node)
    }
}

impl Eq for NamespaceContext {}

fn context_nodes_equal(
    left: &Arc<NamespaceContextNode>,
    right: &Arc<NamespaceContextNode>,
) -> bool {
    if Arc::ptr_eq(left, right) {
        return true;
    }
    if left.bindings != right.bindings || left.depth != right.depth {
        return false;
    }
    if left.declarations.as_ref() != right.declarations.as_ref() {
        return false;
    }
    match (&left.parent, &right.parent) {
        (Some(left), Some(right)) => context_nodes_equal(left, right),
        (None, None) => true,
        _ => false,
    }
}

impl NamespaceContext {
    /// Return an empty namespace scope.
    #[must_use]
    pub fn empty() -> Self {
        Self {
            node: Arc::new(NamespaceContextNode {
                parent: None,
                declarations: Arc::from([]),
                bindings: 0,
                depth: 0,
            }),
        }
    }

    /// Extend a scope with declarations from one XML element.
    ///
    /// Declarations are validated before the new immutable node is published.
    /// Duplicate prefixes on one element are rejected, while a declaration in
    /// a child scope may shadow an ancestor (including `xmlns=""`).
    pub fn child<I>(&self, declarations: I) -> Result<Self>
    where
        I: IntoIterator<Item = Namespace>,
    {
        let declarations = declarations.into_iter();
        let mut local = Vec::new();
        local
            .try_reserve(4)
            .map_err(|source| allocation("SVG namespace context declarations", source))?;
        for namespace in declarations {
            if local
                .iter()
                .any(|old: &Namespace| old.prefix() == namespace.prefix())
            {
                return Err(invalid("SVG namespace context has duplicate declarations"));
            }
            if local.len() >= MAX_NAMESPACE_DECLARATIONS {
                return Err(limit(
                    "SVG namespace context declarations",
                    MAX_NAMESPACE_DECLARATIONS,
                ));
            }
            // `Namespace` is a public validated value.  Re-run the check at
            // this boundary so values assembled by a host cannot smuggle an
            // invalid reserved binding into a shared context.
            let _ = Namespace::new(namespace.prefix(), namespace.uri()).map_err(value_error)?;
            local
                .try_reserve(1)
                .map_err(|source| allocation("SVG namespace context declarations", source))?;
            local.push(namespace);
        }
        let local_count = local.len();
        if local_count == 0 {
            return Ok(self.clone());
        }
        let bindings = self
            .binding_count()
            .checked_add(local_count)
            .ok_or_else(|| limit("SVG namespace context bindings", MAX_CONTEXT_BINDINGS))?;
        if bindings > MAX_CONTEXT_BINDINGS {
            return Err(limit(
                "SVG namespace context bindings",
                MAX_CONTEXT_BINDINGS,
            ));
        }
        let depth = self
            .depth()
            .checked_add(1)
            .ok_or_else(|| limit("SVG namespace context depth", MAX_CONTEXT_DEPTH))?;
        if depth > MAX_CONTEXT_DEPTH {
            return Err(limit("SVG namespace context depth", MAX_CONTEXT_DEPTH));
        }
        Ok(Self {
            node: Arc::new(NamespaceContextNode {
                parent: Some(Arc::clone(&self.node)),
                declarations: Arc::from(local.into_boxed_slice()),
                bindings,
                depth,
            }),
        })
    }

    /// Number of declarations in this scope chain, including shadowed ones.
    #[must_use]
    pub fn binding_count(&self) -> usize {
        self.node.bindings
    }

    /// Number of XML scope frames represented by this handle.
    #[must_use]
    pub fn depth(&self) -> usize {
        self.node.depth
    }

    /// Whether the scope has no declarations.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.binding_count() == 0
    }

    /// Return whether two handles point at the same immutable scope node.
    ///
    /// This is useful to bounded host indexes that need to assert context
    /// sharing without exposing the internal `Arc` or walking namespace bytes.
    #[must_use]
    pub fn shares_storage(&self, other: &Self) -> bool {
        Arc::ptr_eq(&self.node, &other.node)
    }

    /// Visit each effective visible binding exactly once.
    ///
    /// The callback observes the current value for every prefix, including a
    /// default-namespace undeclaration with an empty URI.  The implementation
    /// retains only a bounded vector of scope-node references while walking;
    /// it never materializes a flattened prefix/URI byte table.
    pub fn visit_visible<F>(&self, mut visitor: F) -> Result<()>
    where
        F: FnMut(Option<&str>, &str),
    {
        let mut nodes = Vec::<&NamespaceContextNode>::new();
        nodes
            .try_reserve(self.depth())
            .map_err(|source| allocation("SVG namespace context walk", source))?;
        let mut cursor = Some(self.node.as_ref());
        while let Some(node) = cursor {
            nodes.push(node);
            cursor = node.parent.as_deref();
        }
        // The newest node is first in `nodes`; visit the chain from oldest to
        // newest while skipping a prefix that a descendant shadows.
        for (index, node) in nodes.iter().enumerate().rev() {
            for declaration in node.declarations.iter() {
                if nodes[..index].iter().any(|descendant| {
                    descendant
                        .declarations
                        .iter()
                        .any(|item| item.prefix() == declaration.prefix())
                }) {
                    continue;
                }
                visitor(declaration.prefix(), declaration.uri());
            }
        }
        Ok(())
    }

    fn has_visible_binding(&self, prefix: Option<&str>, uri: &str) -> bool {
        matches!(self.resolve(prefix, true), ContextResolution::Bound(candidate) if candidate == uri)
    }

    fn resolve(&self, prefix: Option<&str>, element: bool) -> ContextResolution<'_> {
        // The default namespace applies to element names only.  An
        // unprefixed attribute is always in no namespace, even when a
        // default binding is visible in the surrounding scope.
        if !element && prefix.is_none() {
            return ContextResolution::Unknown;
        }
        if prefix == Some("xml") {
            return ContextResolution::Bound(XML_NAMESPACE);
        }
        let mut cursor = Some(self.node.as_ref());
        while let Some(node) = cursor {
            if let Some(namespace) = node
                .declarations
                .iter()
                .rev()
                .find(|namespace| namespace.prefix() == prefix)
            {
                if namespace.uri().is_empty() {
                    return ContextResolution::Unbound;
                }
                return ContextResolution::Bound(namespace.uri());
            }
            cursor = node.parent.as_deref();
        }
        ContextResolution::Unknown
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum ContextResolution<'a> {
    Bound(&'a str),
    Unbound,
    Unknown,
}

/// An unknown attribute retained for a future extension vocabulary.
#[derive(Debug, Clone, PartialEq, Eq)]
#[must_use]
pub struct Attribute {
    name: Box<str>,
    value: Box<str>,
}

impl Attribute {
    /// Construct a retained XML attribute.
    ///
    /// # Errors
    ///
    /// Returns an error when the name is not a qualified XML name.
    pub fn new(
        name: impl AsRef<str>,
        value: impl AsRef<str>,
    ) -> std::result::Result<Self, ValueError> {
        let name = name.as_ref();
        if name.len() > MAX_ATTRIBUTE_NAME_BYTES {
            return Err(ValueError::TooLong {
                field: "attribute name",
                limit: MAX_ATTRIBUTE_NAME_BYTES,
            });
        }
        if !litchi_ooxml_common::xml_name::is_qualified_name(name)
            || name == "xmlns"
            || name.starts_with("xmlns:")
        {
            return Err(ValueError::AttributeName {
                value: name.to_owned(),
            });
        }
        if value.as_ref().len() > MAX_ATTRIBUTE_VALUE_BYTES {
            return Err(ValueError::TooLong {
                field: "attribute value",
                limit: MAX_ATTRIBUTE_VALUE_BYTES,
            });
        }
        Ok(Self {
            name: name.into(),
            value: value.as_ref().into(),
        })
    }

    /// Borrow the qualified attribute name.
    #[must_use]
    pub fn name(&self) -> &str {
        &self.name
    }

    /// Borrow the decoded attribute value.
    #[must_use]
    pub fn value(&self) -> &str {
        &self.value
    }
}

/// An unmodeled child content item retained in order for forward compatibility.
#[derive(Debug, Clone, PartialEq, Eq)]
#[must_use]
pub struct Child {
    xml: Box<[u8]>,
}

impl Child {
    /// Borrow the exact child content bytes.
    #[must_use]
    pub fn as_bytes(&self) -> &[u8] {
        &self.xml
    }
}

/// Failure to construct a bounded SVG blip scalar.
#[derive(Debug, Clone, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum ValueError {
    /// A relationship identifier is invalid or exceeds the bound.
    #[error("invalid SVG relationship ID '{value}'")]
    RelationshipId { value: String },
    /// A namespace prefix is invalid.
    #[error("invalid SVG namespace prefix '{value}'")]
    NamespacePrefix { value: String },
    /// A prefixed namespace URI is empty.
    #[error("prefixed SVG namespace URI is empty")]
    EmptyNamespace,
    /// An unknown attribute name is not an XML qualified name.
    #[error("invalid SVG attribute name '{value}'")]
    AttributeName { value: String },
    /// A namespace declaration uses a reserved XML namespace binding.
    #[error("invalid reserved SVG namespace binding")]
    NamespaceBinding,
    /// A retained scalar exceeds its explicit memory bound.
    #[error("SVG {field} exceeds the limit of {limit} bytes")]
    TooLong { field: &'static str, limit: usize },
}

/// Typed and lossless metadata for one `asvg:svgBlip` element.
#[derive(Clone)]
#[must_use]
pub struct SvgBlip {
    reference: Reference,
    namespaces: Arc<[Namespace]>,
    attributes: Vec<Attribute>,
    children: Vec<Child>,
    prefix: Option<Box<str>>,
    source: Option<Arc<[u8]>>,
    contextual: Option<ContextualSource>,
}

#[derive(Clone)]
struct ContextualSource {
    raw: Arc<[u8]>,
    context: NamespaceContext,
    dirty: bool,
}

impl fmt::Debug for SvgBlip {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("SvgBlip")
            .field("reference", &self.reference)
            .field("namespace_count", &self.namespaces.len())
            .field("attribute_count", &self.attributes.len())
            .field("child_count", &self.children.len())
            .field("prefix", &self.prefix)
            .field(
                "source_len",
                &self.source.as_ref().map(|source| source.len()),
            )
            .field(
                "contextual",
                &self.contextual.as_ref().map(|source| ContextualDebug {
                    raw_len: source.raw.len(),
                    context: &source.context,
                    dirty: source.dirty,
                }),
            )
            .finish()
    }
}

struct ContextualDebug<'a> {
    raw_len: usize,
    context: &'a NamespaceContext,
    dirty: bool,
}

impl fmt::Debug for ContextualDebug<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("ContextualSource")
            .field("raw_len", &self.raw_len)
            .field("context", self.context)
            .field("dirty", &self.dirty)
            .finish()
    }
}

impl PartialEq for SvgBlip {
    fn eq(&self, other: &Self) -> bool {
        self.reference == other.reference
            && self.namespaces == other.namespaces
            && self.attributes == other.attributes
            && self.children == other.children
            && self.prefix == other.prefix
            && source_equal(&self.source, &other.source)
            && contextual_equal(&self.contextual, &other.contextual)
    }
}

impl Eq for SvgBlip {}

impl SvgBlip {
    /// Construct a new empty SVG blip with the supplied relationship metadata.
    ///
    /// # Errors
    ///
    /// Returns [`crate::Error::Invalid`] if the reference is not valid.
    pub fn new(reference: Reference) -> Result<Self> {
        validate_reference(&reference)?;
        Ok(Self {
            reference,
            namespaces: Arc::from([]),
            attributes: Vec::new(),
            children: Vec::new(),
            prefix: Some("asvg".into()),
            source: None,
            contextual: None,
        })
    }

    /// Return the typed relationship metadata.
    #[must_use]
    pub const fn reference(&self) -> &Reference {
        &self.reference
    }

    /// Return the embedded SVG relationship, when present.
    #[must_use]
    pub fn embedded(&self) -> Option<&RelationshipId> {
        self.reference.embedded.as_ref()
    }

    /// Return the linked SVG relationship, when present.
    #[must_use]
    pub fn linked(&self) -> Option<&RelationshipId> {
        self.reference.linked.as_ref()
    }

    /// Replace relationship metadata and invalidate the source fast path.
    ///
    /// # Errors
    ///
    /// Returns an error when the reference reuses one relationship ID for both
    /// alternatives.
    pub fn set_reference(&mut self, reference: Reference) -> Result<()> {
        validate_reference(&reference)?;
        self.reference = reference;
        self.source = None;
        if let Some(contextual) = &mut self.contextual {
            contextual.dirty = true;
        }
        Ok(())
    }

    /// Borrow retained namespace declarations.
    #[must_use]
    pub fn namespaces(&self) -> &[Namespace] {
        &self.namespaces
    }

    /// Borrow retained unknown attributes.
    #[must_use]
    pub fn attributes(&self) -> &[Attribute] {
        &self.attributes
    }

    /// Borrow retained unknown children.
    #[must_use]
    pub fn children(&self) -> &[Child] {
        &self.children
    }

    /// Borrow the exact source fragment when this value came from XML.
    #[must_use]
    pub fn source(&self) -> Option<&[u8]> {
        self.source.as_deref()
    }

    /// Borrow the exact raw fragment retained by a contextual read.
    ///
    /// A contextual fragment is intentionally not exposed through
    /// [`Self::source`]: its root may rely on declarations inherited from its
    /// host.  Use [`crate::svg_blip::codec::write_contextual`] to export an
    /// independently readable fragment with those declarations completed.
    #[must_use]
    pub fn raw_source(&self) -> Option<&[u8]> {
        self.source
            .as_deref()
            .or_else(|| self.contextual.as_ref().map(|source| source.raw.as_ref()))
    }

    /// Borrow the shared inherited namespace context of a contextual value.
    #[must_use]
    pub fn namespace_context(&self) -> Option<&NamespaceContext> {
        self.contextual.as_ref().map(|source| &source.context)
    }

    pub(crate) fn from_wire(
        reference: Reference,
        namespaces: Vec<Namespace>,
        attributes: Vec<Attribute>,
        children: Vec<Child>,
        prefix: Option<Box<str>>,
        source: Arc<[u8]>,
    ) -> Result<Self> {
        validate_reference(&reference)?;
        Ok(Self {
            reference,
            namespaces: Arc::from(namespaces.into_boxed_slice()),
            attributes,
            children,
            prefix,
            source: Some(source),
            contextual: None,
        })
    }

    pub(crate) fn from_contextual_wire(
        reference: Reference,
        namespaces: Vec<Namespace>,
        attributes: Vec<Attribute>,
        children: Vec<Child>,
        prefix: Option<Box<str>>,
        raw: Arc<[u8]>,
        context: NamespaceContext,
    ) -> Result<Self> {
        validate_reference(&reference)?;
        Ok(Self {
            reference,
            namespaces: Arc::from(namespaces.into_boxed_slice()),
            attributes,
            children,
            prefix,
            source: None,
            contextual: Some(ContextualSource {
                raw,
                context,
                dirty: false,
            }),
        })
    }
}

fn source_equal(left: &Option<Arc<[u8]>>, right: &Option<Arc<[u8]>>) -> bool {
    match (left, right) {
        (None, None) => true,
        (Some(left), Some(right)) => Arc::ptr_eq(left, right) || left.as_ref() == right.as_ref(),
        _ => false,
    }
}

fn contextual_equal(left: &Option<ContextualSource>, right: &Option<ContextualSource>) -> bool {
    match (left, right) {
        (None, None) => true,
        (Some(left), Some(right)) => {
            (Arc::ptr_eq(&left.raw, &right.raw) || left.raw.as_ref() == right.raw.as_ref())
                && left.context == right.context
                && left.dirty == right.dirty
        },
        _ => false,
    }
}

/// Read one complete SVG blip fragment.
///
/// # Errors
///
/// Returns an error for malformed XML, a wrong root namespace/name, invalid
/// relationship IDs, or exhausted resource bounds.
pub fn read(xml: &[u8]) -> Result<SvgBlip> {
    codec::read(xml)
}

/// Read one `svgBlip` element whose namespace scope is inherited from a host.
///
/// The raw element is retained separately from the contextual namespace
/// handle.  Consequently [`SvgBlip::source`] remains `None` until a caller
/// explicitly exports the value as a standalone fragment, while
/// [`SvgBlip::raw_source`] remains an exact source authority for host-level
/// byte splicing.
pub fn read_contextual(xml: &[u8], context: &NamespaceContext) -> Result<SvgBlip> {
    codec::read_contextual(xml, context)
}

/// Serialize one SVG blip fragment.
///
/// Parsed, unchanged values return their source bytes.  Modified or newly
/// constructed values use deterministic XML with retained unknown content.
///
/// # Errors
///
/// Returns an error when validation fails or the bounded output is too large.
pub fn write(value: &SvgBlip) -> Result<Vec<u8>> {
    codec::write(value)
}

/// Serialize one SVG blip fragment to a caller-provided sink.
///
/// # Errors
///
/// Returns an error when validation or the sink write fails.
pub fn write_to<W: Write>(writer: &mut W, value: &SvgBlip) -> Result<()> {
    codec::write_to(writer, value)
}

/// Serialize a contextual SVG fragment as an independently readable element.
///
/// Namespace completion is bounded before the output allocation.  The
/// contextual value's raw fragment remains unchanged after the root opening
/// tag; inherited bindings are inserted only when the exported root does not
/// already declare that prefix.
pub fn write_contextual(value: &SvgBlip, max_output_bytes: usize) -> Result<Vec<u8>> {
    codec::write_contextual(value, max_output_bytes)
}

/// Serialize a contextual SVG fragment to a caller-owned sink.
///
/// The complete output length is checked before the first sink write.  This
/// keeps a caller cap failure from producing a partial artifact.
pub fn write_contextual_to<W: Write>(
    writer: &mut W,
    value: &SvgBlip,
    max_output_bytes: usize,
) -> Result<()> {
    codec::write_contextual_to(writer, value, max_output_bytes)
}

/// XML codec for [`SvgBlip`].
pub mod codec {
    use super::*;

    /// Read one complete `svgBlip` element.
    pub fn read(xml: &[u8]) -> Result<SvgBlip> {
        if xml.len() > MAX_XML_BYTES {
            return Err(limit("SVG blip XML bytes", MAX_XML_BYTES));
        }
        let mut reader = NsReader::from_reader(xml);
        reader.config_mut().trim_text(false);
        reader.config_mut().check_end_names = true;
        reader.config_mut().check_comments = true;
        let mut buffer = Vec::new();
        let mut root_seen = false;
        let mut root_closed = false;
        let mut value = None;

        loop {
            let event_start = position(&reader, "SVG blip")?;
            let (resolved, event) = reader
                .read_resolved_event_into(&mut buffer)
                .map_err(xml_error)?;
            let resolved = resolved_namespace(&resolved)?;
            let event = event.into_owned();
            let event_end = position(&reader, "SVG blip")?;
            match event {
                Event::Decl(_) if !root_seen => {},
                Event::Start(element) if !root_seen => {
                    let (local, namespace) = resolved_name(&resolved, &element.name())?;
                    require_root(&local, &namespace, element.name().prefix())?;
                    root_seen = true;
                    let root_start = event_start;
                    let child_end = capture_element(&mut reader, &mut buffer)?;
                    value = Some(parse_root(
                        &element,
                        &reader,
                        xml,
                        root_start,
                        child_end,
                        local_prefix(&element.name())?,
                        false,
                    )?);
                    root_closed = true;
                },
                Event::Empty(element) if !root_seen => {
                    let (local, namespace) = resolved_name(&resolved, &element.name())?;
                    require_root(&local, &namespace, element.name().prefix())?;
                    root_seen = true;
                    root_closed = true;
                    value = Some(parse_root(
                        &element,
                        &reader,
                        xml,
                        event_start,
                        event_end,
                        local_prefix(&element.name())?,
                        true,
                    )?);
                },
                Event::Start(_) | Event::Empty(_) if root_seen => {
                    return Err(invalid("SVG blip fragment has more than one root element"));
                },
                Event::Start(_) | Event::Empty(_) => {
                    return Err(invalid("SVG blip fragment root is invalid"));
                },
                Event::Text(text)
                    if (!root_seen || root_closed)
                        && !text.as_ref().iter().all(u8::is_ascii_whitespace) =>
                {
                    return Err(invalid("SVG blip fragment has text outside its root"));
                },
                Event::CData(_) | Event::GeneralRef(_) if !root_seen || root_closed => {
                    return Err(invalid("SVG blip fragment has markup outside its root"));
                },
                Event::DocType(_) | Event::PI(_) => {
                    return Err(invalid(
                        "SVG blip fragment contains forbidden document markup",
                    ));
                },
                Event::Eof => break,
                Event::End(_) | Event::Decl(_) if !root_seen || root_closed => {
                    return Err(invalid("SVG blip fragment has markup outside its root"));
                },
                Event::End(_)
                | Event::Text(_)
                | Event::CData(_)
                | Event::Comment(_)
                | Event::GeneralRef(_) => {},
                Event::Decl(_) => {
                    return Err(invalid("SVG blip fragment has a late declaration"));
                },
            }
            buffer.clear();
        }

        let value = value.ok_or_else(|| invalid("SVG blip fragment has no root element"))?;
        if !root_closed {
            return Err(invalid("SVG blip fragment root is not closed"));
        }
        validate(&value)?;
        Ok(value)
    }

    /// Read one `svgBlip` element against an inherited namespace context.
    ///
    /// This path intentionally does not synthesize a complete XML fragment.
    /// The root and its attributes are resolved by a borrowed resolver over
    /// the persistent context chain, and the raw element is retained once as
    /// the contextual source authority.
    pub fn read_contextual(xml: &[u8], context: &NamespaceContext) -> Result<SvgBlip> {
        if xml.len() > MAX_XML_BYTES {
            return Err(limit("SVG blip XML bytes", MAX_XML_BYTES));
        }
        if context.binding_count() > MAX_CONTEXT_BINDINGS {
            return Err(limit(
                "SVG namespace context bindings",
                MAX_CONTEXT_BINDINGS,
            ));
        }
        if context.depth() > MAX_CONTEXT_DEPTH {
            return Err(limit("SVG namespace context depth", MAX_CONTEXT_DEPTH));
        }
        let mut reader = Reader::from_reader(xml);
        reader.config_mut().trim_text(false);
        reader.config_mut().check_end_names = true;
        let mut buffer = Vec::new();
        let mut root_seen = false;
        let mut root_closed = false;
        let mut value = None;

        loop {
            let event_start = position(&reader, "SVG blip")?;
            let event = reader
                .read_event_into(&mut buffer)
                .map_err(xml_error)?
                .into_owned();
            let event_end = position(&reader, "SVG blip")?;
            match event {
                Event::Decl(_) if !root_seen => {},
                Event::Start(element) if !root_seen => {
                    let resolver = ContextResolver::new(context, &element, reader.decoder())?;
                    let (local, namespace) = resolved_context_name(&resolver, &element.name())?;
                    require_root(&local, &namespace, element.name().prefix())?;
                    root_seen = true;
                    let child_end = capture_plain_element(&mut reader, &mut buffer)?;
                    value = Some(parse_contextual_root(
                        &element,
                        &reader,
                        xml,
                        event_start,
                        child_end,
                        local_prefix(&element.name())?,
                        false,
                        context,
                        resolver,
                    )?);
                    root_closed = true;
                },
                Event::Empty(element) if !root_seen => {
                    let resolver = ContextResolver::new(context, &element, reader.decoder())?;
                    let (local, namespace) = resolved_context_name(&resolver, &element.name())?;
                    require_root(&local, &namespace, element.name().prefix())?;
                    root_seen = true;
                    root_closed = true;
                    value = Some(parse_contextual_root(
                        &element,
                        &reader,
                        xml,
                        event_start,
                        event_end,
                        local_prefix(&element.name())?,
                        true,
                        context,
                        resolver,
                    )?);
                },
                Event::Start(_) | Event::Empty(_) if root_seen => {
                    return Err(invalid("SVG blip fragment has more than one root element"));
                },
                Event::Start(_) | Event::Empty(_) => {
                    return Err(invalid("SVG blip fragment root is invalid"));
                },
                Event::Text(text)
                    if (!root_seen || root_closed)
                        && !text.as_ref().iter().all(u8::is_ascii_whitespace) =>
                {
                    return Err(invalid("SVG blip fragment has text outside its root"));
                },
                Event::CData(_) | Event::GeneralRef(_) if !root_seen || root_closed => {
                    return Err(invalid("SVG blip fragment has markup outside its root"));
                },
                Event::DocType(_) | Event::PI(_) => {
                    return Err(invalid(
                        "SVG blip fragment contains forbidden document markup",
                    ));
                },
                Event::Eof => break,
                Event::End(_) | Event::Decl(_) if !root_seen || root_closed => {
                    return Err(invalid("SVG blip fragment has markup outside its root"));
                },
                Event::End(_)
                | Event::Text(_)
                | Event::CData(_)
                | Event::Comment(_)
                | Event::GeneralRef(_) => {},
                Event::Decl(_) => {
                    return Err(invalid("SVG blip fragment has a late declaration"));
                },
            }
            buffer.clear();
        }

        let value = value.ok_or_else(|| invalid("SVG blip fragment has no root element"))?;
        if !root_closed {
            return Err(invalid("SVG blip fragment root is not closed"));
        }
        validate(&value)?;
        Ok(value)
    }

    /// Serialize one SVG blip fragment.
    pub fn write(value: &SvgBlip) -> Result<Vec<u8>> {
        validate(value)?;
        if let Some(source) = &value.source {
            return Ok(source.to_vec());
        }
        if value.contextual.is_some() {
            return write_contextual(value, MAX_XML_BYTES);
        }
        let plan = output_plan(value)?;
        let size = serialized_size(value, &plan)?;
        let mut output = Vec::new();
        output.reserve_exact(size);
        write_inner(&mut output, value, &plan)?;
        debug_assert_eq!(output.len(), size);
        Ok(output)
    }

    /// Serialize a contextual value with an explicit finite output cap.
    pub fn write_contextual(value: &SvgBlip, max_output_bytes: usize) -> Result<Vec<u8>> {
        validate(value)?;
        let maximum = max_output_bytes.min(MAX_XML_BYTES);
        let Some(contextual) = value.contextual.as_ref() else {
            return write_standalone_bounded(value, maximum);
        };
        if !contextual.dirty {
            return complete_raw(value, contextual, maximum);
        }
        let plan = output_plan(value)?;
        let size = serialized_size_with_limit(value, &plan, maximum)?;
        let mut output = Vec::new();
        output
            .try_reserve_exact(size)
            .map_err(|source| allocation("SVG contextual output", source))?;
        write_inner(&mut output, value, &plan)?;
        debug_assert_eq!(output.len(), size);
        Ok(output)
    }

    /// Stream a contextual value after checking the complete output size.
    pub fn write_contextual_to<W: Write>(
        writer: &mut W,
        value: &SvgBlip,
        max_output_bytes: usize,
    ) -> Result<()> {
        validate(value)?;
        let maximum = max_output_bytes.min(MAX_XML_BYTES);
        let Some(contextual) = value.contextual.as_ref() else {
            if let Some(source) = &value.source {
                if source.len() > maximum {
                    return Err(limit("SVG blip output bytes", maximum));
                }
                writer.write_all(source)?;
                return Ok(());
            }
            let plan = output_plan(value)?;
            let _ = serialized_size_with_limit(value, &plan, maximum)?;
            write_inner_to(writer, value, &plan)?;
            return Ok(());
        };
        if !contextual.dirty {
            let plan = complete_raw_plan(value, contextual, maximum)?;
            writer.write_all(&contextual.raw[..plan.insertion])?;
            write_context_declarations(writer, plan.context, &plan.declared)?;
            writer.write_all(&contextual.raw[plan.insertion..])?;
            return Ok(());
        }
        let plan = output_plan(value)?;
        let _ = serialized_size_with_limit(value, &plan, maximum)?;
        write_inner_to(writer, value, &plan)?;
        Ok(())
    }

    /// Serialize one SVG blip fragment to a caller-owned sink.
    pub fn write_to<W: Write>(writer: &mut W, value: &SvgBlip) -> Result<()> {
        writer.write_all(&write(value)?)?;
        Ok(())
    }

    fn write_standalone_bounded(value: &SvgBlip, maximum: usize) -> Result<Vec<u8>> {
        if let Some(source) = &value.source {
            if source.len() > maximum {
                return Err(limit("SVG blip output bytes", maximum));
            }
            let mut output = Vec::new();
            output
                .try_reserve_exact(source.len())
                .map_err(|source| allocation("SVG contextual output", source))?;
            output.extend_from_slice(source);
            return Ok(output);
        }
        let plan = output_plan(value)?;
        let size = serialized_size_with_limit(value, &plan, maximum)?;
        let mut output = Vec::new();
        output
            .try_reserve_exact(size)
            .map_err(|source| allocation("SVG contextual output", source))?;
        write_inner(&mut output, value, &plan)?;
        debug_assert_eq!(output.len(), size);
        Ok(output)
    }

    fn output_plan(value: &SvgBlip) -> Result<OutputPlan> {
        let root_prefix = choose_root_prefix(value)?;
        let relationship_prefix = if value.reference.is_empty() {
            None
        } else {
            Some(choose_relationship_prefix(value)?.into_boxed_str())
        };
        Ok(OutputPlan {
            root_prefix,
            relationship_prefix,
        })
    }

    fn choose_root_prefix(value: &SvgBlip) -> Result<Option<Box<str>>> {
        let requested = value.prefix.as_deref();
        if !has_namespace_binding(value, requested, NAMESPACE)
            && has_any_namespace_prefix(value, requested)?
        {
            return Ok(Some(
                next_prefix(value, requested.unwrap_or("asvg"))?.into(),
            ));
        }
        Ok(requested.map(Into::into))
    }

    fn choose_relationship_prefix(value: &SvgBlip) -> Result<String> {
        if let Some(namespace) = value.namespaces.iter().find(|namespace| {
            namespace.prefix() == Some("r") && is_relationship_uri(namespace.uri())
        }) {
            return Ok(namespace.prefix().unwrap_or("r").to_owned());
        }
        if let Some(namespace) = value
            .namespaces
            .iter()
            .find(|namespace| is_relationship_uri(namespace.uri()) && namespace.prefix().is_some())
        {
            return Ok(namespace.prefix().unwrap_or("r").to_owned());
        }
        if let Some(prefix) = find_context_relationship_prefix(value)? {
            return Ok(prefix);
        }
        if !has_any_namespace_prefix(value, Some("r"))? {
            return Ok("r".to_owned());
        }
        next_prefix(value, "r")
    }

    fn next_prefix(value: &SvgBlip, base: &str) -> Result<String> {
        let base = if base.is_empty() { "asvg" } else { base };
        for suffix in 2..=MAX_NAMESPACE_DECLARATIONS + 2 {
            let candidate = format!("{base}{suffix}");
            if candidate.len() > MAX_NAMESPACE_BYTES {
                return Err(limit("SVG namespace prefix", MAX_NAMESPACE_BYTES));
            }
            if !has_any_namespace_prefix(value, Some(candidate.as_str()))? {
                return Ok(candidate);
            }
        }
        Err(limit("SVG namespace prefixes", MAX_NAMESPACE_DECLARATIONS))
    }

    fn has_namespace_binding(value: &SvgBlip, prefix: Option<&str>, uri: &str) -> bool {
        if let Some(namespace) = value
            .namespaces
            .iter()
            .find(|namespace| namespace.prefix() == prefix)
        {
            return namespace.uri() == uri;
        }
        let Some(context) = value.contextual.as_ref().map(|source| &source.context) else {
            return false;
        };
        context.has_visible_binding(prefix, uri)
    }

    fn has_relationship_binding(value: &SvgBlip, prefix: &str) -> bool {
        has_namespace_binding(value, Some(prefix), RELATIONSHIP_NAMESPACE)
            || has_namespace_binding(value, Some(prefix), RELATIONSHIP_NAMESPACE_STRICT)
    }

    fn has_any_namespace_prefix(value: &SvgBlip, prefix: Option<&str>) -> Result<bool> {
        if value
            .namespaces
            .iter()
            .any(|namespace| namespace.prefix() == prefix)
        {
            return Ok(true);
        }
        let Some(context) = value.contextual.as_ref().map(|source| &source.context) else {
            return Ok(false);
        };
        let mut found = false;
        context.visit_visible(|candidate, _| {
            if candidate == prefix {
                found = true;
            }
        })?;
        Ok(found)
    }

    fn find_context_relationship_prefix(value: &SvgBlip) -> Result<Option<String>> {
        let Some(context) = value.contextual.as_ref().map(|source| &source.context) else {
            return Ok(None);
        };
        let mut result = None;
        context.visit_visible(|prefix, uri| {
            if result.is_none()
                && prefix.is_some()
                && is_relationship_uri(uri)
                && !value
                    .namespaces
                    .iter()
                    .any(|namespace| namespace.prefix() == prefix)
            {
                result = prefix.map(str::to_owned);
            }
        })?;
        Ok(result)
    }

    fn is_relationship_uri(uri: &str) -> bool {
        uri == RELATIONSHIP_NAMESPACE || uri == RELATIONSHIP_NAMESPACE_STRICT
    }

    fn serialized_size(value: &SvgBlip, plan: &OutputPlan) -> Result<usize> {
        let mut size = 1usize;
        let mut namespace_count = 0usize;
        if let Some(prefix) = plan.root_prefix.as_deref() {
            add_size(&mut size, prefix.len() + 1)?;
        }
        add_size(&mut size, b"svgBlip".len())?;
        for namespace in value.namespaces.iter() {
            add_namespace_count(&mut namespace_count)?;
            namespace_size(&mut size, namespace.prefix(), namespace.uri())?;
        }
        visit_inherited_namespaces(value, |prefix, uri| {
            add_namespace_count(&mut namespace_count)?;
            namespace_size(&mut size, prefix, uri)
        })?;
        if !has_namespace_binding(value, plan.root_prefix.as_deref(), NAMESPACE) {
            add_namespace_count(&mut namespace_count)?;
            namespace_size(&mut size, plan.root_prefix.as_deref(), NAMESPACE)?;
        }
        if let Some(prefix) = plan.relationship_prefix.as_deref()
            && !has_relationship_binding(value, prefix)
        {
            add_namespace_count(&mut namespace_count)?;
            namespace_size(&mut size, Some(prefix), RELATIONSHIP_NAMESPACE)?;
        }
        for attribute in &value.attributes {
            add_size(&mut size, 4 + attribute.name().len())?;
            add_size(&mut size, escaped_size(attribute.value())?)?;
        }
        if let Some(prefix) = plan.relationship_prefix.as_deref() {
            if let Some(id) = &value.reference.embedded {
                relationship_size(&mut size, prefix, "embed", id.as_str())?;
            }
            if let Some(id) = &value.reference.linked {
                relationship_size(&mut size, prefix, "link", id.as_str())?;
            }
        }
        if value.children.is_empty() {
            add_size(&mut size, 2)?;
        } else {
            add_size(&mut size, 1)?;
            for child in &value.children {
                add_size(&mut size, child.as_bytes().len())?;
            }
            add_size(&mut size, 2)?;
            if let Some(prefix) = plan.root_prefix.as_deref() {
                add_size(&mut size, prefix.len() + 1)?;
            }
            add_size(&mut size, b"svgBlip>".len())?;
        }
        Ok(size)
    }

    fn serialized_size_with_limit(
        value: &SvgBlip,
        plan: &OutputPlan,
        maximum: usize,
    ) -> Result<usize> {
        let size = serialized_size(value, plan)?;
        if size > maximum {
            return Err(limit("SVG blip output bytes", maximum));
        }
        Ok(size)
    }

    fn visit_inherited_namespaces<F>(value: &SvgBlip, mut visitor: F) -> Result<()>
    where
        F: FnMut(Option<&str>, &str) -> Result<()>,
    {
        let Some(context) = value.contextual.as_ref().map(|source| &source.context) else {
            return Ok(());
        };
        let mut result = Ok(());
        context.visit_visible(|prefix, uri| {
            if result.is_ok()
                && !value
                    .namespaces
                    .iter()
                    .any(|namespace| namespace.prefix() == prefix)
            {
                result = visitor(prefix, uri);
            }
        })?;
        result
    }

    fn namespace_size(size: &mut usize, prefix: Option<&str>, uri: &str) -> Result<()> {
        add_size(size, 6 + prefix.map_or(0, |prefix| prefix.len() + 1) + 3)?;
        add_size(size, escaped_size(uri)?)
    }

    fn add_namespace_count(count: &mut usize) -> Result<()> {
        *count = count.checked_add(1).ok_or_else(|| {
            limit(
                "SVG blip namespace declarations",
                MAX_NAMESPACE_DECLARATIONS,
            )
        })?;
        if *count > MAX_NAMESPACE_DECLARATIONS {
            return Err(limit(
                "SVG blip namespace declarations",
                MAX_NAMESPACE_DECLARATIONS,
            ));
        }
        Ok(())
    }

    fn relationship_size(size: &mut usize, prefix: &str, name: &str, value: &str) -> Result<()> {
        add_size(size, 5 + prefix.len() + name.len())?;
        add_size(size, escaped_size(value)?)
    }

    fn escaped_size(value: &str) -> Result<usize> {
        let mut size = 0usize;
        for byte in value.bytes() {
            let escaped = match byte {
                b'<' | b'>' => 4,
                b'&' => 5,
                b'\'' | b'"' => 6,
                _ => 1,
            };
            add_size(&mut size, escaped)?;
        }
        Ok(size)
    }

    fn add_size(size: &mut usize, addition: usize) -> Result<()> {
        *size = size
            .checked_add(addition)
            .ok_or_else(|| limit("SVG blip output bytes", MAX_XML_BYTES))?;
        if *size > MAX_XML_BYTES {
            return Err(limit("SVG blip output bytes", MAX_XML_BYTES));
        }
        Ok(())
    }

    struct CompleteRawPlan<'a> {
        context: &'a NamespaceContext,
        insertion: usize,
        declared: Vec<Option<Box<str>>>,
        output_len: usize,
    }

    fn complete_raw(
        value: &SvgBlip,
        contextual: &ContextualSource,
        maximum: usize,
    ) -> Result<Vec<u8>> {
        let plan = complete_raw_plan(value, contextual, maximum)?;
        let mut output = Vec::new();
        output
            .try_reserve_exact(plan.output_len)
            .map_err(|source| allocation("SVG contextual output", source))?;
        output.extend_from_slice(&contextual.raw[..plan.insertion]);
        write_context_declarations(&mut output, plan.context, &plan.declared)?;
        output.extend_from_slice(&contextual.raw[plan.insertion..]);
        debug_assert_eq!(output.len(), plan.output_len);
        Ok(output)
    }

    fn complete_raw_plan<'a>(
        value: &SvgBlip,
        contextual: &'a ContextualSource,
        maximum: usize,
    ) -> Result<CompleteRawPlan<'a>> {
        let fragment = contextual.raw.as_ref();
        if fragment.len() > maximum {
            return Err(limit("SVG blip output bytes", maximum));
        }
        let mut reader = Reader::from_reader(fragment);
        reader.config_mut().trim_text(false);
        reader.config_mut().check_end_names = true;
        let mut buffer = Vec::new();
        let event = reader.read_event_into(&mut buffer).map_err(xml_error)?;
        let empty = matches!(&event, Event::Empty(_));
        let element = match event {
            Event::Start(element) | Event::Empty(element) => element,
            _ => return Err(invalid("SVG contextual fragment has no root element")),
        };
        let root_end = position(&reader, "SVG contextual fragment")?;
        let insertion = if empty {
            root_end
                .checked_sub(2)
                .ok_or_else(|| invalid("SVG contextual self-closing root is truncated"))?
        } else {
            root_end
                .checked_sub(1)
                .ok_or_else(|| invalid("SVG contextual root is truncated"))?
        };
        let mut declared = Vec::new();
        for attribute in element.attributes() {
            let attribute = attribute.map_err(xml_error)?;
            let Some(prefix) = attribute.key.as_namespace_binding() else {
                continue;
            };
            let prefix = match prefix {
                quick_xml::name::PrefixDeclaration::Default => None,
                quick_xml::name::PrefixDeclaration::Named(prefix) => {
                    Some(std::str::from_utf8(prefix).map_err(xml_error)?.into())
                },
            };
            declared
                .try_reserve(1)
                .map_err(|source| allocation("SVG contextual declarations", source))?;
            declared.push(prefix);
        }
        let mut additions = 0usize;
        let mut namespace_count = declared.len();
        if namespace_count > MAX_NAMESPACE_DECLARATIONS {
            return Err(limit(
                "SVG blip namespace declarations",
                MAX_NAMESPACE_DECLARATIONS,
            ));
        }
        let mut planning_error = None;
        contextual.context.visit_visible(|prefix, uri| {
            if planning_error.is_some()
                || declared
                    .iter()
                    .any(|declared| declared.as_deref() == prefix)
            {
                return;
            }
            if namespace_count >= MAX_NAMESPACE_DECLARATIONS {
                planning_error = Some(limit(
                    "SVG blip namespace declarations",
                    MAX_NAMESPACE_DECLARATIONS,
                ));
                return;
            }
            namespace_count += 1;
            let escaped = match escaped_size(uri) {
                Ok(escaped) => escaped,
                Err(error) => {
                    planning_error = Some(error);
                    return;
                },
            };
            let fixed = if prefix.is_none() { 9 } else { 10 };
            additions = match additions
                .checked_add(fixed)
                .and_then(|size| size.checked_add(prefix.map_or(0, str::len)))
                .and_then(|size| size.checked_add(escaped))
            {
                Some(size) => size,
                None => {
                    planning_error = Some(invalid("SVG contextual namespace output overflows"));
                    return;
                },
            };
        })?;
        if let Some(error) = planning_error {
            return Err(error);
        }
        let output_len = fragment
            .len()
            .checked_add(additions)
            .ok_or_else(|| invalid("SVG contextual output overflows"))?;
        if output_len > maximum {
            return Err(limit("SVG blip output bytes", maximum));
        }
        // Keep this argument in the plan so validation callers cannot
        // accidentally complete a value other than the contextual source.
        let _ = value;
        Ok(CompleteRawPlan {
            context: &contextual.context,
            insertion,
            declared,
            output_len,
        })
    }

    fn write_context_declarations<W: Write>(
        writer: &mut W,
        context: &NamespaceContext,
        declared: &[Option<Box<str>>],
    ) -> Result<()> {
        let mut write_error = None;
        context.visit_visible(|prefix, uri| {
            if write_error.is_some()
                || declared
                    .iter()
                    .any(|declared| declared.as_deref() == prefix)
            {
                return;
            }
            if let Err(error) = write_namespace_to(writer, prefix, uri) {
                write_error = Some(error);
            }
        })?;
        write_error.map_or(Ok(()), Err)
    }

    fn parse_root<R: std::io::BufRead>(
        element: &BytesStart<'_>,
        reader: &NsReader<R>,
        xml: &[u8],
        start: usize,
        end: usize,
        prefix: Option<Box<str>>,
        empty: bool,
    ) -> Result<SvgBlip> {
        let decoder = reader.decoder();
        let namespaces = declarations(element, decoder)?;
        if namespaces.len() > MAX_NAMESPACE_DECLARATIONS {
            return Err(limit(
                "SVG blip namespace declarations",
                MAX_NAMESPACE_DECLARATIONS,
            ));
        }
        let reference = reference(element, reader)?;
        let mut attributes = Vec::new();
        for attribute in element.attributes() {
            let attribute = attribute.map_err(xml_error)?;
            let raw_name = attribute.key.as_ref();
            if raw_name == b"xmlns"
                || raw_name.starts_with(b"xmlns:")
                || is_relationship_attribute(attribute.key, reader.resolver())
            {
                continue;
            }
            validate_attribute_namespace(attribute.key, reader.resolver(), prefix.as_deref())?;
            if attribute.value.len() > MAX_ATTRIBUTE_VALUE_BYTES {
                return Err(limit(
                    "SVG blip attribute value bytes",
                    MAX_ATTRIBUTE_VALUE_BYTES,
                ));
            }
            let name = std::str::from_utf8(raw_name).map_err(xml_error)?;
            let value = attribute
                .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
                .map_err(xml_error)?
                .into_owned();
            attributes.push(Attribute::new(name, value).map_err(value_error)?);
            if attributes.len() > MAX_ATTRIBUTES {
                return Err(limit("SVG blip attributes", MAX_ATTRIBUTES));
            }
        }
        let mut children = Vec::new();
        if !empty {
            let raw = xml
                .get(start..end)
                .ok_or_else(|| invalid("SVG blip child range is outside input"))?;
            collect_children(raw, &mut children)?;
        }
        SvgBlip::from_wire(
            reference,
            namespaces,
            attributes,
            children,
            prefix,
            copy_source(
                xml.get(start..end)
                    .ok_or_else(|| invalid("SVG blip source range is outside input"))?,
            )?,
        )
    }

    struct ContextResolver<'a> {
        context: &'a NamespaceContext,
        local: Vec<Namespace>,
    }

    impl<'a> ContextResolver<'a> {
        fn new(
            context: &'a NamespaceContext,
            element: &BytesStart<'_>,
            decoder: quick_xml::encoding::Decoder,
        ) -> Result<Self> {
            Ok(Self {
                context,
                local: declarations(element, decoder)?,
            })
        }

        fn resolve(&self, prefix: Option<&str>, element: bool) -> ContextResolution<'_> {
            // XML's default namespace never applies to an unprefixed
            // attribute.  Keep that distinction visible to relationship and
            // opaque QName validation below.
            if !element && prefix.is_none() {
                return ContextResolution::Unknown;
            }
            if let Some(namespace) = self
                .local
                .iter()
                .find(|namespace| namespace.prefix() == prefix)
            {
                if namespace.uri().is_empty() {
                    return ContextResolution::Unbound;
                }
                return ContextResolution::Bound(namespace.uri());
            }
            self.context.resolve(prefix, element)
        }

        fn resolve_name(&self, name: &QName<'_>, element: bool) -> Result<ContextResolution<'_>> {
            let prefix_value = name.prefix();
            let prefix = match prefix_value.as_ref() {
                Some(prefix) => Some(std::str::from_utf8(prefix.as_ref()).map_err(xml_error)?),
                None => None,
            };
            Ok(self.resolve(prefix, element))
        }
    }

    fn resolved_context_name(
        resolver: &ContextResolver<'_>,
        name: &QName<'_>,
    ) -> Result<(String, String)> {
        let local = std::str::from_utf8(name.local_name().as_ref())
            .map(str::to_owned)
            .map_err(xml_error)?;
        let namespace = match resolver.resolve_name(name, true)? {
            ContextResolution::Bound(namespace) => namespace.to_owned(),
            ContextResolution::Unbound | ContextResolution::Unknown => String::new(),
        };
        Ok((local, namespace))
    }

    fn parse_contextual_root(
        element: &BytesStart<'_>,
        reader: &Reader<&[u8]>,
        xml: &[u8],
        start: usize,
        end: usize,
        prefix: Option<Box<str>>,
        empty: bool,
        context: &NamespaceContext,
        resolver: ContextResolver<'_>,
    ) -> Result<SvgBlip> {
        let decoder = reader.decoder();
        if resolver.local.len() > MAX_NAMESPACE_DECLARATIONS {
            return Err(limit(
                "SVG blip namespace declarations",
                MAX_NAMESPACE_DECLARATIONS,
            ));
        }
        let reference = contextual_reference(element, decoder, &resolver)?;
        let mut attributes = Vec::new();
        for attribute in element.attributes() {
            let attribute = attribute.map_err(xml_error)?;
            let raw_name = attribute.key.as_ref();
            if raw_name == b"xmlns"
                || raw_name.starts_with(b"xmlns:")
                || is_contextual_relationship_attribute(attribute.key, &resolver)
            {
                continue;
            }
            validate_contextual_attribute_namespace(attribute.key, &resolver, prefix.as_deref())?;
            if attribute.value.len() > MAX_ATTRIBUTE_VALUE_BYTES {
                return Err(limit(
                    "SVG blip attribute value bytes",
                    MAX_ATTRIBUTE_VALUE_BYTES,
                ));
            }
            let name = std::str::from_utf8(raw_name).map_err(xml_error)?;
            let value = attribute
                .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
                .map_err(xml_error)?
                .into_owned();
            attributes.push(Attribute::new(name, value).map_err(value_error)?);
            if attributes.len() > MAX_ATTRIBUTES {
                return Err(limit("SVG blip attributes", MAX_ATTRIBUTES));
            }
        }
        let mut children = Vec::new();
        if !empty {
            let raw = xml
                .get(start..end)
                .ok_or_else(|| invalid("SVG blip child range is outside input"))?;
            collect_contextual_children(raw, context, &mut children)?;
        }
        let raw = xml
            .get(start..end)
            .ok_or_else(|| invalid("SVG blip source range is outside input"))?;
        let namespaces = resolver.local;
        SvgBlip::from_contextual_wire(
            reference,
            namespaces,
            attributes,
            children,
            prefix,
            copy_source(raw)?,
            context.clone(),
        )
    }

    fn contextual_reference(
        element: &BytesStart<'_>,
        decoder: quick_xml::encoding::Decoder,
        resolver: &ContextResolver<'_>,
    ) -> Result<Reference> {
        let mut embedded = None;
        let mut linked = None;
        for attribute in element.attributes() {
            let attribute = attribute.map_err(xml_error)?;
            let local = attribute.key.local_name();
            if local.as_ref() != b"embed" && local.as_ref() != b"link" {
                continue;
            }
            if attribute.value.len() > MAX_RELATIONSHIP_ID_BYTES * 4 {
                return Err(limit(
                    "SVG relationship attribute bytes",
                    MAX_RELATIONSHIP_ID_BYTES * 4,
                ));
            }
            if !is_contextual_relationship_attribute(attribute.key, resolver) {
                continue;
            }
            let value = attribute
                .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
                .map_err(xml_error)?;
            let value = RelationshipId::new(value.as_ref()).map_err(value_error)?;
            let slot = if local.as_ref() == b"embed" {
                &mut embedded
            } else {
                &mut linked
            };
            if slot.replace(value).is_some() {
                return Err(invalid("SVG blip has duplicate relationship attributes"));
            }
        }
        Ok(Reference { embedded, linked })
    }

    fn is_contextual_relationship_attribute(
        key: QName<'_>,
        resolver: &ContextResolver<'_>,
    ) -> bool {
        let local = key.local_name();
        if local.as_ref() != b"embed" && local.as_ref() != b"link" {
            return false;
        }
        let prefix_bytes = key.prefix().map(|prefix| prefix.as_ref().to_vec());
        let prefix = prefix_bytes
            .as_deref()
            .and_then(|prefix| std::str::from_utf8(prefix).ok());
        matches!(
            resolver.resolve(prefix, false),
            ContextResolution::Bound(namespace)
                if namespace.as_bytes() == relationships::TRANSITIONAL_NAMESPACE
                    || namespace.as_bytes() == relationships::STRICT_NAMESPACE
        ) || matches!(
            resolver.resolve(prefix, false),
            ContextResolution::Unknown if prefix == Some("r")
        )
    }

    fn validate_contextual_attribute_namespace(
        key: QName<'_>,
        resolver: &ContextResolver<'_>,
        fallback_prefix: Option<&str>,
    ) -> Result<()> {
        let Some(prefix) = key.prefix() else {
            return Ok(());
        };
        let prefix = std::str::from_utf8(prefix.as_ref()).map_err(xml_error)?;
        if prefix == "xml" {
            return Ok(());
        }
        match resolver.resolve(Some(prefix), false) {
            ContextResolution::Bound(_) => Ok(()),
            ContextResolution::Unknown if fallback_prefix == Some(prefix) => Ok(()),
            _ => Err(invalid("SVG blip attribute prefix is not bound")),
        }
    }

    fn collect_children(xml: &[u8], children: &mut Vec<Child>) -> Result<()> {
        let mut reader = Reader::from_reader(xml);
        reader.config_mut().trim_text(false);
        reader.config_mut().check_end_names = true;
        let mut buffer = Vec::new();
        next_element(&mut reader, &mut buffer)?;
        loop {
            let start = position(&reader, "SVG blip child")?;
            let event = reader
                .read_event_into(&mut buffer)
                .map_err(xml_error)?
                .into_owned();
            let end = position(&reader, "SVG blip child")?;
            match event {
                Event::Start(_) => {
                    let end = capture_plain_element(&mut reader, &mut buffer)?;
                    retain_child(xml, start, end, children)?;
                },
                Event::Empty(_) => {
                    retain_child(xml, start, end, children)?;
                },
                Event::End(_) => break,
                Event::DocType(_) | Event::PI(_) | Event::Decl(_) => {
                    return Err(invalid("SVG blip contains forbidden child markup"));
                },
                Event::Eof => return Err(invalid("SVG blip is unterminated")),
                Event::Text(_) | Event::CData(_) | Event::Comment(_) | Event::GeneralRef(_) => {
                    retain_child(xml, start, end, children)?;
                },
            }
            buffer.clear();
        }
        Ok(())
    }

    /// Validate QName bindings in an opaque contextual subtree while keeping
    /// only its direct child byte ranges.  `Reader` is intentionally used for
    /// the raw fragment: unlike a synthetic namespace-completed copy, this
    /// walk resolves names against the borrowed host context and per-element
    /// declarations without allocating the inherited scope for each owner.
    fn collect_contextual_children(
        xml: &[u8],
        context: &NamespaceContext,
        children: &mut Vec<Child>,
    ) -> Result<()> {
        let mut reader = Reader::from_reader(xml);
        reader.config_mut().trim_text(false);
        reader.config_mut().check_end_names = true;
        let mut buffer = Vec::new();
        let mut scope = ContextualScope::new(context);
        let mut starts = Vec::<usize>::new();
        let mut depth = 0usize;
        let mut nodes = 0usize;
        let mut root_seen = false;
        let mut root_closed = false;

        loop {
            let start = position(&reader, "SVG contextual child")?;
            let event = reader
                .read_event_into(&mut buffer)
                .map_err(xml_error)?
                .into_owned();
            let end = position(&reader, "SVG contextual child")?;
            match event {
                Event::Start(element) => {
                    if root_closed {
                        return Err(invalid("SVG contextual child has markup outside its root"));
                    }
                    if depth == 0 {
                        if root_seen {
                            return Err(invalid("SVG contextual child has multiple roots"));
                        }
                        root_seen = true;
                    }
                    let local = declarations(&element, reader.decoder())?;
                    scope.push(local)?;
                    if depth > 0 {
                        validate_contextual_element(&scope, &element)?;
                        validate_contextual_attributes(&scope, &element)?;
                    }
                    depth = depth
                        .checked_add(1)
                        .ok_or_else(|| invalid("SVG contextual child nesting overflow"))?;
                    if depth > MAX_DEPTH {
                        return Err(limit("SVG contextual child depth", MAX_DEPTH));
                    }
                    nodes = nodes
                        .checked_add(1)
                        .ok_or_else(|| limit("SVG contextual child nodes", MAX_NODES))?;
                    if nodes > MAX_NODES {
                        return Err(limit("SVG contextual child nodes", MAX_NODES));
                    }
                    starts
                        .try_reserve(1)
                        .map_err(|source| allocation("SVG contextual child ranges", source))?;
                    starts.push(start);
                },
                Event::Empty(element) => {
                    if root_closed {
                        return Err(invalid("SVG contextual child has markup outside its root"));
                    }
                    if depth == 0 {
                        if root_seen {
                            return Err(invalid("SVG contextual child has multiple roots"));
                        }
                        root_seen = true;
                    }
                    let local = declarations(&element, reader.decoder())?;
                    scope.push(local)?;
                    if depth > 0 {
                        validate_contextual_element(&scope, &element)?;
                        validate_contextual_attributes(&scope, &element)?;
                        if depth == 1 {
                            retain_child(xml, start, end, children)?;
                        }
                    }
                    scope.pop();
                    if depth == 0 {
                        root_closed = true;
                    }
                    nodes = nodes
                        .checked_add(1)
                        .ok_or_else(|| limit("SVG contextual child nodes", MAX_NODES))?;
                    if nodes > MAX_NODES {
                        return Err(limit("SVG contextual child nodes", MAX_NODES));
                    }
                },
                Event::End(_) => {
                    if depth == 0 {
                        return Err(invalid("SVG contextual child has an unexpected end"));
                    }
                    let child_start = starts
                        .pop()
                        .ok_or_else(|| invalid("SVG contextual child start stack is empty"))?;
                    scope.pop();
                    if depth == 2 {
                        retain_child(xml, child_start, end, children)?;
                    }
                    depth -= 1;
                    if depth == 0 {
                        root_closed = true;
                    }
                },
                Event::Text(text) => {
                    if depth == 0 {
                        if !text.as_ref().iter().all(u8::is_ascii_whitespace) {
                            return Err(invalid("SVG contextual child has text outside its root"));
                        }
                    } else if depth == 1 {
                        retain_child(xml, start, end, children)?;
                    }
                },
                Event::CData(_) | Event::Comment(_) | Event::GeneralRef(_) => {
                    if depth == 0 {
                        return Err(invalid("SVG contextual child has markup outside its root"));
                    }
                    if depth == 1 {
                        retain_child(xml, start, end, children)?;
                    }
                },
                Event::DocType(_) | Event::PI(_) | Event::Decl(_) => {
                    return Err(invalid("SVG contextual child contains forbidden markup"));
                },
                Event::Eof => break,
            }
            buffer.clear();
        }
        if !root_seen || !root_closed || depth != 0 {
            return Err(invalid("SVG contextual child is not one complete XML item"));
        }
        Ok(())
    }

    struct ContextualScope<'a> {
        context: &'a NamespaceContext,
        frames: Vec<Vec<Namespace>>,
    }

    impl<'a> ContextualScope<'a> {
        fn new(context: &'a NamespaceContext) -> Self {
            Self {
                context,
                frames: Vec::new(),
            }
        }

        fn push(&mut self, declarations: Vec<Namespace>) -> Result<()> {
            self.frames
                .try_reserve(1)
                .map_err(|source| allocation("SVG contextual namespace frames", source))?;
            self.frames.push(declarations);
            Ok(())
        }

        fn pop(&mut self) {
            let _ = self.frames.pop();
        }

        fn resolve(&self, prefix: Option<&str>, element: bool) -> ContextResolution<'_> {
            if !element && prefix.is_none() {
                return ContextResolution::Unknown;
            }
            if prefix == Some("xml") {
                return ContextResolution::Bound(XML_NAMESPACE);
            }
            for declarations in self.frames.iter().rev() {
                if let Some(namespace) = declarations
                    .iter()
                    .rev()
                    .find(|namespace| namespace.prefix() == prefix)
                {
                    if namespace.uri().is_empty() {
                        return ContextResolution::Unbound;
                    }
                    return ContextResolution::Bound(namespace.uri());
                }
            }
            self.context.resolve(prefix, element)
        }

        fn resolve_name(&self, name: &QName<'_>, element: bool) -> Result<ContextResolution<'_>> {
            let prefix_value = name.prefix();
            let prefix = match prefix_value.as_ref() {
                Some(prefix) => Some(std::str::from_utf8(prefix.as_ref()).map_err(xml_error)?),
                None => None,
            };
            Ok(self.resolve(prefix, element))
        }
    }

    fn validate_contextual_element(
        scope: &ContextualScope<'_>,
        element: &BytesStart<'_>,
    ) -> Result<()> {
        let has_prefix = element.name().prefix().is_some();
        match scope.resolve_name(&element.name(), true)? {
            ContextResolution::Bound(_) => Ok(()),
            ContextResolution::Unbound | ContextResolution::Unknown if !has_prefix => Ok(()),
            ContextResolution::Unbound | ContextResolution::Unknown => {
                Err(invalid("SVG contextual element prefix is not bound"))
            },
        }
    }

    fn validate_contextual_attributes(
        scope: &ContextualScope<'_>,
        element: &BytesStart<'_>,
    ) -> Result<()> {
        for attribute in element.attributes() {
            let attribute = attribute.map_err(xml_error)?;
            let raw_name = attribute.key.as_ref();
            if raw_name == b"xmlns" || raw_name.starts_with(b"xmlns:") {
                continue;
            }
            let Some(prefix) = attribute.key.prefix() else {
                continue;
            };
            let prefix = std::str::from_utf8(prefix.as_ref()).map_err(xml_error)?;
            if !matches!(
                scope.resolve(Some(prefix), false),
                ContextResolution::Bound(_)
            ) {
                return Err(invalid("SVG contextual attribute prefix is not bound"));
            }
        }
        Ok(())
    }

    fn retain_child(xml: &[u8], start: usize, end: usize, children: &mut Vec<Child>) -> Result<()> {
        if children.len() >= MAX_CHILDREN {
            return Err(limit("SVG blip children", MAX_CHILDREN));
        }
        let raw = xml
            .get(start..end)
            .ok_or_else(|| invalid("SVG blip child range is outside input"))?;
        children
            .try_reserve(1)
            .map_err(|source| allocation("SVG blip children", source))?;
        let mut owned = Vec::new();
        owned
            .try_reserve_exact(raw.len())
            .map_err(|source| allocation("SVG blip child bytes", source))?;
        owned.extend_from_slice(raw);
        children.push(Child {
            xml: owned.into_boxed_slice(),
        });
        Ok(())
    }

    struct OutputPlan {
        root_prefix: Option<Box<str>>,
        relationship_prefix: Option<Box<str>>,
    }

    fn write_inner(output: &mut Vec<u8>, value: &SvgBlip, plan: &OutputPlan) -> Result<()> {
        output.extend_from_slice(b"<");
        if let Some(prefix) = plan.root_prefix.as_deref() {
            output.extend_from_slice(prefix.as_bytes());
            output.push(b':');
        }
        output.extend_from_slice(b"svgBlip");
        for namespace in value.namespaces.iter() {
            write_namespace(output, namespace);
        }
        write_inherited_namespaces(output, value)?;
        if !has_namespace_binding(value, plan.root_prefix.as_deref(), NAMESPACE) {
            write_namespace_value(output, plan.root_prefix.as_deref(), NAMESPACE);
        }
        if let Some(prefix) = plan.relationship_prefix.as_deref()
            && !has_relationship_binding(value, prefix)
        {
            write_namespace_value(output, Some(prefix), RELATIONSHIP_NAMESPACE);
        }
        for attribute in &value.attributes {
            output.extend_from_slice(b" ");
            output.extend_from_slice(attribute.name().as_bytes());
            output.extend_from_slice(b"=\"");
            push_escaped(output, attribute.value());
            output.push(b'"');
        }
        write_reference(
            output,
            &value.reference,
            plan.relationship_prefix.as_deref(),
        );
        if value.children.is_empty() {
            output.extend_from_slice(b"/>");
        } else {
            output.push(b'>');
            for child in &value.children {
                output.extend_from_slice(child.as_bytes());
            }
            output.extend_from_slice(b"</");
            if let Some(prefix) = plan.root_prefix.as_deref() {
                output.extend_from_slice(prefix.as_bytes());
                output.push(b':');
            }
            output.extend_from_slice(b"svgBlip>");
        }
        Ok(())
    }

    fn write_inherited_namespaces(output: &mut Vec<u8>, value: &SvgBlip) -> Result<()> {
        visit_inherited_namespaces(value, |prefix, uri| {
            write_namespace_value(output, prefix, uri);
            Ok(())
        })
    }

    fn write_inner_to<W: Write>(writer: &mut W, value: &SvgBlip, plan: &OutputPlan) -> Result<()> {
        writer.write_all(b"<")?;
        if let Some(prefix) = plan.root_prefix.as_deref() {
            writer.write_all(prefix.as_bytes())?;
            writer.write_all(b":")?;
        }
        writer.write_all(b"svgBlip")?;
        for namespace in value.namespaces.iter() {
            write_namespace_to(writer, namespace.prefix(), namespace.uri())?;
        }
        write_inherited_namespaces_to(writer, value)?;
        if !has_namespace_binding(value, plan.root_prefix.as_deref(), NAMESPACE) {
            write_namespace_to(writer, plan.root_prefix.as_deref(), NAMESPACE)?;
        }
        if let Some(prefix) = plan.relationship_prefix.as_deref()
            && !has_relationship_binding(value, prefix)
        {
            write_namespace_to(writer, Some(prefix), RELATIONSHIP_NAMESPACE)?;
        }
        for attribute in &value.attributes {
            writer.write_all(b" ")?;
            writer.write_all(attribute.name().as_bytes())?;
            writer.write_all(b"=\"")?;
            write_escaped_to(writer, attribute.value())?;
            writer.write_all(b"\"")?;
        }
        write_reference_to(
            writer,
            &value.reference,
            plan.relationship_prefix.as_deref(),
        )?;
        if value.children.is_empty() {
            writer.write_all(b"/>")?;
        } else {
            writer.write_all(b">")?;
            for child in &value.children {
                writer.write_all(child.as_bytes())?;
            }
            writer.write_all(b"</")?;
            if let Some(prefix) = plan.root_prefix.as_deref() {
                writer.write_all(prefix.as_bytes())?;
                writer.write_all(b":")?;
            }
            writer.write_all(b"svgBlip>")?;
        }
        Ok(())
    }

    fn write_inherited_namespaces_to<W: Write>(writer: &mut W, value: &SvgBlip) -> Result<()> {
        visit_inherited_namespaces(value, |prefix, uri| write_namespace_to(writer, prefix, uri))
    }

    fn write_reference(output: &mut Vec<u8>, reference: &Reference, prefix: Option<&str>) {
        let Some(prefix) = prefix else {
            return;
        };
        if let Some(id) = &reference.embedded {
            output.extend_from_slice(b" ");
            output.extend_from_slice(prefix.as_bytes());
            output.extend_from_slice(b":embed=\"");
            push_escaped(output, id.as_str());
            output.push(b'"');
        }
        if let Some(id) = &reference.linked {
            output.extend_from_slice(b" ");
            output.extend_from_slice(prefix.as_bytes());
            output.extend_from_slice(b":link=\"");
            push_escaped(output, id.as_str());
            output.push(b'"');
        }
    }

    fn write_reference_to<W: Write>(
        writer: &mut W,
        reference: &Reference,
        prefix: Option<&str>,
    ) -> Result<()> {
        let Some(prefix) = prefix else {
            return Ok(());
        };
        if let Some(id) = &reference.embedded {
            writer.write_all(b" ")?;
            writer.write_all(prefix.as_bytes())?;
            writer.write_all(b":embed=\"")?;
            write_escaped_to(writer, id.as_str())?;
            writer.write_all(b"\"")?;
        }
        if let Some(id) = &reference.linked {
            writer.write_all(b" ")?;
            writer.write_all(prefix.as_bytes())?;
            writer.write_all(b":link=\"")?;
            write_escaped_to(writer, id.as_str())?;
            writer.write_all(b"\"")?;
        }
        Ok(())
    }

    fn write_namespace(output: &mut Vec<u8>, namespace: &Namespace) {
        write_namespace_value(output, namespace.prefix(), namespace.uri());
    }

    fn write_namespace_to<W: Write>(writer: &mut W, prefix: Option<&str>, uri: &str) -> Result<()> {
        writer.write_all(b" xmlns")?;
        if let Some(prefix) = prefix {
            writer.write_all(b":")?;
            writer.write_all(prefix.as_bytes())?;
        }
        writer.write_all(b"=\"")?;
        write_escaped_to(writer, uri)?;
        writer.write_all(b"\"")?;
        Ok(())
    }

    fn write_namespace_value(output: &mut Vec<u8>, prefix: Option<&str>, uri: &str) {
        output.extend_from_slice(b" xmlns");
        if let Some(prefix) = prefix {
            output.push(b':');
            output.extend_from_slice(prefix.as_bytes());
        }
        output.extend_from_slice(b"=\"");
        push_escaped(output, uri);
        output.push(b'"');
    }

    fn push_escaped(output: &mut Vec<u8>, value: &str) {
        output.extend_from_slice(quick_xml::escape::escape(value).as_bytes());
    }

    fn write_escaped_to<W: Write>(writer: &mut W, value: &str) -> Result<()> {
        writer.write_all(quick_xml::escape::escape(value).as_bytes())?;
        Ok(())
    }

    fn reference<R: std::io::BufRead>(
        element: &BytesStart<'_>,
        reader: &NsReader<R>,
    ) -> Result<Reference> {
        for attribute in element.attributes() {
            let attribute = attribute.map_err(xml_error)?;
            if (attribute.key.local_name().as_ref() == b"embed"
                || attribute.key.local_name().as_ref() == b"link")
                && attribute.value.len() > MAX_RELATIONSHIP_ID_BYTES * 4
            {
                return Err(limit(
                    "SVG relationship attribute bytes",
                    MAX_RELATIONSHIP_ID_BYTES * 4,
                ));
            }
        }
        let embedded =
            relationships::attribute_value(element, b"embed", reader.decoder(), reader.resolver())?;
        let linked =
            relationships::attribute_value(element, b"link", reader.decoder(), reader.resolver())?;
        let embedded = embedded
            .map(RelationshipId::new)
            .transpose()
            .map_err(value_error)?;
        let linked = linked
            .map(RelationshipId::new)
            .transpose()
            .map_err(value_error)?;
        Ok(Reference { embedded, linked })
    }

    fn declarations(
        element: &BytesStart<'_>,
        decoder: quick_xml::encoding::Decoder,
    ) -> Result<Vec<Namespace>> {
        let mut result = Vec::new();
        for attribute in element.attributes() {
            let attribute = attribute.map_err(xml_error)?;
            let raw = attribute.key.as_ref();
            let prefix = if raw == b"xmlns" {
                None
            } else if let Some(prefix) = raw.strip_prefix(b"xmlns:") {
                Some(std::str::from_utf8(prefix).map_err(xml_error)?)
            } else {
                continue;
            };
            if result
                .iter()
                .any(|item: &Namespace| item.prefix() == prefix)
            {
                return Err(invalid("SVG blip has duplicate namespace declarations"));
            }
            attribute
                .value
                .len()
                .le(&MAX_NAMESPACE_BYTES)
                .then_some(())
                .ok_or_else(|| limit("SVG blip namespace URI bytes", MAX_NAMESPACE_BYTES))?;
            let uri = attribute
                .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
                .map_err(xml_error)?
                .into_owned();
            result
                .try_reserve(1)
                .map_err(|source| allocation("SVG blip namespaces", source))?;
            result.push(Namespace::new(prefix, uri).map_err(value_error)?);
        }
        Ok(result)
    }

    fn is_relationship_attribute(
        key: QName<'_>,
        resolver: &quick_xml::name::NamespaceResolver,
    ) -> bool {
        let local = key.local_name();
        if local.as_ref() != b"embed" && local.as_ref() != b"link" {
            return false;
        }
        let (namespace, _) = resolver.resolve_attribute(key);
        matches!(
            namespace,
            ResolveResult::Bound(XmlNamespace(value))
                if value == relationships::TRANSITIONAL_NAMESPACE
                    || value == relationships::STRICT_NAMESPACE
        ) || matches!(namespace, ResolveResult::Unknown(prefix) if prefix.as_slice() == b"r")
    }

    fn validate_attribute_namespace(
        key: QName<'_>,
        resolver: &quick_xml::name::NamespaceResolver,
        fallback_prefix: Option<&str>,
    ) -> Result<()> {
        let Some(prefix) = key.prefix() else {
            return Ok(());
        };
        if prefix.as_ref() == b"xml" {
            return Ok(());
        }
        let (namespace, _) = resolver.resolve_attribute(key);
        match namespace {
            ResolveResult::Bound(_) => Ok(()),
            ResolveResult::Unknown(unresolved)
                if fallback_prefix
                    .is_some_and(|fallback| fallback.as_bytes() == unresolved.as_slice()) =>
            {
                Ok(())
            },
            _ => Err(invalid("SVG blip attribute prefix is not bound")),
        }
    }

    fn require_root(local: &str, namespace: &str, prefix: Option<Prefix<'_>>) -> Result<()> {
        if local != "svgBlip"
            || (namespace != NAMESPACE
                && !(namespace.is_empty()
                    && prefix.is_some_and(|prefix| prefix.as_ref() == b"asvg")))
        {
            return Err(invalid("SVG blip root name or namespace is invalid"));
        }
        Ok(())
    }

    fn local_prefix(name: &QName<'_>) -> Result<Option<Box<str>>> {
        name.prefix()
            .map(|prefix| {
                std::str::from_utf8(prefix.as_ref())
                    .map(Into::into)
                    .map_err(xml_error)
            })
            .transpose()
    }

    fn resolved_name(resolved: &str, name: &QName<'_>) -> Result<(String, String)> {
        let local = std::str::from_utf8(name.local_name().as_ref())
            .map(str::to_owned)
            .map_err(xml_error)?;
        Ok((local, resolved.to_owned()))
    }

    fn resolved_namespace(resolved: &ResolveResult<'_>) -> Result<String> {
        match resolved {
            ResolveResult::Bound(namespace) => std::str::from_utf8(namespace.as_ref())
                .map(str::to_owned)
                .map_err(xml_error),
            ResolveResult::Unknown(_) | ResolveResult::Unbound => Ok(String::new()),
        }
    }

    fn next_element(reader: &mut Reader<&[u8]>, buffer: &mut Vec<u8>) -> Result<()> {
        loop {
            let event = reader
                .read_event_into(buffer)
                .map_err(xml_error)?
                .into_owned();
            match event {
                Event::Start(_) | Event::Empty(_) => return Ok(()),
                Event::Text(text) if !text.as_ref().iter().all(u8::is_ascii_whitespace) => {
                    return Err(invalid("SVG blip has text outside its root"));
                },
                Event::Decl(_) => {},
                Event::DocType(_) | Event::PI(_) => {
                    return Err(invalid("SVG blip contains forbidden document markup"));
                },
                Event::Eof => return Err(invalid("SVG blip has no root")),
                Event::End(_)
                | Event::Text(_)
                | Event::CData(_)
                | Event::Comment(_)
                | Event::GeneralRef(_) => {},
            }
            buffer.clear();
        }
    }

    fn capture_element<R: std::io::BufRead>(
        reader: &mut NsReader<R>,
        buffer: &mut Vec<u8>,
    ) -> Result<usize> {
        let mut depth = 1usize;
        let mut nodes = 0usize;
        loop {
            buffer.clear();
            let (_, event) = reader.read_resolved_event_into(buffer).map_err(xml_error)?;
            let event = event.into_owned();
            let end = position(reader, "SVG blip")?;
            match event {
                Event::Start(_) => {
                    depth = depth
                        .checked_add(1)
                        .ok_or_else(|| invalid("SVG blip XML nesting overflow"))?;
                    if depth > MAX_DEPTH {
                        return Err(limit("SVG blip XML depth", MAX_DEPTH));
                    }
                    nodes = nodes.saturating_add(1);
                    if nodes > MAX_NODES {
                        return Err(limit("SVG blip XML nodes", MAX_NODES));
                    }
                },
                Event::Empty(_) => {
                    nodes = nodes.saturating_add(1);
                    if nodes > MAX_NODES {
                        return Err(limit("SVG blip XML nodes", MAX_NODES));
                    }
                },
                Event::End(_) => {
                    depth = depth
                        .checked_sub(1)
                        .ok_or_else(|| invalid("SVG blip XML nesting underflow"))?;
                    if depth == 0 {
                        return Ok(end);
                    }
                },
                Event::DocType(_) | Event::PI(_) => {
                    return Err(invalid("SVG blip contains forbidden document markup"));
                },
                Event::Eof => return Err(invalid("SVG blip root is unterminated")),
                Event::Text(_)
                | Event::CData(_)
                | Event::Comment(_)
                | Event::Decl(_)
                | Event::GeneralRef(_) => {},
            }
        }
    }

    fn capture_plain_element<R: std::io::BufRead>(
        reader: &mut Reader<R>,
        buffer: &mut Vec<u8>,
    ) -> Result<usize> {
        let mut depth = 1usize;
        loop {
            buffer.clear();
            let event = reader
                .read_event_into(buffer)
                .map_err(xml_error)?
                .into_owned();
            let end = position(reader, "SVG blip child")?;
            match event {
                Event::Start(_) => {
                    depth = depth
                        .checked_add(1)
                        .ok_or_else(|| invalid("SVG blip child nesting overflow"))?;
                    if depth > MAX_DEPTH {
                        return Err(limit("SVG blip child depth", MAX_DEPTH));
                    }
                },
                Event::End(_) => {
                    depth = depth
                        .checked_sub(1)
                        .ok_or_else(|| invalid("SVG blip child nesting underflow"))?;
                    if depth == 0 {
                        return Ok(end);
                    }
                },
                Event::DocType(_) | Event::PI(_) => {
                    return Err(invalid("SVG blip child contains forbidden markup"));
                },
                Event::Eof => return Err(invalid("SVG blip child is unterminated")),
                Event::Empty(_)
                | Event::Text(_)
                | Event::CData(_)
                | Event::Comment(_)
                | Event::Decl(_)
                | Event::GeneralRef(_) => {},
            }
        }
    }

    fn position<R: std::io::BufRead>(reader: &Reader<R>, what: &str) -> Result<usize> {
        usize::try_from(reader.buffer_position())
            .map_err(|_| invalid(format!("{what} offset exceeds usize")))
    }

    fn validate(value: &SvgBlip) -> Result<()> {
        validate_reference(&value.reference)?;
        if value.namespaces.len() > MAX_NAMESPACE_DECLARATIONS {
            return Err(limit(
                "SVG blip namespace declarations",
                MAX_NAMESPACE_DECLARATIONS,
            ));
        }
        if value.attributes.len() > MAX_ATTRIBUTES {
            return Err(limit("SVG blip attributes", MAX_ATTRIBUTES));
        }
        if value.children.len() > MAX_CHILDREN {
            return Err(limit("SVG blip children", MAX_CHILDREN));
        }
        for namespace in value.namespaces.iter() {
            let _ = Namespace::new(namespace.prefix(), namespace.uri()).map_err(value_error)?;
        }
        for attribute in &value.attributes {
            let _ = Attribute::new(attribute.name(), attribute.value()).map_err(value_error)?;
        }
        for child in &value.children {
            validate_fragment(child.as_bytes())?;
        }
        if let Some(source) = &value.source {
            if source.len() > MAX_XML_BYTES {
                return Err(limit("SVG blip XML bytes", MAX_XML_BYTES));
            }
        }
        Ok(())
    }

    fn validate_fragment(xml: &[u8]) -> Result<()> {
        if xml.is_empty() || xml.len() > MAX_XML_BYTES {
            return Err(limit("SVG blip child bytes", MAX_XML_BYTES));
        }
        let mut reader = Reader::from_reader(xml);
        reader.config_mut().trim_text(false);
        reader.config_mut().check_end_names = true;
        let mut buffer = Vec::new();
        let mut depth = 0usize;
        let mut top_level_items = 0usize;
        let mut nodes = 0usize;
        loop {
            let event = reader
                .read_event_into(&mut buffer)
                .map_err(xml_error)?
                .into_owned();
            match event {
                Event::Start(_) => {
                    if depth == 0 && top_level_items > 0 {
                        return Err(invalid("SVG blip child has multiple roots"));
                    }
                    if depth == 0 {
                        top_level_items += 1;
                    }
                    depth += 1;
                    nodes += 1;
                    if depth > MAX_DEPTH || nodes > MAX_NODES {
                        return Err(limit("SVG blip child nodes", MAX_NODES));
                    }
                },
                Event::Empty(_) => {
                    if depth == 0 && top_level_items > 0 {
                        return Err(invalid("SVG blip child has multiple roots"));
                    }
                    if depth == 0 {
                        top_level_items += 1;
                    }
                    nodes += 1;
                    if nodes > MAX_NODES {
                        return Err(limit("SVG blip child nodes", MAX_NODES));
                    }
                },
                Event::End(_) => {
                    depth = depth
                        .checked_sub(1)
                        .ok_or_else(|| invalid("SVG blip child has unexpected end"))?;
                },
                Event::Text(_) | Event::CData(_) | Event::Comment(_) | Event::GeneralRef(_)
                    if depth == 0 =>
                {
                    if top_level_items > 0 {
                        return Err(invalid("SVG blip child has multiple roots"));
                    }
                    top_level_items = 1;
                },
                Event::DocType(_) | Event::Decl(_) | Event::PI(_) => {
                    return Err(invalid("SVG blip child contains document markup"));
                },
                Event::Eof => break,
                Event::Text(_) | Event::CData(_) | Event::Comment(_) | Event::GeneralRef(_) => {},
            }
            buffer.clear();
        }
        if depth != 0 || top_level_items != 1 {
            return Err(invalid("SVG blip child is not one complete XML item"));
        }
        Ok(())
    }
}

fn validate_reference(reference: &Reference) -> Result<()> {
    if let (Some(embedded), Some(linked)) = (&reference.embedded, &reference.linked)
        && embedded == linked
    {
        return Err(invalid(
            "SVG blip reuses one relationship ID for embed and link",
        ));
    }
    Ok(())
}

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

fn copy_source(source: &[u8]) -> Result<Arc<[u8]>> {
    let mut owned = Vec::new();
    owned
        .try_reserve_exact(source.len())
        .map_err(|error| allocation("SVG blip source", error))?;
    owned.extend_from_slice(source);
    Ok(Arc::from(owned.into_boxed_slice()))
}

fn allocation(resource: &'static str, source: std::collections::TryReserveError) -> Error {
    Error::Invalid(format!("SVG {resource} allocation failed: {source}"))
}

fn limit(resource: &'static str, limit: usize) -> Error {
    Error::Limit { resource, limit }
}

fn xml_error(error: impl fmt::Display) -> Error {
    Error::Xml(error.to_string())
}

fn value_error(error: impl fmt::Display) -> Error {
    Error::Invalid(error.to_string())
}
