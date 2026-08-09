//! Semantic values for the relationship-bearing portion of `model3d`.

use std::fmt;

use litchi_ooxml_common::xml::is_ncname;
use thiserror::Error;

use super::{MAX_NAMESPACE_TEXT_BYTES, MAX_RELATIONSHIP_ID_BYTES, MAX_RENDERER_TEXT_BYTES};

/// Failure to construct a bounded model3d scalar.
#[derive(Debug, Clone, PartialEq, Eq, Error)]
#[non_exhaustive]
pub enum ValueError {
    /// A relationship identifier is empty, too long, or not an XML `NCName`.
    #[error("invalid model3d relationship ID '{value}'")]
    RelationshipId { value: String },
    /// A text or namespace value exceeds the model3d resource bound.
    #[error("model3d {field} exceeds the limit of {limit} bytes")]
    TooLong {
        /// Semantic field being bounded.
        field: &'static str,
        /// Maximum accepted byte length.
        limit: usize,
    },
    /// A namespace prefix is not an XML `NCName`.
    #[error("invalid model3d namespace prefix '{value}'")]
    NamespacePrefix { value: String },
    /// A namespace URI cannot be empty.
    #[error("model3d namespace URI is empty")]
    EmptyNamespace,
}

/// A bounded `ST_RelationshipId` value.
///
/// The value is retained exactly after XML attribute decoding.  The `NCName`
/// check follows the shared OOXML lexical helper; the finite bound prevents a
/// hostile attribute from becoming an unbounded allocation.
#[derive(Debug, Clone, PartialEq, Eq, Hash, PartialOrd, Ord)]
#[must_use]
pub struct Id(Box<str>);

impl Id {
    /// Construct a checked relationship identifier.
    /// # Errors
    ///
    /// Returns an error when input violates DrawingML constraints, exceeds a configured
    /// bound, or an underlying XML, MCE, I/O, or formatting operation fails.
    pub fn new(value: impl AsRef<str>) -> Result<Self, ValueError> {
        let value = value.as_ref();
        if value.is_empty() || value.len() > MAX_RELATIONSHIP_ID_BYTES || !is_ncname(value) {
            return Err(ValueError::RelationshipId {
                value: value.to_owned(),
            });
        }
        Ok(Self(value.into()))
    }

    /// Borrow the exact lexical identifier.
    #[inline]
    #[must_use]
    pub fn as_str(&self) -> &str {
        &self.0
    }
}

impl AsRef<str> for Id {
    #[inline]
    fn as_ref(&self) -> &str {
        self.as_str()
    }
}

impl fmt::Display for Id {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(self.as_str())
    }
}

impl TryFrom<&str> for Id {
    type Error = ValueError;

    fn try_from(value: &str) -> Result<Self, Self::Error> {
        Self::new(value)
    }
}

impl TryFrom<String> for Id {
    type Error = ValueError;

    fn try_from(value: String) -> Result<Self, Self::Error> {
        Self::new(value)
    }
}

impl From<Id> for String {
    fn from(value: Id) -> Self {
        value.0.into()
    }
}

/// The `AG_Blob` relationship metadata shared by the model and raster blip.
///
/// The schema permits the two independent attributes.  They must not point to
/// the same relationship occurrence; package validation additionally checks
/// that `embedded` resolves internally and `linked` resolves externally.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
#[must_use]
pub struct Reference {
    /// Relationship ID for a package-local binary payload (`r:embed`).
    pub embedded: Option<Id>,
    /// Relationship ID for a non-package binary payload (`r:link`).
    pub linked: Option<Id>,
}

impl Reference {
    /// An empty relationship reference.
    #[inline]
    #[must_use]
    pub const fn none() -> Self {
        Self {
            embedded: None,
            linked: None,
        }
    }

    /// Create an embedded reference from an ordinary lexical ID.
    /// # Errors
    ///
    /// Returns an error when input violates DrawingML constraints, exceeds a configured
    /// bound, or an underlying XML, MCE, I/O, or formatting operation fails.
    pub fn embedded(value: impl AsRef<str>) -> Result<Self, ValueError> {
        Ok(Self {
            embedded: Some(Id::new(value)?),
            linked: None,
        })
    }

    /// Create a linked reference from an ordinary lexical ID.
    /// # Errors
    ///
    /// Returns an error when input violates DrawingML constraints, exceeds a configured
    /// bound, or an underlying XML, MCE, I/O, or formatting operation fails.
    pub fn linked(value: impl AsRef<str>) -> Result<Self, ValueError> {
        Ok(Self {
            embedded: None,
            linked: Some(Id::new(value)?),
        })
    }

    /// Return whether no relationship attribute is present.
    #[inline]
    #[must_use]
    pub const fn is_empty(&self) -> bool {
        self.embedded.is_none() && self.linked.is_none()
    }
}

/// A namespace declaration retained because an opaque child may rely on it.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
#[must_use]
pub struct Namespace {
    prefix: Option<Box<str>>,
    uri: Box<str>,
}

impl Namespace {
    /// Construct a namespace declaration. `None` denotes the default prefix.
    /// # Errors
    ///
    /// Returns an error when input violates DrawingML constraints, exceeds a configured
    /// bound, or an underlying XML, MCE, I/O, or formatting operation fails.
    pub fn new(prefix: Option<&str>, uri: impl AsRef<str>) -> Result<Self, ValueError> {
        if let Some(prefix) = prefix
            && !prefix.is_empty()
            && (prefix.len() > MAX_NAMESPACE_TEXT_BYTES || !is_ncname(prefix))
        {
            return Err(ValueError::NamespacePrefix {
                value: prefix.to_owned(),
            });
        }
        let uri = uri.as_ref();
        if uri.is_empty() {
            return Err(ValueError::EmptyNamespace);
        }
        if uri.len() > MAX_NAMESPACE_TEXT_BYTES {
            return Err(ValueError::TooLong {
                field: "namespace URI",
                limit: MAX_NAMESPACE_TEXT_BYTES,
            });
        }
        Ok(Self {
            prefix: prefix.filter(|value| !value.is_empty()).map(Into::into),
            uri: uri.into(),
        })
    }

    /// Prefix, or `None` for the default namespace.
    #[inline]
    #[must_use]
    pub fn prefix(&self) -> Option<&str> {
        self.prefix.as_deref()
    }

    /// Namespace URI.
    #[inline]
    #[must_use]
    pub fn uri(&self) -> &str {
        &self.uri
    }
}

/// An exact, uninterpreted child element retained by the model3d codec.
///
/// The bytes are one complete XML element, not a document.  Namespace
/// declarations inherited from the model3d root are retained separately in
/// [`Metadata::namespaces`], so the fragment can be replayed without trying
/// to understand a future scene, extension, or rendering payload.
#[derive(Debug, Clone, PartialEq, Eq)]
#[must_use]
pub struct Inert {
    pub(super) xml: Box<[u8]>,
    pub(super) local_name: Box<str>,
    pub(super) namespace: Box<str>,
}

impl Inert {
    /// Borrow the exact retained element bytes.
    #[inline]
    #[must_use]
    pub fn as_bytes(&self) -> &[u8] {
        &self.xml
    }

    /// Local name observed by the namespace-aware reader.
    #[inline]
    #[must_use]
    pub fn local_name(&self) -> &str {
        &self.local_name
    }

    /// Expanded namespace URI observed by the reader.
    #[inline]
    #[must_use]
    pub fn namespace(&self) -> &str {
        &self.namespace
    }

    #[allow(
        clippy::needless_pass_by_value,
        reason = "the wire decoder transfers ownership of renderer strings into the bounded model"
    )]
    pub(super) fn from_wire(
        xml: Vec<u8>,
        local_name: impl Into<Box<str>>,
        namespace: impl Into<Box<str>>,
    ) -> Self {
        Self {
            xml: xml.into_boxed_slice(),
            local_name: local_name.into(),
            namespace: namespace.into(),
        }
    }
}

/// A child in the `CT_Model3D` sequence.
#[derive(Debug, Clone, PartialEq, Eq)]
#[non_exhaustive]
pub enum Child {
    /// A typed raster preview, including its renderer and blip relationship.
    Raster(Raster),
    /// A complete future, host-specific, or not-yet-modeled child element.
    Opaque(Inert),
}

/// The raster preview metadata defined by `CT_Model3DRaster`.
#[derive(Debug, Clone, PartialEq, Eq)]
#[must_use]
pub struct Raster {
    /// Renderer name (`rName`), retained as bounded text.
    pub renderer_name: Box<str>,
    /// Renderer version (`rVer`), retained as bounded text.
    pub renderer_version: Box<str>,
    /// Ordered raster children; the schema's known child is [`RasterChild::Blip`].
    pub children: Vec<RasterChild>,
    /// Namespace declarations needed by retained raster children.
    pub namespaces: Vec<Namespace>,
}

impl Raster {
    /// Construct a raster metadata value without a preview relationship.
    /// # Errors
    ///
    /// Returns an error when input violates DrawingML constraints, exceeds a configured
    /// bound, or an underlying XML, MCE, I/O, or formatting operation fails.
    pub fn new(
        renderer_name: impl AsRef<str>,
        renderer_version: impl AsRef<str>,
    ) -> Result<Self, ValueError> {
        let renderer_name = bounded_text(renderer_name.as_ref(), "renderer name")?;
        let renderer_version = bounded_text(renderer_version.as_ref(), "renderer version")?;
        Ok(Self {
            renderer_name,
            renderer_version,
            children: Vec::new(),
            namespaces: Vec::new(),
        })
    }

    /// Return the typed blip, if the raster carries one.
    #[must_use]
    pub fn blip(&self) -> Option<&Blip> {
        self.children.iter().find_map(|child| match child {
            RasterChild::Blip(blip) => Some(blip),
            RasterChild::Opaque(_) => None,
        })
    }

    /// Replace the typed blip while retaining its position when possible.
    pub fn set_blip(&mut self, blip: Blip) {
        if let Some(child) = self
            .children
            .iter_mut()
            .find(|child| matches!(child, RasterChild::Blip(_)))
        {
            *child = RasterChild::Blip(blip);
        } else {
            self.children.insert(0, RasterChild::Blip(blip));
        }
    }

    pub(super) fn from_wire(
        renderer_name: String,
        renderer_version: String,
        children: Vec<RasterChild>,
        namespaces: Vec<Namespace>,
    ) -> Result<Self, ValueError> {
        Ok(Self {
            renderer_name: bounded_owned_text(renderer_name, "renderer name")?,
            renderer_version: bounded_owned_text(renderer_version, "renderer version")?,
            children,
            namespaces,
        })
    }
}

/// A raster `m3d:blip` element (typed as `a:CT_Blip`) and its extensions.
#[derive(Debug, Clone, PartialEq, Eq)]
#[must_use]
pub struct Blip {
    /// Embedded or linked preview-image relationship metadata.
    pub reference: Reference,
    /// Unmodeled `a:blip` children, retained in order.
    pub children: Vec<Inert>,
    /// Namespace declarations local to the blip element.
    pub namespaces: Vec<Namespace>,
}

impl Blip {
    /// Construct an empty blip reference.
    #[inline]
    #[must_use]
    pub const fn new() -> Self {
        Self {
            reference: Reference::none(),
            children: Vec::new(),
            namespaces: Vec::new(),
        }
    }

    pub(super) fn from_wire(
        reference: Reference,
        children: Vec<Inert>,
        namespaces: Vec<Namespace>,
    ) -> Self {
        Self {
            reference,
            children,
            namespaces,
        }
    }
}

impl Default for Blip {
    fn default() -> Self {
        Self::new()
    }
}

/// Ordered child in a raster preview.
#[derive(Debug, Clone, PartialEq, Eq)]
#[non_exhaustive]
pub enum RasterChild {
    /// The schema's optional preview image relationship.
    Blip(Blip),
    /// A future or extension child retained without interpretation.
    Opaque(Inert),
}

/// Relationship metadata and inert scene payload for one `model3d` element.
#[derive(Debug, Clone, PartialEq, Eq)]
#[must_use]
pub struct Metadata {
    /// The model binary's `AG_Blob` relationship metadata.
    pub reference: Reference,
    /// Ordered scene children; mandatory scene structures remain inert here.
    pub children: Vec<Child>,
    /// Root namespace declarations required by retained child fragments.
    pub namespaces: Vec<Namespace>,
}

impl Metadata {
    /// Start a metadata value with the supplied model relationship reference.
    #[inline]
    #[must_use]
    pub const fn new(reference: Reference) -> Self {
        Self {
            reference,
            children: Vec::new(),
            namespaces: Vec::new(),
        }
    }

    /// Return the typed raster preview, if present.
    #[must_use]
    pub fn raster(&self) -> Option<&Raster> {
        self.children.iter().find_map(|child| match child {
            Child::Raster(raster) => Some(raster),
            Child::Opaque(_) => None,
        })
    }

    /// Replace the typed raster preview while retaining its sequence position.
    pub fn set_raster(&mut self, raster: Raster) {
        if let Some(child) = self
            .children
            .iter_mut()
            .find(|child| matches!(child, Child::Raster(_)))
        {
            *child = Child::Raster(raster);
        } else {
            self.children.push(Child::Raster(raster));
        }
    }
}

fn bounded_text(value: &str, field: &'static str) -> Result<Box<str>, ValueError> {
    if value.len() > MAX_RENDERER_TEXT_BYTES {
        return Err(ValueError::TooLong {
            field,
            limit: MAX_RENDERER_TEXT_BYTES,
        });
    }
    Ok(value.into())
}

fn bounded_owned_text(value: String, field: &'static str) -> Result<Box<str>, ValueError> {
    if value.len() > MAX_RENDERER_TEXT_BYTES {
        return Err(ValueError::TooLong {
            field,
            limit: MAX_RENDERER_TEXT_BYTES,
        });
    }
    Ok(value.into_boxed_str())
}
