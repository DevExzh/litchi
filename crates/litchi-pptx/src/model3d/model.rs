//! Semantic PPTX values for a 3D-model graphic frame.

use std::sync::Arc;

use litchi_drawingml::model3d as drawing;

use crate::{Error, Result};

/// Immutable, bounded bytes for an embedded glTF model or raster preview.
#[derive(Debug, Clone, PartialEq, Eq)]
#[must_use]
pub struct Data(Arc<Vec<u8>>);

impl Data {
    /// Adopt model or preview bytes after applying the owner-independent bound.
    pub fn new(bytes: Vec<u8>) -> Result<Self> {
        Self::from_shared(Arc::new(bytes))
    }

    pub(crate) fn from_shared(bytes: Arc<Vec<u8>>) -> Result<Self> {
        if bytes.len() > super::MAX_MODEL_BYTES {
            return Err(limit("model3d payload bytes", super::MAX_MODEL_BYTES));
        }
        Ok(Self(bytes))
    }

    /// Borrow the inert payload without parsing or executing it.
    #[inline]
    #[must_use]
    pub fn as_slice(&self) -> &[u8] {
        self.0.as_slice()
    }

    /// Return the payload length in bytes.
    #[inline]
    #[must_use]
    pub fn len(&self) -> usize {
        self.0.len()
    }

    /// Return whether the payload is empty.
    #[inline]
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.0.is_empty()
    }

    /// Recover the vector without copying when this is the sole owner.
    pub fn try_into_vec(self) -> std::result::Result<Vec<u8>, Self> {
        Arc::try_unwrap(self.0).map_err(Self)
    }

    pub(crate) fn shared(&self) -> Arc<Vec<u8>> {
        Arc::clone(&self.0)
    }
}

impl AsRef<[u8]> for Data {
    fn as_ref(&self) -> &[u8] {
        self.as_slice()
    }
}

/// An inert external model URL.  Litchi never follows or fetches this value.
#[derive(Debug, Clone, PartialEq, Eq)]
#[must_use]
pub struct Link(Box<str>);

impl Link {
    /// Construct a bounded external target URI.
    pub fn new(value: impl AsRef<str>) -> Result<Self> {
        let value = value.as_ref();
        if value.is_empty()
            || value.len() > super::MAX_LINK_BYTES
            || value.bytes().any(|byte| byte.is_ascii_control())
        {
            return Err(invalid("external model3d target is empty or invalid"));
        }
        Ok(Self(value.into()))
    }

    /// Borrow the exact external target text.
    #[inline]
    #[must_use]
    pub fn as_str(&self) -> &str {
        &self.0
    }
}

/// The model asset references carried by `r:embed` and `r:link`.
///
/// Both fields are retained because the XML schema permits both attributes;
/// ordinary callers can use the concise `embedded` or `linked` constructors.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
#[must_use]
pub struct Asset {
    embedded: Option<Data>,
    linked: Option<Link>,
}

impl Asset {
    /// Construct an asset with one embedded glTF payload.
    #[must_use]
    pub fn embedded(data: Data) -> Self {
        Self {
            embedded: Some(data),
            linked: None,
        }
    }

    /// Construct an asset with one external model target.
    pub fn linked(target: impl AsRef<str>) -> Result<Self> {
        Ok(Self {
            embedded: None,
            linked: Some(Link::new(target)?),
        })
    }

    /// Construct an asset retaining both schema relationship attributes.
    pub fn both(data: Data, target: impl AsRef<str>) -> Result<Self> {
        Ok(Self {
            embedded: Some(data),
            linked: Some(Link::new(target)?),
        })
    }

    /// Borrow the embedded payload, if present.
    #[inline]
    #[must_use]
    pub fn embedded_data(&self) -> Option<&Data> {
        self.embedded.as_ref()
    }

    /// Borrow the external target, if present.
    #[inline]
    #[must_use]
    pub fn linked_target(&self) -> Option<&Link> {
        self.linked.as_ref()
    }

    /// Add or replace the embedded payload.
    pub fn set_embedded(&mut self, data: Option<Data>) {
        self.embedded = data;
    }

    /// Add or replace the external target.
    pub fn set_linked(&mut self, target: Option<Link>) {
        self.linked = target;
    }
}

/// A bounded raster preview carried by the model's `raster` child.
#[derive(Debug, Clone, PartialEq, Eq)]
#[must_use]
pub struct Preview {
    embedded: Option<Data>,
    linked: Option<Link>,
    content_type: Option<Box<str>>,
}

impl Preview {
    /// Construct an inert preview with an image content type.
    pub fn new(data: Data, content_type: impl AsRef<str>) -> Result<Self> {
        Self::embedded(data, content_type)
    }

    /// Construct an embedded raster preview.
    pub fn embedded(data: Data, content_type: impl AsRef<str>) -> Result<Self> {
        let content_type = checked_content_type(content_type.as_ref())?;
        Ok(Self {
            embedded: Some(data),
            linked: None,
            content_type: Some(content_type),
        })
    }

    /// Construct a linked raster preview without attempting to infer a MIME
    /// type from an external target.
    pub fn linked(target: impl AsRef<str>) -> Result<Self> {
        Ok(Self {
            embedded: None,
            linked: Some(Link::new(target)?),
            content_type: None,
        })
    }

    /// Construct a preview retaining both schema relationship attributes.
    pub fn both(
        data: Data,
        content_type: impl AsRef<str>,
        target: impl AsRef<str>,
    ) -> Result<Self> {
        let content_type = checked_content_type(content_type.as_ref())?;
        let content_type = content_type.as_ref();
        Ok(Self {
            embedded: Some(data),
            linked: Some(Link::new(target)?),
            content_type: Some(content_type.into()),
        })
    }

    /// Borrow the inert embedded preview bytes, if present.
    #[inline]
    #[must_use]
    pub fn data(&self) -> Option<&Data> {
        self.embedded.as_ref()
    }

    /// Borrow the external preview target, if present.
    #[inline]
    #[must_use]
    pub fn linked_target(&self) -> Option<&Link> {
        self.linked.as_ref()
    }

    /// Borrow the preview's OPC content type, when the package supplies one.
    #[inline]
    #[must_use]
    pub fn content_type(&self) -> Option<&str> {
        self.content_type.as_deref()
    }

    /// Add or replace the embedded preview bytes.
    pub fn set_embedded(&mut self, data: Option<Data>) {
        self.embedded = data;
    }

    /// Add or replace the external preview target.
    pub fn set_linked(&mut self, target: Option<Link>) {
        self.linked = target;
    }
}

/// The semantic shape anchor of one model3d instance.
#[derive(Debug, Clone, PartialEq, Eq)]
#[must_use]
pub struct Shape {
    index: usize,
    name: Option<Box<str>>,
}

impl Shape {
    /// Return the checked depth-first scene position.
    #[inline]
    #[must_use]
    pub fn index(&self) -> usize {
        self.index
    }

    /// Return the producer-visible shape name, when present.
    #[inline]
    #[must_use]
    pub fn name(&self) -> Option<&str> {
        self.name.as_deref()
    }

    pub(crate) fn from_location(index: usize, name: Option<Box<str>>) -> Self {
        Self { index, name }
    }
}

/// A retained future or unmodeled child of `am3d:model3d`.
#[derive(Debug, Clone, Copy)]
#[must_use]
pub struct Unknown<'a> {
    value: &'a drawing::Inert,
}

impl<'a> Unknown<'a> {
    /// Return the child local name.
    #[inline]
    #[must_use]
    pub fn local_name(self) -> &'a str {
        self.value.local_name()
    }

    /// Return the child namespace URI.
    #[inline]
    #[must_use]
    pub fn namespace(self) -> &'a str {
        self.value.namespace()
    }

    /// Borrow the complete bounded child element.
    #[inline]
    #[must_use]
    pub fn as_bytes(self) -> &'a [u8] {
        self.value.as_bytes()
    }
}

/// The shared, host-neutral model3d scene wrapped without exposing its
/// relationship identifiers.
#[derive(Debug, Clone, PartialEq, Eq)]
#[must_use]
pub struct Scene {
    pub(crate) wire: drawing::Metadata,
}

impl Scene {
    /// Number of ordered children in the model3d sequence.
    #[inline]
    #[must_use]
    pub fn child_count(&self) -> usize {
        self.wire.children.len()
    }

    /// Whether a typed raster child is present.
    #[inline]
    #[must_use]
    pub fn has_raster(&self) -> bool {
        self.wire.raster().is_some()
    }

    /// Iterate top-level future or unmodeled children without interpreting
    /// their content.
    pub fn unknown_children(&self) -> impl Iterator<Item = Unknown<'_>> {
        self.wire.children.iter().filter_map(|child| match child {
            drawing::Child::Opaque(value) if !is_known_scene_child(value) => {
                Some(Unknown { value })
            },
            drawing::Child::Raster(_) => None,
            drawing::Child::Opaque(_) => None,
            _ => None,
        })
    }

    pub(crate) fn from_wire(wire: drawing::Metadata) -> Self {
        Self { wire }
    }
}

/// One contextual PPTX 3D-model instance.
#[derive(Debug, Clone)]
#[must_use]
pub struct Model {
    pub(crate) scene: Scene,
    pub(crate) asset: Asset,
    pub(crate) preview: Option<Preview>,
    pub(crate) shape: Shape,
    pub(crate) base_xml: Arc<Vec<u8>>,
    pub(crate) origin: Origin,
}

impl Model {
    /// Borrow the shared semantic scene.
    #[inline]
    #[must_use]
    pub fn scene(&self) -> &Scene {
        &self.scene
    }

    /// Borrow the embedded or linked model asset metadata.
    #[inline]
    #[must_use]
    pub fn asset(&self) -> &Asset {
        &self.asset
    }

    /// Borrow the optional raster preview.
    #[inline]
    #[must_use]
    pub fn preview(&self) -> Option<&Preview> {
        self.preview.as_ref()
    }

    /// Borrow the semantic shape anchor.
    #[inline]
    #[must_use]
    pub fn shape(&self) -> &Shape {
        &self.shape
    }

    /// Replace the model asset while retaining the scene and producer
    /// snapshot. The package layer allocates relationships and parts.
    pub fn set_asset(&mut self, asset: Asset) {
        self.asset = asset;
    }

    /// Replace the inert raster preview.
    pub fn set_preview(&mut self, preview: Option<Preview>) {
        self.preview = preview;
    }

    pub(crate) fn semantic_eq(&self, other: &Self) -> bool {
        self.scene == other.scene
            && self.asset == other.asset
            && self.preview == other.preview
            && self.shape == other.shape
    }
}

/// A package relationship retained below the ordinary facade for snapshot
/// publication and orphan-safe cleanup.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct Relation {
    pub(crate) id: String,
    pub(crate) target: String,
    pub(crate) external: bool,
}

#[derive(Debug, Clone, Default)]
pub(crate) struct Origin {
    pub(crate) asset: Asset,
    pub(crate) preview: Option<Preview>,
    pub(crate) model_embedded: Option<Relation>,
    pub(crate) model_linked: Option<Relation>,
    pub(crate) preview_embedded: Option<Relation>,
    pub(crate) preview_linked: Option<Relation>,
}

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(format!("PPTX model3d {}", message.into()))
}

fn checked_content_type(value: &str) -> Result<Box<str>> {
    if value.is_empty() || value.len() > 256 || value.bytes().any(|byte| byte.is_ascii_control()) {
        return Err(invalid("model3d preview content type is invalid"));
    }
    Ok(value.into())
}

fn is_known_scene_child(value: &drawing::Inert) -> bool {
    value.namespace() == drawing::NAMESPACE
        && matches!(
            value.local_name(),
            "spPr"
                | "camera"
                | "trans"
                | "attrSrcUrl"
                | "extLst"
                | "objViewport"
                | "winViewport"
                | "ambientLight"
                | "ptLight"
                | "spotLight"
                | "dirLight"
                | "unkLight"
        )
}

fn limit(resource: &'static str, limit: usize) -> Error {
    Error::Limit { resource, limit }
}

/// Convert the shared model3d error into the host error without leaking its
/// package-neutral type through the ordinary PPTX facade.
pub(crate) fn drawing_error(error: litchi_drawingml::Error) -> Error {
    Error::Drawing(error)
}

/// Construct a checked shared relationship identifier for package publication.
pub(crate) fn relationship_id(value: &str) -> Result<drawing::Id> {
    drawing::Id::new(value).map_err(|error| invalid(error.to_string()))
}
