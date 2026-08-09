//! Semantic values for inert `PresentationML` content-part relationships.

use litchi_opc::{PackURI, TargetMode};
use std::sync::Arc;

/// One content-part anchor and its owning slide relationship.
///
/// The payload is deliberately opaque. The model exposes bytes and OPC
/// metadata only; it never parses, follows, renders, or executes the
/// vocabulary stored in the related part.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ContentPart {
    pub(crate) slide_index: usize,
    pub(crate) slide_part_name: PackURI,
    pub(crate) index: usize,
    pub(crate) anchor: Anchor,
    pub(crate) relationship: Relationship,
}

impl ContentPart {
    /// Zero-based slide position containing this content-part anchor.
    #[inline]
    #[must_use]
    pub const fn slide_index(&self) -> usize {
        self.slide_index
    }

    /// Absolute part name of the owning slide.
    #[inline]
    #[must_use]
    pub fn slide_part_name(&self) -> &PackURI {
        &self.slide_part_name
    }

    /// Zero-based content-part position in the owning slide's active markup.
    #[inline]
    #[must_use]
    pub const fn index(&self) -> usize {
        self.index
    }

    /// The selected, bounded `p:contentPart` anchor.
    #[inline]
    #[must_use]
    pub fn anchor(&self) -> &Anchor {
        &self.anchor
    }

    /// The owning slide relationship resolved from the anchor's `r:id`.
    #[inline]
    #[must_use]
    pub fn relationship(&self) -> &Relationship {
        &self.relationship
    }

    /// The relationship identifier written by the anchor.
    #[inline]
    #[must_use]
    pub fn relationship_id(&self) -> &str {
        self.anchor.relationship_id()
    }

    /// The inert target of this content part.
    #[inline]
    #[must_use]
    pub fn target(&self) -> &Target {
        self.relationship.target()
    }

    /// Internal payload bytes, or `None` for an external target.
    #[inline]
    #[must_use]
    pub fn payload(&self) -> Option<&Payload> {
        self.relationship.payload()
    }
}

/// Lossless bounded representation of one active content-part anchor.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Anchor {
    pub(crate) relationship_id: String,
    pub(crate) xml: Vec<u8>,
}

impl Anchor {
    /// Create an opaque content-part anchor from its exact XML bytes.
    ///
    /// The anchor is checked when it is inserted into a snapshot transaction;
    /// construction itself stays allocation-only so callers can assemble a
    /// detached graph without an OPC package.
    #[inline]
    pub fn new(relationship_id: impl Into<String>, xml: impl Into<Vec<u8>>) -> Self {
        Self {
            relationship_id: relationship_id.into(),
            xml: xml.into(),
        }
    }

    /// The `r:id` value resolved against the owning slide relationship graph.
    #[inline]
    #[must_use]
    pub fn relationship_id(&self) -> &str {
        &self.relationship_id
    }

    /// Exact selected anchor bytes after markup-compatibility branch choice.
    #[inline]
    #[must_use]
    pub fn xml(&self) -> &[u8] {
        &self.xml
    }
}

/// Relationship metadata owned by the content-part anchor.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Relationship {
    pub(crate) id: String,
    pub(crate) relationship_type: String,
    pub(crate) target_ref: String,
    pub(crate) target_mode: TargetMode,
    pub(crate) target: Target,
}

impl Relationship {
    /// Create relationship metadata for a detached content-part graph.
    #[inline]
    pub fn new(
        id: impl Into<String>,
        relationship_type: impl Into<String>,
        target_ref: impl Into<String>,
        target_mode: TargetMode,
        target: Target,
    ) -> Self {
        Self {
            id: id.into(),
            relationship_type: relationship_type.into(),
            target_ref: target_ref.into(),
            target_mode,
            target,
        }
    }

    /// Relationship identifier in the owning slide `.rels` part.
    #[inline]
    #[must_use]
    pub fn id(&self) -> &str {
        &self.id
    }

    /// Relationship type URI, retained without vocabulary interpretation.
    #[inline]
    #[must_use]
    pub fn relationship_type(&self) -> &str {
        &self.relationship_type
    }

    /// Original target reference from the owning slide relationship XML.
    #[inline]
    #[must_use]
    pub fn target_ref(&self) -> &str {
        &self.target_ref
    }

    /// OPC target mode; external targets are never contacted.
    #[inline]
    #[must_use]
    pub const fn target_mode(&self) -> TargetMode {
        self.target_mode
    }

    /// Inert internal payload or external URI target.
    #[inline]
    #[must_use]
    pub fn target(&self) -> &Target {
        &self.target
    }

    /// Borrow the payload when the relationship is internal.
    #[inline]
    #[must_use]
    pub fn payload(&self) -> Option<&Payload> {
        match &self.target {
            Target::Internal(payload) => Some(payload),
            Target::External { .. } => None,
        }
    }
}

/// A relationship target retained without following or executing it.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum Target {
    /// Package-owned bytes and the payload part's relationship metadata.
    Internal(Payload),
    /// External URI and mode, preserved inertly.
    External {
        /// Original external target reference.
        target_ref: String,
    },
}

impl Target {
    /// Construct an inert internal target.
    #[inline]
    #[must_use]
    pub fn internal(payload: Payload) -> Self {
        Self::Internal(payload)
    }

    /// Construct an inert external target without contacting its URI.
    #[inline]
    pub fn external(target_ref: impl Into<String>) -> Self {
        Self::External {
            target_ref: target_ref.into(),
        }
    }

    /// Return the internal payload, when present.
    #[inline]
    #[must_use]
    pub fn payload(&self) -> Option<&Payload> {
        match self {
            Self::Internal(payload) => Some(payload),
            Self::External { .. } => None,
        }
    }

    /// Return the preserved external URI, when present.
    #[inline]
    #[must_use]
    pub fn external_ref(&self) -> Option<&str> {
        match self {
            Self::Internal(_) => None,
            Self::External { target_ref } => Some(target_ref),
        }
    }
}

/// Opaque bytes and metadata for one internal content-part target.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Payload {
    pub(crate) part_name: PackURI,
    pub(crate) content_type: Arc<str>,
    pub(crate) bytes: Arc<[u8]>,
    pub(crate) relationships: Arc<[RelationshipMetadata]>,
}

impl Payload {
    /// Create opaque payload bytes and an empty outbound relationship set.
    #[inline]
    pub fn new(
        part_name: PackURI,
        content_type: impl Into<Arc<str>>,
        bytes: impl Into<Vec<u8>>,
    ) -> Self {
        Self {
            part_name,
            content_type: content_type.into(),
            bytes: Arc::<[u8]>::from(bytes.into()),
            relationships: Arc::from([]),
        }
    }

    /// Attach opaque relationship metadata to a payload.
    #[must_use]
    #[inline]
    pub fn with_relationships(
        mut self,
        relationships: impl IntoIterator<Item = RelationshipMetadata>,
    ) -> Self {
        let mut values = relationships.into_iter().collect::<Vec<_>>();
        values.sort_unstable_by(|left, right| left.id.cmp(&right.id));
        self.relationships = Arc::<[RelationshipMetadata]>::from(values);
        self
    }

    /// Absolute OPC part name of this target.
    #[inline]
    #[must_use]
    pub fn part_name(&self) -> &PackURI {
        &self.part_name
    }

    /// Declared OPC content type, without format sniffing.
    #[inline]
    #[must_use]
    pub fn content_type(&self) -> &str {
        &self.content_type
    }

    /// Exact target payload bytes.
    #[inline]
    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        &self.bytes
    }

    /// Relationship metadata owned by the payload part.
    #[inline]
    #[must_use]
    pub fn relationships(&self) -> &[RelationshipMetadata] {
        &self.relationships
    }
}

/// Relationship metadata on an opaque payload part.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RelationshipMetadata {
    pub(crate) id: String,
    pub(crate) relationship_type: String,
    pub(crate) target_ref: String,
    pub(crate) target_mode: TargetMode,
}

impl RelationshipMetadata {
    /// Create outbound relationship metadata without resolving its target.
    #[inline]
    pub fn new(
        id: impl Into<String>,
        relationship_type: impl Into<String>,
        target_ref: impl Into<String>,
        target_mode: TargetMode,
    ) -> Self {
        Self {
            id: id.into(),
            relationship_type: relationship_type.into(),
            target_ref: target_ref.into(),
            target_mode,
        }
    }

    /// Relationship identifier in the payload part's `.rels` graph.
    #[inline]
    #[must_use]
    pub fn id(&self) -> &str {
        &self.id
    }

    /// Relationship type URI.
    #[inline]
    #[must_use]
    pub fn relationship_type(&self) -> &str {
        &self.relationship_type
    }

    /// Original relationship target reference.
    #[inline]
    #[must_use]
    pub fn target_ref(&self) -> &str {
        &self.target_ref
    }

    /// OPC target mode, retained without following the target.
    #[inline]
    #[must_use]
    pub const fn target_mode(&self) -> TargetMode {
        self.target_mode
    }
}
