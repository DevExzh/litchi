//! Semantic values for inert PresentationML content-part relationships.

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
    pub const fn slide_index(&self) -> usize {
        self.slide_index
    }

    /// Absolute part name of the owning slide.
    #[inline]
    pub fn slide_part_name(&self) -> &PackURI {
        &self.slide_part_name
    }

    /// Zero-based content-part position in the owning slide's active markup.
    #[inline]
    pub const fn index(&self) -> usize {
        self.index
    }

    /// The selected, bounded `p:contentPart` anchor.
    #[inline]
    pub fn anchor(&self) -> &Anchor {
        &self.anchor
    }

    /// The owning slide relationship resolved from the anchor's `r:id`.
    #[inline]
    pub fn relationship(&self) -> &Relationship {
        &self.relationship
    }

    /// The relationship identifier written by the anchor.
    #[inline]
    pub fn relationship_id(&self) -> &str {
        self.anchor.relationship_id()
    }

    /// The inert target of this content part.
    #[inline]
    pub fn target(&self) -> &Target {
        self.relationship.target()
    }

    /// Internal payload bytes, or `None` for an external target.
    #[inline]
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
    /// The `r:id` value resolved against the owning slide relationship graph.
    #[inline]
    pub fn relationship_id(&self) -> &str {
        &self.relationship_id
    }

    /// Exact selected anchor bytes after markup-compatibility branch choice.
    #[inline]
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
    /// Relationship identifier in the owning slide `.rels` part.
    #[inline]
    pub fn id(&self) -> &str {
        &self.id
    }

    /// Relationship type URI, retained without vocabulary interpretation.
    #[inline]
    pub fn relationship_type(&self) -> &str {
        &self.relationship_type
    }

    /// Original target reference from the owning slide relationship XML.
    #[inline]
    pub fn target_ref(&self) -> &str {
        &self.target_ref
    }

    /// OPC target mode; external targets are never contacted.
    #[inline]
    pub const fn target_mode(&self) -> TargetMode {
        self.target_mode
    }

    /// Inert internal payload or external URI target.
    #[inline]
    pub fn target(&self) -> &Target {
        &self.target
    }

    /// Borrow the payload when the relationship is internal.
    #[inline]
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
    /// Return the internal payload, when present.
    #[inline]
    pub fn internal(&self) -> Option<&Payload> {
        match self {
            Self::Internal(payload) => Some(payload),
            Self::External { .. } => None,
        }
    }

    /// Return the preserved external URI, when present.
    #[inline]
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
    /// Absolute OPC part name of this target.
    #[inline]
    pub fn part_name(&self) -> &PackURI {
        &self.part_name
    }

    /// Declared OPC content type, without format sniffing.
    #[inline]
    pub fn content_type(&self) -> &str {
        &self.content_type
    }

    /// Exact target payload bytes.
    #[inline]
    pub fn bytes(&self) -> &[u8] {
        &self.bytes
    }

    /// Relationship metadata owned by the payload part.
    #[inline]
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
    /// Relationship identifier in the payload part's `.rels` graph.
    #[inline]
    pub fn id(&self) -> &str {
        &self.id
    }

    /// Relationship type URI.
    #[inline]
    pub fn relationship_type(&self) -> &str {
        &self.relationship_type
    }

    /// Original relationship target reference.
    #[inline]
    pub fn target_ref(&self) -> &str {
        &self.target_ref
    }

    /// OPC target mode, retained without following the target.
    #[inline]
    pub const fn target_mode(&self) -> TargetMode {
        self.target_mode
    }
}
