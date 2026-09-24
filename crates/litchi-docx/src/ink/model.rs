use std::sync::Arc;

use litchi_core::Position;
use litchi_drawingml::ink as shared;

use crate::package::story::{StoryKind, StoryLimits};
use crate::{Error, Result};

/// Resource budgets for one complete annotation inventory.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct Limits {
    /// Existing package/story ownership and source-byte bounds.
    pub stories: StoryLimits,
    /// Maximum active content-part anchors, including generic XML content.
    pub max_annotations: usize,
    /// Maximum bytes in one candidate content payload.
    pub max_payload_bytes: usize,
    /// Maximum distinct candidate content bytes inspected in total.
    pub max_total_payload_bytes: usize,
    /// Maximum XML elements in each story scanner.
    pub max_xml_nodes: usize,
    /// Maximum XML element depth in a story scanner.
    pub max_xml_depth: usize,
    /// Maximum relationships across all package owners.
    pub max_relationships: usize,
}

impl Default for Limits {
    fn default() -> Self {
        Self {
            stories: StoryLimits::default(),
            max_annotations: 16_384,
            max_payload_bytes: shared::MAX_SOURCE_BYTES,
            max_total_payload_bytes: 32 * 1024 * 1024,
            max_xml_nodes: 1_000_000,
            max_xml_depth: 128,
            max_relationships: 32_768,
        }
    }
}

impl Limits {
    pub(crate) fn validate(self) -> Result<Self> {
        self.stories.validate()?;
        for (resource, value, maximum) in [
            ("annotations", self.max_annotations, 65_536),
            (
                "payload bytes",
                self.max_payload_bytes,
                shared::MAX_SOURCE_BYTES,
            ),
            (
                "total payload bytes",
                self.max_total_payload_bytes,
                128 * 1024 * 1024,
            ),
            ("XML nodes", self.max_xml_nodes, 1_000_000),
            ("XML depth", self.max_xml_depth, 128),
            ("relationships", self.max_relationships, 1_000_000),
        ] {
            if value == 0 || value > maximum {
                return Err(Error::Invalid(format!(
                    "DOCX ink {resource} budget must be in 1..={maximum}"
                )));
            }
        }
        Ok(self)
    }
}

/// Snapshot-local position of an annotation's Word story.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct Location {
    pub(crate) kind: StoryKind,
    pub(crate) position: Position,
}

impl Location {
    /// Select a story by semantic role and its zero-based position in that role.
    ///
    /// The edit resolves this selector against its immutable package snapshot.
    #[must_use]
    pub const fn new(kind: StoryKind, position: Position) -> Self {
        Self { kind, position }
    }

    /// Semantic story role, such as header or footnotes.
    #[must_use]
    pub const fn kind(self) -> StoryKind {
        self.kind
    }

    /// Zero-based story position among stories of the same role.
    #[must_use]
    pub const fn position(self) -> Position {
        self.position
    }
}

/// Keep the complete source owner beside its bounded metadata. In particular,
/// a managed PartData reservation must outlive every projected annotation.
#[derive(Debug)]
pub(crate) enum Payload {
    Owned(shared::Document),
    Pinned {
        metadata: shared::Metadata,
        _source: litchi_opc::PartData,
    },
}

impl Payload {
    pub(crate) fn trace_count(&self) -> usize {
        match self {
            Self::Owned(document) => document.trace_count(),
            Self::Pinned { metadata, .. } => metadata.trace_count(),
        }
    }

    pub(crate) fn contexts(&self) -> &[shared::Context] {
        match self {
            Self::Owned(document) => document.contexts(),
            Self::Pinned { metadata, .. } => metadata.contexts(),
        }
    }

    pub(crate) fn brush_properties(&self) -> &[shared::BrushProperty] {
        match self {
            Self::Owned(document) => document.brush_properties(),
            Self::Pinned { metadata, .. } => metadata.brush_properties(),
        }
    }

    pub(crate) fn source(&self) -> &[u8] {
        match self {
            Self::Owned(document) => document.source(),
            Self::Pinned { _source, .. } => _source.as_bytes(),
        }
    }
}

/// One active InkML annotation, retaining its shared immutable payload.
#[derive(Clone, Debug)]
pub struct Annotation {
    pub(crate) location: Location,
    pub(crate) document: Arc<Payload>,
}

impl Annotation {
    /// Story containing this annotation.
    #[must_use]
    pub const fn location(&self) -> Location {
        self.location
    }

    /// Number of discovered traces; coordinates remain uninterpreted.
    #[must_use]
    pub fn trace_count(&self) -> usize {
        self.document.trace_count()
    }

    /// Borrow semantic context metadata in source order.
    pub fn contexts(&self) -> impl ExactSizeIterator<Item = Context<'_>> {
        self.document.contexts().iter().map(Context)
    }

    /// Borrow brush metadata in source order, including preserved future names.
    pub fn brush_properties(&self) -> impl ExactSizeIterator<Item = BrushProperty<'_>> {
        self.document.brush_properties().iter().map(BrushProperty)
    }
}

/// Borrowed context classification without source spans or native identifiers.
#[derive(Clone, Copy, Debug)]
pub struct Context<'a>(&'a shared::Context);

impl Context<'_> {
    #[must_use]
    pub fn kind(&self) -> &shared::ContextKind {
        self.0.kind()
    }
    #[must_use]
    pub fn semantic_type(&self) -> Option<&shared::SemanticType> {
        self.0.semantic_type()
    }
    #[must_use]
    pub fn alignment_level(&self) -> Option<i32> {
        self.0.alignment_level()
    }
    #[must_use]
    pub fn rotation_angle(&self) -> Option<i32> {
        self.0.rotation_angle()
    }
}

/// Borrowed brush metadata; values are not interpreted as rendering commands.
#[derive(Clone, Copy, Debug)]
pub struct BrushProperty<'a>(&'a shared::BrushProperty);

impl BrushProperty<'_> {
    #[must_use]
    pub fn name(&self) -> &shared::BrushPropertyName {
        self.0.name()
    }
    #[must_use]
    pub fn value(&self) -> &str {
        self.0.value()
    }
    #[must_use]
    pub fn units(&self) -> Option<&str> {
        self.0.units()
    }
    #[must_use]
    pub fn ink_effect(&self) -> Option<shared::InkEffect> {
        self.0.ink_effect()
    }
    /// Return the effective MS-ODRAWXML value, including profile defaults.
    #[must_use]
    pub fn effective(&self) -> Option<shared::EffectiveBrushProperty> {
        self.0.effective()
    }
}

/// Immutable package-wide annotation inventory; clones share all retained data.
#[derive(Clone, Debug)]
pub struct Snapshot {
    pub(crate) annotations: Arc<Vec<Annotation>>,
    pub(crate) distinct_payloads: usize,
}

impl Snapshot {
    /// Annotations in main-story-first order, then remaining story order.
    #[must_use]
    pub fn annotations(&self) -> &[Annotation] {
        &self.annotations
    }

    /// Select an annotation by its snapshot-local zero-based position.
    #[must_use]
    pub fn get(&self, position: Position) -> Option<&Annotation> {
        self.annotations.get(position.get())
    }

    /// Number of distinct InkML payloads represented by this inventory.
    #[must_use]
    pub const fn distinct_payload_count(&self) -> usize {
        self.distinct_payloads
    }
}
