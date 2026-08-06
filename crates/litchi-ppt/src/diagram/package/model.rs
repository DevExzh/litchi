//! Immutable slide-boundary values for contextual diagram publication.

use std::ops::Range;
use std::sync::Arc;

use crate::package::Result;

use super::super::model::{Build, EditLimits, Id, Inventory};
use super::super::transaction::Snapshot as DiagramSnapshot;

/// Resource ceilings for one complete owning `SlideContainer` payload.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SlideLimits {
    /// Maximum complete serialized `SlideContainer` size.
    pub max_slide_bytes: usize,
    /// Limits applied to its contextual diagram owner.
    pub diagram: EditLimits,
}

impl Default for SlideLimits {
    fn default() -> Self {
        Self {
            max_slide_bytes: 64 * 1024 * 1024,
            diagram: EditLimits::default(),
        }
    }
}

/// A deterministic fingerprint of one complete owning slide payload.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct SlideRevision(u64);

impl SlideRevision {
    pub(super) fn from_bytes(bytes: &[u8]) -> Self {
        let mut value = 0xcbf2_9ce4_8422_2325u64;
        value ^= bytes.len() as u64;
        value = value.wrapping_mul(0x1000_0000_01b3);
        for byte in bytes {
            value ^= u64::from(*byte);
            value = value.wrapping_mul(0x1000_0000_01b3);
        }
        Self(value)
    }

    /// Returns the compact source fingerprint.
    pub const fn value(self) -> u64 {
        self.0
    }
}

/// An immutable, source-preserving diagram view of one complete slide.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SlideSnapshot {
    pub(super) bytes: Arc<[u8]>,
    pub(super) build_range: Range<usize>,
    pub(super) drawing_range: Range<usize>,
    pub(super) diagram: DiagramSnapshot,
    pub(super) revision: SlideRevision,
    pub(super) limits: SlideLimits,
}

impl SlideSnapshot {
    /// Parse one complete `SlideContainer` with default resource bounds.
    pub fn parse(bytes: impl AsRef<[u8]>) -> Result<Self> {
        Self::from_bytes_with_limits(bytes.as_ref().to_vec(), SlideLimits::default())
    }

    /// Capture one complete slide without requiring a caller-side copy.
    pub fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        Self::from_bytes_with_limits(bytes, SlideLimits::default())
    }

    /// Parse one complete `SlideContainer` under explicit resource bounds.
    pub fn parse_with_limits(bytes: impl AsRef<[u8]>, limits: SlideLimits) -> Result<Self> {
        Self::from_bytes_with_limits(bytes.as_ref().to_vec(), limits)
    }

    /// Capture one complete slide under explicit resource bounds.
    pub fn from_bytes_with_limits(bytes: Vec<u8>, limits: SlideLimits) -> Result<Self> {
        super::validation::validate_limits(limits)?;
        if bytes.len() > limits.max_slide_bytes {
            return Err(crate::package::Error::InvalidFormat(
                "SlideContainer exceeds the configured slide byte limit".into(),
            ));
        }
        let parts = super::codec::locate(&bytes, limits)?;
        let drawing = bytes.get(parts.drawing.clone()).ok_or_else(|| {
            crate::package::Error::Corrupted("PPDrawing range is out of bounds".into())
        })?;
        let build_list = bytes.get(parts.build_list.clone()).ok_or_else(|| {
            crate::package::Error::Corrupted("BuildList range is out of bounds".into())
        })?;
        let diagram = DiagramSnapshot::parse_with_limits(build_list, drawing, limits.diagram)?;
        let bytes: Arc<[u8]> = Arc::from(bytes.into_boxed_slice());
        Ok(Self {
            revision: SlideRevision::from_bytes(&bytes),
            bytes,
            build_range: parts.build_list,
            drawing_range: parts.drawing,
            diagram,
            limits,
        })
    }

    /// Exact serialized `SlideContainer` bytes.
    pub fn bytes(&self) -> &[u8] {
        &self.bytes
    }

    /// Exact sibling `PPDrawing` payload bytes.
    pub fn drawing(&self) -> &[u8] {
        &self.bytes[self.drawing_range.clone()]
    }

    /// Exact `BuildList` record bytes inside the slide's `___PPT10` payload.
    pub fn build_list(&self) -> &[u8] {
        &self.bytes[self.build_range.clone()]
    }

    /// Source fingerprint of the complete owning slide payload.
    pub const fn revision(&self) -> SlideRevision {
        self.revision
    }

    /// Resource ceilings retained by this snapshot.
    pub const fn limits(&self) -> SlideLimits {
        self.limits
    }

    /// Number of typed diagram builds in the owning slide.
    pub fn len(&self) -> usize {
        self.diagram.len()
    }

    /// Whether the owning slide has no typed diagram builds.
    pub fn is_empty(&self) -> bool {
        self.diagram.is_empty()
    }

    /// Iterate typed diagram builds in BuildList source order.
    pub fn builds(&self) -> impl ExactSizeIterator<Item = Build> + '_ {
        self.diagram.builds()
    }

    /// Find one typed diagram build by its contextual identity.
    pub fn get(&self, id: Id) -> Option<Build> {
        self.diagram.get(id)
    }

    /// Recreate the borrowed diagram inventory over the slide-owned bytes.
    pub fn inventory(&self) -> Result<Inventory<'_>> {
        self.diagram.inventory()
    }

    /// Start an isolated edit over this exact slide source.
    pub fn edit(&self) -> super::transaction::SlideEditor {
        super::transaction::SlideEditor::open(self.clone())
    }
}
