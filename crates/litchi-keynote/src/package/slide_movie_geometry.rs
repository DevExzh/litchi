//! Exact-source, selector-first geometry transactions for file-backed movies.
//!
//! Geometry is a rendering-affecting edge: successful edits invalidate the
//! root previews.  This module owns an existing movie's position, displayed
//! size, and the supported rotation/reflection transform.  Other movie flags,
//! media, captions, playback, builds, metadata, and object allocation remain
//! opaque.

#![allow(
    clippy::map_err_ignore,
    clippy::needless_pass_by_value,
    clippy::wildcard_enum_match_arm,
    reason = "The semantic boundary redacts lower-layer failure details."
)]

use std::collections::{HashMap, HashSet};
use std::fmt;
use std::mem::size_of;
use std::sync::Arc;

use litchi_core::Position;
use litchi_iwa_archive::package::{Catalog, Entry, EntryEdit, ExactArtifacts};
use litchi_iwa_common::{decode_varint_from_bytes, varint::encoded_len, wire::WireView};
use litchi_iwa_core::{
    Archive, ArchiveObject, ArchiveReferenceKind, ArchiveReferenceOccurrence,
    ArchiveReferencePolicy, ArchiveReferenceVisitor, RawMessage, SnappyStream,
};
use litchi_iwa_protos::{keynote_movie_geometry_codec, package_metadata_codec};
use thiserror::Error;

use super::{Package, PhysicalSource, ReadError, SemanticLimitKind};
use crate::slide::media::geometry::{MovieFlipAxis, MovieGeometry, MovieTransform};
use crate::slide::media::{Point, Size};
use crate::{MovieKind, MovieSelector, SlideSelector};

const MOVIE_MESSAGE_TYPE: u32 = 3_007;
const MOVIE_STANDIN_MESSAGE_TYPE: u32 = 3_097;
const MOVIE_STYLE_MESSAGE_TYPE: u32 = 2_025;
// Native Keynote media controls use this style archive, while source-built
// fixtures historically used the shape-style archive above.  It is admitted
// only for the audio position path; file-backed movie geometry keeps its
// existing exact role contract.
const MOVIE_AUDIO_STYLE_MESSAGE_TYPE: u32 = 3_016;
const SLIDE_NODE_MESSAGE_TYPE: u32 = 4;
const SLIDE_MESSAGE_TYPE: u32 = 5;
const SLIDE_OWNED_DRAWABLES_FIELD: u32 = 7;
const SLIDE_Z_ORDER_FIELD: u32 = 42;
const MOVIE_SUPER_FIELD: u32 = 1;
const DRAWABLE_PARENT_FIELD: u32 = 2;
const MOVIE_REFLECTION_FLAG: u32 = 1 << 2;
const PACKAGE_METADATA_MESSAGE_TYPE: u32 = 11_006;
const TABLE_INFO_MESSAGE_TYPE: u32 = 6_000;
const TABLE_MODEL_MESSAGE_TYPE: u32 = 6_001;
const HEADER_BUCKET_MESSAGE_TYPE: u32 = 6_006;
const TABLE_STYLE_MESSAGE_TYPE: u32 = 6_003;
const TABLE_STYLE_PRESET_MESSAGE_TYPE: u32 = 6_008;
const TABLE_STYLE_NETWORK_MESSAGE_TYPE: u32 = 6_247;
const STYLESHEET_MESSAGE_TYPE: u32 = 401;
const MOVIE_GEOMETRY_COMPLETE: () = ();

/// Resource categories reported by a movie-geometry transaction.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum SlideMovieGeometryLimitKind {
    InputBytes,
    OutputBytes,
    WireBytes,
    Entries,
    EntryBytes,
    TotalBytes,
    Slides,
    References,
    WireFields,
    WireNesting,
    WireWork,
    GeometryBytes,
    Allocations,
    Retained,
    Scratch,
    Components,
}

impl fmt::Display for SlideMovieGeometryLimitKind {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.write_str(match self {
            Self::InputBytes => "input bytes",
            Self::OutputBytes => "output bytes",
            Self::WireBytes => "wire bytes",
            Self::Entries => "entries",
            Self::EntryBytes => "entry bytes",
            Self::TotalBytes => "total bytes",
            Self::Slides => "slides",
            Self::References => "references",
            Self::WireFields => "wire fields",
            Self::WireNesting => "wire nesting depth",
            Self::WireWork => "wire work",
            Self::GeometryBytes => "movie geometry bytes",
            Self::Allocations => "allocations",
            Self::Retained => "retained bytes",
            Self::Scratch => "scratch bytes",
            Self::Components => "components",
        })
    }
}

/// Content-redacted failure raised by a movie-geometry transaction.
#[derive(Debug, Clone, PartialEq, Eq, Error)]
#[non_exhaustive]
pub enum SlideMovieGeometryError {
    #[error("this Keynote source does not support physical movie-geometry edits")]
    UnsupportedSource,
    #[error("the requested Keynote movie-geometry graph is unsupported")]
    UnsupportedDependency,
    #[error("the selected Keynote movie is locked")]
    Locked,
    #[error("the Keynote movie-geometry selector is ambiguous")]
    AmbiguousSelector,
    #[error("the Keynote slide selector name cannot be empty")]
    EmptySlideName,
    #[error("the Keynote show has no slide matching the requested name")]
    SlideNameNotFound,
    #[error("the Keynote show has no slide at position {position:?}")]
    SlidePositionNotFound { position: Position },
    #[error("the selected Keynote slide has no movie at position {position:?}")]
    MoviePositionNotFound { position: Position },
    #[error("the Keynote movie-geometry source is invalid")]
    InvalidSource,
    #[error("Keynote movie geometry {kind} limit exceeded: observed {observed}, maximum {maximum}")]
    LimitExceeded {
        kind: SlideMovieGeometryLimitKind,
        observed: u64,
        maximum: u64,
    },
    #[error("could not allocate {amount} units for the Keynote movie-geometry transaction")]
    Allocation { amount: usize },
    #[error("the edited Keynote movie geometry failed semantic verification")]
    Verification,
    #[error("the Keynote movie-geometry patch does not match the exact source package")]
    PatchConflict,
}

#[derive(Debug, Clone, Copy)]
pub(crate) struct GeometryBudget {
    max_input: usize,
    max_output: usize,
    max_fields: usize,
    max_work: usize,
    max_nesting: usize,
    max_references: usize,
    max_allocations: usize,
    max_retained: usize,
    max_scratch: usize,
    input: usize,
    output: usize,
    fields: usize,
    work: usize,
    nesting: usize,
    references: usize,
    allocations: usize,
    retained: usize,
    scratch: usize,
}

impl GeometryBudget {
    pub(crate) fn new(package: &Package) -> Result<Self, SlideMovieGeometryError> {
        std::hint::black_box(MOVIE_GEOMETRY_COMPLETE);
        let wire = package.wire_limits().map_err(map_wire_error)?;
        let source: usize = package
            .state
            .options
            .archive()
            .max_input_bytes()
            .try_into()
            .map_err(|_| SlideMovieGeometryError::InvalidSource)?;
        let aggregate = source
            .checked_mul(4)
            .ok_or(SlideMovieGeometryError::InvalidSource)?;
        Ok(Self {
            max_input: aggregate,
            max_output: aggregate,
            max_fields: wire.max_fields(),
            max_work: wire.max_rewrite_work(),
            max_nesting: wire.max_nesting(),
            max_references: package.semantic_limits().max_references(),
            max_allocations: aggregate,
            max_retained: aggregate,
            max_scratch: aggregate,
            input: 0,
            output: 0,
            fields: 0,
            work: 0,
            nesting: 0,
            references: 0,
            allocations: 0,
            retained: 0,
            scratch: 0,
        })
    }

    fn add(
        current: &mut usize,
        amount: usize,
        maximum: usize,
        kind: SlideMovieGeometryLimitKind,
    ) -> Result<(), SlideMovieGeometryError> {
        let value = current
            .checked_add(amount)
            .ok_or(SlideMovieGeometryError::InvalidSource)?;
        if value > maximum {
            return Err(SlideMovieGeometryError::LimitExceeded {
                kind,
                observed: value as u64,
                maximum: maximum as u64,
            });
        }
        *current = value;
        Ok(())
    }

    pub(crate) fn source(&mut self, bytes: usize) -> Result<(), SlideMovieGeometryError> {
        Self::add(
            &mut self.input,
            bytes,
            self.max_input,
            SlideMovieGeometryLimitKind::InputBytes,
        )
    }

    pub(crate) fn output(&mut self, bytes: usize) -> Result<(), SlideMovieGeometryError> {
        Self::add(
            &mut self.output,
            bytes,
            self.max_output,
            SlideMovieGeometryLimitKind::OutputBytes,
        )
    }

    pub(crate) fn fields(&mut self, amount: usize) -> Result<(), SlideMovieGeometryError> {
        Self::add(
            &mut self.fields,
            amount,
            self.max_fields,
            SlideMovieGeometryLimitKind::WireFields,
        )
    }

    pub(crate) fn preflight_output(&self, bytes: usize) -> Result<(), SlideMovieGeometryError> {
        let observed = self
            .output
            .checked_add(bytes)
            .ok_or(SlideMovieGeometryError::InvalidSource)?;
        if observed > self.max_output {
            return Err(SlideMovieGeometryError::LimitExceeded {
                kind: SlideMovieGeometryLimitKind::OutputBytes,
                observed: observed as u64,
                maximum: self.max_output as u64,
            });
        }
        Ok(())
    }

    pub(crate) fn preflight_work(&self, bytes: usize) -> Result<(), SlideMovieGeometryError> {
        let observed = self
            .work
            .checked_add(bytes)
            .ok_or(SlideMovieGeometryError::InvalidSource)?;
        if observed > self.max_work {
            return Err(SlideMovieGeometryError::LimitExceeded {
                kind: SlideMovieGeometryLimitKind::WireWork,
                observed: observed as u64,
                maximum: self.max_work as u64,
            });
        }
        Ok(())
    }

    pub(crate) fn references(&mut self, amount: usize) -> Result<(), SlideMovieGeometryError> {
        Self::add(
            &mut self.references,
            amount,
            self.max_references,
            SlideMovieGeometryLimitKind::References,
        )
    }

    pub(crate) fn work(&mut self, bytes: usize) -> Result<(), SlideMovieGeometryError> {
        Self::add(
            &mut self.work,
            bytes,
            self.max_work,
            SlideMovieGeometryLimitKind::WireWork,
        )
    }

    pub(crate) fn allocations(&mut self, amount: usize) -> Result<(), SlideMovieGeometryError> {
        Self::add(
            &mut self.allocations,
            amount,
            self.max_allocations,
            SlideMovieGeometryLimitKind::Allocations,
        )
    }

    pub(crate) fn retained(&mut self, amount: usize) -> Result<(), SlideMovieGeometryError> {
        Self::add(
            &mut self.retained,
            amount,
            self.max_retained,
            SlideMovieGeometryLimitKind::Retained,
        )
    }

    pub(crate) fn scratch(&mut self, amount: usize) -> Result<(), SlideMovieGeometryError> {
        Self::add(
            &mut self.scratch,
            amount,
            self.max_scratch,
            SlideMovieGeometryLimitKind::Scratch,
        )
    }

    pub(crate) fn physical(&mut self, bytes: usize) -> Result<(), SlideMovieGeometryError> {
        self.source(bytes)?;
        self.work(bytes)?;
        self.output(
            bytes
                .checked_mul(2)
                .ok_or(SlideMovieGeometryError::InvalidSource)?,
        )
    }

    pub(crate) fn reassembly(
        &mut self,
        requirements: litchi_iwa_archive::package::ReassemblyExecutionRequirements,
    ) -> Result<(), SlideMovieGeometryError> {
        self.output(requirements.output_bytes())?;
        self.allocations(requirements.allocations())?;
        self.retained(requirements.retained_bytes())?;
        self.scratch(requirements.scratch_bytes())?;
        self.work(requirements.output_bytes())
    }

    pub(crate) fn candidate_reopen(&mut self, bytes: usize) -> Result<(), SlideMovieGeometryError> {
        self.source(bytes)?;
        self.work(bytes)
    }

    /// Debit the complete package scan performed by semantic validation.
    ///
    /// Validation walks the retained physical source independently of the
    /// focused graph selectors, so callers must account for that scan before
    /// invoking [`Package::validate`].
    pub(crate) fn validate_package(
        &mut self,
        package: &Package,
    ) -> Result<(), SlideMovieGeometryError> {
        let mut work = 0usize;
        let mut allocations = 0usize;
        let mut fields = 0usize;
        let mut logical_bytes = 0usize;
        let mut object_count = 0usize;
        let mut message_count = 0usize;
        for component in package.state.source.components().iter() {
            let archive_bytes = component.archive().encoded_len().map_err(map_core_error)?;
            work = work
                .checked_add(component.name().len())
                .and_then(|value| value.checked_add(archive_bytes))
                .ok_or(SlideMovieGeometryError::InvalidSource)?;
            logical_bytes = logical_bytes
                .checked_add(component.name().len())
                .and_then(|value| value.checked_add(archive_bytes))
                .ok_or(SlideMovieGeometryError::InvalidSource)?;
            allocations = allocations
                .checked_add(1)
                .ok_or(SlideMovieGeometryError::InvalidSource)?;
            for object in &component.archive().objects {
                object_count = object_count
                    .checked_add(1)
                    .ok_or(SlideMovieGeometryError::InvalidSource)?;
                work = work
                    .checked_add(1)
                    .and_then(|value| value.checked_add(object.messages.len()))
                    .and_then(|value| value.checked_add(object.archive_info.message_infos.len()))
                    .ok_or(SlideMovieGeometryError::InvalidSource)?;
                allocations = allocations
                    .checked_add(1)
                    .and_then(|value| value.checked_add(object.messages.len()))
                    .and_then(|value| value.checked_add(object.archive_info.message_infos.len()))
                    .ok_or(SlideMovieGeometryError::InvalidSource)?;
                for (message_index, message) in object.messages.iter().enumerate() {
                    message_count = message_count
                        .checked_add(1)
                        .ok_or(SlideMovieGeometryError::InvalidSource)?;
                    work = work
                        .checked_add(message.data.len())
                        .ok_or(SlideMovieGeometryError::InvalidSource)?;
                    logical_bytes = logical_bytes
                        .checked_add(message.data.len())
                        .ok_or(SlideMovieGeometryError::InvalidSource)?;
                    let info = object
                        .archive_info
                        .message_infos
                        .get(message_index)
                        .ok_or(SlideMovieGeometryError::InvalidSource)?;
                    fields = fields
                        .checked_add(info.field_infos.len())
                        .ok_or(SlideMovieGeometryError::InvalidSource)?;
                    let mut references = info
                        .object_references
                        .len()
                        .checked_add(info.data_references.len())
                        .and_then(|value| value.checked_add(info.field_infos.len()))
                        .ok_or(SlideMovieGeometryError::InvalidSource)?;
                    for field in &info.field_infos {
                        references = references
                            .checked_add(field.path.path.len())
                            .and_then(|value| value.checked_add(field.object_references.len()))
                            .and_then(|value| value.checked_add(field.data_references.len()))
                            .ok_or(SlideMovieGeometryError::InvalidSource)?;
                    }
                    work = work
                        .checked_add(references)
                        .ok_or(SlideMovieGeometryError::InvalidSource)?;
                }
            }
        }
        let object_storage = object_count
            .checked_mul(size_of::<ArchiveObject>())
            .ok_or(SlideMovieGeometryError::InvalidSource)?;
        let message_storage = message_count
            .checked_mul(size_of::<RawMessage>())
            .ok_or(SlideMovieGeometryError::InvalidSource)?;
        let semantic_bytes = logical_bytes
            .checked_add(object_storage)
            .and_then(|value| value.checked_add(message_storage))
            .ok_or(SlideMovieGeometryError::InvalidSource)?;
        // Source/candidate bytes are charged at the physical ingress and
        // reopen seams.  This pass accounts for semantic traversal work and
        // its known archive/message staging. Reference counts are deliberately
        // not charged here: validation independently enforces the package-wide
        // semantic reference ceiling, while repeated traversal is work above.
        // Semantic wire fields and nesting are per-payload ceilings enforced
        // by Package::validate's decoders. Opaque archive payload bytes are
        // work/storage, not parsed fields; only the header inventory traversed
        // above contributes to this operation's aggregate field ledger.
        self.fields(fields)?;
        self.work(work)?;
        self.allocations(allocations)?;
        self.retained(semantic_bytes)?;
        self.scratch(semantic_bytes)?;
        package.validate().map_err(map_read_error)
    }

    fn residual(
        &self,
        package: &Package,
    ) -> Result<litchi_iwa_common::WireLimits, SlideMovieGeometryError> {
        let base = package.wire_limits().map_err(map_wire_error)?;
        let input = self.remaining_wire(
            self.input,
            self.max_input,
            SlideMovieGeometryLimitKind::InputBytes,
        )?;
        let fields = self.remaining_wire(
            self.fields,
            self.max_fields,
            SlideMovieGeometryLimitKind::WireFields,
        )?;
        let work = self.remaining_wire(
            self.work,
            self.max_work,
            SlideMovieGeometryLimitKind::WireWork,
        )?;
        let nesting = self.remaining_wire(
            self.nesting,
            self.max_nesting,
            SlideMovieGeometryLimitKind::WireNesting,
        )?;
        base.with_input_bytes(base.max_input_bytes().min(input))
            .and_then(|v| v.with_fields(base.max_fields().min(fields)))
            .and_then(|v| v.with_rewrite_work(base.max_rewrite_work().min(work)))
            .and_then(|v| v.with_nesting(base.max_nesting().min(nesting)))
            .map_err(map_wire_error)
    }

    fn remaining_wire(
        &self,
        used: usize,
        maximum: usize,
        kind: SlideMovieGeometryLimitKind,
    ) -> Result<usize, SlideMovieGeometryError> {
        maximum
            .checked_sub(used)
            .filter(|remaining| *remaining > 0)
            .ok_or(SlideMovieGeometryError::LimitExceeded {
                kind,
                observed: used.saturating_add(1) as u64,
                maximum: maximum as u64,
            })
    }

    fn remaining_output(&self) -> Result<usize, SlideMovieGeometryError> {
        self.max_output
            .checked_sub(self.output)
            .filter(|value| *value > 0)
            .ok_or(SlideMovieGeometryError::LimitExceeded {
                kind: SlideMovieGeometryLimitKind::OutputBytes,
                observed: self.output.saturating_add(1) as u64,
                maximum: self.max_output as u64,
            })
    }

    fn remaining_allocations(&self) -> Result<usize, SlideMovieGeometryError> {
        self.max_allocations
            .checked_sub(self.allocations)
            .filter(|value| *value > 0)
            .ok_or(SlideMovieGeometryError::LimitExceeded {
                kind: SlideMovieGeometryLimitKind::Allocations,
                observed: self.allocations.saturating_add(1) as u64,
                maximum: self.max_allocations as u64,
            })
    }

    fn remaining_retained(&self) -> Result<usize, SlideMovieGeometryError> {
        self.max_retained
            .checked_sub(self.retained)
            .filter(|value| *value > 0)
            .ok_or(SlideMovieGeometryError::LimitExceeded {
                kind: SlideMovieGeometryLimitKind::Retained,
                observed: self.retained.saturating_add(1) as u64,
                maximum: self.max_retained as u64,
            })
    }

    fn remaining_scratch(&self) -> Result<usize, SlideMovieGeometryError> {
        self.max_scratch
            .checked_sub(self.scratch)
            .filter(|value| *value > 0)
            .ok_or(SlideMovieGeometryError::LimitExceeded {
                kind: SlideMovieGeometryLimitKind::Scratch,
                observed: self.scratch.saturating_add(1) as u64,
                maximum: self.max_scratch as u64,
            })
    }

    fn remaining_references(&self) -> Result<usize, SlideMovieGeometryError> {
        self.max_references
            .checked_sub(self.references)
            .filter(|value| *value > 0)
            .ok_or(SlideMovieGeometryError::LimitExceeded {
                kind: SlideMovieGeometryLimitKind::References,
                observed: self.references.saturating_add(1) as u64,
                maximum: self.max_references as u64,
            })
    }

    pub(crate) fn codec_report(
        &mut self,
        report: keynote_movie_geometry_codec::DecodeReport,
    ) -> Result<(), SlideMovieGeometryError> {
        self.source(report.input_bytes())?;
        self.allocations(report.allocations())?;
        self.retained(report.retained_bytes())?;
        self.scratch(report.scratch_bytes())?;
        Self::add(
            &mut self.fields,
            report.fields(),
            self.max_fields,
            SlideMovieGeometryLimitKind::WireFields,
        )?;
        Self::add(
            &mut self.work,
            report.work_bytes(),
            self.max_work,
            SlideMovieGeometryLimitKind::WireWork,
        )?;
        Self::add(
            &mut self.nesting,
            report.max_depth() as usize,
            self.max_nesting,
            SlideMovieGeometryLimitKind::WireNesting,
        )?;
        Ok(())
    }

    pub(crate) fn codec_requirements(
        &mut self,
        requirements: keynote_movie_geometry_codec::RewriteExecutionRequirements,
    ) -> Result<(), SlideMovieGeometryError> {
        self.output(requirements.output_bytes)?;
        Self::add(
            &mut self.fields,
            requirements.fields,
            self.max_fields,
            SlideMovieGeometryLimitKind::WireFields,
        )?;
        Self::add(
            &mut self.work,
            requirements.work_bytes,
            self.max_work,
            SlideMovieGeometryLimitKind::WireWork,
        )?;
        self.allocations(requirements.allocations)?;
        self.retained(requirements.retained_bytes)?;
        self.scratch(requirements.scratch_bytes)?;
        Self::add(
            &mut self.nesting,
            usize::try_from(requirements.max_depth).unwrap_or(usize::MAX),
            self.max_nesting,
            SlideMovieGeometryLimitKind::WireNesting,
        )?;
        Ok(())
    }

    fn metadata_report(
        &mut self,
        report: package_metadata_codec::RewriteReport,
    ) -> Result<(), SlideMovieGeometryError> {
        self.source(report.input_bytes())?;
        self.output(report.output_bytes())?;
        self.fields(report.fields())?;
        self.work(report.work_bytes())?;
        self.references(report.references_scanned())?;
        self.allocations(report.allocations())?;
        self.retained(report.retained_bytes())?;
        self.scratch(report.scratch_bytes())?;
        self.nesting = self.nesting.max(report.max_depth() as usize);
        if self.nesting > self.max_nesting {
            return Err(SlideMovieGeometryError::LimitExceeded {
                kind: SlideMovieGeometryLimitKind::WireNesting,
                observed: self.nesting as u64,
                maximum: self.max_nesting as u64,
            });
        }
        Ok(())
    }
}

/// One mutable geometry value staged against an immutable package snapshot.
pub struct SlideMovieGeometryEdit<'a> {
    source: &'a Package,
    budget: GeometryBudget,
    selection: GeometrySelection,
    after: MovieGeometry,
    after_transform: MovieTransform,
}

impl fmt::Debug for SlideMovieGeometryEdit<'_> {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.debug_struct("SlideMovieGeometryEdit")
            .field("slide_position", &self.selection.slide_position)
            .field("movie_position", &self.selection.movie_position)
            .finish_non_exhaustive()
    }
}

impl<'a> SlideMovieGeometryEdit<'a> {
    fn new<'slide>(
        source: &'a Package,
        slide: impl Into<SlideSelector<'slide>>,
        movie: impl Into<MovieSelector>,
    ) -> Result<Self, SlideMovieGeometryError> {
        let mut budget = GeometryBudget::new(source)?;
        let source_bytes = physical_catalog(source)?.source_bytes().len();
        budget.source(source_bytes)?;
        let selection = select_movie_with_budget(source, slide.into(), movie.into(), &mut budget)?;
        let before = selection
            .before
            .ok_or(SlideMovieGeometryError::UnsupportedDependency)?;
        let before_transform = selection
            .before_transform
            .ok_or(SlideMovieGeometryError::UnsupportedDependency)?;
        Ok(Self {
            source,
            budget,
            selection,
            after: before,
            after_transform: before_transform,
        })
    }

    #[must_use]
    pub const fn slide_position(&self) -> Position {
        self.selection.slide_position
    }
    #[must_use]
    pub const fn movie_position(&self) -> Position {
        self.selection.movie_position
    }
    #[must_use]
    pub const fn before(&self) -> Option<MovieGeometry> {
        self.selection.before
    }
    #[must_use]
    pub const fn after(&self) -> MovieGeometry {
        self.after
    }

    /// Return the transform observed when this edit began.
    #[must_use]
    pub const fn before_transform(&self) -> Option<MovieTransform> {
        self.selection.before_transform
    }

    /// Return the transform currently staged by this edit.
    #[must_use]
    pub const fn after_transform(&self) -> MovieTransform {
        self.after_transform
    }

    pub fn set(mut self, geometry: MovieGeometry) -> Result<Self, SlideMovieGeometryError> {
        self.after = geometry;
        Ok(self)
    }

    /// Stage a transform while preserving the existing position and size.
    pub fn set_transform(
        mut self,
        transform: MovieTransform,
    ) -> Result<Self, SlideMovieGeometryError> {
        self.after_transform = transform;
        Ok(self)
    }

    /// Stage one native Arrange flip while preserving the existing geometry.
    pub fn flip(self, axis: MovieFlipAxis) -> Result<Self, SlideMovieGeometryError> {
        if self.selection.native_flags.is_none() {
            return Err(SlideMovieGeometryError::UnsupportedDependency);
        }
        let transform = self.after_transform.flipped(axis);
        self.set_transform(transform)
    }

    pub fn commit(self) -> Result<SlideMovieGeometryCommit, SlideMovieGeometryError> {
        commit_edit(
            self.source,
            &self.selection,
            self.after,
            self.after_transform,
            self.budget,
        )
    }
}

/// Exact-source checked reversible movie-geometry patch.
#[derive(Clone, PartialEq)]
pub struct SlideMovieGeometryPatch {
    artifacts: ExactArtifacts,
    selection: GeometrySelection,
    before: MovieGeometry,
    after: MovieGeometry,
    before_transform: MovieTransform,
    after_transform: MovieTransform,
    touched_components: usize,
    deleted_previews: usize,
    source_previews_absent: bool,
    target_previews_absent: bool,
}

impl fmt::Debug for SlideMovieGeometryPatch {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.debug_struct("SlideMovieGeometryPatch")
            .field("slide_position", &self.selection.slide_position)
            .field("movie_position", &self.selection.movie_position)
            .finish_non_exhaustive()
    }
}

impl SlideMovieGeometryPatch {
    #[must_use]
    pub const fn slide_position(&self) -> Position {
        self.selection.slide_position
    }
    #[must_use]
    pub const fn movie_position(&self) -> Position {
        self.selection.movie_position
    }
    #[must_use]
    pub const fn before(&self) -> MovieGeometry {
        self.before
    }
    #[must_use]
    pub const fn after(&self) -> MovieGeometry {
        self.after
    }

    /// Return the transform observed before this patch.
    #[must_use]
    pub const fn before_transform(&self) -> MovieTransform {
        self.before_transform
    }

    /// Return the transform produced by this patch.
    #[must_use]
    pub const fn after_transform(&self) -> MovieTransform {
        self.after_transform
    }
    #[must_use]
    pub const fn source_fingerprint(&self) -> u64 {
        self.artifacts.source_fingerprint()
    }
    #[must_use]
    pub const fn target_fingerprint(&self) -> u64 {
        self.artifacts.target_fingerprint()
    }
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before == self.after
            && self.before_transform == self.after_transform
            && self.artifacts.is_byte_noop()
    }
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            artifacts: self.artifacts.inverse(),
            selection: self.selection.clone(),
            before: self.after,
            after: self.before,
            before_transform: self.after_transform,
            after_transform: self.before_transform,
            touched_components: self.touched_components,
            deleted_previews: self.deleted_previews,
            source_previews_absent: self.target_previews_absent,
            target_previews_absent: self.source_previews_absent,
        }
    }
}

/// Compact movie-geometry publication diagnostics.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SlideMovieGeometryDiagnostics {
    changed: bool,
    touched_components: usize,
    deleted_previews: usize,
    full_reparse_performed: bool,
}

impl SlideMovieGeometryDiagnostics {
    const fn unchanged() -> Self {
        Self {
            changed: false,
            touched_components: 0,
            deleted_previews: 0,
            full_reparse_performed: false,
        }
    }
    const fn published(touched_components: usize, deleted_previews: usize) -> Self {
        Self {
            changed: true,
            touched_components,
            deleted_previews,
            full_reparse_performed: true,
        }
    }
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }
    #[must_use]
    pub const fn touched_components(self) -> usize {
        self.touched_components
    }
    #[must_use]
    pub const fn deleted_previews(self) -> usize {
        self.deleted_previews
    }
    #[must_use]
    pub const fn full_reparse_performed(self) -> bool {
        self.full_reparse_performed
    }
}

/// Fully verified result of one movie-geometry transaction.
#[must_use = "a Keynote movie-geometry commit contains the validated package snapshot"]
#[derive(Debug)]
pub struct SlideMovieGeometryCommit {
    package: Package,
    patch: SlideMovieGeometryPatch,
    diagnostics: SlideMovieGeometryDiagnostics,
}

impl SlideMovieGeometryCommit {
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.package
    }
    #[must_use]
    pub fn into_package(self) -> Package {
        self.package
    }
    #[must_use]
    pub const fn patch(&self) -> &SlideMovieGeometryPatch {
        &self.patch
    }
    #[must_use]
    pub const fn diagnostics(&self) -> &SlideMovieGeometryDiagnostics {
        &self.diagnostics
    }
}

#[derive(Clone, PartialEq)]
pub(crate) struct GeometrySelection {
    pub(crate) slide_position: Position,
    pub(crate) movie_position: Position,
    pub(crate) slide_identifier: u64,
    pub(crate) node_identifier: u64,
    pub(crate) movie_identifier: u64,
    pub(crate) message_index: usize,
    pub(crate) slide_component_name: Arc<str>,
    pub(crate) before: Option<MovieGeometry>,
    pub(crate) before_position: Option<Point>,
    pub(crate) before_transform: Option<MovieTransform>,
    pub(crate) native_flags: Option<u32>,
    pub(crate) native_angle: Option<f32>,
    pub(crate) locked: bool,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum GeometryMediaKind {
    FileMovie,
    Audio,
}

#[derive(Clone, Copy)]
struct MovieArchiveReference {
    identifier: u64,
    expected_message_type: u32,
}

impl fmt::Debug for GeometrySelection {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.debug_struct("GeometrySelection")
            .field("slide_position", &self.slide_position)
            .field("movie_position", &self.movie_position)
            .field("has_before", &self.before.is_some())
            .finish_non_exhaustive()
    }
}

impl Package {
    /// Read the validated position and displayed size of one file-backed movie.
    pub fn slide_movie_geometry<'slide>(
        &self,
        slide: impl Into<SlideSelector<'slide>>,
        movie: impl Into<MovieSelector>,
    ) -> Result<Option<MovieGeometry>, SlideMovieGeometryError> {
        let mut budget = GeometryBudget::new(self)?;
        let source_bytes = physical_catalog(self)?.source_bytes().len();
        budget.source(source_bytes)?;
        Ok(select_movie_with_budget(self, slide.into(), movie.into(), &mut budget)?.before)
    }

    /// Read the validated rotation and reflection state of one file-backed movie.
    pub fn slide_movie_transform<'slide>(
        &self,
        slide: impl Into<SlideSelector<'slide>>,
        movie: impl Into<MovieSelector>,
    ) -> Result<Option<MovieTransform>, SlideMovieGeometryError> {
        let mut budget = GeometryBudget::new(self)?;
        let source_bytes = physical_catalog(self)?.source_bytes().len();
        budget.source(source_bytes)?;
        Ok(
            select_movie_with_budget(self, slide.into(), movie.into(), &mut budget)?
                .before_transform,
        )
    }

    /// Begin an exact immutable edit of an existing file-backed movie geometry.
    pub fn edit_slide_movie_geometry<'slide>(
        &self,
        slide: impl Into<SlideSelector<'slide>>,
        movie: impl Into<MovieSelector>,
    ) -> Result<SlideMovieGeometryEdit<'_>, SlideMovieGeometryError> {
        SlideMovieGeometryEdit::new(self, slide, movie)
    }

    /// Apply an exact-source checked movie-geometry patch.
    pub fn apply_slide_movie_geometry(
        &self,
        patch: &SlideMovieGeometryPatch,
    ) -> Result<SlideMovieGeometryCommit, SlideMovieGeometryError> {
        let mut budget = GeometryBudget::new(self)?;
        let catalog = physical_catalog(self)?;
        let source_len = catalog.source_bytes().len();
        budget.source(source_len)?;
        // An exact-artifact authorization may compare the full source when
        // the caller hands us a byte-equal, independently owned Arc.  Debit
        // that comparison before consulting the patch so a rejected patch
        // cannot bypass the operation ledger.
        budget.work(source_len)?;
        let shared_source = catalog.shared_source();
        if !patch.artifacts.authorizes_source(&shared_source) {
            return Err(SlideMovieGeometryError::PatchConflict);
        }
        if previews_absent(self)? != patch.source_previews_absent {
            return Err(SlideMovieGeometryError::PatchConflict);
        }
        let current = select_movie_with_budget(
            self,
            SlideSelector::position(patch.selection.slide_position),
            MovieSelector::position(patch.selection.movie_position),
            &mut budget,
        )?;
        if !same_selection(&current, &patch.selection)
            || current.before != Some(patch.before)
            || current.before_transform != Some(patch.before_transform)
        {
            return Err(SlideMovieGeometryError::PatchConflict);
        }
        if patch.is_noop() {
            self.validate().map_err(map_read_error)?;
            return Ok(SlideMovieGeometryCommit {
                package: self.snapshot(),
                patch: patch.clone(),
                diagnostics: SlideMovieGeometryDiagnostics::unchanged(),
            });
        }
        if patch.selection.locked {
            return Err(SlideMovieGeometryError::Locked);
        }
        if !catalog.source_is_exact() {
            return Err(SlideMovieGeometryError::UnsupportedSource);
        }
        reopen_target_patch(self, patch, &mut budget)
    }
}

fn commit_edit(
    source: &Package,
    selection: &GeometrySelection,
    after: MovieGeometry,
    after_transform: MovieTransform,
    mut budget: GeometryBudget,
) -> Result<SlideMovieGeometryCommit, SlideMovieGeometryError> {
    let before = selection
        .before
        .ok_or(SlideMovieGeometryError::UnsupportedDependency)?;
    let before_transform = selection
        .before_transform
        .ok_or(SlideMovieGeometryError::UnsupportedDependency)?;
    if before == after && before_transform == after_transform {
        let catalog = physical_catalog(source)?;
        // ExactArtifacts fingerprints the retained source.  Reuse the
        // catalog's immutable Arc rather than copying the complete package
        // on the no-op path, while charging the fingerprint walk itself.
        budget.work(source.source_bytes().len())?;
        let bytes = catalog.shared_source();
        return Ok(SlideMovieGeometryCommit {
            package: source.snapshot(),
            patch: SlideMovieGeometryPatch {
                artifacts: ExactArtifacts::new(Arc::clone(&bytes), bytes),
                selection: selection.clone(),
                before,
                after,
                before_transform,
                after_transform,
                touched_components: 0,
                deleted_previews: 0,
                source_previews_absent: previews_absent(source)?,
                target_previews_absent: previews_absent(source)?,
            },
            diagnostics: SlideMovieGeometryDiagnostics::unchanged(),
        });
    }
    if selection.locked {
        return Err(SlideMovieGeometryError::Locked);
    }
    let catalog = physical_catalog(source)?;
    if !catalog.source_is_exact() {
        return Err(SlideMovieGeometryError::UnsupportedSource);
    }
    let current = select_movie_with_budget(
        source,
        SlideSelector::position(selection.slide_position),
        MovieSelector::position(selection.movie_position),
        &mut budget,
    )?;
    if !same_selection(&current, selection)
        || current.before != Some(before)
        || current.before_transform != Some(before_transform)
    {
        return Err(SlideMovieGeometryError::PatchConflict);
    }
    let (candidate, deleted_previews) =
        rewrite_movie(source, selection, after, after_transform, &mut budget)?;
    candidate.validate().map_err(map_read_error)?;
    if !previews_absent(&candidate)? {
        return Err(SlideMovieGeometryError::Verification);
    }
    let selected = select_movie_with_budget(
        &candidate,
        SlideSelector::position(selection.slide_position),
        MovieSelector::position(selection.movie_position),
        &mut budget,
    )?;
    if !same_selection(&selected, selection)
        || selected.before != Some(after)
        || selected.before_transform != Some(after_transform)
    {
        return Err(SlideMovieGeometryError::Verification);
    }
    verify_locality(
        source,
        &candidate,
        selection.slide_component_name.as_ref(),
        selection.movie_identifier,
        selection.message_index,
        true,
        &mut budget,
    )?;
    let source_bytes = catalog.shared_source();
    let target = physical_catalog(&candidate)?.shared_source();
    Ok(SlideMovieGeometryCommit {
        package: candidate,
        patch: SlideMovieGeometryPatch {
            artifacts: ExactArtifacts::new(source_bytes, target),
            selection: selection.clone(),
            before,
            after,
            before_transform,
            after_transform,
            touched_components: 1,
            deleted_previews,
            source_previews_absent: previews_absent(source)?,
            target_previews_absent: true,
        },
        diagnostics: SlideMovieGeometryDiagnostics::published(1, deleted_previews),
    })
}

fn reopen_target_patch(
    source: &Package,
    patch: &SlideMovieGeometryPatch,
    budget: &mut GeometryBudget,
) -> Result<SlideMovieGeometryCommit, SlideMovieGeometryError> {
    budget.candidate_reopen(patch.artifacts.target().len())?;
    budget.allocations(1)?;
    budget.retained(patch.artifacts.target().len())?;
    budget.scratch(patch.artifacts.target().len())?;
    let candidate =
        Package::from_source_with_options(patch.artifacts.target(), source.state.options)
            .map_err(map_read_error)?;
    candidate.validate().map_err(map_read_error)?;
    if previews_absent(&candidate)? != patch.target_previews_absent {
        return Err(SlideMovieGeometryError::Verification);
    }
    let selected = select_movie_with_budget(
        &candidate,
        SlideSelector::position(patch.selection.slide_position),
        MovieSelector::position(patch.selection.movie_position),
        budget,
    )?;
    if !same_selection(&selected, &patch.selection)
        || selected.before != Some(patch.after)
        || selected.before_transform != Some(patch.after_transform)
    {
        return Err(SlideMovieGeometryError::Verification);
    }
    verify_locality(
        source,
        &candidate,
        patch.selection.slide_component_name.as_ref(),
        patch.selection.movie_identifier,
        patch.selection.message_index,
        patch.target_previews_absent,
        budget,
    )?;
    Ok(SlideMovieGeometryCommit {
        package: candidate,
        patch: patch.clone(),
        diagnostics: SlideMovieGeometryDiagnostics::published(
            patch.touched_components,
            patch.deleted_previews,
        ),
    })
}

fn is_media_data_name(name: &str, media_kind: GeometryMediaKind) -> bool {
    let bytes = name.as_bytes();
    if bytes.len() < 5 || !bytes[..5].eq_ignore_ascii_case(b"data/") {
        return false;
    }
    if matches!(media_kind, GeometryMediaKind::Audio) {
        // Audio data names are producer-defined and may have no reliable
        // extension.  The selected movie's data references and metadata
        // ownership remain the authoritative admission proof below.
        return true;
    }
    debug_assert!(matches!(media_kind, GeometryMediaKind::FileMovie));
    let suffixes: &[&[u8]] = &[b".mov".as_slice(), b".mp4", b".m4v", b".mpeg", b".mpg"];
    suffixes.iter().any(|suffix| {
        bytes.len() >= suffix.len()
            && bytes[bytes.len() - suffix.len()..].eq_ignore_ascii_case(suffix)
    })
}

fn select_movie_with_budget(
    package: &Package,
    slide_selector: SlideSelector<'_>,
    movie_selector: MovieSelector,
    budget: &mut GeometryBudget,
) -> Result<GeometrySelection, SlideMovieGeometryError> {
    select_media_with_budget(
        package,
        slide_selector,
        movie_selector,
        GeometryMediaKind::FileMovie,
        budget,
    )
}

pub(crate) fn select_audio_with_budget(
    package: &Package,
    slide_selector: SlideSelector<'_>,
    movie_selector: MovieSelector,
    budget: &mut GeometryBudget,
) -> Result<GeometrySelection, SlideMovieGeometryError> {
    select_media_with_budget(
        package,
        slide_selector,
        movie_selector,
        GeometryMediaKind::Audio,
        budget,
    )
}

fn select_media_with_budget(
    package: &Package,
    slide_selector: SlideSelector<'_>,
    movie_selector: MovieSelector,
    media_kind: GeometryMediaKind,
    budget: &mut GeometryBudget,
) -> Result<GeometrySelection, SlideMovieGeometryError> {
    let catalog = physical_catalog(package)?;
    let mut has_movie_media = false;
    for entry in catalog.package().iter() {
        budget.work(
            entry
                .name()
                .len()
                .checked_add(entry.data().len())
                .and_then(|value| value.checked_add(1))
                .ok_or(SlideMovieGeometryError::InvalidSource)?,
        )?;
        if !entry.is_opaque() && is_media_data_name(entry.name(), media_kind) {
            has_movie_media = true;
        }
    }
    if !has_movie_media {
        return Err(SlideMovieGeometryError::UnsupportedDependency);
    }
    let slide_position = resolve_slide_position(package, slide_selector, budget)?;
    let record = package
        .slide_record_at(slide_position.get())
        .map_err(map_read_error)?
        .ok_or(SlideMovieGeometryError::SlidePositionNotFound {
            position: slide_position,
        })?;
    let (component_name, slide) = package
        .object_with_component(record.slide_identifier)
        .ok_or(SlideMovieGeometryError::InvalidSource)?;
    let (slide_message_index, slide_payload) = unique_message(slide, SLIDE_MESSAGE_TYPE, budget)?;
    let limits = package.wire_limits().map_err(map_wire_error)?;
    let ids = repeated_references(slide_payload, SLIDE_OWNED_DRAWABLES_FIELD, limits, budget)?;
    let z_order = repeated_references(slide_payload, SLIDE_Z_ORDER_FIELD, limits, budget)?;
    validate_slide_archive_info_references(slide, slide_message_index, &ids, &z_order, budget)?;
    let drawable_capacity = ids
        .len()
        .checked_add(z_order.len())
        .ok_or(SlideMovieGeometryError::InvalidSource)?;
    budget.allocations(drawable_capacity)?;
    let mut validated_drawables = HashSet::new();
    validated_drawables
        .try_reserve(drawable_capacity)
        .map_err(|_| SlideMovieGeometryError::Allocation {
            amount: drawable_capacity,
        })?;
    for identifier in ids.iter().chain(z_order.iter()) {
        if !validated_drawables.insert(*identifier) {
            continue;
        }
        let (_, object) = package
            .object_with_component(*identifier)
            .ok_or(SlideMovieGeometryError::InvalidSource)?;
        validate_drawable_message_headers(object, budget)?;
        if object.messages.iter().any(|message| {
            is_known_movie_role(message.type_) && message.type_ != MOVIE_MESSAGE_TYPE
        }) {
            return Err(SlideMovieGeometryError::UnsupportedDependency);
        }
    }
    budget.references(
        ids.len()
            .checked_add(z_order.len())
            .ok_or(SlideMovieGeometryError::InvalidSource)?,
    )?;
    let mut movies = Vec::new();
    budget.allocations(ids.len())?;
    movies
        .try_reserve_exact(ids.len())
        .map_err(|_| SlideMovieGeometryError::Allocation { amount: ids.len() })?;
    let mut data_identifier_counts: HashMap<u64, usize> = HashMap::new();
    data_identifier_counts
        .try_reserve(ids.len())
        .map_err(|_| SlideMovieGeometryError::Allocation { amount: ids.len() })?;
    budget.allocations(ids.len())?;
    let mut audio_data_witnesses = Vec::new();
    if matches!(media_kind, GeometryMediaKind::Audio) {
        let witness_capacity = ids
            .len()
            .checked_mul(2)
            .ok_or(SlideMovieGeometryError::InvalidSource)?;
        budget.allocations(witness_capacity)?;
        let witness_bytes = witness_capacity
            .checked_mul(size_of::<(u64, u64)>())
            .ok_or(SlideMovieGeometryError::InvalidSource)?;
        budget.retained(witness_bytes)?;
        budget.scratch(witness_bytes)?;
        audio_data_witnesses
            .try_reserve_exact(witness_capacity)
            .map_err(|_| SlideMovieGeometryError::Allocation {
                amount: witness_capacity,
            })?;
    }
    for identifier in &ids {
        let (movie_component, object) = package
            .object_with_component(*identifier)
            .ok_or(SlideMovieGeometryError::InvalidSource)?;
        if !object
            .messages
            .iter()
            .any(|message| message.type_ == MOVIE_MESSAGE_TYPE)
        {
            continue;
        }
        if movie_component != component_name {
            return Err(SlideMovieGeometryError::UnsupportedDependency);
        }
        let (message_index, payload) = unique_message(object, MOVIE_MESSAGE_TYPE, budget)?;
        validate_movie_wire_framing(payload, limits, budget)?;
        let preflight = super::preflight_movie(
            payload,
            limits,
            super::SemanticPath::SlideDrawable {
                slide: slide_position.get(),
                index: movies.len(),
            },
        )
        .map_err(map_read_error)?;
        let (info, _) = super::decode_movie_info(
            payload,
            limits,
            super::SemanticPath::SlideDrawable {
                slide: slide_position.get(),
                index: movies.len(),
            },
        )
        .map_err(map_read_error)?;
        let kind = info.kind();
        if matches!(media_kind, GeometryMediaKind::FileMovie) && kind != MovieKind::File {
            continue;
        }
        if matches!(media_kind, GeometryMediaKind::Audio) && kind != MovieKind::Audio {
            // Audio selectors count every movie archive in slide source order.
            // Keep non-audio siblings in the typed inventory so selecting one
            // fails as a wrong-kind graph instead of silently renumbering the
            // audio list.
            movies.push((
                *identifier,
                message_index,
                kind,
                None,
                None,
                None,
                preflight.geometry_flags,
                preflight.geometry_angle,
                preflight.locked.unwrap_or(false),
            ));
            continue;
        }
        if preflight.display_size_fields > 0
            && (preflight.display_width.is_some_and(|width| width < 0.0)
                || preflight.display_height.is_some_and(|height| height < 0.0))
        {
            return Err(SlideMovieGeometryError::InvalidSource);
        }
        if preflight.movie_data_fields != 1 || preflight.data_references == 0 {
            return Err(SlideMovieGeometryError::UnsupportedDependency);
        }
        let data_identifiers = movie_data_references(payload, limits, budget)?;
        for data_identifier in data_identifiers {
            let count = data_identifier_counts.entry(data_identifier).or_insert(0);
            *count = (*count)
                .checked_add(1)
                .ok_or(SlideMovieGeometryError::InvalidSource)?;
            if matches!(media_kind, GeometryMediaKind::Audio) {
                if audio_data_witnesses
                    .iter()
                    .any(|(known_data, known_movie)| {
                        *known_data == data_identifier && *known_movie == *identifier
                    })
                {
                    return Err(SlideMovieGeometryError::InvalidSource);
                }
                audio_data_witnesses.push((data_identifier, *identifier));
            }
        }
        if movie_parent(payload, limits, budget)? != record.slide_identifier {
            return Err(SlideMovieGeometryError::InvalidSource);
        }
        let options = codec_options(package, payload, budget)?;
        let (geometry, before_position, transform, native_flags, native_angle) = match media_kind {
            GeometryMediaKind::FileMovie => {
                let (snapshot, report) =
                    keynote_movie_geometry_codec::decode_movie_geometry_with_report(
                        payload, options,
                    )
                    .map_err(map_geometry_codec_error)?;
                budget.codec_report(report)?;
                let geometry = MovieGeometry::new(
                    Point {
                        x: snapshot.x(),
                        y: snapshot.y(),
                    },
                    Size {
                        width: snapshot.width(),
                        height: snapshot.height(),
                    },
                )
                .map(Some)
                .map_err(|_| SlideMovieGeometryError::InvalidSource)?;
                let position = geometry
                    .ok_or(SlideMovieGeometryError::InvalidSource)?
                    .position();
                let angle = preflight.geometry_angle.unwrap_or(0.0);
                let reflected =
                    preflight.geometry_flags.unwrap_or_default() & MOVIE_REFLECTION_FLAG != 0;
                let transform = MovieTransform::new(angle, reflected)
                    .map_err(|_| SlideMovieGeometryError::InvalidSource)?;
                (
                    geometry,
                    position,
                    Some(transform),
                    preflight.geometry_flags,
                    preflight.geometry_angle,
                )
            },
            GeometryMediaKind::Audio => {
                let (snapshot, report) =
                    keynote_movie_geometry_codec::decode_movie_position_with_report(
                        payload, options,
                    )
                    .map_err(map_geometry_codec_error)?;
                budget.codec_report(report)?;
                let position = Point {
                    x: snapshot.x(),
                    y: snapshot.y(),
                };
                if !position.x.is_finite() || !position.y.is_finite() {
                    return Err(SlideMovieGeometryError::InvalidSource);
                }
                (None, position, None, None, None)
            },
        };
        movies.push((
            *identifier,
            message_index,
            kind,
            geometry,
            Some(before_position),
            transform,
            native_flags,
            native_angle,
            preflight.locked.unwrap_or(false),
        ));
    }
    let movie_position = movie_selector.as_position();
    let (
        movie_identifier,
        message_index,
        movie_kind,
        before,
        before_position,
        before_transform,
        native_flags,
        native_angle,
        locked,
    ) = *movies.get(movie_position.get()).ok_or(
        SlideMovieGeometryError::MoviePositionNotFound {
            position: movie_position,
        },
    )?;
    if matches!(media_kind, GeometryMediaKind::Audio) && movie_kind != MovieKind::Audio {
        return Err(SlideMovieGeometryError::UnsupportedDependency);
    }
    let before_position = before_position.ok_or(SlideMovieGeometryError::UnsupportedDependency)?;
    ensure_unique_movie_identity(package, component_name, movie_identifier, budget)?;
    ensure_unique_movie_owner(
        package,
        record.slide_identifier,
        movie_identifier,
        SLIDE_OWNED_DRAWABLES_FIELD,
        limits,
        budget,
    )?;
    ensure_unique_movie_owner(
        package,
        record.slide_identifier,
        movie_identifier,
        SLIDE_Z_ORDER_FIELD,
        limits,
        budget,
    )?;
    let (_, movie_object) = package
        .object_with_component(movie_identifier)
        .ok_or(SlideMovieGeometryError::InvalidSource)?;
    let (_, movie_payload) = unique_message(movie_object, MOVIE_MESSAGE_TYPE, budget)?;
    let movie_data_ids = movie_data_references(movie_payload, limits, budget)?;
    validate_movie_data_dependencies(
        movie_object,
        message_index,
        &movie_data_ids,
        &data_identifier_counts,
        budget,
    )?;
    let movie_refs = movie_archive_references(movie_payload, limits, budget)?;
    let mut movie_ref_ids = Vec::new();
    budget.allocations(movie_refs.len())?;
    movie_ref_ids
        .try_reserve_exact(movie_refs.len())
        .map_err(|_| SlideMovieGeometryError::Allocation {
            amount: movie_refs.len(),
        })?;
    for reference in &movie_refs {
        movie_ref_ids.push(reference.identifier);
    }
    validate_archive_info_references(
        movie_object,
        message_index,
        &movie_ref_ids,
        false,
        &[],
        matches!(media_kind, GeometryMediaKind::Audio),
        budget,
    )?;
    for reference in &movie_refs {
        let identifier = reference.identifier;
        package
            .object_with_component(identifier)
            .ok_or(SlideMovieGeometryError::InvalidSource)?;
        let (_, referenced_object) = package
            .object_with_component(identifier)
            .ok_or(SlideMovieGeometryError::InvalidSource)?;
        unique_movie_reference_message(
            referenced_object,
            reference.expected_message_type,
            media_kind,
            budget,
        )?;
        let owners = package
            .state
            .source
            .components()
            .iter()
            .flat_map(|component| component.archive().objects.iter())
            .filter(|object| object.archive_info.identifier == Some(identifier))
            .count();
        if owners != 1 {
            return Err(SlideMovieGeometryError::UnsupportedDependency);
        }
    }
    if before.is_none() && !matches!(media_kind, GeometryMediaKind::Audio) {
        return Err(SlideMovieGeometryError::UnsupportedDependency);
    }
    let metadata_capacity = if matches!(media_kind, GeometryMediaKind::Audio) {
        2
    } else {
        ids.len()
            .checked_add(z_order.len())
            .and_then(|value| value.checked_add(1))
            .ok_or(SlideMovieGeometryError::InvalidSource)?
    };
    budget.allocations(metadata_capacity)?;
    budget.retained(
        metadata_capacity
            .checked_mul(size_of::<u64>())
            .ok_or(SlideMovieGeometryError::InvalidSource)?,
    )?;
    let mut metadata_targets = Vec::new();
    metadata_targets
        .try_reserve_exact(metadata_capacity)
        .map_err(|_| SlideMovieGeometryError::Allocation {
            amount: metadata_capacity,
        })?;
    metadata_targets.push(record.slide_identifier);
    if matches!(media_kind, GeometryMediaKind::Audio) {
        if !metadata_targets.contains(&movie_identifier) {
            metadata_targets.push(movie_identifier);
        }
    } else {
        for identifier in ids.iter().chain(z_order.iter()) {
            if !metadata_targets.contains(identifier) {
                metadata_targets.push(*identifier);
            }
        }
    }
    let graph = MovieReferenceGraph {
        component_name,
        slide_identifier: record.slide_identifier,
        slide_message_index,
        movie_identifier,
        movie_message_index: message_index,
        data_ids: &movie_data_ids,
        audio_data_witnesses: &audio_data_witnesses,
        media_kind,
    };
    validate_movie_metadata(package, &graph, &metadata_targets, budget)?;
    validate_global_movie_references(package, &graph, budget)?;
    Ok(GeometrySelection {
        slide_position,
        movie_position,
        slide_identifier: record.slide_identifier,
        node_identifier: record.node_identifier,
        movie_identifier,
        message_index,
        slide_component_name: Arc::from(component_name),
        before,
        before_position: Some(before_position),
        before_transform,
        native_flags,
        native_angle,
        locked,
    })
}

fn rewrite_movie(
    source: &Package,
    selection: &GeometrySelection,
    after: MovieGeometry,
    after_transform: MovieTransform,
    budget: &mut GeometryBudget,
) -> Result<(Package, usize), SlideMovieGeometryError> {
    let geometry_changed = selection.before != Some(after);
    let transform_changed = selection.before_transform != Some(after_transform);
    rewrite_geometry_message(
        source,
        selection.slide_component_name.as_ref(),
        selection.movie_identifier,
        selection.message_index,
        |original, budget| {
            let mut rewritten = None;
            if geometry_changed {
                let options = codec_options(source, original, budget)?;
                let write = keynote_movie_geometry_codec::MovieGeometryWrite::from_values(
                    after.position().x,
                    after.position().y,
                    after.size().width,
                    after.size().height,
                );
                let prepared = keynote_movie_geometry_codec::prepare_movie_geometry_rewrite(
                    original, write, options,
                )
                .map_err(map_geometry_codec_error)?;
                budget.codec_report(prepared.prepare_report())?;
                let requirements = prepared.execution_requirements();
                budget.codec_requirements(requirements)?;
                rewritten = Some(
                    prepared
                        .execute(requirements.exact_limits())
                        .map_err(map_geometry_codec_error)?
                        .into_output(),
                );
            }
            if transform_changed {
                let transform_source = rewritten.as_deref().unwrap_or(original);
                let options = codec_options(source, transform_source, budget)?;
                let write = transform_write(selection, after_transform)?;
                let prepared = keynote_movie_geometry_codec::prepare_movie_transform_rewrite(
                    transform_source,
                    write,
                    options,
                )
                .map_err(map_geometry_codec_error)?;
                budget.codec_report(prepared.prepare_report())?;
                let requirements = prepared.execution_requirements();
                budget.codec_requirements(requirements)?;
                rewritten = Some(
                    prepared
                        .execute(requirements.exact_limits())
                        .map_err(map_geometry_codec_error)?
                        .into_output(),
                );
            }
            rewritten.ok_or(SlideMovieGeometryError::InvalidSource)
        },
        budget,
    )
}

/// Rewrite one selected movie/audio message while reusing the exact physical
/// archive transaction, preview invalidation, and operation accounting.
pub(crate) fn rewrite_geometry_message<F>(
    source: &Package,
    component_name: &str,
    object_identifier: u64,
    message_index: usize,
    rewrite: F,
    budget: &mut GeometryBudget,
) -> Result<(Package, usize), SlideMovieGeometryError>
where
    F: FnOnce(&[u8], &mut GeometryBudget) -> Result<Vec<u8>, SlideMovieGeometryError>,
{
    let catalog = physical_catalog(source)?;
    let entries = entry_index(catalog.package(), budget)?;
    let entry = entries
        .get(component_name)
        .copied()
        .ok_or(SlideMovieGeometryError::InvalidSource)?;
    let physical_limits = source.state.options.archive();
    let archive_limits = physical_limits
        .effective_archive_limits()
        .map_err(map_archive_error)?;
    let snappy_limits = physical_limits.snappy_limits().map_err(map_archive_error)?;
    budget.physical(entry.data().len())?;
    let stream = SnappyStream::decompress_with_limits(entry.data(), snappy_limits)
        .map_err(map_core_error)?;
    budget.physical(stream.as_bytes().len())?;
    budget.allocations(1)?;
    budget.retained(stream.as_bytes().len())?;
    budget.scratch(stream.as_bytes().len())?;
    let mut archive =
        Archive::parse_with_limits(stream.as_bytes(), archive_limits).map_err(map_core_error)?;
    charge_archive_inventory(&archive, budget)?;
    let object = archive
        .object(object_identifier)
        .ok_or(SlideMovieGeometryError::InvalidSource)?;
    let original = object
        .messages
        .get(message_index)
        .ok_or(SlideMovieGeometryError::InvalidSource)?
        .data
        .as_slice();
    let encoded_bound = archive
        .encoded_len_with_limits(archive_limits)
        .map_err(map_core_error)?
        .checked_add(original.len())
        .ok_or(SlideMovieGeometryError::InvalidSource)?;
    let compressed_bound =
        SnappyStream::maximum_compressed_len(encoded_bound).map_err(map_core_error)?;
    if compressed_bound > snappy_limits.max_compressed_stream() {
        return Err(SlideMovieGeometryError::LimitExceeded {
            kind: SlideMovieGeometryLimitKind::EntryBytes,
            observed: compressed_bound as u64,
            maximum: snappy_limits.max_compressed_stream() as u64,
        });
    }
    let package_bound = source
        .source_bytes()
        .len()
        .checked_sub(entry.data().len())
        .and_then(|value| value.checked_add(compressed_bound))
        .ok_or(SlideMovieGeometryError::InvalidSource)?;
    budget.preflight_output(
        encoded_bound
            .checked_add(compressed_bound)
            .ok_or(SlideMovieGeometryError::InvalidSource)?,
    )?;
    budget.preflight_output(package_bound)?;
    budget.preflight_work(
        encoded_bound
            .checked_add(compressed_bound)
            .ok_or(SlideMovieGeometryError::InvalidSource)?,
    )?;
    budget.allocations(1)?;
    budget.retained(encoded_bound)?;
    budget.scratch(encoded_bound)?;
    let rewritten = rewrite(original, budget)?;
    archive
        .object_mut(object_identifier)
        .ok_or(SlideMovieGeometryError::InvalidSource)?
        .replace_message_preserving_header_with_limits(
            message_index,
            RawMessage {
                type_: MOVIE_MESSAGE_TYPE,
                data: rewritten,
            },
            archive_limits,
        )
        .map_err(map_core_error)?;
    let bytes = archive
        .to_bytes_with_limits(archive_limits)
        .map_err(map_core_error)?;
    budget.physical(bytes.len())?;
    budget.work(compressed_bound)?;
    budget.allocations(1)?;
    budget.retained(compressed_bound)?;
    budget.scratch(compressed_bound)?;
    let compressed = SnappyStream::compress(&bytes).map_err(map_core_error)?;
    let previews = super::rendering_invalidation::root_preview_deletions(catalog.package())
        .map_err(|_| SlideMovieGeometryError::InvalidSource)?;
    let edit = EntryEdit::new(component_name, compressed.as_slice());
    let edits = [edit];
    let plan_units = previews
        .len()
        .checked_add(edits.len())
        .ok_or(SlideMovieGeometryError::InvalidSource)?;
    budget.allocations(1)?;
    budget.retained(
        plan_units
            .checked_mul(size_of::<&str>())
            .ok_or(SlideMovieGeometryError::InvalidSource)?,
    )?;
    budget.scratch(
        plan_units
            .checked_mul(size_of::<&str>())
            .ok_or(SlideMovieGeometryError::InvalidSource)?,
    )?;
    let prepared = catalog
        .prepare_reassembly_with_deletions(&edits, previews.names(), physical_limits)
        .map_err(map_archive_error)?;
    let req = prepared.execution_requirements();
    budget.reassembly(req)?;
    budget.candidate_reopen(req.output_bytes())?;
    let output = prepared
        .execute(req.exact_limits())
        .map_err(map_archive_error)?;
    let candidate = Package::from_source_with_options(output.into(), source.state.options)
        .map_err(map_read_error)?;
    Ok((candidate, previews.len()))
}

fn transform_write(
    selection: &GeometrySelection,
    after: MovieTransform,
) -> Result<keynote_movie_geometry_codec::MovieTransformWrite, SlideMovieGeometryError> {
    let before = selection
        .before_transform
        .ok_or(SlideMovieGeometryError::UnsupportedDependency)?;
    let flags = match selection.native_flags {
        Some(value) if (value & MOVIE_REFLECTION_FLAG != 0) == after.is_reflected() => {
            keynote_movie_geometry_codec::TransformField::Preserve
        },
        Some(value) => keynote_movie_geometry_codec::TransformField::Set(if after.is_reflected() {
            value | MOVIE_REFLECTION_FLAG
        } else {
            value & !MOVIE_REFLECTION_FLAG
        }),
        None if !after.is_reflected() => keynote_movie_geometry_codec::TransformField::Preserve,
        None => keynote_movie_geometry_codec::TransformField::Set(MOVIE_REFLECTION_FLAG),
    };
    let angle = if selection.native_angle == Some(after.angle_degrees())
        || (selection.native_angle.is_none() && after.angle_degrees() == 0.0)
    {
        keynote_movie_geometry_codec::TransformField::Preserve
    } else {
        keynote_movie_geometry_codec::TransformField::Set(after.angle_degrees())
    };
    if before.is_reflected() == after.is_reflected()
        && selection.native_flags.is_none()
        && !matches!(
            flags,
            keynote_movie_geometry_codec::TransformField::Preserve
        )
    {
        return Err(SlideMovieGeometryError::UnsupportedDependency);
    }
    Ok(keynote_movie_geometry_codec::MovieTransformWrite::with_updates(flags, angle))
}

pub(crate) fn codec_options(
    package: &Package,
    payload: &[u8],
    budget: &GeometryBudget,
) -> Result<keynote_movie_geometry_codec::DecodeOptions, SlideMovieGeometryError> {
    let limits = budget.residual(package)?;
    let recursion =
        u32::try_from(limits.max_nesting()).map_err(|_| SlideMovieGeometryError::InvalidSource)?;
    let output = budget.remaining_output()?.min(limits.max_output_bytes());
    let allocations = budget.remaining_allocations()?;
    let retained = budget.remaining_retained()?;
    let scratch = budget.remaining_scratch()?;
    Ok(keynote_movie_geometry_codec::DecodeOptions::new(
        limits.max_input_bytes().min(payload.len().max(1)),
        limits.max_fields(),
        limits.max_rewrite_work(),
        recursion,
    )
    .with_max_output_bytes(output)
    .with_max_allocations(allocations)
    .with_max_retained_bytes(retained)
    .with_max_scratch_bytes(scratch))
}

pub(crate) fn map_geometry_codec_error(
    error: keynote_movie_geometry_codec::DecodeError,
) -> SlideMovieGeometryError {
    if let Some(limit) = error.resource_limit() {
        let (kind, observed, maximum) = match limit {
            keynote_movie_geometry_codec::DecodeLimit::Bytes { observed, maximum } => {
                (SlideMovieGeometryLimitKind::WireBytes, observed, maximum)
            },
            keynote_movie_geometry_codec::DecodeLimit::Fields { observed, maximum } => {
                (SlideMovieGeometryLimitKind::WireFields, observed, maximum)
            },
            keynote_movie_geometry_codec::DecodeLimit::Work { observed, maximum } => {
                (SlideMovieGeometryLimitKind::WireWork, observed, maximum)
            },
            keynote_movie_geometry_codec::DecodeLimit::Nesting { observed, maximum } => (
                SlideMovieGeometryLimitKind::WireNesting,
                observed as usize,
                maximum as usize,
            ),
            keynote_movie_geometry_codec::DecodeLimit::Allocations { observed, maximum } => {
                (SlideMovieGeometryLimitKind::Allocations, observed, maximum)
            },
            keynote_movie_geometry_codec::DecodeLimit::Retained { observed, maximum } => {
                (SlideMovieGeometryLimitKind::Retained, observed, maximum)
            },
            keynote_movie_geometry_codec::DecodeLimit::Scratch { observed, maximum } => {
                (SlideMovieGeometryLimitKind::Scratch, observed, maximum)
            },
            keynote_movie_geometry_codec::DecodeLimit::Output { observed, maximum } => {
                (SlideMovieGeometryLimitKind::OutputBytes, observed, maximum)
            },
            _ => return SlideMovieGeometryError::InvalidSource,
        };
        return SlideMovieGeometryError::LimitExceeded {
            kind,
            observed: observed as u64,
            maximum: maximum as u64,
        };
    }
    SlideMovieGeometryError::InvalidSource
}

pub(crate) fn verify_locality(
    source: &Package,
    candidate: &Package,
    component_name: &str,
    object_identifier: u64,
    message_index: usize,
    target_previews_absent: bool,
    budget: &mut GeometryBudget,
) -> Result<(), SlideMovieGeometryError> {
    let source_catalog = physical_catalog(source)?;
    let candidate_catalog = physical_catalog(candidate)?;
    let previews = super::rendering_invalidation::root_preview_deletions(source_catalog.package())
        .map_err(|_| SlideMovieGeometryError::Verification)?;
    let candidate_previews =
        super::rendering_invalidation::root_preview_deletions(candidate_catalog.package())
            .map_err(|_| SlideMovieGeometryError::Verification)?;

    let source_entries = entry_index(source_catalog.package(), budget)?;
    let candidate_entries = entry_index(candidate_catalog.package(), budget)?;
    for entry in source_catalog.package().iter() {
        budget.work(
            entry
                .data()
                .len()
                .checked_add(entry.name().len())
                .ok_or(SlideMovieGeometryError::InvalidSource)?,
        )?;
        let candidate_entry = candidate_entries.get(entry.name()).copied();
        if target_previews_absent && previews.names().contains(&entry.name()) {
            if candidate_entry.is_some() {
                return Err(SlideMovieGeometryError::Verification);
            }
            continue;
        }
        let other = candidate_entry.ok_or(SlideMovieGeometryError::Verification)?;
        budget.work(other.data().len())?;
        if entry.name() != component_name && entry.data() != other.data() {
            return Err(SlideMovieGeometryError::Verification);
        }
    }
    for entry in candidate_catalog.package().iter() {
        if !source_entries.contains_key(entry.name()) {
            if target_previews_absent || !candidate_previews.names().contains(&entry.name()) {
                return Err(SlideMovieGeometryError::Verification);
            }
        }
    }

    let source_archive = component_archive(source, component_name, budget)?;
    let candidate_archive = component_archive(candidate, component_name, budget)?;
    if source_archive.objects.len() != candidate_archive.objects.len() {
        return Err(SlideMovieGeometryError::Verification);
    }
    let source_objects = archive_object_index(&source_archive, budget)?;
    let candidate_objects = archive_object_index(&candidate_archive, budget)?;
    let archive_limits = source
        .state
        .options
        .archive()
        .effective_archive_limits()
        .map_err(map_archive_error)?;
    for source_object in &source_archive.objects {
        let identifier = source_object
            .archive_info
            .identifier
            .ok_or(SlideMovieGeometryError::Verification)?;
        let source_object = source_objects
            .get(&identifier)
            .copied()
            .ok_or(SlideMovieGeometryError::Verification)?;
        let candidate_object = candidate_objects
            .get(&identifier)
            .copied()
            .ok_or(SlideMovieGeometryError::Verification)?;
        if identifier == object_identifier {
            let source_message = source_object
                .messages
                .get(message_index)
                .ok_or(SlideMovieGeometryError::Verification)?;
            let candidate_message = candidate_object
                .messages
                .get(message_index)
                .ok_or(SlideMovieGeometryError::Verification)?;
            if source_message.type_ != candidate_message.type_ {
                return Err(SlideMovieGeometryError::Verification);
            }
            let mut expected = source_object.clone();
            expected
                .replace_message_preserving_header_with_limits(
                    message_index,
                    candidate_message.clone(),
                    archive_limits,
                )
                .map_err(map_core_error)?;
            expected.header_length = candidate_object.header_length;
            expected.data_length = candidate_object.data_length;
            if !expected.same_content_ignoring_offsets(candidate_object) {
                return Err(SlideMovieGeometryError::Verification);
            }
        } else if !source_object.same_content_ignoring_offsets(candidate_object) {
            return Err(SlideMovieGeometryError::Verification);
        }
        budget.work(source_object.messages.len())?;
    }
    Ok(())
}

fn component_archive(
    package: &Package,
    name: &str,
    budget: &mut GeometryBudget,
) -> Result<Archive, SlideMovieGeometryError> {
    let catalog = physical_catalog(package)?;
    let entry = catalog
        .package()
        .iter()
        .find(|entry| entry.name() == name)
        .ok_or(SlideMovieGeometryError::InvalidSource)?;
    if entry.is_opaque() {
        return Err(SlideMovieGeometryError::InvalidSource);
    }
    let snappy = package
        .state
        .options
        .archive()
        .snappy_limits()
        .map_err(map_archive_error)?;
    let archive_limits = package
        .state
        .options
        .archive()
        .effective_archive_limits()
        .map_err(map_archive_error)?;
    budget.physical(entry.data().len())?;
    let stream =
        SnappyStream::decompress_with_limits(entry.data(), snappy).map_err(map_core_error)?;
    budget.physical(stream.as_bytes().len())?;
    let archive =
        Archive::parse_with_limits(stream.as_bytes(), archive_limits).map_err(map_core_error)?;
    charge_archive_inventory(&archive, budget)?;
    Ok(archive)
}

fn entry_index<'a>(
    catalog: &'a Catalog,
    budget: &mut GeometryBudget,
) -> Result<HashMap<&'a str, &'a Entry>, SlideMovieGeometryError> {
    let count = catalog.iter().count();
    budget.allocations(usize::from(count != 0))?;
    budget.retained(
        count
            .checked_mul(size_of::<(&str, &Entry)>())
            .ok_or(SlideMovieGeometryError::InvalidSource)?,
    )?;
    let mut index = HashMap::new();
    index
        .try_reserve(count)
        .map_err(|_| SlideMovieGeometryError::Allocation { amount: count })?;
    for entry in catalog.iter() {
        budget.work(
            entry
                .name()
                .len()
                .checked_add(1)
                .ok_or(SlideMovieGeometryError::InvalidSource)?,
        )?;
        if index.insert(entry.name(), entry).is_some() {
            return Err(SlideMovieGeometryError::Verification);
        }
    }
    Ok(index)
}

fn archive_object_index<'a>(
    archive: &'a Archive,
    budget: &mut GeometryBudget,
) -> Result<HashMap<u64, &'a ArchiveObject>, SlideMovieGeometryError> {
    let count = archive.objects.len();
    budget.allocations(usize::from(count != 0))?;
    budget.retained(
        count
            .checked_mul(size_of::<(u64, &ArchiveObject)>())
            .ok_or(SlideMovieGeometryError::InvalidSource)?,
    )?;
    let mut index = HashMap::new();
    index
        .try_reserve(count)
        .map_err(|_| SlideMovieGeometryError::Allocation { amount: count })?;
    for object in &archive.objects {
        let identifier = object
            .archive_info
            .identifier
            .ok_or(SlideMovieGeometryError::Verification)?;
        if index.insert(identifier, object).is_some() {
            return Err(SlideMovieGeometryError::Verification);
        }
    }
    Ok(index)
}

fn charge_archive_inventory(
    archive: &Archive,
    budget: &mut GeometryBudget,
) -> Result<(), SlideMovieGeometryError> {
    budget.work(archive.objects.len())?;
    budget.allocations(usize::from(!archive.objects.is_empty()))?;
    for object in &archive.objects {
        budget.work(object.messages.len())?;
        let mut fields = 0usize;
        let mut references = 0usize;
        for info in &object.archive_info.message_infos {
            fields = fields
                .checked_add(info.field_infos.len())
                .ok_or(SlideMovieGeometryError::InvalidSource)?;
            references = references
                .checked_add(info.object_references.len())
                .and_then(|value| value.checked_add(info.data_references.len()))
                .ok_or(SlideMovieGeometryError::InvalidSource)?;
            for field in &info.field_infos {
                references = references
                    .checked_add(field.object_references.len())
                    .and_then(|value| value.checked_add(field.data_references.len()))
                    .ok_or(SlideMovieGeometryError::InvalidSource)?;
            }
        }
        budget.fields(fields)?;
        budget.references(references)?;
        budget.work(
            fields
                .checked_add(references)
                .ok_or(SlideMovieGeometryError::InvalidSource)?,
        )?;
    }
    Ok(())
}

fn resolve_slide_position(
    package: &Package,
    selector: SlideSelector<'_>,
    budget: &mut GeometryBudget,
) -> Result<Position, SlideMovieGeometryError> {
    match selector {
        SlideSelector::Position(position) => Ok(position),
        SlideSelector::Name(name) => {
            if name.is_empty() {
                return Err(SlideMovieGeometryError::EmptySlideName);
            }
            budget.work(package.source_bytes().len())?;
            package
                .show()
                .map_err(map_read_error)?
                .select_slide(selector)
                .map_err(map_slide_selector_error)?
                .map(|slide| Position::new(slide.index()))
                .ok_or(SlideMovieGeometryError::SlideNameNotFound)
        },
    }
}

fn same_selection(left: &GeometrySelection, right: &GeometrySelection) -> bool {
    left.slide_position == right.slide_position
        && left.movie_position == right.movie_position
        && left.slide_identifier == right.slide_identifier
        && left.node_identifier == right.node_identifier
        && left.movie_identifier == right.movie_identifier
        && left.message_index == right.message_index
        && left.slide_component_name == right.slide_component_name
        && left.locked == right.locked
}

fn unique_message<'a>(
    object: &'a ArchiveObject,
    message_type: u32,
    budget: &mut GeometryBudget,
) -> Result<(usize, &'a [u8]), SlideMovieGeometryError> {
    if object.messages.len() != object.archive_info.message_infos.len() {
        return Err(SlideMovieGeometryError::InvalidSource);
    }
    let mut selected = None;
    for (index, message) in object.messages.iter().enumerate() {
        budget.work(
            message
                .data
                .len()
                .checked_add(1)
                .ok_or(SlideMovieGeometryError::InvalidSource)?,
        )?;
        let info = object
            .archive_info
            .message_infos
            .get(index)
            .ok_or(SlideMovieGeometryError::InvalidSource)?;
        if message.type_ != info.type_
            || usize::try_from(info.length).ok() != Some(message.data.len())
        {
            return Err(SlideMovieGeometryError::InvalidSource);
        }
        if is_known_movie_role(message.type_) && message.type_ != message_type {
            return Err(SlideMovieGeometryError::UnsupportedDependency);
        }
        if message.type_ == message_type
            && selected.replace((index, message.data.as_slice())).is_some()
        {
            return Err(SlideMovieGeometryError::InvalidSource);
        }
    }
    selected.ok_or(SlideMovieGeometryError::InvalidSource)
}

fn unique_movie_reference_message<'a>(
    object: &'a ArchiveObject,
    message_type: u32,
    media_kind: GeometryMediaKind,
    budget: &mut GeometryBudget,
) -> Result<(usize, &'a [u8]), SlideMovieGeometryError> {
    if !matches!(media_kind, GeometryMediaKind::Audio) || message_type != MOVIE_STYLE_MESSAGE_TYPE {
        return unique_message(object, message_type, budget);
    }

    if object.messages.len() != object.archive_info.message_infos.len() {
        return Err(SlideMovieGeometryError::InvalidSource);
    }

    let mut selected = None;
    for (index, message) in object.messages.iter().enumerate() {
        budget.work(
            message
                .data
                .len()
                .checked_add(1)
                .ok_or(SlideMovieGeometryError::InvalidSource)?,
        )?;
        let info = object
            .archive_info
            .message_infos
            .get(index)
            .ok_or(SlideMovieGeometryError::InvalidSource)?;
        if message.type_ != info.type_
            || usize::try_from(info.length).ok() != Some(message.data.len())
        {
            return Err(SlideMovieGeometryError::InvalidSource);
        }

        let is_style = matches!(
            message.type_,
            MOVIE_STYLE_MESSAGE_TYPE | MOVIE_AUDIO_STYLE_MESSAGE_TYPE
        );
        if is_known_movie_role(message.type_) && !is_style {
            return Err(SlideMovieGeometryError::UnsupportedDependency);
        }
        if is_style && selected.replace((index, message.data.as_slice())).is_some() {
            return Err(SlideMovieGeometryError::InvalidSource);
        }
    }

    selected.ok_or(SlideMovieGeometryError::InvalidSource)
}

fn validate_drawable_message_headers(
    object: &ArchiveObject,
    budget: &mut GeometryBudget,
) -> Result<(), SlideMovieGeometryError> {
    if object.messages.len() != object.archive_info.message_infos.len() {
        return Err(SlideMovieGeometryError::InvalidSource);
    }
    for (index, message) in object.messages.iter().enumerate() {
        budget.work(
            message
                .data
                .len()
                .checked_add(1)
                .ok_or(SlideMovieGeometryError::InvalidSource)?,
        )?;
        let info = object
            .archive_info
            .message_infos
            .get(index)
            .ok_or(SlideMovieGeometryError::InvalidSource)?;
        if message.type_ != info.type_
            || usize::try_from(info.length).ok() != Some(message.data.len())
        {
            return Err(SlideMovieGeometryError::InvalidSource);
        }
    }
    Ok(())
}

fn is_known_movie_role(message_type: u32) -> bool {
    matches!(
        message_type,
        SLIDE_MESSAGE_TYPE
            | SLIDE_NODE_MESSAGE_TYPE
            | MOVIE_MESSAGE_TYPE
            | MOVIE_STANDIN_MESSAGE_TYPE
            | MOVIE_STYLE_MESSAGE_TYPE
            | TABLE_INFO_MESSAGE_TYPE
            | TABLE_MODEL_MESSAGE_TYPE
            | HEADER_BUCKET_MESSAGE_TYPE
            | TABLE_STYLE_MESSAGE_TYPE
            | TABLE_STYLE_PRESET_MESSAGE_TYPE
            | TABLE_STYLE_NETWORK_MESSAGE_TYPE
            | STYLESHEET_MESSAGE_TYPE
    )
}

fn repeated_references(
    payload: &[u8],
    field_number: u32,
    limits: litchi_iwa_common::WireLimits,
    budget: &mut GeometryBudget,
) -> Result<Vec<u64>, SlideMovieGeometryError> {
    let fields = WireView::parse_with_limits(payload, limits).map_err(map_wire_error)?;
    let count = fields
        .fields()
        .filter(|field| field.number() == field_number)
        .count();
    budget.fields(count)?;
    budget.work(payload.len())?;
    let mut result = Vec::new();
    budget.allocations(usize::from(count != 0))?;
    result
        .try_reserve_exact(count)
        .map_err(|_| SlideMovieGeometryError::Allocation { amount: count })?;
    for field in fields
        .fields()
        .filter(|field| field.number() == field_number)
    {
        if field.wire_type() != 2 {
            return Err(SlideMovieGeometryError::InvalidSource);
        }
        result.push(strict_reference_payload(field.payload(), limits)?);
    }
    Ok(result)
}

fn movie_parent(
    payload: &[u8],
    limits: litchi_iwa_common::WireLimits,
    budget: &mut GeometryBudget,
) -> Result<u64, SlideMovieGeometryError> {
    let fields = WireView::parse_with_limits(payload, limits).map_err(map_wire_error)?;
    budget.work(payload.len())?;
    let mut super_fields = fields
        .fields()
        .filter(|field| field.number() == MOVIE_SUPER_FIELD);
    let super_field = super_fields.next();
    if super_field.is_none() || super_fields.next().is_some() {
        return Err(SlideMovieGeometryError::InvalidSource);
    }
    let super_field = super_field.ok_or(SlideMovieGeometryError::InvalidSource)?;
    if super_field.wire_type() != 2 {
        return Err(SlideMovieGeometryError::InvalidSource);
    }
    super_field
        .validate_canonical_framing()
        .map_err(map_wire_error)?;
    let nested =
        WireView::parse_with_limits(super_field.payload(), limits).map_err(map_wire_error)?;
    let mut parent = nested
        .fields()
        .filter(|field| field.number() == DRAWABLE_PARENT_FIELD);
    let parent_field = parent.next();
    if parent_field.is_none() || parent.next().is_some() {
        return Err(SlideMovieGeometryError::InvalidSource);
    }
    let parent_field = parent_field.ok_or(SlideMovieGeometryError::InvalidSource)?;
    if parent_field.wire_type() != 2 {
        return Err(SlideMovieGeometryError::InvalidSource);
    }
    parent_field
        .validate_canonical_framing()
        .map_err(map_wire_error)?;
    strict_reference_payload(parent_field.payload(), limits)
}

fn strict_reference_payload(
    payload: &[u8],
    limits: litchi_iwa_common::WireLimits,
) -> Result<u64, SlideMovieGeometryError> {
    let fields = WireView::parse_with_limits(payload, limits).map_err(map_wire_error)?;
    let mut identifier = None;
    for field in fields.fields() {
        field.validate_canonical_framing().map_err(map_wire_error)?;
        match field.number() {
            1 => {
                if field.wire_type() != 0 || identifier.is_some() {
                    return Err(SlideMovieGeometryError::InvalidSource);
                }
                let (value, width) = decode_varint_from_bytes(field.payload())
                    .map_err(|_| SlideMovieGeometryError::InvalidSource)?;
                if width != encoded_len(value) || value == 0 {
                    return Err(SlideMovieGeometryError::InvalidSource);
                }
                identifier = Some(value);
            },
            2 | 3 => return Err(SlideMovieGeometryError::InvalidSource),
            _ => {},
        }
    }
    identifier.ok_or(SlideMovieGeometryError::InvalidSource)
}

fn validate_movie_wire_framing(
    payload: &[u8],
    limits: litchi_iwa_common::WireLimits,
    budget: &mut GeometryBudget,
) -> Result<(), SlideMovieGeometryError> {
    let root = WireView::parse_with_limits(payload, limits).map_err(map_wire_error)?;
    budget.work(payload.len())?;
    for field in root.fields() {
        if matches!(field.number(), 1 | 14 | 15 | 20 | 21) {
            field.validate_canonical_framing().map_err(map_wire_error)?;
        }
        if field.number() != 1 || field.wire_type() != 2 {
            continue;
        }
        let drawable =
            WireView::parse_with_limits(field.payload(), limits).map_err(map_wire_error)?;
        for drawable_field in drawable.fields() {
            if drawable_field.number() != 1 || drawable_field.wire_type() != 2 {
                continue;
            }
            drawable_field
                .validate_canonical_framing()
                .map_err(map_wire_error)?;
            let geometry = WireView::parse_with_limits(drawable_field.payload(), limits)
                .map_err(map_wire_error)?;
            for geometry_field in geometry.fields() {
                if matches!(geometry_field.number(), 1 | 2) {
                    geometry_field
                        .validate_canonical_framing()
                        .map_err(map_wire_error)?;
                    if geometry_field.wire_type() == 2 {
                        let coordinates =
                            WireView::parse_with_limits(geometry_field.payload(), limits)
                                .map_err(map_wire_error)?;
                        for coordinate in coordinates.fields() {
                            coordinate
                                .validate_canonical_framing()
                                .map_err(map_wire_error)?;
                        }
                    }
                } else if matches!(geometry_field.number(), 3 | 4) {
                    geometry_field
                        .validate_canonical_framing()
                        .map_err(map_wire_error)?;
                }
            }
        }
    }
    Ok(())
}

fn validate_archive_info_references(
    object: &ArchiveObject,
    message_index: usize,
    expected: &[u64],
    allow_unselected: bool,
    allowed_field_paths: &[u32],
    order_insensitive: bool,
    budget: &mut GeometryBudget,
) -> Result<(), SlideMovieGeometryError> {
    let info = object
        .archive_info
        .message_infos
        .get(message_index)
        .ok_or(SlideMovieGeometryError::InvalidSource)?;
    let field_count = info.field_infos.len();
    let reference_count = info
        .object_references
        .len()
        .checked_add(info.data_references.len())
        .and_then(|value| {
            info.field_infos.iter().try_fold(value, |total, field| {
                total
                    .checked_add(field.object_references.len())
                    .and_then(|value| value.checked_add(field.data_references.len()))
            })
        })
        .ok_or(SlideMovieGeometryError::InvalidSource)?;
    budget.fields(field_count)?;
    budget.references(reference_count)?;
    budget.work(
        field_count
            .checked_add(reference_count)
            .and_then(|value| value.checked_add(expected.len()))
            .ok_or(SlideMovieGeometryError::InvalidSource)?,
    )?;

    let mut info_seen = HashSet::new();
    budget.allocations(usize::from(!info.object_references.is_empty()))?;
    info_seen
        .try_reserve(info.object_references.len())
        .map_err(|_| SlideMovieGeometryError::Allocation {
            amount: info.object_references.len(),
        })?;
    for identifier in &info.object_references {
        if !info_seen.insert(*identifier) {
            return Err(SlideMovieGeometryError::InvalidSource);
        }
    }

    let mut expected_set = HashSet::new();
    budget.allocations(usize::from(!expected.is_empty()))?;
    expected_set
        .try_reserve(expected.len())
        .map_err(|_| SlideMovieGeometryError::Allocation {
            amount: expected.len(),
        })?;
    for identifier in expected {
        expected_set.insert(*identifier);
    }
    if expected_set
        .iter()
        .any(|identifier| !info_seen.contains(identifier))
    {
        return Err(SlideMovieGeometryError::InvalidSource);
    }
    if !allow_unselected {
        let exact_order = info.object_references.as_slice() == expected;
        let exact_set = order_insensitive
            && info.object_references.len() == expected.len()
            && expected_set.len() == expected.len()
            && info
                .object_references
                .iter()
                .all(|identifier| expected_set.contains(identifier));
        if !exact_order && !exact_set {
            return Err(SlideMovieGeometryError::InvalidSource);
        }
    }

    let max_field_references = info
        .field_infos
        .iter()
        .map(|field| field.object_references.len())
        .max()
        .unwrap_or(0);
    let mut field_seen = HashSet::new();
    budget.allocations(usize::from(max_field_references != 0))?;
    field_seen.try_reserve(max_field_references).map_err(|_| {
        SlideMovieGeometryError::Allocation {
            amount: max_field_references,
        }
    })?;
    for field in &info.field_infos {
        field_seen.clear();
        if !allow_unselected
            && (!field.object_references.is_empty() || !field.data_references.is_empty())
        {
            return Err(SlideMovieGeometryError::InvalidSource);
        }
        for identifier in &field.object_references {
            if !field_seen.insert(*identifier)
                || !info_seen.contains(identifier)
                || (expected_set.contains(identifier)
                    && (!allow_unselected
                        || !allowed_field_paths
                            .iter()
                            .any(|path| field.path.as_slice() == [*path])))
            {
                return Err(SlideMovieGeometryError::InvalidSource);
            }
        }
        if field
            .data_references
            .iter()
            .any(|identifier| expected_set.contains(identifier))
        {
            return Err(SlideMovieGeometryError::InvalidSource);
        }
    }
    Ok(())
}

fn validate_slide_archive_info_references(
    object: &ArchiveObject,
    message_index: usize,
    owned: &[u64],
    z_order: &[u64],
    budget: &mut GeometryBudget,
) -> Result<(), SlideMovieGeometryError> {
    let expected_len = owned
        .len()
        .checked_add(z_order.len())
        .ok_or(SlideMovieGeometryError::InvalidSource)?;
    budget.allocations(expected_len)?;
    budget.work(expected_len)?;
    let mut expected = Vec::new();
    expected
        .try_reserve_exact(expected_len)
        .map_err(|_| SlideMovieGeometryError::Allocation {
            amount: expected_len,
        })?;
    expected.extend_from_slice(owned);
    expected.extend_from_slice(z_order);

    let mut unique = HashSet::new();
    budget.allocations(expected_len)?;
    unique
        .try_reserve(expected_len)
        .map_err(|_| SlideMovieGeometryError::Allocation {
            amount: expected_len,
        })?;
    for value in owned {
        if !unique.insert(*value) {
            return Err(SlideMovieGeometryError::InvalidSource);
        }
    }
    unique.clear();
    for value in z_order {
        if !unique.insert(*value) {
            return Err(SlideMovieGeometryError::InvalidSource);
        }
    }
    validate_archive_info_references(
        object,
        message_index,
        &expected,
        true,
        &[SLIDE_OWNED_DRAWABLES_FIELD, SLIDE_Z_ORDER_FIELD],
        false,
        budget,
    )?;
    let info = object
        .archive_info
        .message_infos
        .get(message_index)
        .ok_or(SlideMovieGeometryError::InvalidSource)?;
    for field in &info.field_infos {
        if field.path.as_slice() == [SLIDE_OWNED_DRAWABLES_FIELD]
            && field.object_references.as_slice() != owned
        {
            return Err(SlideMovieGeometryError::InvalidSource);
        }
        if field.path.as_slice() == [SLIDE_Z_ORDER_FIELD]
            && field.object_references.as_slice() != z_order
        {
            return Err(SlideMovieGeometryError::InvalidSource);
        }
    }
    Ok(())
}

fn movie_archive_references(
    payload: &[u8],
    limits: litchi_iwa_common::WireLimits,
    budget: &mut GeometryBudget,
) -> Result<Vec<MovieArchiveReference>, SlideMovieGeometryError> {
    let fields = WireView::parse_with_limits(payload, limits).map_err(map_wire_error)?;
    budget.work(payload.len())?;
    budget.allocations(1)?;
    let mut super_fields = fields
        .fields()
        .filter(|field| field.number() == MOVIE_SUPER_FIELD);
    let super_field = super_fields.next();
    if super_field.is_none() || super_fields.next().is_some() {
        return Err(SlideMovieGeometryError::InvalidSource);
    }
    let super_field = super_field.ok_or(SlideMovieGeometryError::InvalidSource)?;
    if super_field.wire_type() != 2 {
        return Err(SlideMovieGeometryError::InvalidSource);
    }
    super_field
        .validate_canonical_framing()
        .map_err(map_wire_error)?;
    let drawable =
        WireView::parse_with_limits(super_field.payload(), limits).map_err(map_wire_error)?;
    let mut references = Vec::new();
    references
        .try_reserve_exact(3)
        .map_err(|_| SlideMovieGeometryError::Allocation { amount: 3 })?;
    let mut title_seen = false;
    let mut caption_seen = false;
    for field in drawable
        .fields()
        .filter(|field| matches!(field.number(), 10 | 11))
    {
        field.validate_canonical_key().map_err(map_wire_error)?;
        if field.wire_type() != 2 {
            return Err(SlideMovieGeometryError::InvalidSource);
        }
        field.validate_canonical_framing().map_err(map_wire_error)?;
        let seen = if field.number() == 10 {
            &mut title_seen
        } else {
            &mut caption_seen
        };
        if std::mem::replace(seen, true) {
            return Err(SlideMovieGeometryError::InvalidSource);
        }
        references.push(MovieArchiveReference {
            identifier: strict_reference_payload(field.payload(), limits)?,
            expected_message_type: MOVIE_STANDIN_MESSAGE_TYPE,
        });
    }
    let mut style_seen = false;
    for field in fields.fields().filter(|field| field.number() == 19) {
        field.validate_canonical_key().map_err(map_wire_error)?;
        if field.wire_type() != 2 {
            return Err(SlideMovieGeometryError::InvalidSource);
        }
        field.validate_canonical_framing().map_err(map_wire_error)?;
        if std::mem::replace(&mut style_seen, true) {
            return Err(SlideMovieGeometryError::InvalidSource);
        }
        references.push(MovieArchiveReference {
            identifier: strict_reference_payload(field.payload(), limits)?,
            expected_message_type: MOVIE_STYLE_MESSAGE_TYPE,
        });
    }
    if references
        .windows(2)
        .any(|pair| pair[0].identifier == pair[1].identifier)
    {
        return Err(SlideMovieGeometryError::InvalidSource);
    }
    Ok(references)
}

fn movie_data_references(
    payload: &[u8],
    limits: litchi_iwa_common::WireLimits,
    budget: &mut GeometryBudget,
) -> Result<Vec<u64>, SlideMovieGeometryError> {
    let fields = WireView::parse_with_limits(payload, limits).map_err(map_wire_error)?;
    let mut result = Vec::new();
    budget.allocations(1)?;
    budget.work(payload.len())?;
    result
        .try_reserve_exact(2)
        .map_err(|_| SlideMovieGeometryError::Allocation { amount: 2 })?;
    let mut movie_data_seen = false;
    let mut poster_data_seen = false;
    for field in fields
        .fields()
        .filter(|field| matches!(field.number(), 14 | 15))
    {
        field.validate_canonical_framing().map_err(map_wire_error)?;
        if field.wire_type() != 2 {
            return Err(SlideMovieGeometryError::InvalidSource);
        }
        let seen = if field.number() == 14 {
            &mut movie_data_seen
        } else {
            &mut poster_data_seen
        };
        if std::mem::replace(seen, true) {
            return Err(SlideMovieGeometryError::InvalidSource);
        }
        result.push(strict_reference_payload(field.payload(), limits)?);
    }
    if !movie_data_seen || result.windows(2).any(|pair| pair[0] == pair[1]) {
        return Err(SlideMovieGeometryError::UnsupportedDependency);
    }
    Ok(result)
}

fn validate_movie_data_dependencies(
    object: &ArchiveObject,
    message_index: usize,
    expected: &[u64],
    counts: &HashMap<u64, usize>,
    budget: &mut GeometryBudget,
) -> Result<(), SlideMovieGeometryError> {
    let info = object
        .archive_info
        .message_infos
        .get(message_index)
        .ok_or(SlideMovieGeometryError::InvalidSource)?;
    let mut seen = HashSet::new();
    budget.allocations(expected.len())?;
    budget.work(expected.len())?;
    seen.try_reserve(expected.len())
        .map_err(|_| SlideMovieGeometryError::Allocation {
            amount: expected.len(),
        })?;
    for identifier in expected {
        if !seen.insert(*identifier) {
            return Err(SlideMovieGeometryError::InvalidSource);
        }
        let occurrences = counts.get(identifier).copied().unwrap_or_default();
        if occurrences == 0 || (occurrences == 1 && !info.data_references.contains(identifier)) {
            return Err(SlideMovieGeometryError::UnsupportedDependency);
        }
    }
    if !info.data_references.is_empty() {
        if info
            .data_references
            .iter()
            .enumerate()
            .any(|(index, identifier)| info.data_references[..index].contains(identifier))
            || info.data_references.as_slice() != expected
        {
            return Err(SlideMovieGeometryError::InvalidSource);
        }
    }
    Ok(())
}

#[derive(Clone, Copy)]
struct MovieReferenceGraph<'a> {
    component_name: &'a str,
    slide_identifier: u64,
    slide_message_index: usize,
    movie_identifier: u64,
    movie_message_index: usize,
    data_ids: &'a [u64],
    audio_data_witnesses: &'a [(u64, u64)],
    media_kind: GeometryMediaKind,
}

fn validate_movie_metadata(
    package: &Package,
    graph: &MovieReferenceGraph<'_>,
    target_ids: &[u64],
    budget: &mut GeometryBudget,
) -> Result<(), SlideMovieGeometryError> {
    let MovieReferenceGraph {
        component_name,
        data_ids,
        audio_data_witnesses,
        media_kind,
        movie_identifier: selected_movie_identifier,
        ..
    } = *graph;
    let metadata = metadata_payload(package, budget)?;
    if matches!(media_kind, GeometryMediaKind::Audio) {
        // A metadata record consumes at least a key and a value byte. Bound
        // callback storage and witness lookups before the visitor allocates;
        // the maps themselves grow fallibly only as records are encountered.
        let units = (metadata.len() / 2)
            .checked_add(
                target_ids
                    .len()
                    .checked_mul(3)
                    .ok_or(SlideMovieGeometryError::InvalidSource)?,
            )
            .and_then(|value| value.checked_add(data_ids.len().checked_mul(2)?))
            .and_then(|value| value.checked_add(audio_data_witnesses.len()))
            .ok_or(SlideMovieGeometryError::InvalidSource)?;
        let bytes = units
            .checked_mul(size_of::<(u64, u64, usize)>())
            .ok_or(SlideMovieGeometryError::InvalidSource)?;
        let lookup_width = audio_data_witnesses
            .len()
            .checked_add(target_ids.len())
            .and_then(|value| value.checked_add(data_ids.len()))
            .and_then(|value| value.checked_add(2))
            .ok_or(SlideMovieGeometryError::InvalidSource)?;
        budget.allocations(units)?;
        budget.retained(bytes)?;
        budget.scratch(bytes)?;
        budget.work(
            metadata
                .len()
                .checked_mul(lookup_width)
                .ok_or(SlideMovieGeometryError::InvalidSource)?,
        )?;
    }
    let limits = budget.residual(package)?;
    let components = WireView::parse_with_limits(metadata, limits)
        .map_err(map_wire_error)?
        .fields()
        .filter(|field| matches!(field.number(), 3 | 11))
        .count()
        .checked_mul(2)
        .ok_or(SlideMovieGeometryError::InvalidSource)?;
    let references = budget.remaining_references()?;
    let options = package_metadata_codec::RewriteOptions::new(
        metadata.len().min(limits.max_input_bytes()),
        budget.remaining_output()?.min(limits.max_output_bytes()),
        limits.max_fields(),
        limits.max_rewrite_work(),
        u32::try_from(limits.max_nesting()).unwrap_or(u32::MAX),
        components,
        references,
        references,
    );
    budget.allocations(target_ids.len())?;
    if matches!(media_kind, GeometryMediaKind::Audio) {
        budget.allocations(audio_data_witnesses.len())?;
    }
    let mut visitor = MovieMetadataAuthorityVisitor::new(
        target_ids,
        component_name,
        data_ids,
        audio_data_witnesses,
        matches!(media_kind, GeometryMediaKind::Audio),
        selected_movie_identifier,
    )?;
    let inspection = package_metadata_codec::inspect_package_metadata_with_visitor(
        metadata,
        options,
        &mut visitor,
    )
    .map_err(map_metadata_error)?;
    budget.metadata_report(inspection.report())?;
    let audio_data_owners_valid = !matches!(media_kind, GeometryMediaKind::Audio)
        || data_ids.iter().all(|identifier| {
            let parent_count = visitor
                .audio_data_parent_counts
                .get(identifier)
                .copied()
                .unwrap_or(0);
            parent_count == 0 || visitor.audio_data_owner_counts.get(identifier).copied() == Some(1)
        });
    if visitor.invalid
        || visitor.target_counts.iter().any(|count| *count != 1)
        || !audio_data_owners_valid
    {
        return Err(SlideMovieGeometryError::UnsupportedDependency);
    }
    Ok(())
}

fn metadata_payload<'a>(
    package: &'a Package,
    budget: &mut GeometryBudget,
) -> Result<&'a [u8], SlideMovieGeometryError> {
    let mut payload = None;
    for component in package.state.source.components().iter() {
        for object in &component.archive().objects {
            if object.messages.len() != object.archive_info.message_infos.len() {
                return Err(SlideMovieGeometryError::InvalidSource);
            }
            for (index, message) in object.messages.iter().enumerate() {
                budget.work(
                    message
                        .data
                        .len()
                        .checked_add(1)
                        .ok_or(SlideMovieGeometryError::InvalidSource)?,
                )?;
                let info = object
                    .archive_info
                    .message_infos
                    .get(index)
                    .ok_or(SlideMovieGeometryError::InvalidSource)?;
                if message.type_ != info.type_
                    || usize::try_from(info.length).ok() != Some(message.data.len())
                {
                    return Err(SlideMovieGeometryError::InvalidSource);
                }
                if message.type_ == PACKAGE_METADATA_MESSAGE_TYPE
                    && payload.replace(message.data.as_slice()).is_some()
                {
                    return Err(SlideMovieGeometryError::InvalidSource);
                }
            }
        }
    }
    payload.ok_or(SlideMovieGeometryError::InvalidSource)
}

fn metadata_component_matches_physical(metadata: &str, physical: &str) -> bool {
    fn basename(name: &str) -> &str {
        name.rsplit('/').next().unwrap_or(name)
    }
    fn without_iwa(name: &str) -> &str {
        name.strip_suffix(".iwa").unwrap_or(name)
    }
    without_iwa(basename(metadata)) == without_iwa(basename(physical))
}

struct MovieMetadataAuthorityVisitor<'a> {
    target_ids: &'a [u64],
    component_name: &'a str,
    data_ids: &'a [u64],
    audio_data_witnesses: &'a [(u64, u64)],
    target_counts: Vec<usize>,
    target_pairs: Vec<Option<(u64, u64)>>,
    seen_pairs: HashMap<(u64, u64), u64>,
    allow_audio_data_owners: bool,
    selected_movie_identifier: u64,
    audio_data_parent_counts: HashMap<u64, usize>,
    audio_data_owner_counts: HashMap<u64, usize>,
    audio_data_explicit_owners: HashSet<(u64, u64)>,
    invalid: bool,
}

impl<'a> MovieMetadataAuthorityVisitor<'a> {
    fn new(
        target_ids: &'a [u64],
        component_name: &'a str,
        data_ids: &'a [u64],
        audio_data_witnesses: &'a [(u64, u64)],
        allow_audio_data_owners: bool,
        selected_movie_identifier: u64,
    ) -> Result<Self, SlideMovieGeometryError> {
        let mut target_counts = Vec::new();
        target_counts
            .try_reserve_exact(target_ids.len())
            .map_err(|_| SlideMovieGeometryError::Allocation {
                amount: target_ids.len(),
            })?;
        target_counts.resize(target_ids.len(), 0);
        let mut target_pairs = Vec::new();
        target_pairs
            .try_reserve_exact(target_ids.len())
            .map_err(|_| SlideMovieGeometryError::Allocation {
                amount: target_ids.len(),
            })?;
        target_pairs.resize(target_ids.len(), None);
        let mut seen_pairs = HashMap::new();
        seen_pairs.try_reserve(target_ids.len()).map_err(|_| {
            SlideMovieGeometryError::Allocation {
                amount: target_ids.len(),
            }
        })?;
        let mut audio_data_owner_counts = HashMap::new();
        let mut audio_data_parent_counts = HashMap::new();
        let mut audio_data_explicit_owners = HashSet::new();
        if allow_audio_data_owners {
            audio_data_parent_counts
                .try_reserve(data_ids.len())
                .map_err(|_| SlideMovieGeometryError::Allocation {
                    amount: data_ids.len(),
                })?;
            audio_data_owner_counts
                .try_reserve(data_ids.len())
                .map_err(|_| SlideMovieGeometryError::Allocation {
                    amount: data_ids.len(),
                })?;
            audio_data_explicit_owners
                .try_reserve(audio_data_witnesses.len())
                .map_err(|_| SlideMovieGeometryError::Allocation {
                    amount: audio_data_witnesses.len(),
                })?;
        }
        Ok(Self {
            target_ids,
            component_name,
            data_ids,
            audio_data_witnesses,
            target_counts,
            target_pairs,
            seen_pairs,
            allow_audio_data_owners,
            selected_movie_identifier,
            audio_data_parent_counts,
            audio_data_owner_counts,
            audio_data_explicit_owners,
            invalid: false,
        })
    }

    fn target_index(&self, identifier: u64) -> Option<usize> {
        self.target_ids
            .iter()
            .position(|target| *target == identifier)
    }

    fn expected_audio_owner_count(&self, data_identifier: u64) -> usize {
        self.audio_data_witnesses
            .iter()
            .filter(|(identifier, _)| *identifier == data_identifier)
            .count()
    }

    fn expected_audio_owner(&self, data_identifier: u64, object_identifier: u64) -> bool {
        self.audio_data_witnesses.iter().any(|(identifier, owner)| {
            *identifier == data_identifier && *owner == object_identifier
        })
    }
}

impl package_metadata_codec::PackageMetadataVisitor for MovieMetadataAuthorityVisitor<'_> {
    fn visit_object_uuid(
        &mut self,
        binding: package_metadata_codec::ObjectUuidDescriptor<'_>,
    ) -> Result<(), package_metadata_codec::RewriteError> {
        let object_identifier = binding.object_identifier();
        let uuid = binding.uuid();
        let pair = (uuid.lower(), uuid.upper());
        if !self.seen_pairs.contains_key(&pair) {
            self.seen_pairs
                .try_reserve(1)
                .map_err(|_| package_metadata_codec::RewriteError::allocation(1))?;
        }
        if let Some(previous) = self.seen_pairs.insert(pair, object_identifier) {
            if previous != object_identifier
                && (self.target_index(previous).is_some()
                    || self.target_index(object_identifier).is_some())
            {
                self.invalid = true;
            }
        }
        if let Some(index) = self.target_index(object_identifier) {
            if !binding.component().is_current()
                || !metadata_component_matches_physical(
                    binding.component().effective_locator(),
                    self.component_name,
                )
            {
                self.invalid = true;
            }
            if self.target_counts[index] != 0 {
                self.invalid = true;
            }
            self.target_counts[index] = self.target_counts[index].saturating_add(1);
            self.target_pairs[index] = Some(pair);
        }
        Ok(())
    }

    fn visit_external_reference(
        &mut self,
        reference: package_metadata_codec::ExternalReferenceDescriptor<'_>,
    ) -> Result<(), package_metadata_codec::RewriteError> {
        if reference
            .object_identifier()
            .is_some_and(|identifier| self.target_index(identifier).is_some())
        {
            self.invalid = true;
        }
        Ok(())
    }

    fn visit_data_reference(
        &mut self,
        reference: package_metadata_codec::DataReferenceDescriptor<'_>,
    ) -> Result<(), package_metadata_codec::RewriteError> {
        if self.data_ids.contains(&reference.data_identifier())
            && (!reference.component().is_current()
                || !metadata_component_matches_physical(
                    reference.component().effective_locator(),
                    self.component_name,
                ))
        {
            self.invalid = true;
        }
        if self.allow_audio_data_owners && self.data_ids.contains(&reference.data_identifier()) {
            let data_identifier = reference.data_identifier();
            let parent_count = self.expected_audio_owner_count(data_identifier);
            if self
                .audio_data_parent_counts
                .insert(data_identifier, reference.owner_count())
                .is_some()
            {
                self.invalid = true;
            }
            if reference.owner_count() != 0 && reference.owner_count() != parent_count {
                self.invalid = true;
            }
        }
        Ok(())
    }

    fn visit_data_reference_owner(
        &mut self,
        owner: package_metadata_codec::DataReferenceOwnerDescriptor<'_>,
    ) -> Result<(), package_metadata_codec::RewriteError> {
        let audio_owner = self.allow_audio_data_owners
            && self.data_ids.contains(&owner.data_identifier())
            && self.expected_audio_owner(owner.data_identifier(), owner.object_identifier())
            && owner.count() == 1
            && owner.component().is_current()
            && metadata_component_matches_physical(
                owner.component().effective_locator(),
                self.component_name,
            );
        if audio_owner {
            if !self
                .audio_data_explicit_owners
                .insert((owner.data_identifier(), owner.object_identifier()))
            {
                self.invalid = true;
            }
            if owner.object_identifier() == self.selected_movie_identifier {
                let count = self
                    .audio_data_owner_counts
                    .entry(owner.data_identifier())
                    .or_insert(0);
                *count = count.saturating_add(1);
                if *count != 1 {
                    self.invalid = true;
                }
            }
        } else if self.target_index(owner.object_identifier()).is_some()
            || (self.allow_audio_data_owners && self.data_ids.contains(&owner.data_identifier()))
        {
            self.invalid = true;
        }
        Ok(())
    }

    fn visit_ambiguous_object_identifier(
        &mut self,
        _component: package_metadata_codec::ComponentDescriptor<'_>,
        identifier: u64,
    ) -> Result<(), package_metadata_codec::RewriteError> {
        if self.target_index(identifier).is_some() {
            self.invalid = true;
        }
        Ok(())
    }

    fn visit_data_metadata_map(
        &mut self,
        object_identifier: u64,
        _has_unknown_fields: bool,
    ) -> Result<(), package_metadata_codec::RewriteError> {
        if self.target_index(object_identifier).is_some() {
            self.invalid = true;
        }
        Ok(())
    }
}

fn validate_global_movie_references(
    package: &Package,
    graph: &MovieReferenceGraph<'_>,
    budget: &mut GeometryBudget,
) -> Result<(), SlideMovieGeometryError> {
    let MovieReferenceGraph {
        component_name,
        slide_identifier,
        slide_message_index,
        movie_identifier,
        movie_message_index,
        data_ids,
        audio_data_witnesses,
        media_kind,
    } = *graph;
    let archive_limits = package
        .state
        .options
        .archive()
        .effective_archive_limits()
        .map_err(map_archive_error)?;
    let expected_movie_references = package
        .object_with_component(slide_identifier)
        .and_then(|(_, object)| object.archive_info.message_infos.get(slide_message_index))
        .map(|info| {
            info.object_references
                .iter()
                .filter(|identifier| **identifier == movie_identifier)
                .count()
                + info
                    .field_infos
                    .iter()
                    .flat_map(|field| field.object_references.iter())
                    .filter(|identifier| **identifier == movie_identifier)
                    .count()
        })
        .ok_or(SlideMovieGeometryError::InvalidSource)?;
    let expected_data_references = package
        .object_with_component(movie_identifier)
        .and_then(|(_, object)| object.archive_info.message_infos.get(movie_message_index))
        .map(|info| {
            info.data_references
                .iter()
                .filter(|identifier| data_ids.contains(identifier))
                .count()
                + info
                    .field_infos
                    .iter()
                    .flat_map(|field| field.data_references.iter())
                    .filter(|identifier| data_ids.contains(identifier))
                    .count()
        })
        .ok_or(SlideMovieGeometryError::InvalidSource)?;
    let mut visitor = MovieInboundReferenceVisitor {
        package,
        slide_identifier,
        slide_message_index,
        movie_identifier,
        movie_message_index,
        data_ids,
        audio_data_witnesses,
        component_name,
        allow_audio_shared_data: matches!(media_kind, GeometryMediaKind::Audio),
        movie_references: 0,
        data_references: 0,
        invalid: false,
    };
    for component in package.state.source.components().iter() {
        for object in &component.archive().objects {
            let mut fields = 0usize;
            let mut references = 0usize;
            let mut message_bytes = 0usize;
            for (index, message) in object.messages.iter().enumerate() {
                let info = object
                    .archive_info
                    .message_infos
                    .get(index)
                    .ok_or(SlideMovieGeometryError::InvalidSource)?;
                if message.type_ != info.type_
                    || usize::try_from(info.length).ok() != Some(message.data.len())
                {
                    return Err(SlideMovieGeometryError::InvalidSource);
                }
                fields = fields
                    .checked_add(info.field_infos.len())
                    .ok_or(SlideMovieGeometryError::InvalidSource)?;
                references = references
                    .checked_add(info.object_references.len())
                    .and_then(|value| value.checked_add(info.data_references.len()))
                    .ok_or(SlideMovieGeometryError::InvalidSource)?;
                for field in &info.field_infos {
                    references = references
                        .checked_add(field.object_references.len())
                        .and_then(|value| value.checked_add(field.data_references.len()))
                        .ok_or(SlideMovieGeometryError::InvalidSource)?;
                }
                message_bytes = message_bytes
                    .checked_add(message.data.len())
                    .ok_or(SlideMovieGeometryError::InvalidSource)?;
            }
            budget.fields(fields)?;
            budget.references(references)?;
            budget.work(
                message_bytes
                    .checked_add(fields)
                    .and_then(|value| value.checked_add(references))
                    .ok_or(SlideMovieGeometryError::InvalidSource)?,
            )?;
            budget.allocations(
                fields
                    .checked_add(references)
                    .and_then(|value| value.checked_add(1))
                    .ok_or(SlideMovieGeometryError::InvalidSource)?,
            )?;
            budget.retained(message_bytes)?;
            budget.scratch(message_bytes)?;
            if matches!(media_kind, GeometryMediaKind::Audio) {
                budget.work(
                    references
                        .checked_mul(audio_data_witnesses.len())
                        .ok_or(SlideMovieGeometryError::InvalidSource)?,
                )?;
            }
            object
                .inspect_references_with_policy_and_limits(
                    &mut visitor,
                    ArchiveReferencePolicy::RejectUnknownMetadata,
                    archive_limits,
                )
                .map_err(map_core_error)?;
        }
    }
    if visitor.invalid
        || visitor.movie_references != expected_movie_references
        || visitor.data_references != expected_data_references
    {
        return Err(SlideMovieGeometryError::UnsupportedDependency);
    }
    Ok(())
}

struct MovieInboundReferenceVisitor<'a> {
    package: &'a Package,
    slide_identifier: u64,
    slide_message_index: usize,
    movie_identifier: u64,
    movie_message_index: usize,
    data_ids: &'a [u64],
    audio_data_witnesses: &'a [(u64, u64)],
    component_name: &'a str,
    allow_audio_shared_data: bool,
    movie_references: usize,
    data_references: usize,
    invalid: bool,
}

impl ArchiveReferenceVisitor for MovieInboundReferenceVisitor<'_> {
    fn visit_reference(
        &mut self,
        occurrence: ArchiveReferenceOccurrence,
    ) -> litchi_iwa_core::Result<()> {
        if occurrence.referenced_identifier == 0 {
            self.invalid = true;
            return Ok(());
        }
        match occurrence.kind {
            ArchiveReferenceKind::Object => {
                if self
                    .package
                    .object_with_component(occurrence.referenced_identifier)
                    .is_none()
                {
                    let selected_scope = occurrence.object_identifier == self.slide_identifier
                        || occurrence.object_identifier == self.movie_identifier
                        || occurrence.referenced_identifier == self.movie_identifier;
                    if !self.allow_audio_shared_data || selected_scope {
                        self.invalid = true;
                    }
                }
                if occurrence.referenced_identifier == self.movie_identifier {
                    if occurrence.object_identifier != self.slide_identifier
                        || occurrence.message_index != self.slide_message_index
                    {
                        self.invalid = true;
                    } else {
                        self.movie_references = self.movie_references.saturating_add(1);
                    }
                }
            },
            ArchiveReferenceKind::Data => {
                if occurrence.referenced_identifier == self.movie_identifier
                    || occurrence.referenced_identifier == self.slide_identifier
                {
                    self.invalid = true;
                } else if self.data_ids.contains(&occurrence.referenced_identifier) {
                    let selected_movie_reference = occurrence.object_identifier
                        == self.movie_identifier
                        && occurrence.message_index == self.movie_message_index;
                    let shared_audio_reference = self.allow_audio_shared_data
                        && self.audio_data_witnesses.iter().any(
                            |(data_identifier, movie_identifier)| {
                                *data_identifier == occurrence.referenced_identifier
                                    && *movie_identifier == occurrence.object_identifier
                            },
                        )
                        && self
                            .package
                            .object_with_component(occurrence.object_identifier)
                            .is_some_and(|(component, object)| {
                                component == self.component_name
                                    && object
                                        .archive_info
                                        .message_infos
                                        .get(occurrence.message_index)
                                        .is_some_and(|info| info.type_ == MOVIE_MESSAGE_TYPE)
                            });
                    if selected_movie_reference {
                        self.data_references = self.data_references.saturating_add(1);
                    } else if !shared_audio_reference {
                        self.invalid = true;
                    }
                }
            },
        }
        Ok(())
    }
}

fn ensure_unique_movie_identity(
    package: &Package,
    component: &str,
    identifier: u64,
    budget: &mut GeometryBudget,
) -> Result<(), SlideMovieGeometryError> {
    let mut total = 0usize;
    let mut selected = 0usize;
    for current in package.state.source.components().iter() {
        budget.work(current.archive().objects.len())?;
        for object in &current.archive().objects {
            budget.work(1)?;
            if object.archive_info.identifier == Some(identifier) {
                total += 1;
                if current.name() == component {
                    selected += 1;
                }
            }
        }
    }
    if total != 1 || selected != 1 {
        return Err(SlideMovieGeometryError::UnsupportedDependency);
    }
    Ok(())
}

fn ensure_unique_movie_owner(
    package: &Package,
    slide_identifier: u64,
    movie_identifier: u64,
    field_number: u32,
    limits: litchi_iwa_common::WireLimits,
    budget: &mut GeometryBudget,
) -> Result<(), SlideMovieGeometryError> {
    let mut total = 0usize;
    let mut selected = 0usize;
    for component in package.state.source.components().iter() {
        for object in &component.archive().objects {
            for message in &object.messages {
                if message.type_ != SLIDE_MESSAGE_TYPE {
                    continue;
                }
                for id in repeated_references(&message.data, field_number, limits, budget)? {
                    if id == movie_identifier {
                        total += 1;
                        if object.archive_info.identifier == Some(slide_identifier) {
                            selected += 1;
                        }
                    }
                }
            }
        }
    }
    if total != 1 || selected != 1 {
        return Err(SlideMovieGeometryError::InvalidSource);
    }
    Ok(())
}

pub(crate) fn physical_catalog(
    package: &Package,
) -> Result<&litchi_iwa_archive::SourceCatalog, SlideMovieGeometryError> {
    match &package.state.source {
        PhysicalSource::Package(source) => Ok(source),
        PhysicalSource::Semantic(_) => Err(SlideMovieGeometryError::UnsupportedSource),
    }
}

pub(crate) fn previews_absent(package: &Package) -> Result<bool, SlideMovieGeometryError> {
    let catalog = physical_catalog(package)?;
    super::rendering_invalidation::root_previews_absent(catalog.package())
        .map_err(|_| SlideMovieGeometryError::Verification)
}

fn map_slide_selector_error(error: crate::SlideSelectorError) -> SlideMovieGeometryError {
    match error {
        crate::SlideSelectorError::DuplicateSlideName { .. } => {
            SlideMovieGeometryError::AmbiguousSelector
        },
        crate::SlideSelectorError::EmptySlideName => SlideMovieGeometryError::EmptySlideName,
    }
}
fn map_read_error(error: ReadError) -> SlideMovieGeometryError {
    match error {
        ReadError::SemanticLimit {
            kind,
            observed,
            maximum,
            ..
        } => SlideMovieGeometryError::LimitExceeded {
            kind: match kind {
                SemanticLimitKind::Slides => SlideMovieGeometryLimitKind::Slides,
                SemanticLimitKind::References => SlideMovieGeometryLimitKind::References,
                _ => SlideMovieGeometryLimitKind::WireWork,
            },
            observed: observed as u64,
            maximum: maximum as u64,
        },
        ReadError::PayloadLimit {
            kind,
            observed,
            maximum,
            ..
        } => SlideMovieGeometryError::LimitExceeded {
            kind: match kind {
                super::PayloadLimitKind::Bytes => SlideMovieGeometryLimitKind::WireBytes,
                super::PayloadLimitKind::Fields => SlideMovieGeometryLimitKind::WireFields,
                super::PayloadLimitKind::Nesting => SlideMovieGeometryLimitKind::WireNesting,
                super::PayloadLimitKind::Work => SlideMovieGeometryLimitKind::WireWork,
            },
            observed: observed as u64,
            maximum: maximum as u64,
        },
        ReadError::Allocation { amount, .. } => SlideMovieGeometryError::Allocation { amount },
        _ => SlideMovieGeometryError::InvalidSource,
    }
}
fn map_wire_error(_error: litchi_iwa_common::Error) -> SlideMovieGeometryError {
    SlideMovieGeometryError::InvalidSource
}

fn map_metadata_error(error: package_metadata_codec::RewriteError) -> SlideMovieGeometryError {
    if let Some(limit) = error.resource_limit() {
        let (kind, observed, maximum) = match limit {
            package_metadata_codec::RewriteLimit::InputBytes { observed, maximum } => {
                (SlideMovieGeometryLimitKind::WireBytes, observed, maximum)
            },
            package_metadata_codec::RewriteLimit::OutputBytes { observed, maximum } => {
                (SlideMovieGeometryLimitKind::OutputBytes, observed, maximum)
            },
            package_metadata_codec::RewriteLimit::Fields { observed, maximum } => {
                (SlideMovieGeometryLimitKind::WireFields, observed, maximum)
            },
            package_metadata_codec::RewriteLimit::Work { observed, maximum } => {
                (SlideMovieGeometryLimitKind::WireWork, observed, maximum)
            },
            package_metadata_codec::RewriteLimit::Nesting { observed, maximum } => (
                SlideMovieGeometryLimitKind::WireNesting,
                observed as usize,
                maximum as usize,
            ),
            package_metadata_codec::RewriteLimit::Components { observed, maximum } => {
                (SlideMovieGeometryLimitKind::Components, observed, maximum)
            },
            package_metadata_codec::RewriteLimit::References { observed, maximum } => {
                (SlideMovieGeometryLimitKind::References, observed, maximum)
            },
            package_metadata_codec::RewriteLimit::Additions { observed, maximum } => {
                (SlideMovieGeometryLimitKind::Entries, observed, maximum)
            },
            _ => return SlideMovieGeometryError::InvalidSource,
        };
        return SlideMovieGeometryError::LimitExceeded {
            kind,
            observed: observed as u64,
            maximum: maximum as u64,
        };
    }
    if let Some(amount) = error.allocation_request() {
        return SlideMovieGeometryError::Allocation { amount };
    }
    SlideMovieGeometryError::InvalidSource
}

fn map_archive_error(error: litchi_iwa_archive::Error) -> SlideMovieGeometryError {
    match error {
        litchi_iwa_archive::Error::Limit {
            kind,
            observed,
            maximum,
        } => SlideMovieGeometryError::LimitExceeded {
            kind: match kind {
                litchi_iwa_archive::LimitKind::InputBytes => {
                    SlideMovieGeometryLimitKind::InputBytes
                },
                litchi_iwa_archive::LimitKind::OutputBytes => {
                    SlideMovieGeometryLimitKind::OutputBytes
                },
                litchi_iwa_archive::LimitKind::Entries => SlideMovieGeometryLimitKind::Entries,
                litchi_iwa_archive::LimitKind::EntryBytes
                | litchi_iwa_archive::LimitKind::CompressedEntryBytes => {
                    SlideMovieGeometryLimitKind::EntryBytes
                },
                litchi_iwa_archive::LimitKind::TotalBytes => {
                    SlideMovieGeometryLimitKind::TotalBytes
                },
                _ => SlideMovieGeometryLimitKind::WireBytes,
            },
            observed,
            maximum,
        },
        litchi_iwa_archive::Error::Iwa(error) => map_core_error(error),
        litchi_iwa_archive::Error::Allocation { amount, .. } => {
            SlideMovieGeometryError::Allocation { amount }
        },
        _ => SlideMovieGeometryError::InvalidSource,
    }
}
fn map_core_error(error: litchi_iwa_core::Error) -> SlideMovieGeometryError {
    match error {
        litchi_iwa_core::Error::Limit {
            kind,
            observed,
            maximum,
        } => SlideMovieGeometryError::LimitExceeded {
            kind: match kind {
                litchi_iwa_core::LimitKind::Objects
                | litchi_iwa_core::LimitKind::Messages
                | litchi_iwa_core::LimitKind::MessagesPerObject => {
                    SlideMovieGeometryLimitKind::Entries
                },
                litchi_iwa_core::LimitKind::HeaderFields
                | litchi_iwa_core::LimitKind::MetadataItems => {
                    SlideMovieGeometryLimitKind::WireFields
                },
                litchi_iwa_core::LimitKind::HeaderNesting => {
                    SlideMovieGeometryLimitKind::WireNesting
                },
                litchi_iwa_core::LimitKind::SnappyCompressedChunkBytes
                | litchi_iwa_core::LimitKind::SnappyCompressedStreamBytes
                | litchi_iwa_core::LimitKind::SnappyFrames => {
                    SlideMovieGeometryLimitKind::EntryBytes
                },
                _ => SlideMovieGeometryLimitKind::WireBytes,
            },
            observed: observed as u64,
            maximum: maximum as u64,
        },
        litchi_iwa_core::Error::Allocation { requested, .. } => {
            SlideMovieGeometryError::Allocation { amount: requested }
        },
        _ => SlideMovieGeometryError::InvalidSource,
    }
}

#[cfg(test)]
mod validation_budget_tests {
    use super::{GeometryBudget, Package, SlideMovieGeometryError, SlideMovieGeometryLimitKind};

    #[test]
    fn semantic_validation_charges_inventory_before_decoding() {
        let source = include_bytes!(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../test-data/iwork/keynote/media-comments-baseline-native.key"
        ));
        let package = Package::from_bytes(source).unwrap();
        let payload_bytes: usize = package
            .state
            .source
            .components()
            .iter()
            .flat_map(|component| &component.archive().objects)
            .flat_map(|object| &object.messages)
            .map(|message| message.data.len())
            .sum();
        let mut measured = GeometryBudget::new(&package).unwrap();
        measured.validate_package(&package).unwrap();
        assert!(payload_bytes > 0);
        let header_fields: usize = package
            .state
            .source
            .components()
            .iter()
            .flat_map(|component| &component.archive().objects)
            .flat_map(|object| &object.archive_info.message_infos)
            .map(|info| info.field_infos.len())
            .sum();
        assert!(header_fields > 0);
        assert_eq!(measured.fields, header_fields);
        assert!(measured.work >= payload_bytes);
        assert!(measured.retained >= payload_bytes);
        assert!(measured.scratch >= payload_bytes);

        let mut exact = GeometryBudget::new(&package).unwrap();
        exact.max_fields = measured.fields;
        exact.validate_package(&package).unwrap();
        let mut short = GeometryBudget::new(&package).unwrap();
        short.max_fields = measured.fields - 1;
        assert!(matches!(
            short.validate_package(&package),
            Err(SlideMovieGeometryError::LimitExceeded {
                kind: SlideMovieGeometryLimitKind::WireFields,
                ..
            })
        ));
        assert_eq!(package.source_bytes(), source);
    }
}
