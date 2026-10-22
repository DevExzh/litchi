//! Immutable, selector-first transactions for Keynote slide order.

use std::fmt;
use std::sync::Arc;

use litchi_core::Position;
use litchi_iwa_archive::package::EntryEdit;
use litchi_iwa_common::{WireLimits, wire::WireView};
use litchi_iwa_core::SnappyStream;
use thiserror::Error as ThisError;

use super::{
    Package, PhysicalSource, ReadError, SHOW_MESSAGE_TYPE, decode_show_snapshot, unique_payload,
};
use crate::{SlideSelector, SlideSelectorError};

/// A finite resource governed while a slide-order transaction is prepared.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum LimitKind {
    /// Complete input package bytes.
    InputBytes,
    /// Complete rewritten package bytes.
    OutputBytes,
    /// Bytes in one protobuf payload.
    WireBytes,
    /// ZIP members or IWA objects.
    Entries,
    /// Bytes in one package member or IWA value.
    EntryBytes,
    /// Aggregate package or IWA bytes.
    TotalBytes,
    /// Semantic slides.
    Slides,
    /// Semantic graph references.
    References,
    /// Semantic text-storage objects.
    TextStorages,
    /// Semantic rich-text fragments.
    TextFragments,
    /// Aggregate retained semantic text.
    TextBytes,
    /// Parsed protobuf fields.
    WireFields,
    /// Protobuf nesting depth.
    WireNesting,
    /// Aggregate protobuf rewrite work.
    WireWork,
}

impl fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::InputBytes => "input bytes",
            Self::OutputBytes => "output bytes",
            Self::WireBytes => "wire bytes",
            Self::Entries => "entries",
            Self::EntryBytes => "entry bytes",
            Self::TotalBytes => "total bytes",
            Self::Slides => "slides",
            Self::References => "references",
            Self::TextStorages => "text storages",
            Self::TextFragments => "text fragments",
            Self::TextBytes => "text bytes",
            Self::WireFields => "wire fields",
            Self::WireNesting => "wire nesting",
            Self::WireWork => "wire work",
        })
    }
}

/// An error raised while staging or committing a slide-order transaction.
///
/// Error values contain only semantic positions and resource measurements;
/// native object identities, member names, and lower wire types stay private.
#[derive(Debug, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// A well-formed request this route refuses; [`Refusal::kind`] classifies it
    /// (record 0351 §2.4, record 0361).
    #[error(transparent)]
    Refused(Refusal),
    /// The source package or selected wire payload is structurally invalid.
    #[error("the Keynote source cannot be reordered safely")]
    InvalidSource,
    /// A finite resource ceiling was exceeded.
    #[error("Keynote slide-order {kind} limit exceeded: observed {observed}, maximum {maximum}")]
    Limit {
        /// Resource category that exceeded its limit.
        kind: LimitKind,
        /// Observed or requested amount.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
    },
    /// A bounded destination allocation failed before publication.
    #[error("could not allocate {amount} units for the Keynote slide-order transaction")]
    Allocation {
        /// Elements or bytes requested.
        amount: usize,
    },
    /// Full candidate reopening did not reproduce the requested order.
    #[error("the reordered Keynote candidate failed semantic verification")]
    Verification,
    /// The patch does not belong to this exact immutable package artifact.
    #[error("the Keynote slide-order patch does not match the exact source package")]
    PatchConflict,
}

/// A well-formed request this family's route refuses, classified by
/// [`Refusal::kind`] (record 0351 §2.4, record 0361).
#[derive(Debug, ThisError)]
#[non_exhaustive]
pub enum Refusal {
    /// This package was prepared for semantic reading only and has no editable
    /// physical source.
    #[error("this Keynote source does not support physical edits")]
    UnsupportedSource,
    /// The presentation uses a secondary or hierarchical slide-order model
    /// that this focused transaction does not rewrite.
    #[error("this Keynote slide topology is not supported for reordering")]
    UnsupportedTopology,
    /// An exact-name selector matched more than one slide.
    #[error("the Keynote slide selector is ambiguous")]
    AmbiguousSelector,
    /// An exact-name selector was empty.
    #[error("the Keynote slide selector name cannot be empty")]
    EmptySlideName,
    /// An exact-name selector matched no slide.
    #[error("the Keynote show has no slide matching the requested name")]
    SlideNameNotFound,
    /// A checked source position does not exist.
    #[error("the Keynote show has no slide at position {position:?}")]
    SlidePositionNotFound {
        /// Missing checked source position.
        position: Position,
    },
    /// The requested final destination is outside the unchanged slide count.
    #[error(
        "slide-order destination {position:?} is outside the presentation's {slide_count} slides"
    )]
    DestinationOutOfRange {
        /// Invalid checked final position.
        position: Position,
        /// Number of slides in the immutable base snapshot.
        slide_count: usize,
    },
    /// A second operation was staged in the same bounded transaction.
    #[error("the Keynote slide-order transaction already has a staged operation")]
    OperationAlreadyStaged,
    /// Commit was requested without a staged operation.
    #[error("the Keynote slide-order transaction has no staged operation")]
    NoStagedOperation,
}

impl Refusal {
    /// The shared refusal kind litchi-ops classifies this reason by.
    #[must_use]
    pub const fn kind(&self) -> litchi_core::edit::RefusalKind {
        use litchi_core::edit::RefusalKind;
        match self {
            Self::UnsupportedSource { .. } | Self::UnsupportedTopology { .. } => {
                RefusalKind::UnsupportedContent
            },
            Self::AmbiguousSelector { .. }
            | Self::EmptySlideName { .. }
            | Self::DestinationOutOfRange { .. } => RefusalKind::InvalidValue,
            Self::SlideNameNotFound { .. } | Self::SlidePositionNotFound { .. } => {
                RefusalKind::NotFound
            },
            Self::OperationAlreadyStaged { .. } | Self::NoStagedOperation { .. } => {
                RefusalKind::UnsupportedOperation
            },
        }
    }
}

impl From<Refusal> for Error {
    fn from(refusal: Refusal) -> Self {
        Self::Refused(refusal)
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct Intent {
    source: Position,
    destination: Position,
}

/// A bounded slide-order edit staged against an immutable Keynote snapshot.
#[derive(Debug)]
pub struct Edit<'a> {
    source: &'a Package,
    intent: Option<Intent>,
}

impl<'a> Edit<'a> {
    pub(super) const fn new(source: &'a Package) -> Self {
        Self {
            source,
            intent: None,
        }
    }

    /// Stage one slide move to a final checked presentation position.
    ///
    /// `destination` is interpreted in the final presentation order, matching
    /// `Vec::remove(source); Vec::insert(destination, value)`.
    ///
    /// # Errors
    ///
    /// Returns a typed error when the selector is missing or ambiguous, the
    /// destination is outside the base slide count, or an operation is already
    /// staged.
    pub fn move_slide<'selector>(
        &mut self,
        selector_input: impl Into<SlideSelector<'selector>>,
        destination: Position,
    ) -> Result<&mut Self, Error> {
        if self.intent.is_some() {
            return Err(Error::Refused(Refusal::OperationAlreadyStaged));
        }

        let selector = selector_input.into();
        let (source, slide_count) = match selector {
            SlideSelector::Position(position) => {
                let order = self.source.private_slide_node_order()?;
                if position.get() >= order.len() {
                    return Err(Error::Refused(Refusal::SlidePositionNotFound { position }));
                }
                (position, order.len())
            },
            SlideSelector::Name(_) => {
                let show = self.source.show().map_err(map_read_error)?;
                let selected = show
                    .select_slide(selector)
                    .map_err(map_selector_error)?
                    .ok_or(Error::Refused(Refusal::SlideNameNotFound))?;
                (Position::new(selected.index()), show.slides().len())
            },
        };
        if destination.get() >= slide_count {
            return Err(Error::Refused(Refusal::DestinationOutOfRange {
                position: destination,
                slide_count,
            }));
        }
        self.intent = Some(Intent {
            source,
            destination,
        });
        Ok(self)
    }

    /// Validate and atomically publish the staged immutable candidate.
    ///
    /// An exact same-position move reuses the source package allocation and
    /// bytes. A changed candidate is fully reopened and semantically read back
    /// under the original physical and semantic limits before publication.
    ///
    /// # Errors
    ///
    /// Returns an error without changing the source when staging is empty or
    /// any source, topology, wire, allocation, limit, or readback invariant
    /// fails.
    pub fn commit(self) -> Result<Commit, Error> {
        let intent = self
            .intent
            .ok_or(Error::Refused(Refusal::NoStagedOperation))?;
        let source_bytes = self.source.physical_shared_source()?;
        let source_fingerprint = fingerprint(&source_bytes);
        let before = self.source.private_slide_node_order()?;
        if intent.source.get() >= before.len() || intent.destination.get() >= before.len() {
            return Err(Error::InvalidSource);
        }
        if intent.source == intent.destination {
            self.source.validate().map_err(map_read_error)?;
            return Ok(Commit {
                package: self.source.snapshot(),
                patch: Patch {
                    source_bytes: Arc::clone(&source_bytes),
                    target_bytes: source_bytes,
                    source_fingerprint,
                    target_fingerprint: source_fingerprint,
                    source_position: intent.source,
                    destination: intent.destination,
                },
                diagnostics: Diagnostics::unchanged(),
            });
        }

        self.source.editable_source_catalog()?;
        // Reject malformed or over-budget downstream semantics before doing
        // any decompression, recompression, or package reassembly work. The
        // candidate is validated again after rewriting as the publication
        // boundary.
        self.source.validate().map_err(map_read_error)?;
        self.source.validate_flat_slide_order(&before)?;
        let after = moved_order(&before, intent)?;
        let package = rewrite_slide_order(self.source, intent, &after)?;
        let target_fingerprint = fingerprint(package.source_bytes());
        Ok(Commit {
            patch: Patch {
                source_bytes,
                target_bytes: package.editable_shared_source()?,
                source_fingerprint,
                target_fingerprint,
                source_position: intent.source,
                destination: intent.destination,
            },
            package,
            diagnostics: Diagnostics::published(),
        })
    }
}

/// An exact-source-checked, reversible semantic slide-order patch.
///
/// Native identities and package member names remain private. Public metadata
/// contains only the checked semantic source and destination positions.
#[derive(Clone, PartialEq, Eq)]
pub struct Patch {
    source_bytes: Arc<[u8]>,
    target_bytes: Arc<[u8]>,
    source_fingerprint: u64,
    target_fingerprint: u64,
    source_position: Position,
    destination: Position,
}

impl fmt::Debug for Patch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Patch")
            .field("source_position", &self.source_position)
            .field("destination", &self.destination)
            .finish_non_exhaustive()
    }
}

impl Patch {
    /// Return the base package's compact diagnostic fingerprint.
    #[must_use]
    pub const fn source_fingerprint(&self) -> u64 {
        self.source_fingerprint
    }

    /// Return the committed package's compact diagnostic fingerprint.
    #[must_use]
    pub const fn target_fingerprint(&self) -> u64 {
        self.target_fingerprint
    }

    /// Return the selected base-snapshot slide position.
    #[must_use]
    pub const fn source_position(&self) -> Position {
        self.source_position
    }

    /// Return the selected final destination position.
    #[must_use]
    pub const fn destination_position(&self) -> Position {
        self.destination
    }

    /// Return whether this patch preserves the exact source order and bytes.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.source_position.get() == self.destination.get()
            && (Arc::ptr_eq(&self.source_bytes, &self.target_bytes)
                || self.source_bytes.as_ref() == self.target_bytes.as_ref())
    }

    /// Return an exact reversible patch from the committed package back to its
    /// immutable source.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            source_bytes: Arc::clone(&self.target_bytes),
            target_bytes: Arc::clone(&self.source_bytes),
            source_fingerprint: self.target_fingerprint,
            target_fingerprint: self.source_fingerprint,
            source_position: self.destination,
            destination: self.source_position,
        }
    }
}

/// Compact evidence describing one slide-order commit.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Diagnostics {
    changed: bool,
    touched_components: usize,
    full_reparse_performed: bool,
}

impl Diagnostics {
    const fn unchanged() -> Self {
        Self {
            changed: false,
            touched_components: 0,
            full_reparse_performed: false,
        }
    }

    const fn published() -> Self {
        Self {
            changed: true,
            touched_components: 1,
            full_reparse_performed: true,
        }
    }

    /// Return whether the committed package differs from its source.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }

    /// Return the number of physical IWA components rewritten.
    #[must_use]
    pub const fn touched_components(self) -> usize {
        self.touched_components
    }

    /// Return whether the complete candidate was reopened before publication.
    #[must_use]
    pub const fn full_reparse_performed(self) -> bool {
        self.full_reparse_performed
    }
}

/// The fully verified result of one immutable slide-order transaction.
#[must_use = "a Keynote slide-order commit contains the validated package snapshot"]
#[derive(Debug)]
pub struct Commit {
    package: Package,
    patch: Patch,
    diagnostics: Diagnostics,
}

impl Commit {
    /// Borrow the fully reopened immutable package snapshot.
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.package
    }

    /// Consume this commit and return its immutable package snapshot.
    #[must_use]
    pub fn into_package(self) -> Package {
        self.package
    }

    /// Consume the commit and return its package and patch.
    #[must_use]
    pub fn into_parts(self) -> (Package, Patch) {
        (self.package, self.patch)
    }

    /// Borrow the reversible exact-source patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Borrow compact publication diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> &Diagnostics {
        &self.diagnostics
    }
}

impl Package {
    /// Start one selector-first slide-order edit.
    #[must_use]
    pub const fn edit_slide_order(&self) -> Edit<'_> {
        Edit::new(self)
    }

    /// Apply an exact-source-checked slide-order patch.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] unless this package is the
    /// exact immutable source captured by `patch`. The retained target is fully
    /// reopened and semantically verified under this package's original limits.
    pub fn apply_slide_order(&self, patch: &Patch) -> Result<Commit, Error> {
        let source = self.physical_shared_source()?;
        if fingerprint(self.source_bytes()) != patch.source_fingerprint
            || self.source_bytes() != patch.source_bytes.as_ref()
            || source.as_ref() != patch.source_bytes.as_ref()
        {
            return Err(Error::PatchConflict);
        }
        let current = self.private_slide_node_order()?;

        if patch.is_noop() {
            if patch.source_bytes.as_ref() != patch.target_bytes.as_ref() {
                return Err(Error::PatchConflict);
            }
            self.validate().map_err(map_read_error)?;
            return Ok(Commit {
                package: self.snapshot(),
                patch: patch.clone(),
                diagnostics: Diagnostics::unchanged(),
            });
        }

        self.editable_source_catalog()?;
        if fingerprint(&patch.target_bytes) != patch.target_fingerprint {
            return Err(Error::PatchConflict);
        }
        let intent = Intent {
            source: patch.source_position,
            destination: patch.destination,
        };
        if intent.source.get() >= current.len() || intent.destination.get() >= current.len() {
            return Err(Error::PatchConflict);
        }
        self.validate_flat_slide_order(&current)?;
        let expected = moved_order(&current, intent)?;
        let candidate =
            Package::from_source_with_options(Arc::clone(&patch.target_bytes), self.state.options)
                .map_err(map_read_error)?;
        candidate.validate().map_err(map_read_error)?;
        let readback = candidate.private_slide_node_order()?;
        if readback.as_ref() != expected.as_ref() {
            return Err(Error::Verification);
        }
        Ok(Commit {
            package: candidate,
            patch: patch.clone(),
            diagnostics: Diagnostics::published(),
        })
    }

    fn editable_shared_source(&self) -> Result<Arc<[u8]>, Error> {
        Ok(self.editable_source_catalog()?.shared_source())
    }

    #[allow(
        clippy::unnecessary_wraps,
        reason = "the internal prepared-source feature adds a semantic-only failure branch"
    )]
    fn physical_shared_source(&self) -> Result<Arc<[u8]>, Error> {
        match &self.state.source {
            PhysicalSource::Package(source) => Ok(source.shared_source()),
            PhysicalSource::Semantic(_) => Err(Error::Refused(Refusal::UnsupportedSource)),
        }
    }

    fn editable_source_catalog(&self) -> Result<&litchi_iwa_archive::SourceCatalog, Error> {
        let source = match &self.state.source {
            PhysicalSource::Package(source) => source,
            PhysicalSource::Semantic(_) => return Err(Error::Refused(Refusal::UnsupportedSource)),
        };
        if !source.source_is_exact() {
            return Err(Error::Refused(Refusal::UnsupportedSource));
        }
        Ok(source)
    }

    pub(super) fn private_slide_node_order(&self) -> Result<Box<[u64]>, Error> {
        let show_identifier = self.root_show_identifier().map_err(map_read_error)?;
        if show_identifier == 0 {
            return Ok(Box::default());
        }
        let object = self
            .required_object(show_identifier, "Keynote show")
            .map_err(map_read_error)?;
        let payload = unique_payload(&object.messages, &[SHOW_MESSAGE_TYPE], "Keynote show")
            .map_err(map_read_error)?;
        let mut budget = super::SemanticBudget::new(self.semantic_limits());
        budget
            .charge_references(1, super::SemanticPath::Show)
            .map_err(map_read_error)?;
        let preflight_count = super::preflight_show(
            payload,
            self.semantic_wire_limits().map_err(map_read_error)?,
            &mut budget,
        )
        .map_err(map_read_error)?;
        let snapshot = decode_show_snapshot(
            payload,
            self.semantic_limits().max_slides(),
            self.semantic_wire_limits().map_err(map_read_error)?,
        )
        .map_err(map_read_error)?;
        if snapshot.slide_node_identifiers().len() != preflight_count {
            return Err(Error::InvalidSource);
        }
        clone_identifiers(snapshot.slide_node_identifiers())
    }

    fn validate_flat_slide_order(&self, expected: &[u64]) -> Result<(), Error> {
        let mut unique = Vec::new();
        unique
            .try_reserve_exact(expected.len())
            .map_err(|_allocation| Error::Allocation {
                amount: expected.len(),
            })?;
        unique.extend_from_slice(expected);
        unique.sort_unstable();
        if unique.first() == Some(&0) || unique.windows(2).any(|pair| pair[0] == pair[1]) {
            return Err(Error::Refused(Refusal::UnsupportedTopology));
        }
        let show_identifier = self.root_show_identifier().map_err(map_read_error)?;
        let show_object = self
            .required_object(show_identifier, "Keynote show")
            .map_err(map_read_error)?;
        let show_payload =
            unique_payload(&show_object.messages, &[SHOW_MESSAGE_TYPE], "Keynote show")
                .map_err(map_read_error)?;
        let snapshot = decode_show_snapshot(
            show_payload,
            self.semantic_limits().max_slides(),
            self.semantic_wire_limits().map_err(map_read_error)?,
        )
        .map_err(map_read_error)?;
        if snapshot.has_deprecated_root_slide_node()
            || snapshot.has_slide_list()
            || snapshot.slide_node_identifiers() != expected
        {
            return Err(Error::Refused(Refusal::UnsupportedTopology));
        }
        let mut slide_identifiers = Vec::new();
        slide_identifiers
            .try_reserve_exact(expected.len())
            .map_err(|_allocation| Error::Allocation {
                amount: expected.len(),
            })?;
        for &identifier in expected {
            let node_object = self
                .required_object(identifier, "Keynote slide node")
                .map_err(map_read_error)?;
            let node_payload = unique_payload(&node_object.messages, &[4], "Keynote slide node")
                .map_err(map_read_error)?;
            let slide_identifier = flat_slide_identifier(
                node_payload,
                self.semantic_wire_limits().map_err(map_read_error)?,
            )?;
            if slide_identifier == 0 {
                return Err(Error::Refused(Refusal::UnsupportedTopology));
            }
            slide_identifiers.push(slide_identifier);
        }
        slide_identifiers.sort_unstable();
        if slide_identifiers.windows(2).any(|pair| pair[0] == pair[1]) {
            return Err(Error::Refused(Refusal::UnsupportedTopology));
        }
        Ok(())
    }
}

fn flat_slide_identifier(payload: &[u8], limits: WireLimits) -> Result<u64, Error> {
    let view = WireView::parse(payload, limits).map_err(map_wire_error)?;
    let mut slide_identifier = None;
    let mut depth_fields = 0usize;
    for field in view.fields() {
        match field.number() {
            1 => return Err(Error::Refused(Refusal::UnsupportedTopology)),
            2 => {
                if slide_identifier.is_some() {
                    return Err(Error::Refused(Refusal::UnsupportedTopology));
                }
                super::require_canonical_length_delimited(field, "Keynote slide reference")
                    .map_err(map_wire_error)?;
                slide_identifier = Some(
                    super::validate_reference_payload(
                        field.payload(),
                        limits,
                        "Keynote slide reference",
                    )
                    .map_err(map_wire_error)?,
                );
            },
            21 => {
                let depth = super::require_unique_uint64(
                    field,
                    &mut depth_fields,
                    "Keynote slide-node depth",
                )
                .map_err(map_wire_error)?;
                if depth != 1 {
                    return Err(Error::Refused(Refusal::UnsupportedTopology));
                }
            },
            _ => {},
        }
    }
    slide_identifier.ok_or(Error::Refused(Refusal::UnsupportedTopology))
}

fn moved_order(before: &[u64], intent: Intent) -> Result<Box<[u64]>, Error> {
    let mut after = Vec::new();
    after
        .try_reserve_exact(before.len())
        .map_err(|_allocation| Error::Allocation {
            amount: before.len(),
        })?;
    after.extend_from_slice(before);
    let identifier = after.remove(intent.source.get());
    after.insert(intent.destination.get(), identifier);
    Ok(after.into_boxed_slice())
}

fn rewrite_slide_order(source: &Package, intent: Intent, after: &[u64]) -> Result<Package, Error> {
    let source_catalog = source.editable_source_catalog()?;
    let snappy_limits = source_catalog
        .limits()
        .snappy_limits()
        .map_err(map_archive_error)?;
    let show_identifier = source.root_show_identifier().map_err(map_read_error)?;
    let mut components = source_catalog
        .components()
        .iter()
        .filter(|component| component.archive().object(show_identifier).is_some());
    let component = components.next().ok_or(Error::InvalidSource)?;
    if components.next().is_some() {
        return Err(Error::InvalidSource);
    }
    let component_name = component.name();
    let entry = source_catalog
        .package()
        .iter()
        .find(|entry| entry.name() == component_name)
        .ok_or(Error::InvalidSource)?;
    if entry.is_opaque() {
        return Err(Error::InvalidSource);
    }

    let stream = SnappyStream::decompress(
        entry.data(),
        source_catalog
            .limits()
            .snappy_limits()
            .map_err(map_archive_error)?,
    )
    .map_err(map_core_error)?;
    let object = component
        .archive()
        .object(show_identifier)
        .ok_or(Error::InvalidSource)?;
    let mut messages = object
        .messages
        .iter()
        .enumerate()
        .filter(|(_index, message)| message.type_ == SHOW_MESSAGE_TYPE);
    let (message_index, message) = messages.next().ok_or(Error::InvalidSource)?;
    if messages.next().is_some() {
        return Err(Error::InvalidSource);
    }
    let rewritten = permute_slide_reference_records(
        &message.data,
        after.len(),
        intent,
        source.wire_limits().map_err(map_wire_error)?,
    )?;
    let readback = decode_show_snapshot(
        &rewritten,
        source.semantic_limits().max_slides(),
        source.semantic_wire_limits().map_err(map_read_error)?,
    )
    .map_err(map_read_error)?;
    if readback.slide_node_identifiers() != after {
        return Err(Error::Verification);
    }

    if rewritten.len() != message.data.len() {
        return Err(Error::Verification);
    }
    let data_offset =
        usize::try_from(object.data_offset).map_err(|_conversion| Error::InvalidSource)?;
    let preceding = object
        .messages
        .iter()
        .take(message_index)
        .try_fold(0usize, |total, current| {
            total.checked_add(current.data.len())
        })
        .ok_or(Error::InvalidSource)?;
    let message_start = data_offset
        .checked_add(preceding)
        .ok_or(Error::InvalidSource)?;
    let message_end = message_start
        .checked_add(message.data.len())
        .ok_or(Error::InvalidSource)?;
    if stream.as_bytes().get(message_start..message_end) != Some(message.data.as_slice()) {
        return Err(Error::InvalidSource);
    }
    let mut rewritten_stream = stream.into_bytes();
    rewritten_stream[message_start..message_end].copy_from_slice(&rewritten);
    let compressed =
        SnappyStream::compress(&rewritten_stream, snappy_limits).map_err(map_core_error)?;
    let output = source_catalog
        .package()
        .reassemble_to_bytes(
            &[EntryEdit::new(component_name, &compressed)],
            source_catalog.limits(),
        )
        .map_err(map_archive_error)?;
    let candidate = Package::from_source_with_options(output.into(), source.state.options)
        .map_err(map_read_error)?;
    candidate.validate().map_err(map_read_error)?;
    if candidate.private_slide_node_order()?.as_ref() != after {
        return Err(Error::Verification);
    }
    Ok(candidate)
}

fn permute_slide_reference_records(
    show: &[u8],
    slide_count: usize,
    intent: Intent,
    limits: WireLimits,
) -> Result<Vec<u8>, Error> {
    if intent.source.get() >= slide_count || intent.destination.get() >= slide_count {
        return Err(Error::InvalidSource);
    }
    let show_view = WireView::parse(show, limits).map_err(map_wire_error)?;
    let mut slide_trees = show_view.fields().filter(|field| field.number() == 3);
    let slide_tree = slide_trees.next().ok_or(Error::InvalidSource)?;
    if slide_trees.next().is_some() || slide_tree.wire_type() != 2 {
        return Err(Error::InvalidSource);
    }
    let tree_view = WireView::parse(slide_tree.payload(), limits).map_err(map_wire_error)?;
    let work = show_view
        .len()
        .checked_add(tree_view.len())
        .and_then(|value| value.checked_add(slide_count))
        .ok_or(Error::Limit {
            kind: LimitKind::WireWork,
            observed: u64::MAX,
            maximum: usize_to_u64(limits.max_rewrite_work()),
        })?;
    if work > limits.max_rewrite_work() {
        return Err(Error::Limit {
            kind: LimitKind::WireWork,
            observed: usize_to_u64(work),
            maximum: usize_to_u64(limits.max_rewrite_work()),
        });
    }

    let mut records = Vec::new();
    records
        .try_reserve_exact(slide_count)
        .map_err(|_allocation| Error::Allocation {
            amount: slide_count,
        })?;
    for field in tree_view.fields().filter(|field| field.number() == 2) {
        if field.wire_type() != 2 {
            return Err(Error::InvalidSource);
        }
        records.push(field.raw());
    }
    if records.len() != slide_count {
        return Err(Error::InvalidSource);
    }

    let mut rewritten_tree = allocate_exact(slide_tree.payload().len())?;
    let mut slide_slot = 0usize;
    for field in tree_view.fields() {
        if field.number() == 2 {
            let source_slot = source_slot_for_destination(slide_slot, intent);
            rewritten_tree.extend_from_slice(
                records
                    .get(source_slot)
                    .copied()
                    .ok_or(Error::InvalidSource)?,
            );
            slide_slot += 1;
        } else {
            rewritten_tree.extend_from_slice(field.raw());
        }
    }
    if slide_slot != records.len() || rewritten_tree.len() != slide_tree.payload().len() {
        return Err(Error::Verification);
    }

    let mut rewritten_show = allocate_exact(show.len())?;
    for field in show_view.fields() {
        if field.number() == 3 {
            let raw = field.raw();
            let header_length = raw
                .len()
                .checked_sub(field.payload().len())
                .ok_or(Error::InvalidSource)?;
            rewritten_show.extend_from_slice(&raw[..header_length]);
            rewritten_show.extend_from_slice(&rewritten_tree);
        } else {
            rewritten_show.extend_from_slice(field.raw());
        }
    }
    if rewritten_show.len() != show.len() || rewritten_show.len() > limits.max_output_bytes() {
        return Err(Error::Limit {
            kind: LimitKind::OutputBytes,
            observed: usize_to_u64(rewritten_show.len()),
            maximum: usize_to_u64(limits.max_output_bytes()),
        });
    }
    Ok(rewritten_show)
}

const fn source_slot_for_destination(slot: usize, intent: Intent) -> usize {
    let source = intent.source.get();
    let destination = intent.destination.get();
    if destination < source {
        if slot == destination {
            source
        } else if slot > destination && slot <= source {
            slot - 1
        } else {
            slot
        }
    } else if slot == destination {
        source
    } else if slot >= source && slot < destination {
        slot + 1
    } else {
        slot
    }
}

fn allocate_exact(capacity: usize) -> Result<Vec<u8>, Error> {
    let mut output = Vec::new();
    output
        .try_reserve_exact(capacity)
        .map_err(|_allocation| Error::Allocation { amount: capacity })?;
    Ok(output)
}

fn clone_identifiers(identifiers: &[u64]) -> Result<Box<[u64]>, Error> {
    let mut cloned = Vec::new();
    cloned
        .try_reserve_exact(identifiers.len())
        .map_err(|_allocation| Error::Allocation {
            amount: identifiers.len(),
        })?;
    cloned.extend_from_slice(identifiers);
    Ok(cloned.into_boxed_slice())
}

fn map_selector_error(error: SlideSelectorError) -> Error {
    match error {
        SlideSelectorError::EmptySlideName => Error::Refused(Refusal::EmptySlideName),
        SlideSelectorError::DuplicateSlideName { .. } => Error::Refused(Refusal::AmbiguousSelector),
    }
}

fn map_read_error(error: ReadError) -> Error {
    match error {
        ReadError::SemanticLimit {
            kind,
            observed,
            maximum,
            ..
        } => Error::Limit {
            kind: match kind {
                super::SemanticLimitKind::Objects => LimitKind::Entries,
                super::SemanticLimitKind::Slides => LimitKind::Slides,
                super::SemanticLimitKind::References => LimitKind::References,
                super::SemanticLimitKind::TextStorages => LimitKind::TextStorages,
                super::SemanticLimitKind::TextFragments => LimitKind::TextFragments,
                super::SemanticLimitKind::TextBytes => LimitKind::TextBytes,
            },
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        },
        ReadError::PayloadLimit {
            kind,
            observed,
            maximum,
            ..
        } => Error::Limit {
            kind: match kind {
                super::PayloadLimitKind::Bytes => LimitKind::WireBytes,
                super::PayloadLimitKind::Fields => LimitKind::WireFields,
                super::PayloadLimitKind::Nesting => LimitKind::WireNesting,
                super::PayloadLimitKind::Work => LimitKind::WireWork,
            },
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        },
        ReadError::Allocation { amount, .. } => Error::Allocation { amount },
        ReadError::Archive(archive_error) => map_archive_error(archive_error),
        ReadError::Io(_)
        | ReadError::Detection(_)
        | ReadError::UnexpectedFormat { .. }
        | ReadError::InvalidFormat(_)
        | ReadError::Decode(_)
        | ReadError::TextStorage { .. }
        | ReadError::Metadata(_) => Error::InvalidSource,
    }
}

fn map_archive_error(error: litchi_iwa_archive::Error) -> Error {
    match error {
        litchi_iwa_archive::Error::Limit {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: match kind {
                litchi_iwa_archive::LimitKind::InputBytes => LimitKind::InputBytes,
                litchi_iwa_archive::LimitKind::OutputBytes => LimitKind::OutputBytes,
                litchi_iwa_archive::LimitKind::Entries => LimitKind::Entries,
                litchi_iwa_archive::LimitKind::MemberNameBytes
                | litchi_iwa_archive::LimitKind::MetadataBytes
                | litchi_iwa_archive::LimitKind::CompressedEntryBytes
                | litchi_iwa_archive::LimitKind::EntryBytes
                | litchi_iwa_archive::LimitKind::IwaStreamBytes => LimitKind::EntryBytes,
                litchi_iwa_archive::LimitKind::TotalBytes
                | litchi_iwa_archive::LimitKind::IwaTotalBytes => LimitKind::TotalBytes,
            },
            observed,
            maximum,
        },
        litchi_iwa_archive::Error::Allocation { amount, .. } => Error::Allocation { amount },
        litchi_iwa_archive::Error::Iwa(core_error) => map_core_error(core_error),
        _ => Error::InvalidSource,
    }
}

#[allow(
    clippy::needless_pass_by_value,
    reason = "used directly as a Result::map_err conversion"
)]
fn map_core_error(error: litchi_iwa_core::Error) -> Error {
    match error {
        litchi_iwa_core::Error::Limit {
            observed, maximum, ..
        } => Error::Limit {
            kind: LimitKind::EntryBytes,
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        },
        litchi_iwa_core::Error::Allocation { requested, .. } => {
            Error::Allocation { amount: requested }
        },
        _ => Error::InvalidSource,
    }
}

#[allow(
    clippy::needless_pass_by_value,
    reason = "used directly as a Result::map_err conversion"
)]
fn map_wire_error(error: litchi_iwa_common::Error) -> Error {
    match error {
        litchi_iwa_common::Error::LimitExceeded {
            kind,
            observed,
            limit,
        } => Error::Limit {
            kind: match kind {
                litchi_iwa_common::LimitKind::InputBytes => LimitKind::WireBytes,
                litchi_iwa_common::LimitKind::OutputBytes => LimitKind::OutputBytes,
                litchi_iwa_common::LimitKind::Fields
                | litchi_iwa_common::LimitKind::TableRows
                | litchi_iwa_common::LimitKind::TableColumns
                | litchi_iwa_common::LimitKind::TableCells
                | litchi_iwa_common::LimitKind::MaterializedCells => LimitKind::WireFields,
                litchi_iwa_common::LimitKind::Nesting => LimitKind::WireNesting,
                litchi_iwa_common::LimitKind::RewriteWork => LimitKind::WireWork,
            },
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(limit),
        },
        litchi_iwa_common::Error::Allocation { amount, .. } => Error::Allocation { amount },
        _ => Error::InvalidSource,
    }
}

fn usize_to_u64(value: usize) -> u64 {
    u64::try_from(value).unwrap_or(u64::MAX)
}

fn fingerprint(bytes: &[u8]) -> u64 {
    let mut value = 0xcbf2_9ce4_8422_2325_u64;
    for byte in bytes {
        value ^= u64::from(*byte);
        value = value.wrapping_mul(0x0000_0100_0000_01b3);
    }
    value
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn same_position_with_different_target_bytes_is_not_a_noop() {
        let source = Arc::<[u8]>::from(&b"source bytes"[..]);
        let target = Arc::<[u8]>::from(&b"target bytes"[..]);
        let patch = Patch {
            source_bytes: Arc::clone(&source),
            target_bytes: target,
            source_fingerprint: fingerprint(&source),
            target_fingerprint: fingerprint(b"target bytes"),
            source_position: Position::new(1),
            destination: Position::new(1),
        };

        assert!(!patch.is_noop());
    }

    #[test]
    fn same_position_with_equal_target_bytes_is_a_noop() {
        let source = Arc::<[u8]>::from(&b"source bytes"[..]);
        let target = Arc::<[u8]>::from(&b"source bytes"[..]);
        let patch = Patch {
            source_bytes: Arc::clone(&source),
            target_bytes: target,
            source_fingerprint: fingerprint(&source),
            target_fingerprint: fingerprint(&source),
            source_position: Position::new(1),
            destination: Position::new(1),
        };

        assert!(patch.is_noop());
    }
}
