//! Exact-source Keynote soundtrack playback-order transactions.
//!
//! This adapter changes only the repeated soundtrack media-reference records.
//! It deliberately does not own soundtrack media assets or playback settings.
//! The Buffa-backed soundtrack projection validates the selected graph and
//! streams its references; the caller-owned raw payload remains authoritative
//! for the focused order-preserving rewrite.

use std::path::Path;
use std::{collections::HashMap, fmt, path::Component as PathComponent, sync::Arc};

use litchi_core::Position;
use litchi_iwa_archive::{Limits, SourceCatalog, package::EntryEdit};
use litchi_iwa_common::{
    WireLimits, decode_varint_from_bytes,
    varint::encoded_len,
    wire::{WireDescent, WirePreflight, WireView, preflight_wire_tree_with_limits},
};
use litchi_iwa_core::{Archive, ArchiveObject, RawMessage, SnappyStream};
use litchi_iwa_protos::keynote_soundtrack_settings_codec as soundtrack_codec;
use thiserror::Error as ThisError;

use super::{DOCUMENT_MESSAGE_TYPE, Package, PhysicalSource, ReadError, SHOW_MESSAGE_TYPE};

const SOUNDTRACK_MESSAGE_TYPE: u32 = 21;
const DOCUMENT_SHOW_FIELD: u32 = 2;
const SHOW_SOUNDTRACK_FIELD: u32 = 17;
const SOUNDTRACK_MEDIA_FIELD: u32 = 3;
const LENGTH_DELIMITED_WIRE_TYPE: u8 = 2;
const METADATA_COMPONENT: &str = "Index/Metadata.iwa";
const PACKAGE_METADATA_MESSAGE_TYPE: u32 = 11_006;

// Reassembly may change only these selected-record ranges.  The local and
// central CRC/size triples are rewritten for the edited payload; the central
// local-header offset also follows the retained local record's physical move.
const ZIP_LOCAL_CRC_AND_SIZES: std::ops::Range<usize> = 14..26;
const ZIP_CENTRAL_CRC_AND_SIZES: std::ops::Range<usize> = 16..28;
const ZIP_CENTRAL_LOCAL_OFFSET: std::ops::Range<usize> = 42..46;

/// A finite resource governed by a soundtrack-order transaction.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Complete package input bytes.
    InputBytes,
    /// Complete package output bytes.
    OutputBytes,
    /// Retained package entries.
    Entries,
    /// One retained entry or encoded record.
    EntryBytes,
    /// Aggregate retained package bytes.
    TotalBytes,
    /// Semantic slides traversed while reopening a candidate.
    Slides,
    /// Semantic graph references traversed while reopening a candidate.
    References,
    /// Semantic text-storage objects traversed while reopening a candidate.
    TextStorages,
    /// Semantic rich-text fragments retained while reopening a candidate.
    TextFragments,
    /// Aggregate semantic text bytes retained while reopening a candidate.
    TextBytes,
    /// Parsed wire bytes.
    WireBytes,
    /// Parsed wire fields.
    WireFields,
    /// Wire nesting depth.
    WireNesting,
    /// Aggregate wire traversal and rewrite work.
    WireWork,
    /// Soundtrack media-reference records.
    Items,
}

impl fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::InputBytes => "input bytes",
            Self::OutputBytes => "output bytes",
            Self::Entries => "entries",
            Self::EntryBytes => "entry bytes",
            Self::TotalBytes => "total bytes",
            Self::Slides => "slides",
            Self::References => "references",
            Self::TextStorages => "text storages",
            Self::TextFragments => "text fragments",
            Self::TextBytes => "text bytes",
            Self::WireBytes => "wire bytes",
            Self::WireFields => "wire fields",
            Self::WireNesting => "wire nesting depth",
            Self::WireWork => "wire work",
            Self::Items => "soundtrack items",
        })
    }
}

/// A typed, content-redacted soundtrack-order transaction failure.
#[derive(Debug, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// The package has no exact physical source that can be changed.
    #[error("this Keynote source does not support physical soundtrack-order edits")]
    UnsupportedSource,
    /// The presentation has no soundtrack object.
    #[error("the Keynote presentation has no soundtrack")]
    SoundtrackNotFound,
    /// The selected source position is outside the soundtrack sequence.
    #[error("soundtrack source position {position:?} does not exist")]
    SourcePositionNotFound { position: Position },
    /// The requested final position is outside the soundtrack sequence.
    #[error("soundtrack destination position {position:?} is outside {item_count} items")]
    DestinationOutOfRange {
        position: Position,
        item_count: usize,
    },
    /// One edit may stage only one move.
    #[error("a soundtrack-order operation is already staged")]
    OperationAlreadyStaged,
    /// Commit was requested before staging a move.
    #[error("no soundtrack-order operation is staged")]
    NoStagedOperation,
    /// The selected graph or wire payload is not safely reorderable.
    #[error("the Keynote soundtrack graph cannot be reordered safely")]
    InvalidSource,
    /// A finite resource ceiling was exceeded.
    #[error(
        "Keynote soundtrack-order {kind} limit exceeded: observed {observed}, maximum {maximum}"
    )]
    LimitExceeded {
        kind: LimitKind,
        observed: u64,
        maximum: u64,
    },
    /// A bounded destination allocation failed.
    #[error("could not allocate {amount} soundtrack-order items or bytes")]
    Allocation { amount: usize },
    /// Candidate readback did not reproduce the requested order.
    #[error("the reordered Keynote soundtrack failed verification")]
    Verification,
    /// The patch was applied to a different exact source.
    #[error("the Keynote soundtrack-order patch conflicts with this package")]
    PatchConflict,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
struct Intent {
    source: Position,
    destination: Position,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
struct ReferenceFacts {
    identifier: u64,
    external: Option<bool>,
}

impl ReferenceFacts {
    /// Return whether this reference carries the legacy external marker.
    ///
    /// `TSP.Reference.deprecated_is_external` is still the only native
    /// external-reference bit in the wire schema.  Keep this check beside
    /// the strict parser so callers cannot accidentally treat a reference
    /// with an explicit `true` marker as an in-package object edge.
    fn is_external(self) -> bool {
        self.external == Some(true)
    }
}

#[derive(Clone, Copy, Debug)]
struct TransactionBudget {
    fields: usize,
    work: usize,
    references: usize,
    max_fields: usize,
    max_work: usize,
    max_references: usize,
    max_nesting: usize,
}

impl TransactionBudget {
    fn new(package: &Package) -> Result<Self, Error> {
        let limits = package.wire_limits().map_err(map_wire_error)?;
        Ok(Self {
            fields: 0,
            work: 0,
            references: 0,
            max_fields: limits.max_fields(),
            max_work: limits.max_rewrite_work(),
            max_references: package.semantic_limits().max_references(),
            max_nesting: limits.max_nesting(),
        })
    }

    fn charge_fields(&mut self, amount: usize) -> Result<(), Error> {
        self.fields = checked_charge(self.fields, amount, self.max_fields, LimitKind::WireFields)?;
        Ok(())
    }

    fn charge_work(&mut self, amount: usize) -> Result<(), Error> {
        self.work = checked_charge(self.work, amount, self.max_work, LimitKind::WireWork)?;
        Ok(())
    }

    fn charge_references(&mut self, amount: usize) -> Result<(), Error> {
        self.references = checked_charge(
            self.references,
            amount,
            self.max_references,
            LimitKind::References,
        )?;
        Ok(())
    }

    fn remaining_fields(self) -> usize {
        self.max_fields.saturating_sub(self.fields)
    }

    fn remaining_work(self) -> usize {
        self.max_work.saturating_sub(self.work)
    }

    fn require_depth(self, depth: usize) -> Result<(), Error> {
        if depth > self.max_nesting {
            return Err(Error::LimitExceeded {
                kind: LimitKind::WireNesting,
                observed: depth as u64,
                maximum: self.max_nesting as u64,
            });
        }
        Ok(())
    }

    fn merge_codec(&mut self, report: soundtrack_codec::DecodeReport) -> Result<(), Error> {
        self.charge_fields(report.fields())?;
        self.charge_work(report.work_bytes())
    }
}

/// Preflight one borrowed wire message before constructing a `WireView`.
///
/// `WireView` keeps one span allocation for every field.  The preflight scan
/// therefore runs first and checks the complete field/work/depth cost against
/// the transaction's remaining budget.  The caller still charges the actual
/// parse below; the scan is intentionally a guard, not a second logical
/// operation in the transaction accounting.
fn preflight_wire_message(
    source: &[u8],
    limits: WireLimits,
    budget: &TransactionBudget,
) -> Result<WirePreflight, Error> {
    let bounded_limits = remaining_wire_limits(limits, budget)?;
    let report =
        preflight_wire_tree_with_limits(source, bounded_limits, |_visit| Ok(WireDescent::Skip))
            .map_err(map_wire_error)?;
    ensure_wire_preflight_budget(report, budget)?;
    budget.require_depth(1)?;
    Ok(report)
}

/// Preflight a soundtrack payload and its field-3 data-reference messages.
///
/// The selected media records are the only deferred messages the focused
/// transaction descends into.  Unknown length-delimited fields remain opaque,
/// preserving the source's forward-compatible bytes while still proving the
/// exact nesting cost of every record that will be decoded and retained.
fn preflight_soundtrack_payload(
    source: &[u8],
    limits: WireLimits,
    budget: &TransactionBudget,
) -> Result<WirePreflight, Error> {
    let bounded_limits = remaining_wire_limits(limits, budget)?;
    let report = preflight_wire_tree_with_limits(source, bounded_limits, |visit| {
        let field = visit.field();
        if field.wire_type() == 3 || field.wire_type() == 4 {
            return Err(litchi_iwa_common::Error::InvalidFormat(
                "group-bearing soundtrack payload".to_owned(),
            ));
        }
        if visit.path().is_empty() && field.number() == SOUNDTRACK_MEDIA_FIELD {
            if field.wire_type() != LENGTH_DELIMITED_WIRE_TYPE {
                return Err(litchi_iwa_common::Error::InvalidFormat(
                    "soundtrack media field is not length-delimited".to_owned(),
                ));
            }
            field.validate_canonical_framing()?;
            Ok(WireDescent::Descend)
        } else {
            Ok(WireDescent::Skip)
        }
    })
    .map_err(map_wire_error)?;
    ensure_wire_preflight_budget(report, budget)?;
    // The common preflight reports the root as depth zero.  Buffa's strict
    // soundtrack decoder reports the same root as depth one and each media
    // reference as depth two, so make that convention explicit here.
    let observed_depth = report
        .max_depth()
        .checked_add(1)
        .ok_or(Error::InvalidSource)?;
    budget.require_depth(observed_depth)?;
    Ok(report)
}

fn ensure_wire_preflight_budget(
    report: WirePreflight,
    budget: &TransactionBudget,
) -> Result<(), Error> {
    if report.fields() > budget.remaining_fields() {
        let observed = budget
            .fields
            .checked_add(report.fields())
            .ok_or(Error::InvalidSource)?;
        return Err(Error::LimitExceeded {
            kind: LimitKind::WireFields,
            observed: observed as u64,
            maximum: budget.max_fields as u64,
        });
    }
    if report.scanned_bytes() > budget.remaining_work() {
        let observed = budget
            .work
            .checked_add(report.scanned_bytes())
            .ok_or(Error::InvalidSource)?;
        return Err(Error::LimitExceeded {
            kind: LimitKind::WireWork,
            observed: observed as u64,
            maximum: budget.max_work as u64,
        });
    }
    Ok(())
}

fn check_wire_output_bound(output_bytes: usize, limits: WireLimits) -> Result<(), Error> {
    if output_bytes > limits.max_output_bytes() {
        return Err(Error::LimitExceeded {
            kind: LimitKind::OutputBytes,
            observed: output_bytes as u64,
            maximum: limits.max_output_bytes() as u64,
        });
    }
    Ok(())
}

fn checked_charge(
    current: usize,
    amount: usize,
    maximum: usize,
    kind: LimitKind,
) -> Result<usize, Error> {
    let observed = current.checked_add(amount).ok_or(Error::InvalidSource)?;
    if observed > maximum {
        return Err(Error::LimitExceeded {
            kind,
            observed: observed as u64,
            maximum: maximum as u64,
        });
    }
    Ok(observed)
}

/// One immutable edit that stages exactly one soundtrack item move.
pub struct Edit<'a> {
    source: &'a Package,
    intent: Option<Intent>,
    selection: Option<Selection<'a>>,
    budget: Option<TransactionBudget>,
}

impl fmt::Debug for Edit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("source_position", &self.intent.map(|intent| intent.source))
            .field(
                "destination_position",
                &self.intent.map(|intent| intent.destination),
            )
            .finish_non_exhaustive()
    }
}

impl Edit<'_> {
    /// Stage one soundtrack item move to a checked final position.
    ///
    /// The destination is interpreted after removal, matching
    /// `Vec::remove(source); Vec::insert(destination, value)`. Equal source
    /// and destination positions are an exact no-op.
    pub fn move_item(
        &mut self,
        source: Position,
        destination: Position,
    ) -> Result<&mut Self, Error> {
        if self.intent.is_some() {
            return Err(Error::OperationAlreadyStaged);
        }
        let mut budget = TransactionBudget::new(self.source)?;
        let Some(selection) = resolve_selection(self.source, &mut budget, false)? else {
            return Err(Error::SoundtrackNotFound);
        };
        let item_count = selection.media.len();
        if source.get() >= item_count {
            return Err(Error::SourcePositionNotFound { position: source });
        }
        if destination.get() >= item_count {
            return Err(Error::DestinationOutOfRange {
                position: destination,
                item_count,
            });
        }
        self.intent = Some(Intent {
            source,
            destination,
        });
        self.selection = Some(selection);
        self.budget = Some(budget);
        Ok(self)
    }

    /// Validate and atomically publish the staged move.
    pub fn commit(self) -> Result<Commit, Error> {
        let intent = self.intent.ok_or(Error::NoStagedOperation)?;
        let source = physical_source(self.source)?;
        let source_bytes = source.shared_source();
        let selection = self.selection.ok_or(Error::InvalidSource)?;
        let mut budget = self.budget.ok_or(Error::InvalidSource)?;
        let source_ids = selection.media;
        validate_intent(intent, source_ids.len())?;

        if intent.source == intent.destination {
            return Ok(Commit {
                package: self.source.snapshot(),
                patch: Patch {
                    artifacts: litchi_iwa_archive::package::ExactArtifacts::new(
                        Arc::clone(&source_bytes),
                        source_bytes,
                    ),
                    source_position: intent.source,
                    destination_position: intent.destination,
                    item_count: source_ids.len(),
                },
                diagnostics: Diagnostics::unchanged(),
            });
        }
        if !source.source_is_exact() {
            return Err(Error::UnsupportedSource);
        }
        preflight_rewrite_output(physical_source(self.source)?)?;
        validate_mutation_source(self.source, &selection, &mut budget)?;
        let wire_limits = self.source.wire_limits().map_err(map_wire_error)?;
        // Admit the selected payload's complete field/work/depth shape before
        // retaining the moved-id vector or entering the rewrite allocator.
        preflight_soundtrack_payload(selection.soundtrack_payload, wire_limits, &budget)?;
        budget.charge_work(source_ids.len())?;
        let after = moved_ids(source_ids, intent)?;
        let package = rewrite(
            self.source,
            source,
            &selection,
            source_ids,
            &after,
            intent,
            &mut budget,
        )?;
        let target = physical_source(&package)?.shared_source();
        let target_selection =
            resolve_selection(&package, &mut budget, true)?.ok_or(Error::Verification)?;
        let target_ids = target_selection.media;
        if target_ids != after {
            return Err(Error::Verification);
        }
        verify_preserved_source(
            self.source,
            &package,
            &selection,
            &target_selection,
            &after,
            &mut budget,
        )?;
        Ok(Commit {
            package,
            patch: Patch {
                artifacts: litchi_iwa_archive::package::ExactArtifacts::new(source_bytes, target),
                source_position: intent.source,
                destination_position: intent.destination,
                item_count: source_ids.len(),
            },
            diagnostics: Diagnostics::published(),
        })
    }
}

/// An exact-source-checked, reversible soundtrack-order patch.
#[derive(Clone, PartialEq, Eq)]
pub struct Patch {
    artifacts: litchi_iwa_archive::package::ExactArtifacts,
    source_position: Position,
    destination_position: Position,
    item_count: usize,
}

impl fmt::Debug for Patch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Patch")
            .field("source_position", &self.source_position)
            .field("destination_position", &self.destination_position)
            .finish_non_exhaustive()
    }
}

impl Patch {
    /// Return the selected source position.
    #[must_use]
    pub const fn source_position(&self) -> Position {
        self.source_position
    }

    /// Return the selected final destination position.
    #[must_use]
    pub const fn destination_position(&self) -> Position {
        self.destination_position
    }

    /// Return the source package's compact diagnostic fingerprint.
    #[must_use]
    pub const fn source_fingerprint(&self) -> u64 {
        self.artifacts.source_fingerprint()
    }

    /// Return the target package's compact diagnostic fingerprint.
    #[must_use]
    pub const fn target_fingerprint(&self) -> u64 {
        self.artifacts.target_fingerprint()
    }

    /// Return whether this patch preserves the source bytes and order.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.source_position == self.destination_position && self.artifacts.is_byte_noop()
    }

    /// Return the exact target-to-source inverse in shared-handle work.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            artifacts: self.artifacts.inverse(),
            source_position: self.destination_position,
            destination_position: self.source_position,
            item_count: self.item_count,
        }
    }
}

/// Compact evidence describing one soundtrack-order commit.
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

    /// Return whether package bytes changed.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }

    /// Return the number of rewritten components.
    #[must_use]
    pub const fn touched_components(self) -> usize {
        self.touched_components
    }

    /// Return whether the candidate was fully reopened.
    #[must_use]
    pub const fn full_reparse_performed(self) -> bool {
        self.full_reparse_performed
    }
}

/// One successfully validated immutable soundtrack-order publication.
#[must_use = "a soundtrack-order commit contains the validated package snapshot"]
#[derive(Debug)]
pub struct Commit {
    package: Package,
    patch: Patch,
    diagnostics: Diagnostics,
}

impl Commit {
    /// Borrow the validated package snapshot.
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.package
    }

    /// Consume the commit and return its package snapshot.
    #[must_use]
    pub fn into_package(self) -> Package {
        self.package
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
    /// Start one immutable soundtrack-order edit.
    #[must_use]
    pub const fn edit_soundtrack_order(&self) -> Edit<'_> {
        Edit {
            source: self,
            intent: None,
            selection: None,
            budget: None,
        }
    }

    /// Apply an exact-source-checked soundtrack-order patch.
    pub fn apply_soundtrack_order(&self, patch: &Patch) -> Result<Commit, Error> {
        let source = physical_source(self)?;
        let source_bytes = source.shared_source();
        if !patch.artifacts.authorizes_source(&source_bytes) {
            return Err(Error::PatchConflict);
        }
        let mut budget = TransactionBudget::new(self)?;
        let before_selection =
            resolve_selection(self, &mut budget, false)?.ok_or(Error::PatchConflict)?;
        let before = before_selection.media;
        if before.len() != patch.item_count {
            return Err(Error::PatchConflict);
        }
        validate_intent(
            Intent {
                source: patch.source_position,
                destination: patch.destination_position,
            },
            before.len(),
        )?;
        if patch.is_noop() {
            // A no-op patch must carry an actually identical target.  Keep
            // this check before publishing the source snapshot so a forged
            // in-memory patch cannot bypass source/order validation.
            if !patch.artifacts.is_byte_noop() {
                return Err(Error::PatchConflict);
            }
            return Ok(Commit {
                package: self.snapshot(),
                patch: patch.clone(),
                diagnostics: Diagnostics::unchanged(),
            });
        }
        if !source.source_is_exact() {
            return Err(Error::PatchConflict);
        }
        validate_mutation_source(self, &before_selection, &mut budget)?;
        budget.charge_work(before.len())?;
        let expected = moved_ids(
            before,
            Intent {
                source: patch.source_position,
                destination: patch.destination_position,
            },
        )?;
        let target = patch.artifacts.target();
        if litchi_iwa_archive::package::ExactArtifacts::new(
            Arc::clone(&source_bytes),
            Arc::clone(&target),
        )
        .target_fingerprint()
            != patch.target_fingerprint()
        {
            return Err(Error::PatchConflict);
        }
        let target_len = target.len();
        budget.charge_work(
            target_len
                .checked_add(source_bytes.len())
                .ok_or(Error::InvalidSource)?,
        )?;
        let candidate = Package::from_source_with_options(target, self.state.options)
            .map_err(map_read_error)?;
        charge_catalog_reopen_cost(physical_source(&candidate)?, target_len, None, &mut budget)?;
        let after_selection =
            resolve_selection(&candidate, &mut budget, true)?.ok_or(Error::Verification)?;
        let after = after_selection.media;
        if after != expected {
            return Err(Error::Verification);
        }
        verify_preserved_source(
            self,
            &candidate,
            &before_selection,
            &after_selection,
            &expected,
            &mut budget,
        )?;
        Ok(Commit {
            package: candidate,
            patch: patch.clone(),
            diagnostics: Diagnostics::published(),
        })
    }
}

struct Selection<'a> {
    soundtrack_component: &'a str,
    show_identifier: u64,
    soundtrack_identifier: u64,
    root: &'a ArchiveObject,
    root_message_index: usize,
    show: &'a ArchiveObject,
    show_message_index: usize,
    soundtrack: &'a ArchiveObject,
    soundtrack_message_index: usize,
    soundtrack_payload: &'a [u8],
    media: &'a [u64],
}

fn resolve_selection<'a>(
    package: &'a Package,
    budget: &mut TransactionBudget,
    enforce_ownership: bool,
) -> Result<Option<Selection<'a>>, Error> {
    let catalog = physical_source(package)?;
    budget.charge_work(catalog.components().len())?;
    let mut roots = catalog
        .components()
        .iter()
        .filter(|component| component.name().rsplit('/').next() == Some("Document.iwa"));
    let root_component = roots.next().ok_or(Error::InvalidSource)?;
    if roots.next().is_some() {
        return Err(Error::InvalidSource);
    }
    let root = root_component
        .archive()
        .object(1)
        .ok_or(Error::InvalidSource)?;
    let limits = package.wire_limits().map_err(map_wire_error)?;
    let (root_message_index, root_payload) = selected_message(root, DOCUMENT_MESSAGE_TYPE)?;
    let root_reference =
        strict_optional_reference(root_payload, DOCUMENT_SHOW_FIELD, limits, budget)?;
    let Some(root_reference) = root_reference else {
        return Err(Error::InvalidSource);
    };
    if root_reference.is_external() {
        return Err(Error::InvalidSource);
    }
    let show_identifier = root_reference.identifier;
    if show_identifier == 0 || show_identifier == 1 {
        return Err(Error::InvalidSource);
    }
    let (_, show) = unique_object(catalog, show_identifier)?;
    let (show_message_index, show_payload) = selected_message(show, SHOW_MESSAGE_TYPE)?;
    let soundtrack_reference =
        strict_optional_reference(show_payload, SHOW_SOUNDTRACK_FIELD, limits, budget)?;
    let Some(soundtrack_reference) = soundtrack_reference else {
        return Ok(None);
    };
    if soundtrack_reference.is_external() {
        return Err(Error::InvalidSource);
    }
    let soundtrack_identifier = soundtrack_reference.identifier;
    if soundtrack_identifier == 0
        || soundtrack_identifier == 1
        || soundtrack_identifier == show_identifier
    {
        return Err(Error::InvalidSource);
    }
    let (soundtrack_component, soundtrack) = unique_object(catalog, soundtrack_identifier)?;
    let (soundtrack_message_index, soundtrack_payload) =
        selected_message(soundtrack, SOUNDTRACK_MESSAGE_TYPE)?;
    charge_message_info(root, root_message_index, budget)?;
    charge_message_info(show, show_message_index, budget)?;
    charge_message_info(soundtrack, soundtrack_message_index, budget)?;
    let nesting = u32::try_from(limits.max_nesting()).map_err(|_| Error::InvalidSource)?;
    let options = soundtrack_codec::DecodeOptions::new(
        soundtrack_payload.len(),
        budget.remaining_fields(),
        budget.remaining_work(),
        nesting,
    );
    let info = soundtrack
        .archive_info
        .message_infos
        .get(soundtrack_message_index)
        .ok_or(Error::InvalidSource)?;
    check_item_budget(package, info.data_references.len())?;
    budget.charge_references(info.data_references.len())?;
    let mut media_index = 0usize;
    let mut metadata_matches = true;
    let report = soundtrack_codec::visit_soundtrack_media_identifiers(
        soundtrack_payload,
        options,
        &mut |identifier| {
            if info.data_references.get(media_index).copied() != Some(identifier) {
                metadata_matches = false;
            }
            media_index = media_index.saturating_add(1);
            Ok(())
        },
    )
    .map_err(|error| map_codec_error(error, budget))?;
    if !metadata_matches
        || media_index != info.data_references.len()
        || report.media_references() != media_index
    {
        return Err(Error::InvalidSource);
    }
    budget.merge_codec(report)?;
    validate_soundtrack_media_references(
        soundtrack_payload,
        limits,
        info.data_references.len(),
        budget,
    )?;
    let selection = Selection {
        soundtrack_component,
        show_identifier,
        soundtrack_identifier,
        root,
        root_message_index,
        show,
        show_message_index,
        soundtrack,
        soundtrack_message_index,
        soundtrack_payload,
        media: &info.data_references,
    };
    if enforce_ownership {
        // The selected objects are only safe to rewrite when the incoming
        // graph proves the exact root -> show -> soundtrack ownership chain.
        // Keep these mutation-admission checks out of the selector itself so
        // exact no-op commits retain their source snapshot even when the
        // source has unrelated mutation metadata damage.
        validate_selected_metadata(root, root_message_index)?;
        validate_selected_metadata(show, show_message_index)?;
        validate_selected_metadata(soundtrack, soundtrack_message_index)?;
        validate_object_reference_metadata(
            root,
            root_message_index,
            show_identifier,
            DOCUMENT_SHOW_FIELD,
        )?;
        validate_object_reference_metadata(
            show,
            show_message_index,
            soundtrack_identifier,
            SHOW_SOUNDTRACK_FIELD,
        )?;
        validate_soundtrack_metadata(soundtrack, soundtrack_message_index)?;
        validate_media_closure(package, &selection, budget)?;
        validate_reference_role_disjointness(&selection, limits, budget)?;
    }
    Ok(Some(selection))
}

/// Validate the complete media-reference records, including fields that the
/// focused scalar codec intentionally ignores.  A native soundtrack record
/// may carry the legacy `deprecated_is_external` marker as an unknown field;
/// accepting that marker would make the archive index appear in-package while
/// the payload still points outside the selected host.
fn validate_soundtrack_media_references(
    source: &[u8],
    limits: WireLimits,
    expected: usize,
    budget: &mut TransactionBudget,
) -> Result<(), Error> {
    preflight_soundtrack_payload(source, limits, budget)?;
    budget.charge_work(source.len())?;
    let view = WireView::parse_with_limits(source, remaining_wire_limits(limits, budget)?)
        .map_err(map_wire_error)?;
    budget.charge_fields(view.len())?;
    let mut observed = 0usize;
    for field in view.fields() {
        if field.number() != SOUNDTRACK_MEDIA_FIELD {
            continue;
        }
        if field.wire_type() != LENGTH_DELIMITED_WIRE_TYPE {
            return Err(Error::InvalidSource);
        }
        field.validate_canonical_framing().map_err(map_wire_error)?;
        let reference = strict_reference(field.payload(), limits, budget)?;
        if reference.is_external() {
            return Err(Error::InvalidSource);
        }
        observed = observed.checked_add(1).ok_or(Error::InvalidSource)?;
    }
    if observed != expected {
        return Err(Error::InvalidSource);
    }
    Ok(())
}

fn check_item_budget(package: &Package, item_count: usize) -> Result<(), Error> {
    let maximum = package.semantic_limits().max_references();
    if item_count > maximum {
        return Err(Error::LimitExceeded {
            kind: LimitKind::Items,
            observed: item_count as u64,
            maximum: maximum as u64,
        });
    }
    Ok(())
}

fn validate_mutation_source(
    package: &Package,
    selection: &Selection<'_>,
    budget: &mut TransactionBudget,
) -> Result<(), Error> {
    let catalog = physical_source(package)?;
    if !catalog.source_is_exact() {
        return Err(Error::UnsupportedSource);
    }
    budget.charge_work(catalog.components().len())?;
    validate_selected_metadata(selection.root, selection.root_message_index)?;
    validate_selected_metadata(selection.show, selection.show_message_index)?;
    validate_selected_metadata(selection.soundtrack, selection.soundtrack_message_index)?;
    validate_object_reference_metadata(
        selection.root,
        selection.root_message_index,
        selection.show_identifier,
        DOCUMENT_SHOW_FIELD,
    )?;
    validate_object_reference_metadata(
        selection.show,
        selection.show_message_index,
        selection.soundtrack_identifier,
        SHOW_SOUNDTRACK_FIELD,
    )?;
    validate_soundtrack_metadata(selection.soundtrack, selection.soundtrack_message_index)?;
    validate_media_closure(package, selection, budget)?;
    let limits = package.wire_limits().map_err(map_wire_error)?;
    validate_reference_role_disjointness(selection, limits, budget)
}

fn charge_catalog_structure(
    catalog: &SourceCatalog,
    budget: &mut TransactionBudget,
) -> Result<(), Error> {
    for component in catalog.components().iter() {
        let archive = component.archive();
        budget.charge_work(archive.objects.len())?;
        for object in &archive.objects {
            budget.charge_work(
                object
                    .messages
                    .len()
                    .checked_add(object.archive_info.message_infos.len())
                    .ok_or(Error::InvalidSource)?,
            )?;
            for info in &object.archive_info.message_infos {
                let references = info
                    .object_references
                    .len()
                    .checked_add(info.data_references.len())
                    .ok_or(Error::InvalidSource)?;
                budget.charge_references(references)?;
                budget.charge_fields(info.field_infos.len())?;
                budget.charge_work(
                    info.versions
                        .len()
                        .checked_add(info.diff_merge_version.len())
                        .and_then(|amount| amount.checked_add(info.diff_read_version.len()))
                        .and_then(|amount| amount.checked_add(references))
                        .and_then(|amount| {
                            amount.checked_add(
                                info.diff_field_path
                                    .as_ref()
                                    .map_or(0, |path| path.path.len()),
                            )
                        })
                        .and_then(|amount| amount.checked_add(info.fields_to_remove.len()))
                        .ok_or(Error::InvalidSource)?,
                )?;
                for field in &info.field_infos {
                    let field_references = field
                        .object_references
                        .len()
                        .checked_add(field.data_references.len())
                        .ok_or(Error::InvalidSource)?;
                    budget.charge_references(field_references)?;
                    budget.charge_work(
                        1usize
                            .checked_add(field.path.path.len())
                            .and_then(|amount| amount.checked_add(field.known_field_version.len()))
                            .and_then(|amount| amount.checked_add(field_references))
                            .and_then(|amount| {
                                amount.checked_add(
                                    field
                                        .known_field_feature_identifier
                                        .as_ref()
                                        .map_or(0, String::len),
                                )
                            })
                            .ok_or(Error::InvalidSource)?,
                    )?;
                }
            }
        }
    }
    Ok(())
}

fn charge_catalog_reopen_cost(
    catalog: &SourceCatalog,
    raw_package_bytes: usize,
    replacement: Option<(&str, usize, usize)>,
    budget: &mut TransactionBudget,
) -> Result<(), Error> {
    let mut logical_bytes = 0usize;
    for entry in catalog.package().iter() {
        let bytes = replacement
            .filter(|(name, _, _)| *name == entry.name())
            .map_or(entry.data().len(), |(_, logical, _)| logical);
        logical_bytes = logical_bytes
            .checked_add(bytes)
            .ok_or(Error::InvalidSource)?;
    }
    let mut iwa_bytes = 0usize;
    for component in catalog.components().iter() {
        let bytes = replacement
            .filter(|(name, _, _)| *name == component.name())
            .map_or_else(
                || archive_stream_extent(component.archive()),
                |(_, _, decoded)| decoded,
            );
        iwa_bytes = iwa_bytes.checked_add(bytes).ok_or(Error::InvalidSource)?;
    }
    budget.charge_work(
        raw_package_bytes
            .checked_add(logical_bytes.checked_mul(2).ok_or(Error::InvalidSource)?)
            .and_then(|amount| amount.checked_add(iwa_bytes.checked_mul(2)?))
            .ok_or(Error::InvalidSource)?,
    )?;
    charge_catalog_structure(catalog, budget)
}

fn archive_stream_extent(archive: &Archive) -> usize {
    archive
        .objects
        .iter()
        .filter_map(|object| {
            let payload = object.messages.iter().try_fold(0usize, |amount, message| {
                amount.checked_add(message.data.len())
            })?;
            usize::try_from(object.data_offset)
                .ok()?
                .checked_add(payload)
        })
        .max()
        .unwrap_or(0)
}

fn charge_message_info(
    object: &ArchiveObject,
    index: usize,
    budget: &mut TransactionBudget,
) -> Result<(), Error> {
    let info = object
        .archive_info
        .message_infos
        .get(index)
        .ok_or(Error::InvalidSource)?;
    let aggregate_references = info
        .object_references
        .len()
        .checked_add(info.data_references.len())
        .ok_or(Error::InvalidSource)?;
    budget.charge_work(
        1usize
            .checked_add(aggregate_references)
            .ok_or(Error::InvalidSource)?,
    )?;
    budget.charge_references(aggregate_references)?;
    for field in &info.field_infos {
        let references = field
            .object_references
            .len()
            .checked_add(field.data_references.len())
            .ok_or(Error::InvalidSource)?;
        budget.charge_work(
            1usize
                .checked_add(field.path.path.len())
                .and_then(|amount| amount.checked_add(references))
                .ok_or(Error::InvalidSource)?,
        )?;
        budget.charge_references(references)?;
    }
    Ok(())
}

#[derive(Debug)]
struct MediaClosureState<'a> {
    payload_occurrences: usize,
    component_declarations: usize,
    owner_occurrences: usize,
    owner_count: usize,
    data_declarations: usize,
    filename: Option<&'a [u8]>,
    materialized_length: Option<usize>,
}

fn validate_media_closure(
    package: &Package,
    selection: &Selection<'_>,
    budget: &mut TransactionBudget,
) -> Result<(), Error> {
    let info = selection
        .soundtrack
        .archive_info
        .message_infos
        .get(selection.soundtrack_message_index)
        .ok_or(Error::InvalidSource)?;
    if info.data_references.is_empty() {
        return Ok(());
    }

    let mut states = HashMap::new();
    states
        .try_reserve(info.data_references.len())
        .map_err(|_| Error::Allocation {
            amount: info.data_references.len(),
        })?;
    for identifier in &info.data_references {
        let state = states.entry(*identifier).or_insert(MediaClosureState {
            payload_occurrences: 0,
            component_declarations: 0,
            owner_occurrences: 0,
            owner_count: 0,
            data_declarations: 0,
            filename: None,
            materialized_length: None,
        });
        state.payload_occurrences = state
            .payload_occurrences
            .checked_add(1)
            .ok_or(Error::InvalidSource)?;
    }

    let catalog = physical_source(package)?;
    let metadata_component = catalog
        .components()
        .get(METADATA_COMPONENT)
        .ok_or(Error::InvalidSource)?;
    let mut selected_metadata_payload = None;
    for object in &metadata_component.archive().objects {
        if object.messages.len() != object.archive_info.message_infos.len() {
            return Err(Error::InvalidSource);
        }
        for (index, (message, message_info)) in object
            .messages
            .iter()
            .zip(&object.archive_info.message_infos)
            .enumerate()
        {
            budget.charge_work(
                message
                    .data
                    .len()
                    .checked_add(1)
                    .ok_or(Error::InvalidSource)?,
            )?;
            if message.type_ != message_info.type_
                || usize::try_from(message_info.length).ok() != Some(message.data.len())
            {
                return Err(Error::InvalidSource);
            }
            if message.type_ == PACKAGE_METADATA_MESSAGE_TYPE {
                if selected_metadata_payload
                    .replace(message.data.as_slice())
                    .is_some()
                {
                    return Err(Error::InvalidSource);
                }
                charge_message_info(object, index, budget)?;
                validate_selected_metadata(object, index)?;
            }
        }
    }
    let metadata_payload = selected_metadata_payload.ok_or(Error::InvalidSource)?;
    let locator = selection
        .soundtrack_component
        .strip_prefix("Index/")
        .and_then(|name| name.strip_suffix(".iwa"))
        .ok_or(Error::InvalidSource)?;
    let selected_components = stream_package_metadata(
        metadata_payload,
        locator.as_bytes(),
        selection.soundtrack_identifier,
        &mut states,
        budget,
    )?;
    if selected_components != 1 {
        return Err(Error::InvalidSource);
    }

    let mut materialized = HashMap::new();
    materialized
        .try_reserve(states.len())
        .map_err(|_| Error::Allocation {
            amount: states.len(),
        })?;
    for state in states.values() {
        let filename = state.filename.ok_or(Error::InvalidSource)?;
        if materialized.insert(filename, (0usize, None)).is_some() {
            return Err(Error::InvalidSource);
        }
    }
    for entry in catalog.package().iter() {
        budget.charge_work(1)?;
        if let Some(filename) = entry.name().strip_prefix("Data/")
            && let Some((count, length)) = materialized.get_mut(filename.as_bytes())
        {
            *count = count.checked_add(1).ok_or(Error::InvalidSource)?;
            *length = Some(entry.data().len());
        }
    }
    for state in states.values() {
        let filename = state.filename.ok_or(Error::InvalidSource)?;
        if state.component_declarations != 1
            || state.owner_occurrences != 1
            || state.owner_count != state.payload_occurrences
            || state.data_declarations != 1
            || materialized.get(filename).copied() != Some((1, state.materialized_length))
        {
            return Err(Error::InvalidSource);
        }
    }
    Ok(())
}

fn stream_package_metadata<'a>(
    source: &'a [u8],
    locator: &[u8],
    soundtrack_identifier: u64,
    states: &mut HashMap<u64, MediaClosureState<'a>>,
    budget: &mut TransactionBudget,
) -> Result<usize, Error> {
    budget.require_depth(1)?;
    budget.charge_work(source.len())?;
    let mut input = source;
    let mut selected_components = 0usize;
    while let Some(field) = next_raw_field(&mut input, budget)? {
        match field.number {
            3 if field.wire == LENGTH_DELIMITED_WIRE_TYPE => {
                let component = field.bytes.ok_or(Error::InvalidSource)?;
                if stream_component_info(
                    component,
                    true,
                    locator,
                    soundtrack_identifier,
                    states,
                    budget,
                )? {
                    selected_components = selected_components
                        .checked_add(1)
                        .ok_or(Error::InvalidSource)?;
                }
            },
            11 if field.wire == LENGTH_DELIMITED_WIRE_TYPE => {
                // Versioned component declarations describe historical
                // owners.  They must never satisfy the closure for a
                // currently selected physical component.
                let component = field.bytes.ok_or(Error::InvalidSource)?;
                stream_component_info(
                    component,
                    false,
                    locator,
                    soundtrack_identifier,
                    states,
                    budget,
                )?;
            },
            4 if field.wire == LENGTH_DELIMITED_WIRE_TYPE => {
                stream_data_info(field.bytes.ok_or(Error::InvalidSource)?, states, budget)?;
            },
            3 | 4 | 11 => return Err(Error::InvalidSource),
            _ => {},
        }
    }
    Ok(selected_components)
}

fn stream_component_info(
    source: &[u8],
    scan_references: bool,
    locator: &[u8],
    soundtrack_identifier: u64,
    states: &mut HashMap<u64, MediaClosureState<'_>>,
    budget: &mut TransactionBudget,
) -> Result<bool, Error> {
    budget.require_depth(2)?;
    budget.charge_work(source.len())?;
    let mut input = source;
    let mut preferred = None;
    let mut current = None;
    while let Some(field) = next_raw_field(&mut input, budget)? {
        match field.number {
            2 if field.wire == LENGTH_DELIMITED_WIRE_TYPE && preferred.is_none() => {
                preferred = field.bytes
            },
            3 if field.wire == LENGTH_DELIMITED_WIRE_TYPE && current.is_none() => {
                current = field.bytes
            },
            2 | 3 => return Err(Error::InvalidSource),
            _ => {},
        }
    }
    let selected = scan_references && current.or(preferred) == Some(locator);
    if !selected {
        return Ok(false);
    }

    budget.charge_work(source.len())?;
    let mut reference_input = source;
    while let Some(field) = next_raw_field(&mut reference_input, budget)? {
        if field.number == 7 {
            if field.wire != LENGTH_DELIMITED_WIRE_TYPE {
                return Err(Error::InvalidSource);
            }
            stream_component_data_reference(
                field.bytes.ok_or(Error::InvalidSource)?,
                soundtrack_identifier,
                states,
                budget,
            )?;
        }
    }
    Ok(true)
}

fn stream_component_data_reference(
    source: &[u8],
    soundtrack_identifier: u64,
    states: &mut HashMap<u64, MediaClosureState<'_>>,
    budget: &mut TransactionBudget,
) -> Result<(), Error> {
    budget.require_depth(3)?;
    budget.charge_work(source.len())?;
    let mut input = source;
    let mut data_identifier = None;
    while let Some(field) = next_raw_field(&mut input, budget)? {
        match field.number {
            1 if field.wire == 0 && data_identifier.is_none() => data_identifier = field.varint,
            2 if field.wire == LENGTH_DELIMITED_WIRE_TYPE => {},
            1 | 2 => return Err(Error::InvalidSource),
            _ => {},
        }
    }
    let Some(state) = data_identifier.and_then(|identifier| states.get_mut(&identifier)) else {
        return Ok(());
    };
    state.component_declarations = state
        .component_declarations
        .checked_add(1)
        .ok_or(Error::InvalidSource)?;
    budget.charge_work(source.len())?;
    let mut owner_input = source;
    while let Some(owner_field) = next_raw_field(&mut owner_input, budget)? {
        if owner_field.number != 2 {
            continue;
        }
        let owner = owner_field.bytes.ok_or(Error::InvalidSource)?;
        budget.require_depth(4)?;
        budget.charge_work(owner.len())?;
        let mut owner_bytes = owner;
        let mut object = None;
        let mut count = None;
        while let Some(owner_value) = next_raw_field(&mut owner_bytes, budget)? {
            match owner_value.number {
                1 if owner_value.wire == 0 && object.is_none() => object = owner_value.varint,
                2 if owner_value.wire == 0 && count.is_none() => count = owner_value.varint,
                1 | 2 => return Err(Error::InvalidSource),
                _ => {},
            }
        }
        budget.charge_references(1)?;
        if object == Some(soundtrack_identifier) {
            state.owner_occurrences = state
                .owner_occurrences
                .checked_add(1)
                .ok_or(Error::InvalidSource)?;
            state.owner_count = usize::try_from(count.ok_or(Error::InvalidSource)?)
                .map_err(|_| Error::InvalidSource)?;
        }
    }
    Ok(())
}

fn stream_data_info<'a>(
    source: &'a [u8],
    states: &mut HashMap<u64, MediaClosureState<'a>>,
    budget: &mut TransactionBudget,
) -> Result<(), Error> {
    budget.require_depth(2)?;
    budget.charge_work(source.len())?;
    let mut input = source;
    let mut identifier = None;
    let mut digest = None;
    let mut preferred = None;
    let mut current = None;
    let mut materialized_length = None;
    while let Some(field) = next_raw_field(&mut input, budget)? {
        match field.number {
            1 if field.wire == 0 && identifier.is_none() => identifier = field.varint,
            2 if field.wire == LENGTH_DELIMITED_WIRE_TYPE && digest.is_none() => {
                digest = field.bytes
            },
            3 if field.wire == LENGTH_DELIMITED_WIRE_TYPE && preferred.is_none() => {
                preferred = field.bytes
            },
            4 if field.wire == LENGTH_DELIMITED_WIRE_TYPE && current.is_none() => {
                current = field.bytes
            },
            18 if field.wire == 0 && materialized_length.is_none() => {
                materialized_length = field.varint
            },
            1 | 2 | 3 | 4 | 18 => return Err(Error::InvalidSource),
            _ => {},
        }
    }
    let Some(state) = identifier.and_then(|data_identifier| states.get_mut(&data_identifier))
    else {
        return Ok(());
    };
    state.data_declarations = state
        .data_declarations
        .checked_add(1)
        .ok_or(Error::InvalidSource)?;
    if digest.is_none_or(|digest_bytes| digest_bytes.len() != 20) {
        return Err(Error::InvalidSource);
    }
    state.materialized_length = Some(
        usize::try_from(materialized_length.ok_or(Error::InvalidSource)?)
            .map_err(|_| Error::InvalidSource)?,
    );
    let filename_bytes = current.or(preferred).ok_or(Error::InvalidSource)?;
    let filename = std::str::from_utf8(filename_bytes).map_err(|_| Error::InvalidSource)?;
    if filename.is_empty()
        || filename.contains(['\0', '\\'])
        || Path::new(filename).is_absolute()
        || Path::new(filename)
            .components()
            .any(|component| !matches!(component, PathComponent::Normal(_)))
    {
        return Err(Error::InvalidSource);
    }
    state.filename = Some(filename.as_bytes());
    Ok(())
}

#[derive(Clone, Copy)]
struct RawField<'a> {
    number: u32,
    wire: u8,
    varint: Option<u64>,
    bytes: Option<&'a [u8]>,
}

fn next_raw_field<'a>(
    input: &mut &'a [u8],
    budget: &mut TransactionBudget,
) -> Result<Option<RawField<'a>>, Error> {
    if input.is_empty() {
        return Ok(None);
    }
    budget.charge_fields(1)?;
    let tag = take_canonical_varint(input)?;
    let number = u32::try_from(tag >> 3).map_err(|_| Error::InvalidSource)?;
    let wire = u8::try_from(tag & 7).map_err(|_| Error::InvalidSource)?;
    if number == 0 || number > 0x1fff_ffff {
        return Err(Error::InvalidSource);
    }
    let mut field = RawField {
        number,
        wire,
        varint: None,
        bytes: None,
    };
    match wire {
        0 => field.varint = Some(take_canonical_varint(input)?),
        1 => field.bytes = Some(take_bytes(input, 8)?),
        LENGTH_DELIMITED_WIRE_TYPE => {
            let length =
                usize::try_from(take_canonical_varint(input)?).map_err(|_| Error::InvalidSource)?;
            field.bytes = Some(take_bytes(input, length)?);
        },
        5 => field.bytes = Some(take_bytes(input, 4)?),
        _ => return Err(Error::InvalidSource),
    }
    Ok(Some(field))
}

fn take_canonical_varint(input: &mut &[u8]) -> Result<u64, Error> {
    let (value, consumed) = decode_varint_from_bytes(input).map_err(|_| Error::InvalidSource)?;
    if consumed != encoded_len(value) {
        return Err(Error::InvalidSource);
    }
    *input = input.get(consumed..).ok_or(Error::InvalidSource)?;
    Ok(value)
}

fn take_bytes<'a>(input: &mut &'a [u8], count: usize) -> Result<&'a [u8], Error> {
    let value = input.get(..count).ok_or(Error::InvalidSource)?;
    *input = input.get(count..).ok_or(Error::InvalidSource)?;
    Ok(value)
}

fn validate_soundtrack_metadata(object: &ArchiveObject, index: usize) -> Result<(), Error> {
    let info = object
        .archive_info
        .message_infos
        .get(index)
        .ok_or(Error::InvalidSource)?;
    if !info.object_references.is_empty() {
        return Err(Error::InvalidSource);
    }
    let mut media_path = false;
    for field in &info.field_infos {
        if !field.object_references.is_empty() {
            return Err(Error::InvalidSource);
        }
        if field.path.as_slice() == [SOUNDTRACK_MEDIA_FIELD] {
            if media_path
                || field.r#type != Some(litchi_iwa_core::FieldType::DataReference)
                || field.data_references != info.data_references
            {
                return Err(Error::InvalidSource);
            }
            media_path = true;
        } else if !field.data_references.is_empty() {
            return Err(Error::InvalidSource);
        }
    }
    if !media_path {
        return Err(Error::InvalidSource);
    }
    Ok(())
}

/// Reject message-level merge/diff metadata that would make a local payload
/// rewrite depend on an unmodelled base message or overlay.
fn validate_selected_metadata(object: &ArchiveObject, index: usize) -> Result<(), Error> {
    let info = object
        .archive_info
        .message_infos
        .get(index)
        .ok_or(Error::InvalidSource)?;
    if object.archive_info.should_merge == Some(true)
        || info.base_message_index.is_some()
        || !info.diff_merge_version.is_empty()
        || info.diff_field_path.is_some()
        || !info.fields_to_remove.is_empty()
        || !info.diff_read_version.is_empty()
    {
        return Err(Error::InvalidSource);
    }
    Ok(())
}

/// Prove that one selected object reference has exactly one owning field.
///
/// The physical archive index is intentionally not treated as an ownership
/// proof: duplicate aggregate references or a second field-local reference
/// would make the focused soundtrack rewrite ambiguous.
fn validate_object_reference_metadata(
    object: &ArchiveObject,
    index: usize,
    identifier: u64,
    path: u32,
) -> Result<(), Error> {
    let info = object
        .archive_info
        .message_infos
        .get(index)
        .ok_or(Error::InvalidSource)?;
    if info
        .object_references
        .iter()
        .filter(|candidate| **candidate == identifier)
        .count()
        != 1
    {
        return Err(Error::InvalidSource);
    }
    let mut selected_path = false;
    for field in &info.field_infos {
        if field.path.as_slice() == [path] {
            if selected_path
                || field.r#type != Some(litchi_iwa_core::FieldType::ObjectReference)
                || field.object_references.as_slice() != [identifier]
            {
                return Err(Error::InvalidSource);
            }
            selected_path = true;
        } else if field.object_references.contains(&identifier) {
            return Err(Error::InvalidSource);
        }
    }
    if !selected_path {
        return Err(Error::InvalidSource);
    }
    Ok(())
}

fn validate_reference_role_disjointness(
    selection: &Selection<'_>,
    limits: WireLimits,
    budget: &mut TransactionBudget,
) -> Result<(), Error> {
    let (_, root_payload) = selected_message(selection.root, DOCUMENT_MESSAGE_TYPE)?;
    reject_selected_identifier_in_reference_fields(
        root_payload,
        &[4],
        selection.soundtrack_identifier,
        limits,
        budget,
    )?;
    let (_, show_payload) = selected_message(selection.show, SHOW_MESSAGE_TYPE)?;
    preflight_wire_message(show_payload, limits, budget)?;
    reject_selected_identifier_in_reference_fields(
        show_payload,
        &[1, 2, 5, 7, 19],
        selection.soundtrack_identifier,
        limits,
        budget,
    )?;
    budget.charge_work(show_payload.len())?;
    let bounded_limits = remaining_wire_limits(limits, budget)?;
    let show = WireView::parse_with_limits(show_payload, bounded_limits).map_err(map_wire_error)?;
    budget.charge_fields(show.len())?;
    let mut slide_tree_payload = None;
    for field in show.fields() {
        if field.number() != 3 {
            continue;
        }
        if slide_tree_payload.is_some() || field.wire_type() != LENGTH_DELIMITED_WIRE_TYPE {
            return Err(Error::InvalidSource);
        }
        field.validate_canonical_framing().map_err(map_wire_error)?;
        slide_tree_payload = Some(field.payload());
    }
    let tree_payload = slide_tree_payload.ok_or(Error::InvalidSource)?;
    reject_selected_identifier_in_reference_fields(
        tree_payload,
        &[1, 2],
        selection.soundtrack_identifier,
        limits,
        budget,
    )
}

fn reject_selected_identifier_in_reference_fields(
    source: &[u8],
    reference_fields: &[u32],
    selected: u64,
    limits: WireLimits,
    budget: &mut TransactionBudget,
) -> Result<(), Error> {
    preflight_wire_message(source, limits, budget)?;
    budget.charge_work(source.len())?;
    let bounded_limits = remaining_wire_limits(limits, budget)?;
    let view = WireView::parse_with_limits(source, bounded_limits).map_err(map_wire_error)?;
    budget.charge_fields(view.len())?;
    for field in view.fields() {
        if !reference_fields.contains(&field.number()) {
            continue;
        }
        if field.wire_type() != LENGTH_DELIMITED_WIRE_TYPE {
            return Err(Error::InvalidSource);
        }
        field.validate_canonical_framing().map_err(map_wire_error)?;
        let reference = strict_reference(field.payload(), limits, budget)?;
        if reference.identifier == selected || reference.is_external() {
            return Err(Error::InvalidSource);
        }
    }
    Ok(())
}

fn selected_message(object: &ArchiveObject, kind: u32) -> Result<(usize, &[u8]), Error> {
    if object.messages.len() != object.archive_info.message_infos.len() {
        return Err(Error::InvalidSource);
    }
    let mut selected = None;
    for (index, (message, info)) in object
        .messages
        .iter()
        .zip(&object.archive_info.message_infos)
        .enumerate()
    {
        if message.type_ != info.type_
            || usize::try_from(info.length).ok() != Some(message.data.len())
        {
            return Err(Error::InvalidSource);
        }
        if message.type_ == kind && selected.replace((index, message.data.as_slice())).is_some() {
            return Err(Error::InvalidSource);
        }
    }
    selected.ok_or(Error::InvalidSource)
}

fn unique_object(
    catalog: &SourceCatalog,
    identifier: u64,
) -> Result<(&str, &ArchiveObject), Error> {
    let mut selected = None;
    for component in catalog.components().iter() {
        if let Some(object) = component.archive().object(identifier) {
            if selected.is_some() {
                return Err(Error::InvalidSource);
            }
            selected = Some((component.name(), object));
        }
    }
    selected.ok_or(Error::InvalidSource)
}

/// Check the exact-preservation boundary after a candidate has been
/// reopened.  Reassembly is allowed to change the selected component's
/// compressed bytes and the selected soundtrack payload/header metadata, but
/// every other logical package byte and archive object must remain untouched.
fn verify_preserved_source(
    source: &Package,
    candidate: &Package,
    before_selection: &Selection<'_>,
    after_selection: &Selection<'_>,
    expected_media: &[u64],
    budget: &mut TransactionBudget,
) -> Result<(), Error> {
    let before_catalog = physical_source(source)?;
    let after_catalog = physical_source(candidate)?;
    budget.charge_work(
        before_catalog
            .shared_source()
            .len()
            .checked_add(after_catalog.shared_source().len())
            .ok_or(Error::Verification)?,
    )?;
    if before_selection.soundtrack_component != after_selection.soundtrack_component
        || before_selection.soundtrack_identifier != after_selection.soundtrack_identifier
        || before_selection.soundtrack_message_index != after_selection.soundtrack_message_index
        || after_selection.media != expected_media
    {
        return Err(Error::Verification);
    }
    verify_zip_boundaries(before_catalog, after_catalog, budget)?;
    validate_soundtrack_metadata(
        after_selection.soundtrack,
        after_selection.soundtrack_message_index,
    )?;

    let mut before_entries = before_catalog.package().iter();
    let mut after_entries = after_catalog.package().iter();
    loop {
        match (before_entries.next(), after_entries.next()) {
            (Some(before), Some(after)) if before.name() == after.name() => {
                budget.charge_work(
                    before
                        .data()
                        .len()
                        .checked_add(after.data().len())
                        .and_then(|amount| {
                            amount.checked_add(before.raw_record().local_record().len())
                        })
                        .and_then(|amount| {
                            amount.checked_add(after.raw_record().local_record().len())
                        })
                        .and_then(|amount| {
                            amount.checked_add(before.raw_record().central_directory_record().len())
                        })
                        .and_then(|amount| {
                            amount.checked_add(after.raw_record().central_directory_record().len())
                        })
                        .ok_or(Error::Verification)?,
                )?;
                if before.name() != before_selection.soundtrack_component {
                    let local_offset_delta = record_offset_delta(
                        before_catalog.source_bytes(),
                        after_catalog.source_bytes(),
                        before.raw_record().local_record(),
                        after.raw_record().local_record(),
                    )
                    .ok_or(Error::Verification)?;
                    if before.data() != after.data()
                        || before.raw_name() != after.raw_name()
                        || before.raw_record().local_record() != after.raw_record().local_record()
                        || !central_record_preserved_except_offset(
                            before.raw_record().central_directory_record(),
                            after.raw_record().central_directory_record(),
                            local_offset_delta,
                        )
                    {
                        return Err(Error::Verification);
                    }
                } else if before.raw_name() != after.raw_name()
                    || before.metadata().local() != after.metadata().local()
                    || before.metadata().central() != after.metadata().central()
                    || !selected_local_record_preserved(
                        before,
                        after,
                        zip_entry_fields(before).ok_or(Error::Verification)?,
                        zip_entry_fields(after).ok_or(Error::Verification)?,
                    )
                    || !selected_central_record_preserved(
                        before.raw_record().central_directory_record(),
                        after.raw_record().central_directory_record(),
                        zip_entry_fields(before).ok_or(Error::Verification)?,
                        zip_entry_fields(after).ok_or(Error::Verification)?,
                        record_offset_delta(
                            before_catalog.source_bytes(),
                            after_catalog.source_bytes(),
                            before.raw_record().local_record(),
                            after.raw_record().local_record(),
                        )
                        .ok_or(Error::Verification)?,
                    )
                {
                    return Err(Error::Verification);
                }
            },
            (None, None) => break,
            _ => return Err(Error::Verification),
        }
    }

    let before_component = before_catalog
        .components()
        .get(before_selection.soundtrack_component)
        .ok_or(Error::Verification)?;
    let after_component = after_catalog
        .components()
        .get(after_selection.soundtrack_component)
        .ok_or(Error::Verification)?;
    if before_component.archive().objects.len() != after_component.archive().objects.len() {
        return Err(Error::Verification);
    }
    for (before_object, after_object) in before_component
        .archive()
        .objects
        .iter()
        .zip(&after_component.archive().objects)
    {
        budget.charge_work(1)?;
        if before_object.header_offset != after_object.header_offset
            || before_object.header_length != after_object.header_length
            || before_object.data_offset != after_object.data_offset
            || before_object.data_length != after_object.data_length
        {
            return Err(Error::Verification);
        }
        if before_object.archive_info.identifier != after_object.archive_info.identifier {
            return Err(Error::Verification);
        }
        if before_object.archive_info.identifier != Some(before_selection.soundtrack_identifier) {
            if before_object.archive_info != after_object.archive_info
                || before_object.messages != after_object.messages
            {
                return Err(Error::Verification);
            }
            continue;
        }
        if before_object.messages.len() != after_object.messages.len()
            || before_object.archive_info.should_merge != after_object.archive_info.should_merge
            || before_object.archive_info.message_infos.len()
                != after_object.archive_info.message_infos.len()
        {
            return Err(Error::Verification);
        }
        for (index, ((before_info, after_info), (before_message, after_message))) in before_object
            .archive_info
            .message_infos
            .iter()
            .zip(&after_object.archive_info.message_infos)
            .zip(before_object.messages.iter().zip(&after_object.messages))
            .enumerate()
        {
            if before_message.type_ != after_message.type_ {
                return Err(Error::Verification);
            }
            if index != before_selection.soundtrack_message_index {
                if before_info != after_info || before_message != after_message {
                    return Err(Error::Verification);
                }
            } else if before_info.data_references != before_selection.media
                || after_info.data_references != expected_media
                || !message_info_preserved_except_soundtrack_media_data_references(
                    before_info,
                    after_info,
                    expected_media,
                )
                || before_message.data.len() != after_message.data.len()
            {
                return Err(Error::Verification);
            }
        }
    }
    Ok(())
}

/// Reassembly patches only the four-byte relative local-header offset in a
/// retained central-directory record, and that field must move by the same
/// checked delta as the corresponding local record. All other bytes remain
/// exact source provenance, including names, flags, timestamps, extras,
/// comments, and archive-level attributes.
fn central_record_preserved_except_offset(
    before: &[u8],
    after: &[u8],
    expected_offset_delta: i128,
) -> bool {
    before.len() >= 46
        && before.len() == after.len()
        && is_central_record(before)
        && is_central_record(after)
        && bytes_equal_outside_ranges(before, after, ZIP_CENTRAL_LOCAL_OFFSET, None)
        && central_offset_delta(before, after) == Some(expected_offset_delta)
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
struct ZipEntryFields {
    crc32: u32,
    compressed_size: u32,
    uncompressed_size: u32,
}

fn zip_entry_fields(entry: &litchi_iwa_archive::package::Entry) -> Option<ZipEntryFields> {
    Some(ZipEntryFields {
        crc32: zip_crc32(entry.data()),
        compressed_size: u32::try_from(entry.raw_record().compressed_data().len()).ok()?,
        uncompressed_size: u32::try_from(entry.data().len()).ok()?,
    })
}

fn selected_local_record_preserved(
    before: &litchi_iwa_archive::package::Entry,
    after: &litchi_iwa_archive::package::Entry,
    before_fields: ZipEntryFields,
    after_fields: ZipEntryFields,
) -> bool {
    // CRC and size fields are the only mutable local-header bytes.  Their
    // exact before/after values are tied to the parsed entry facts below;
    // descriptor-bearing entries retain the local fields and move the same
    // three values into their descriptor instead.
    let before_record = before.raw_record().local_record();
    let after_record = after.raw_record().local_record();
    if !is_local_record(before_record) || !is_local_record(after_record) {
        return false;
    }
    let Some(before_header_length) = zip_local_header_length(before_record) else {
        return false;
    };
    let Some(after_header_length) = zip_local_header_length(after_record) else {
        return false;
    };
    if before_header_length != after_header_length
        || !bytes_equal_outside_ranges(
            &before_record[..before_header_length],
            &after_record[..after_header_length],
            ZIP_LOCAL_CRC_AND_SIZES,
            None,
        )
    {
        return false;
    }
    let Some(before_header_fields) = local_header_fields(before_record) else {
        return false;
    };
    let Some(after_header_fields) = local_header_fields(after_record) else {
        return false;
    };
    let flags = before.metadata().local().flags();
    if flags != after.metadata().local().flags() {
        return false;
    }
    if flags & 0x0008 == 0 {
        if before_header_fields != before_fields || after_header_fields != after_fields {
            return false;
        }
    } else if before_header_fields != after_header_fields {
        return false;
    }
    let Some(before_payload_end) = before_header_length
        .checked_add(before.raw_record().compressed_data().len())
        .filter(|end| *end <= before_record.len())
    else {
        return false;
    };
    let Some(after_payload_end) = after_header_length
        .checked_add(after.raw_record().compressed_data().len())
        .filter(|end| *end <= after_record.len())
    else {
        return false;
    };
    selected_local_suffix_preserved(
        flags,
        &before_record[before_payload_end..],
        &after_record[after_payload_end..],
        before_fields,
        after_fields,
    )
}

fn zip_local_header_length(record: &[u8]) -> Option<usize> {
    if record.get(..4)? != b"PK\x03\x04" {
        return None;
    }
    let name_length = usize::from(u16::from_le_bytes(record.get(26..28)?.try_into().ok()?));
    let extra_length = usize::from(u16::from_le_bytes(record.get(28..30)?.try_into().ok()?));
    30usize
        .checked_add(name_length)?
        .checked_add(extra_length)
        .filter(|length| *length <= record.len())
}

fn local_header_fields(record: &[u8]) -> Option<ZipEntryFields> {
    Some(ZipEntryFields {
        crc32: raw_u32(record, 14)?,
        compressed_size: raw_u32(record, 18)?,
        uncompressed_size: raw_u32(record, 22)?,
    })
}

fn selected_local_suffix_preserved(
    flags: u16,
    before: &[u8],
    after: &[u8],
    before_fields: ZipEntryFields,
    after_fields: ZipEntryFields,
) -> bool {
    if flags & 0x0008 == 0 {
        return before == after;
    }
    let before_descriptor = if before.starts_with(b"PK\x07\x08") {
        4
    } else {
        0
    };
    let after_descriptor = if after.starts_with(b"PK\x07\x08") {
        4
    } else {
        0
    };
    before_descriptor == after_descriptor
        && before.len() == after.len()
        && before.len() >= before_descriptor + 12
        && before[..before_descriptor] == after[..after_descriptor]
        && raw_descriptor_fields(before, before_descriptor) == Some(before_fields)
        && raw_descriptor_fields(after, after_descriptor) == Some(after_fields)
        && before[before_descriptor + 12..] == after[after_descriptor + 12..]
}

fn selected_central_record_preserved(
    before: &[u8],
    after: &[u8],
    before_fields: ZipEntryFields,
    after_fields: ZipEntryFields,
    expected_offset_delta: i128,
) -> bool {
    before.len() == after.len()
        && before.len() >= ZIP_CENTRAL_LOCAL_OFFSET.end
        && is_central_record(before)
        && is_central_record(after)
        && bytes_equal_outside_ranges(
            before,
            after,
            ZIP_CENTRAL_CRC_AND_SIZES,
            Some(ZIP_CENTRAL_LOCAL_OFFSET),
        )
        && central_fields(before) == Some(before_fields)
        && central_fields(after) == Some(after_fields)
        && central_offset_delta(before, after) == Some(expected_offset_delta)
}

fn is_local_record(record: &[u8]) -> bool {
    record.get(..4) == Some(b"PK\x03\x04")
}

fn is_central_record(record: &[u8]) -> bool {
    record.get(..4) == Some(b"PK\x01\x02")
}

fn bytes_equal_outside_ranges(
    before: &[u8],
    after: &[u8],
    first: std::ops::Range<usize>,
    second: Option<std::ops::Range<usize>>,
) -> bool {
    if before.len() != after.len() || first.start > first.end || first.end > before.len() {
        return false;
    }
    let second = second.unwrap_or(first.end..first.end);
    if second.start > second.end || second.end > before.len() || first.end > second.start {
        return false;
    }
    before[..first.start] == after[..first.start]
        && before[first.end..second.start] == after[first.end..second.start]
        && before[second.end..] == after[second.end..]
}

fn central_fields(record: &[u8]) -> Option<ZipEntryFields> {
    Some(ZipEntryFields {
        crc32: raw_u32(record, 16)?,
        compressed_size: raw_u32(record, 20)?,
        uncompressed_size: raw_u32(record, 24)?,
    })
}

fn raw_descriptor_fields(record: &[u8], start: usize) -> Option<ZipEntryFields> {
    Some(ZipEntryFields {
        crc32: raw_u32(record, start)?,
        compressed_size: raw_u32(record, start.checked_add(4)?)?,
        uncompressed_size: raw_u32(record, start.checked_add(8)?)?,
    })
}

fn raw_u32(record: &[u8], start: usize) -> Option<u32> {
    let end = start.checked_add(4)?;
    Some(u32::from_le_bytes(record.get(start..end)?.try_into().ok()?))
}

fn central_offset_delta(before: &[u8], after: &[u8]) -> Option<i128> {
    let before = i128::from(raw_u32(before, 42)?);
    let after = i128::from(raw_u32(after, 42)?);
    after.checked_sub(before)
}

fn record_offset_delta(
    before_source: &[u8],
    after_source: &[u8],
    before_record: &[u8],
    after_record: &[u8],
) -> Option<i128> {
    let before = i128::try_from(raw_slice_offset(before_source, before_record)?).ok()?;
    let after = i128::try_from(raw_slice_offset(after_source, after_record)?).ok()?;
    after.checked_sub(before)
}

fn zip_crc32(bytes: &[u8]) -> u32 {
    let mut crc = u32::MAX;
    for byte in bytes {
        crc ^= u32::from(*byte);
        for _ in 0..8 {
            crc = if crc & 1 == 0 {
                crc >> 1
            } else {
                (crc >> 1) ^ 0xedb8_8320
            };
        }
    }
    !crc
}

/// Prove that the ZIP envelope outside the selected member's mutable bytes
/// remains source-backed. Entry-level checks cover member records; this
/// boundary check covers bytes that do not belong to an Entry API, including
/// the prelude and end-of-central-directory tail.
fn verify_zip_boundaries(
    before: &SourceCatalog,
    after: &SourceCatalog,
    budget: &mut TransactionBudget,
) -> Result<(), Error> {
    let before_bytes = before.source_bytes();
    let after_bytes = after.source_bytes();
    let before_first = before.package().iter().next().ok_or(Error::Verification)?;
    let after_first = after.package().iter().next().ok_or(Error::Verification)?;
    let before_prelude_end =
        raw_slice_offset(before_bytes, before_first.raw_record().local_record())
            .ok_or(Error::Verification)?;
    let after_prelude_end = raw_slice_offset(after_bytes, after_first.raw_record().local_record())
        .ok_or(Error::Verification)?;
    budget.charge_work(
        before_prelude_end
            .checked_add(after_prelude_end)
            .ok_or(Error::Verification)?,
    )?;
    if before_bytes.get(..before_prelude_end) != after_bytes.get(..after_prelude_end) {
        return Err(Error::Verification);
    }

    let before_tail = eocd_tail(before_bytes).ok_or(Error::Verification)?;
    let after_tail = eocd_tail(after_bytes).ok_or(Error::Verification)?;
    budget.charge_work(
        before_tail
            .len()
            .checked_add(after_tail.len())
            .ok_or(Error::Verification)?,
    )?;
    if before_tail.len() != after_tail.len()
        || before_tail[..16] != after_tail[..16]
        || before_tail[20..] != after_tail[20..]
    {
        return Err(Error::Verification);
    }
    Ok(())
}

fn raw_slice_offset(source: &[u8], raw: &[u8]) -> Option<usize> {
    if raw.is_empty() {
        return Some(0);
    }
    source
        .windows(raw.len())
        .enumerate()
        .find_map(|(offset, slice)| std::ptr::eq(slice.as_ptr(), raw.as_ptr()).then_some(offset))
}

fn eocd_tail(source: &[u8]) -> Option<&[u8]> {
    source
        .windows(22)
        .enumerate()
        .rev()
        .find_map(|(offset, fixed)| {
            (fixed.get(..4) == Some(b"PK\x05\x06"))
                .then(|| {
                    let comment_length =
                        usize::from(u16::from_le_bytes(fixed.get(20..22)?.try_into().ok()?));
                    let end = offset.checked_add(22)?.checked_add(comment_length)?;
                    (end == source.len()).then(|| &source[offset..])
                })
                .flatten()
        })
}

fn preflight_rewrite_output(catalog: &SourceCatalog) -> Result<(), Error> {
    let limits = catalog.limits();
    let replacement_bound = limits
        .max_iwa_stream_bytes()
        .checked_add(32)
        .ok_or(Error::InvalidSource)?;
    let output_bound =
        archive_output_upper_bound(catalog.shared_source().len(), replacement_bound)?;
    check_archive_output_bound(output_bound, limits)
}

fn charge_output_upper_bound(
    source_bytes: usize,
    replacement_bytes: usize,
    limits: Limits,
    budget: &mut TransactionBudget,
) -> Result<usize, Error> {
    let output_bound = archive_output_upper_bound(source_bytes, replacement_bytes)?;
    check_archive_output_bound(output_bound, limits)?;
    budget.charge_work(output_bound)?;
    Ok(output_bound)
}

fn archive_output_upper_bound(
    source_bytes: usize,
    replacement_bytes: usize,
) -> Result<usize, Error> {
    // ZIP reassembly may re-encode the replacement member (for example with
    // Deflate), so the logical replacement length is not itself a physical
    // output bound. Keep a conservative checked bound before the reassembly
    // output buffer or candidate package is allocated.
    let replacement_bound = replacement_bytes
        .checked_mul(2)
        .and_then(|amount| amount.checked_add(1_024))
        .ok_or(Error::InvalidSource)?;
    let output_bound = source_bytes
        .checked_add(replacement_bound)
        .ok_or(Error::InvalidSource)?;
    Ok(output_bound)
}

fn check_archive_output_bound(output_bound: usize, limits: Limits) -> Result<(), Error> {
    let observed = u64::try_from(output_bound).map_err(|_| Error::InvalidSource)?;
    if observed > limits.max_input_bytes() {
        return Err(Error::LimitExceeded {
            kind: LimitKind::OutputBytes,
            observed,
            maximum: limits.max_input_bytes(),
        });
    }
    Ok(())
}

fn message_info_preserved_except_soundtrack_media_data_references(
    before: &litchi_iwa_core::MessageInfo,
    after: &litchi_iwa_core::MessageInfo,
    expected_media: &[u64],
) -> bool {
    before.type_ == after.type_
        && before.versions == after.versions
        && before.length == after.length
        && before.field_infos.len() == after.field_infos.len()
        && before.object_references == after.object_references
        && before.base_message_index == after.base_message_index
        && before.diff_merge_version == after.diff_merge_version
        && before.diff_field_path == after.diff_field_path
        && before.fields_to_remove == after.fields_to_remove
        && before.diff_read_version == after.diff_read_version
        && before
            .field_infos
            .iter()
            .zip(&after.field_infos)
            .all(|(before_field, after_field)| {
                before_field.path == after_field.path
                    && before_field.r#type == after_field.r#type
                    && before_field.unknown_field_rule == after_field.unknown_field_rule
                    && before_field.object_references == after_field.object_references
                    && if before_field.path.as_slice() == [SOUNDTRACK_MEDIA_FIELD] {
                        after_field.data_references == expected_media
                    } else {
                        before_field.data_references == after_field.data_references
                    }
                    && before_field.known_field_rule == after_field.known_field_rule
                    && before_field.known_field_version == after_field.known_field_version
                    && before_field.known_field_feature_identifier
                        == after_field.known_field_feature_identifier
            })
}

fn strict_optional_reference(
    source: &[u8],
    field_number: u32,
    limits: WireLimits,
    budget: &mut TransactionBudget,
) -> Result<Option<ReferenceFacts>, Error> {
    budget.require_depth(1)?;
    preflight_wire_message(source, limits, budget)?;
    budget.charge_work(source.len())?;
    let bounded_limits = remaining_wire_limits(limits, budget)?;
    let view = WireView::parse_with_limits(source, bounded_limits).map_err(map_wire_error)?;
    budget.charge_fields(view.len())?;
    let mut result = None;
    for field in view.fields() {
        if field.number() != field_number {
            continue;
        }
        if result.is_some() || field.wire_type() != LENGTH_DELIMITED_WIRE_TYPE {
            return Err(Error::InvalidSource);
        }
        field.validate_canonical_framing().map_err(map_wire_error)?;
        result = Some(strict_reference(field.payload(), limits, budget)?);
    }
    Ok(result)
}

fn strict_reference(
    source: &[u8],
    limits: WireLimits,
    budget: &mut TransactionBudget,
) -> Result<ReferenceFacts, Error> {
    // The native codec counts the containing soundtrack message as depth one
    // and this reference as depth two.  Check that depth before creating the
    // reference's span vector.
    budget.require_depth(2)?;
    preflight_wire_message(source, limits, budget)?;
    budget.charge_work(source.len())?;
    let bounded_limits = remaining_wire_limits(limits, budget)?;
    let view = WireView::parse_with_limits(source, bounded_limits).map_err(map_wire_error)?;
    budget.charge_fields(view.len())?;
    let mut identifier = None;
    let mut deprecated_type = None;
    let mut external = None;
    for field in view.fields() {
        field.validate_canonical_key().map_err(map_wire_error)?;
        field.validate_canonical_framing().map_err(map_wire_error)?;
        if field.wire_type() != 0 {
            if matches!(field.number(), 1..=3) {
                return Err(Error::InvalidSource);
            }
            continue;
        }
        let value = canonical_varint(field.payload())?;
        match field.number() {
            1 => {
                if identifier.is_some() || value == 0 {
                    return Err(Error::InvalidSource);
                }
                identifier = Some(value);
            },
            2 => {
                if deprecated_type.is_some() || !canonical_int32(value) {
                    return Err(Error::InvalidSource);
                }
                deprecated_type = Some(value);
            },
            3 => {
                if external.is_some() {
                    return Err(Error::InvalidSource);
                }
                external = Some(match value {
                    0 => false,
                    1 => true,
                    _ => return Err(Error::InvalidSource),
                });
            },
            _ => {},
        }
    }
    // Zero is not an object/data identifier.  Reject it here so every
    // reference role (including media records) receives the same strict
    // validation instead of relying on one caller-specific range check.
    let identifier = identifier
        .filter(|identifier| *identifier != 0)
        .ok_or(Error::InvalidSource)?;
    let _ = deprecated_type;
    budget.charge_references(1)?;
    Ok(ReferenceFacts {
        identifier,
        external,
    })
}

fn canonical_varint(source: &[u8]) -> Result<u64, Error> {
    let (value, consumed) = decode_varint_from_bytes(source).map_err(|_| Error::InvalidSource)?;
    if consumed != source.len() || consumed != encoded_len(value) {
        return Err(Error::InvalidSource);
    }
    Ok(value)
}

fn canonical_int32(value: u64) -> bool {
    value <= i64::from(i32::MAX) as u64 || value >= 0xffff_ffff_8000_0000
}

fn permute_payload(
    source: &[u8],
    intent: Intent,
    item_count: usize,
    limits: WireLimits,
    budget: &mut TransactionBudget,
) -> Result<Vec<u8>, Error> {
    if source.len() > limits.max_input_bytes() {
        return Err(Error::LimitExceeded {
            kind: LimitKind::WireBytes,
            observed: source.len() as u64,
            maximum: limits.max_input_bytes() as u64,
        });
    }
    // The permutation is length-preserving, so the final payload can never
    // be smaller than this source buffer. Check the independent output
    // ceiling before retaining the item records or reserving the destination
    // buffer; otherwise a tight output profile would still permit a
    // source-sized allocation before returning the same limit error below.
    check_wire_output_bound(source.len(), limits)?;
    // Prove the complete root-plus-media wire shape before retaining the
    // parsed span vector, the media-record references, or the destination
    // payload buffer.  The final parse below remains responsible for charging
    // the actual rewrite work.
    preflight_soundtrack_payload(source, limits, budget)?;
    budget.charge_work(source.len())?;
    let view = WireView::parse_with_limits(source, remaining_wire_limits(limits, budget)?)
        .map_err(map_wire_error)?;
    budget.charge_fields(view.len())?;
    budget.charge_work(
        source
            .len()
            .checked_add(item_count)
            .ok_or(Error::InvalidSource)?,
    )?;
    budget.charge_references(item_count)?;
    let mut media = Vec::new();
    media
        .try_reserve_exact(item_count)
        .map_err(|_| Error::Allocation { amount: item_count })?;
    for field in view.fields() {
        if field.number() != SOUNDTRACK_MEDIA_FIELD {
            continue;
        }
        if field.wire_type() != LENGTH_DELIMITED_WIRE_TYPE {
            return Err(Error::InvalidSource);
        }
        field
            .validate_canonical_framing()
            .map_err(|_| Error::InvalidSource)?;
        let reference = strict_reference(field.payload(), limits, budget)?;
        if reference.is_external() {
            return Err(Error::InvalidSource);
        }
        media.push(field.raw());
    }
    if media.len() != item_count {
        return Err(Error::InvalidSource);
    }
    let mut output = Vec::new();
    output
        .try_reserve_exact(source.len())
        .map_err(|_| Error::Allocation {
            amount: source.len(),
        })?;
    budget.charge_work(source.len())?;
    let mut slot = 0usize;
    for field in view.fields() {
        if field.number() == SOUNDTRACK_MEDIA_FIELD {
            let source_slot = source_slot_for_destination(slot, intent);
            output.extend_from_slice(
                media
                    .get(source_slot)
                    .copied()
                    .ok_or(Error::InvalidSource)?,
            );
            slot += 1;
        } else {
            output.extend_from_slice(field.raw());
        }
    }
    if slot != media.len() || output.len() != source.len() {
        return Err(Error::Verification);
    }
    check_wire_output_bound(output.len(), limits)?;
    Ok(output)
}

fn rewrite(
    source: &Package,
    catalog: &SourceCatalog,
    selection: &Selection<'_>,
    before: &[u64],
    after: &[u64],
    intent: Intent,
    budget: &mut TransactionBudget,
) -> Result<Package, Error> {
    if selection.soundtrack_identifier == 0 || before.len() != after.len() {
        return Err(Error::InvalidSource);
    }
    let entry = catalog
        .package()
        .iter()
        .find(|entry| entry.name() == selection.soundtrack_component)
        .ok_or(Error::InvalidSource)?;
    if entry.is_opaque() {
        return Err(Error::InvalidSource);
    }
    let physical_limits = catalog.limits();
    let archive_limits = physical_limits
        .effective_archive_limits()
        .map_err(map_archive_error)?;
    // The decoded component is bounded by the physical profile even before
    // its Snappy buffer exists.  Check the conservative package-output bound
    // at this point so a tight output profile fails before decompression,
    // archive parsing, or any rewrite buffer is allocated.
    preflight_rewrite_output(catalog)?;
    let snappy_limits = transaction_snappy_limits(source, budget, entry.data().len(), 2)?;
    let stream = SnappyStream::decompress_with_limits(entry.data(), snappy_limits)
        .map_err(|error| map_transaction_snappy_error(error, budget))?;
    let stream_bytes = stream.as_bytes();
    budget.charge_work(
        entry
            .data()
            .len()
            .checked_add(
                stream_bytes
                    .len()
                    .checked_mul(2)
                    .ok_or(Error::InvalidSource)?,
            )
            .ok_or(Error::InvalidSource)?,
    )?;
    let mut archive =
        Archive::parse_with_limits(stream_bytes, archive_limits).map_err(map_core_error)?;
    archive
        .validate_canonical_object_framing(stream_bytes)
        .map_err(map_core_error)?;
    let object = archive
        .object(selection.soundtrack_identifier)
        .ok_or(Error::InvalidSource)?;
    let (message_index, payload) = selected_message(object, SOUNDTRACK_MESSAGE_TYPE)?;
    if message_index != selection.soundtrack_message_index
        || payload != selection.soundtrack_payload
    {
        return Err(Error::InvalidSource);
    }
    let serialized_bound = stream_bytes
        .len()
        .checked_add(32)
        .ok_or(Error::InvalidSource)?;
    // Reordering preserves every payload and archive-info length.  The small
    // allowance covers a defensive header-length change, while the output
    // bound itself covers worst-case Snappy expansion.  Charge this before
    // retaining media records, serializing the candidate archive, or
    // allocating the reassembled package output.
    let output_bound = charge_output_upper_bound(
        catalog.shared_source().len(),
        serialized_bound,
        physical_limits,
        budget,
    )?;
    let wire_limits = source.wire_limits().map_err(map_wire_error)?;
    let rewritten = permute_payload(payload, intent, before.len(), wire_limits, budget)?;
    budget.charge_work(
        serialized_bound
            .checked_mul(2)
            .ok_or(Error::InvalidSource)?,
    )?;
    let object = archive
        .object_mut(selection.soundtrack_identifier)
        .ok_or(Error::InvalidSource)?;
    object
        .replace_message_reordering_data_references_preserving_header_with_limits(
            message_index,
            RawMessage {
                type_: SOUNDTRACK_MESSAGE_TYPE,
                data: rewritten,
            },
            after,
            archive_limits,
        )
        .map_err(map_core_error)?;
    let serialized = archive
        .to_bytes_with_limits(archive_limits)
        .map_err(map_core_error)?;
    if serialized.len() > serialized_bound {
        return Err(Error::Verification);
    }
    let compressed = SnappyStream::compress(&serialized).map_err(map_core_error)?;
    let output = catalog
        .package()
        .reassemble_to_bytes(
            &[EntryEdit::new(selection.soundtrack_component, &compressed)],
            physical_limits,
        )
        .map_err(map_archive_error)?;
    let target: Arc<[u8]> = output.into();
    if target.len() > output_bound {
        return Err(Error::Verification);
    }
    charge_catalog_reopen_cost(
        catalog,
        target.len(),
        Some((
            selection.soundtrack_component,
            compressed.len(),
            serialized.len(),
        )),
        budget,
    )?;
    let candidate =
        Package::from_source_with_options(target, source.state.options).map_err(map_read_error)?;
    Ok(candidate)
}

fn transaction_snappy_limits(
    package: &Package,
    budget: &TransactionBudget,
    compressed_bytes: usize,
    decoded_passes: usize,
) -> Result<litchi_iwa_core::SnappyLimits, Error> {
    let base = package
        .limits()
        .snappy_limits()
        .map_err(map_archive_error)?;
    let remaining_work = budget.remaining_work();
    let decoded_allowance = remaining_work
        .checked_sub(compressed_bytes)
        .map_or(0, |available| available / decoded_passes);
    if decoded_allowance == 0 {
        return Err(Error::LimitExceeded {
            kind: LimitKind::WireWork,
            observed: budget
                .work
                .saturating_add(compressed_bytes)
                .saturating_add(decoded_passes) as u64,
            maximum: budget.max_work as u64,
        });
    }
    litchi_iwa_core::SnappyLimits::new(
        base.max_uncompressed_chunk().min(decoded_allowance),
        decoded_allowance,
    )
    .and_then(|limits| {
        limits.with_input_limits(
            base.max_compressed_chunk(),
            base.max_compressed_stream(),
            base.max_frames(),
        )
    })
    .map_err(map_core_error)
}

fn map_transaction_snappy_error(
    error: litchi_iwa_core::Error,
    budget: &TransactionBudget,
) -> Error {
    match error {
        litchi_iwa_core::Error::Limit {
            kind:
                litchi_iwa_core::LimitKind::SnappyChunkBytes
                | litchi_iwa_core::LimitKind::SnappyStreamBytes,
            observed,
            ..
        } => Error::LimitExceeded {
            kind: LimitKind::WireWork,
            observed: budget.work.saturating_add(observed) as u64,
            maximum: budget.max_work as u64,
        },
        other => map_core_error(other),
    }
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

fn moved_ids(before: &[u64], intent: Intent) -> Result<Vec<u64>, Error> {
    validate_intent(intent, before.len())?;
    let mut after = Vec::new();
    after
        .try_reserve_exact(before.len())
        .map_err(|_| Error::Allocation {
            amount: before.len(),
        })?;
    after.extend_from_slice(before);
    let value = after.remove(intent.source.get());
    after.insert(intent.destination.get(), value);
    Ok(after)
}

fn validate_intent(intent: Intent, item_count: usize) -> Result<(), Error> {
    if intent.source.get() >= item_count {
        return Err(Error::SourcePositionNotFound {
            position: intent.source,
        });
    }
    if intent.destination.get() >= item_count {
        return Err(Error::DestinationOutOfRange {
            position: intent.destination,
            item_count,
        });
    }
    Ok(())
}

fn remaining_wire_limits(
    limits: WireLimits,
    budget: &TransactionBudget,
) -> Result<WireLimits, Error> {
    if budget.remaining_fields() == 0 {
        return Err(Error::LimitExceeded {
            kind: LimitKind::WireFields,
            observed: budget.max_fields.saturating_add(1) as u64,
            maximum: budget.max_fields as u64,
        });
    }
    if budget.remaining_work() == 0 {
        return Err(Error::LimitExceeded {
            kind: LimitKind::WireWork,
            observed: budget.max_work.saturating_add(1) as u64,
            maximum: budget.max_work as u64,
        });
    }
    limits
        .with_fields(budget.remaining_fields())
        .and_then(|bounded| bounded.with_rewrite_work(budget.remaining_work()))
        .map_err(map_wire_error)
}

fn physical_source(package: &Package) -> Result<&SourceCatalog, Error> {
    match &package.state.source {
        PhysicalSource::Package(source) => Ok(source),
        PhysicalSource::Semantic(_) => Err(Error::UnsupportedSource),
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
        } => Error::LimitExceeded {
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
            observed: observed as u64,
            maximum: limit as u64,
        },
        litchi_iwa_common::Error::Allocation { amount, .. } => Error::Allocation { amount },
        litchi_iwa_common::Error::InvalidFormat(_)
        | litchi_iwa_common::Error::InvalidLimit { .. } => Error::InvalidSource,
    }
}

#[allow(
    clippy::needless_pass_by_value,
    reason = "used directly as a Result::map_err conversion"
)]
fn map_codec_error(error: soundtrack_codec::DecodeError, budget: &TransactionBudget) -> Error {
    let Some(limit) = error.resource_limit() else {
        return Error::InvalidSource;
    };
    match limit {
        soundtrack_codec::DecodeLimit::Bytes { observed, maximum } => Error::LimitExceeded {
            kind: LimitKind::WireBytes,
            observed: observed as u64,
            maximum: maximum as u64,
        },
        soundtrack_codec::DecodeLimit::Fields { observed, maximum } => Error::LimitExceeded {
            kind: LimitKind::WireFields,
            observed: budget.fields.saturating_add(observed) as u64,
            maximum: budget.fields.saturating_add(maximum) as u64,
        },
        soundtrack_codec::DecodeLimit::Work { observed, maximum } => Error::LimitExceeded {
            kind: LimitKind::WireWork,
            observed: budget.work.saturating_add(observed) as u64,
            maximum: budget.work.saturating_add(maximum) as u64,
        },
        soundtrack_codec::DecodeLimit::Nesting { observed, maximum } => Error::LimitExceeded {
            kind: LimitKind::WireNesting,
            observed: u64::from(observed),
            maximum: u64::from(maximum),
        },
    }
}

#[allow(
    clippy::needless_pass_by_value,
    reason = "used directly as a Result::map_err conversion"
)]
fn map_read_error(error: ReadError) -> Error {
    match error {
        ReadError::SemanticLimit {
            kind,
            observed,
            maximum,
            ..
        } => Error::LimitExceeded {
            kind: match kind {
                super::SemanticLimitKind::Objects => LimitKind::Entries,
                super::SemanticLimitKind::Slides => LimitKind::Slides,
                super::SemanticLimitKind::References => LimitKind::References,
                super::SemanticLimitKind::TextStorages => LimitKind::TextStorages,
                super::SemanticLimitKind::TextFragments => LimitKind::TextFragments,
                super::SemanticLimitKind::TextBytes => LimitKind::TextBytes,
            },
            observed: observed as u64,
            maximum: maximum as u64,
        },
        ReadError::PayloadLimit {
            kind,
            observed,
            maximum,
            ..
        } => Error::LimitExceeded {
            kind: match kind {
                super::PayloadLimitKind::Bytes => LimitKind::WireBytes,
                super::PayloadLimitKind::Fields => LimitKind::WireFields,
                super::PayloadLimitKind::Nesting => LimitKind::WireNesting,
                super::PayloadLimitKind::Work => LimitKind::WireWork,
            },
            observed: observed as u64,
            maximum: maximum as u64,
        },
        ReadError::Allocation { amount, .. } => Error::Allocation { amount },
        ReadError::Archive(archive_error) => map_archive_error(archive_error),
        ReadError::Io(_)
        | ReadError::Detection(_)
        | ReadError::NotKeynote
        | ReadError::InvalidFormat(_)
        | ReadError::Decode(_)
        | ReadError::TextStorage { .. }
        | ReadError::Metadata(_) => Error::InvalidSource,
    }
}

#[allow(
    clippy::needless_pass_by_value,
    reason = "used directly as a Result::map_err conversion"
)]
fn map_archive_error(error: litchi_iwa_archive::Error) -> Error {
    match error {
        litchi_iwa_archive::Error::Limit {
            kind,
            observed,
            maximum,
        } => Error::LimitExceeded {
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
        litchi_iwa_archive::Error::Io(_)
        | litchi_iwa_archive::Error::Zip { .. }
        | litchi_iwa_archive::Error::InvalidLimits(_)
        | litchi_iwa_archive::Error::Encrypted
        | litchi_iwa_archive::Error::SourceChanged { .. }
        | litchi_iwa_archive::Error::DirectoryChanged { .. }
        | litchi_iwa_archive::Error::Reassembly(_)
        | litchi_iwa_archive::Error::InvalidBundle(_) => Error::InvalidSource,
    }
}

#[allow(
    clippy::needless_pass_by_value,
    reason = "used directly as a Result::map_err conversion"
)]
fn map_core_error(error: litchi_iwa_core::Error) -> Error {
    match error {
        litchi_iwa_core::Error::Limit {
            kind,
            observed,
            maximum,
        } => Error::LimitExceeded {
            kind: match kind {
                litchi_iwa_core::LimitKind::Objects
                | litchi_iwa_core::LimitKind::Messages
                | litchi_iwa_core::LimitKind::MessagesPerObject
                | litchi_iwa_core::LimitKind::MetadataItems => LimitKind::Entries,
                litchi_iwa_core::LimitKind::MessageBytes => LimitKind::WireBytes,
                litchi_iwa_core::LimitKind::HeaderFields => LimitKind::WireFields,
                litchi_iwa_core::LimitKind::HeaderNesting => LimitKind::WireNesting,
                litchi_iwa_core::LimitKind::ArchiveBytes
                | litchi_iwa_core::LimitKind::ObjectBytes
                | litchi_iwa_core::LimitKind::HeaderBytes
                | litchi_iwa_core::LimitKind::HeaderMemoryBytes
                | litchi_iwa_core::LimitKind::SnappyChunkBytes
                | litchi_iwa_core::LimitKind::SnappyStreamBytes
                | litchi_iwa_core::LimitKind::SnappyCompressedChunkBytes
                | litchi_iwa_core::LimitKind::SnappyCompressedStreamBytes
                | litchi_iwa_core::LimitKind::SnappyFrames => LimitKind::EntryBytes,
            },
            observed: observed as u64,
            maximum: maximum as u64,
        },
        litchi_iwa_core::Error::Allocation { requested, .. } => {
            Error::Allocation { amount: requested }
        },
        litchi_iwa_core::Error::InvalidArchive { .. }
        | litchi_iwa_core::Error::InvalidLimits { .. }
        | litchi_iwa_core::Error::HeaderCodec { .. }
        | litchi_iwa_core::Error::Io(_)
        | litchi_iwa_core::Error::Snappy { .. } => Error::InvalidSource,
    }
}

#[cfg(test)]
mod tests {
    use std::collections::HashMap;

    use litchi_iwa_common::{WireLimits, encode_varint_into};

    use super::{
        Error, Intent, Limits, MediaClosureState, TransactionBudget, ZipEntryFields,
        bytes_equal_outside_ranges, central_record_preserved_except_offset,
        charge_output_upper_bound, check_wire_output_bound, permute_payload,
        preflight_soundtrack_payload, selected_central_record_preserved,
        selected_local_suffix_preserved, stream_package_metadata,
        validate_soundtrack_media_references,
    };

    const SOUNDTRACK_IDENTIFIER: u64 = 10;
    const DATA_IDENTIFIER: u64 = 42;
    const LOCATOR: &[u8] = b"Document";

    fn budget() -> TransactionBudget {
        TransactionBudget {
            fields: 0,
            work: 0,
            references: 0,
            max_fields: WireLimits::MAX_FIELDS,
            max_work: WireLimits::MAX_REWRITE_WORK,
            max_references: 1_000_000,
            max_nesting: WireLimits::MAX_NESTING,
        }
    }

    #[test]
    fn versioned_component_cannot_claim_selected_media_ownership() {
        let mut owner = Vec::new();
        push_varint_field(&mut owner, 1, SOUNDTRACK_IDENTIFIER);
        push_varint_field(&mut owner, 2, 1);

        let mut data_reference = Vec::new();
        push_varint_field(&mut data_reference, 1, DATA_IDENTIFIER);
        push_bytes_field(&mut data_reference, 2, &owner);

        let mut component = Vec::new();
        push_bytes_field(&mut component, 2, LOCATOR);
        push_bytes_field(&mut component, 7, &data_reference);

        let mut metadata = Vec::new();
        // PackageMetadata.versioned_components, not current components.
        push_bytes_field(&mut metadata, 11, &component);

        let mut states = HashMap::from([(
            DATA_IDENTIFIER,
            MediaClosureState {
                payload_occurrences: 1,
                component_declarations: 0,
                owner_occurrences: 0,
                owner_count: 0,
                data_declarations: 0,
                filename: None,
                materialized_length: None,
            },
        )]);
        let mut transaction_budget = budget();
        assert_eq!(
            stream_package_metadata(
                &metadata,
                LOCATOR,
                SOUNDTRACK_IDENTIFIER,
                &mut states,
                &mut transaction_budget,
            )
            .expect("versioned metadata is parseable"),
            0
        );
        let state = states.get(&DATA_IDENTIFIER).expect("state remains present");
        assert_eq!(state.component_declarations, 0);
        assert_eq!(state.owner_occurrences, 0);
    }

    #[test]
    fn media_reference_external_marker_is_rejected_before_rewrite() {
        let mut reference = Vec::new();
        push_varint_field(&mut reference, 1, DATA_IDENTIFIER);
        push_varint_field(&mut reference, 3, 1);
        let mut payload = Vec::new();
        push_bytes_field(&mut payload, 3, &reference);
        let mut transaction_budget = budget();
        assert!(matches!(
            validate_soundtrack_media_references(
                &payload,
                WireLimits::default(),
                1,
                &mut transaction_budget,
            ),
            Err(Error::InvalidSource)
        ));
    }

    #[test]
    fn output_upper_bound_rejects_one_over_before_budget_charge() {
        let limits = Limits::new(1_063, 1, 1_024, 1_024, 1_024).expect("test limits are valid");
        let mut transaction_budget = budget();
        assert!(matches!(
            charge_output_upper_bound(8, 16, limits, &mut transaction_budget),
            Err(Error::LimitExceeded {
                kind: super::LimitKind::OutputBytes,
                observed: 1_064,
                maximum: 1_063,
            })
        ));
        assert_eq!(transaction_budget.work, 0);
    }

    #[test]
    fn output_upper_bound_reports_checked_arithmetic_overflow() {
        let mut transaction_budget = budget();
        assert!(matches!(
            charge_output_upper_bound(usize::MAX, 1, Limits::default(), &mut transaction_budget),
            Err(Error::InvalidSource)
        ));
        assert_eq!(transaction_budget.work, 0);
    }

    #[test]
    fn selected_raw_zip_delta_ranges_are_explicit() {
        let mut local_before = [0u8; 30];
        local_before[..4].copy_from_slice(b"PK\x03\x04");
        local_before[12] = 0x5a;
        let mut local_after = local_before;
        local_after[14..26].copy_from_slice(&[1; 12]);
        assert!(bytes_equal_outside_ranges(
            &local_before,
            &local_after,
            14..26,
            None,
        ));
        local_after[12] = 0x6b;
        assert!(!bytes_equal_outside_ranges(
            &local_before,
            &local_after,
            14..26,
            None,
        ));

        let mut central_before = [0u8; 46];
        central_before[..4].copy_from_slice(b"PK\x01\x02");
        let mut central_after = central_before;
        central_after[16..28].copy_from_slice(&[2; 12]);
        central_after[42..46].copy_from_slice(&[3; 4]);
        assert!(bytes_equal_outside_ranges(
            &central_before,
            &central_after,
            16..28,
            Some(42..46),
        ));
        central_after[30] = 0x7c;
        assert!(!bytes_equal_outside_ranges(
            &central_before,
            &central_after,
            16..28,
            Some(42..46),
        ));
    }

    #[test]
    fn soundtrack_wire_preflight_rejects_limits_before_budget_usage() {
        let mut reference = Vec::new();
        push_varint_field(&mut reference, 1, DATA_IDENTIFIER);
        let mut payload = Vec::new();
        push_bytes_field(&mut payload, 3, &reference);

        let mut fields = budget();
        fields.max_fields = 1;
        assert!(matches!(
            preflight_soundtrack_payload(&payload, WireLimits::default(), &fields),
            Err(Error::LimitExceeded {
                kind: super::LimitKind::WireFields,
                ..
            })
        ));
        assert_eq!(fields.fields, 0);
        assert_eq!(fields.work, 0);

        let mut work = budget();
        work.max_work = payload.len() - 1;
        assert!(matches!(
            preflight_soundtrack_payload(&payload, WireLimits::default(), &work),
            Err(Error::LimitExceeded {
                kind: super::LimitKind::WireWork,
                ..
            })
        ));
        assert_eq!(work.fields, 0);
        assert_eq!(work.work, 0);

        let mut nesting = budget();
        nesting.max_nesting = 1;
        assert!(matches!(
            preflight_soundtrack_payload(&payload, WireLimits::default(), &nesting),
            Err(Error::LimitExceeded {
                kind: super::LimitKind::WireNesting,
                ..
            })
        ));
        assert_eq!(nesting.fields, 0);
        assert_eq!(nesting.work, 0);
    }

    #[test]
    fn wire_output_preflight_rejects_before_destination_allocation() {
        let transaction_budget = budget();
        assert!(matches!(
            check_wire_output_bound(4, WireLimits::default().with_output_bytes(3).unwrap()),
            Err(Error::LimitExceeded {
                kind: super::LimitKind::OutputBytes,
                observed: 4,
                maximum: 3,
            })
        ));
        assert_eq!(transaction_budget.fields, 0);
        assert_eq!(transaction_budget.work, 0);
    }

    #[test]
    fn permutation_limits_fail_before_any_destination_usage() {
        let mut reference = Vec::new();
        push_varint_field(&mut reference, 1, DATA_IDENTIFIER);
        let mut payload = Vec::new();
        push_bytes_field(&mut payload, 3, &reference);
        let intent = Intent {
            source: litchi_core::Position::new(0),
            destination: litchi_core::Position::new(0),
        };

        let mut fields = budget();
        fields.max_fields = 1;
        assert!(matches!(
            permute_payload(&payload, intent, 1, WireLimits::default(), &mut fields,),
            Err(Error::LimitExceeded {
                kind: super::LimitKind::WireFields,
                ..
            })
        ));
        assert_eq!(fields.fields, 0);
        assert_eq!(fields.work, 0);
        assert_eq!(fields.references, 0);

        let mut work = budget();
        work.max_work = payload.len() - 1;
        assert!(matches!(
            permute_payload(&payload, intent, 1, WireLimits::default(), &mut work,),
            Err(Error::LimitExceeded {
                kind: super::LimitKind::WireWork,
                ..
            })
        ));
        assert_eq!(work.fields, 0);
        assert_eq!(work.work, 0);
        assert_eq!(work.references, 0);

        let mut nesting = budget();
        nesting.max_nesting = 1;
        assert!(matches!(
            permute_payload(&payload, intent, 1, WireLimits::default(), &mut nesting,),
            Err(Error::LimitExceeded {
                kind: super::LimitKind::WireNesting,
                ..
            })
        ));
        assert_eq!(nesting.fields, 0);
        assert_eq!(nesting.work, 0);
        assert_eq!(nesting.references, 0);

        let mut output = budget();
        let output_limits = WireLimits::default()
            .with_output_bytes(payload.len() - 1)
            .expect("payload length is non-zero");
        assert!(matches!(
            permute_payload(&payload, intent, 1, output_limits, &mut output),
            Err(Error::LimitExceeded {
                kind: super::LimitKind::OutputBytes,
                ..
            })
        ));
        assert_eq!(output.fields, 0);
        assert_eq!(output.work, 0);
        assert_eq!(output.references, 0);
    }

    #[test]
    fn selected_zip_suffix_and_central_records_allow_only_reassembly_fields() {
        let unchanged_fields = ZipEntryFields {
            crc32: 1,
            compressed_size: 2,
            uncompressed_size: 3,
        };
        assert!(selected_local_suffix_preserved(
            0,
            b"gap",
            b"gap",
            unchanged_fields,
            unchanged_fields,
        ));
        assert!(!selected_local_suffix_preserved(
            0,
            b"gap",
            b"changed",
            unchanged_fields,
            unchanged_fields,
        ));

        let before_suffix = [
            b'P', b'K', 7, 8, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, b'g', b'a', b'p',
        ];
        let before_fields = ZipEntryFields {
            crc32: 0x0403_0201,
            compressed_size: 0x0807_0605,
            uncompressed_size: 0x0c0b_0a09,
        };
        let after_fields = ZipEntryFields {
            crc32: 0x1817_1615,
            compressed_size: 0x1c1b_1a19,
            uncompressed_size: 0x201f_1e1d,
        };
        let mut after_suffix = before_suffix;
        after_suffix[4..16].copy_from_slice(&[21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32]);
        assert!(selected_local_suffix_preserved(
            0x0008,
            &before_suffix,
            &after_suffix,
            before_fields,
            after_fields,
        ));
        after_suffix[18] = b'X';
        assert!(!selected_local_suffix_preserved(
            0x0008,
            &before_suffix,
            &after_suffix,
            before_fields,
            after_fields,
        ));

        let mut before_central = [0u8; 46];
        before_central[..4].copy_from_slice(b"PK\x01\x02");
        before_central[28..42].copy_from_slice(&[1; 14]);
        put_u32(&mut before_central, 16, 1);
        put_u32(&mut before_central, 20, 2);
        put_u32(&mut before_central, 24, 3);
        put_u32(&mut before_central, 42, 7);
        let mut after_central = before_central;
        put_u32(&mut after_central, 16, 4);
        put_u32(&mut after_central, 20, 5);
        put_u32(&mut after_central, 24, 6);
        assert!(selected_central_record_preserved(
            &before_central,
            &after_central,
            unchanged_fields,
            ZipEntryFields {
                crc32: 4,
                compressed_size: 5,
                uncompressed_size: 6,
            },
            0,
        ));
        after_central[28] = 9;
        assert!(!selected_central_record_preserved(
            &before_central,
            &after_central,
            unchanged_fields,
            ZipEntryFields {
                crc32: 4,
                compressed_size: 5,
                uncompressed_size: 6,
            },
            0,
        ));
        after_central[28] = 1;
        put_u32(&mut after_central, 42, 8);
        assert!(!selected_central_record_preserved(
            &before_central,
            &after_central,
            unchanged_fields,
            ZipEntryFields {
                crc32: 4,
                compressed_size: 5,
                uncompressed_size: 6,
            },
            0,
        ));
        assert!(central_record_preserved_except_offset(
            &before_central,
            &after_central,
            1,
        ));
    }

    #[test]
    fn wire_budget_boundaries_reject_one_over_without_mutating_usage() {
        let mut fields = budget();
        fields.max_fields = 1;
        assert!(fields.charge_fields(1).is_ok());
        assert!(matches!(
            fields.charge_fields(1),
            Err(Error::LimitExceeded {
                kind: super::LimitKind::WireFields,
                observed: 2,
                maximum: 1,
            })
        ));
        assert_eq!(fields.fields, 1);

        let mut work = budget();
        work.max_work = 1;
        assert!(work.charge_work(1).is_ok());
        assert!(matches!(
            work.charge_work(1),
            Err(Error::LimitExceeded {
                kind: super::LimitKind::WireWork,
                observed: 2,
                maximum: 1,
            })
        ));
        assert_eq!(work.work, 1);

        let mut nesting = budget();
        nesting.max_nesting = 1;
        assert!(nesting.require_depth(1).is_ok());
        assert!(matches!(
            nesting.require_depth(2),
            Err(Error::LimitExceeded {
                kind: super::LimitKind::WireNesting,
                observed: 2,
                maximum: 1,
            })
        ));
    }

    fn put_u32(bytes: &mut [u8], start: usize, value: u32) {
        bytes[start..start + 4].copy_from_slice(&value.to_le_bytes());
    }

    fn push_varint_field(output: &mut Vec<u8>, number: u32, value: u64) {
        encode_varint_into(output, u64::from(number) << 3);
        encode_varint_into(output, value);
    }

    fn push_bytes_field(output: &mut Vec<u8>, number: u32, value: &[u8]) {
        encode_varint_into(output, (u64::from(number) << 3) | 2);
        encode_varint_into(
            output,
            u64::try_from(value.len()).expect("test field length fits u64"),
        );
        output.extend_from_slice(value);
    }
}
