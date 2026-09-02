//! Shared physical validation primitives for Keynote soundtrack adapters.
//!
//! The settings, order, and future soundtrack-item adapters all operate on
//! the same native graph.  Keeping the admission rules here gives those
//! adapters one strict implementation of the root → show → soundtrack
//! selection, canonical reference parsing, bounded wire traversal, and
//! package-media closure checks.  This module is intentionally private to the
//! package layer: IDs, archive objects, and wire payloads must not cross the
//! public semantic API.
//!
//! The helpers borrow the already parsed [`Package`] source.  They do not
//! decode or re-encode an entire package and they never treat the physical
//! archive index as a substitute for a validated field-level ownership proof.

#![allow(
    dead_code,
    reason = "Shared soundtrack primitives are consumed by the order, settings, and item adapters as they migrate to this seam."
)]

use std::{
    collections::{HashMap, HashSet},
    fmt,
    ops::Range,
    path::{Component as PathComponent, Path},
};

use litchi_iwa_archive::{Limits, SourceCatalog};
use litchi_iwa_common::{
    WireLimits, decode_varint_from_bytes,
    varint::encoded_len,
    wire::{WireDescent, WirePreflight, WireView, preflight_wire_tree_with_limits},
};
use litchi_iwa_core::{ArchiveObject, FieldType};
use thiserror::Error as ThisError;

use super::{DOCUMENT_MESSAGE_TYPE, Package, PhysicalSource, SHOW_MESSAGE_TYPE};

/// Native Keynote soundtrack object type.
pub(crate) const SOUNDTRACK_MESSAGE_TYPE: u32 = 21;
/// Document → show reference field.
pub(crate) const DOCUMENT_SHOW_FIELD: u32 = 2;
/// Show → soundtrack reference field.
pub(crate) const SHOW_SOUNDTRACK_FIELD: u32 = 17;
/// Soundtrack repeated media-reference field.
pub(crate) const SOUNDTRACK_MEDIA_FIELD: u32 = 3;
/// Package metadata object type.
pub(crate) const PACKAGE_METADATA_MESSAGE_TYPE: u32 = 11_006;
/// Metadata component containing data declarations and owners.
pub(crate) const METADATA_COMPONENT: &str = "Index/Metadata.iwa";

const LENGTH_DELIMITED_WIRE_TYPE: u8 = 2;
const DATA_METADATA_MAP_MESSAGE_TYPE: u32 = 11_015;

/// Mutable ZIP-local-header bytes permitted when a selected member is
/// reassembled.
pub(crate) const ZIP_LOCAL_CRC_AND_SIZES: Range<usize> = 14..26;
/// Mutable ZIP-central-directory bytes permitted when a selected member is
/// reassembled.
pub(crate) const ZIP_CENTRAL_CRC_AND_SIZES: Range<usize> = 16..28;
/// Mutable ZIP-central-directory relative local-header offset.
pub(crate) const ZIP_CENTRAL_LOCAL_OFFSET: Range<usize> = 42..46;

/// Resource category charged by a shared physical soundtrack transaction.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub(crate) enum LimitKind {
    /// Input wire or package bytes.
    InputBytes,
    /// Output package bytes.
    OutputBytes,
    /// Retained package entries/components.
    Entries,
    /// One encoded entry or archive record.
    EntryBytes,
    /// Aggregate retained package bytes.
    TotalBytes,
    /// Object references traversed in archive metadata.
    References,
    /// Soundtrack item records traversed.
    Items,
    /// Parsed wire bytes.
    WireBytes,
    /// Parsed wire fields.
    WireFields,
    /// Wire nesting depth.
    WireNesting,
    /// Aggregate wire traversal and rewrite work.
    WireWork,
}

impl fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::InputBytes => "input bytes",
            Self::OutputBytes => "output bytes",
            Self::Entries => "entries",
            Self::EntryBytes => "entry bytes",
            Self::TotalBytes => "total bytes",
            Self::References => "references",
            Self::Items => "soundtrack items",
            Self::WireBytes => "wire bytes",
            Self::WireFields => "wire fields",
            Self::WireNesting => "wire nesting depth",
            Self::WireWork => "wire work",
        })
    }
}

/// Content-free failure returned by shared physical validation.
#[derive(Debug, ThisError)]
#[non_exhaustive]
pub(crate) enum Error {
    /// The package was produced from a semantic component snapshot and has no
    /// exact ZIP source that can be physically edited.
    #[error("this Keynote source does not support physical soundtrack operations")]
    UnsupportedSource,
    /// The selected presentation graph or wire metadata is unsafe to edit.
    #[error("the Keynote soundtrack graph is not safely editable")]
    InvalidSource,
    /// The presentation has no soundtrack object.
    #[error("the Keynote presentation has no soundtrack")]
    SoundtrackNotFound,
    /// A finite resource ceiling was exceeded.
    #[error(
        "Keynote soundtrack physical {kind} limit exceeded: observed {observed}, maximum {maximum}"
    )]
    LimitExceeded {
        /// Resource category that exceeded its ceiling.
        kind: LimitKind,
        /// Checked observed amount.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
    },
    /// A bounded vector/map allocation failed before publication.
    #[error("could not allocate {amount} soundtrack physical units")]
    Allocation {
        /// Number of elements or bytes requested.
        amount: usize,
    },
}

/// Whether selection is used only for a read or as a rewrite admission gate.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum SelectionPolicy {
    /// Resolve canonical graph and payload references without requiring media
    /// ownership metadata; suitable for a read-only adapter.
    Read,
    /// Require selected metadata, role disjointness, and complete media
    /// closure before a mutating adapter may stage a rewrite.
    Rewrite,
}

/// Operation-local bounded accounting shared by physical soundtrack adapters.
///
/// All checked charges happen before the corresponding retained view,
/// collection, or output buffer is allocated by callers.  The counters are
/// deliberately independent from the generic package read budget so a
/// settings transaction cannot consume a later order or item transaction's
/// allowance.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct Budget {
    fields: usize,
    work: usize,
    references: usize,
    items: usize,
    max_fields: usize,
    max_work: usize,
    max_references: usize,
    max_items: usize,
    max_nesting: usize,
}

impl Budget {
    /// Build an operation-local budget from the package's already validated
    /// wire and semantic profiles.
    pub(crate) fn new(package: &Package) -> Result<Self, Error> {
        let limits = package.wire_limits().map_err(map_wire_error)?;
        let max_references = package.semantic_limits().max_references();
        Ok(Self {
            fields: 0,
            work: 0,
            references: 0,
            items: 0,
            max_fields: limits.max_fields(),
            max_work: limits.max_rewrite_work(),
            max_references,
            max_items: max_references,
            max_nesting: limits.max_nesting(),
        })
    }

    /// Construct a test/adaptor budget from explicit ceilings.
    pub(crate) const fn from_limits(
        max_fields: usize,
        max_work: usize,
        max_references: usize,
        max_nesting: usize,
    ) -> Self {
        Self {
            fields: 0,
            work: 0,
            references: 0,
            items: 0,
            max_fields,
            max_work,
            max_references,
            max_items: max_references,
            max_nesting,
        }
    }

    /// Number of fields already charged.
    #[must_use]
    pub(crate) const fn fields(self) -> usize {
        self.fields
    }

    /// Aggregate work already charged.
    #[must_use]
    pub(crate) const fn work(self) -> usize {
        self.work
    }

    /// Number of references already charged.
    #[must_use]
    pub(crate) const fn references(self) -> usize {
        self.references
    }

    /// Number of soundtrack items already charged.
    #[must_use]
    pub(crate) const fn items(self) -> usize {
        self.items
    }

    /// Remaining field allowance.
    #[must_use]
    pub(crate) const fn remaining_fields(self) -> usize {
        self.max_fields.saturating_sub(self.fields)
    }

    /// Remaining aggregate work allowance.
    #[must_use]
    pub(crate) const fn remaining_work(self) -> usize {
        self.max_work.saturating_sub(self.work)
    }

    /// Remaining reference allowance.
    #[must_use]
    pub(crate) const fn remaining_references(self) -> usize {
        self.max_references.saturating_sub(self.references)
    }

    /// Charge parsed fields before retaining their span view.
    pub(crate) fn charge_fields(&mut self, amount: usize) -> Result<(), Error> {
        self.fields = checked_charge(self.fields, amount, self.max_fields, LimitKind::WireFields)?;
        Ok(())
    }

    /// Charge aggregate work, including source scans and rewrite operations.
    pub(crate) fn charge_work(&mut self, amount: usize) -> Result<(), Error> {
        self.work = checked_charge(self.work, amount, self.max_work, LimitKind::WireWork)?;
        Ok(())
    }

    /// Charge graph references.
    pub(crate) fn charge_references(&mut self, amount: usize) -> Result<(), Error> {
        self.references = checked_charge(
            self.references,
            amount,
            self.max_references,
            LimitKind::References,
        )?;
        Ok(())
    }

    /// Charge soundtrack item records independently from generic graph refs.
    pub(crate) fn charge_items(&mut self, amount: usize) -> Result<(), Error> {
        self.items = checked_charge(self.items, amount, self.max_items, LimitKind::Items)?;
        Ok(())
    }

    /// Require a native-message depth before creating a nested view.
    pub(crate) fn require_depth(self, depth: usize) -> Result<(), Error> {
        if depth > self.max_nesting {
            return Err(Error::LimitExceeded {
                kind: LimitKind::WireNesting,
                observed: depth as u64,
                maximum: self.max_nesting as u64,
            });
        }
        Ok(())
    }

    /// Charge a common preflight report after the caller has admitted its
    /// shape and before retaining the corresponding `WireView`.
    pub(crate) fn merge_preflight(&mut self, report: WirePreflight) -> Result<(), Error> {
        self.charge_fields(report.fields())?;
        self.charge_work(report.scanned_bytes())?;
        self.require_depth(
            report
                .max_depth()
                .checked_add(1)
                .ok_or(Error::InvalidSource)?,
        )
    }
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

/// Borrowed facts from one strict native `TSP.Reference` message.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) struct ReferenceFacts {
    /// Native object/data identifier, guaranteed nonzero.
    pub(crate) identifier: u64,
    /// Legacy external marker when it was explicitly encoded.
    pub(crate) external: Option<bool>,
}

impl ReferenceFacts {
    /// Return whether the native reference explicitly points outside the
    /// in-package object/data graph.
    #[must_use]
    pub(crate) const fn is_external(self) -> bool {
        matches!(self.external, Some(true))
    }
}

/// One selected root → show → soundtrack graph with all payload slices tied
/// to the same immutable package source.
#[derive(Debug)]
pub(crate) struct Selection<'a> {
    /// Normalized component containing the document root object.
    pub(crate) root_component: &'a str,
    /// Normalized component containing the selected show object.
    pub(crate) show_component: &'a str,
    /// Normalized component containing the soundtrack object.
    pub(crate) soundtrack_component: &'a str,
    /// Root show identifier (kept crate-private for wire adapters).
    pub(crate) show_identifier: u64,
    /// Selected soundtrack identifier.
    pub(crate) soundtrack_identifier: u64,
    /// Root object and selected message index.
    pub(crate) root: &'a ArchiveObject,
    pub(crate) root_message_index: usize,
    /// Show object and selected message index.
    pub(crate) show: &'a ArchiveObject,
    pub(crate) show_message_index: usize,
    /// Soundtrack object, selected message index, and source payload.
    pub(crate) soundtrack: &'a ArchiveObject,
    pub(crate) soundtrack_message_index: usize,
    pub(crate) soundtrack_payload: &'a [u8],
    /// Aggregate data-reference list paired with the payload's repeated
    /// field-3 records.
    pub(crate) media: &'a [u64],
}

impl Selection<'_> {
    /// Number of selected soundtrack item references.
    #[must_use]
    pub(crate) const fn item_count(&self) -> usize {
        self.media.len()
    }

    /// Borrow the source aggregate reference list without copying IDs.
    #[must_use]
    pub(crate) const fn media_ids(&self) -> &[u64] {
        self.media
    }
}

/// Resolve the one strict physical root → show → soundtrack chain.
pub(crate) fn select_soundtrack<'a>(
    package: &'a Package,
    budget: &mut Budget,
    policy: SelectionPolicy,
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
        strict_optional_reference(root_payload, DOCUMENT_SHOW_FIELD, limits, budget)?
            .ok_or(Error::InvalidSource)?;
    if root_reference.is_external()
        || root_reference.identifier == 0
        || root_reference.identifier == 1
    {
        return Err(Error::InvalidSource);
    }
    let show_identifier = root_reference.identifier;
    let (show_component, show) = unique_object(catalog, show_identifier)?;
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
    let info = soundtrack
        .archive_info
        .message_infos
        .get(soundtrack_message_index)
        .ok_or(Error::InvalidSource)?;
    // Item count is charged once at selection admission. The closure scan
    // below only validates ownership and package records; it must not charge
    // the same logical item list a second time when an adapter composes both
    // checks.
    budget.charge_items(info.data_references.len())?;
    budget.charge_references(info.data_references.len())?;
    validate_soundtrack_media_references(
        soundtrack_payload,
        limits,
        &info.data_references,
        budget,
    )?;
    let selection = Selection {
        root_component: root_component.name(),
        show_component,
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
    if policy == SelectionPolicy::Rewrite {
        validate_selected_metadata(root, root_message_index)?;
        validate_selected_metadata(show, show_message_index)?;
        validate_selected_metadata(soundtrack, soundtrack_message_index)?;
        validate_unmodified_object_reference_metadata(
            root,
            root_message_index,
            show_identifier,
            DOCUMENT_SHOW_FIELD,
        )?;
        validate_unmodified_object_reference_metadata(
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

/// Return a physically backed package source or fail closed for semantic-only
/// snapshots.
pub(crate) fn physical_source(package: &Package) -> Result<&SourceCatalog, Error> {
    match &package.state.source {
        PhysicalSource::Package(source) => Ok(source),
        PhysicalSource::Semantic(_) => Err(Error::UnsupportedSource),
    }
}

fn unique_object(
    catalog: &SourceCatalog,
    identifier: u64,
) -> Result<(&str, &ArchiveObject), Error> {
    if identifier == 0 {
        return Err(Error::InvalidSource);
    }
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

/// Validate one object's message/header pairing and select a unique message
/// of the requested native type.
pub(crate) fn selected_message(object: &ArchiveObject, kind: u32) -> Result<(usize, &[u8]), Error> {
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

/// Strictly parse one optional length-delimited `TSP.Reference` field.
pub(crate) fn strict_optional_reference(
    source: &[u8],
    field_number: u32,
    limits: WireLimits,
    budget: &mut Budget,
) -> Result<Option<ReferenceFacts>, Error> {
    budget.require_depth(1)?;
    let _ = preflight_wire_message(source, limits, budget)?;
    budget.charge_work(source.len())?;
    let bounded = remaining_wire_limits(limits, budget)?;
    let view = WireView::parse_with_limits(source, bounded).map_err(map_wire_error)?;
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

/// Strictly parse one native reference.  IDs are nonzero, known scalar fields
/// are unique and canonical, and an explicit external marker is retained for
/// callers to reject according to graph role.
pub(crate) fn strict_reference(
    source: &[u8],
    limits: WireLimits,
    budget: &mut Budget,
) -> Result<ReferenceFacts, Error> {
    budget.require_depth(2)?;
    let _ = preflight_wire_message(source, limits, budget)?;
    budget.charge_work(source.len())?;
    let bounded = remaining_wire_limits(limits, budget)?;
    let view = WireView::parse_with_limits(source, bounded).map_err(map_wire_error)?;
    budget.charge_fields(view.len())?;
    let mut identifier = None;
    let mut deprecated_type = None;
    let mut external = None;
    for field in view.fields() {
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
    let identifier = identifier.ok_or(Error::InvalidSource)?;
    let _ = deprecated_type;
    budget.charge_references(1)?;
    Ok(ReferenceFacts {
        identifier,
        external,
    })
}

/// Preflight a message whose selected descendants are not schema-known.
pub(crate) fn preflight_wire_message(
    source: &[u8],
    limits: WireLimits,
    budget: &Budget,
) -> Result<WirePreflight, Error> {
    let bounded = remaining_wire_limits(limits, budget)?;
    let report = preflight_wire_tree_with_limits(source, bounded, |_visit| Ok(WireDescent::Skip))
        .map_err(map_wire_error)?;
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
    budget.require_depth(
        report
            .max_depth()
            .checked_add(1)
            .ok_or(Error::InvalidSource)?,
    )?;
    Ok(report)
}

/// Preflight a soundtrack message and descend only through repeated field-3
/// reference records.  Unknown length-delimited fields stay opaque and are
/// therefore preserved byte-for-byte by operation-specific rewriters.
pub(crate) fn preflight_soundtrack_payload(
    source: &[u8],
    limits: WireLimits,
    budget: &Budget,
) -> Result<WirePreflight, Error> {
    let bounded = remaining_wire_limits(limits, budget)?;
    let report = preflight_wire_tree_with_limits(source, bounded, |visit| {
        let field = visit.field();
        if matches!(field.wire_type(), 3 | 4) {
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
    if report.fields() > budget.remaining_fields()
        || report.scanned_bytes() > budget.remaining_work()
    {
        let fields = budget
            .fields
            .checked_add(report.fields())
            .ok_or(Error::InvalidSource)?;
        if fields > budget.max_fields {
            return Err(Error::LimitExceeded {
                kind: LimitKind::WireFields,
                observed: fields as u64,
                maximum: budget.max_fields as u64,
            });
        }
        let work = budget
            .work
            .checked_add(report.scanned_bytes())
            .ok_or(Error::InvalidSource)?;
        return Err(Error::LimitExceeded {
            kind: LimitKind::WireWork,
            observed: work as u64,
            maximum: budget.max_work as u64,
        });
    }
    // The preflight root is depth zero; Buffa/native message adapters expose
    // the containing soundtrack as depth one and each reference as depth two.
    budget.require_depth(
        report
            .max_depth()
            .checked_add(1)
            .ok_or(Error::InvalidSource)?,
    )?;
    Ok(report)
}

/// Validate every repeated soundtrack media-reference record and reconcile it
/// with the archive-info aggregate list.
pub(crate) fn validate_soundtrack_media_references(
    source: &[u8],
    limits: WireLimits,
    expected: &[u64],
    budget: &mut Budget,
) -> Result<(), Error> {
    let _ = preflight_soundtrack_payload(source, limits, budget)?;
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
        if reference.is_external() || expected.get(observed).copied() != Some(reference.identifier)
        {
            return Err(Error::InvalidSource);
        }
        observed = observed.checked_add(1).ok_or(Error::InvalidSource)?;
    }
    if observed != expected.len() {
        return Err(Error::InvalidSource);
    }
    Ok(())
}

/// Validate selected message-level metadata before a local payload rewrite.
///
/// Merge/diff overlays are intentionally rejected because a focused physical
/// adapter does not carry a base-message resolver.  Refusing such a source is
/// safer than silently changing only one layer of a merged native value.
pub(crate) fn validate_selected_metadata(
    object: &ArchiveObject,
    index: usize,
) -> Result<(), Error> {
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

/// Prove that one selected object-reference field is the sole owner of its
/// aggregate identifier and that no second field-local edge aliases it.
pub(crate) fn validate_object_reference_metadata(
    object: &ArchiveObject,
    index: usize,
    identifier: u64,
    path: u32,
) -> Result<(), Error> {
    validate_object_reference_metadata_with_policy(object, index, identifier, path, true)
}

fn validate_unmodified_object_reference_metadata(
    object: &ArchiveObject,
    index: usize,
    identifier: u64,
    path: u32,
) -> Result<(), Error> {
    validate_object_reference_metadata_with_policy(object, index, identifier, path, false)
}

fn validate_object_reference_metadata_with_policy(
    object: &ArchiveObject,
    index: usize,
    identifier: u64,
    path: u32,
    require_selected_path: bool,
) -> Result<(), Error> {
    if identifier == 0 {
        return Err(Error::InvalidSource);
    }
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
                || field.r#type != Some(FieldType::ObjectReference)
                || field.object_references.as_slice() != [identifier]
            {
                return Err(Error::InvalidSource);
            }
            selected_path = true;
        } else if field.object_references.contains(&identifier) {
            return Err(Error::InvalidSource);
        }
    }
    if require_selected_path && !selected_path {
        return Err(Error::InvalidSource);
    }
    Ok(())
}

/// Validate the selected soundtrack's archive-info data-reference closure.
///
/// This scan is intentionally narrow and borrowed: it visits only the
/// selected `PackageMetadata` object, current component declarations, owner
/// records, and matching `DataInfo` records. Versioned component declarations
/// (field 11) are parsed for shape but never satisfy current ownership.
pub(crate) fn validate_media_closure(
    package: &Package,
    selection: &Selection<'_>,
    budget: &mut Budget,
) -> Result<(), Error> {
    let limits = package.wire_limits().map_err(map_wire_error)?;
    let info = selection
        .soundtrack
        .archive_info
        .message_infos
        .get(selection.soundtrack_message_index)
        .ok_or(Error::InvalidSource)?;
    let mut states = HashMap::new();
    budget.charge_work(info.data_references.len())?;
    states
        .try_reserve(info.data_references.len())
        .map_err(|_| Error::Allocation {
            amount: info.data_references.len(),
        })?;
    for identifier in &info.data_references {
        if *identifier == 0 {
            return Err(Error::InvalidSource);
        }
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
    validate_data_metadata_map(catalog, metadata_payload, limits, budget)?;
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
    budget.charge_work(states.len())?;
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

/// Validate the optional root `PackageMetadata.data_metadata_map` closure.
///
/// The root edge, selected map object, map entries, and each pointed-to
/// metadata object are all required to be unique and in-package. The map is
/// optional and its entries may describe non-audio data; therefore this check
/// does not require every selected soundtrack asset to have a cosmetic
/// `DataMetadata` entry. It does, however, reject duplicate map fields/entries,
/// zero or external references, and dangling map targets before a lifecycle
/// adapter is allowed to alter media ownership.
pub(crate) fn validate_data_metadata_map(
    catalog: &SourceCatalog,
    metadata_payload: &[u8],
    limits: WireLimits,
    budget: &mut Budget,
) -> Result<(), Error> {
    let mut input = metadata_payload;
    let mut map_identifier = None;
    while let Some(field) = next_raw_field(&mut input, budget)? {
        if field.number != 10 {
            continue;
        }
        if field.wire != LENGTH_DELIMITED_WIRE_TYPE || map_identifier.is_some() {
            return Err(Error::InvalidSource);
        }
        let reference = strict_reference(field.bytes.ok_or(Error::InvalidSource)?, limits, budget)?;
        if reference.is_external() {
            return Err(Error::InvalidSource);
        }
        map_identifier = Some(reference.identifier);
    }
    let Some(map_identifier) = map_identifier else {
        return Ok(());
    };
    let (_, map_object) = unique_object(catalog, map_identifier)?;
    let (map_message_index, map_payload) =
        selected_message(map_object, DATA_METADATA_MAP_MESSAGE_TYPE)?;
    validate_selected_metadata(map_object, map_message_index)?;

    let mut seen_data_identifiers = HashSet::new();
    budget.charge_work(map_payload.len())?;
    let mut map_input = map_payload;
    while let Some(field) = next_raw_field(&mut map_input, budget)? {
        if field.number != 1 {
            continue;
        }
        if field.wire != LENGTH_DELIMITED_WIRE_TYPE {
            return Err(Error::InvalidSource);
        }
        let (data_identifier, metadata_identifier) =
            parse_data_metadata_entry(field.bytes.ok_or(Error::InvalidSource)?, limits, budget)?;
        if !seen_data_identifiers.insert(data_identifier) {
            return Err(Error::InvalidSource);
        }
        // `DataMetadata` is a native object edge, not an opaque scalar. A
        // missing or duplicate target would make later owner reclamation
        // ambiguous, so prove physical uniqueness while the catalog is open.
        let (_, metadata_object) = unique_object(catalog, metadata_identifier)?;
        if metadata_object.messages.is_empty() {
            return Err(Error::InvalidSource);
        }
    }
    Ok(())
}

fn parse_data_metadata_entry(
    source: &[u8],
    limits: WireLimits,
    budget: &mut Budget,
) -> Result<(u64, u64), Error> {
    budget.require_depth(3)?;
    budget.charge_work(source.len())?;
    let mut input = source;
    let mut data_identifier = None;
    let mut metadata_identifier = None;
    while let Some(field) = next_raw_field(&mut input, budget)? {
        match field.number {
            1 if field.wire == 0 && data_identifier.is_none() => {
                let value = field.varint.ok_or(Error::InvalidSource)?;
                if value == 0 {
                    return Err(Error::InvalidSource);
                }
                data_identifier = Some(value);
            },
            2 if field.wire == LENGTH_DELIMITED_WIRE_TYPE && metadata_identifier.is_none() => {
                let reference =
                    strict_reference(field.bytes.ok_or(Error::InvalidSource)?, limits, budget)?;
                if reference.is_external() {
                    return Err(Error::InvalidSource);
                }
                metadata_identifier = Some(reference.identifier);
            },
            1 | 2 => return Err(Error::InvalidSource),
            _ => {},
        }
    }
    Ok((
        data_identifier.ok_or(Error::InvalidSource)?,
        metadata_identifier.ok_or(Error::InvalidSource)?,
    ))
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

fn stream_package_metadata<'a>(
    source: &'a [u8],
    locator: &[u8],
    soundtrack_identifier: u64,
    states: &mut HashMap<u64, MediaClosureState<'a>>,
    budget: &mut Budget,
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
                // Versioned declarations describe historical ownership and
                // are never accepted as the current selected component.
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
    budget: &mut Budget,
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
    budget: &mut Budget,
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
            let count = count.ok_or(Error::InvalidSource)?;
            if count == 0 {
                return Err(Error::InvalidSource);
            }
            state.owner_count = usize::try_from(count).map_err(|_| Error::InvalidSource)?;
        }
    }
    Ok(())
}

fn stream_data_info<'a>(
    source: &'a [u8],
    states: &mut HashMap<u64, MediaClosureState<'a>>,
    budget: &mut Budget,
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

/// Validate the selected soundtrack's field-local role metadata and reject
/// data references attached to an unmodelled field path.
pub(crate) fn validate_soundtrack_metadata(
    object: &ArchiveObject,
    index: usize,
) -> Result<(), Error> {
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
                || field.r#type != Some(FieldType::DataReference)
                || field.data_references != info.data_references
            {
                return Err(Error::InvalidSource);
            }
            media_path = true;
        } else if !field.data_references.is_empty() {
            return Err(Error::InvalidSource);
        }
    }
    // Native Keynote may omit field-local attribution and carry the complete
    // soundtrack media list only in MessageInfo.data_references. The item
    // writer preserves that producer-selected omission while transitioning
    // the aggregate list. Any present field remains strict and unique, and no
    // unrelated field may alias a soundtrack data reference.
    Ok(())
}

/// Reject a selected soundtrack identifier appearing in unrelated show/root
/// reference roles. An external marker is also rejected in every visited
/// role, because a local rewrite cannot preserve an external edge safely.
pub(crate) fn validate_reference_role_disjointness(
    selection: &Selection<'_>,
    limits: WireLimits,
    budget: &mut Budget,
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
    let _ = preflight_wire_message(show_payload, limits, budget)?;
    reject_selected_identifier_in_reference_fields(
        show_payload,
        &[1, 2, 5, 7, 19],
        selection.soundtrack_identifier,
        limits,
        budget,
    )?;
    budget.charge_work(show_payload.len())?;
    let show = WireView::parse_with_limits(show_payload, remaining_wire_limits(limits, budget)?)
        .map_err(map_wire_error)?;
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
    budget: &mut Budget,
) -> Result<(), Error> {
    let _ = preflight_wire_message(source, limits, budget)?;
    budget.charge_work(source.len())?;
    let view = WireView::parse_with_limits(source, remaining_wire_limits(limits, budget)?)
        .map_err(map_wire_error)?;
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

fn charge_message_info(
    object: &ArchiveObject,
    index: usize,
    budget: &mut Budget,
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

/// Charge the parsed archive-info topology before a rewriter retains any
/// additional candidate state. This walks metadata only; payload bytes remain
/// owned by the source catalog and are charged by the operation that decodes
/// them.
pub(crate) fn charge_catalog_structure(
    catalog: &SourceCatalog,
    budget: &mut Budget,
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

/// Charge the bounded logical/IWA reopen work for a candidate package. The
/// optional replacement tuple is `(member name, logical bytes, decoded IWA
/// bytes)` and lets a rewriter charge its changed member without reopening it
/// twice just to estimate cost.
pub(crate) fn charge_catalog_reopen_cost(
    catalog: &SourceCatalog,
    raw_package_bytes: usize,
    replacement: Option<(&str, usize, usize)>,
    budget: &mut Budget,
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

fn archive_stream_extent(archive: &litchi_iwa_core::Archive) -> usize {
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

/// Borrow one raw protobuf field while requiring canonical key, length, and
/// scalar varint framing. The parser has no allocation and keeps all payloads
/// tied to the supplied source slice.
#[derive(Clone, Copy, Debug)]
struct RawField<'a> {
    number: u32,
    wire: u8,
    varint: Option<u64>,
    bytes: Option<&'a [u8]>,
}

fn next_raw_field<'a>(
    input: &mut &'a [u8],
    budget: &mut Budget,
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

pub(crate) fn canonical_varint(source: &[u8]) -> Result<u64, Error> {
    let (value, consumed) = decode_varint_from_bytes(source).map_err(|_| Error::InvalidSource)?;
    if consumed != source.len() || consumed != encoded_len(value) {
        return Err(Error::InvalidSource);
    }
    Ok(value)
}

pub(crate) const fn canonical_int32(value: u64) -> bool {
    value <= i32::MAX as u64 || value >= 0xffff_ffff_8000_0000
}

pub(crate) fn remaining_wire_limits(
    limits: WireLimits,
    budget: &Budget,
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
        .and_then(|value| value.with_rewrite_work(budget.remaining_work()))
        .map_err(map_wire_error)
}

/// Check a length-preserving or bounded wire replacement before allocating
/// its output buffer.
pub(crate) fn check_wire_output_bound(
    output_bytes: usize,
    limits: WireLimits,
) -> Result<(), Error> {
    if output_bytes > limits.max_output_bytes() {
        return Err(Error::LimitExceeded {
            kind: LimitKind::OutputBytes,
            observed: output_bytes as u64,
            maximum: limits.max_output_bytes() as u64,
        });
    }
    Ok(())
}

fn map_wire_error(error: litchi_iwa_common::Error) -> Error {
    match error {
        litchi_iwa_common::Error::LimitExceeded {
            kind,
            observed,
            limit,
        } => Error::LimitExceeded {
            kind: match kind {
                litchi_iwa_common::LimitKind::InputBytes => LimitKind::InputBytes,
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

/// Checked upper bound for ZIP reassembly before a replacement buffer is
/// allocated. The replacement member may be re-encoded by the archive layer,
/// so its logical length is doubled with a fixed metadata allowance.
pub(crate) fn archive_output_upper_bound(
    source_bytes: usize,
    replacement_bytes: usize,
) -> Result<usize, Error> {
    let replacement_bound = replacement_bytes
        .checked_mul(2)
        .and_then(|amount| amount.checked_add(1_024))
        .ok_or(Error::InvalidSource)?;
    source_bytes
        .checked_add(replacement_bound)
        .ok_or(Error::InvalidSource)
}

/// Check a ZIP reassembly upper bound against the physical package profile.
pub(crate) fn check_archive_output_bound(output_bound: usize, limits: Limits) -> Result<(), Error> {
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

/// Perform the conservative package-output preflight used before decompression
/// or archive/reassembly allocations.
pub(crate) fn preflight_rewrite_output(catalog: &SourceCatalog) -> Result<(), Error> {
    let limits = catalog.limits();
    let replacement_bound = limits
        .max_iwa_stream_bytes()
        .checked_add(32)
        .ok_or(Error::InvalidSource)?;
    let output_bound =
        archive_output_upper_bound(catalog.shared_source().len(), replacement_bound)?;
    check_archive_output_bound(output_bound, limits)
}

/// Charge the same upper bound used by [`preflight_rewrite_output`].
pub(crate) fn charge_output_upper_bound(
    source_bytes: usize,
    replacement_bytes: usize,
    limits: Limits,
    budget: &mut Budget,
) -> Result<usize, Error> {
    let output_bound = archive_output_upper_bound(source_bytes, replacement_bytes)?;
    check_archive_output_bound(output_bound, limits)?;
    budget.charge_work(output_bound)?;
    Ok(output_bound)
}

/// ZIP fields derived from a decoded entry's current logical payload.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) struct ZipEntryFields {
    /// CRC-32 of the logical entry bytes.
    pub(crate) crc32: u32,
    /// Compressed member length in the raw record.
    pub(crate) compressed_size: u32,
    /// Logical member length.
    pub(crate) uncompressed_size: u32,
}

/// Derive the expected ZIP CRC and size triple for one entry.
pub(crate) fn zip_entry_fields(
    entry: &litchi_iwa_archive::package::Entry,
) -> Option<ZipEntryFields> {
    Some(ZipEntryFields {
        crc32: zip_crc32(entry.data()),
        compressed_size: u32::try_from(entry.raw_record().compressed_data().len()).ok()?,
        uncompressed_size: u32::try_from(entry.data().len()).ok()?,
    })
}

/// Verify a retained central-directory record changed only in its relative
/// local-header offset and that the offset moved by the expected checked
/// delta.
pub(crate) fn central_record_preserved_except_offset(
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

/// Verify the selected local record, including descriptor fields for a
/// data-descriptor member. Names, flags, timestamps, extras, and all suffix
/// bytes remain source-exact.
pub(crate) fn selected_local_record_preserved(
    before: &litchi_iwa_archive::package::Entry,
    after: &litchi_iwa_archive::package::Entry,
    before_fields: ZipEntryFields,
    after_fields: ZipEntryFields,
) -> bool {
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

/// Verify a selected central-directory record, allowing only CRC/size and
/// relative local-header offset changes.
pub(crate) fn selected_central_record_preserved(
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

/// Verify a selected local suffix, including optional data-descriptor values.
pub(crate) fn selected_local_suffix_preserved(
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

/// Compare two byte records while ignoring one or two explicit, disjoint
/// ranges. Invalid or overlapping ranges fail closed.
pub(crate) fn bytes_equal_outside_ranges(
    before: &[u8],
    after: &[u8],
    first: Range<usize>,
    second: Option<Range<usize>>,
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

fn is_local_record(record: &[u8]) -> bool {
    record.get(..4) == Some(b"PK\x03\x04")
}

fn is_central_record(record: &[u8]) -> bool {
    record.get(..4) == Some(b"PK\x01\x02")
}

fn zip_local_header_length(record: &[u8]) -> Option<usize> {
    if !is_local_record(record) {
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

/// Compute the checked physical offset delta between two source-backed raw
/// records.
pub(crate) fn record_offset_delta(
    before_source: &[u8],
    after_source: &[u8],
    before_record: &[u8],
    after_record: &[u8],
) -> Option<i128> {
    let before = i128::try_from(raw_slice_offset(before_source, before_record)?).ok()?;
    let after = i128::try_from(raw_slice_offset(after_source, after_record)?).ok()?;
    after.checked_sub(before)
}

/// Compute ZIP CRC-32 without allocating.
pub(crate) fn zip_crc32(bytes: &[u8]) -> u32 {
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

/// Verify package bytes outside entry records (ZIP prelude and EOCD tail)
/// remain exact after reassembly.
pub(crate) fn verify_zip_boundaries(
    before: &SourceCatalog,
    after: &SourceCatalog,
    budget: &mut Budget,
) -> Result<(), Error> {
    let before_bytes = before.source_bytes();
    let after_bytes = after.source_bytes();
    let before_first = before.package().iter().next().ok_or(Error::InvalidSource)?;
    let after_first = after.package().iter().next().ok_or(Error::InvalidSource)?;
    let before_prelude_end =
        raw_slice_offset(before_bytes, before_first.raw_record().local_record())
            .ok_or(Error::InvalidSource)?;
    let after_prelude_end = raw_slice_offset(after_bytes, after_first.raw_record().local_record())
        .ok_or(Error::InvalidSource)?;
    budget.charge_work(
        before_prelude_end
            .checked_add(after_prelude_end)
            .ok_or(Error::InvalidSource)?,
    )?;
    if before_bytes.get(..before_prelude_end) != after_bytes.get(..after_prelude_end) {
        return Err(Error::InvalidSource);
    }
    let before_tail = eocd_tail(before_bytes).ok_or(Error::InvalidSource)?;
    let after_tail = eocd_tail(after_bytes).ok_or(Error::InvalidSource)?;
    budget.charge_work(
        before_tail
            .len()
            .checked_add(after_tail.len())
            .ok_or(Error::InvalidSource)?,
    )?;
    if before_tail.len() != after_tail.len()
        || before_tail[..16] != after_tail[..16]
        || before_tail[20..] != after_tail[20..]
    {
        return Err(Error::InvalidSource);
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use std::collections::HashMap;

    use litchi_iwa_common::{WireLimits, encode_varint_into};

    use super::{
        Budget, Error, LimitKind, MediaClosureState, SOUNDTRACK_MEDIA_FIELD,
        bytes_equal_outside_ranges, canonical_int32, selected_local_suffix_preserved,
        stream_package_metadata, strict_reference, validate_soundtrack_media_references,
    };

    fn budget() -> Budget {
        Budget::from_limits(
            WireLimits::MAX_FIELDS,
            WireLimits::MAX_REWRITE_WORK,
            1_000_000,
            WireLimits::MAX_NESTING,
        )
    }

    fn push_varint_field(output: &mut Vec<u8>, number: u32, value: u64) {
        encode_varint_into(output, u64::from(number) << 3);
        encode_varint_into(output, value);
    }

    fn push_bytes_field(output: &mut Vec<u8>, number: u32, value: &[u8]) {
        encode_varint_into(output, (u64::from(number) << 3) | 2);
        encode_varint_into(output, value.len() as u64);
        output.extend_from_slice(value);
    }

    fn reference(identifier: u64, external: Option<bool>) -> Vec<u8> {
        let mut output = Vec::new();
        push_varint_field(&mut output, 1, identifier);
        if let Some(external) = external {
            push_varint_field(&mut output, 3, u64::from(external));
        }
        output
    }

    fn soundtrack_payload(references: &[u64]) -> Vec<u8> {
        let mut output = Vec::new();
        for identifier in references {
            let value = reference(*identifier, None);
            push_bytes_field(&mut output, SOUNDTRACK_MEDIA_FIELD, &value);
        }
        output
    }

    #[test]
    fn strict_reference_rejects_zero_identifier() {
        let source = reference(0, None);
        let mut transaction_budget = budget();
        assert!(matches!(
            strict_reference(&source, WireLimits::default(), &mut transaction_budget),
            Err(Error::InvalidSource)
        ));
    }

    #[test]
    fn strict_reference_retains_external_marker_without_accepting_unknown_value() {
        let source = reference(7, Some(true));
        let mut transaction_budget = budget();
        let facts = strict_reference(&source, WireLimits::default(), &mut transaction_budget)
            .expect("strict reference");
        assert_eq!(facts.identifier, 7);
        assert!(facts.is_external());

        let invalid = reference(7, Some(true));
        let mut malformed = invalid.clone();
        // The final value is the external marker. Replace it with a value
        // outside the native bool domain while preserving canonical framing.
        *malformed.last_mut().expect("marker byte") = 2;
        let mut transaction_budget = budget();
        assert!(matches!(
            strict_reference(&malformed, WireLimits::default(), &mut transaction_budget),
            Err(Error::InvalidSource)
        ));
    }

    #[test]
    fn strict_reference_rejects_duplicate_known_fields() {
        let mut source = reference(7, None);
        push_varint_field(&mut source, 1, 8);
        let mut transaction_budget = budget();
        assert!(matches!(
            strict_reference(&source, WireLimits::default(), &mut transaction_budget),
            Err(Error::InvalidSource)
        ));
    }

    #[test]
    fn media_records_reconcile_with_expected_count_and_preserve_unknown_fields() {
        let mut source = soundtrack_payload(&[7, 9]);
        push_bytes_field(&mut source, 99, b"future");
        let mut transaction_budget = budget();
        validate_soundtrack_media_references(
            &source,
            WireLimits::default(),
            &[7, 9],
            &mut transaction_budget,
        )
        .expect("media records are valid");

        let mut transaction_budget = budget();
        assert!(matches!(
            validate_soundtrack_media_references(
                &source,
                WireLimits::default(),
                &[7],
                &mut transaction_budget,
            ),
            Err(Error::InvalidSource)
        ));
    }

    #[test]
    fn media_records_reject_external_reference() {
        let mut source = Vec::new();
        let value = reference(7, Some(true));
        push_bytes_field(&mut source, SOUNDTRACK_MEDIA_FIELD, &value);
        let mut transaction_budget = budget();
        assert!(matches!(
            validate_soundtrack_media_references(
                &source,
                WireLimits::default(),
                &[7],
                &mut transaction_budget,
            ),
            Err(Error::InvalidSource)
        ));
    }

    #[test]
    fn media_records_reject_identifier_or_order_mismatch() {
        let source = soundtrack_payload(&[7, 9]);
        for expected in [&[7, 8][..], &[9, 7][..]] {
            let mut transaction_budget = budget();
            assert!(matches!(
                validate_soundtrack_media_references(
                    &source,
                    WireLimits::default(),
                    expected,
                    &mut transaction_budget,
                ),
                Err(Error::InvalidSource)
            ));
        }
    }

    #[test]
    fn package_metadata_versioned_component_does_not_prove_current_ownership() {
        const SOUNDTRACK_IDENTIFIER: u64 = 10;
        const DATA_IDENTIFIER: u64 = 42;
        const LOCATOR: &[u8] = b"Document";
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
            .expect("metadata is parseable"),
            0
        );
        let state = states.get(&DATA_IDENTIFIER).expect("state remains");
        assert_eq!(state.component_declarations, 0);
        assert_eq!(state.owner_occurrences, 0);
    }

    #[test]
    fn package_metadata_rejects_zero_owner_count() {
        const SOUNDTRACK_IDENTIFIER: u64 = 10;
        const DATA_IDENTIFIER: u64 = 42;
        const LOCATOR: &[u8] = b"Document";
        let mut owner = Vec::new();
        push_varint_field(&mut owner, 1, SOUNDTRACK_IDENTIFIER);
        push_varint_field(&mut owner, 2, 0);
        let mut data_reference = Vec::new();
        push_varint_field(&mut data_reference, 1, DATA_IDENTIFIER);
        push_bytes_field(&mut data_reference, 2, &owner);
        let mut component = Vec::new();
        push_bytes_field(&mut component, 2, LOCATOR);
        push_bytes_field(&mut component, 7, &data_reference);
        let mut metadata = Vec::new();
        push_bytes_field(&mut metadata, 3, &component);
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
        assert!(matches!(
            stream_package_metadata(
                &metadata,
                LOCATOR,
                SOUNDTRACK_IDENTIFIER,
                &mut states,
                &mut transaction_budget,
            ),
            Err(Error::InvalidSource)
        ));
    }

    #[test]
    fn budget_uses_checked_charges_and_independent_item_counter() {
        let mut transaction_budget = Budget::from_limits(2, 3, 4, 5);
        transaction_budget.charge_fields(2).expect("field budget");
        transaction_budget.charge_work(3).expect("work budget");
        transaction_budget.charge_items(4).expect("item budget");
        assert_eq!(transaction_budget.fields(), 2);
        assert_eq!(transaction_budget.work(), 3);
        assert_eq!(transaction_budget.items(), 4);
        assert!(matches!(
            transaction_budget.charge_fields(1),
            Err(Error::LimitExceeded {
                kind: LimitKind::WireFields,
                ..
            })
        ));
        assert!(matches!(
            transaction_budget.charge_work(1),
            Err(Error::LimitExceeded {
                kind: LimitKind::WireWork,
                ..
            })
        ));
        assert!(matches!(
            transaction_budget.charge_items(1),
            Err(Error::LimitExceeded {
                kind: LimitKind::Items,
                ..
            })
        ));
    }

    #[test]
    fn bytes_equal_outside_ranges_rejects_invalid_or_overlapping_ranges() {
        let before = b"0123456789";
        let mut after = before.to_vec();
        after[4] = b'x';
        assert!(bytes_equal_outside_ranges(before, &after, 4..5, None));
        assert!(!bytes_equal_outside_ranges(
            before,
            &after,
            4..7,
            Some(6..8)
        ));
        assert!(!bytes_equal_outside_ranges(before, &after, 11..11, None));
    }

    #[test]
    fn descriptor_suffix_requires_exact_descriptor_and_tail() {
        let fields = super::ZipEntryFields {
            crc32: 0x0102_0304,
            compressed_size: 5,
            uncompressed_size: 9,
        };
        let mut before = b"PK\x07\x08".to_vec();
        before.extend(fields.crc32.to_le_bytes());
        before.extend(fields.compressed_size.to_le_bytes());
        before.extend(fields.uncompressed_size.to_le_bytes());
        before.extend_from_slice(b"tail");
        let after = before.clone();
        assert!(selected_local_suffix_preserved(
            0x0008, &before, &after, fields, fields
        ));

        let mut changed = after;
        *changed.last_mut().expect("tail") = b'!';
        assert!(!selected_local_suffix_preserved(
            0x0008, &before, &changed, fields, fields
        ));
    }

    #[test]
    fn canonical_signed_int32_accepts_native_twos_complement_range_only() {
        assert!(canonical_int32(0));
        assert!(canonical_int32(i64::from(i32::MAX) as u64));
        assert!(canonical_int32(0xffff_ffff_8000_0000));
        assert!(!canonical_int32(0xffff_ffff_7fff_ffff));
    }
}
