//! Shared Numbers chart title and axis-title transaction machinery.
//!
//! The public title and axis-title modules are thin selector facades over this
//! owner.  Native chart references and IWA objects stay here; the wire-level
//! title engines are supplied by `litchi-iwa-style::chart_text`.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "The focused chart-text transaction keeps its bounded phases together."
)]

use std::fmt;
use std::sync::Arc;

use litchi_core::Position;
use litchi_iwa_archive::package::{EntryEdit, OwnedExactArtifacts};
use litchi_iwa_archive::{Error as ArchiveError, LimitKind as ArchiveLimitKind};
use litchi_iwa_common::chart::axis::Axis;
use litchi_iwa_core::{ArchiveObject, RawMessage, SnappyStream};
use litchi_iwa_protos::chart_style_codec::{self as chart_codec, Archive as ChartArchive};
use litchi_iwa_style::chart_text::{self, ChartTextKind};
use thiserror::Error as ThisError;

use super::Package;
use super::chart_arrangement::{
    self, ChartBudget, ChartSelection, map_package_error as map_arrangement_package_error,
    object_from_resolved, select_chart_with_budget, selected_chart_payload, unique_typed_message,
    validate_message_metadata,
};
use crate::selector::{ChartSelector, SheetSelector};

const GENERATED_EXTENSION_FIELD: u32 = 10_000;
const CHART_MESSAGE_TYPE: u32 = 5_021;

/// The chart-text property addressed by a transaction.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub(super) enum Kind {
    Title,
    Axis(Axis),
}

impl Kind {
    const fn message_type(self) -> u32 {
        match self {
            Self::Title => ChartArchive::ChartNonStyle.message_type(),
            Self::Axis(_) => ChartArchive::AxisNonStyle.message_type(),
        }
    }

    const fn reference_field(self) -> u32 {
        match self {
            Self::Title => 10,
            Self::Axis(Axis::Value) => 14,
            Self::Axis(Axis::Category) => 16,
        }
    }
}

/// A finite resource governed by one Numbers chart-text transaction.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Complete package bytes inspected or retained.
    PackageBytes,
    /// One decoded IWA component payload.
    PayloadBytes,
    /// Semantic graph items inspected during candidate reopen.
    PayloadItems,
    /// Native chart references inspected by the ownership proof.
    PayloadReferences,
    /// Strict generated-extension bytes.
    WireBytes,
    /// Rewritten output bytes.
    OutputBytes,
    /// Strict generated-extension fields visited.
    WireFields,
    /// Strict generated-extension nesting.
    WireNesting,
    /// Strict generated-extension work.
    WireWork,
    /// Retained title text bytes.
    TextBytes,
    /// Bounded temporary allocations.
    Allocations,
    /// Source and candidate bytes retained together.
    RetainedBytes,
    /// Candidate scratch bytes.
    ScratchBytes,
}

impl fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::PackageBytes => "package bytes",
            Self::PayloadBytes => "payload bytes",
            Self::PayloadItems => "payload items",
            Self::PayloadReferences => "payload references",
            Self::WireBytes => "wire bytes",
            Self::OutputBytes => "output bytes",
            Self::WireFields => "wire fields",
            Self::WireNesting => "wire nesting",
            Self::WireWork => "wire work",
            Self::TextBytes => "text bytes",
            Self::Allocations => "allocations",
            Self::RetainedBytes => "retained bytes",
            Self::ScratchBytes => "scratch bytes",
        })
    }
}

/// A typed refusal from a Numbers chart-text route.
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Refusal {
    /// No sheet matched the semantic selector.
    #[error("sheet selector did not resolve")]
    SheetNotFound,
    /// No chart matched the semantic selector.
    #[error("chart selector did not resolve")]
    ChartNotFound,
    /// The selected chart does not carry the requested axis.
    #[error("the chart has no requested axis")]
    UnsupportedAxis,
    /// The selected native text object is shared by another chart owner.
    #[error("the chart text object is shared by another chart")]
    SharedObject,
    /// The source does not retain an exact physical artifact.
    #[error("this Numbers source does not support chart-text edits")]
    UnsupportedSource,
    /// The edit has no staged operation; an unstaged edit is an exact no-op.
    #[error("no chart-text operation was staged")]
    NoStagedOperation,
    /// The supplied text contains a native object marker.
    #[error("chart text contains a reserved native marker")]
    ReservedMarker,
}

impl Refusal {
    /// Classify the refusal for unified operations.
    #[must_use]
    pub const fn kind(self) -> litchi_core::edit::RefusalKind {
        use litchi_core::edit::RefusalKind;
        match self {
            Self::SheetNotFound | Self::ChartNotFound => RefusalKind::NotFound,
            Self::UnsupportedAxis => RefusalKind::UnsupportedTarget,
            Self::SharedObject | Self::UnsupportedSource => RefusalKind::UnsupportedContent,
            Self::NoStagedOperation => RefusalKind::UnsupportedOperation,
            Self::ReservedMarker => RefusalKind::InvalidValue,
        }
    }
}

/// Failure from a Numbers chart-text read or transaction.
#[derive(Debug, Clone, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// A well-formed request this route refuses.
    #[error(transparent)]
    Refused(Refusal),
    /// The rooted chart graph or selected generated extension is malformed.
    #[error("the selected Numbers chart-text source is invalid")]
    InvalidSource,
    /// A finite operation ceiling was exceeded.
    #[error("Numbers chart text {kind} limit exceeded: observed {observed}, maximum {maximum}")]
    Limit {
        /// Resource category.
        kind: LimitKind,
        /// Observed amount.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
    },
    /// A bounded temporary allocation failed.
    #[error("could not allocate {amount} Numbers chart-text units")]
    Allocation { amount: usize },
    /// Candidate reopen or semantic readback failed.
    #[error("the edited Numbers chart text failed semantic verification")]
    Verification,
    /// A patch was applied to a different exact source.
    #[error("the Numbers chart-text patch does not match its exact source")]
    PatchConflict,
}

impl From<Refusal> for Error {
    fn from(refusal: Refusal) -> Self {
        Self::Refused(refusal)
    }
}

/// A selected chart text owner.  All identifiers remain private to this
/// package module and never appear in the public edit or patch.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(super) struct Target {
    pub(super) sheet: Position,
    pub(super) chart: Position,
    pub(super) kind: Kind,
    pub(super) component_index: usize,
    pub(super) component_name: Arc<str>,
    pub(super) object_index: usize,
    pub(super) object_identifier: u64,
    pub(super) message_index: usize,
    pub(super) message_type: u32,
    pub(super) before: Option<Box<str>>,
}

impl Target {
    fn object<'a>(&self, package: &'a Package) -> Result<&'a ArchiveObject, Error> {
        let component = package
            .state
            .components
            .catalog()
            .get_index(self.component_index)
            .ok_or(Error::InvalidSource)?;
        if component.name() != self.component_name.as_ref() {
            return Err(Error::InvalidSource);
        }
        let object = component
            .archive()
            .objects
            .get(self.object_index)
            .ok_or(Error::InvalidSource)?;
        if object.archive_info.identifier != Some(self.object_identifier) {
            return Err(Error::InvalidSource);
        }
        let message = object
            .messages
            .get(self.message_index)
            .ok_or(Error::InvalidSource)?;
        if message.type_ != self.message_type {
            return Err(Error::InvalidSource);
        }
        Ok(object)
    }

    fn payload<'a>(&self, package: &'a Package) -> Result<&'a [u8], Error> {
        Ok(self
            .object(package)?
            .messages
            .get(self.message_index)
            .ok_or(Error::InvalidSource)?
            .data
            .as_slice())
    }
}

/// Resolve a chart text target and prove the selected reference is owned by
/// exactly one chart graph.
pub(super) fn resolve<'sheet>(
    package: &Package,
    sheet_selector: impl Into<SheetSelector<'sheet>>,
    chart_selector: impl Into<ChartSelector>,
    kind: Kind,
) -> Result<Target, Error> {
    let mut budget = ChartBudget::for_package(package).map_err(map_arrangement_error)?;
    let selection = select_chart_with_budget(
        package,
        sheet_selector.into(),
        chart_selector.into(),
        false,
        &mut budget,
    )
    .map_err(map_arrangement_error)?;
    let payload = selected_chart_payload(package, &selection).map_err(map_arrangement_error)?;
    let mut reference_budget = budget;
    let references = decode_references(payload, &mut reference_budget)?;
    let identifier = match kind {
        Kind::Title => references.chart_non_style(),
        Kind::Axis(Axis::Category) => references.category_axis_non_styles().first().copied(),
        Kind::Axis(Axis::Value) => references.value_axis_non_styles().first().copied(),
    }
    .filter(|identifier| *identifier != 0);
    let identifier = identifier.ok_or_else(|| match kind {
        Kind::Title => Error::InvalidSource,
        Kind::Axis(_) => Refusal::UnsupportedAxis.into(),
    })?;
    prove_reference_exclusive(
        package,
        &selection,
        &references,
        identifier,
        &mut reference_budget,
    )?;
    let chart_object = package
        .state
        .components
        .catalog()
        .get_index(selection.component_index)
        .and_then(|component| component.archive().objects.get(selection.object_index))
        .ok_or(Error::InvalidSource)?;
    let chart_info = chart_object
        .archive_info
        .message_infos
        .get(selection.message_index)
        .ok_or(Error::InvalidSource)?;
    if !owns_reference(chart_info, identifier, kind.reference_field()) {
        return Err(Error::InvalidSource);
    }
    let resolved = package
        .state
        .index
        .resolve_ref_id(&package.state.components, identifier)
        .map_err(|error| map_arrangement_error(map_arrangement_package_error(error)))?
        .ok_or(Error::InvalidSource)?;
    let object = object_from_resolved(package, resolved).map_err(map_arrangement_error)?;
    validate_message_metadata(object).map_err(map_arrangement_error)?;
    let (message_index, message) = unique_typed_message(object, kind.message_type())
        .map_err(map_arrangement_error)?
        .ok_or(Error::InvalidSource)?;
    let component = package
        .state
        .components
        .catalog()
        .get_index(resolved.component_index)
        .ok_or(Error::InvalidSource)?;
    let before = read_payload(message.data.as_slice(), kind)?;
    Ok(Target {
        sheet: selection.sheet_position,
        chart: selection.chart_position,
        kind,
        component_index: resolved.component_index,
        component_name: Arc::from(component.name()),
        object_index: resolved.object_index,
        object_identifier: identifier,
        message_index,
        message_type: kind.message_type(),
        before,
    })
}

fn owns_reference(info: &litchi_iwa_core::MessageInfo, identifier: u64, field: u32) -> bool {
    let aggregate = info
        .object_references
        .iter()
        .filter(|candidate| **candidate == identifier)
        .count();
    if aggregate != 1 || info.data_references.contains(&identifier) {
        return false;
    }
    let mut fields = 0usize;
    for field_info in &info.field_infos {
        let objects = field_info
            .object_references
            .iter()
            .filter(|candidate| **candidate == identifier)
            .count();
        let data = field_info
            .data_references
            .iter()
            .filter(|candidate| **candidate == identifier)
            .count();
        if objects == 0 && data == 0 {
            continue;
        }
        if data != 0
            || objects != 1
            || field_info.path.path.as_slice() != [GENERATED_EXTENSION_FIELD, field]
        {
            return false;
        }
        fields = fields.saturating_add(objects);
    }
    fields <= 1
}

fn reference_count(references: &chart_codec::StyleReferences, identifier: u64) -> usize {
    let singular = [
        references.chart_style(),
        references.chart_non_style(),
        references.legend_style(),
        references.legend_non_style(),
    ]
    .into_iter()
    .flatten()
    .filter(|candidate| *candidate == identifier)
    .count();
    let repeated = references
        .value_axis_styles()
        .iter()
        .chain(references.value_axis_non_styles())
        .chain(references.category_axis_styles())
        .chain(references.category_axis_non_styles())
        .chain(references.series_theme_styles())
        .chain(references.paragraph_styles())
        .filter(|candidate| **candidate == identifier)
        .count();
    let sparse = references
        .series_private_styles()
        .into_iter()
        .flat_map(|entries| entries.entries())
        .chain(
            references
                .series_non_styles()
                .into_iter()
                .flat_map(|entries| entries.entries()),
        )
        .filter(|(_, candidate)| *candidate == identifier)
        .count();
    singular
        .checked_add(repeated)
        .and_then(|count| count.checked_add(sparse))
        .unwrap_or(usize::MAX)
}

#[derive(Debug, Default)]
struct ReferenceOwners {
    count: usize,
}

impl ReferenceOwners {
    fn add(
        &mut self,
        references: &chart_codec::StyleReferences,
        identifier: u64,
    ) -> Result<(), Error> {
        self.count = self
            .count
            .checked_add(reference_count(references, identifier))
            .ok_or(Error::InvalidSource)?;
        if self.count > 1 {
            return Err(Refusal::SharedObject.into());
        }
        Ok(())
    }

    fn finish(self) -> Result<(), Error> {
        if self.count == 1 {
            Ok(())
        } else {
            Err(Refusal::SharedObject.into())
        }
    }
}

fn reference_vector_len(references: &chart_codec::StyleReferences) -> usize {
    let singular = [
        references.chart_style(),
        references.chart_non_style(),
        references.legend_style(),
        references.legend_non_style(),
    ]
    .into_iter()
    .filter(Option::is_some)
    .count();
    let repeated = references
        .value_axis_styles()
        .len()
        .saturating_add(references.value_axis_non_styles().len())
        .saturating_add(references.category_axis_styles().len())
        .saturating_add(references.category_axis_non_styles().len())
        .saturating_add(references.series_theme_styles().len())
        .saturating_add(references.paragraph_styles().len());
    let sparse = references
        .series_private_styles()
        .map_or(0, |entries| entries.entries().len())
        .saturating_add(
            references
                .series_non_styles()
                .map_or(0, |entries| entries.entries().len()),
        );
    singular.saturating_add(repeated).saturating_add(sparse)
}

fn decode_references(
    payload: &[u8],
    budget: &mut ChartBudget,
) -> Result<chart_codec::StyleReferences, Error> {
    budget
        .preflight_reference_scan(payload)
        .map_err(map_arrangement_error)?;
    let limits = budget
        .reference_codec_limits(payload)
        .map_err(map_arrangement_error)?;
    let references = chart_codec::references(payload, limits).map_err(map_chart_codec_error)?;
    budget
        .charge_reference_scan(payload, reference_vector_len(&references))
        .map_err(map_arrangement_error)?;
    Ok(references)
}

fn prove_reference_exclusive(
    package: &Package,
    selection: &ChartSelection,
    selected: &chart_codec::StyleReferences,
    identifier: u64,
    budget: &mut ChartBudget,
) -> Result<(), Error> {
    let mut owners = ReferenceOwners::default();
    owners.add(selected, identifier)?;
    for (component_index, (_name, archive)) in package.state.components.iter_archives().enumerate()
    {
        for (object_index, object) in archive.objects.iter().enumerate() {
            for (message_index, message) in object.messages.iter().enumerate() {
                if message.type_ != CHART_MESSAGE_TYPE
                    || (component_index == selection.component_index
                        && object_index == selection.object_index
                        && message_index == selection.message_index)
                {
                    continue;
                }
                let references = decode_references(message.data.as_slice(), budget)?;
                owners.add(&references, identifier)?;
            }
        }
    }
    owners.finish()
}

fn read_payload(payload: &[u8], kind: Kind) -> Result<Option<Box<str>>, Error> {
    let (snapshot, _report) = chart_text::decode(
        payload,
        shared_kind(kind),
        chart_text::Limits::for_source(payload),
    )
    .map_err(map_chart_text_error)?;
    Ok(snapshot.text().map(|text| text.into()))
}

fn rewrite_payload(payload: &[u8], kind: Kind, after: Option<&str>) -> Result<Vec<u8>, Error> {
    chart_text::rewrite(
        payload,
        shared_kind(kind),
        after,
        rewrite_limits(payload, after)?,
    )
    .map(|rewrite| rewrite.into_output())
    .map_err(map_chart_text_error)
}

const fn shared_kind(kind: Kind) -> ChartTextKind {
    match kind {
        Kind::Title => ChartTextKind::Title,
        Kind::Axis(axis) => ChartTextKind::Axis(axis),
    }
}

fn rewrite_limits(payload: &[u8], after: Option<&str>) -> Result<chart_text::Limits, Error> {
    let replacement = after.map_or(0, str::len);
    let output = payload
        .len()
        .checked_add(replacement)
        .and_then(|value| value.checked_add(96))
        .ok_or(Error::InvalidSource)?
        .max(1);
    let retained = payload
        .len()
        .checked_add(output)
        .ok_or(Error::InvalidSource)?
        .max(1);
    let work = payload
        .len()
        .checked_mul(16)
        .and_then(|value| value.checked_add(output))
        .ok_or(Error::InvalidSource)?
        .max(1);
    Ok(chart_text::Limits::new(
        payload.len().max(1),
        payload.len().saturating_mul(8).max(1),
        output,
        8,
        work,
        output,
        output,
        retained,
        output,
    ))
}

fn map_chart_text_error(error: chart_text::Error) -> Error {
    match error {
        chart_text::Error::InvalidSource => Error::InvalidSource,
        chart_text::Error::Allocation { amount } => Error::Allocation { amount },
        chart_text::Error::Limit {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: match kind {
                chart_text::LimitKind::InputBytes => LimitKind::PayloadBytes,
                chart_text::LimitKind::Fields => LimitKind::WireFields,
                chart_text::LimitKind::OutputBytes => LimitKind::OutputBytes,
                chart_text::LimitKind::Nesting => LimitKind::WireNesting,
                chart_text::LimitKind::Work => LimitKind::WireWork,
                chart_text::LimitKind::TextBytes => LimitKind::TextBytes,
                chart_text::LimitKind::Allocations => LimitKind::Allocations,
                chart_text::LimitKind::RetainedBytes => LimitKind::RetainedBytes,
                chart_text::LimitKind::ScratchBytes => LimitKind::ScratchBytes,
                _ => return Error::InvalidSource,
            },
            observed,
            maximum,
        },
        _ => Error::InvalidSource,
    }
}

fn map_chart_codec_error(error: chart_codec::Error) -> Error {
    match error {
        chart_codec::Error::Limit {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: match kind {
                chart_codec::LimitKind::Bytes => LimitKind::WireBytes,
                chart_codec::LimitKind::Fields => LimitKind::WireFields,
                _ => return Error::InvalidSource,
            },
            observed: count(observed),
            maximum: count(maximum),
        },
        chart_codec::Error::Allocation { bytes } => Error::Allocation { amount: bytes },
        _ => Error::InvalidSource,
    }
}

fn map_archive_error(error: ArchiveError) -> Error {
    match error {
        ArchiveError::Limit {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: match kind {
                ArchiveLimitKind::InputBytes
                | ArchiveLimitKind::OutputBytes
                | ArchiveLimitKind::Entries
                | ArchiveLimitKind::MemberNameBytes
                | ArchiveLimitKind::MetadataBytes
                | ArchiveLimitKind::CompressedEntryBytes
                | ArchiveLimitKind::EntryBytes
                | ArchiveLimitKind::TotalBytes => LimitKind::PackageBytes,
                ArchiveLimitKind::IwaStreamBytes | ArchiveLimitKind::IwaTotalBytes => {
                    LimitKind::PayloadBytes
                },
            },
            observed,
            maximum,
        },
        ArchiveError::Allocation { amount, .. } => Error::Allocation { amount },
        ArchiveError::Iwa(error) => map_core_error(error),
        ArchiveError::Reassembly(_) => Refusal::UnsupportedSource.into(),
        _ => Error::InvalidSource,
    }
}

fn map_core_error(error: litchi_iwa_core::Error) -> Error {
    match error {
        litchi_iwa_core::Error::Limit {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: match kind {
                litchi_iwa_core::LimitKind::HeaderFields => LimitKind::WireFields,
                litchi_iwa_core::LimitKind::HeaderNesting => LimitKind::WireNesting,
                litchi_iwa_core::LimitKind::Objects
                | litchi_iwa_core::LimitKind::Messages
                | litchi_iwa_core::LimitKind::MessagesPerObject
                | litchi_iwa_core::LimitKind::MetadataItems
                | litchi_iwa_core::LimitKind::SnappyFrames => LimitKind::PayloadBytes,
                _ => LimitKind::WireBytes,
            },
            observed: count(observed),
            maximum: count(maximum),
        },
        litchi_iwa_core::Error::Allocation { requested, .. } => {
            Error::Allocation { amount: requested }
        },
        _ => Error::InvalidSource,
    }
}

fn count(value: usize) -> u64 {
    u64::try_from(value).unwrap_or(u64::MAX)
}

fn map_candidate_error(error: super::Error) -> Error {
    if let Some(resource) = error.resource_error() {
        return match resource {
            super::ResourceError::Limit {
                kind,
                observed,
                maximum,
            } => Error::Limit {
                kind: match kind {
                    super::PayloadLimitKind::InputBytes => LimitKind::PayloadBytes,
                    super::PayloadLimitKind::Fields => LimitKind::WireFields,
                    super::PayloadLimitKind::OutputBytes => LimitKind::OutputBytes,
                    super::PayloadLimitKind::Nesting => LimitKind::WireNesting,
                    super::PayloadLimitKind::RewriteWork => LimitKind::WireWork,
                    super::PayloadLimitKind::TableRows
                    | super::PayloadLimitKind::TableColumns
                    | super::PayloadLimitKind::TableCells
                    | super::PayloadLimitKind::MaterializedCells => LimitKind::PayloadItems,
                },
                observed: count(observed),
                maximum: count(maximum),
            },
            super::ResourceError::Allocation { amount } => Error::Allocation { amount },
        };
    }
    match error {
        super::Error::Archive(error) => map_archive_error(error),
        super::Error::InputTooLarge { observed, maximum } => Error::Limit {
            kind: LimitKind::PackageBytes,
            observed,
            maximum,
        },
        super::Error::SemanticLimit {
            observed, maximum, ..
        } => Error::Limit {
            kind: LimitKind::PayloadBytes,
            observed: count(observed),
            maximum: count(maximum),
        },
        _ => Error::InvalidSource,
    }
}

fn map_arrangement_error(error: chart_arrangement::Error) -> Error {
    match error {
        chart_arrangement::Error::Refused(chart_arrangement::Refusal::UnsupportedSource) => {
            Refusal::UnsupportedSource.into()
        },
        chart_arrangement::Error::Refused(chart_arrangement::Refusal::SheetNameNotFound)
        | chart_arrangement::Error::Refused(chart_arrangement::Refusal::SheetPositionNotFound {
            ..
        }) => Refusal::SheetNotFound.into(),
        chart_arrangement::Error::Refused(chart_arrangement::Refusal::ChartPositionNotFound {
            ..
        }) => Refusal::ChartNotFound.into(),
        chart_arrangement::Error::Limit {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: match kind {
                chart_arrangement::LimitKind::WireBytes => LimitKind::WireBytes,
                chart_arrangement::LimitKind::WireFields => LimitKind::WireFields,
                chart_arrangement::LimitKind::WireNesting => LimitKind::WireNesting,
                chart_arrangement::LimitKind::WireWork => LimitKind::WireWork,
                chart_arrangement::LimitKind::OutputBytes => LimitKind::OutputBytes,
                chart_arrangement::LimitKind::Allocations => LimitKind::Allocations,
                chart_arrangement::LimitKind::RetainedBytes => LimitKind::RetainedBytes,
                chart_arrangement::LimitKind::ScratchBytes => LimitKind::ScratchBytes,
                chart_arrangement::LimitKind::References => LimitKind::PayloadReferences,
                chart_arrangement::LimitKind::CompressedBytes => LimitKind::PackageBytes,
            },
            observed,
            maximum,
        },
        chart_arrangement::Error::Allocation { amount } => Error::Allocation { amount },
        _ => Error::InvalidSource,
    }
}

fn publish(source: &Package, target: &Target, rewritten: Vec<u8>) -> Result<Package, Error> {
    let source_catalog = source
        .state
        .components
        .physical()
        .ok_or(Error::InvalidSource)?;
    if !source_catalog.source_is_exact() {
        return Err(Refusal::UnsupportedSource.into());
    }
    let component = source
        .state
        .components
        .catalog()
        .get_index(target.component_index)
        .ok_or(Error::InvalidSource)?;
    if component.name() != target.component_name.as_ref() {
        return Err(Error::InvalidSource);
    }
    let archive_limits = source_catalog
        .limits()
        .effective_archive_limits()
        .map_err(map_archive_error)?;
    let snappy_limits = source_catalog
        .limits()
        .snappy_limits()
        .map_err(map_archive_error)?;
    let mut archive = component.archive().clone();
    archive
        .object_mut(target.object_identifier)
        .ok_or(Error::InvalidSource)?
        .replace_message_preserving_header(
            target.message_index,
            RawMessage {
                type_: target.message_type,
                data: rewritten,
            },
            archive_limits,
        )
        .map_err(map_core_error)?;
    let encoded = archive.to_bytes(archive_limits).map_err(map_core_error)?;
    let compressed = SnappyStream::compress(&encoded, snappy_limits).map_err(map_core_error)?;
    let edits = [EntryEdit::new(target.component_name.as_ref(), &compressed)];
    let prepared = source_catalog
        .prepare_reassembly_with_deletions(&edits, &[], source_catalog.limits())
        .map_err(map_archive_error)?;
    let execution_limits = prepared.execution_requirements().exact_limits();
    let output = prepared
        .execute(execution_limits)
        .map_err(map_archive_error)?;
    Package::from_owned_bytes_with_options(output, source.state.options)
        .map_err(map_candidate_error)
}

/// The immutable chart-text edit shared by title and axis-title facades.
pub struct Edit<'a> {
    source: &'a Package,
    target: Target,
    operation: Option<Option<Box<str>>>,
}

impl fmt::Debug for Edit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("sheet", &self.target.sheet)
            .field("chart", &self.target.chart)
            .field("kind", &self.target.kind)
            .field("staged", &self.operation.is_some())
            .finish_non_exhaustive()
    }
}

impl Edit<'_> {
    /// Text observed when the edit began.
    #[must_use]
    pub fn before(&self) -> Option<&str> {
        self.target.before.as_deref()
    }

    /// Text staged for publication.
    #[must_use]
    pub fn after(&self) -> Option<&str> {
        self.operation
            .as_ref()
            .map(|value| value.as_deref())
            .unwrap_or(self.target.before.as_deref())
    }

    /// Stage a visible title.
    pub fn set(mut self, text: impl AsRef<str>) -> Result<Self, Error> {
        if text
            .as_ref()
            .chars()
            .any(|character| matches!(character, '\u{fffc}' | '\u{000e}' | '\u{0004}'))
        {
            return Err(Refusal::ReservedMarker.into());
        }
        self.operation = Some(Some(text.as_ref().into()));
        Ok(self)
    }

    /// Stage removal of a visible title.
    pub fn clear(mut self) -> Result<Self, Error> {
        self.operation = Some(None);
        Ok(self)
    }

    /// Validate and publish the exact-source candidate.
    pub fn commit(self) -> Result<Commit, Error> {
        let after = self
            .operation
            .clone()
            .unwrap_or_else(|| self.target.before.clone());
        if after == self.target.before {
            return Ok(Commit::unchanged(self.source, self.target, after));
        }
        let current = resolve(
            self.source,
            SheetSelector::position(self.target.sheet),
            ChartSelector::index(self.target.chart.get()),
            self.target.kind,
        )?;
        if current != self.target {
            return Err(Error::PatchConflict);
        }
        let payload = current.payload(self.source)?;
        let rewritten = rewrite_payload(payload, current.kind, after.as_deref())?;
        if read_payload(&rewritten, current.kind)?.as_deref() != after.as_deref() {
            return Err(Error::Verification);
        }
        let package = publish(self.source, &current, rewritten)?;
        let verified = resolve(
            &package,
            SheetSelector::position(self.target.sheet),
            ChartSelector::index(self.target.chart.get()),
            self.target.kind,
        )?;
        if verified.before.as_deref() != after.as_deref() {
            return Err(Error::Verification);
        }
        let mut locality_budget =
            ChartBudget::for_package(self.source).map_err(map_arrangement_error)?;
        chart_arrangement::verify_message_locality(
            self.source,
            &package,
            current.component_name.as_ref(),
            current.component_index,
            current.object_index,
            current.object_identifier,
            current.message_index,
            current.message_type,
            verified.payload(&package)?,
            &mut locality_budget,
        )
        .map_err(map_arrangement_error)?;
        Commit::published(self.source, package, current, after)
    }
}

/// Exact-source reversible chart-text patch.
#[derive(Clone, PartialEq, Eq)]
pub struct Patch {
    artifacts: OwnedExactArtifacts,
    sheet: Position,
    chart: Position,
    kind: Kind,
    before: Option<Box<str>>,
    after: Option<Box<str>>,
}

impl fmt::Debug for Patch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Patch")
            .field("sheet", &self.sheet)
            .field("chart", &self.chart)
            .field("kind", &self.kind)
            .finish_non_exhaustive()
    }
}

impl Patch {
    /// Whether the text and exact artifact are unchanged.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before == self.after && self.artifacts.is_byte_noop()
    }

    /// Text observed before this patch.
    #[must_use]
    pub fn before(&self) -> Option<&str> {
        self.before.as_deref()
    }

    /// Text produced by this patch.
    #[must_use]
    pub fn after(&self) -> Option<&str> {
        self.after.as_deref()
    }

    /// The exact target-to-source inverse.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            artifacts: self.artifacts.inverse(),
            sheet: self.sheet,
            chart: self.chart,
            kind: self.kind,
            before: self.after.clone(),
            after: self.before.clone(),
        }
    }
}

/// Compact publication diagnostics.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Diagnostics {
    changed: bool,
}

impl Diagnostics {
    /// Whether the candidate changed package bytes.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }
}

/// A verified immutable chart-text commit.
#[must_use = "a chart-text commit contains the validated package snapshot"]
#[derive(Debug)]
pub struct Commit {
    package: Package,
    patch: Patch,
    diagnostics: Diagnostics,
}

impl Commit {
    /// Borrow the reopened package.
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.package
    }

    /// Consume the commit and return its package.
    #[must_use]
    pub fn into_package(self) -> Package {
        self.package
    }

    /// Consume the commit and return package plus patch.
    #[must_use]
    pub fn into_parts(self) -> (Package, Patch) {
        (self.package, self.patch)
    }

    /// Borrow the exact-source patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Borrow publication diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> &Diagnostics {
        &self.diagnostics
    }

    fn unchanged(source: &Package, target: Target, value: Option<Box<str>>) -> Self {
        let owner = source.state.source.clone();
        Self {
            package: source.clone(),
            patch: Patch {
                artifacts: OwnedExactArtifacts::new(owner.clone(), owner),
                sheet: target.sheet,
                chart: target.chart,
                kind: target.kind,
                before: target.before,
                after: value,
            },
            diagnostics: Diagnostics { changed: false },
        }
    }

    fn published(
        source: &Package,
        package: Package,
        target: Target,
        after: Option<Box<str>>,
    ) -> Result<Self, Error> {
        let source_owner = source.state.source.clone();
        let target_owner = package.state.source.clone();
        Ok(Self {
            package,
            patch: Patch {
                artifacts: OwnedExactArtifacts::new(source_owner, target_owner),
                sheet: target.sheet,
                chart: target.chart,
                kind: target.kind,
                before: target.before,
                after,
            },
            diagnostics: Diagnostics { changed: true },
        })
    }
}

/// Apply one exact-source chart-text patch.
pub(super) fn apply_patch(package: &Package, patch: &Patch) -> Result<Commit, Error> {
    if !patch.artifacts.authorizes_owner(&package.state.source) {
        return Err(Error::PatchConflict);
    }
    let current = resolve(
        package,
        SheetSelector::position(patch.sheet),
        ChartSelector::index(patch.chart.get()),
        patch.kind,
    )?;
    if current.before != patch.before {
        return Err(Error::PatchConflict);
    }
    if patch.is_noop() {
        return Ok(Commit::unchanged(package, current, patch.after.clone()));
    }
    let target = patch.artifacts.target_owner();
    let mut budget = ChartBudget::for_package(package).map_err(map_arrangement_error)?;
    budget
        .candidate_reopen(target.len())
        .map_err(map_arrangement_error)?;
    let candidate = Package::from_source_owner_with_options(target, package.state.options)
        .map_err(map_candidate_error)?;
    let verified = resolve(
        &candidate,
        SheetSelector::position(patch.sheet),
        ChartSelector::index(patch.chart.get()),
        patch.kind,
    )?;
    if verified.before != patch.after {
        return Err(Error::Verification);
    }
    let mut locality_budget = ChartBudget::for_package(package).map_err(map_arrangement_error)?;
    chart_arrangement::verify_message_locality(
        package,
        &candidate,
        current.component_name.as_ref(),
        current.component_index,
        current.object_index,
        current.object_identifier,
        current.message_index,
        current.message_type,
        verified.payload(&candidate)?,
        &mut locality_budget,
    )
    .map_err(map_arrangement_error)?;
    Ok(Commit {
        package: candidate,
        patch: patch.clone(),
        diagnostics: Diagnostics { changed: true },
    })
}

/// Read a selected chart text value.
pub(super) fn read<'sheet>(
    package: &Package,
    sheet: impl Into<SheetSelector<'sheet>>,
    chart: impl Into<ChartSelector>,
    kind: Kind,
) -> Result<Option<String>, Error> {
    Ok(resolve(package, sheet, chart, kind)?
        .before
        .map(|value| value.into()))
}

/// Begin a selected chart-text edit.
pub(super) fn edit<'a, 'sheet>(
    package: &'a Package,
    sheet: impl Into<SheetSelector<'sheet>>,
    chart: impl Into<ChartSelector>,
    kind: Kind,
) -> Result<Edit<'a>, Error> {
    Ok(Edit {
        target: resolve(package, sheet, chart, kind)?,
        source: package,
        operation: None,
    })
}

#[cfg(test)]
mod tests {
    use super::{Error, ReferenceOwners, Refusal};
    use litchi_iwa_common::{WireLimits, wire::append_length_delimited_field};

    fn chart_payload(identifier: u64) -> Vec<u8> {
        let mut reference = Vec::new();
        litchi_iwa_common::wire::append_varint_field(&mut reference, 1, identifier).unwrap();
        let mut chart = Vec::new();
        append_length_delimited_field(&mut chart, 10, &reference, WireLimits::DEFAULT).unwrap();
        let mut drawable = Vec::new();
        append_length_delimited_field(&mut drawable, 10_000, &chart, WireLimits::DEFAULT).unwrap();
        drawable
    }

    #[test]
    fn shared_reference_owner_is_refused() {
        let first =
            super::chart_codec::references(&chart_payload(77), super::chart_codec::Limits::DEFAULT)
                .unwrap();
        let second =
            super::chart_codec::references(&chart_payload(77), super::chart_codec::Limits::DEFAULT)
                .unwrap();
        let mut owners = ReferenceOwners::default();
        owners.add(&first, 77).unwrap();
        assert!(matches!(
            owners.add(&second, 77),
            Err(Error::Refused(Refusal::SharedObject))
        ));
    }
}
