//! Private, format-neutral graph authority for Keynote chart-axis owners.
//!
//! This module deliberately owns the security-sensitive chart graph proof:
//! selector resolution is delegated to the existing chart-title selector,
//! while native chart, axis, stand-in, stylesheet, metadata, and archive
//! locality checks stay private to the package adapter. A caller supplies its
//! own aggregate budget; no public API exposes IDs or wire objects.

#![allow(
    clippy::cast_sign_loss,
    clippy::map_err_ignore,
    clippy::needless_pass_by_value,
    clippy::shadow_reuse,
    clippy::wildcard_enum_match_arm,
    reason = "The graph adapter redacts lower-layer failures and keeps native graph details private."
)]

use std::collections::HashSet;

use litchi_core::Position;
use litchi_iwa_archive::{SourceCatalog, package::Entry};
use litchi_iwa_common::{
    WireLimits, decode_varint_from_bytes,
    varint::encoded_len,
    wire::{WireField, WireView, parse_wire_fields_with_limits},
};
use litchi_iwa_core::{
    Archive, ArchiveObject, ArchiveReferenceKind, ArchiveReferenceOccurrence,
    ArchiveReferencePolicy, ArchiveReferenceScope, ArchiveReferenceVisitor, RawMessage,
    SnappyStream,
};
use litchi_iwa_protos::package_metadata_codec;

use super::{Package, PhysicalSource};
use crate::{Axis, ChartSelector, SlideSelector};

const CHART_MESSAGE_TYPE: u32 = 5_021;
const CHART_NON_STYLE_MESSAGE_TYPE: u32 = 5_023;
const CHART_AXIS_MESSAGE_TYPE: u32 = 5_027;
const STANDIN_MESSAGE_TYPE: u32 = 3_097;
const STYLESHEET_MESSAGE_TYPE: u32 = 401;
const DRAWABLE_SUPER_FIELD: u32 = 1;
const DRAWABLE_TITLE_FIELD: u32 = 10;
const CHART_EXTENSION_FIELD: u32 = 10_000;
const CHART_AXIS_VALUE_FIELD: u32 = 14;
const CHART_AXIS_CATEGORY_FIELD: u32 = 16;
const DRAWABLE_LOCKED_FIELD: u32 = 5;
const PACKAGE_METADATA_MESSAGE_TYPE: u32 = 11_006;

/// Resource categories charged by shared graph scans.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum AxisSupportLimitKind {
    InputBytes,
    OutputBytes,
    WireBytes,
    Entries,
    EntryBytes,
    TotalBytes,
    Slides,
    References,
    TextStorages,
    TextFragments,
    TextBytes,
    WireFields,
    WireNesting,
    WireWork,
    TitleBytes,
}

/// Selector failures are kept neutral until the owning semantic adapter maps
/// them to its own error type.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum AxisSupportSelectorError {
    UnsupportedSource,
    AmbiguousSelector,
    EmptySlideName,
    SlideNameNotFound,
    SlidePositionNotFound { position: Position },
    ChartNameNotFound,
    ChartPositionNotFound { position: Position },
    EmptyChartName,
}

/// A lower-layer failure from the private graph proof. The owner maps this
/// to its public error without exposing graph or wire implementation details.
#[derive(Debug)]
pub(super) enum AxisSupportError {
    Selector(AxisSupportSelectorError),
    InvalidSource,
    LimitExceeded {
        kind: AxisSupportLimitKind,
        observed: u64,
        maximum: u64,
    },
    Allocation {
        amount: usize,
    },
}

/// A caller-owned aggregate budget. Implementations must charge before
/// allocating or retaining data and must keep all counters in one operation
/// ledger.
pub(super) trait AxisSupportBudget {
    fn charge_selection_scans(
        &mut self,
        package: &Package,
        mutation_guards: bool,
    ) -> Result<(), AxisSupportError>;
    #[allow(dead_code)]
    fn charge_input(&mut self, amount: usize) -> Result<(), AxisSupportError>;
    fn charge_wire_vector(&mut self, payload: usize) -> Result<(), AxisSupportError>;
    fn finish_wire_scan(&mut self, fields: usize) -> Result<(), AxisSupportError>;
    fn charge_reference_vector(&mut self, capacity: usize) -> Result<(), AxisSupportError>;
    fn charge_references(&mut self, amount: usize) -> Result<(), AxisSupportError>;
    fn charge_scan_pass(
        &mut self,
        package: &Package,
        retained_vectors: usize,
    ) -> Result<(), AxisSupportError>;
    fn charge_locality_scan(&mut self, package: &Package) -> Result<(), AxisSupportError>;
    fn charge_work(&mut self, amount: usize) -> Result<(), AxisSupportError>;
    fn metadata_options(
        &self,
        package: &Package,
    ) -> Result<package_metadata_codec::RewriteOptions, AxisSupportError>;
    fn charge_metadata_report(
        &mut self,
        report: package_metadata_codec::RewriteReport,
    ) -> Result<(), AxisSupportError>;
}

fn usize_to_u64(value: usize) -> u64 {
    u64::try_from(value).unwrap_or(u64::MAX)
}

pub(super) fn map_chart_title_error(
    error: super::slide_chart_title::ChartTitleError,
) -> AxisSupportError {
    use super::slide_chart_title::{ChartTitleError, ChartTitleLimitKind};
    match error {
        ChartTitleError::UnsupportedSource => {
            AxisSupportError::Selector(AxisSupportSelectorError::UnsupportedSource)
        },
        ChartTitleError::AmbiguousSelector => {
            AxisSupportError::Selector(AxisSupportSelectorError::AmbiguousSelector)
        },
        ChartTitleError::EmptySlideName => {
            AxisSupportError::Selector(AxisSupportSelectorError::EmptySlideName)
        },
        ChartTitleError::SlideNameNotFound => {
            AxisSupportError::Selector(AxisSupportSelectorError::SlideNameNotFound)
        },
        ChartTitleError::SlidePositionNotFound { position } => {
            AxisSupportError::Selector(AxisSupportSelectorError::SlidePositionNotFound { position })
        },
        ChartTitleError::ChartNameNotFound => {
            AxisSupportError::Selector(AxisSupportSelectorError::ChartNameNotFound)
        },
        ChartTitleError::ChartPositionNotFound { position } => {
            AxisSupportError::Selector(AxisSupportSelectorError::ChartPositionNotFound { position })
        },
        ChartTitleError::EmptyChartName => {
            AxisSupportError::Selector(AxisSupportSelectorError::EmptyChartName)
        },
        ChartTitleError::LimitExceeded {
            kind,
            observed,
            maximum,
        } => AxisSupportError::LimitExceeded {
            kind: match kind {
                ChartTitleLimitKind::InputBytes => AxisSupportLimitKind::InputBytes,
                ChartTitleLimitKind::OutputBytes => AxisSupportLimitKind::OutputBytes,
                ChartTitleLimitKind::WireBytes => AxisSupportLimitKind::WireBytes,
                ChartTitleLimitKind::Entries => AxisSupportLimitKind::Entries,
                ChartTitleLimitKind::EntryBytes => AxisSupportLimitKind::EntryBytes,
                ChartTitleLimitKind::TotalBytes => AxisSupportLimitKind::TotalBytes,
                ChartTitleLimitKind::Slides => AxisSupportLimitKind::Slides,
                ChartTitleLimitKind::References => AxisSupportLimitKind::References,
                ChartTitleLimitKind::TextStorages => AxisSupportLimitKind::TextStorages,
                ChartTitleLimitKind::TextFragments => AxisSupportLimitKind::TextFragments,
                ChartTitleLimitKind::TextBytes => AxisSupportLimitKind::TextBytes,
                ChartTitleLimitKind::WireFields => AxisSupportLimitKind::WireFields,
                ChartTitleLimitKind::WireNesting => AxisSupportLimitKind::WireNesting,
                ChartTitleLimitKind::WireWork => AxisSupportLimitKind::WireWork,
                ChartTitleLimitKind::TitleBytes => AxisSupportLimitKind::TitleBytes,
            },
            observed,
            maximum,
        },
        ChartTitleError::Allocation { amount } => AxisSupportError::Allocation { amount },
        ChartTitleError::InvalidSource
        | ChartTitleError::Verification
        | ChartTitleError::PatchConflict => AxisSupportError::InvalidSource,
    }
}

fn map_wire_error(error: litchi_iwa_common::Error) -> AxisSupportError {
    match error {
        litchi_iwa_common::Error::LimitExceeded {
            kind,
            observed,
            limit,
        } => AxisSupportError::LimitExceeded {
            kind: match kind {
                litchi_iwa_common::LimitKind::InputBytes => AxisSupportLimitKind::WireBytes,
                litchi_iwa_common::LimitKind::OutputBytes => AxisSupportLimitKind::OutputBytes,
                litchi_iwa_common::LimitKind::Fields
                | litchi_iwa_common::LimitKind::TableRows
                | litchi_iwa_common::LimitKind::TableColumns
                | litchi_iwa_common::LimitKind::TableCells
                | litchi_iwa_common::LimitKind::MaterializedCells => {
                    AxisSupportLimitKind::WireFields
                },
                litchi_iwa_common::LimitKind::Nesting => AxisSupportLimitKind::WireNesting,
                litchi_iwa_common::LimitKind::RewriteWork => AxisSupportLimitKind::WireWork,
            },
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(limit),
        },
        litchi_iwa_common::Error::Allocation { amount, .. } => {
            AxisSupportError::Allocation { amount }
        },
        litchi_iwa_common::Error::InvalidFormat(_)
        | litchi_iwa_common::Error::InvalidLimit { .. } => AxisSupportError::InvalidSource,
    }
}

fn map_archive_error(error: litchi_iwa_archive::Error) -> AxisSupportError {
    match error {
        litchi_iwa_archive::Error::Limit {
            kind,
            observed,
            maximum,
        } => AxisSupportError::LimitExceeded {
            kind: match kind {
                litchi_iwa_archive::LimitKind::InputBytes => AxisSupportLimitKind::InputBytes,
                litchi_iwa_archive::LimitKind::OutputBytes => AxisSupportLimitKind::OutputBytes,
                litchi_iwa_archive::LimitKind::Entries => AxisSupportLimitKind::Entries,
                litchi_iwa_archive::LimitKind::EntryBytes => AxisSupportLimitKind::EntryBytes,
                litchi_iwa_archive::LimitKind::TotalBytes => AxisSupportLimitKind::TotalBytes,
                _ => AxisSupportLimitKind::WireBytes,
            },
            observed,
            maximum,
        },
        litchi_iwa_archive::Error::Allocation { amount, .. } => {
            AxisSupportError::Allocation { amount }
        },
        _ => AxisSupportError::InvalidSource,
    }
}

fn map_core_error(error: litchi_iwa_core::Error) -> AxisSupportError {
    match error {
        litchi_iwa_core::Error::Limit {
            kind,
            observed,
            maximum,
        } => AxisSupportError::LimitExceeded {
            kind: match kind {
                litchi_iwa_core::LimitKind::ArchiveBytes => AxisSupportLimitKind::TotalBytes,
                litchi_iwa_core::LimitKind::Objects
                | litchi_iwa_core::LimitKind::Messages
                | litchi_iwa_core::LimitKind::MessagesPerObject
                | litchi_iwa_core::LimitKind::MetadataItems => AxisSupportLimitKind::Entries,
                litchi_iwa_core::LimitKind::ObjectBytes
                | litchi_iwa_core::LimitKind::MessageBytes
                | litchi_iwa_core::LimitKind::HeaderBytes
                | litchi_iwa_core::LimitKind::HeaderMemoryBytes
                | litchi_iwa_core::LimitKind::SnappyChunkBytes
                | litchi_iwa_core::LimitKind::SnappyStreamBytes
                | litchi_iwa_core::LimitKind::SnappyCompressedChunkBytes
                | litchi_iwa_core::LimitKind::SnappyCompressedStreamBytes => {
                    AxisSupportLimitKind::WireBytes
                },
                litchi_iwa_core::LimitKind::HeaderFields => AxisSupportLimitKind::WireFields,
                litchi_iwa_core::LimitKind::HeaderNesting => AxisSupportLimitKind::WireNesting,
                litchi_iwa_core::LimitKind::SnappyFrames => AxisSupportLimitKind::WireWork,
            },
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        },
        litchi_iwa_core::Error::Allocation { requested, .. } => {
            AxisSupportError::Allocation { amount: requested }
        },
        _ => AxisSupportError::InvalidSource,
    }
}

fn map_metadata_error(error: package_metadata_codec::RewriteError) -> AxisSupportError {
    if let Some(limit) = error.resource_limit() {
        let (kind, observed, maximum) = match limit {
            package_metadata_codec::RewriteLimit::InputBytes { observed, maximum } => {
                (AxisSupportLimitKind::WireBytes, observed, maximum)
            },
            package_metadata_codec::RewriteLimit::OutputBytes { observed, maximum } => {
                (AxisSupportLimitKind::OutputBytes, observed, maximum)
            },
            package_metadata_codec::RewriteLimit::Fields { observed, maximum } => {
                (AxisSupportLimitKind::WireFields, observed, maximum)
            },
            package_metadata_codec::RewriteLimit::Work { observed, maximum } => {
                (AxisSupportLimitKind::WireWork, observed, maximum)
            },
            package_metadata_codec::RewriteLimit::Nesting { observed, maximum } => {
                return AxisSupportError::LimitExceeded {
                    kind: AxisSupportLimitKind::WireNesting,
                    observed: u64::from(observed),
                    maximum: u64::from(maximum),
                };
            },
            package_metadata_codec::RewriteLimit::Components { observed, maximum }
            | package_metadata_codec::RewriteLimit::References { observed, maximum }
            | package_metadata_codec::RewriteLimit::Additions { observed, maximum } => {
                (AxisSupportLimitKind::References, observed, maximum)
            },
            _ => return AxisSupportError::InvalidSource,
        };
        return AxisSupportError::LimitExceeded {
            kind,
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        };
    }
    if let Some(amount) = error.allocation_request() {
        return AxisSupportError::Allocation { amount };
    }
    AxisSupportError::InvalidSource
}

pub(super) fn physical_catalog(package: &Package) -> Result<&SourceCatalog, AxisSupportError> {
    match &package.state.source {
        PhysicalSource::Package(source) => Ok(source),
        PhysicalSource::Semantic(_) => Err(AxisSupportError::Selector(
            AxisSupportSelectorError::UnsupportedSource,
        )),
    }
}

/// Validate canonical object framing after reserving the bounded per-object
/// varint work in the caller's aggregate ledger.
pub(super) fn validate_canonical_object_length_prefixes_with_budget(
    source: &[u8],
    archive: &Archive,
    budget: &mut dyn AxisSupportBudget,
) -> Result<(), AxisSupportError> {
    budget.charge_work(
        archive
            .objects
            .len()
            .checked_mul(10)
            .ok_or(AxisSupportError::InvalidSource)?,
    )?;
    validate_canonical_object_length_prefixes_unbudgeted(source, archive)
}

fn validate_canonical_object_length_prefixes_unbudgeted(
    source: &[u8],
    archive: &Archive,
) -> Result<(), AxisSupportError> {
    for object in &archive.objects {
        let offset =
            usize::try_from(object.header_offset).map_err(|_| AxisSupportError::InvalidSource)?;
        let remaining = source
            .get(offset..)
            .ok_or(AxisSupportError::InvalidSource)?;
        let (header_bytes, prefix_bytes) =
            decode_varint_from_bytes(remaining).map_err(|_| AxisSupportError::InvalidSource)?;
        if prefix_bytes != encoded_len(header_bytes) {
            return Err(AxisSupportError::InvalidSource);
        }
        let framed_header_bytes = header_bytes
            .checked_add(u64::try_from(prefix_bytes).map_err(|_| AxisSupportError::InvalidSource)?)
            .ok_or(AxisSupportError::InvalidSource)?;
        if framed_header_bytes != object.header_length
            || object
                .header_offset
                .checked_add(object.header_length)
                .ok_or(AxisSupportError::InvalidSource)?
                != object.data_offset
        {
            return Err(AxisSupportError::InvalidSource);
        }
    }
    Ok(())
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub(super) struct AxisSelection {
    pub(super) slide_position: Position,
    pub(super) chart_position: Position,
    pub(super) slide_identifier: u64,
    pub(super) chart_identifier: u64,
    pub(super) non_style_identifier: u64,
    pub(super) chart_component_name: String,
    pub(super) chart_message_index: usize,
    pub(super) axis: Axis,
    pub(super) axis_identifier: u64,
    pub(super) axis_component_name: String,
    pub(super) axis_message_index: usize,
}

pub(super) fn select_axis(
    package: &Package,
    slide_selector: SlideSelector<'_>,
    chart_selector: ChartSelector<'_>,
    axis: Axis,
    mutation_guards: bool,
    budget: &mut dyn AxisSupportBudget,
) -> Result<AxisSelection, AxisSupportError> {
    budget.charge_selection_scans(package, mutation_guards)?;
    let chart = super::slide_chart_title::select_chart_with_budget(
        package,
        slide_selector,
        chart_selector,
        mutation_guards,
        budget,
    )?;
    let (chart_component, chart_object) = package
        .object_with_component(chart.chart_identifier)
        .ok_or(AxisSupportError::InvalidSource)?;
    if chart_component != chart.slide_component_name {
        return Err(AxisSupportError::InvalidSource);
    }
    let (chart_message_index, chart_message) =
        unique_message(chart_object, CHART_MESSAGE_TYPE, budget)?;
    validate_selected_message_metadata(chart_object, chart_message_index)?;
    let limits = package.wire_limits().map_err(map_wire_error)?;
    let outer = accounted_wire_fields(&chart_message.data, limits, budget)?;
    let super_payload =
        unique_length_delimited_field(&outer, &chart_message.data, DRAWABLE_SUPER_FIELD, budget)?
            .ok_or(AxisSupportError::InvalidSource)?;
    if mutation_guards {
        validate_unlocked_drawable(super_payload, limits, budget)?;
    }
    let title_identifier = required_reference(super_payload, DRAWABLE_TITLE_FIELD, limits, budget)?;
    validate_graph_object(package, title_identifier, STANDIN_MESSAGE_TYPE, budget)?;
    validate_graph_object(
        package,
        chart.non_style_identifier,
        CHART_NON_STYLE_MESSAGE_TYPE,
        budget,
    )?;
    let chart_payload =
        unique_length_delimited_field(&outer, &chart_message.data, CHART_EXTENSION_FIELD, budget)?
            .ok_or(AxisSupportError::InvalidSource)?;
    let category = repeated_references(chart_payload, CHART_AXIS_CATEGORY_FIELD, limits, budget)?;
    let value = repeated_references(chart_payload, CHART_AXIS_VALUE_FIELD, limits, budget)?;
    validate_axis_roles(&category, &value, budget)?;
    let selected = match axis {
        Axis::Category => &category,
        Axis::Value => &value,
    };
    budget.charge_work(selected.len())?;
    let axis_identifier = selected
        .iter()
        .copied()
        .find(|identifier| *identifier != 0)
        .ok_or(AxisSupportError::InvalidSource)?;
    let (axis_component, axis_object) = unique_object(package, axis_identifier)?;
    if axis_object.messages.len() != 1 {
        return Err(AxisSupportError::InvalidSource);
    }
    let (axis_message_index, _) = unique_message(axis_object, CHART_AXIS_MESSAGE_TYPE, budget)?;
    validate_selected_message_metadata(axis_object, axis_message_index)?;
    if mutation_guards {
        prove_unique_primary_axis(package, axis_identifier, budget)?;
        let stylesheet_component_name = validate_global_axis_references(
            package,
            chart.chart_identifier,
            chart_message_index,
            axis_identifier,
            budget,
        )?;
        validate_axis_metadata(
            package,
            chart_component,
            axis_component,
            stylesheet_component_name,
            axis_identifier,
            budget,
        )?;
    }
    if axis_identifier == chart.chart_identifier
        || axis_identifier == chart.non_style_identifier
        || axis_identifier == chart.slide_identifier
    {
        return Err(AxisSupportError::InvalidSource);
    }
    Ok(AxisSelection {
        slide_position: chart.slide_position,
        chart_position: chart.chart_position,
        slide_identifier: chart.slide_identifier,
        chart_identifier: chart.chart_identifier,
        non_style_identifier: chart.non_style_identifier,
        chart_component_name: chart_component.to_owned(),
        chart_message_index,
        axis,
        axis_identifier,
        axis_component_name: axis_component.to_owned(),
        axis_message_index,
    })
}

pub(super) fn repeated_references(
    payload: &[u8],
    field_number: u32,
    limits: WireLimits,
    budget: &mut dyn AxisSupportBudget,
) -> Result<Vec<u64>, AxisSupportError> {
    // `WireView` retains a field/span vector just like the general wire
    // parser. Reserve and charge that vector before parsing, then charge the
    // decoded field scan below. This keeps repeated-reference parsing on the
    // same ledger as `accounted_wire_fields`.
    budget.charge_wire_vector(payload.len())?;
    let fields = WireView::parse_with_limits(payload, limits).map_err(map_wire_error)?;
    budget.finish_wire_scan(fields.len())?;
    let mut references = Vec::new();
    budget.charge_reference_vector(fields.len())?;
    references
        .try_reserve(fields.len())
        .map_err(|_error| AxisSupportError::Allocation {
            amount: fields.len(),
        })?;
    for field in fields.fields() {
        if field.number() != field_number {
            continue;
        }
        if field.wire_type() != 2 {
            return Err(AxisSupportError::InvalidSource);
        }
        field.validate_canonical_framing().map_err(map_wire_error)?;
        references.push(validate_reference_payload(field.payload(), limits, budget)?);
    }
    budget.charge_references(references.len())?;
    Ok(references)
}

pub(super) fn accounted_wire_fields(
    payload: &[u8],
    limits: WireLimits,
    budget: &mut dyn AxisSupportBudget,
) -> Result<Vec<WireField>, AxisSupportError> {
    budget.charge_wire_vector(payload.len())?;
    let fields = parse_wire_fields_with_limits(payload, limits).map_err(map_wire_error)?;
    budget.finish_wire_scan(fields.len())?;
    Ok(fields)
}

pub(super) fn unique_length_delimited_field<'a>(
    fields: &[WireField],
    source: &'a [u8],
    number: u32,
    budget: &mut dyn AxisSupportBudget,
) -> Result<Option<&'a [u8]>, AxisSupportError> {
    // The caller normally supplies fields from `accounted_wire_fields`, but
    // keep this reusable helper independently bounded: every field is
    // compared before duplicate/type rejection can terminate the scan.
    budget.charge_work(fields.len())?;
    let mut selected = None;
    for field in fields
        .iter()
        .copied()
        .filter(|field| field.number() == number)
    {
        if selected.is_some() || field.wire_type() != 2 {
            return Err(AxisSupportError::InvalidSource);
        }
        field
            .validate_canonical_framing(source)
            .map_err(map_wire_error)?;
        selected = Some(field.payload(source).map_err(map_wire_error)?);
    }
    Ok(selected)
}

pub(super) fn unique_message<'a>(
    object: &'a ArchiveObject,
    message_type: u32,
    budget: &mut dyn AxisSupportBudget,
) -> Result<(usize, &'a RawMessage), AxisSupportError> {
    // This helper must inspect every message to reject duplicate typed
    // payloads. Charge the full comparison bound before entering the scan so
    // malformed packages cannot spend unaccounted work on a late duplicate.
    budget.charge_work(object.messages.len())?;
    let mut selected = None;
    for (index, message) in object.messages.iter().enumerate() {
        if message.type_ != message_type {
            continue;
        }
        if selected.replace((index, message)).is_some() {
            return Err(AxisSupportError::InvalidSource);
        }
    }
    selected.ok_or(AxisSupportError::InvalidSource)
}

pub(super) fn required_reference(
    payload: &[u8],
    field_number: u32,
    limits: WireLimits,
    budget: &mut dyn AxisSupportBudget,
) -> Result<u64, AxisSupportError> {
    let fields = accounted_wire_fields(payload, limits, budget)?;
    let reference = unique_length_delimited_field(&fields, payload, field_number, budget)?
        .ok_or(AxisSupportError::InvalidSource)?;
    let identifier = validate_reference_payload(reference, limits, budget)?;
    if identifier == 0 {
        return Err(AxisSupportError::InvalidSource);
    }
    Ok(identifier)
}

pub(super) fn validate_graph_object(
    package: &Package,
    identifier: u64,
    message_type: u32,
    budget: &mut dyn AxisSupportBudget,
) -> Result<(), AxisSupportError> {
    let (_component, object) = unique_object(package, identifier)?;
    if object.messages.len() != 1 {
        return Err(AxisSupportError::InvalidSource);
    }
    let (message_index, _message) = unique_message(object, message_type, budget)?;
    validate_selected_message_metadata(object, message_index)
}

pub(super) fn validate_unlocked_drawable(
    payload: &[u8],
    limits: WireLimits,
    budget: &mut dyn AxisSupportBudget,
) -> Result<(), AxisSupportError> {
    budget.charge_wire_vector(payload.len())?;
    let view = WireView::parse_with_limits(payload, limits).map_err(map_wire_error)?;
    budget.finish_wire_scan(view.len())?;
    let mut locked = None;
    for field in view
        .fields()
        .filter(|field| field.number() == DRAWABLE_LOCKED_FIELD)
    {
        if locked.is_some() || field.wire_type() != 0 {
            return Err(AxisSupportError::InvalidSource);
        }
        field.validate_canonical_framing().map_err(map_wire_error)?;
        let (value, bytes) = decode_varint_from_bytes(field.payload())
            .map_err(|_error| AxisSupportError::InvalidSource)?;
        if bytes != field.payload().len() || value > 1 {
            return Err(AxisSupportError::InvalidSource);
        }
        locked = Some(value != 0);
    }
    if locked == Some(true) {
        return Err(AxisSupportError::InvalidSource);
    }
    Ok(())
}

/// Validate one nested archive reference after reserving the wire parser's
/// span/vector allocation. The package-level validator is intentionally
/// reused for its canonical identifier/deprecated-field rules; this wrapper
/// accounts its second-level parse without exposing those low-level details
/// to either semantic owner.
pub(super) fn validate_reference_payload(
    payload: &[u8],
    limits: WireLimits,
    budget: &mut dyn AxisSupportBudget,
) -> Result<u64, AxisSupportError> {
    budget.charge_wire_vector(payload.len())?;
    // A malformed payload can contain one minimal field per byte. Charge the
    // conservative field-scan bound before delegating to the package-level
    // validator, whose parser is deliberately kept private to `package`.
    budget.finish_wire_scan(payload.len())?;
    super::validate_reference_payload(payload, limits, "Keynote chart drawable")
        .map_err(map_wire_error)
}

pub(super) fn validate_axis_roles(
    category: &[u64],
    value: &[u64],
    budget: &mut dyn AxisSupportBudget,
) -> Result<(), AxisSupportError> {
    let roles = category
        .len()
        .checked_add(value.len())
        .ok_or(AxisSupportError::InvalidSource)?;
    budget.charge_work(
        roles
            .checked_mul(2)
            .ok_or(AxisSupportError::InvalidSource)?,
    )?;
    let category_primary = category.iter().copied().find(|identifier| *identifier != 0);
    let value_primary = value.iter().copied().find(|identifier| *identifier != 0);
    if category_primary.is_none() || value_primary.is_none() {
        return Err(AxisSupportError::InvalidSource);
    }
    let mut seen = HashSet::new();
    budget.charge_reference_vector(roles)?;
    seen.try_reserve(roles)
        .map_err(|_error| AxisSupportError::Allocation { amount: roles })?;
    for identifier in category.iter().chain(value).copied() {
        if identifier != 0 && !seen.insert(identifier) {
            return Err(AxisSupportError::InvalidSource);
        }
    }
    Ok(())
}

pub(super) fn unique_object(
    package: &Package,
    identifier: u64,
) -> Result<(&str, &ArchiveObject), AxisSupportError> {
    // `Package` builds and validates one sorted object index at ingestion,
    // rejecting duplicate native identities before exposing this lookup. Use
    // that index here: rescanning every component for each graph edge was
    // both redundant and an easy source of under-accounted quadratic work.
    package
        .object_with_component(identifier)
        .ok_or(AxisSupportError::InvalidSource)
}

pub(super) fn validate_selected_message_metadata(
    object: &ArchiveObject,
    message_index: usize,
) -> Result<(), AxisSupportError> {
    let message = object
        .messages
        .get(message_index)
        .ok_or(AxisSupportError::InvalidSource)?;
    let info = object
        .archive_info
        .message_infos
        .get(message_index)
        .ok_or(AxisSupportError::InvalidSource)?;
    if object.archive_info.should_merge == Some(true)
        || info.type_ != message.type_
        || usize::try_from(info.length).ok() != Some(message.data.len())
        || info.base_message_index.is_some()
        || !info.diff_merge_version.is_empty()
        || info.diff_field_path.is_some()
        || !info.fields_to_remove.is_empty()
        || !info.diff_read_version.is_empty()
    {
        return Err(AxisSupportError::InvalidSource);
    }
    Ok(())
}

pub(super) fn prove_unique_primary_axis(
    package: &Package,
    axis_identifier: u64,
    budget: &mut dyn AxisSupportBudget,
) -> Result<(), AxisSupportError> {
    budget.charge_scan_pass(package, 0)?;
    let limits = package.wire_limits().map_err(map_wire_error)?;
    let mut primary_owners = 0usize;
    for component in package.state.source.components().iter() {
        for object in &component.archive().objects {
            for message in &object.messages {
                if message.type_ != CHART_MESSAGE_TYPE {
                    continue;
                }
                let outer = accounted_wire_fields(&message.data, limits, budget)?;
                let chart = unique_length_delimited_field(
                    &outer,
                    &message.data,
                    CHART_EXTENSION_FIELD,
                    budget,
                )?
                .ok_or(AxisSupportError::InvalidSource)?;
                let category =
                    repeated_references(chart, CHART_AXIS_CATEGORY_FIELD, limits, budget)?;
                let value = repeated_references(chart, CHART_AXIS_VALUE_FIELD, limits, budget)?;
                validate_axis_roles(&category, &value, budget)?;
                let role_scan = category
                    .len()
                    .checked_add(value.len())
                    .and_then(|roles| roles.checked_mul(2))
                    .ok_or(AxisSupportError::InvalidSource)?;
                budget.charge_work(role_scan)?;
                for roles in [&category, &value] {
                    let primary = roles.iter().copied().find(|identifier| *identifier != 0);
                    if primary == Some(axis_identifier) {
                        primary_owners = primary_owners
                            .checked_add(1)
                            .ok_or(AxisSupportError::InvalidSource)?;
                    } else if roles.contains(&axis_identifier) {
                        return Err(AxisSupportError::InvalidSource);
                    }
                }
            }
        }
    }
    if primary_owners == 1 {
        Ok(())
    } else {
        Err(AxisSupportError::InvalidSource)
    }
}

/// Charge the fixed package envelope before a package-wide metadata loop
/// starts. `charge_scan_pass` is implemented by each semantic owner by
/// building an inventory, so calling it immediately before another loop
/// would inspect the same package twice without charging the second pass.
/// The source byte bound covers the byte-sized framing/callback work; the
/// parsed object count covers the package index walk. Message counts are
/// charged at each object immediately before the lower-layer scan, because
/// they are not retained in [`Package::state`].
fn charge_package_metadata_loop_bound(
    package: &Package,
    budget: &mut dyn AxisSupportBudget,
) -> Result<(), AxisSupportError> {
    let source_bytes = match &package.state.source {
        PhysicalSource::Package(source) => source.source_bytes().len(),
        PhysicalSource::Semantic(_) => 0,
    };
    let bound = source_bytes
        .checked_add(package.state.source.components().len())
        .and_then(|value| value.checked_add(package.state.total_objects))
        .ok_or(AxisSupportError::InvalidSource)?;
    budget.charge_work(bound)
}

/// Account one metadata-info traversal and the following source-authoritative
/// equality/reference traversal. Lengths are read before either traversal is
/// entered, and every nested vector contributes its own checked byte/item
/// bound. `repetitions` is two for the local accounting pass plus the
/// lower-layer pass that follows it.
fn charge_message_info_metadata_work(
    info: &litchi_iwa_core::MessageInfo,
    repetitions: usize,
    budget: &mut dyn AxisSupportBudget,
) -> Result<usize, AxisSupportError> {
    let mut work = 16usize;
    work = work
        .checked_add(info.versions.len())
        .and_then(|value| value.checked_add(info.object_references.len()))
        .and_then(|value| value.checked_add(info.data_references.len()))
        .and_then(|value| value.checked_add(info.diff_merge_version.len()))
        .and_then(|value| value.checked_add(info.diff_read_version.len()))
        .and_then(|value| value.checked_add(info.field_infos.len()))
        .and_then(|value| value.checked_add(usize::from(info.diff_field_path.is_some())))
        .and_then(|value| value.checked_add(info.fields_to_remove.len()))
        .ok_or(AxisSupportError::InvalidSource)?;
    if let Some(path) = &info.diff_field_path {
        work = work
            .checked_add(path.path.len())
            .ok_or(AxisSupportError::InvalidSource)?;
    }
    for path in &info.fields_to_remove {
        work = work
            .checked_add(path.path.len())
            .ok_or(AxisSupportError::InvalidSource)?;
    }
    budget.charge_work(
        work.checked_mul(repetitions)
            .ok_or(AxisSupportError::InvalidSource)?,
    )?;
    for field in &info.field_infos {
        let field_work = 8usize
            .checked_add(field.path.path.len())
            .and_then(|value| value.checked_add(field.object_references.len()))
            .and_then(|value| value.checked_add(field.data_references.len()))
            .and_then(|value| value.checked_add(field.known_field_version.len()))
            .and_then(|value| {
                field
                    .known_field_feature_identifier
                    .as_ref()
                    .map_or(Some(value), |feature| value.checked_add(feature.len()))
            })
            .ok_or(AxisSupportError::InvalidSource)?;
        budget.charge_work(
            field_work
                .checked_mul(repetitions)
                .ok_or(AxisSupportError::InvalidSource)?,
        )?;
        work = work
            .checked_add(field_work)
            .ok_or(AxisSupportError::InvalidSource)?;
    }
    Ok(work)
}

/// Charge the complete archive-reference inspection for one object before
/// invoking the core visitor. The core path canonicalizes/decode-checks each
/// ArchiveInfo and compares the decoded metadata with the retained object;
/// those buffers and equality bytes are not represented by an
/// `AxisSupportBudget` report, so reserve them explicitly here.
fn charge_archive_reference_inspection(
    object: &ArchiveObject,
    budget: &mut dyn AxisSupportBudget,
) -> Result<(), AxisSupportError> {
    let mut metadata_work = 16usize;
    let mut references = 0usize;
    for info in &object.archive_info.message_infos {
        metadata_work = metadata_work
            .checked_add(charge_message_info_metadata_work(info, 2, budget)?)
            .ok_or(AxisSupportError::InvalidSource)?;
        references = references
            .checked_add(info.object_references.len())
            .and_then(|value| value.checked_add(info.data_references.len()))
            .ok_or(AxisSupportError::InvalidSource)?;
        for field in &info.field_infos {
            references = references
                .checked_add(field.object_references.len())
                .and_then(|value| value.checked_add(field.data_references.len()))
                .ok_or(AxisSupportError::InvalidSource)?;
        }
    }
    let source_header_bytes =
        usize::try_from(object.header_length).map_err(|_| AxisSupportError::InvalidSource)?;
    // Every metadata item can require multiple bytes in canonical protobuf
    // framing. Eight is deliberately conservative while remaining tied to
    // the parsed object rather than a global maximum-header allocation.
    let metadata_header_bytes = metadata_work
        .checked_mul(8)
        .ok_or(AxisSupportError::InvalidSource)?;
    let header_bytes = source_header_bytes.max(metadata_header_bytes).max(1);
    budget.charge_work(
        header_bytes
            .checked_mul(8)
            .and_then(|value| value.checked_add(object.messages.len().checked_mul(2)?))
            .ok_or(AxisSupportError::InvalidSource)?,
    )?;
    // Canonical encoding and the bounded Buffa decode each retain a
    // header-sized logical buffer. `charge_reference_vector` is the shared
    // interface's allocation/retained/scratch primitive; round up from bytes
    // to its u64 capacity without allowing a zero-capacity special case.
    let buffer_bytes = header_bytes
        .checked_mul(4)
        .ok_or(AxisSupportError::InvalidSource)?;
    let buffer_capacity = buffer_bytes.div_ceil(size_of::<u64>()).max(1);
    budget.charge_reference_vector(buffer_capacity)?;
    budget.charge_reference_vector(buffer_capacity)?;
    budget.charge_references(references)
}

struct StrictArchiveReferenceVisitor;

impl ArchiveReferenceVisitor for StrictArchiveReferenceVisitor {
    fn visit_reference(
        &mut self,
        _occurrence: ArchiveReferenceOccurrence,
    ) -> litchi_iwa_core::Result<()> {
        Ok(())
    }
}

/// Inspect one archive object with the deletion-grade metadata policy.
///
/// Callers that perform their own ownership census still need the complete
/// source metadata proof: an unknown ArchiveInfo, MessageInfo, or FieldInfo
/// field may contain an owner edge the focused census cannot see. Charge the
/// same retained header/reference work as the axis proof before invoking the
/// core inspector, then discard known occurrences after the strict policy has
/// established that the projection is complete.
pub(super) fn inspect_archive_references_strict(
    package: &Package,
    object: &ArchiveObject,
    budget: &mut dyn AxisSupportBudget,
) -> Result<(), AxisSupportError> {
    budget.charge_work(1)?;
    charge_archive_reference_inspection(object, budget)?;
    let archive_limits = package
        .state
        .options
        .archive()
        .effective_archive_limits()
        .map_err(map_archive_error)?;
    let mut visitor = StrictArchiveReferenceVisitor;
    object
        .inspect_references_with_policy_and_limits(
            &mut visitor,
            ArchiveReferencePolicy::RejectUnknownMetadata,
            archive_limits,
        )
        .map_err(map_core_error)?;
    Ok(())
}

/// Charge the aggregate/field reference census used to reject aliases before
/// the lower-layer visitor starts. The census compares each identifier value,
/// so count both the vector walk and each value comparison.
fn charge_reference_census(
    info: &litchi_iwa_core::MessageInfo,
    budget: &mut dyn AxisSupportBudget,
) -> Result<(), AxisSupportError> {
    budget.charge_work(
        info.object_references
            .len()
            .checked_add(info.data_references.len())
            .and_then(|value| value.checked_add(info.field_infos.len()))
            .and_then(|value| value.checked_add(1))
            .ok_or(AxisSupportError::InvalidSource)?,
    )?;
    for field in &info.field_infos {
        budget.charge_work(
            field
                .object_references
                .len()
                .checked_add(field.data_references.len())
                .and_then(|value| value.checked_add(1))
                .ok_or(AxisSupportError::InvalidSource)?,
        )?;
    }
    Ok(())
}

pub(super) fn validate_global_axis_references<'a>(
    package: &'a Package,
    chart_identifier: u64,
    chart_message_index: usize,
    axis_identifier: u64,
    budget: &mut dyn AxisSupportBudget,
) -> Result<Option<&'a str>, AxisSupportError> {
    charge_package_metadata_loop_bound(package, budget)?;
    let chart_object = package
        .object(chart_identifier)
        .ok_or(AxisSupportError::InvalidSource)?;
    let chart_info = chart_object
        .archive_info
        .message_infos
        .get(chart_message_index)
        .ok_or(AxisSupportError::InvalidSource)?;
    charge_reference_census(chart_info, budget)?;
    let aggregate_edges = chart_info
        .object_references
        .iter()
        .filter(|identifier| **identifier == axis_identifier)
        .count();
    let aggregate_data_edges = chart_info
        .data_references
        .iter()
        .filter(|identifier| **identifier == axis_identifier)
        .count();
    let mut field_edges = 0usize;
    let mut field_data_edges = 0usize;
    for field in &chart_info.field_infos {
        field_edges = field_edges
            .checked_add(
                field
                    .object_references
                    .iter()
                    .filter(|identifier| **identifier == axis_identifier)
                    .count(),
            )
            .ok_or(AxisSupportError::InvalidSource)?;
        field_data_edges = field_data_edges
            .checked_add(
                field
                    .data_references
                    .iter()
                    .filter(|identifier| **identifier == axis_identifier)
                    .count(),
            )
            .ok_or(AxisSupportError::InvalidSource)?;
    }
    if aggregate_edges != 1
        || aggregate_data_edges != 0
        || field_edges != 0
        || field_data_edges != 0
    {
        return Err(AxisSupportError::InvalidSource);
    }
    let archive_limits = package
        .state
        .options
        .archive()
        .effective_archive_limits()
        .map_err(map_archive_error)?;
    let mut visitor = AxisInboundReferenceVisitor {
        package,
        chart_identifier,
        chart_message_index,
        axis_identifier,
        selected_references: 0,
        stylesheet_references: 0,
        stylesheet_component_name: None,
        invalid: false,
    };
    for component in package.state.source.components().iter() {
        for object in &component.archive().objects {
            budget.charge_work(1)?;
            charge_archive_reference_inspection(object, budget)?;
            object
                .inspect_references_with_policy_and_limits(
                    &mut visitor,
                    ArchiveReferencePolicy::RejectUnknownMetadata,
                    archive_limits,
                )
                .map_err(map_core_error)?;
        }
    }
    if visitor.invalid || visitor.selected_references != 1 {
        return Err(AxisSupportError::InvalidSource);
    }
    Ok(visitor.stylesheet_component_name)
}

struct AxisInboundReferenceVisitor<'a> {
    package: &'a Package,
    chart_identifier: u64,
    chart_message_index: usize,
    axis_identifier: u64,
    selected_references: usize,
    stylesheet_references: usize,
    stylesheet_component_name: Option<&'a str>,
    invalid: bool,
}

impl ArchiveReferenceVisitor for AxisInboundReferenceVisitor<'_> {
    fn visit_reference(
        &mut self,
        occurrence: ArchiveReferenceOccurrence,
    ) -> litchi_iwa_core::Result<()> {
        // Data references occupy a separate namespace, so unrelated chart
        // datasets do not participate in object ownership.  Still reject a
        // numeric collision with the selected axis: this focused owner does
        // not rewrite or prove that ambiguous cross-namespace relationship.
        if occurrence.kind == ArchiveReferenceKind::Data
            && occurrence.referenced_identifier == self.axis_identifier
        {
            self.invalid = true;
        }
        if occurrence.referenced_identifier == self.axis_identifier {
            if occurrence.kind == ArchiveReferenceKind::Object
                && occurrence.scope == ArchiveReferenceScope::Message
                && occurrence.object_identifier == self.chart_identifier
                && occurrence.message_index == self.chart_message_index
            {
                if let Some(count) = self.selected_references.checked_add(1) {
                    self.selected_references = count;
                } else {
                    self.invalid = true;
                }
            } else if let Some(component_name) =
                stylesheet_registration_component(self.package, occurrence)
            {
                if let Some(count) = self.stylesheet_references.checked_add(1) {
                    self.stylesheet_references = count;
                } else {
                    self.invalid = true;
                }
                self.stylesheet_component_name = Some(component_name);
                if self.stylesheet_references > 1 {
                    self.invalid = true;
                }
            } else {
                self.invalid = true;
            }
        }
        Ok(())
    }
}

pub(super) fn stylesheet_registration_component(
    package: &Package,
    occurrence: ArchiveReferenceOccurrence,
) -> Option<&str> {
    if occurrence.kind != ArchiveReferenceKind::Object
        || occurrence.scope != ArchiveReferenceScope::Message
    {
        return None;
    }
    let (component, object) = package.object_with_component(occurrence.object_identifier)?;
    (object.messages.len() == 1
        && object
            .messages
            .get(occurrence.message_index)
            .is_some_and(|message| message.type_ == STYLESHEET_MESSAGE_TYPE)
        && validate_selected_message_metadata(object, occurrence.message_index).is_ok())
    .then_some(component)
}

pub(super) fn validate_axis_metadata(
    package: &Package,
    chart_component_name: &str,
    axis_component_name: &str,
    stylesheet_component_name: Option<&str>,
    axis_identifier: u64,
    budget: &mut dyn AxisSupportBudget,
) -> Result<(), AxisSupportError> {
    let Some(metadata) = package_metadata_payload(package, budget)? else {
        return Ok(());
    };
    let owner_external_component_name =
        (chart_component_name != axis_component_name).then_some(chart_component_name);
    let registry_external_component_name = stylesheet_component_name.filter(|component_name| {
        *component_name != axis_component_name
            && Some(*component_name) != owner_external_component_name
    });
    let mut selected = SelectedAxisMetadataVisitor {
        axis_component_name,
        owner_external_component_name,
        registry_external_component_name,
        axis_identifier,
        selected_uuid: None,
        selected_component_identifier: None,
        external_target_component_identifier: None,
        owner_external_seen: false,
        registry_external_seen: false,
        count: 0,
        invalid: false,
    };
    charge_metadata_inspection_envelope(metadata, budget)?;
    let options = budget.metadata_options(package)?;
    let inspection = package_metadata_codec::inspect_package_metadata_with_visitor(
        metadata,
        options,
        &mut selected,
    )
    .map_err(map_metadata_error)?;
    budget.charge_metadata_report(inspection.report())?;
    let expects_external =
        owner_external_component_name.is_some() || registry_external_component_name.is_some();
    if selected.owner_external_seen != owner_external_component_name.is_some()
        || selected.registry_external_seen != registry_external_component_name.is_some()
        || (expects_external
            && selected.external_target_component_identifier
                != selected.selected_component_identifier)
        || (!expects_external && selected.external_target_component_identifier.is_some())
    {
        selected.invalid = true;
    }
    let selected_uuid = selected
        .selected_uuid
        .filter(|_uuid| !selected.invalid && selected.count == 1)
        .ok_or(AxisSupportError::InvalidSource)?;

    let mut authority = AxisMetadataAuthorityVisitor {
        axis_identifier,
        selected_uuid,
        selected_pair_count: 0,
        selected_object_count: 0,
        invalid: false,
    };
    charge_metadata_inspection_envelope(metadata, budget)?;
    let options = budget.metadata_options(package)?;
    let inspection = package_metadata_codec::inspect_package_metadata_with_visitor(
        metadata,
        options,
        &mut authority,
    )
    .map_err(map_metadata_error)?;
    budget.charge_metadata_report(inspection.report())?;
    if authority.invalid
        || authority.selected_pair_count != 1
        || authority.selected_object_count != 1
    {
        return Err(AxisSupportError::InvalidSource);
    }
    Ok(())
}

/// Reserve the work and temporary buffers that the strict metadata inspector
/// performs outside its public [`RewriteReport`]. Inspection has a canonical
/// parser, a parity decode, and a second visitor pass; callbacks also compare
/// source component locators byte-for-byte. Keep this precharge before options
/// are derived so the codec receives only the remaining owner budget.
fn charge_metadata_inspection_envelope(
    metadata: &[u8],
    budget: &mut dyn AxisSupportBudget,
) -> Result<(), AxisSupportError> {
    let bytes = metadata.len().max(1);
    budget.charge_work(
        bytes
            .checked_mul(8)
            .and_then(|value| value.checked_add(64))
            .ok_or(AxisSupportError::InvalidSource)?,
    )?;
    // The codec's canonical/decode views are borrowed in the report and do
    // not expose their temporary capacity. Reserve two logical buffers at a
    // conservative four-times-source envelope through the shared allocation
    // primitive. This remains compatible with both semantic owners without
    // widening AxisSupportBudget with owner-specific allocation methods.
    let buffer_bytes = bytes
        .checked_mul(4)
        .ok_or(AxisSupportError::InvalidSource)?;
    let capacity = buffer_bytes.div_ceil(size_of::<u64>()).max(1);
    budget.charge_reference_vector(capacity)?;
    budget.charge_reference_vector(capacity)
}

pub(super) fn package_metadata_payload<'a>(
    package: &'a Package,
    budget: &mut dyn AxisSupportBudget,
) -> Result<Option<&'a [u8]>, AxisSupportError> {
    // The owner-level scan-pass helpers construct an inventory by traversing
    // every object/message. Calling one immediately before the lookup below
    // would perform a duplicate pass and leave this second pass uncharged.
    // Precharge the fixed package envelope once, then charge each discovered
    // object/message before its bounded checks.
    charge_package_metadata_loop_bound(package, budget)?;
    let mut payload = None;
    for component in package.state.source.components().iter() {
        budget.charge_work(1)?;
        for object in &component.archive().objects {
            budget.charge_work(1)?;
            if object.messages.len() != object.archive_info.message_infos.len() {
                return Err(AxisSupportError::InvalidSource);
            }
            for (index, message) in object.messages.iter().enumerate() {
                budget.charge_work(1)?;
                let info = object
                    .archive_info
                    .message_infos
                    .get(index)
                    .ok_or(AxisSupportError::InvalidSource)?;
                if message.type_ != info.type_
                    || usize::try_from(info.length).ok() != Some(message.data.len())
                {
                    return Err(AxisSupportError::InvalidSource);
                }
                if message.type_ == PACKAGE_METADATA_MESSAGE_TYPE {
                    validate_selected_message_metadata(object, index)?;
                    if payload.replace(message.data.as_slice()).is_some() {
                        return Err(AxisSupportError::InvalidSource);
                    }
                }
            }
        }
    }
    Ok(payload)
}

pub(super) fn metadata_component_matches_physical(
    component: package_metadata_codec::ComponentDescriptor<'_>,
    physical: &str,
) -> bool {
    let Some(expected) = physical
        .strip_prefix("Index/")
        .and_then(|value| value.strip_suffix(".iwa"))
    else {
        return false;
    };
    // Native PackageMetadata keeps a human-readable preferred locator (for
    // example, `Slide`) alongside an explicit archive locator (for example,
    // `Slide-2652150`).  When the explicit locator is present it is the
    // physical identity; requiring the preferred locator to repeat it rejects
    // otherwise valid native packages.  A preferred locator alone still has
    // to be the exact physical basename, and an explicit locator must agree
    // with the effective locator, so generic or conflicting descriptors stay
    // rejected.
    let preferred_matches = component.preferred_locator() == expected;
    let explicit_matches = component
        .locator()
        .is_some_and(|locator| locator == expected);
    (preferred_matches || explicit_matches)
        && component
            .locator()
            .is_none_or(|locator| locator == expected)
        && component.effective_locator() == expected
}

struct SelectedAxisMetadataVisitor<'a> {
    axis_component_name: &'a str,
    owner_external_component_name: Option<&'a str>,
    registry_external_component_name: Option<&'a str>,
    axis_identifier: u64,
    selected_uuid: Option<package_metadata_codec::UuidBits>,
    selected_component_identifier: Option<u64>,
    external_target_component_identifier: Option<u64>,
    owner_external_seen: bool,
    registry_external_seen: bool,
    count: usize,
    invalid: bool,
}

impl package_metadata_codec::PackageMetadataVisitor for SelectedAxisMetadataVisitor<'_> {
    fn visit_unknown_field(&mut self) -> Result<(), package_metadata_codec::RewriteError> {
        // PackageMetadata is extensible in native files. Authority is granted
        // only by the explicitly projected UUID/component/reference records;
        // opaque extensions are preserved but never treated as evidence.
        Ok(())
    }

    fn visit_object_uuid(
        &mut self,
        binding: package_metadata_codec::ObjectUuidDescriptor<'_>,
    ) -> Result<(), package_metadata_codec::RewriteError> {
        if binding.object_identifier() == self.axis_identifier {
            if let Some(count) = self.count.checked_add(1) {
                self.count = count;
            } else {
                self.invalid = true;
            }
            if self.selected_uuid.replace(binding.uuid()).is_some() {
                self.invalid = true;
            }
            if self
                .selected_component_identifier
                .replace(binding.component().identifier())
                .is_some()
            {
                self.invalid = true;
            }
            if !binding.component().is_current()
                || !metadata_component_matches_physical(
                    binding.component(),
                    self.axis_component_name,
                )
            {
                self.invalid = true;
            }
        }
        Ok(())
    }

    fn visit_external_reference(
        &mut self,
        reference: package_metadata_codec::ExternalReferenceDescriptor<'_>,
    ) -> Result<(), package_metadata_codec::RewriteError> {
        if reference.object_identifier() == Some(self.axis_identifier) {
            let target = reference.target_component_identifier();
            if self
                .external_target_component_identifier
                .is_some_and(|identifier| identifier != target)
            {
                self.invalid = true;
            } else {
                self.external_target_component_identifier = Some(target);
            }
            let source_matches = |component_name: &str| {
                reference.source().is_current()
                    && metadata_component_matches_physical(reference.source(), component_name)
            };
            if self
                .owner_external_component_name
                .is_some_and(source_matches)
            {
                if self.owner_external_seen {
                    self.invalid = true;
                }
                self.owner_external_seen = true;
            } else if self
                .registry_external_component_name
                .is_some_and(source_matches)
            {
                if self.registry_external_seen {
                    self.invalid = true;
                }
                self.registry_external_seen = true;
            } else {
                self.invalid = true;
            }
            if reference.is_weak() == Some(true) || reference.is_versioned() {
                self.invalid = true;
            }
        }
        Ok(())
    }

    fn visit_data_reference_owner(
        &mut self,
        owner: package_metadata_codec::DataReferenceOwnerDescriptor<'_>,
    ) -> Result<(), package_metadata_codec::RewriteError> {
        if owner.object_identifier() == self.axis_identifier {
            self.invalid = true;
        }
        Ok(())
    }

    fn visit_ambiguous_object_identifier(
        &mut self,
        _component: package_metadata_codec::ComponentDescriptor<'_>,
        identifier: u64,
    ) -> Result<(), package_metadata_codec::RewriteError> {
        if identifier == self.axis_identifier {
            self.invalid = true;
        }
        Ok(())
    }

    fn visit_data_metadata_map(
        &mut self,
        object_identifier: u64,
        _has_unknown_fields: bool,
    ) -> Result<(), package_metadata_codec::RewriteError> {
        if object_identifier == self.axis_identifier {
            self.invalid = true;
        }
        Ok(())
    }
}

struct AxisMetadataAuthorityVisitor {
    axis_identifier: u64,
    selected_uuid: package_metadata_codec::UuidBits,
    selected_pair_count: usize,
    selected_object_count: usize,
    invalid: bool,
}

impl package_metadata_codec::PackageMetadataVisitor for AxisMetadataAuthorityVisitor {
    fn visit_unknown_field(&mut self) -> Result<(), package_metadata_codec::RewriteError> {
        // Unknown extensions neither grant nor contradict the known authority
        // records. Physical archive/reference checks remain independently
        // required before a mutation is admitted.
        Ok(())
    }

    fn visit_object_uuid(
        &mut self,
        binding: package_metadata_codec::ObjectUuidDescriptor<'_>,
    ) -> Result<(), package_metadata_codec::RewriteError> {
        if binding.uuid() == self.selected_uuid {
            if let Some(count) = self.selected_pair_count.checked_add(1) {
                self.selected_pair_count = count;
            } else {
                self.invalid = true;
            }
            if binding.object_identifier() != self.axis_identifier {
                self.invalid = true;
            }
        }
        if binding.object_identifier() == self.axis_identifier {
            if let Some(count) = self.selected_object_count.checked_add(1) {
                self.selected_object_count = count;
            } else {
                self.invalid = true;
            }
            if binding.uuid() != self.selected_uuid {
                self.invalid = true;
            }
        }
        Ok(())
    }
}

/// Verify that a chart-axis rewrite touched only its selected component and
/// permitted root rendering previews. This is independent of the semantic
/// value being edited, so title and value-axis owners share one locality proof.
pub(super) fn verify_package_locality(
    source: &Package,
    candidate: &Package,
    selection: &AxisSelection,
    previews_must_be_absent: bool,
    budget: &mut dyn AxisSupportBudget,
) -> Result<(), AxisSupportError> {
    verify_package_locality_with_expected_message(
        source,
        candidate,
        selection,
        previews_must_be_absent,
        None,
        budget,
    )
}

/// The minimal accounting surface needed by the physical locality proof.
///
/// Axis owners have a full aggregate ledger, while a few legacy semantic
/// owners only expose a bounded wire-work ledger. Keeping this smaller seam
/// private lets those owners reuse the same ZIP/object/message proof without
/// duplicating its source-preservation logic or manufacturing an unrelated
/// axis selection.
pub(super) trait LocalityBudget {
    fn charge_locality_scan(&mut self, package: &Package) -> Result<(), AxisSupportError>;
    fn charge_reference_vector(&mut self, capacity: usize) -> Result<(), AxisSupportError>;
    fn charge_wire_vector(&mut self, payload: usize) -> Result<(), AxisSupportError>;
    fn finish_wire_scan(&mut self, fields: usize) -> Result<(), AxisSupportError>;
    fn charge_work(&mut self, amount: usize) -> Result<(), AxisSupportError>;
}

impl<T: AxisSupportBudget + ?Sized> LocalityBudget for T {
    fn charge_locality_scan(&mut self, package: &Package) -> Result<(), AxisSupportError> {
        AxisSupportBudget::charge_locality_scan(self, package)
    }

    fn charge_reference_vector(&mut self, capacity: usize) -> Result<(), AxisSupportError> {
        AxisSupportBudget::charge_reference_vector(self, capacity)
    }

    fn charge_wire_vector(&mut self, payload: usize) -> Result<(), AxisSupportError> {
        AxisSupportBudget::charge_wire_vector(self, payload)
    }

    fn finish_wire_scan(&mut self, fields: usize) -> Result<(), AxisSupportError> {
        AxisSupportBudget::finish_wire_scan(self, fields)
    }

    fn charge_work(&mut self, amount: usize) -> Result<(), AxisSupportError> {
        AxisSupportBudget::charge_work(self, amount)
    }
}

/// Verify locality with an optional source-authoritative selected message.
///
/// Chart-axis title rewrites predate value-axis transactions and retain their
/// own semantic readback proof, so the compatibility wrapper above leaves
/// their selected message open.  Value-axis transactions additionally pass
/// the exact message bytes produced by their deterministic rewrite.  This
/// keeps semantic readback from becoming an authorization for changing an
/// unrelated outer field or an opaque span inside the selected extension.
pub(super) fn verify_package_locality_with_expected_message(
    source: &Package,
    candidate: &Package,
    selection: &AxisSelection,
    previews_must_be_absent: bool,
    expected_selected_message: Option<&[u8]>,
    budget: &mut dyn AxisSupportBudget,
) -> Result<(), AxisSupportError> {
    verify_package_locality_for_component(
        source,
        candidate,
        &selection.axis_component_name,
        selection.axis_identifier,
        selection.axis_message_index,
        previews_must_be_absent,
        expected_selected_message,
        budget,
    )
}

/// Verify physical locality for one selected object/message in one component.
///
/// The selected component may be a native slide component or a separately
/// stored imported chart component. Every other package member and every
/// unselected object/message must remain source-authoritative. When the caller
/// supplies `Some` expected bytes, the selected message must match those bytes
/// exactly; `None` retains the semantic-only compatibility behavior used by
/// older axis owners.
pub(super) fn verify_package_locality_for_component<B: LocalityBudget + ?Sized>(
    source: &Package,
    candidate: &Package,
    selected_component_name: &str,
    selected_identifier: u64,
    selected_message_index: usize,
    previews_must_be_absent: bool,
    expected_selected_message: Option<&[u8]>,
    budget: &mut B,
) -> Result<(), AxisSupportError> {
    budget.charge_locality_scan(source)?;
    budget.charge_locality_scan(candidate)?;
    let source_catalog = physical_catalog(source)?.package();
    let candidate_catalog = physical_catalog(candidate)?.package();
    verify_zip_envelope_locality(source_catalog, candidate_catalog, budget)?;
    // Preview discovery tests each member against the three native preview
    // names. Reserve that complete comparison bound before the helper can
    // inspect either catalog (and reserve the bounded result vectors too).
    budget.charge_work(
        source_catalog
            .len()
            .checked_add(candidate_catalog.len())
            .and_then(|value| value.checked_mul(3))
            .ok_or(AxisSupportError::InvalidSource)?,
    )?;
    // Preview discovery allocates bounded name vectors and compares every
    // archive member against the deletion set. Reserve that work before the
    // helper can inspect either catalog.
    budget.charge_reference_vector(3)?;
    budget.charge_reference_vector(3)?;
    let source_previews = super::rendering_invalidation::root_preview_deletions(source_catalog)
        .map_err(|_| AxisSupportError::InvalidSource)?;
    let candidate_previews =
        super::rendering_invalidation::root_preview_deletions(candidate_catalog)
            .map_err(|_| AxisSupportError::InvalidSource)?;
    // Compare the filtered catalogs in order. Reassembly preserves member
    // order; a name lookup would both admit reordering and turn malformed
    // input into an unbounded quadratic scan. The linear bound is charged
    // before either iterator starts consuming entries.
    budget.charge_work(
        source_catalog
            .len()
            .checked_add(candidate_catalog.len())
            .and_then(|value| value.checked_mul(3))
            .ok_or(AxisSupportError::InvalidSource)?,
    )?;
    let mut source_entries = source_catalog
        .iter()
        .filter(|entry| !source_previews.names().contains(&entry.name()));
    let mut candidate_entries = candidate_catalog
        .iter()
        .filter(|entry| !candidate_previews.names().contains(&entry.name()));
    loop {
        match (source_entries.next(), candidate_entries.next()) {
            (Some(entry), Some(other)) => {
                let selected_component = entry.name() == selected_component_name;
                budget.charge_work(entry_comparison_work(entry, other)?)?;
                if entry.name() != other.name()
                    || entry.raw_name() != other.raw_name()
                    || entry.is_opaque() != other.is_opaque()
                    || (!selected_component
                        && (entry.metadata().local() != other.metadata().local()
                            || entry.metadata().central() != other.metadata().central()
                            || entry.data() != other.data()
                            || entry.metadata() != other.metadata()
                            || entry.raw_record().local_record()
                                != other.raw_record().local_record()
                            || !central_record_preserved_except_offset(entry, other)))
                    || (selected_component && !selected_entry_records_compatible(entry, other))
                {
                    return Err(AxisSupportError::InvalidSource);
                }
            },
            (None, None) => break,
            _ => return Err(AxisSupportError::InvalidSource),
        }
    }
    if previews_must_be_absent && !candidate_previews.names().is_empty() {
        return Err(AxisSupportError::InvalidSource);
    }
    let component_lookup_work = source
        .state
        .source
        .components()
        .len()
        .checked_add(candidate.state.source.components().len())
        .ok_or(AxisSupportError::InvalidSource)?;
    budget.charge_work(component_lookup_work)?;
    let source_component = source
        .state
        .source
        .components()
        .iter()
        .find(|component| component.name() == selected_component_name)
        .ok_or(AxisSupportError::InvalidSource)?;
    let candidate_component = candidate
        .state
        .source
        .components()
        .iter()
        .find(|component| component.name() == selected_component_name)
        .ok_or(AxisSupportError::InvalidSource)?;
    let source_objects = &source_component.archive().objects;
    let candidate_objects = &candidate_component.archive().objects;
    budget.charge_work(
        source_objects
            .len()
            .checked_add(candidate_objects.len())
            .ok_or(AxisSupportError::InvalidSource)?,
    )?;
    if source_objects.len() != candidate_objects.len() {
        return Err(AxisSupportError::InvalidSource);
    }
    for (source_object, candidate_object) in source_objects.iter().zip(candidate_objects) {
        budget.charge_work(archive_info_comparison_work(
            source_object,
            candidate_object,
        )?)?;
        if source_object.archive_info.identifier != candidate_object.archive_info.identifier {
            return Err(AxisSupportError::InvalidSource);
        }
        let selected_object = source_object.archive_info.identifier == Some(selected_identifier);
        if !archive_info_compatible(
            source_object,
            candidate_object,
            selected_object.then_some(selected_message_index),
        ) || source_object.messages.len() != candidate_object.messages.len()
            || (!selected_object && !source_object.same_content_ignoring_offsets(candidate_object))
        {
            return Err(AxisSupportError::InvalidSource);
        }
        for (message_index, (source_message, candidate_message)) in source_object
            .messages
            .iter()
            .zip(&candidate_object.messages)
            .enumerate()
        {
            if selected_object && message_index == selected_message_index {
                budget.charge_work(
                    source_message
                        .data
                        .len()
                        .checked_add(candidate_message.data.len())
                        .and_then(|value| value.checked_add(1))
                        .ok_or(AxisSupportError::InvalidSource)?,
                )?;
                if source_message.type_ != candidate_message.type_ {
                    return Err(AxisSupportError::InvalidSource);
                }
                if let Some(expected) = expected_selected_message {
                    budget.charge_work(expected.len())?;
                    if candidate_message.data.as_slice() != expected {
                        return Err(AxisSupportError::InvalidSource);
                    }
                }
            } else {
                budget.charge_work(
                    source_message
                        .data
                        .len()
                        .checked_add(candidate_message.data.len())
                        .and_then(|value| value.checked_add(1))
                        .ok_or(AxisSupportError::InvalidSource)?,
                )?;
                if source_message != candidate_message {
                    return Err(AxisSupportError::InvalidSource);
                }
            }
        }
    }
    verify_selected_archive_info_framing(
        source,
        candidate,
        selected_component_name,
        selected_identifier,
        selected_message_index,
        budget,
    )?;
    Ok(())
}

/// Return the structural ZIP envelope boundaries needed by the locality
/// proof.  The archive reader intentionally keeps these details private, so
/// this small, allocation-free parser uses the already validated raw source
/// retained by the physical catalog.  The central records are walked to
/// derive the first local-record offset instead of trusting a byte-pattern
/// search through an arbitrary prelude.
#[derive(Debug, Clone, Copy)]
struct ZipEnvelopeShape {
    prelude_end: usize,
    eocd_offset: usize,
    comment_start: usize,
    comment_end: usize,
}

fn verify_zip_envelope_locality<B: LocalityBudget + ?Sized>(
    source: &litchi_iwa_archive::package::Catalog,
    candidate: &litchi_iwa_archive::package::Catalog,
    budget: &mut B,
) -> Result<(), AxisSupportError> {
    let source_bytes = source.source_bytes();
    let candidate_bytes = candidate.source_bytes();
    budget.charge_work(
        source_bytes
            .len()
            .checked_add(candidate_bytes.len())
            .ok_or(AxisSupportError::InvalidSource)?,
    )?;
    let source_shape =
        parse_zip_envelope_shape(source_bytes).ok_or(AxisSupportError::InvalidSource)?;
    let candidate_shape =
        parse_zip_envelope_shape(candidate_bytes).ok_or(AxisSupportError::InvalidSource)?;

    let source_prelude = source_bytes
        .get(..source_shape.prelude_end)
        .ok_or(AxisSupportError::InvalidSource)?;
    let candidate_prelude = candidate_bytes
        .get(..candidate_shape.prelude_end)
        .ok_or(AxisSupportError::InvalidSource)?;
    if source_prelude != candidate_prelude {
        return Err(AxisSupportError::InvalidSource);
    }

    // Reassembly may move the central directory and its end record, and may
    // change the entry count/central size when previews are deleted.  Those
    // four EOCD words are therefore deliberately excluded.  Disk fields,
    // comment length, comment bytes, and every byte after the comment remain
    // source-authoritative.
    let source_eocd = source_bytes
        .get(source_shape.eocd_offset..)
        .ok_or(AxisSupportError::InvalidSource)?;
    let candidate_eocd = candidate_bytes
        .get(candidate_shape.eocd_offset..)
        .ok_or(AxisSupportError::InvalidSource)?;
    if source_eocd.len() < 22 || candidate_eocd.len() < 22 {
        return Err(AxisSupportError::InvalidSource);
    }
    if source_eocd[..8] != candidate_eocd[..8] || source_eocd[20..22] != candidate_eocd[20..22] {
        return Err(AxisSupportError::InvalidSource);
    }
    let source_comment = source_bytes
        .get(source_shape.comment_start..source_shape.comment_end)
        .ok_or(AxisSupportError::InvalidSource)?;
    let candidate_comment = candidate_bytes
        .get(candidate_shape.comment_start..candidate_shape.comment_end)
        .ok_or(AxisSupportError::InvalidSource)?;
    if source_comment != candidate_comment
        || source_bytes.get(source_shape.comment_end..)
            != candidate_bytes.get(candidate_shape.comment_end..)
    {
        return Err(AxisSupportError::InvalidSource);
    }
    Ok(())
}

fn parse_zip_envelope_shape(bytes: &[u8]) -> Option<ZipEnvelopeShape> {
    if bytes.len() < 22 {
        return None;
    }
    let mut offset = bytes.len().checked_sub(22)?;
    loop {
        if read_u32(bytes, offset) == Some(0x0605_4b50) {
            if let Some(shape) = parse_zip_eocd_at(bytes, offset) {
                return Some(shape);
            }
        }
        if offset == 0 {
            break;
        }
        offset -= 1;
    }
    None
}

fn parse_zip_eocd_at(bytes: &[u8], eocd_offset: usize) -> Option<ZipEnvelopeShape> {
    let tail = bytes.get(eocd_offset..)?;
    if tail.len() < 22
        || read_u16(tail, 4) != Some(0)
        || read_u16(tail, 6) != Some(0)
        || read_u16(tail, 8) != read_u16(tail, 10)
    {
        return None;
    }
    let entry_count = usize::from(read_u16(tail, 10)?);
    let central_size = usize::try_from(read_u32(tail, 12)?).ok()?;
    let raw_central_offset = u64::from(read_u32(tail, 16)?);
    let comment_length = usize::from(read_u16(tail, 20)?);
    let comment_start = eocd_offset.checked_add(22)?;
    let comment_end = comment_start.checked_add(comment_length)?;
    if comment_end > bytes.len() {
        return None;
    }
    let central_start = eocd_offset.checked_sub(central_size)?;
    let central_base = u64::try_from(central_start)
        .ok()?
        .checked_sub(raw_central_offset)?;
    let mut cursor = central_start;
    let mut first_local: Option<usize> = None;
    for _ in 0..entry_count {
        let record = bytes.get(cursor..cursor.checked_add(46)?)?;
        if read_u32(record, 0) != Some(0x0201_4b50) {
            return None;
        }
        let name_length = usize::from(read_u16(record, 28)?);
        let extra_length = usize::from(read_u16(record, 30)?);
        let file_comment_length = usize::from(read_u16(record, 32)?);
        let record_length = 46usize
            .checked_add(name_length)?
            .checked_add(extra_length)?
            .checked_add(file_comment_length)?;
        let record_end = cursor.checked_add(record_length)?;
        if record_end > eocd_offset {
            return None;
        }
        let raw_local_offset = u64::from(read_u32(record, 42)?);
        let local_offset = central_base.checked_add(raw_local_offset)?;
        let local_offset = usize::try_from(local_offset).ok()?;
        if local_offset >= central_start || read_u32(bytes, local_offset) != Some(0x0403_4b50) {
            return None;
        }
        first_local = Some(first_local.map_or(local_offset, |old| old.min(local_offset)));
        cursor = record_end;
    }
    if cursor != eocd_offset {
        return None;
    }
    Some(ZipEnvelopeShape {
        prelude_end: first_local.unwrap_or(central_start),
        eocd_offset,
        comment_start,
        comment_end,
    })
}

/// Compare the selected object's raw `ArchiveInfo` framing.  The parsed
/// neutral projection intentionally omits unknown field bytes, so decoded
/// equality alone cannot prove that a candidate preserved those bytes.  The
/// only permitted difference here is the selected `MessageInfo.length`
/// scalar; its enclosing length prefixes are checked against the same
/// source-width-preserving varint rule used by the IWA rewriter.
fn verify_selected_archive_info_framing<B: LocalityBudget + ?Sized>(
    source: &Package,
    candidate: &Package,
    selected_component_name: &str,
    selected_identifier: u64,
    selected_message_index: usize,
    budget: &mut B,
) -> Result<(), AxisSupportError> {
    let source_catalog = physical_catalog(source)?.package();
    let candidate_catalog = physical_catalog(candidate)?.package();
    let source_entry = source_catalog
        .iter()
        .find(|entry| entry.name() == selected_component_name)
        .ok_or(AxisSupportError::InvalidSource)?;
    let candidate_entry = candidate_catalog
        .iter()
        .find(|entry| entry.name() == selected_component_name)
        .ok_or(AxisSupportError::InvalidSource)?;
    if source_entry.is_opaque() || candidate_entry.is_opaque() {
        return Err(AxisSupportError::InvalidSource);
    }
    let source_component = source
        .state
        .source
        .components()
        .iter()
        .find(|component| component.name() == selected_component_name)
        .ok_or(AxisSupportError::InvalidSource)?;
    let candidate_component = candidate
        .state
        .source
        .components()
        .iter()
        .find(|component| component.name() == selected_component_name)
        .ok_or(AxisSupportError::InvalidSource)?;
    let source_object = source_component
        .archive()
        .objects
        .iter()
        .find(|object| object.archive_info.identifier == Some(selected_identifier))
        .ok_or(AxisSupportError::InvalidSource)?;
    let candidate_object = candidate_component
        .archive()
        .objects
        .iter()
        .find(|object| object.archive_info.identifier == Some(selected_identifier))
        .ok_or(AxisSupportError::InvalidSource)?;
    if source_object.header_offset != candidate_object.header_offset {
        return Err(AxisSupportError::InvalidSource);
    }
    let archive_limits = source
        .state
        .options
        .archive()
        .effective_archive_limits()
        .map_err(map_archive_error)?;
    let source_bound = source_component
        .archive()
        .encoded_len_with_limits(archive_limits)
        .map_err(map_core_error)?;
    let candidate_bound = candidate_component
        .archive()
        .encoded_len_with_limits(archive_limits)
        .map_err(map_core_error)?;
    let source_compressed = source_entry.data();
    let candidate_compressed = candidate_entry.data();
    budget.charge_work(
        source_compressed
            .len()
            .checked_add(candidate_compressed.len())
            .and_then(|value| value.checked_add(source_bound))
            .and_then(|value| value.checked_add(candidate_bound))
            .ok_or(AxisSupportError::InvalidSource)?,
    )?;
    let source_snappy_limits = source
        .state
        .options
        .archive()
        .snappy_limits()
        .map_err(map_archive_error)?;
    let candidate_snappy_limits = candidate
        .state
        .options
        .archive()
        .snappy_limits()
        .map_err(map_archive_error)?;
    let source_stream =
        SnappyStream::decompress_with_limits(source_compressed, source_snappy_limits)
            .map_err(map_core_error)?;
    let candidate_stream =
        SnappyStream::decompress_with_limits(candidate_compressed, candidate_snappy_limits)
            .map_err(map_core_error)?;
    budget.charge_work(
        source_stream
            .as_bytes()
            .len()
            .checked_add(candidate_stream.as_bytes().len())
            .ok_or(AxisSupportError::InvalidSource)?,
    )?;
    let source_header = framed_object_header(source_stream.as_bytes(), source_object)
        .ok_or(AxisSupportError::InvalidSource)?;
    let candidate_header = framed_object_header(candidate_stream.as_bytes(), candidate_object)
        .ok_or(AxisSupportError::InvalidSource)?;
    let source_length = source_object
        .archive_info
        .message_infos
        .get(selected_message_index)
        .ok_or(AxisSupportError::InvalidSource)?
        .length;
    let candidate_length = candidate_object
        .archive_info
        .message_infos
        .get(selected_message_index)
        .ok_or(AxisSupportError::InvalidSource)?
        .length;
    compare_archive_info_headers(
        source_header,
        candidate_header,
        selected_message_index,
        source_length,
        candidate_length,
        source.wire_limits().map_err(map_wire_error)?,
        budget,
    )
}

fn framed_object_header<'a>(stream: &'a [u8], object: &ArchiveObject) -> Option<&'a [u8]> {
    let start = usize::try_from(object.header_offset).ok()?;
    let framed_length = usize::try_from(object.header_length).ok()?;
    let framed = stream.get(start..start.checked_add(framed_length)?)?;
    let (header_length, prefix_length) = decode_varint_from_bytes(framed).ok()?;
    let header_length = usize::try_from(header_length).ok()?;
    if prefix_length != encoded_len(header_length as u64)
        || prefix_length.checked_add(header_length) != Some(framed_length)
    {
        return None;
    }
    framed.get(prefix_length..)
}

fn compare_archive_info_headers<B: LocalityBudget + ?Sized>(
    source: &[u8],
    candidate: &[u8],
    selected_message_index: usize,
    source_length: u32,
    candidate_length: u32,
    limits: WireLimits,
    budget: &mut B,
) -> Result<(), AxisSupportError> {
    budget.charge_wire_vector(source.len())?;
    budget.charge_wire_vector(candidate.len())?;
    let source_fields = parse_wire_fields_with_limits(source, limits).map_err(map_wire_error)?;
    let candidate_fields =
        parse_wire_fields_with_limits(candidate, limits).map_err(map_wire_error)?;
    budget.finish_wire_scan(source_fields.len())?;
    budget.finish_wire_scan(candidate_fields.len())?;
    if source_fields.len() != candidate_fields.len() {
        return Err(AxisSupportError::InvalidSource);
    }
    let mut message_index = 0usize;
    for (source_field, candidate_field) in source_fields.iter().zip(&candidate_fields) {
        if source_field.number() != candidate_field.number()
            || source_field.wire_type() != candidate_field.wire_type()
        {
            return Err(AxisSupportError::InvalidSource);
        }
        if source_field.number() == 2 && source_field.wire_type() == 2 {
            let selected = message_index == selected_message_index;
            if selected {
                compare_selected_message_info(
                    *source_field,
                    *candidate_field,
                    source,
                    candidate,
                    source_length,
                    candidate_length,
                    limits,
                    budget,
                )?;
            } else if source_field.raw(source).map_err(map_wire_error)?
                != candidate_field.raw(candidate).map_err(map_wire_error)?
            {
                return Err(AxisSupportError::InvalidSource);
            }
            message_index = message_index
                .checked_add(1)
                .ok_or(AxisSupportError::InvalidSource)?;
        } else if source_field.raw(source).map_err(map_wire_error)?
            != candidate_field.raw(candidate).map_err(map_wire_error)?
        {
            return Err(AxisSupportError::InvalidSource);
        }
    }
    if message_index <= selected_message_index {
        return Err(AxisSupportError::InvalidSource);
    }
    Ok(())
}

fn compare_selected_message_info<B: LocalityBudget + ?Sized>(
    source_field: WireField,
    candidate_field: WireField,
    source: &[u8],
    candidate: &[u8],
    source_length: u32,
    candidate_length: u32,
    limits: WireLimits,
    budget: &mut B,
) -> Result<(), AxisSupportError> {
    if source_field.key(source).map_err(map_wire_error)?
        != candidate_field.key(candidate).map_err(map_wire_error)?
    {
        return Err(AxisSupportError::InvalidSource);
    }
    let source_payload = source_field.payload(source).map_err(map_wire_error)?;
    let candidate_payload = candidate_field.payload(candidate).map_err(map_wire_error)?;
    compare_length_prefix(
        source,
        candidate,
        source_field,
        candidate_field,
        source_payload.len(),
        candidate_payload.len(),
    )?;
    budget.charge_wire_vector(source_payload.len())?;
    budget.charge_wire_vector(candidate_payload.len())?;
    let source_fields =
        parse_wire_fields_with_limits(source_payload, limits).map_err(map_wire_error)?;
    let candidate_fields =
        parse_wire_fields_with_limits(candidate_payload, limits).map_err(map_wire_error)?;
    budget.finish_wire_scan(source_fields.len())?;
    budget.finish_wire_scan(candidate_fields.len())?;
    if source_fields.len() != candidate_fields.len() {
        return Err(AxisSupportError::InvalidSource);
    }
    let selected_length_field = source_fields
        .iter()
        .enumerate()
        .filter(|(_, field)| field.number() == 3 && field.wire_type() == 0)
        .map(|(index, _)| index)
        .next_back()
        .ok_or(AxisSupportError::InvalidSource)?;
    if candidate_fields
        .iter()
        .enumerate()
        .filter(|(_, field)| field.number() == 3 && field.wire_type() == 0)
        .map(|(index, _)| index)
        .next_back()
        != Some(selected_length_field)
    {
        return Err(AxisSupportError::InvalidSource);
    }
    for (index, (source_field, candidate_field)) in
        source_fields.iter().zip(&candidate_fields).enumerate()
    {
        if source_field.number() != candidate_field.number()
            || source_field.wire_type() != candidate_field.wire_type()
        {
            return Err(AxisSupportError::InvalidSource);
        }
        if index != selected_length_field {
            if source_field.raw(source_payload).map_err(map_wire_error)?
                != candidate_field
                    .raw(candidate_payload)
                    .map_err(map_wire_error)?
            {
                return Err(AxisSupportError::InvalidSource);
            }
            continue;
        }
        if source_field.key(source_payload).map_err(map_wire_error)?
            != candidate_field
                .key(candidate_payload)
                .map_err(map_wire_error)?
        {
            return Err(AxisSupportError::InvalidSource);
        }
        let source_value = exact_varint(
            source_field
                .payload(source_payload)
                .map_err(map_wire_error)?,
        )
        .ok_or(AxisSupportError::InvalidSource)?;
        let candidate_value = exact_varint(
            candidate_field
                .payload(candidate_payload)
                .map_err(map_wire_error)?,
        )
        .ok_or(AxisSupportError::InvalidSource)?;
        if source_value != u64::from(source_length)
            || candidate_value != u64::from(candidate_length)
            || !varint_matches_source_width(
                u64::from(candidate_length),
                candidate_field
                    .payload(candidate_payload)
                    .map_err(map_wire_error)?,
                source_field
                    .payload(source_payload)
                    .map_err(map_wire_error)?
                    .len(),
            )
        {
            return Err(AxisSupportError::InvalidSource);
        }
    }
    Ok(())
}

fn compare_length_prefix(
    source: &[u8],
    candidate: &[u8],
    source_field: WireField,
    candidate_field: WireField,
    source_payload_length: usize,
    candidate_payload_length: usize,
) -> Result<(), AxisSupportError> {
    let source_prefix = source
        .get(source_field.key_end()..source_field.payload_start())
        .ok_or(AxisSupportError::InvalidSource)?;
    let candidate_prefix = candidate
        .get(candidate_field.key_end()..candidate_field.payload_start())
        .ok_or(AxisSupportError::InvalidSource)?;
    if exact_varint(source_prefix)
        != Some(u64::try_from(source_payload_length).map_err(|_| AxisSupportError::InvalidSource)?)
        || exact_varint(candidate_prefix)
            != Some(
                u64::try_from(candidate_payload_length)
                    .map_err(|_| AxisSupportError::InvalidSource)?,
            )
    {
        return Err(AxisSupportError::InvalidSource);
    }
    if source_payload_length == candidate_payload_length {
        if source_prefix != candidate_prefix {
            return Err(AxisSupportError::InvalidSource);
        }
    } else if !varint_matches_source_width(
        u64::try_from(candidate_payload_length).map_err(|_| AxisSupportError::InvalidSource)?,
        candidate_prefix,
        source_prefix.len(),
    ) {
        return Err(AxisSupportError::InvalidSource);
    }
    Ok(())
}

fn exact_varint(bytes: &[u8]) -> Option<u64> {
    let (value, width) = decode_varint_from_bytes(bytes).ok()?;
    (width == bytes.len()).then_some(value)
}

fn varint_matches_source_width(value: u64, bytes: &[u8], source_width: usize) -> bool {
    let width = encoded_len(value).max(source_width.min(10));
    if bytes.len() != width {
        return false;
    }
    let mut remaining = value;
    for (index, byte) in bytes.iter().copied().enumerate() {
        let mut expected = (remaining & 0x7f) as u8;
        remaining >>= 7;
        if index + 1 != width {
            expected |= 0x80;
        }
        if byte != expected {
            return false;
        }
    }
    true
}

fn archive_info_compatible(
    source: &ArchiveObject,
    candidate: &ArchiveObject,
    selected_message_index: Option<usize>,
) -> bool {
    source.archive_info.identifier == candidate.archive_info.identifier
        && source.archive_info.should_merge == candidate.archive_info.should_merge
        && source.archive_info.message_infos.len() == candidate.archive_info.message_infos.len()
        && source
            .archive_info
            .message_infos
            .iter()
            .zip(&candidate.archive_info.message_infos)
            .enumerate()
            .all(|(index, (source_info, candidate_info))| {
                if selected_message_index == Some(index) {
                    message_info_compatible_except_length(source_info, candidate_info)
                } else {
                    source_info == candidate_info
                }
            })
}

/// Return a conservative byte/comparison bound for one object metadata
/// equality check. `MessageInfo`/`FieldInfo` derive equality over nested
/// vectors and paths, so charging only the number of top-level messages would
/// under-account a late mismatch in a large metadata record.
fn archive_info_comparison_work(
    source: &ArchiveObject,
    candidate: &ArchiveObject,
) -> Result<usize, AxisSupportError> {
    let mut work = 4usize; // identifier, merge flag, and two length checks
    for info in source
        .archive_info
        .message_infos
        .iter()
        .chain(&candidate.archive_info.message_infos)
    {
        work = work
            .checked_add(message_info_comparison_work(info)?)
            .ok_or(AxisSupportError::InvalidSource)?;
    }
    Ok(work)
}

fn message_info_comparison_work(
    info: &litchi_iwa_core::MessageInfo,
) -> Result<usize, AxisSupportError> {
    let mut work = 16usize; // scalar and option-presence comparisons
    work = work
        .checked_add(info.versions.len())
        .and_then(|value| value.checked_add(info.object_references.len()))
        .and_then(|value| value.checked_add(info.data_references.len()))
        .and_then(|value| value.checked_add(info.diff_merge_version.len()))
        .and_then(|value| value.checked_add(info.diff_read_version.len()))
        .ok_or(AxisSupportError::InvalidSource)?;
    if let Some(path) = &info.diff_field_path {
        work = work
            .checked_add(path.path.len())
            .ok_or(AxisSupportError::InvalidSource)?;
    }
    for path in &info.fields_to_remove {
        work = work
            .checked_add(path.path.len())
            .ok_or(AxisSupportError::InvalidSource)?;
    }
    for field in &info.field_infos {
        work = work
            .checked_add(8)
            .and_then(|value| value.checked_add(field.path.path.len()))
            .and_then(|value| value.checked_add(field.object_references.len()))
            .and_then(|value| value.checked_add(field.data_references.len()))
            .and_then(|value| value.checked_add(field.known_field_version.len()))
            .and_then(|value| {
                field
                    .known_field_feature_identifier
                    .as_ref()
                    .map_or(Some(value), |feature| value.checked_add(feature.len()))
            })
            .ok_or(AxisSupportError::InvalidSource)?;
    }
    Ok(work)
}

fn message_info_compatible_except_length(
    source: &litchi_iwa_core::MessageInfo,
    candidate: &litchi_iwa_core::MessageInfo,
) -> bool {
    source.type_ == candidate.type_
        && source.versions == candidate.versions
        && source.field_infos == candidate.field_infos
        && source.object_references == candidate.object_references
        && source.data_references == candidate.data_references
        && source.base_message_index == candidate.base_message_index
        && source.diff_merge_version == candidate.diff_merge_version
        && source.diff_field_path == candidate.diff_field_path
        && source.fields_to_remove == candidate.fields_to_remove
        && source.diff_read_version == candidate.diff_read_version
}

fn entry_comparison_work(source: &Entry, candidate: &Entry) -> Result<usize, AxisSupportError> {
    source
        .data()
        .len()
        .checked_add(candidate.data().len())
        .and_then(|value| source.raw_record().local_record().len().checked_add(value))
        .and_then(|value| {
            source
                .raw_record()
                .central_directory_record()
                .len()
                .checked_add(value)
        })
        .and_then(|value| {
            candidate
                .raw_record()
                .local_record()
                .len()
                .checked_add(value)
        })
        .and_then(|value| {
            candidate
                .raw_record()
                .central_directory_record()
                .len()
                .checked_add(value)
        })
        .and_then(|value| value.checked_add(128))
        .ok_or(AxisSupportError::InvalidSource)
}

fn selected_entry_records_compatible(source: &Entry, candidate: &Entry) -> bool {
    if source.name() != candidate.name()
        || source.raw_name() != candidate.raw_name()
        || source.is_opaque() != candidate.is_opaque()
    {
        return false;
    }
    compatible_local_record(source, candidate) && compatible_central_record(source, candidate)
}

fn compatible_local_record(source: &Entry, candidate: &Entry) -> bool {
    let source_record = source.raw_record().local_record();
    let candidate_record = candidate.raw_record().local_record();
    if source_record.len() < 30 || candidate_record.len() < 30 {
        return false;
    }
    let Some(source_name_len) = read_u16(source_record, 26) else {
        return false;
    };
    let Some(source_extra_len) = read_u16(source_record, 28) else {
        return false;
    };
    let Some(candidate_name_len) = read_u16(candidate_record, 26) else {
        return false;
    };
    let Some(candidate_extra_len) = read_u16(candidate_record, 28) else {
        return false;
    };
    let source_header_len = 30usize
        .checked_add(usize::from(source_name_len))
        .and_then(|value| value.checked_add(usize::from(source_extra_len)));
    let candidate_header_len = 30usize
        .checked_add(usize::from(candidate_name_len))
        .and_then(|value| value.checked_add(usize::from(candidate_extra_len)));
    let (Some(source_header_len), Some(candidate_header_len)) =
        (source_header_len, candidate_header_len)
    else {
        return false;
    };
    if source_header_len != candidate_header_len
        || source_header_len > source_record.len()
        || candidate_header_len > candidate_record.len()
    {
        return false;
    }
    let source_compressed_len = source.raw_record().compressed_data().len();
    let candidate_compressed_len = candidate.raw_record().compressed_data().len();
    let Some(source_suffix_start) = source_header_len.checked_add(source_compressed_len) else {
        return false;
    };
    let Some(candidate_suffix_start) = candidate_header_len.checked_add(candidate_compressed_len)
    else {
        return false;
    };
    if source_suffix_start > source_record.len() || candidate_suffix_start > candidate_record.len()
    {
        return false;
    }
    // CRC and the two size words are the only local-header values changed by
    // an edited member. Every name/extra/header byte, including method,
    // flags, and timestamps, remains source-authoritative.
    if !equal_except(
        &source_record[..source_header_len],
        &candidate_record[..candidate_header_len],
        std::slice::from_ref(&(14..26)),
    ) {
        return false;
    }
    let source_suffix = &source_record[source_suffix_start..];
    let candidate_suffix = &candidate_record[candidate_suffix_start..];
    descriptor_shape_compatible(
        source.metadata().local().flags(),
        source_suffix,
        candidate_suffix,
    )
}

fn compatible_central_record(source: &Entry, candidate: &Entry) -> bool {
    let source_record = source.raw_record().central_directory_record();
    let candidate_record = candidate.raw_record().central_directory_record();
    if source_record.len() < 46
        || candidate_record.len() != source_record.len()
        || read_u16(source_record, 28) != read_u16(candidate_record, 28)
        || read_u16(source_record, 30) != read_u16(candidate_record, 30)
        || read_u16(source_record, 32) != read_u16(candidate_record, 32)
    {
        return false;
    }
    // CRC, compressed/uncompressed sizes, and the local-header offset are
    // rewritten by ZIP reassembly. Attributes, ordering fields, timestamps,
    // names, extras, comments, and all signatures stay exact.
    equal_except(source_record, candidate_record, &[(16..28), (42..46)])
}

fn central_record_preserved_except_offset(source: &Entry, candidate: &Entry) -> bool {
    let source_record = source.raw_record().central_directory_record();
    let candidate_record = candidate.raw_record().central_directory_record();
    equal_except(
        source_record,
        candidate_record,
        std::slice::from_ref(&(42..46)),
    )
}

fn descriptor_shape_compatible(flags: u16, source: &[u8], candidate: &[u8]) -> bool {
    if flags & 0x0008 == 0 {
        return source.is_empty() && candidate.is_empty();
    }
    if source.len() != candidate.len() || !matches!(source.len(), 12 | 16) {
        return false;
    }
    if source.len() == 16 {
        source[..4] == candidate[..4]
            && read_u32(source, 0) == Some(0x0807_4b50)
            && read_u32(candidate, 0) == Some(0x0807_4b50)
    } else {
        true
    }
}

fn equal_except(source: &[u8], candidate: &[u8], ignored: &[std::ops::Range<usize>]) -> bool {
    source.len() == candidate.len()
        && source.iter().enumerate().all(|(index, byte)| {
            ignored.iter().any(|range| range.contains(&index)) || Some(byte) == candidate.get(index)
        })
}

fn read_u16(bytes: &[u8], offset: usize) -> Option<u16> {
    let value = bytes.get(offset..offset.checked_add(2)?)?;
    Some(u16::from_le_bytes([value[0], value[1]]))
}

fn read_u32(bytes: &[u8], offset: usize) -> Option<u32> {
    let value = bytes.get(offset..offset.checked_add(4)?)?;
    Some(u32::from_le_bytes([value[0], value[1], value[2], value[3]]))
}

#[cfg(test)]
mod tests {
    use litchi_iwa_common::wire::{append_length_delimited_field, append_varint_field};
    use litchi_iwa_protos::package_metadata_codec::{
        self, ComponentDescriptor, PackageMetadataVisitor, RewriteError, RewriteOptions,
    };

    use super::metadata_component_matches_physical;

    fn put_varint(output: &mut Vec<u8>, field: u32, value: u64) {
        append_varint_field(output, field, value).expect("test varint fits wire limits");
    }

    fn put_bytes(output: &mut Vec<u8>, field: u32, value: &[u8]) {
        append_length_delimited_field(output, field, value)
            .expect("test bytes field fits wire limits");
    }

    fn component_metadata(preferred: &str, locator: Option<&str>) -> Vec<u8> {
        let mut component = Vec::new();
        put_varint(&mut component, 1, 1);
        put_bytes(&mut component, 2, preferred.as_bytes());
        if let Some(locator) = locator {
            put_bytes(&mut component, 3, locator.as_bytes());
        }

        let mut metadata = Vec::new();
        put_varint(&mut metadata, 1, 1);
        put_bytes(&mut metadata, 3, &component);
        metadata
    }

    struct MatchVisitor<'source> {
        physical: &'source str,
        matched: Option<bool>,
    }

    impl PackageMetadataVisitor for MatchVisitor<'_> {
        fn visit_component(
            &mut self,
            component: ComponentDescriptor<'_>,
        ) -> Result<(), RewriteError> {
            self.matched = Some(metadata_component_matches_physical(
                component,
                self.physical,
            ));
            Ok(())
        }
    }

    fn matches(preferred: &str, locator: Option<&str>) -> bool {
        let source = component_metadata(preferred, locator);
        let mut visitor = MatchVisitor {
            physical: "Index/Slide-2652150.iwa",
            matched: None,
        };
        package_metadata_codec::inspect_package_metadata_with_visitor(
            &source,
            RewriteOptions::new(source.len(), source.len(), 32, 4096, 8, 4, 4, 0),
            &mut visitor,
        )
        .expect("minimal PackageMetadata component should inspect");
        visitor.matched.expect("component callback should run")
    }

    #[test]
    fn metadata_component_matching_accepts_native_explicit_locator() {
        assert!(matches("Slide", Some("Slide-2652150")));
        assert!(matches("Slide-2652150", None));
    }

    #[test]
    fn metadata_component_matching_rejects_generic_or_conflicting_locator() {
        assert!(!matches("Slide", None));
        assert!(!matches("Slide", Some("Slide-2652151")));
        assert!(!matches("Slide-2652150", Some("Slide-2652151")));
        assert!(!matches("Other", Some("Slide-2652151")));
    }
}
