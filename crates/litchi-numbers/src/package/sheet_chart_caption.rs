//! Selector-first Numbers sheet-chart caption transactions.
//!
//! A chart caption is a native chart edge plus a bounded four-object text
//! graph.  Existing captions rewrite only their private storage.  Creating a
//! caption authors the shared inline graph, registers its fresh objects, and
//! repoints only the selected chart edge.  Clearing leaves the old graph in
//! place and attaches a fresh stand-in, matching the native undo-preserving
//! route.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "The focused chart-caption transaction keeps graph phases together."
)]

use std::fmt;
use std::num::NonZeroU64;
use std::sync::Arc;

use litchi_core::Position;
use litchi_iwa_archive::package::OwnedExactArtifacts;
use litchi_iwa_common::WireLimits;
use litchi_iwa_common::wire::WireView;
use litchi_iwa_core::{ArchiveObject, RawMessage};
use litchi_iwa_protos::pages_movie_caption_codec;
use litchi_iwa_style::chart_text::{caption as caption_codec, caption_graph as graph_codec};
use thiserror::Error as ThisError;

use super::Package;
use super::chart_arrangement::{
    self as arrangement, ChartBudget, ChartSelection, object_from_resolved,
    parse_wire_view_with_budget, select_chart_with_budget, selected_chart_payload,
    unique_typed_message, validate_message_metadata,
};
use super::sheet_drawable_graph as drawable_graph;
use super::structure::{
    self, Fault, FaultLimit, Working,
    metadata::{self, Registry, Transition},
    references,
};
use crate::selector::{ChartSelector, SheetSelector};

const CAPTION_INFO_MESSAGE_TYPE: u32 = 633;
const CAPTION_PLACEMENT_MESSAGE_TYPE: u32 = 634;
const SHAPE_STYLE_MESSAGE_TYPE: u32 = 2_025;
const STORAGE_MESSAGE_TYPE: u32 = 2_001;
const STANDIN_MESSAGE_TYPE: u32 = 3_097;
const THEME_MESSAGE_TYPE: u32 = 12_009;
const THEME_APPLICATION_FIELD: u32 = 210;
const THEME_CAPTION_STYLE_FIELD: u32 = 1;
const DOCUMENT_STYLESHEET_FIELD: u32 = 4;
const DOCUMENT_THEME_FIELD: u32 = 6;
const DOCUMENT_SUPER_FIELD: u32 = 8;
const DOCUMENT_LANGUAGE_FIELD: u32 = 3;
const CHART_DRAWABLE_FIELD: u32 = 1;
const DRAWABLE_CAPTION_FIELD: u32 = 11;
const DRAWABLE_GEOMETRY_FIELD: u32 = 1;
const GEOMETRY_SIZE_FIELD: u32 = 2;
const SIZE_WIDTH_FIELD: u32 = 1;
const STANDARD_VERSION: [u32; 3] = [1, 0, 5];
const STANDIN_VERSION: [u32; 3] = [10, 1, 0];
const CHART_MESSAGE_TYPE: u32 = 5_021;

/// A finite resource governed by one Numbers chart-caption transaction.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Complete source or candidate package bytes.
    PackageBytes,
    /// One decoded IWA component or graph payload.
    PayloadBytes,
    /// Authored graph/output bytes.
    OutputBytes,
    /// Fresh object identifiers.
    Identifiers,
    /// Bounded temporary allocations.
    Allocations,
}

impl fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::PackageBytes => "package bytes",
            Self::PayloadBytes => "payload bytes",
            Self::OutputBytes => "output bytes",
            Self::Identifiers => "object identifiers",
            Self::Allocations => "allocations",
        })
    }
}

/// A well-formed request this focused route refuses.
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Refusal {
    /// No sheet matched the semantic selector.
    #[error("sheet selector did not resolve")]
    SheetNotFound,
    /// No chart matched the semantic selector.
    #[error("chart selector did not resolve")]
    ChartNotFound,
    /// The selected chart has no caption reference edge.
    #[error("the selected chart has no caption edge")]
    NoCaptionEdge,
    /// The selected caption graph is shared or malformed.
    #[error("the selected chart caption graph is shared or malformed")]
    SharedObject,
    /// A source layout outside the admitted inline authoring profile.
    #[error("the selected chart uses an unsupported caption graph layout")]
    UnsupportedLayout,
    /// A new caption has no finite positive displayed width.
    #[error("the selected chart has no displayed width for a new caption")]
    NoWidth,
    /// The replacement contains a native object marker.
    #[error("the caption text contains a reserved native marker")]
    ReservedMarker,
    /// Replacing text would remove dependent native object content.
    #[error("the existing caption contains dependent native content")]
    DependentContent,
    /// The source does not retain an exact physical artifact.
    #[error("this Numbers source does not support exact chart-caption edits")]
    UnsupportedSource,
}

impl Refusal {
    /// Classify the refusal for unified operations.
    #[must_use]
    pub const fn kind(self) -> litchi_core::edit::RefusalKind {
        use litchi_core::edit::RefusalKind;
        match self {
            Self::SheetNotFound | Self::ChartNotFound => RefusalKind::NotFound,
            Self::ReservedMarker => RefusalKind::InvalidValue,
            Self::NoCaptionEdge
            | Self::SharedObject
            | Self::UnsupportedLayout
            | Self::NoWidth
            | Self::DependentContent
            | Self::UnsupportedSource => RefusalKind::UnsupportedContent,
        }
    }
}

/// Failure from a Numbers chart-caption read or transaction.
#[derive(Debug, Clone, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// A well-formed request this route refuses.
    #[error(transparent)]
    Refused(Refusal),
    /// The rooted chart graph or selected caption graph is malformed.
    #[error("the selected Numbers chart-caption source is invalid")]
    InvalidSource,
    /// A finite operation ceiling was exceeded.
    #[error("Numbers chart caption {kind} limit exceeded: observed {observed}, maximum {maximum}")]
    Limit {
        /// Resource category.
        kind: LimitKind,
        /// Observed amount.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
    },
    /// A bounded temporary allocation failed.
    #[error("could not allocate {amount} Numbers chart-caption units")]
    Allocation { amount: usize },
    /// Candidate reopen or semantic readback failed.
    #[error("the edited Numbers chart caption failed semantic verification")]
    Verification,
    /// A patch was applied to a different exact source.
    #[error("the Numbers chart-caption patch does not match its exact source")]
    PatchConflict,
}

impl From<Refusal> for Error {
    fn from(value: Refusal) -> Self {
        Self::Refused(value)
    }
}

fn map_fault(fault: Fault) -> Error {
    match fault {
        Fault::Unsupported => Refusal::UnsupportedSource.into(),
        Fault::Invalid | Fault::Dependency => Error::InvalidSource,
        Fault::Limit {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: match kind {
                FaultLimit::PackageBytes => LimitKind::PackageBytes,
                FaultLimit::PayloadBytes => LimitKind::PayloadBytes,
                FaultLimit::Identifiers => LimitKind::Identifiers,
            },
            observed,
            maximum,
        },
        Fault::Allocation { amount } => Error::Allocation { amount },
        Fault::Verification => Error::Verification,
    }
}

fn map_arrangement_error(error: arrangement::Error) -> Error {
    match error {
        arrangement::Error::Refused(arrangement::Refusal::UnsupportedSource) => {
            Refusal::UnsupportedSource.into()
        },
        arrangement::Error::Refused(arrangement::Refusal::SheetNameNotFound)
        | arrangement::Error::Refused(arrangement::Refusal::SheetPositionNotFound { .. }) => {
            Refusal::SheetNotFound.into()
        },
        arrangement::Error::Refused(arrangement::Refusal::ChartPositionNotFound { .. }) => {
            Refusal::ChartNotFound.into()
        },
        arrangement::Error::Limit {
            observed, maximum, ..
        } => Error::Limit {
            kind: LimitKind::PayloadBytes,
            observed,
            maximum,
        },
        arrangement::Error::Allocation { amount } => Error::Allocation { amount },
        _ => Error::InvalidSource,
    }
}

#[derive(Debug, Clone, PartialEq)]
struct Target {
    selection: ChartSelection,
    component_index: usize,
    component_name: Arc<str>,
    chart_identifier: NonZeroU64,
    chart_message_index: usize,
    chart_message_type: u32,
    caption_identifier: NonZeroU64,
    storage: Option<NonZeroU64>,
    text: Option<Box<str>>,
    width: Option<f32>,
}

impl Target {
    fn same_route(&self, other: &Self) -> bool {
        self.selection == other.selection
            && self.component_index == other.component_index
            && self.component_name == other.component_name
            && self.chart_identifier == other.chart_identifier
            && self.chart_message_index == other.chart_message_index
            && self.chart_message_type == other.chart_message_type
            && self.caption_identifier == other.caption_identifier
    }
}

fn unique_nonzero(value: u64) -> Result<NonZeroU64, Error> {
    NonZeroU64::new(value).ok_or(Error::InvalidSource)
}

fn first_reference(payload: &[u8], path: &[u32]) -> Result<u64, Error> {
    let values = references::identifiers(payload, path).map_err(map_fault)?;
    values.first().copied().ok_or(Error::InvalidSource)
}

fn document_language(payload: &[u8]) -> Result<Option<String>, Error> {
    let super_payload = match references::records(payload, &[DOCUMENT_SUPER_FIELD])
        .map_err(map_fault)?
        .as_slice()
    {
        [record] => record,
        [] => return Ok(None),
        _ => return Err(Error::InvalidSource),
    };
    let view =
        WireView::parse(super_payload, WireLimits::DEFAULT).map_err(|_| Error::InvalidSource)?;
    let field =
        drawable_graph::unique_field(&view, DOCUMENT_LANGUAGE_FIELD, 2).map_err(map_fault)?;
    field
        .map(|field| String::from_utf8(field.payload().to_vec()).map_err(|_| Error::InvalidSource))
        .transpose()
}

/// Resolve the document stylesheet, caption paragraph preset and language.
fn caption_theme(package: &Package) -> Result<(u64, u64, Option<String>), Error> {
    let document = drawable_graph::object(package, 1).map_err(map_fault)?;
    let document_index = structure::unique_message(document, &[1]).map_err(map_fault)?;
    let document_payload = &document.messages[document_index].data;
    let stylesheet = first_reference(document_payload, &[DOCUMENT_STYLESHEET_FIELD])?;
    let theme = first_reference(document_payload, &[DOCUMENT_THEME_FIELD])?;
    let theme_object = drawable_graph::object(package, theme).map_err(map_fault)?;
    let theme_index =
        structure::unique_message(theme_object, &[THEME_MESSAGE_TYPE]).map_err(map_fault)?;
    let theme_payload = &theme_object.messages[theme_index].data;
    let paragraph = references::identifiers(
        theme_payload,
        &[1, THEME_APPLICATION_FIELD, THEME_CAPTION_STYLE_FIELD],
    )
    .map_err(map_fault)?
    .first()
    .copied()
    .ok_or(Error::InvalidSource)?;
    unique_nonzero(stylesheet)?;
    unique_nonzero(paragraph)?;
    Ok((stylesheet, paragraph, document_language(document_payload)?))
}

fn chart_width(package: &Package, payload: &[u8]) -> Result<Option<f32>, Error> {
    let mut budget = ChartBudget::for_package(package).map_err(map_arrangement_error)?;
    let chart =
        parse_wire_view_with_budget(payload, 1, &mut budget).map_err(map_arrangement_error)?;
    let Some(drawable) =
        drawable_graph::unique_field(&chart, CHART_DRAWABLE_FIELD, 2).map_err(map_fault)?
    else {
        return Ok(None);
    };
    let drawable = parse_wire_view_with_budget(drawable.payload(), 2, &mut budget)
        .map_err(map_arrangement_error)?;
    let Some(geometry) =
        drawable_graph::unique_field(&drawable, DRAWABLE_GEOMETRY_FIELD, 2).map_err(map_fault)?
    else {
        return Ok(None);
    };
    let geometry = parse_wire_view_with_budget(geometry.payload(), 3, &mut budget)
        .map_err(map_arrangement_error)?;
    let Some(size) =
        drawable_graph::unique_field(&geometry, GEOMETRY_SIZE_FIELD, 2).map_err(map_fault)?
    else {
        return Ok(None);
    };
    let size = parse_wire_view_with_budget(size.payload(), 3, &mut budget)
        .map_err(map_arrangement_error)?;
    let Some(width) =
        drawable_graph::unique_field(&size, SIZE_WIDTH_FIELD, 5).map_err(map_fault)?
    else {
        return Ok(None);
    };
    let bytes: [u8; 4] = width
        .payload()
        .try_into()
        .map_err(|_| Error::InvalidSource)?;
    Ok(Some(f32::from_le_bytes(bytes)))
}

/// Resolve a semantic chart selector and prove ownership of its caption graph.
fn resolve<'sheet>(
    package: &Package,
    sheet: impl Into<SheetSelector<'sheet>>,
    chart: impl Into<ChartSelector>,
) -> Result<Target, Error> {
    let mut budget = ChartBudget::for_package(package).map_err(map_arrangement_error)?;
    let selection =
        select_chart_with_budget(package, sheet.into(), chart.into(), false, &mut budget)
            .map_err(map_arrangement_error)?;
    let payload = selected_chart_payload(package, &selection).map_err(map_arrangement_error)?;
    let edge = caption_codec::decode_chart_caption(
        payload,
        caption_codec::DecodeOptions::for_source(payload),
    )
    .map_err(|_| Error::InvalidSource)?;
    let caption_identifier =
        unique_nonzero(edge.caption_identifier().ok_or(Refusal::NoCaptionEdge)?)?;
    let resolved = package
        .state
        .index
        .resolve_ref_id(&package.state.components, caption_identifier.get())
        .map_err(|_| Error::InvalidSource)?
        .ok_or(Error::InvalidSource)?;
    let caption = object_from_resolved(package, resolved).map_err(map_arrangement_error)?;
    if resolved.component_index != selection.component_index {
        return Err(Error::InvalidSource);
    }
    validate_message_metadata(caption).map_err(map_arrangement_error)?;
    let component = package
        .state
        .components
        .catalog()
        .get_index(selection.component_index)
        .ok_or(Error::InvalidSource)?;
    let component_name: Arc<str> = Arc::from(component.name());
    if let Some((_info_index, info)) =
        unique_typed_message(caption, CAPTION_INFO_MESSAGE_TYPE).map_err(map_arrangement_error)?
    {
        let facts = pages_movie_caption_codec::decode_caption_info(
            &info.data,
            pages_movie_caption_codec::DecodeOptions::for_source(&info.data),
        )
        .map_err(|_| Error::InvalidSource)?;
        if facts.child_info_kind() != Some(1)
            || facts.parent_identifier() != Some(selection.chart_identifier)
            || facts.is_text_box() != Some(true)
        {
            return Err(Error::InvalidSource);
        }
        let storage = unique_nonzero(
            facts
                .owned_storage_identifier()
                .ok_or(Error::InvalidSource)?,
        )?;
        let style = unique_nonzero(facts.style_identifier().ok_or(Error::InvalidSource)?)?;
        let placement = unique_nonzero(facts.placement_identifier().ok_or(Error::InvalidSource)?)?;
        if facts.deprecated_storage_identifier() != Some(storage.get()) {
            return Err(Error::InvalidSource);
        }
        let identifiers = [caption_identifier, style, storage, placement];
        for (index, identifier) in identifiers.iter().enumerate() {
            if identifiers[..index].contains(identifier) {
                return Err(Error::InvalidSource);
            }
        }
        for (identifier, message_type, private) in [
            (storage, STORAGE_MESSAGE_TYPE, true),
            (placement, CAPTION_PLACEMENT_MESSAGE_TYPE, true),
            (style, SHAPE_STYLE_MESSAGE_TYPE, false),
        ] {
            let located = package
                .state
                .index
                .resolve_ref_id(&package.state.components, identifier.get())
                .map_err(|_| Error::InvalidSource)?
                .ok_or(Error::InvalidSource)?;
            if private && located.component_index != selection.component_index {
                return Err(Error::InvalidSource);
            }
            let object = object_from_resolved(package, located).map_err(map_arrangement_error)?;
            let (message_index, _) = unique_typed_message(object, message_type)
                .map_err(map_arrangement_error)?
                .ok_or(Error::InvalidSource)?;
            validate_message_metadata(object).map_err(map_arrangement_error)?;
            if message_index >= object.messages.len() {
                return Err(Error::InvalidSource);
            }
        }
        let storage_object = drawable_graph::object(package, storage.get()).map_err(map_fault)?;
        let storage_index = structure::unique_message(storage_object, &[STORAGE_MESSAGE_TYPE])
            .map_err(map_fault)?;
        let storage_payload = &storage_object.messages[storage_index].data;
        let text = litchi_iwa_text_wire::from_bytes(
            storage_payload,
            litchi_iwa_text_wire::Limits::default(),
        )
        .map_err(|_| Error::InvalidSource)?;
        return Ok(Target {
            selection: selection.clone(),
            component_index: resolved.component_index,
            component_name,
            chart_identifier: unique_nonzero(selection.chart_identifier)?,
            chart_message_index: selection.message_index,
            chart_message_type: CHART_MESSAGE_TYPE,
            caption_identifier,
            storage: Some(storage),
            text: Some(text.text().into()),
            width: chart_width(package, payload)?,
        });
    }
    let standin =
        unique_typed_message(caption, STANDIN_MESSAGE_TYPE).map_err(map_arrangement_error)?;
    if standin.is_none() || caption.messages.len() != 1 {
        return Err(Refusal::SharedObject.into());
    }
    Ok(Target {
        selection: selection.clone(),
        component_index: resolved.component_index,
        component_name,
        chart_identifier: unique_nonzero(selection.chart_identifier)?,
        chart_message_index: selection.message_index,
        chart_message_type: CHART_MESSAGE_TYPE,
        caption_identifier,
        storage: None,
        text: None,
        width: chart_width(package, payload)?,
    })
}

fn rewrite_chart_edge(payload: &[u8], replacement: u64) -> Result<Vec<u8>, Error> {
    if replacement == 0 {
        return Err(Error::InvalidSource);
    }
    let options = caption_codec::DecodeOptions::new(
        payload.len().saturating_add(32).max(1),
        payload.len().saturating_mul(4).saturating_add(8).max(1),
        payload.len().saturating_mul(8).saturating_add(8).max(1),
        8,
    )
    .with_max_output_bytes(payload.len().saturating_add(32).max(1));
    let output = caption_codec::rewrite_chart_caption(
        payload,
        caption_codec::ChartCaptionWrite::new(replacement),
        options,
    )
    .map_err(|_| Error::InvalidSource)?;
    let verified = caption_codec::decode_chart_caption(
        &output,
        caption_codec::DecodeOptions::for_source(&output),
    )
    .map_err(|_| Error::Verification)?;
    if verified.caption_identifier() != Some(replacement) {
        return Err(Error::Verification);
    }
    Ok(output)
}

/// Replace exactly one caption edge and its archive-header reference facts.
fn replace_chart_edge(
    working: &mut Working<'_>,
    target: &Target,
    replacement: u64,
) -> Result<(), Error> {
    let object = working
        .object(target.chart_identifier.get())
        .map_err(map_fault)?;
    let info = object
        .archive_info
        .message_infos
        .get(target.chart_message_index)
        .ok_or(Error::InvalidSource)?
        .clone();
    let message = object
        .messages
        .get(target.chart_message_index)
        .ok_or(Error::InvalidSource)?;
    if message.type_ != target.chart_message_type {
        return Err(Error::InvalidSource);
    }
    if info
        .data_references
        .contains(&target.caption_identifier.get())
        || info
            .object_references
            .iter()
            .filter(|identifier| **identifier == target.caption_identifier.get())
            .count()
            != 1
    {
        return Err(Error::InvalidSource);
    }
    let mut field_path = None;
    for field in &info.field_infos {
        let count = field
            .object_references
            .iter()
            .filter(|identifier| **identifier == target.caption_identifier.get())
            .count();
        if field
            .data_references
            .contains(&target.caption_identifier.get())
            || field.object_references.contains(&replacement)
        {
            return Err(Error::InvalidSource);
        }
        if count == 0 {
            continue;
        }
        if count != 1
            || !matches!(
                field.path.path.as_slice(),
                [DRAWABLE_CAPTION_FIELD, 1]
                    | [1, DRAWABLE_CAPTION_FIELD, 1]
                    | [1, 1, DRAWABLE_CAPTION_FIELD, 1]
            )
        {
            return Err(Error::InvalidSource);
        }
        if field_path.replace(field.path.path.clone()).is_some() {
            return Err(Error::InvalidSource);
        }
    }
    let field_path = field_path.ok_or(Error::InvalidSource)?;
    let current = caption_codec::decode_chart_caption(
        &message.data,
        caption_codec::DecodeOptions::for_source(&message.data),
    )
    .map_err(|_| Error::InvalidSource)?;
    if current.caption_identifier() != Some(target.caption_identifier.get()) {
        return Err(Error::PatchConflict);
    }
    let data = rewrite_chart_edge(&message.data, replacement)?;
    let limits = working.archive_limits();
    let rewritten = working
        .object_mut(target.chart_identifier.get())
        .map_err(map_fault)?;
    references::replace(
        rewritten,
        target.chart_message_index,
        data,
        &[target.caption_identifier.get()],
        &[replacement],
        &field_path,
        limits,
    )
    .map_err(map_fault)
}

fn new_versioned_object(
    limits: litchi_iwa_core::Limits,
    identifier: u64,
    message_type: u32,
    data: Vec<u8>,
    references: Vec<u64>,
) -> Result<ArchiveObject, Error> {
    let mut object =
        structure::new_object(identifier, None, message_type, data, references, limits)
            .map_err(map_fault)?;
    let info = object
        .archive_info
        .message_infos
        .first_mut()
        .ok_or(Error::InvalidSource)?;
    info.versions = STANDARD_VERSION.to_vec().into();
    Ok(object)
}

fn caption_graph(
    source: &Package,
    target: &Target,
    first: u64,
    text: &str,
) -> Result<Vec<ArchiveObject>, Error> {
    let width = target
        .width
        .filter(|width| width.is_finite() && *width > 0.0)
        .ok_or(Refusal::NoWidth)?;
    let (stylesheet, paragraph, language) = caption_theme(source)?;
    let working = Working::new(source).map_err(map_fault)?;
    let limits = working.archive_limits();
    let ids = [
        first,
        first.checked_add(1).ok_or(Error::InvalidSource)?,
        first.checked_add(2).ok_or(Error::InvalidSource)?,
        first.checked_add(3).ok_or(Error::InvalidSource)?,
    ];
    let output = graph_codec::encode_caption_graph_with_kind_with_report(
        graph_codec::CaptionGraphWrite {
            drawable_identifier: target.chart_identifier.get(),
            style_identifier: ids[0],
            info_identifier: ids[1],
            storage_identifier: ids[2],
            placement_identifier: ids[3],
            stylesheet_identifier: stylesheet,
            paragraph_style_identifier: paragraph,
            drawable_width: width,
            text,
            language: language.as_deref(),
        },
        graph_codec::CaptionGraphKind::Caption,
        graph_codec::EncodeOptions::for_text(text),
    )
    .map_err(|error| match error {
        graph_codec::EncodeError::Resource(_) => Error::Limit {
            kind: LimitKind::OutputBytes,
            observed: text.len() as u64,
            maximum: text.len() as u64,
        },
        graph_codec::EncodeError::Allocation { amount } => Error::Allocation { amount },
        _ => Error::InvalidSource,
    })?;
    let [style, info, storage, placement] = output.into_payloads().into_parts();
    let mut objects = Vec::new();
    objects
        .try_reserve_exact(4)
        .map_err(|_| Error::Allocation { amount: 4 })?;
    objects.push(new_versioned_object(
        limits,
        ids[0],
        SHAPE_STYLE_MESSAGE_TYPE,
        style,
        vec![paragraph],
    )?);
    objects.push(new_versioned_object(
        limits,
        ids[1],
        CAPTION_INFO_MESSAGE_TYPE,
        info,
        vec![ids[0], ids[2], ids[3]],
    )?);
    objects.push(new_versioned_object(
        limits,
        ids[2],
        STORAGE_MESSAGE_TYPE,
        storage,
        vec![paragraph],
    )?);
    objects.push(new_versioned_object(
        limits,
        ids[3],
        CAPTION_PLACEMENT_MESSAGE_TYPE,
        placement,
        Vec::new(),
    )?);
    Ok(objects)
}

fn standin_object(working: &Working<'_>, identifier: u64) -> Result<ArchiveObject, Error> {
    new_versioned_object(
        working.archive_limits(),
        identifier,
        STANDIN_MESSAGE_TYPE,
        graph_codec::canonical_standin_payload().to_vec(),
        Vec::new(),
    )
    .map(|mut object| {
        if let Some(info) = object.archive_info.message_infos.first_mut() {
            info.versions = STANDIN_VERSION.to_vec().into();
        }
        object
    })
}

fn register_new_objects(
    working: &mut Working<'_>,
    registry: &Registry,
    transition: &mut Transition,
    component: &str,
    objects: &[ArchiveObject],
) -> Result<(), Error> {
    let component_id = registry
        .component_for_entry(component)
        .map_err(map_fault)?
        .identifier;
    for object in objects {
        let identifier = object.archive_info.identifier.ok_or(Error::InvalidSource)?;
        working
            .insert(component, object.clone())
            .map_err(map_fault)?;
        let seed = metadata::identity_seed(working, identifier).map_err(map_fault)?;
        transition
            .register_uuid(registry, component_id, identifier, seed)
            .map_err(map_fault)?;
        let message = object
            .archive_info
            .message_infos
            .first()
            .ok_or(Error::InvalidSource)?;
        for reference in &message.object_references {
            let target_entry = working
                .locate(*reference)
                .map_err(map_fault)?
                .ok_or(Error::InvalidSource)?;
            let target_component = registry
                .component_for_entry(&target_entry)
                .map_err(map_fault)?
                .identifier;
            transition
                .reference(registry, component_id, target_component, *reference)
                .map_err(map_fault)?;
        }
    }
    Ok(())
}

fn publish_graph(
    source: &Package,
    target: &Target,
    replacement: u64,
    objects: &[ArchiveObject],
) -> Result<Package, Error> {
    let mut working = Working::new(source).map_err(map_fault)?;
    working.keep_previews();
    let registry = metadata::registry(&working)
        .map_err(map_fault)?
        .ok_or(Refusal::UnsupportedSource)?;
    let mut transition = Transition::default();
    register_new_objects(
        &mut working,
        &registry,
        &mut transition,
        target.component_name.as_ref(),
        objects,
    )?;
    replace_chart_edge(&mut working, target, replacement)?;
    working
        .publish(&transition)
        .map(|published| published.package)
        .map_err(map_fault)
}

fn publish_storage(source: &Package, target: &Target, replacement: &str) -> Result<Package, Error> {
    let storage = target.storage.ok_or(Error::InvalidSource)?;
    let storage_object = drawable_graph::object(source, storage.get()).map_err(map_fault)?;
    let storage_index =
        structure::unique_message(storage_object, &[STORAGE_MESSAGE_TYPE]).map_err(map_fault)?;
    let payload = &storage_object.messages[storage_index].data;
    let before = litchi_iwa_text_wire::from_bytes(payload, litchi_iwa_text_wire::Limits::default())
        .map_err(|_| Error::InvalidSource)?;
    let end = before.text().encode_utf16().count();
    let prepared = litchi_iwa_text_wire::prepare_storage_text_rewrite_with_behavior(
        payload,
        0..end,
        replacement,
        litchi_iwa_text_wire::RewriteBehavior::ReplaceSelection,
        litchi_iwa_text_wire::RewriteLimits::default(),
    )
    .map_err(|_| Error::InvalidSource)?;
    let requirements = prepared.execution_requirements();
    let rewritten = prepared
        .execute(litchi_iwa_text_wire::StorageRewriteExecutionLimits {
            max_output_bytes: requirements.output_bytes(),
            max_retained_elements: requirements.retained_elements(),
            max_retained_bytes: requirements.retained_bytes(),
            max_peak_scratch_bytes: requirements.peak_scratch_bytes(),
            max_allocations: requirements.allocations(),
            max_work: requirements.work(),
        })
        .map_err(|_| Error::InvalidSource)?;
    if !rewritten.removed_object_references().is_empty()
        || !rewritten.removed_object_references_by_field().is_empty()
        || rewritten.object_reference_occurrences_before()
            != rewritten.object_reference_occurrences_after()
    {
        return Err(Refusal::DependentContent.into());
    }
    let mut working = Working::new(source).map_err(map_fault)?;
    working.keep_previews();
    let limits = working.archive_limits();
    let object = working.object_mut(storage.get()).map_err(map_fault)?;
    object
        .replace_message_preserving_header(
            storage_index,
            RawMessage {
                type_: STORAGE_MESSAGE_TYPE,
                data: rewritten.into_bytes(),
            },
            limits,
        )
        .map_err(|_| Error::InvalidSource)?;
    working
        .publish(&Transition::default())
        .map(|published| published.package)
        .map_err(map_fault)
}

fn chart_componentized(package: &Package, target: &Target) -> Result<bool, Error> {
    let component = package
        .state
        .components
        .catalog()
        .get_index(target.component_index)
        .ok_or(Error::InvalidSource)?;
    Ok(component
        .archive()
        .objects
        .iter()
        .any(|object| object.messages.iter().any(|message| message.type_ == 4_000)))
}

/// An immutable edit staged against one exact Numbers package snapshot.
pub struct Edit<'a> {
    source: &'a Package,
    target: Target,
    operation: Option<Option<Box<str>>>,
}

impl fmt::Debug for Edit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("sheet", &self.target.selection.sheet_position)
            .field("chart", &self.target.selection.chart_position)
            .field("staged", &self.operation.is_some())
            .finish_non_exhaustive()
    }
}

impl Edit<'_> {
    /// Caption text observed when this edit began.
    #[must_use]
    pub fn before(&self) -> Option<&str> {
        self.target.text.as_deref()
    }

    /// Caption text staged for publication.
    #[must_use]
    pub fn after(&self) -> Option<&str> {
        self.operation
            .as_ref()
            .map(|value| value.as_deref())
            .unwrap_or(self.target.text.as_deref())
    }

    /// Stage a visible caption.
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

    /// Stage removal of the visible caption.
    pub fn clear(mut self) -> Result<Self, Error> {
        self.operation = Some(None);
        Ok(self)
    }

    /// Validate and publish the exact-source candidate.
    pub fn commit(self) -> Result<Commit, Error> {
        let after = self
            .operation
            .clone()
            .unwrap_or_else(|| self.target.text.clone());
        if after == self.target.text {
            return Ok(Commit::unchanged(self.source, self.target, after));
        }
        if !self.source.state.source.source_is_exact() {
            return Err(Refusal::UnsupportedSource.into());
        }
        let current = resolve(
            self.source,
            SheetSelector::position(self.target.selection.sheet_position),
            ChartSelector::position(self.target.selection.chart_position),
        )?;
        if !current.same_route(&self.target) || current.text != self.target.text {
            return Err(Error::PatchConflict);
        }
        let package = match (&after, current.storage) {
            (Some(text), Some(_storage)) => publish_storage(self.source, &current, text)?,
            (Some(text), None) => {
                if chart_componentized(self.source, &current)? {
                    return Err(Refusal::UnsupportedLayout.into());
                }
                let working = Working::new(self.source).map_err(map_fault)?;
                let first = working.next_identifier();
                let objects = caption_graph(self.source, &current, first, text)?;
                let ids = objects
                    .iter()
                    .map(|object| object.archive_info.identifier.ok_or(Error::InvalidSource))
                    .collect::<Result<Vec<_>, _>>()?;
                let replacement = *ids.get(1).ok_or(Error::InvalidSource)?;
                publish_graph(self.source, &current, replacement, &objects)?
            },
            (None, Some(_storage)) => {
                let working = Working::new(self.source).map_err(map_fault)?;
                let standin = working.next_identifier();
                let object = standin_object(&working, standin)?;
                publish_graph(self.source, &current, standin, &[object])?
            },
            (None, None) => return Ok(Commit::unchanged(self.source, current, after)),
        };
        let verified = resolve(
            &package,
            SheetSelector::position(current.selection.sheet_position),
            ChartSelector::position(current.selection.chart_position),
        )?;
        if verified.text.as_deref() != after.as_deref() {
            return Err(Error::Verification);
        }
        Ok(Commit::published(self.source, package, current, after))
    }
}

/// Exact-source reversible chart-caption patch.
#[derive(Clone, PartialEq)]
pub struct Patch {
    artifacts: OwnedExactArtifacts,
    sheet: Position,
    chart: Position,
    before: Option<Box<str>>,
    after: Option<Box<str>>,
}

impl fmt::Debug for Patch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Patch")
            .field("sheet", &self.sheet)
            .field("chart", &self.chart)
            .finish_non_exhaustive()
    }
}

impl Patch {
    /// Whether both semantic text and exact package bytes are unchanged.
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

    /// The exact target-to-source inverse patch.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            artifacts: self.artifacts.inverse(),
            sheet: self.sheet,
            chart: self.chart,
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
    /// Whether the candidate package bytes changed.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }
}

/// A verified immutable chart-caption commit.
#[must_use = "a chart-caption commit contains the validated package snapshot"]
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

    /// Consume the commit and return package plus exact-source patch.
    #[must_use]
    pub fn into_parts(self) -> (Package, Patch) {
        (self.package, self.patch)
    }

    /// Borrow the reversible exact-source patch.
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
                sheet: target.selection.sheet_position,
                chart: target.selection.chart_position,
                before: target.text,
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
    ) -> Self {
        let source_owner = source.state.source.clone();
        let target_owner = package.state.source.clone();
        Self {
            package,
            patch: Patch {
                artifacts: OwnedExactArtifacts::new(source_owner, target_owner),
                sheet: target.selection.sheet_position,
                chart: target.selection.chart_position,
                before: target.text,
                after,
            },
            diagnostics: Diagnostics { changed: true },
        }
    }
}

fn read<'sheet>(
    package: &Package,
    sheet: impl Into<SheetSelector<'sheet>>,
    chart: impl Into<ChartSelector>,
) -> Result<Option<String>, Error> {
    Ok(resolve(package, sheet, chart)?.text.map(Into::into))
}

fn edit<'a>(
    package: &'a Package,
    sheet: impl Into<SheetSelector<'a>>,
    chart: impl Into<ChartSelector>,
) -> Result<Edit<'a>, Error> {
    Ok(Edit {
        source: package,
        target: resolve(package, sheet, chart)?,
        operation: None,
    })
}

fn apply_patch(package: &Package, patch: &Patch) -> Result<Commit, Error> {
    if !patch.artifacts.authorizes_owner(&package.state.source) {
        return Err(Error::PatchConflict);
    }
    let current = resolve(
        package,
        SheetSelector::position(patch.sheet),
        ChartSelector::position(patch.chart),
    )?;
    if current.text != patch.before {
        return Err(Error::PatchConflict);
    }
    if patch.is_noop() {
        return Ok(Commit::unchanged(package, current, patch.after.clone()));
    }
    let candidate = Package::from_source_owner_with_options(
        patch.artifacts.target_owner(),
        package.state.options,
    )
    .map_err(|_| Error::Verification)?;
    let verified = resolve(
        &candidate,
        SheetSelector::position(patch.sheet),
        ChartSelector::position(patch.chart),
    )?;
    if verified.text != patch.after {
        return Err(Error::Verification);
    }
    Ok(Commit {
        package: candidate,
        patch: patch.clone(),
        diagnostics: Diagnostics { changed: true },
    })
}

impl Package {
    /// Read one selected sheet-chart caption.
    pub fn sheet_chart_caption<'sheet>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        chart: impl Into<ChartSelector>,
    ) -> Result<Option<String>, Error> {
        read(self, sheet, chart)
    }

    /// Begin an immutable exact-source edit of one selected chart caption.
    pub fn edit_sheet_chart_caption<'sheet>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        chart: impl Into<ChartSelector>,
    ) -> Result<Edit<'_>, Error> {
        edit(self, sheet, chart)
    }

    /// Apply an exact-source chart-caption patch to this package.
    pub fn apply_sheet_chart_caption(&self, patch: &Patch) -> Result<Commit, Error> {
        apply_patch(self, patch)
    }
}
