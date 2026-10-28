//! Selector-first Pages body-chart caption transactions.
//!
//! A chart caption is a native reference edge plus a bounded four-object text
//! graph.  Existing captions rewrite only their private storage.  Creating a
//! caption authors the same inline graph as the existing body-drawable label
//! family, registers its fresh objects, and repoints only the selected chart
//! edge.  Clearing keeps the old graph and attaches a fresh native stand-in,
//! matching the application's undo-preserving behavior.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "The focused chart-caption transaction keeps graph phases together."
)]

use std::fmt;
use std::num::NonZeroU64;
use std::sync::Arc;

use litchi_core::Position;
use litchi_iwa_archive::package::OwnedExactArtifacts;
use litchi_iwa_core::{Archive, ArchiveObject, RawMessage};
use litchi_iwa_protos::pages_movie_caption_codec;
use litchi_iwa_style::chart_text::{caption as caption_codec, caption_graph as graph_codec};
use thiserror::Error as ThisError;

use super::Package;
use super::body_chart_arrangement::{self as arrangement, ArrangementBudget, ChartTarget};
use super::body_drawable_graph::{self as graph, Budget as GraphBudget, ObjectIndex};
use super::body_storage_objects as objects;
use crate::selector::BodyChartSelector;

const CAPTION_INFO_MESSAGE_TYPE: u32 = 633;
const CAPTION_PLACEMENT_MESSAGE_TYPE: u32 = 634;
const SHAPE_STYLE_MESSAGE_TYPE: u32 = 2_025;
const STORAGE_MESSAGE_TYPE: u32 = 2_001;
const STANDIN_MESSAGE_TYPE: u32 = 3_097;
const CHART_DRAWABLE_FIELD: u32 = 1;
const DRAWABLE_CAPTION_FIELD: u32 = 11;
const DRAWABLE_GEOMETRY_FIELD: u32 = 1;
const GEOMETRY_SIZE_FIELD: u32 = 2;
const SIZE_WIDTH_FIELD: u32 = 1;
const STANDARD_VERSION: [u32; 3] = [1, 0, 5];
const STANDIN_VERSION: [u32; 3] = [10, 1, 0];

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// One native payload or graph object inspected.
    PayloadBytes,
    /// Authored graph/output bytes.
    OutputBytes,
    /// Bounded graph allocations.
    Allocations,
}

impl fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::PayloadBytes => "payload bytes",
            Self::OutputBytes => "output bytes",
            Self::Allocations => "allocations",
        })
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Refusal {
    #[error("the Pages body has no chart at the requested position")]
    ChartNotFound,
    #[error("the selected chart has no caption edge")]
    NoCaptionEdge,
    #[error("the selected chart caption graph is shared or malformed")]
    SharedObject,
    #[error("the selected chart has no displayed width for a new caption")]
    NoWidth,
    #[error("the caption text contains a reserved native marker")]
    ReservedMarker,
    #[error("the existing caption contains dependent native content")]
    DependentContent,
    #[error("this Pages source does not support exact chart-caption edits")]
    UnsupportedSource,
}

impl Refusal {
    #[must_use]
    pub const fn kind(self) -> litchi_core::edit::RefusalKind {
        use litchi_core::edit::RefusalKind;
        match self {
            Self::ChartNotFound => RefusalKind::NotFound,
            Self::NoCaptionEdge | Self::SharedObject | Self::NoWidth | Self::UnsupportedSource => {
                RefusalKind::UnsupportedContent
            },
            Self::ReservedMarker => RefusalKind::InvalidValue,
            Self::DependentContent => RefusalKind::UnsupportedContent,
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Error {
    #[error(transparent)]
    Refused(Refusal),
    #[error("the selected Pages body-chart caption source is invalid")]
    InvalidSource,
    #[error(
        "Pages body-chart caption {kind} limit exceeded: observed {observed}, maximum {maximum}"
    )]
    Limit {
        kind: LimitKind,
        observed: u64,
        maximum: u64,
    },
    #[error("could not allocate {amount} Pages body-chart caption units")]
    Allocation { amount: usize },
    #[error("the edited Pages body-chart caption failed semantic verification")]
    Verification,
    #[error("the Pages body-chart caption patch does not match its exact source")]
    PatchConflict,
}

impl From<Refusal> for Error {
    fn from(value: Refusal) -> Self {
        Self::Refused(value)
    }
}

#[derive(Debug, Clone, PartialEq)]
struct Target {
    chart: ChartTarget,
    component_index: usize,
    component_name: Arc<str>,
    drawable_object_index: usize,
    drawable_identifier: NonZeroU64,
    drawable_message_index: usize,
    drawable_message_type: u32,
    caption_identifier: NonZeroU64,
    storage: Option<NonZeroU64>,
    text: Option<Box<str>>,
    width: Option<f32>,
}

impl Target {
    fn same_route(&self, other: &Self) -> bool {
        self.chart.same_identity(&other.chart)
            && self.component_index == other.component_index
            && self.component_name == other.component_name
            && self.drawable_object_index == other.drawable_object_index
            && self.drawable_identifier == other.drawable_identifier
            && self.drawable_message_index == other.drawable_message_index
            && self.drawable_message_type == other.drawable_message_type
            && self.caption_identifier == other.caption_identifier
    }
}

fn map_arrangement_error(error: arrangement::Error) -> Error {
    match error {
        arrangement::Error::Refused(arrangement::Refusal::ChartNotFound { .. })
        | arrangement::Error::Refused(arrangement::Refusal::ChartPositionNotFound { .. }) => {
            Refusal::ChartNotFound.into()
        },
        arrangement::Error::Refused(arrangement::Refusal::UnsupportedSource) => {
            Refusal::UnsupportedSource.into()
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

fn map_graph_error(error: graph::GraphError) -> Error {
    match error {
        graph::GraphError::Limit {
            observed, maximum, ..
        } => Error::Limit {
            kind: LimitKind::PayloadBytes,
            observed,
            maximum,
        },
        graph::GraphError::Allocation { amount } => Error::Allocation { amount },
        graph::GraphError::InvalidSource => Error::InvalidSource,
    }
}

fn width(package: &Package, payload: &[u8]) -> Result<Option<f32>, Error> {
    let mut budget = GraphBudget::new(package).map_err(map_graph_error)?;
    let chart = budget.parse(payload, 1).map_err(map_graph_error)?;
    let Some(drawable) =
        graph::unique_field(&chart, CHART_DRAWABLE_FIELD, 2).map_err(map_graph_error)?
    else {
        return Ok(None);
    };
    let drawable = budget
        .parse(drawable.payload(), 2)
        .map_err(map_graph_error)?;
    let Some(geometry) =
        graph::unique_field(&drawable, DRAWABLE_GEOMETRY_FIELD, 2).map_err(map_graph_error)?
    else {
        return Ok(None);
    };
    let geometry = budget
        .parse(geometry.payload(), 3)
        .map_err(map_graph_error)?;
    let Some(size) =
        graph::unique_field(&geometry, GEOMETRY_SIZE_FIELD, 2).map_err(map_graph_error)?
    else {
        return Ok(None);
    };
    let size = budget.parse(size.payload(), 3).map_err(map_graph_error)?;
    let Some(value) = graph::unique_field(&size, SIZE_WIDTH_FIELD, 5).map_err(map_graph_error)?
    else {
        return Ok(None);
    };
    let bytes: [u8; 4] = value
        .payload()
        .try_into()
        .map_err(|_| Error::InvalidSource)?;
    Ok(Some(f32::from_le_bytes(bytes)))
}

fn resolve(package: &Package, selector: BodyChartSelector) -> Result<Target, Error> {
    let mut arrangement_budget = ArrangementBudget::new(package).map_err(map_arrangement_error)?;
    let chart = arrangement::resolve_target(package, selector, &mut arrangement_budget)
        .map_err(map_arrangement_error)?;
    let payload =
        arrangement::selected_chart_payload(package, &chart).map_err(map_arrangement_error)?;
    let snapshot = caption_codec::decode_chart_caption(
        payload,
        caption_codec::DecodeOptions::for_source(payload),
    )
    .map_err(|_| Error::InvalidSource)?;
    let caption_identifier = snapshot
        .caption_identifier()
        .and_then(NonZeroU64::new)
        .ok_or(Refusal::NoCaptionEdge)?;
    let mut graph_budget = GraphBudget::new(package).map_err(map_graph_error)?;
    let index = ObjectIndex::new(package, &mut graph_budget).map_err(map_graph_error)?;
    let caption = index.get(caption_identifier).map_err(map_graph_error)?;
    if caption.component_index != chart.component_index {
        return Err(Error::InvalidSource);
    }
    let caption_info = graph::unique_optional_message(caption.object, CAPTION_INFO_MESSAGE_TYPE)
        .map_err(map_graph_error)?;
    if let Some((message_index, message)) = caption_info {
        graph::validate_message_metadata(caption.object, message_index).map_err(map_graph_error)?;
        let facts = pages_movie_caption_codec::decode_caption_info(
            &message.data,
            pages_movie_caption_codec::DecodeOptions::for_source(&message.data),
        )
        .map_err(|_| Error::InvalidSource)?;
        if facts.child_info_kind() != Some(1)
            || facts.parent_identifier() != Some(chart.drawable_identifier.get())
            || facts.is_text_box() != Some(true)
        {
            return Err(Error::InvalidSource);
        }
        let storage = facts
            .owned_storage_identifier()
            .and_then(NonZeroU64::new)
            .ok_or(Error::InvalidSource)?;
        let style = facts
            .style_identifier()
            .and_then(NonZeroU64::new)
            .ok_or(Error::InvalidSource)?;
        let placement = facts
            .placement_identifier()
            .and_then(NonZeroU64::new)
            .ok_or(Error::InvalidSource)?;
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
            let located = index.get(identifier).map_err(map_graph_error)?;
            if private && located.component_index != chart.component_index {
                return Err(Error::InvalidSource);
            }
            let (index, _) =
                graph::unique_message(located.object, message_type).map_err(map_graph_error)?;
            graph::validate_message_metadata(located.object, index).map_err(map_graph_error)?;
        }
        let storage_data =
            objects::storage_by_identifier(package, storage.get()).map_err(map_graph_error)?;
        return Ok(Target {
            chart,
            component_index: chart.component_index,
            component_name: Arc::from(
                package
                    .state
                    .source
                    .components()
                    .get_index(chart.component_index)
                    .ok_or(Error::InvalidSource)?
                    .name(),
            ),
            drawable_object_index: chart.drawable_object_index,
            drawable_identifier: chart.drawable_identifier,
            drawable_message_index: chart.drawable_message_index,
            drawable_message_type: chart.drawable_message_type,
            caption_identifier,
            storage: Some(storage),
            text: Some(storage_data.text.into_boxed_str()),
            width: width(package, payload)?,
        });
    }
    let standin = graph::unique_optional_message(caption.object, STANDIN_MESSAGE_TYPE)
        .map_err(map_graph_error)?;
    if standin.is_some() {
        if caption.object.messages.len() != 1 {
            return Err(Error::InvalidSource);
        }
        Ok(Target {
            chart,
            component_index: chart.component_index,
            component_name: Arc::from(
                package
                    .state
                    .source
                    .components()
                    .get_index(chart.component_index)
                    .ok_or(Error::InvalidSource)?
                    .name(),
            ),
            drawable_object_index: chart.drawable_object_index,
            drawable_identifier: chart.drawable_identifier,
            drawable_message_index: chart.drawable_message_index,
            drawable_message_type: chart.drawable_message_type,
            caption_identifier,
            storage: None,
            text: None,
            width: width(package, payload)?,
        })
    } else {
        Err(Refusal::SharedObject.into())
    }
}

fn rewrite_chart_edge(
    source: &Package,
    target: &Target,
    replacement: u64,
) -> Result<Vec<u8>, Error> {
    let component = source
        .state
        .source
        .components()
        .get_index(target.component_index)
        .ok_or(Error::InvalidSource)?;
    let object = component
        .archive()
        .objects
        .get(target.drawable_object_index)
        .ok_or(Error::InvalidSource)?;
    let message = object
        .messages
        .get(target.drawable_message_index)
        .ok_or(Error::InvalidSource)?;
    if message.type_ != target.drawable_message_type {
        return Err(Error::InvalidSource);
    }
    let options = caption_codec::DecodeOptions::new(
        message.data.len().saturating_add(32).max(1),
        message
            .data
            .len()
            .saturating_mul(4)
            .saturating_add(8)
            .max(1),
        message
            .data
            .len()
            .saturating_mul(8)
            .saturating_add(8)
            .max(1),
        8,
    )
    .with_max_output_bytes(message.data.len().saturating_add(32).max(1));
    caption_codec::rewrite_chart_caption(
        &message.data,
        caption_codec::ChartCaptionWrite::new(replacement),
        options,
    )
    .map_err(|_| Error::InvalidSource)
}

fn replace_reference(references: &mut Vec<u64>, old: u64, new: u64) -> Result<(), Error> {
    let Some(index) = references.iter().position(|identifier| *identifier == old) else {
        return Err(Error::InvalidSource);
    };
    if references
        .iter()
        .filter(|identifier| **identifier == old)
        .count()
        != 1
        || references.contains(&new)
    {
        return Err(Error::SharedObject.into());
    }
    references[index] = new;
    Ok(())
}

fn replace_chart_message_references(
    object: &mut ArchiveObject,
    message_index: usize,
    old: u64,
    new: u64,
) -> Result<(), Error> {
    let info = object
        .archive_info
        .message_infos
        .get_mut(message_index)
        .ok_or(Error::InvalidSource)?;
    replace_reference(&mut info.object_references, old, new)?;
    let mut found = false;
    for field in &mut info.field_infos {
        let count = field
            .object_references
            .iter()
            .filter(|identifier| **identifier == old)
            .count();
        if field.data_references.contains(&old)
            || field.data_references.contains(&new)
            || field.object_references.contains(&new)
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
        if found {
            return Err(Error::InvalidSource);
        }
        found = true;
        replace_reference(&mut field.object_references, old, new)?;
    }
    if !found {
        return Err(Error::InvalidSource);
    }
    Ok(())
}

fn new_versioned_object(
    package: &Package,
    identifier: u64,
    message_type: u32,
    data: Vec<u8>,
    version: [u32; 3],
    references: Vec<u64>,
) -> Result<ArchiveObject, Error> {
    let mut object =
        objects::new_object(package, identifier, message_type, data).map_err(map_graph_error)?;
    let info = object
        .archive_info
        .message_infos
        .first_mut()
        .ok_or(Error::InvalidSource)?;
    info.versions = version.to_vec().into();
    info.object_references = references;
    Ok(object)
}

fn caption_graph(
    package: &Package,
    target: &Target,
    first: u64,
    text: &str,
) -> Result<(Vec<ArchiveObject>, [u64; 4]), Error> {
    let width = target
        .width
        .filter(|width| width.is_finite() && *width > 0.0)
        .ok_or(Refusal::NoWidth)?;
    let (stylesheet, paragraph_style, language) =
        super::body_drawable_caption::caption_theme(package).map_err(|_| Error::InvalidSource)?;
    let ids = [
        first,
        first.checked_add(1).ok_or(Error::InvalidSource)?,
        first.checked_add(2).ok_or(Error::InvalidSource)?,
        first.checked_add(3).ok_or(Error::InvalidSource)?,
    ];
    let output = graph_codec::encode_caption_graph_with_kind_with_report(
        graph_codec::CaptionGraphWrite {
            drawable_identifier: target.drawable_identifier.get(),
            style_identifier: ids[0],
            info_identifier: ids[1],
            storage_identifier: ids[2],
            placement_identifier: ids[3],
            stylesheet_identifier: stylesheet,
            paragraph_style_identifier: paragraph_style,
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
    let objects = vec![
        new_versioned_object(
            package,
            ids[0],
            SHAPE_STYLE_MESSAGE_TYPE,
            style,
            STANDARD_VERSION,
            vec![paragraph_style],
        )?,
        new_versioned_object(
            package,
            ids[1],
            CAPTION_INFO_MESSAGE_TYPE,
            info,
            STANDARD_VERSION,
            vec![ids[0], ids[2], ids[3]],
        )?,
        new_versioned_object(
            package,
            ids[2],
            STORAGE_MESSAGE_TYPE,
            storage,
            STANDARD_VERSION,
            vec![paragraph_style],
        )?,
        new_versioned_object(
            package,
            ids[3],
            CAPTION_PLACEMENT_MESSAGE_TYPE,
            placement,
            STANDARD_VERSION,
            Vec::new(),
        )?,
    ];
    Ok((objects, ids))
}

fn publish_graph(
    source: &Package,
    target: &Target,
    replacement: u64,
    created: Vec<ArchiveObject>,
) -> Result<Package, Error> {
    let archive_limits = source
        .state
        .source
        .limits()
        .effective_archive_limits()
        .map_err(|_| Error::InvalidSource)?;
    let mut archive =
        graph::editable_archive(source, target.component_name.as_ref()).map_err(map_graph_error)?;
    let mut created_ids = Vec::with_capacity(created.len());
    for object in &created {
        created_ids.push(object.archive_info.identifier.ok_or(Error::InvalidSource)?);
    }
    if created_ids.is_empty() || created_ids.iter().any(|id| *id == 0) {
        return Err(Error::InvalidSource);
    }
    for object in created {
        archive
            .insert_object(object, archive_limits)
            .map_err(|_| Error::InvalidSource)?;
    }
    let payload = rewrite_chart_edge(source, target, replacement)?;
    let chart = archive
        .object_mut(target.drawable_identifier.get())
        .ok_or(Error::InvalidSource)?;
    chart
        .replace_message_preserving_header(
            target.drawable_message_index,
            RawMessage {
                type_: target.drawable_message_type,
                data: payload,
            },
            archive_limits,
        )
        .map_err(|_| Error::InvalidSource)?;
    replace_chart_message_references(
        chart,
        target.drawable_message_index,
        target.caption_identifier.get(),
        replacement,
    )?;

    let mut archives = vec![(Arc::clone(&target.component_name), archive)];
    if let Some(metadata) = objects::metadata(source).map_err(map_graph_error)? {
        let component = metadata
            .component_for_entry(target.component_name.as_ref())
            .map_err(map_graph_error)?
            .ok_or(Error::InvalidSource)?;
        let last = created_ids
            .iter()
            .copied()
            .max()
            .ok_or(Error::InvalidSource)?;
        let data = metadata
            .registered(source, component, &created_ids, last)
            .map_err(map_graph_error)?;
        let mut metadata_archive = graph::editable_archive(source, metadata.component.as_ref())
            .map_err(map_graph_error)?;
        metadata_archive
            .object_mut(metadata.object)
            .ok_or(Error::InvalidSource)?
            .replace_message_preserving_header(
                metadata.message_index,
                RawMessage {
                    type_: 11_006,
                    data,
                },
                archive_limits,
            )
            .map_err(|_| Error::InvalidSource)?;
        archives.push((Arc::clone(&metadata.component), metadata_archive));
    }
    graph::publish_archives(source, archives).map_err(map_graph_error)
}

fn publish_standin(source: &Package, target: &Target, standin: u64) -> Result<Package, Error> {
    let object = new_versioned_object(
        source,
        standin,
        STANDIN_MESSAGE_TYPE,
        graph_codec::canonical_standin_payload().to_vec(),
        STANDIN_VERSION,
        Vec::new(),
    )?;
    publish_graph(source, target, standin, vec![object])
}

pub struct Edit<'a> {
    source: &'a Package,
    target: Target,
    operation: Option<Option<Box<str>>>,
}

impl fmt::Debug for Edit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("position", &self.target.chart.position)
            .field("staged", &self.operation.is_some())
            .finish_non_exhaustive()
    }
}

impl Edit<'_> {
    #[must_use]
    pub fn before(&self) -> Option<&str> {
        self.target.text.as_deref()
    }

    #[must_use]
    pub fn after(&self) -> Option<&str> {
        self.operation
            .as_ref()
            .map(|value| value.as_deref())
            .unwrap_or(self.target.text.as_deref())
    }

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

    pub fn clear(mut self) -> Result<Self, Error> {
        self.operation = Some(None);
        Ok(self)
    }

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
            BodyChartSelector::position(self.target.chart.position),
        )?;
        if !current.same_route(&self.target) || current.text != self.target.text {
            return Err(Error::PatchConflict);
        }
        let candidate = match (&after, current.storage) {
            (Some(text), Some(storage)) => {
                let mut located = objects::storage_by_identifier(self.source, storage.get())
                    .map_err(map_graph_error)?;
                let end = located.utf16_len();
                let removed = objects::splice_text(self.source, &mut located, 0, end, text)
                    .map_err(map_graph_error)?;
                if !removed.is_empty() {
                    return Err(Refusal::DependentContent.into());
                }
                graph::publish(
                    self.source,
                    vec![graph::MessageRewrite {
                        component: located.component,
                        object: located.identifier,
                        message_index: located.message_index,
                        message_type: located.message_type,
                        data: located.payload,
                    }],
                )
                .map_err(map_graph_error)?
            },
            (Some(text), None) => {
                let metadata = objects::metadata(self.source).map_err(map_graph_error)?;
                let first = objects::next_identifier(self.source, metadata.as_ref())
                    .map_err(map_graph_error)?;
                let (objects, ids) = caption_graph(self.source, &current, first, text)?;
                publish_graph(self.source, &current, ids[1], objects)?
            },
            (None, Some(_)) => {
                let metadata = objects::metadata(self.source).map_err(map_graph_error)?;
                let standin = objects::next_identifier(self.source, metadata.as_ref())
                    .map_err(map_graph_error)?;
                publish_standin(self.source, &current, standin)?
            },
            (None, None) => return Ok(Commit::unchanged(self.source, current, after)),
        };
        let verified = resolve(
            &candidate,
            BodyChartSelector::position(current.chart.position),
        )?;
        if verified.text.as_deref() != after.as_deref() {
            return Err(Error::Verification);
        }
        Ok(Commit::published(self.source, candidate, current, after))
    }
}

#[derive(Clone, PartialEq, Eq)]
pub struct Patch {
    artifacts: OwnedExactArtifacts,
    position: Position,
    before: Option<Box<str>>,
    after: Option<Box<str>>,
}

impl fmt::Debug for Patch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Patch")
            .field("position", &self.position)
            .finish_non_exhaustive()
    }
}

impl Patch {
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before == self.after && self.artifacts.is_byte_noop()
    }

    #[must_use]
    pub fn before(&self) -> Option<&str> {
        self.before.as_deref()
    }

    #[must_use]
    pub fn after(&self) -> Option<&str> {
        self.after.as_deref()
    }

    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            artifacts: self.artifacts.inverse(),
            position: self.position,
            before: self.after.clone(),
            after: self.before.clone(),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Diagnostics {
    changed: bool,
}

impl Diagnostics {
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }
}

#[must_use = "a chart-caption commit contains the validated package snapshot"]
#[derive(Debug)]
pub struct Commit {
    package: Package,
    patch: Patch,
    diagnostics: Diagnostics,
}

impl Commit {
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.package
    }

    #[must_use]
    pub fn into_package(self) -> Package {
        self.package
    }

    #[must_use]
    pub fn into_parts(self) -> (Package, Patch) {
        (self.package, self.patch)
    }

    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    #[must_use]
    pub const fn diagnostics(&self) -> &Diagnostics {
        &self.diagnostics
    }

    fn unchanged(source: &Package, target: Target, value: Option<Box<str>>) -> Self {
        let source_bytes = source.state.source.shared_source();
        let owner =
            litchi_iwa_archive::package::SharedBytes::from_shared_slice(Arc::clone(&source_bytes));
        Self {
            package: source.clone(),
            patch: Patch {
                artifacts: OwnedExactArtifacts::new(owner.clone(), owner),
                position: target.chart.position,
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
        let source_owner = litchi_iwa_archive::package::SharedBytes::from_shared_slice(
            source.state.source.shared_source(),
        );
        let target_owner = litchi_iwa_archive::package::SharedBytes::from_shared_slice(
            package.state.source.shared_source(),
        );
        Self {
            package,
            patch: Patch {
                artifacts: OwnedExactArtifacts::new(source_owner, target_owner),
                position: target.chart.position,
                before: target.text,
                after,
            },
            diagnostics: Diagnostics { changed: true },
        }
    }
}

pub(super) fn read(
    package: &Package,
    selector: BodyChartSelector,
) -> Result<Option<String>, Error> {
    Ok(resolve(package, selector)?.text.map(Into::into))
}

pub(super) fn edit(package: &Package, selector: BodyChartSelector) -> Result<Edit<'_>, Error> {
    Ok(Edit {
        source: package,
        target: resolve(package, selector)?,
        operation: None,
    })
}

pub(super) fn apply_patch(package: &Package, patch: &Patch) -> Result<Commit, Error> {
    let source = litchi_iwa_archive::package::SharedBytes::from_shared_slice(
        package.state.source.shared_source(),
    );
    if !patch.artifacts.authorizes_owner(&source) {
        return Err(Error::PatchConflict);
    }
    let current = resolve(package, BodyChartSelector::position(patch.position))?;
    if current.text != patch.before {
        return Err(Error::PatchConflict);
    }
    if patch.is_noop() {
        return Ok(Commit::unchanged(package, current, patch.after.clone()));
    }
    let candidate = graph::reopen(package, Arc::from(patch.artifacts.target_owner().as_ref()))
        .map_err(map_graph_error)?;
    let verified = resolve(&candidate, BodyChartSelector::position(patch.position))?;
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
    pub fn body_chart_caption(
        &self,
        selector: impl Into<BodyChartSelector>,
    ) -> Result<Option<String>, Error> {
        read(self, selector.into())
    }

    pub fn edit_body_chart_caption(
        &self,
        selector: impl Into<BodyChartSelector>,
    ) -> Result<Edit<'_>, Error> {
        edit(self, selector.into())
    }

    pub fn apply_body_chart_caption(&self, patch: &Patch) -> Result<Commit, Error> {
        apply_patch(self, patch)
    }
}
