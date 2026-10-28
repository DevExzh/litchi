//! Shared Pages body-chart title and axis-title transaction machinery.
//!
//! The rooted body-chart selector and component publication remain Pages
//! responsibilities.  Generated title and axis-title wire work is delegated
//! to the neutral `litchi-iwa-style::chart_text` engines, so Keynote, Numbers
//! and Pages use the same presence-preserving codec contract.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "The focused chart-text transaction keeps its bounded phases together."
)]

use std::fmt;
use std::num::NonZeroU64;
use std::sync::Arc;

use litchi_core::Position;
use litchi_iwa_archive::package::SharedBytes;
use litchi_iwa_common::chart::axis::Axis;
use litchi_iwa_common::wire::patch_length_delimited_field;
use litchi_iwa_protos::chart_style_codec::{self as chart_codec, Archive as ChartArchive};
use litchi_iwa_style::chart_text::{axis_title, title};
use thiserror::Error as ThisError;

use super::Package;
use super::body_chart_arrangement::{self as arrangement, ArrangementBudget, ChartTarget};
use super::body_drawable_graph::{self as graph, Budget as GraphBudget, ObjectIndex};
use crate::selector::BodyChartSelector;

const GENERATED_EXTENSION_FIELD: u32 = 10_000;

const fn codec_axis(axis: Axis) -> axis_title::AxisTitleKind {
    match axis {
        Axis::Category => axis_title::AxisTitleKind::Category,
        Axis::Value => axis_title::AxisTitleKind::Value,
    }
}

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

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// One native payload inspected by the selector or codec.
    PayloadBytes,
    /// Generated-extension wire bytes.
    WireBytes,
    /// Candidate package bytes.
    OutputBytes,
    /// Bounded temporary allocations.
    Allocations,
}

impl fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::PayloadBytes => "payload bytes",
            Self::WireBytes => "wire bytes",
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
    #[error("the selected chart has no requested axis")]
    UnsupportedAxis,
    #[error("the selected chart text object is shared by another chart")]
    SharedObject,
    #[error("this Pages source does not support exact chart-text edits")]
    UnsupportedSource,
    #[error("chart text contains a reserved native marker")]
    ReservedMarker,
}

impl Refusal {
    #[must_use]
    pub const fn kind(self) -> litchi_core::edit::RefusalKind {
        use litchi_core::edit::RefusalKind;
        match self {
            Self::ChartNotFound => RefusalKind::NotFound,
            Self::UnsupportedAxis | Self::SharedObject | Self::UnsupportedSource => {
                RefusalKind::UnsupportedContent
            },
            Self::ReservedMarker => RefusalKind::InvalidValue,
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Error {
    #[error(transparent)]
    Refused(Refusal),
    #[error("the selected Pages body-chart text source is invalid")]
    InvalidSource,
    #[error("Pages body-chart text {kind} limit exceeded: observed {observed}, maximum {maximum}")]
    Limit {
        kind: LimitKind,
        observed: u64,
        maximum: u64,
    },
    #[error("could not allocate {amount} Pages body-chart text units")]
    Allocation { amount: usize },
    #[error("the edited Pages body-chart text failed semantic verification")]
    Verification,
    #[error("the Pages body-chart text patch does not match its exact source")]
    PatchConflict,
}

impl From<Refusal> for Error {
    fn from(value: Refusal) -> Self {
        Self::Refused(value)
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct Target {
    chart: ChartTarget,
    kind: Kind,
    component_index: usize,
    component_name: Arc<str>,
    object_index: usize,
    object_identifier: NonZeroU64,
    message_index: usize,
    message_type: u32,
    before: Option<Box<str>>,
}

impl Target {
    fn same_identity(&self, other: &Self) -> bool {
        self.chart.same_identity(&other.chart)
            && self.kind == other.kind
            && self.component_index == other.component_index
            && self.component_name == other.component_name
            && self.object_index == other.object_index
            && self.object_identifier == other.object_identifier
            && self.message_index == other.message_index
            && self.message_type == other.message_type
    }

    fn payload<'a>(&self, package: &'a Package) -> Result<&'a [u8], Error> {
        let component = package
            .state
            .source
            .components()
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
        if object.archive_info.identifier != Some(self.object_identifier.get()) {
            return Err(Error::InvalidSource);
        }
        let message = object
            .messages
            .get(self.message_index)
            .ok_or(Error::InvalidSource)?;
        if message.type_ != self.message_type {
            return Err(Error::InvalidSource);
        }
        Ok(message.data.as_slice())
    }
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
        if data != 0 || objects != 1 || field_info.path.path.as_slice() != [10000, field] {
            return false;
        }
        fields = fields.saturating_add(objects);
    }
    fields == 1
}

fn extension(payload: &[u8]) -> Result<Option<&[u8]>, Error> {
    let view =
        litchi_iwa_common::wire::WireView::parse(payload, litchi_iwa_common::WireLimits::DEFAULT)
            .map_err(|_| Error::InvalidSource)?;
    let mut selected = None;
    for field in view.fields() {
        if field.number() != GENERATED_EXTENSION_FIELD {
            continue;
        }
        if selected.is_some() || field.wire_type() != 2 {
            return Err(Error::InvalidSource);
        }
        field
            .validate_canonical_framing()
            .map_err(|_| Error::InvalidSource)?;
        selected = Some(field.payload());
    }
    Ok(selected)
}

fn read_payload(payload: &[u8], kind: Kind) -> Result<Option<Box<str>>, Error> {
    let Some(extension) = extension(payload)? else {
        return Ok(None);
    };
    match kind {
        Kind::Title => {
            title::decode_chart_title(extension, title::DecodeOptions::for_source(extension))
                .map_err(|_| Error::InvalidSource)
                .map(|snapshot| snapshot.visible_title().map(Into::into))
        },
        Kind::Axis(axis) => axis_title::decode_axis_title(
            extension,
            axis_title::DecodeOptions::for_source(extension),
        )
        .map_err(|_| Error::InvalidSource)
        .map(|snapshot| snapshot.visible_title(codec_axis(axis)).map(Into::into)),
    }
}

fn rewrite_payload(payload: &[u8], kind: Kind, after: Option<&str>) -> Result<Vec<u8>, Error> {
    let old = extension(payload)?;
    let output = old
        .map_or(0, |value| value.len())
        .checked_add(after.map_or(0, str::len))
        .and_then(|value| value.checked_add(64))
        .ok_or(Error::InvalidSource)?
        .max(1);
    let rewritten = match kind {
        Kind::Title => title::rewrite_chart_title(
            old.unwrap_or_default(),
            title::ChartTitleWrite::new(Some(after.is_some()), after),
            title::DecodeOptions::new(
                output,
                output.saturating_mul(8),
                output.saturating_mul(16),
                8,
            )
            .with_max_output_bytes(output)
            .with_max_title_bytes(output),
        )
        .map_err(|_| Error::InvalidSource)?,
        Kind::Axis(axis) => axis_title::rewrite_axis_title(
            old.unwrap_or_default(),
            axis_title::AxisTitleWrite::preserve().with_title(codec_axis(axis), after),
            axis_title::DecodeOptions::new(
                output,
                output.saturating_mul(8),
                output.saturating_mul(16),
                8,
            )
            .with_max_output_bytes(output)
            .with_max_title_bytes(output)
            .with_max_allocations(output)
            .with_max_retained_bytes(output.saturating_mul(2))
            .with_max_scratch_bytes(output),
        )
        .map_err(|_| Error::InvalidSource)?,
    };
    patch_length_delimited_field(
        payload,
        GENERATED_EXTENSION_FIELD,
        old.is_some(),
        Some(rewritten.as_slice()),
    )
    .map_err(|_| Error::InvalidSource)
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

fn resolve(package: &Package, selector: BodyChartSelector, kind: Kind) -> Result<Target, Error> {
    let mut arrangement_budget = ArrangementBudget::new(package).map_err(map_arrangement_error)?;
    let chart = arrangement::resolve_target(package, selector, &mut arrangement_budget)
        .map_err(map_arrangement_error)?;
    let payload =
        arrangement::selected_chart_payload(package, &chart).map_err(map_arrangement_error)?;
    let references = chart_codec::references(payload, chart_codec::Limits::DEFAULT)
        .map_err(|_| Error::InvalidSource)?;
    let identifier = match kind {
        Kind::Title => references.chart_non_style(),
        Kind::Axis(Axis::Category) => references.category_axis_non_styles().first().copied(),
        Kind::Axis(Axis::Value) => references.value_axis_non_styles().first().copied(),
    }
    .filter(|identifier| *identifier != 0)
    .ok_or_else(|| match kind {
        Kind::Title => Error::InvalidSource,
        Kind::Axis(_) => Refusal::UnsupportedAxis.into(),
    })?;
    let all_references: &[u64] = match kind {
        Kind::Title => std::slice::from_ref(&identifier),
        Kind::Axis(Axis::Category) => references.category_axis_non_styles(),
        Kind::Axis(Axis::Value) => references.value_axis_non_styles(),
    };
    if all_references
        .iter()
        .filter(|candidate| **candidate == identifier)
        .count()
        != 1
    {
        return Err(Refusal::SharedObject.into());
    }
    let drawable_component = package
        .state
        .source
        .components()
        .get_index(chart.component_index)
        .ok_or(Error::InvalidSource)?;
    let drawable_object = drawable_component
        .archive()
        .objects
        .get(chart.drawable_object_index)
        .ok_or(Error::InvalidSource)?;
    let drawable_info = drawable_object
        .archive_info
        .message_infos
        .get(chart.drawable_message_index)
        .ok_or(Error::InvalidSource)?;
    if !owns_reference(drawable_info, identifier, kind.reference_field()) {
        return Err(Error::InvalidSource);
    }

    let mut graph_budget = GraphBudget::new(package).map_err(map_graph_error)?;
    let index = ObjectIndex::new(package, &mut graph_budget).map_err(map_graph_error)?;
    let located = index
        .get(NonZeroU64::new(identifier).ok_or(Error::InvalidSource)?)
        .map_err(map_graph_error)?;
    graph::validate_message_metadata(
        located.object,
        graph::unique_message(located.object, kind.message_type())
            .map_err(map_graph_error)?
            .0,
    )
    .map_err(map_graph_error)?;
    let (message_index, message) =
        graph::unique_message(located.object, kind.message_type()).map_err(map_graph_error)?;
    let component_name = index
        .component_name(located.component_index)
        .map_err(map_graph_error)?;
    let component = package
        .state
        .source
        .components()
        .get_index(located.component_index)
        .ok_or(Error::InvalidSource)?;
    let object_index = component
        .archive()
        .objects
        .iter()
        .position(|object| object.archive_info.identifier == Some(identifier))
        .ok_or(Error::InvalidSource)?;
    let before = read_payload(message.data.as_slice(), kind)?;
    Ok(Target {
        chart,
        kind,
        component_index: located.component_index,
        component_name,
        object_index,
        object_identifier: NonZeroU64::new(identifier).ok_or(Error::InvalidSource)?,
        message_index,
        message_type: kind.message_type(),
        before,
    })
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
            .field("kind", &self.target.kind)
            .field("staged", &self.operation.is_some())
            .finish_non_exhaustive()
    }
}

impl Edit<'_> {
    #[must_use]
    pub fn before(&self) -> Option<&str> {
        self.target.before.as_deref()
    }

    #[must_use]
    pub fn after(&self) -> Option<&str> {
        self.operation
            .as_ref()
            .map(|value| value.as_deref())
            .unwrap_or(self.target.before.as_deref())
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
            .unwrap_or_else(|| self.target.before.clone());
        if after == self.target.before {
            return Ok(Commit::unchanged(self.source, self.target, after));
        }
        if !self.source.state.source.source_is_exact() {
            return Err(Refusal::UnsupportedSource.into());
        }
        let current = resolve(
            self.source,
            BodyChartSelector::position(self.target.chart.position),
            self.target.kind,
        )?;
        if !current.same_identity(&self.target) || current.before != self.target.before {
            return Err(Error::PatchConflict);
        }
        let payload = current.payload(self.source)?;
        let rewritten = rewrite_payload(payload, current.kind, after.as_deref())?;
        if read_payload(&rewritten, current.kind)?.as_deref() != after.as_deref() {
            return Err(Error::Verification);
        }
        let candidate = graph::publish(
            self.source,
            vec![graph::MessageRewrite {
                component: Arc::clone(&current.component_name),
                object: current.object_identifier,
                message_index: current.message_index,
                message_type: current.message_type,
                data: rewritten,
            }],
        )
        .map_err(map_graph_error)?;
        let verified = resolve(
            &candidate,
            BodyChartSelector::position(current.chart.position),
            current.kind,
        )?;
        if !verified.same_identity(&current) || verified.before.as_deref() != after.as_deref() {
            return Err(Error::Verification);
        }
        Ok(Commit::published(self.source, candidate, current, after))
    }
}

#[derive(Clone, PartialEq, Eq)]
pub struct Patch {
    artifacts: litchi_iwa_archive::package::OwnedExactArtifacts,
    position: Position,
    kind: Kind,
    before: Option<Box<str>>,
    after: Option<Box<str>>,
}

impl fmt::Debug for Patch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Patch")
            .field("position", &self.position)
            .field("kind", &self.kind)
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
            kind: self.kind,
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

#[must_use = "a chart-text commit contains the validated package snapshot"]
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
        let package = source;
        let source = package.state.source.shared_source();
        let owner = SharedBytes::from_shared_slice(Arc::clone(&source));
        Self {
            package: package.clone(),
            patch: Patch {
                artifacts: litchi_iwa_archive::package::OwnedExactArtifacts::new(
                    owner.clone(),
                    owner,
                ),
                position: target.chart.position,
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
    ) -> Self {
        let source_bytes = source.state.source.shared_source();
        let target_bytes = package.state.source.shared_source();
        let source_owner = SharedBytes::from_shared_slice(source_bytes);
        let target_owner = SharedBytes::from_shared_slice(target_bytes);
        Self {
            package,
            patch: Patch {
                artifacts: litchi_iwa_archive::package::OwnedExactArtifacts::new(
                    source_owner,
                    target_owner,
                ),
                position: target.chart.position,
                kind: target.kind,
                before: target.before,
                after,
            },
            diagnostics: Diagnostics { changed: true },
        }
    }
}

pub(super) fn read(
    package: &Package,
    selector: BodyChartSelector,
    kind: Kind,
) -> Result<Option<String>, Error> {
    Ok(resolve(package, selector, kind)?.before.map(Into::into))
}

pub(super) fn edit(
    package: &Package,
    selector: BodyChartSelector,
    kind: Kind,
) -> Result<Edit<'_>, Error> {
    Ok(Edit {
        source: package,
        target: resolve(package, selector, kind)?,
        operation: None,
    })
}

pub(super) fn apply_patch(package: &Package, patch: &Patch) -> Result<Commit, Error> {
    let source = package.state.source.shared_source();
    let owner = SharedBytes::from_shared_slice(source);
    if !patch.artifacts.authorizes_owner(&owner) {
        return Err(Error::PatchConflict);
    }
    let current = resolve(
        package,
        BodyChartSelector::position(patch.position),
        patch.kind,
    )?;
    if current.before != patch.before {
        return Err(Error::PatchConflict);
    }
    if patch.is_noop() {
        return Ok(Commit::unchanged(package, current, patch.after.clone()));
    }
    let candidate = graph::reopen(package, Arc::from(patch.artifacts.target_owner().as_ref()))
        .map_err(map_graph_error)?;
    let verified = resolve(
        &candidate,
        BodyChartSelector::position(patch.position),
        patch.kind,
    )?;
    if verified.before != patch.after {
        return Err(Error::Verification);
    }
    Ok(Commit {
        package: candidate,
        patch: patch.clone(),
        diagnostics: Diagnostics { changed: true },
    })
}

impl Package {
    pub fn body_chart_title(
        &self,
        selector: impl Into<BodyChartSelector>,
    ) -> Result<Option<String>, Error> {
        read(self, selector.into(), Kind::Title)
    }

    pub fn edit_body_chart_title(
        &self,
        selector: impl Into<BodyChartSelector>,
    ) -> Result<Edit<'_>, Error> {
        edit(self, selector.into(), Kind::Title)
    }

    pub fn apply_body_chart_title(&self, patch: &Patch) -> Result<Commit, Error> {
        apply_patch(self, patch)
    }

    pub fn body_chart_axis_title(
        &self,
        selector: impl Into<BodyChartSelector>,
        axis: Axis,
    ) -> Result<Option<String>, Error> {
        read(self, selector.into(), Kind::Axis(axis))
    }

    pub fn edit_body_chart_axis_title(
        &self,
        selector: impl Into<BodyChartSelector>,
        axis: Axis,
    ) -> Result<Edit<'_>, Error> {
        edit(self, selector.into(), Kind::Axis(axis))
    }

    pub fn apply_body_chart_axis_title(&self, patch: &Patch) -> Result<Commit, Error> {
        apply_patch(self, patch)
    }
}
