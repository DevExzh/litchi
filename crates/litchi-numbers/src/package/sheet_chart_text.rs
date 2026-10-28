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
use litchi_iwa_common::chart::axis::Axis;
use litchi_iwa_common::wire::patch_length_delimited_field;
use litchi_iwa_core::{ArchiveObject, RawMessage, SnappyStream};
use litchi_iwa_protos::chart_style_codec::{self as chart_codec, Archive as ChartArchive};
use litchi_iwa_style::chart_text::{axis_title, title};
use thiserror::Error as ThisError;

use super::Package;
use super::chart_arrangement::{
    self, ChartBudget, ChartSelection, object_from_resolved, select_chart_with_budget,
    selected_chart_payload, unique_typed_message, validate_message_metadata,
};
use crate::selector::{ChartSelector, SheetSelector};

const GENERATED_EXTENSION_FIELD: u32 = 10_000;

const fn codec_axis(axis: Axis) -> axis_title::AxisTitleKind {
    match axis {
        Axis::Category => axis_title::AxisTitleKind::Category,
        Axis::Value => axis_title::AxisTitleKind::Value,
    }
}

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
    /// Strict generated-extension bytes.
    WireBytes,
    /// Rewritten output bytes.
    OutputBytes,
    /// Retained title text bytes.
    TextBytes,
    /// Bounded temporary allocations.
    Allocations,
}

impl fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::PackageBytes => "package bytes",
            Self::PayloadBytes => "payload bytes",
            Self::WireBytes => "wire bytes",
            Self::OutputBytes => "output bytes",
            Self::TextBytes => "text bytes",
            Self::Allocations => "allocations",
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
    let references = chart_codec::references(payload, chart_codec::Limits::DEFAULT)
        .map_err(|_| Error::InvalidSource)?;
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
    let all_references = match kind {
        Kind::Title => vec![references.chart_non_style().unwrap_or_default()],
        Kind::Axis(Axis::Category) => references.category_axis_non_styles().to_vec(),
        Kind::Axis(Axis::Value) => references.value_axis_non_styles().to_vec(),
    };
    if all_references
        .iter()
        .filter(|candidate| **candidate == identifier)
        .count()
        != 1
    {
        return Err(Refusal::SharedObject.into());
    }
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
        .map_err(|_| Error::InvalidSource)?
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
            let snapshot =
                title::decode_chart_title(extension, title::DecodeOptions::for_source(extension))
                    .map_err(|_| Error::InvalidSource)?;
            Ok(snapshot.visible_title().map(str::into))
        },
        Kind::Axis(axis) => {
            let snapshot = axis_title::decode_axis_title(
                extension,
                axis_title::DecodeOptions::for_source(extension),
            )
            .map_err(|_| Error::InvalidSource)?;
            Ok(snapshot.visible_title(codec_axis(axis)).map(str::into))
        },
    }
}

fn rewrite_payload(payload: &[u8], kind: Kind, after: Option<&str>) -> Result<Vec<u8>, Error> {
    let old = extension(payload)?;
    let old_len = old.map_or(0, |value| value.len());
    let output = old_len
        .checked_add(after.map_or(0, str::len))
        .and_then(|value| value.checked_add(64))
        .ok_or(Error::InvalidSource)?
        .max(1);
    let extension = match kind {
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
        Some(extension.as_slice()),
    )
    .map_err(|_| Error::InvalidSource)
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
                chart_arrangement::LimitKind::OutputBytes => LimitKind::OutputBytes,
                chart_arrangement::LimitKind::Allocations => LimitKind::Allocations,
                chart_arrangement::LimitKind::RetainedBytes => LimitKind::PayloadBytes,
                _ => LimitKind::WireBytes,
            },
            observed,
            maximum,
        },
        chart_arrangement::Error::Allocation { amount } => Error::Allocation { amount },
        _ => Error::InvalidSource,
    }
}

fn publish(source: &Package, target: &Target, rewritten: Vec<u8>) -> Result<Package, Error> {
    if !source.state.source.source_is_exact() {
        return Err(Refusal::UnsupportedSource.into());
    }
    let source_catalog = source
        .state
        .components
        .physical()
        .ok_or(Error::InvalidSource)?;
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
        .map_err(|_| Error::InvalidSource)?;
    let snappy_limits = source_catalog
        .limits()
        .snappy_limits()
        .map_err(|_| Error::InvalidSource)?;
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
        .map_err(|_| Error::InvalidSource)?;
    let encoded = archive
        .to_bytes(archive_limits)
        .map_err(|_| Error::InvalidSource)?;
    let compressed =
        SnappyStream::compress(&encoded, snappy_limits).map_err(|_| Error::InvalidSource)?;
    let edits = [EntryEdit::new(target.component_name.as_ref(), &compressed)];
    let prepared = source_catalog
        .prepare_reassembly_with_deletions(&edits, &[], source_catalog.limits())
        .map_err(|_| Error::InvalidSource)?;
    let output = prepared
        .execute(prepared.execution_requirements().exact_limits())
        .map_err(|_| Error::InvalidSource)?;
    Package::from_owned_bytes_with_options(output, source.state.options)
        .map_err(|_| Error::Verification)
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
    let candidate = Package::from_source_owner_with_options(target, package.state.options)
        .map_err(|_| Error::Verification)?;
    let verified = resolve(
        &candidate,
        SheetSelector::position(patch.sheet),
        ChartSelector::index(patch.chart.get()),
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
pub(super) fn edit<'a>(
    package: &'a Package,
    sheet: impl Into<SheetSelector<'a>>,
    chart: impl Into<ChartSelector>,
    kind: Kind,
) -> Result<Edit<'a>, Error> {
    Ok(Edit {
        target: resolve(package, sheet, chart, kind)?,
        source: package,
        operation: None,
    })
}
