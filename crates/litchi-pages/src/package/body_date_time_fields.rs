//! Selector-first Date & Time smart fields of the rooted Pages body.
//!
//! A Date & Time field is a `TSWP.DateTimeSmartFieldArchive` object the body
//! storage's smart-field table references over a nonempty UTF-16 range of
//! its display text. Fields are selected by body order; the public values
//! carry the range and the archive-free [`Settings`]. Adding a field creates
//! the object above the package watermark; inserting one first inserts its
//! display text; removing one keeps the text (see
//! [`super::body_storage_objects`]).

use std::fmt;
use std::sync::Arc;

use litchi_core::Position;
use litchi_iwa_archive::package::ExactArtifacts;
use litchi_iwa_common::wire::{
    WireView, patch_length_delimited_field, patch_nested_fixed64_field, patch_varint_field,
};
use litchi_iwa_text::date_time::{
    DisplayText, Format, FormatterStyle, Instant, LocaleIdentifier, Settings, UpdatePlan,
};
use litchi_iwa_text::{TextPosition, TextRange};
use litchi_iwa_text_wire::{RewriteBehavior, StorageRewriteExecutionLimits};
use thiserror::Error as ThisError;

use super::Package;
use super::body_drawable_graph::{self as graph, GraphError, GraphLimit, map_wire_error};
use super::body_storage_objects::{self as objects, Storage};
use litchi_iwa_text_wire::ranged;

const DATE_TIME_MESSAGE_TYPE: u32 = 2_034;
const SMART_FIELD_FIELD: u32 = 1;
const SMART_FIELD_UUID_FIELD: u32 = 1;
const FORMAT_FIELD: u32 = 2;
const LOCALE_FIELD: u32 = 3;
const DATE_STYLE_FIELD: u32 = 4;
const TIME_STYLE_FIELD: u32 = 5;
const UPDATE_PLAN_FIELD: u32 = 6;
const NEEDS_UPDATE_FIELD: u32 = 7;
const DATE_FIELD: u32 = 8;
const DATE_SECONDS_FIELD: u32 = 1;

/// A finite resource charged by one body date-time field operation.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Bytes in native payloads inspected.
    InputBytes,
    /// Bytes in the rewritten payload or package artifact.
    OutputBytes,
    /// Parsed wire fields.
    WireFields,
    /// Wire nesting depth.
    WireNesting,
    /// Aggregate wire traversal and reassembly work.
    WireWork,
    /// Semantic section or body-storage cardinality.
    PayloadItems,
    /// Semantic UTF-8 text retained by the reopened document.
    TextBytes,
}

impl fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::InputBytes => "input bytes",
            Self::OutputBytes => "output bytes",
            Self::WireFields => "wire fields",
            Self::WireNesting => "wire nesting",
            Self::WireWork => "wire work",
            Self::PayloadItems => "payload items",
            Self::TextBytes => "text bytes",
        })
    }
}

/// Failure from a body date-time field read or transaction.
#[derive(Debug, Clone, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// A well-formed request this route refuses; [`Refusal::kind`] classifies
    /// it (record 0351 §2.4).
    #[error(transparent)]
    Refused(Refusal),
    /// The body storage or a date-time field object is malformed.
    #[error("the Pages body date-time field source is invalid")]
    InvalidSource,
    /// A finite operation resource ceiling was exceeded.
    #[error(
        "Pages body date-time field {kind} limit exceeded: observed {observed}, maximum {maximum}"
    )]
    Limit {
        /// Resource category.
        kind: LimitKind,
        /// Observed or requested amount.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
    },
    /// A bounded temporary allocation failed.
    #[error("could not allocate {amount} units for Pages body date-time fields")]
    Allocation {
        /// Requested units.
        amount: usize,
    },
    /// Reopening the candidate did not reproduce the requested date-time fields.
    #[error("the edited Pages body date-time fields failed semantic verification")]
    Verification,
    /// The patch was produced from another exact package artifact.
    #[error("the Pages body date-time field patch does not match the exact source package")]
    PatchConflict,
}

/// A well-formed request this family refuses, classified by [`Refusal::kind`].
#[derive(Debug, Clone, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Refusal {
    /// No date-time field matches the selector.
    #[error("no Pages body date-time field matches the selector")]
    NotFound,
    /// The range is outside the body text or splits a character.
    #[error("the date-time field range is outside the body text or splits a character")]
    InvalidRange,
    /// The range overlaps another date-time field.
    #[error("the date-time field range overlaps another date-time field")]
    Overlap,
    /// The date-time field object is referenced by more than the body storage.
    #[error("the date-time field object is shared outside the body storage")]
    SharedObject,
    /// A second operation was staged on one edit.
    #[error("one body date-time field edit stages one operation")]
    OperationAlreadyStaged,
    /// The source does not keep an exact artifact for a changed publication.
    #[error("this Pages source does not support exact body date-time field edits")]
    UnsupportedSource,
}

impl Refusal {
    /// The shared refusal kind litchi-ops classifies this reason by.
    #[must_use]
    pub const fn kind(&self) -> litchi_core::edit::RefusalKind {
        use litchi_core::edit::RefusalKind;
        match self {
            Self::NotFound => RefusalKind::NotFound,
            Self::InvalidRange | Self::Overlap => RefusalKind::InvalidValue,
            Self::SharedObject | Self::UnsupportedSource => RefusalKind::UnsupportedContent,
            Self::OperationAlreadyStaged => RefusalKind::UnsupportedOperation,
        }
    }
}

impl From<Refusal> for Error {
    fn from(refusal: Refusal) -> Self {
        Self::Refused(refusal)
    }
}

impl From<GraphError> for Error {
    fn from(error: GraphError) -> Self {
        match error {
            GraphError::InvalidSource => Self::InvalidSource,
            GraphError::Allocation { amount } => Self::Allocation { amount },
            GraphError::Limit {
                kind,
                observed,
                maximum,
            } => Self::Limit {
                kind: match kind {
                    GraphLimit::InputBytes => LimitKind::InputBytes,
                    GraphLimit::OutputBytes => LimitKind::OutputBytes,
                    GraphLimit::WireFields => LimitKind::WireFields,
                    GraphLimit::WireNesting => LimitKind::WireNesting,
                    GraphLimit::WireWork => LimitKind::WireWork,
                    GraphLimit::SemanticItems => LimitKind::PayloadItems,
                    GraphLimit::SemanticTextBytes => LimitKind::TextBytes,
                },
                observed,
                maximum,
            },
        }
    }
}

/// One Date & Time field of the body.
#[derive(Debug, Clone, PartialEq)]
pub struct Field {
    range: TextRange,
    settings: Settings,
}

impl Field {
    /// The UTF-16 range of the field's display text in the body.
    #[must_use]
    pub const fn range(&self) -> TextRange {
        self.range
    }

    /// The field's formatter and update settings.
    #[must_use]
    pub const fn settings(&self) -> &Settings {
        &self.settings
    }
}

#[derive(Debug, Clone, PartialEq)]
enum Operation {
    Add(TextRange, Settings),
    Insert(TextPosition, DisplayText, Settings),
    Update(usize, TextRange, Settings),
    Remove(usize),
}

/// One Date & Time field operation staged against an immutable package.
pub struct Edit<'a> {
    source: &'a Package,
    before: Arc<[Field]>,
    operation: Option<Operation>,
}

impl fmt::Debug for Edit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("operation", &self.operation)
            .finish_non_exhaustive()
    }
}

impl Edit<'_> {
    /// The fields when the edit began.
    #[must_use]
    pub fn fields(&self) -> &[Field] {
        &self.before
    }

    fn stage(&mut self, operation: Operation) -> Result<&mut Self, Error> {
        if self.operation.is_some() {
            return Err(Refusal::OperationAlreadyStaged.into());
        }
        self.operation = Some(operation);
        Ok(self)
    }

    fn select(&self, position: Position) -> Result<usize, Error> {
        if position.get() < self.before.len() {
            Ok(position.get())
        } else {
            Err(Refusal::NotFound.into())
        }
    }

    /// Stage a field over existing body text.
    ///
    /// # Errors
    ///
    /// Refuses a second operation; range checks run at commit.
    pub fn add(&mut self, range: TextRange, settings: Settings) -> Result<&mut Self, Error> {
        self.stage(Operation::Add(range, settings))
    }

    /// Stage insertion of display text at a body position and a field over it.
    ///
    /// # Errors
    ///
    /// Refuses a second operation.
    pub fn insert(
        &mut self,
        position: TextPosition,
        display_text: DisplayText,
        settings: Settings,
    ) -> Result<&mut Self, Error> {
        self.stage(Operation::Insert(position, display_text, settings))
    }

    /// Stage a new range and settings for the field at `position` (body
    /// order).
    ///
    /// # Errors
    ///
    /// Refuses a missing field and a second operation.
    pub fn update(
        &mut self,
        position: impl Into<Position>,
        range: TextRange,
        settings: Settings,
    ) -> Result<&mut Self, Error> {
        let index = self.select(position.into())?;
        self.stage(Operation::Update(index, range, settings))
    }

    /// Stage removal of the field at `position`; its display text stays.
    ///
    /// # Errors
    ///
    /// Refuses a missing field and a second operation.
    pub fn remove(&mut self, position: impl Into<Position>) -> Result<&mut Self, Error> {
        let index = self.select(position.into())?;
        self.stage(Operation::Remove(index))
    }

    /// Validate and publish the staged operation.
    pub fn commit(self) -> Result<Commit, Error> {
        commit_edit(self)
    }
}

/// Exact-source reversible body date-time field patch.
#[derive(Clone, PartialEq)]
pub struct Patch {
    artifacts: ExactArtifacts,
    before: Arc<[Field]>,
    after: Arc<[Field]>,
}

impl fmt::Debug for Patch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Patch")
            .field("before", &self.before.len())
            .field("after", &self.after.len())
            .finish_non_exhaustive()
    }
}

impl Patch {
    /// The fields required before this patch applies.
    #[must_use]
    pub fn before(&self) -> &[Field] {
        &self.before
    }

    /// The fields this patch produces.
    #[must_use]
    pub fn after(&self) -> &[Field] {
        &self.after
    }

    /// Whether both the fields and the exact artifact are unchanged.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before == self.after && self.artifacts.is_byte_noop()
    }

    /// The exact target-to-source inverse.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            artifacts: self.artifacts.inverse(),
            before: Arc::clone(&self.after),
            after: Arc::clone(&self.before),
        }
    }
}

/// Content-free diagnostics of one body date-time field publication.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Diagnostics {
    changed: bool,
}

impl Diagnostics {
    /// Whether package bytes and fields changed.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }
}

/// The validated immutable result of one body date-time field transaction.
#[must_use = "a body date-time field commit contains the validated package snapshot"]
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

    /// Consume the commit and return its package and patch.
    #[must_use]
    pub fn into_parts(self) -> (Package, Patch) {
        (self.package, self.patch)
    }

    /// Borrow the exact-source reversible patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Borrow the publication diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> &Diagnostics {
        &self.diagnostics
    }
}

impl Package {
    /// Read the body's Date & Time fields in body order.
    ///
    /// # Errors
    ///
    /// Returns [`Error`] when the body storage or a field is malformed.
    pub fn body_date_time_fields(&self) -> Result<Vec<Field>, Error> {
        let (storage, _) = objects::body_storage(self)?;
        Ok(read(self, &storage)?
            .into_iter()
            .map(|(_, field)| field)
            .collect())
    }

    /// Begin one Date & Time field add, insert, update or removal.
    ///
    /// # Errors
    ///
    /// Returns [`Error`] when the body storage or a field is malformed.
    pub fn edit_body_date_time_fields(&self) -> Result<Edit<'_>, Error> {
        Ok(Edit {
            source: self,
            before: self.body_date_time_fields()?.into(),
            operation: None,
        })
    }

    /// Replay an exact-source Date & Time field patch.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when the patch was made against
    /// another artifact or the fields no longer match its `before` state.
    pub fn apply_body_date_time_fields(&self, patch: &Patch) -> Result<Commit, Error> {
        if !patch
            .artifacts
            .authorizes_source(&self.state.source.shared_source())
        {
            return Err(Error::PatchConflict);
        }
        if *self.body_date_time_fields()? != *patch.before {
            return Err(Error::PatchConflict);
        }
        if patch.is_noop() {
            return Ok(Commit {
                package: self.clone(),
                patch: patch.clone(),
                diagnostics: Diagnostics { changed: false },
            });
        }
        let candidate = graph::reopen(self, patch.artifacts.target())?;
        if *candidate.body_date_time_fields()? != *patch.after {
            return Err(Error::Verification);
        }
        Ok(Commit {
            package: candidate,
            patch: patch.clone(),
            diagnostics: Diagnostics { changed: true },
        })
    }
}

fn read(package: &Package, storage: &Storage) -> Result<Vec<(u64, Field)>, Error> {
    let entries = ranged::entries(&storage.payload, ranged::Table::SmartField, wire_limits())
        .map_err(map_ranged_error)?;
    let archive = package
        .state
        .source
        .components()
        .get(storage.component.as_ref())
        .ok_or(Error::InvalidSource)?
        .archive();
    read_from(archive, &entries)
}

fn read_from(
    archive: &litchi_iwa_core::Archive,
    entries: &[ranged::Entry],
) -> Result<Vec<(u64, Field)>, Error> {
    let mut fields = Vec::new();
    let mut seen = std::collections::HashSet::new();
    for entry in entries {
        let object = entry.object();
        let found = archive.object(object).ok_or(Error::InvalidSource)?;
        let Some(settings) = field_settings(found)? else {
            continue;
        };
        if !seen.insert(object) {
            return Err(Error::InvalidSource);
        }
        fields.push((
            object,
            Field {
                range: entry.range(),
                settings,
            },
        ));
    }
    Ok(fields)
}

fn varint(view: &WireView<'_>, number: u32) -> Result<Option<u64>, Error> {
    graph::unique_field(view, number, 0)?
        .map(|field| graph::canonical_varint(field.payload()))
        .transpose()
        .map_err(Error::from)
}

#[allow(clippy::cast_possible_truncation, reason = "native enums are 32-bit")]
fn enum_value(value: Option<u64>) -> Option<i32> {
    value.map(|value| value as i64 as i32)
}

fn field_settings(object: &litchi_iwa_core::ArchiveObject) -> Result<Option<Settings>, Error> {
    let payloads = object
        .messages
        .iter()
        .filter(|message| message.type_ == DATE_TIME_MESSAGE_TYPE)
        .collect::<Vec<_>>();
    let message = match payloads.as_slice() {
        [] => return Ok(None),
        [message] if object.messages.len() == 1 => message,
        _ => return Err(Error::InvalidSource),
    };
    let limits = litchi_iwa_common::WireLimits::default();
    let view = WireView::parse(&message.data, limits).map_err(map_wire_error)?;
    let smart = graph::unique_field(&view, SMART_FIELD_FIELD, 2)?.ok_or(Error::InvalidSource)?;
    let smart = WireView::parse(smart.payload(), limits).map_err(map_wire_error)?;
    let uuid =
        graph::unique_field(&smart, SMART_FIELD_UUID_FIELD, 2)?.ok_or(Error::InvalidSource)?;
    let uuid = std::str::from_utf8(uuid.payload()).map_err(|_| Error::InvalidSource)?;
    if !objects::valid_text_attribute_uuid(uuid) {
        return Err(Error::InvalidSource);
    }
    let text = |number: u32| -> Result<Option<&str>, Error> {
        graph::unique_field(&view, number, 2)?
            .map(|field| std::str::from_utf8(field.payload()).map_err(|_| Error::InvalidSource))
            .transpose()
    };
    let format = text(FORMAT_FIELD)?
        .map(Format::new)
        .transpose()
        .map_err(|_| Error::InvalidSource)?;
    let locale = text(LOCALE_FIELD)?
        .map(LocaleIdentifier::new)
        .transpose()
        .map_err(|_| Error::InvalidSource)?;
    let instant = match graph::unique_field(&view, DATE_FIELD, 2)? {
        Some(date) => {
            let date = WireView::parse(date.payload(), limits).map_err(map_wire_error)?;
            let seconds =
                graph::unique_field(&date, DATE_SECONDS_FIELD, 1)?.ok_or(Error::InvalidSource)?;
            let seconds = f64::from_le_bytes(
                seconds
                    .payload()
                    .try_into()
                    .map_err(|_| Error::InvalidSource)?,
            );
            Some(Instant::from_reference_date_seconds(seconds).map_err(|_| Error::InvalidSource)?)
        },
        None => None,
    };
    let needs_update = varint(&view, NEEDS_UPDATE_FIELD)?.map(|value| value != 0);
    let settings = Settings::new(
        format,
        locale,
        enum_value(varint(&view, DATE_STYLE_FIELD)?).map(FormatterStyle::from_raw),
        enum_value(varint(&view, TIME_STYLE_FIELD)?).map(FormatterStyle::from_raw),
        enum_value(varint(&view, UPDATE_PLAN_FIELD)?).map(UpdatePlan::from_raw),
        needs_update,
        instant,
    )
    .map_err(|_| Error::InvalidSource)?;
    Ok(Some(settings))
}

const fn i32_varint(value: i32) -> u64 {
    value as i64 as u64
}

/// A `TSP.Date` payload: field 1, a fixed64 double of reference-date seconds.
fn encode_date(instant: Instant) -> Result<Vec<u8>, Error> {
    let mut date = Vec::with_capacity(9);
    date.push(0x09);
    date.extend_from_slice(&instant.reference_date_seconds().to_le_bytes());
    Ok(date)
}

fn encode_field(uuid: &str, settings: &Settings) -> Result<Vec<u8>, Error> {
    use litchi_iwa_common::wire::{append_length_delimited_field, append_varint_field};
    let limits = litchi_iwa_common::WireLimits::default();
    let mut smart = Vec::new();
    append_length_delimited_field(&mut smart, SMART_FIELD_UUID_FIELD, uuid.as_bytes(), limits)?;
    let mut data = Vec::new();
    append_length_delimited_field(&mut data, SMART_FIELD_FIELD, &smart, limits)?;
    if let Some(format) = settings.format() {
        append_length_delimited_field(&mut data, FORMAT_FIELD, format.as_str().as_bytes(), limits)?;
    }
    if let Some(locale) = settings.locale_identifier() {
        append_length_delimited_field(&mut data, LOCALE_FIELD, locale.as_str().as_bytes(), limits)?;
    }
    for (number, value) in [
        (
            DATE_STYLE_FIELD,
            settings.date_style().map(FormatterStyle::as_raw),
        ),
        (
            TIME_STYLE_FIELD,
            settings.time_style().map(FormatterStyle::as_raw),
        ),
        (
            UPDATE_PLAN_FIELD,
            settings.update_plan().map(UpdatePlan::as_raw),
        ),
    ] {
        if let Some(value) = value {
            append_varint_field(&mut data, number, i32_varint(value))?;
        }
    }
    if let Some(needs_update) = settings.needs_update() {
        append_varint_field(&mut data, NEEDS_UPDATE_FIELD, u64::from(needs_update))?;
    }
    if let Some(instant) = settings.instant() {
        append_length_delimited_field(&mut data, DATE_FIELD, &encode_date(instant)?, limits)?;
    }
    Ok(data)
}

fn patch_settings(original: &[u8], settings: &Settings) -> Result<Vec<u8>, Error> {
    let limits = litchi_iwa_common::WireLimits::default();
    let view = WireView::parse(original, limits).map_err(map_wire_error)?;
    let present = |number: u32| view.fields().any(|field| field.number() == number);
    let (format, locale, date_style, time_style, update_plan, needs_update, date) = (
        present(FORMAT_FIELD),
        present(LOCALE_FIELD),
        present(DATE_STYLE_FIELD),
        present(TIME_STYLE_FIELD),
        present(UPDATE_PLAN_FIELD),
        present(NEEDS_UPDATE_FIELD),
        present(DATE_FIELD),
    );
    let mut data = patch_length_delimited_field(
        original,
        FORMAT_FIELD,
        format,
        settings.format().map(|value| value.as_str().as_bytes()),
    )?;
    data = patch_length_delimited_field(
        &data,
        LOCALE_FIELD,
        locale,
        settings
            .locale_identifier()
            .map(|value| value.as_str().as_bytes()),
    )?;
    for (number, present, replacement) in [
        (
            DATE_STYLE_FIELD,
            date_style,
            settings.date_style().map(FormatterStyle::as_raw),
        ),
        (
            TIME_STYLE_FIELD,
            time_style,
            settings.time_style().map(FormatterStyle::as_raw),
        ),
        (
            UPDATE_PLAN_FIELD,
            update_plan,
            settings.update_plan().map(UpdatePlan::as_raw),
        ),
    ] {
        data = patch_varint_field(&data, number, present, replacement.map(i32_varint))?;
    }
    data = patch_varint_field(
        &data,
        NEEDS_UPDATE_FIELD,
        needs_update,
        settings.needs_update().map(u64::from),
    )?;
    data = match (date, settings.instant()) {
        (true, Some(instant)) => patch_nested_fixed64_field(
            &data,
            &[DATE_FIELD, DATE_SECONDS_FIELD],
            true,
            Some(instant.reference_date_seconds().to_bits()),
        )?,
        (false, Some(instant)) => {
            patch_length_delimited_field(&data, DATE_FIELD, false, Some(&encode_date(instant)?))?
        },
        (true, None) => patch_length_delimited_field(&data, DATE_FIELD, true, None)?,
        (false, None) => data,
    };
    Ok(data)
}

impl From<litchi_iwa_common::Error> for Error {
    fn from(error: litchi_iwa_common::Error) -> Self {
        map_wire_error(error).into()
    }
}

fn check_range(storage: &Storage, range: TextRange) -> Result<(u32, u32), Error> {
    let (start, end) = (range.start().utf16_index(), range.end().utf16_index());
    if end > storage.utf16_len() || !objects::scalar_boundaries(&storage.text, [start, end]) {
        return Err(Refusal::InvalidRange.into());
    }
    Ok((start, end))
}

/// Insert `text` at `position` in the storage payload through the shared
/// TSWP rewrite (attachment and attribute tables shift with it).
fn insert_text(
    source: &Package,
    storage: &mut Storage,
    position: u32,
    text: &str,
) -> Result<(), Error> {
    if position > storage.utf16_len() || !objects::scalar_boundaries(&storage.text, [position]) {
        return Err(Refusal::InvalidRange.into());
    }
    let limits = super::storage_rewrite_limits(source.state.source.limits())
        .map_err(|_| Error::InvalidSource)?;
    let start = position as usize;
    let prepared = litchi_iwa_text_wire::prepare_storage_text_rewrite_with_behavior(
        &storage.payload,
        start..start,
        text,
        RewriteBehavior::ReplaceSelection,
        limits,
    )
    .map_err(|_| Error::InvalidSource)?;
    let requirements = prepared.execution_requirements();
    let rewritten = prepared
        .execute(StorageRewriteExecutionLimits {
            max_output_bytes: requirements.output_bytes(),
            max_retained_elements: requirements.retained_elements(),
            max_retained_bytes: requirements.retained_bytes(),
            max_peak_scratch_bytes: requirements.peak_scratch_bytes(),
            max_allocations: requirements.allocations(),
            max_work: requirements.work(),
        })
        .map_err(|_| Error::InvalidSource)?;
    let mut updated = String::with_capacity(storage.text.len() + text.len());
    let byte = storage
        .text
        .char_indices()
        .scan(0u32, |units, (index, character)| {
            let at = *units;
            *units += character.len_utf16() as u32;
            Some((at, index))
        })
        .find(|(units, _)| *units == position)
        .map_or(storage.text.len(), |(_, index)| index);
    updated.push_str(&storage.text[..byte]);
    updated.push_str(text);
    updated.push_str(&storage.text[byte..]);
    storage.payload = rewritten.into_bytes();
    storage.text = updated;
    Ok(())
}
fn commit_edit(edit: Edit<'_>) -> Result<Commit, Error> {
    let source = edit.source;
    let source_bytes = source.state.source.shared_source();
    let unchanged = |before: Arc<[Field]>| Commit {
        package: source.clone(),
        patch: Patch {
            artifacts: ExactArtifacts::new(Arc::clone(&source_bytes), Arc::clone(&source_bytes)),
            before: Arc::clone(&before),
            after: before,
        },
        diagnostics: Diagnostics { changed: false },
    };
    let Some(operation) = edit.operation else {
        return Ok(unchanged(edit.before));
    };
    if let Operation::Update(position, range, settings) = &operation
        && edit.before[*position].range == *range
        && edit.before[*position].settings == *settings
    {
        return Ok(unchanged(edit.before));
    }
    if !source.state.source.source_is_exact() {
        return Err(Refusal::UnsupportedSource.into());
    }
    let (mut storage, _) = objects::body_storage(source)?;
    let current = read(source, &storage)?;
    if current
        .iter()
        .map(|(_, field)| field)
        .ne(edit.before.iter())
    {
        return Err(Error::PatchConflict);
    }
    let operation = match operation {
        Operation::Insert(position, display, settings) => {
            insert_text(
                source,
                &mut storage,
                position.utf16_index(),
                display.as_str(),
            )?;
            let start = position.utf16_index() as usize;
            let range = TextRange::from_utf16_indexes(
                start,
                start + display.as_str().encode_utf16().count(),
            )
            .map_err(|_| Refusal::InvalidRange)?;
            Operation::Add(range, settings)
        },
        other => other,
    };
    let metadata = objects::metadata(source)?;
    let archive_limits = source
        .state
        .source
        .limits()
        .effective_archive_limits()
        .map_err(graph::map_archive_error)?;
    let mut archive = graph::editable_archive(source, &storage.component)?;
    let mut metadata_payload = metadata.as_ref().map(|metadata| metadata.payload.clone());
    let mut expected = read(source, &storage)?
        .into_iter()
        .map(|(_, field)| field)
        .collect::<Vec<_>>();
    let mut added = None;
    let mut removed = None;
    let rewrite;
    match operation {
        Operation::Add(range, settings) => {
            check_range(&storage, range)?;
            let identifier = objects::next_identifier(source, metadata.as_ref())?;
            rewrite = table_change(
                &storage,
                &ranged::Change::Add {
                    object: identifier,
                    range,
                },
            )?;
            let uuid = objects::text_attribute_uuid(identifier, DATE_TIME_MESSAGE_TYPE)?;
            let object = objects::new_object(
                source,
                identifier,
                DATE_TIME_MESSAGE_TYPE,
                encode_field(&uuid, &settings)?,
            )?;
            archive
                .insert_object(object, archive_limits)
                .map_err(graph::map_core_error)?;
            if let Some(metadata) = &metadata {
                metadata_payload = Some(metadata.raised(source, identifier)?);
            }
            added = Some(identifier);
            let position = expected.partition_point(|field| field.range.start() < range.start());
            expected.insert(position, Field { range, settings });
        },
        Operation::Update(position, range, settings) => {
            let identifier = current[position].0;
            check_range(&storage, range)?;
            require_exclusive(source, storage.identifier.get(), identifier)?;
            rewrite = table_change(
                &storage,
                &ranged::Change::Move {
                    object: identifier,
                    range,
                },
            )?;
            let object = archive.object_mut(identifier).ok_or(Error::InvalidSource)?;
            let original = object.messages[0].data.clone();
            let data = patch_settings(&original, &settings)?;
            object
                .replace_message_preserving_header(
                    0,
                    litchi_iwa_core::RawMessage {
                        type_: DATE_TIME_MESSAGE_TYPE,
                        data,
                    },
                    archive_limits,
                )
                .map_err(graph::map_core_error)?;
            expected[position] = Field { range, settings };
            expected.sort_by_key(|field| field.range.start());
        },
        Operation::Insert(..) => return Err(Error::InvalidSource),
        Operation::Remove(position) => {
            let identifier = current[position].0;
            require_exclusive(source, storage.identifier.get(), identifier)?;
            rewrite = table_change(&storage, &ranged::Change::Remove { object: identifier })?;
            if let Some(metadata) = &metadata {
                let mut payload = metadata.payload.clone();
                if let Some(component) = metadata.component_for_entry(&storage.component)? {
                    payload = objects::remove_external_references(&payload, component, identifier)?;
                }
                if let Some(component) = metadata.uuid_owner(identifier)? {
                    payload = objects::remove_object_uuid(&payload, component, identifier)?;
                }
                metadata_payload = Some(payload);
            }
            archive
                .remove_object(identifier)
                .ok_or(Error::InvalidSource)?;
            removed = Some(identifier);
            expected.remove(position);
        },
    }
    // The storage's field table and its object-reference metadata.
    let payload = rewrite.into_payload();
    let storage_object = archive
        .object_mut(storage.identifier.get())
        .ok_or(Error::InvalidSource)?;
    storage_object
        .replace_message_preserving_header(
            storage.message_index,
            litchi_iwa_core::RawMessage {
                type_: storage.message_type,
                data: payload,
            },
            archive_limits,
        )
        .map_err(graph::map_core_error)?;
    let info = storage_object
        .archive_info
        .message_infos
        .get_mut(storage.message_index)
        .ok_or(Error::InvalidSource)?;
    if let Some(identifier) = added {
        if info.object_references.contains(&identifier) {
            return Err(Error::InvalidSource);
        }
        info.object_references.push(identifier);
    }
    if let Some(identifier) = removed {
        if info
            .object_references
            .iter()
            .filter(|candidate| **candidate == identifier)
            .count()
            != 1
        {
            return Err(Error::InvalidSource);
        }
        info.object_references
            .retain(|candidate| *candidate != identifier);
        for field in &mut info.field_infos {
            field
                .object_references
                .retain(|candidate| *candidate != identifier);
        }
    }
    let mut archives = vec![(Arc::clone(&storage.component), archive)];
    if let (Some(metadata), Some(mut payload)) = (&metadata, metadata_payload) {
        if let Some(identifier) = removed {
            let remaining = objects::max_identifier(source, &archives);
            payload = objects::released(&payload, metadata.last, &[identifier], remaining)?;
        }
        if payload != metadata.payload {
            let mut metadata_archive = graph::editable_archive(source, &metadata.component)?;
            metadata_archive
                .object_mut(metadata.object)
                .ok_or(Error::InvalidSource)?
                .replace_message_preserving_header(
                    metadata.message_index,
                    litchi_iwa_core::RawMessage {
                        type_: 11_006,
                        data: payload,
                    },
                    archive_limits,
                )
                .map_err(graph::map_core_error)?;
            archives.push((Arc::clone(&metadata.component), metadata_archive));
        }
    }
    let candidate = graph::publish_archives(source, archives)?;
    let after = candidate.body_date_time_fields()?;
    if after != expected {
        return Err(Error::Verification);
    }
    let target_bytes = candidate.state.source.shared_source();
    Ok(Commit {
        package: candidate,
        patch: Patch {
            artifacts: ExactArtifacts::new(source_bytes, target_bytes),
            before: edit.before,
            after: after.into(),
        },
        diagnostics: Diagnostics { changed: true },
    })
}

/// Refuse a field object that another object's metadata references.
fn require_exclusive(package: &Package, storage: u64, object: u64) -> Result<(), Error> {
    let owners = package
        .state
        .source
        .components()
        .iter()
        .flat_map(|component| component.archive().objects.iter())
        .filter(|candidate| {
            candidate
                .archive_info
                .message_infos
                .iter()
                .any(|info| info.object_references.contains(&object))
        })
        .filter_map(|candidate| candidate.archive_info.identifier)
        .collect::<Vec<_>>();
    if owners != [storage] {
        return Err(Refusal::SharedObject.into());
    }
    Ok(())
}

const fn wire_limits() -> litchi_iwa_common::WireLimits {
    litchi_iwa_common::WireLimits::DEFAULT
}

/// One change of the storage's smart-field table through the shared engine.
fn table_change(storage: &Storage, change: &ranged::Change) -> Result<ranged::Rewrite, Error> {
    ranged::rewrite(
        &storage.payload,
        ranged::Table::SmartField,
        change,
        wire_limits(),
    )
    .map_err(map_ranged_error)
}

fn map_ranged_error(error: ranged::Error) -> Error {
    match error {
        ranged::Error::Refused(ranged::Refusal::NotABoundary) => Refusal::InvalidRange.into(),
        ranged::Error::Refused(ranged::Refusal::Occupied) => Refusal::Overlap.into(),
        ranged::Error::Refused(_) => Refusal::SharedObject.into(),
        ranged::Error::Allocation { amount } => Error::Allocation { amount },
        ranged::Error::Limit {
            observed, maximum, ..
        } => Error::Limit {
            kind: LimitKind::WireWork,
            observed: observed as u64,
            maximum: maximum as u64,
        },
        _ => Error::InvalidSource,
    }
}
