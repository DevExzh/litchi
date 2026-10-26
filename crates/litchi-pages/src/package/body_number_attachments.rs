//! Selector-first page-number and page-count attachments of a Pages text
//! storage: the body, or one header/footer slot.
//!
//! A number attachment is a `TSWP.NumberAttachmentArchive` object the
//! storage's attachment table references at a U+FFFC character. Inserting
//! one adds the character and the object above the package watermark;
//! removing one deletes the character, the object and its metadata
//! registrations (see [`super::body_storage_objects`]). Attachments are
//! selected by their order in the storage; no native identifier is public.

use std::fmt;
use std::sync::Arc;

use litchi_core::Position;
use litchi_iwa_archive::package::ExactArtifacts;
use litchi_iwa_common::wire::{
    WireView, patch_length_delimited_field, patch_nested_length_delimited_field,
    patch_nested_varint_field, patch_varint_field, repeated_length_delimited_payloads,
};
use litchi_iwa_text::TextPosition;
use litchi_iwa_text::number_attachment::{
    TextNumberAttachmentFormat, TextNumberAttachmentKind, TextNumberAttachmentSettings,
    TextNumberAttachmentText,
};
use thiserror::Error as ThisError;

use super::Package;
use super::body_drawable_graph::{self as graph, GraphError, GraphLimit, map_wire_error};
use super::body_storage_objects::{self as objects, Storage};
use crate::header_footer::HeaderFooterSelector;

const ATTACHMENT_TABLE_FIELD: u32 = 9;
const TABLE_ENTRIES_FIELD: u32 = 1;
const ENTRY_INDEX_FIELD: u32 = 1;
const ENTRY_OBJECT_FIELD: u32 = 2;
const NUMBER_ATTACHMENT_MESSAGE_TYPE: u32 = 2_043;
const SUPER_FIELD: u32 = 1;
const STRING_EQUIVALENT_FIELD: u32 = 1;
const KIND_FIELD: u32 = 2;
const NUMBER_FORMAT_FIELD: u32 = 2;
const STRING_VALUE_FIELD: u32 = 3;
const NUMBER_FORMAT_NAME_FIELD: u32 = 4;
const OBJECT_REPLACEMENT: &str = "\u{fffc}";

/// A finite resource charged by one body number attachment operation.
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
}

impl fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::InputBytes => "input bytes",
            Self::OutputBytes => "output bytes",
            Self::WireFields => "wire fields",
            Self::WireNesting => "wire nesting",
            Self::WireWork => "wire work",
        })
    }
}

/// Failure from a body number attachment read or transaction.
#[derive(Debug, Clone, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// A well-formed request this route refuses; [`Refusal::kind`] classifies
    /// it (record 0351 §2.4).
    #[error(transparent)]
    Refused(Refusal),
    /// The body storage or a number attachment object is malformed.
    #[error("the Pages body number attachment source is invalid")]
    InvalidSource,
    /// A finite operation resource ceiling was exceeded.
    #[error(
        "Pages body number attachment {kind} limit exceeded: observed {observed}, maximum {maximum}"
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
    #[error("could not allocate {amount} units for Pages body number attachments")]
    Allocation {
        /// Requested units.
        amount: usize,
    },
    /// Reopening the candidate did not reproduce the requested number attachments.
    #[error("the edited Pages body number attachments failed semantic verification")]
    Verification,
    /// The patch was produced from another exact package artifact.
    #[error("the Pages body number attachment patch does not match the exact source package")]
    PatchConflict,
}

/// A well-formed request this family refuses, classified by [`Refusal::kind`].
#[derive(Debug, Clone, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Refusal {
    /// No number attachment matches the selector.
    #[error("no Pages body number attachment matches the selector")]
    NotFound,
    /// The position is outside the text or splits a character.
    #[error("the number attachment position is outside the text or splits a character")]
    InvalidRange,
    /// Another attachment already sits at the position.
    #[error("the text position already holds an attachment")]
    Overlap,
    /// The number attachment object is referenced by more than the body storage.
    #[error("the number attachment object is shared outside the body storage")]
    SharedObject,
    /// A second operation was staged on one edit.
    #[error("one body number attachment edit stages one operation")]
    OperationAlreadyStaged,
    /// The source does not keep an exact artifact for a changed publication.
    #[error("this Pages source does not support exact body number attachment edits")]
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
                },
                observed,
                maximum,
            },
        }
    }
}

/// The text storage whose number attachments an operation reads or edits.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum Target<'a> {
    /// The rooted body storage.
    Body,
    /// One header or footer slot.
    HeaderFooter(HeaderFooterSelector<'a>),
}

/// One number attachment of a storage.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Attachment {
    position: TextPosition,
    settings: TextNumberAttachmentSettings,
}

impl Attachment {
    /// The UTF-16 position of the attachment's U+FFFC character.
    #[must_use]
    pub const fn position(&self) -> TextPosition {
        self.position
    }

    /// The attachment's kind, format and cached text.
    #[must_use]
    pub const fn settings(&self) -> &TextNumberAttachmentSettings {
        &self.settings
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
enum Operation {
    Insert(TextPosition, TextNumberAttachmentSettings),
    Update(usize, TextNumberAttachmentSettings),
    Remove(usize),
}

/// One number-attachment operation staged against an immutable package.
pub struct Edit<'a> {
    source: &'a Package,
    storage: u64,
    before: Arc<[Attachment]>,
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
    /// The attachments when the edit began.
    #[must_use]
    pub fn attachments(&self) -> &[Attachment] {
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

    /// Stage a new attachment (and its U+FFFC) at a text position.
    ///
    /// # Errors
    ///
    /// Refuses a second operation; position checks run at commit.
    pub fn insert(
        &mut self,
        position: TextPosition,
        settings: TextNumberAttachmentSettings,
    ) -> Result<&mut Self, Error> {
        self.stage(Operation::Insert(position, settings))
    }

    /// Stage new settings for the attachment at `position` (storage order).
    ///
    /// # Errors
    ///
    /// Refuses a missing attachment and a second operation.
    pub fn update(
        &mut self,
        position: impl Into<Position>,
        settings: TextNumberAttachmentSettings,
    ) -> Result<&mut Self, Error> {
        let index = self.select(position.into())?;
        self.stage(Operation::Update(index, settings))
    }

    /// Stage removal of the attachment at `position` and its U+FFFC.
    ///
    /// # Errors
    ///
    /// Refuses a missing attachment and a second operation.
    pub fn remove(&mut self, position: impl Into<Position>) -> Result<&mut Self, Error> {
        let index = self.select(position.into())?;
        self.stage(Operation::Remove(index))
    }

    /// Validate and publish the staged operation.
    pub fn commit(self) -> Result<Commit, Error> {
        commit_edit(self)
    }
}

/// Exact-source reversible body number attachment patch.
#[derive(Clone, PartialEq)]
pub struct Patch {
    artifacts: ExactArtifacts,
    storage: u64,
    before: Arc<[Attachment]>,
    after: Arc<[Attachment]>,
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
    /// The attachments required before this patch applies.
    #[must_use]
    pub fn before(&self) -> &[Attachment] {
        &self.before
    }

    /// The attachments this patch produces.
    #[must_use]
    pub fn after(&self) -> &[Attachment] {
        &self.after
    }

    /// Whether both the attachments and the exact artifact are unchanged.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before == self.after && self.artifacts.is_byte_noop()
    }

    /// The exact target-to-source inverse.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            artifacts: self.artifacts.inverse(),
            storage: self.storage,
            before: Arc::clone(&self.after),
            after: Arc::clone(&self.before),
        }
    }
}

/// Content-free diagnostics of one body number attachment publication.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Diagnostics {
    changed: bool,
}

impl Diagnostics {
    /// Whether package bytes and attachments changed.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }
}

/// The validated immutable result of one body number attachment transaction.
#[must_use = "a body number attachment commit contains the validated package snapshot"]
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
    /// Read the number attachments of a storage in text order.
    ///
    /// # Errors
    ///
    /// Returns [`Error`] when the target or an attachment is malformed.
    pub fn number_attachments(&self, target: Target<'_>) -> Result<Vec<Attachment>, Error> {
        let storage = storage(self, storage_identifier(self, target)?)?;
        Ok(read(self, &storage)?
            .into_iter()
            .map(|(_, attachment)| attachment)
            .collect())
    }

    /// Begin one number-attachment insert, update or removal.
    ///
    /// # Errors
    ///
    /// Returns [`Error`] when the target or an attachment is malformed.
    pub fn edit_number_attachments(&self, target: Target<'_>) -> Result<Edit<'_>, Error> {
        let identifier = storage_identifier(self, target)?;
        let storage = storage(self, identifier)?;
        Ok(Edit {
            source: self,
            storage: identifier,
            before: read(self, &storage)?
                .into_iter()
                .map(|(_, attachment)| attachment)
                .collect::<Vec<_>>()
                .into(),
            operation: None,
        })
    }

    /// Replay an exact-source number-attachment patch.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when the patch was made against
    /// another artifact or the attachments no longer match its `before`
    /// state.
    pub fn apply_number_attachments(&self, patch: &Patch) -> Result<Commit, Error> {
        if !patch
            .artifacts
            .authorizes_source(&self.state.source.shared_source())
        {
            return Err(Error::PatchConflict);
        }
        let attachments = |package: &Package| -> Result<Vec<Attachment>, Error> {
            Ok(read(package, &storage(package, patch.storage)?)?
                .into_iter()
                .map(|(_, attachment)| attachment)
                .collect())
        };
        if *attachments(self)? != *patch.before {
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
        if *attachments(&candidate)? != *patch.after {
            return Err(Error::Verification);
        }
        Ok(Commit {
            package: candidate,
            patch: patch.clone(),
            diagnostics: Diagnostics { changed: true },
        })
    }
}

fn storage_identifier(package: &Package, target: Target<'_>) -> Result<u64, Error> {
    match target {
        Target::Body => Ok(objects::body_storage(package)?.0.identifier.get()),
        Target::HeaderFooter(selector) => {
            super::header_footer_text::header_footer_storage(package, selector).map_err(|error| {
                match error {
                    super::header_footer_text::Error::Refused(_) => Refusal::NotFound.into(),
                    _ => Error::InvalidSource,
                }
            })
        },
    }
}

fn storage(package: &Package, identifier: u64) -> Result<Storage, Error> {
    Ok(objects::storage_by_identifier(package, identifier)?)
}

/// Raw attachment-table entries `(index, object, raw)` of a storage payload.
fn entries(payload: &[u8]) -> Result<(Option<Vec<u8>>, Vec<(u32, u64, Vec<u8>)>), Error> {
    let tables = repeated_length_delimited_payloads(payload, ATTACHMENT_TABLE_FIELD)?;
    let table = match tables.as_slice() {
        [] => return Ok((None, Vec::new())),
        [table] => table.to_vec(),
        _ => return Err(Error::InvalidSource),
    };
    let limits = litchi_iwa_common::WireLimits::default();
    let mut entries = Vec::new();
    for raw in repeated_length_delimited_payloads(&table, TABLE_ENTRIES_FIELD)? {
        let view = WireView::parse(raw, limits).map_err(map_wire_error)?;
        let index =
            graph::unique_field(&view, ENTRY_INDEX_FIELD, 0)?.ok_or(Error::InvalidSource)?;
        let index = u32::try_from(graph::canonical_varint(index.payload())?)
            .map_err(|_| Error::InvalidSource)?;
        let Some(object) = graph::unique_field(&view, ENTRY_OBJECT_FIELD, 2)? else {
            continue;
        };
        let object = WireView::parse(object.payload(), limits).map_err(map_wire_error)?;
        let identifier = graph::unique_field(&object, 1, 0)?.ok_or(Error::InvalidSource)?;
        entries.push((
            index,
            graph::canonical_varint(identifier.payload())?,
            raw.to_vec(),
        ));
    }
    Ok((Some(table), entries))
}

fn read(package: &Package, storage: &Storage) -> Result<Vec<(u64, Attachment)>, Error> {
    let (_, entries) = entries(&storage.payload)?;
    let archive = package
        .state
        .source
        .components()
        .get(storage.component.as_ref())
        .ok_or(Error::InvalidSource)?
        .archive();
    let mut attachments = Vec::new();
    let mut seen = std::collections::HashSet::new();
    for (index, object, _) in entries {
        let Some(found) = archive.object(object) else {
            continue;
        };
        let Some(settings) = attachment_settings(found)? else {
            continue;
        };
        if !seen.insert(object) {
            return Err(Error::InvalidSource);
        }
        attachments.push((
            object,
            Attachment {
                position: TextPosition::from_utf16_code_units(index),
                settings,
            },
        ));
    }
    attachments.sort_by_key(|(_, attachment)| attachment.position);
    Ok(attachments)
}

fn attachment_settings(
    object: &litchi_iwa_core::ArchiveObject,
) -> Result<Option<TextNumberAttachmentSettings>, Error> {
    let payloads = object
        .messages
        .iter()
        .filter(|message| message.type_ == NUMBER_ATTACHMENT_MESSAGE_TYPE)
        .collect::<Vec<_>>();
    let message = match payloads.as_slice() {
        [] => return Ok(None),
        [message] if object.messages.len() == 1 => message,
        _ => return Err(Error::InvalidSource),
    };
    let limits = litchi_iwa_common::WireLimits::default();
    let view = WireView::parse(&message.data, limits).map_err(map_wire_error)?;
    let textual = graph::unique_field(&view, SUPER_FIELD, 2)?.ok_or(Error::InvalidSource)?;
    let textual = WireView::parse(textual.payload(), limits).map_err(map_wire_error)?;
    let kind = graph::unique_field(&textual, KIND_FIELD, 0)?.ok_or(Error::InvalidSource)?;
    #[allow(clippy::cast_possible_truncation, reason = "native enum is 32-bit")]
    let kind = graph::canonical_varint(kind.payload())? as i64 as i32;
    let text =
        |view: &WireView<'_>, number: u32| -> Result<Option<TextNumberAttachmentText>, Error> {
            graph::unique_field(view, number, 2)?
                .map(|field| {
                    let value =
                        std::str::from_utf8(field.payload()).map_err(|_| Error::InvalidSource)?;
                    TextNumberAttachmentText::new(value).map_err(|_| Error::InvalidSource)
                })
                .transpose()
        };
    let number_format = graph::unique_field(&view, NUMBER_FORMAT_FIELD, 0)?
        .map(|field| graph::canonical_varint(field.payload()))
        .transpose()?
        .map(|value| u32::try_from(value).map_err(|_| Error::InvalidSource))
        .transpose()?
        .map(TextNumberAttachmentFormat::from_native_value);
    Ok(Some(TextNumberAttachmentSettings {
        kind: TextNumberAttachmentKind::from_raw(kind),
        string_equivalent: text(&textual, STRING_EQUIVALENT_FIELD)?,
        number_format,
        string_value: text(&view, STRING_VALUE_FIELD)?,
        number_format_name: text(&view, NUMBER_FORMAT_NAME_FIELD)?,
    }))
}

const fn i32_varint(value: i32) -> u64 {
    value as i64 as u64
}

fn encode_attachment(settings: &TextNumberAttachmentSettings) -> Result<Vec<u8>, Error> {
    use litchi_iwa_common::wire::{append_length_delimited_field, append_varint_field};
    let limits = litchi_iwa_common::WireLimits::default();
    let mut textual = Vec::new();
    if let Some(value) = &settings.string_equivalent {
        append_length_delimited_field(
            &mut textual,
            STRING_EQUIVALENT_FIELD,
            value.as_str().as_bytes(),
            limits,
        )?;
    }
    append_varint_field(&mut textual, KIND_FIELD, i32_varint(settings.kind.as_raw()))?;
    let mut data = Vec::new();
    append_length_delimited_field(&mut data, SUPER_FIELD, &textual, limits)?;
    if let Some(format) = settings.number_format {
        append_varint_field(
            &mut data,
            NUMBER_FORMAT_FIELD,
            u64::from(format.native_value()),
        )?;
    }
    if let Some(value) = &settings.string_value {
        append_length_delimited_field(
            &mut data,
            STRING_VALUE_FIELD,
            value.as_str().as_bytes(),
            limits,
        )?;
    }
    if let Some(value) = &settings.number_format_name {
        append_length_delimited_field(
            &mut data,
            NUMBER_FORMAT_NAME_FIELD,
            value.as_str().as_bytes(),
            limits,
        )?;
    }
    Ok(data)
}

fn patch_attachment(
    original: &[u8],
    settings: &TextNumberAttachmentSettings,
) -> Result<Vec<u8>, Error> {
    let current = attachment_settings_from_payload(original)?;
    let mut data = patch_nested_length_delimited_field(
        original,
        &[SUPER_FIELD, STRING_EQUIVALENT_FIELD],
        current.string_equivalent.is_some(),
        settings
            .string_equivalent
            .as_ref()
            .map(|value| value.as_str().as_bytes()),
    )?;
    data = patch_nested_varint_field(
        &data,
        &[SUPER_FIELD, KIND_FIELD],
        true,
        Some(i32_varint(settings.kind.as_raw())),
    )?;
    data = patch_varint_field(
        &data,
        NUMBER_FORMAT_FIELD,
        current.number_format.is_some(),
        settings
            .number_format
            .map(|value| u64::from(value.native_value())),
    )?;
    data = patch_length_delimited_field(
        &data,
        STRING_VALUE_FIELD,
        current.string_value.is_some(),
        settings
            .string_value
            .as_ref()
            .map(|value| value.as_str().as_bytes()),
    )?;
    data = patch_length_delimited_field(
        &data,
        NUMBER_FORMAT_NAME_FIELD,
        current.number_format_name.is_some(),
        settings
            .number_format_name
            .as_ref()
            .map(|value| value.as_str().as_bytes()),
    )?;
    Ok(data)
}

fn attachment_settings_from_payload(payload: &[u8]) -> Result<TextNumberAttachmentSettings, Error> {
    let object = litchi_iwa_core::ArchiveObject::new(
        1,
        vec![litchi_iwa_core::RawMessage {
            type_: NUMBER_ATTACHMENT_MESSAGE_TYPE,
            data: payload.to_vec(),
        }],
        litchi_iwa_core::Limits::default(),
    )
    .map_err(graph::map_core_error)?;
    attachment_settings(&object)?.ok_or(Error::InvalidSource)
}

impl From<litchi_iwa_common::Error> for Error {
    fn from(error: litchi_iwa_common::Error) -> Self {
        map_wire_error(error).into()
    }
}

fn encode_reference(identifier: u64) -> Result<Vec<u8>, Error> {
    let mut reference = Vec::new();
    litchi_iwa_common::wire::append_varint_field(&mut reference, 1, identifier)?;
    Ok(reference)
}

fn commit_edit(edit: Edit<'_>) -> Result<Commit, Error> {
    let source = edit.source;
    let source_bytes = source.state.source.shared_source();
    let unchanged = |before: Arc<[Attachment]>| Commit {
        package: source.clone(),
        patch: Patch {
            artifacts: ExactArtifacts::new(Arc::clone(&source_bytes), Arc::clone(&source_bytes)),
            storage: edit.storage,
            before: Arc::clone(&before),
            after: before,
        },
        diagnostics: Diagnostics { changed: false },
    };
    let Some(operation) = edit.operation.clone() else {
        return Ok(unchanged(edit.before));
    };
    if let Operation::Update(position, settings) = &operation
        && edit.before[*position].settings == *settings
    {
        return Ok(unchanged(edit.before));
    }
    if !source.state.source.source_is_exact() {
        return Err(Refusal::UnsupportedSource.into());
    }
    let mut storage = storage(source, edit.storage)?;
    let current = read(source, &storage)?;
    if current
        .iter()
        .map(|(_, attachment)| attachment)
        .ne(edit.before.iter())
    {
        return Err(Error::PatchConflict);
    }
    let metadata = objects::metadata(source)?;
    let archive_limits = source
        .state
        .source
        .limits()
        .effective_archive_limits()
        .map_err(graph::map_archive_error)?;
    let mut archive = graph::editable_archive(source, &storage.component)?;
    let mut metadata_payload = metadata.as_ref().map(|metadata| metadata.payload.clone());
    let mut removed = None;
    let mut pruned = Vec::new();
    let mut added = None;
    match operation {
        Operation::Insert(position, settings) => {
            let at = position.utf16_index();
            let (_, existing) = entries(&storage.payload)?;
            if existing.iter().any(|(index, _, _)| *index == at) {
                return Err(Refusal::Overlap.into());
            }
            if at > storage.utf16_len() || !objects::scalar_boundaries(&storage.text, [at]) {
                return Err(Refusal::InvalidRange.into());
            }
            objects::splice_text(source, &mut storage, at, at, OBJECT_REPLACEMENT)?;
            let identifier = objects::next_identifier(source, metadata.as_ref())?;
            let (table, mut rows) = entries(&storage.payload)?;
            let mut entry = Vec::new();
            litchi_iwa_common::wire::append_varint_field(
                &mut entry,
                ENTRY_INDEX_FIELD,
                u64::from(at),
            )?;
            litchi_iwa_common::wire::append_length_delimited_field(
                &mut entry,
                ENTRY_OBJECT_FIELD,
                &encode_reference(identifier)?,
                litchi_iwa_common::WireLimits::default(),
            )?;
            rows.push((at, identifier, entry));
            rows.sort_by_key(|(index, _, _)| *index);
            let raw = rows.into_iter().map(|(_, _, raw)| raw).collect::<Vec<_>>();
            let table_payload = match &table {
                Some(table) => litchi_iwa_common::wire::rewrite_repeated_length_delimited_fields(
                    table,
                    TABLE_ENTRIES_FIELD,
                    &raw,
                )?,
                None => {
                    let mut table = Vec::new();
                    for entry in &raw {
                        litchi_iwa_common::wire::append_length_delimited_field(
                            &mut table,
                            TABLE_ENTRIES_FIELD,
                            entry,
                            litchi_iwa_common::WireLimits::default(),
                        )?;
                    }
                    table
                },
            };
            storage.payload = patch_length_delimited_field(
                &storage.payload,
                ATTACHMENT_TABLE_FIELD,
                table.is_some(),
                Some(&table_payload),
            )?;
            let object = objects::new_object(
                source,
                identifier,
                NUMBER_ATTACHMENT_MESSAGE_TYPE,
                encode_attachment(&settings)?,
            )?;
            if let Some(metadata) = &metadata {
                metadata_payload = Some(metadata.raised(source, identifier)?);
            }
            added = Some((identifier, object));
        },
        Operation::Update(position, settings) => {
            let identifier = current[position].0;
            require_exclusive(source, storage.identifier.get(), identifier)?;
            let object = archive.object_mut(identifier).ok_or(Error::InvalidSource)?;
            let data = patch_attachment(&object.messages[0].data, &settings)?;
            object
                .replace_message_preserving_header(
                    0,
                    litchi_iwa_core::RawMessage {
                        type_: NUMBER_ATTACHMENT_MESSAGE_TYPE,
                        data,
                    },
                    archive_limits,
                )
                .map_err(graph::map_core_error)?;
        },
        Operation::Remove(position) => {
            let (identifier, attachment) = &current[position];
            let identifier = *identifier;
            require_exclusive(source, storage.identifier.get(), identifier)?;
            let at = attachment.position.utf16_index();
            let references = objects::splice_text(source, &mut storage, at, at + 1, "")?;
            if references != [identifier] {
                return Err(Refusal::SharedObject.into());
            }
            pruned.push(identifier);
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
            removed = Some(identifier);
        },
    }
    let storage_object = archive
        .object_mut(storage.identifier.get())
        .ok_or(Error::InvalidSource)?;
    let message = litchi_iwa_core::RawMessage {
        type_: storage.message_type,
        data: storage.payload.clone(),
    };
    if pruned.is_empty() {
        storage_object
            .replace_message_preserving_header(storage.message_index, message, archive_limits)
            .map_err(graph::map_core_error)?;
    } else {
        storage_object
            .replace_message_pruning_object_references_preserving_header(
                storage.message_index,
                message,
                &pruned,
                archive_limits,
            )
            .map_err(graph::map_core_error)?;
    }
    if let Some((identifier, object)) = added {
        let info = storage_object
            .archive_info
            .message_infos
            .get_mut(storage.message_index)
            .ok_or(Error::InvalidSource)?;
        if info.object_references.contains(&identifier) {
            return Err(Error::InvalidSource);
        }
        info.object_references.push(identifier);
        archive
            .insert_object(object, archive_limits)
            .map_err(graph::map_core_error)?;
    }
    if let Some(identifier) = removed {
        archive
            .remove_object(identifier)
            .ok_or(Error::InvalidSource)?;
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
    let after = read(&candidate, &self::storage(&candidate, edit.storage)?)?
        .into_iter()
        .map(|(_, attachment)| attachment)
        .collect::<Vec<_>>();
    let expected_count = match edit.operation {
        Some(Operation::Insert(..)) => edit.before.len() + 1,
        Some(Operation::Remove(_)) => edit.before.len() - 1,
        _ => edit.before.len(),
    };
    if after.len() != expected_count {
        return Err(Error::Verification);
    }
    let target_bytes = candidate.state.source.shared_source();
    Ok(Commit {
        package: candidate,
        patch: Patch {
            artifacts: ExactArtifacts::new(source_bytes, target_bytes),
            storage: edit.storage,
            before: edit.before,
            after: after.into(),
        },
        diagnostics: Diagnostics { changed: true },
    })
}

/// Refuse an attachment object that another object's metadata references.
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
