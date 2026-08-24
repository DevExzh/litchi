//! Exact-source transactions for Pages body-table title settings.
//!
//! The rooted body-table graph is resolved by `table_lock`; this module only
//! owns the title payload projection and the immutable transaction around it.
//! The strict title codec validates the selected scalar/reference fields while
//! the original model payload remains the source-preservation authority.

use std::fmt;
use std::sync::Arc;

use litchi_iwa_archive::{SourceCatalog, package::EntryEdit};
use litchi_iwa_common::{
    WireLimits,
    wire::{NestedFieldEdit, NestedFieldReplacement, patch_nested_fields_batched_with_limits},
};
use litchi_iwa_core::RawMessage;
use litchi_iwa_protos::numbers_table_title_codec::{self, DecodeLimit, TableTitleSettingsSnapshot};
use thiserror::Error;

use super::{Package, PackageError, page_layout, table_lock};
use crate::selector::BodyTableSelector;
use crate::table::title::Settings;

const TITLE_VISIBLE_FIELD: u32 = 22;
const TITLE_OUTLINED_FIELD: u32 = 37;
const PREVIEW_ENTRY_NAMES: [&str; 3] = ["preview.jpg", "preview-micro.jpg", "preview-web.jpg"];

/// Finite resource categories enforced by a body-table title transaction.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum BodyTableTitleLimitKind {
    /// Complete source package bytes.
    InputBytes,
    /// Complete candidate package bytes.
    OutputBytes,
    /// Physical package entries.
    Entries,
    /// Bytes in one physical package entry.
    EntryBytes,
    /// Aggregate physical package bytes.
    TotalEntryBytes,
    /// Decoded native payload bytes.
    PayloadBytes,
    /// Aggregate decoded native payload bytes.
    TotalPayloadBytes,
    /// Native payload objects inspected.
    PayloadObjects,
    /// Native payload messages inspected.
    PayloadMessages,
    /// Native framing/metadata items inspected.
    PayloadItems,
    /// Native references inspected.
    PayloadReferences,
    /// Strict title wire bytes.
    WireBytes,
    /// Strict title output bytes.
    WireOutputBytes,
    /// Strict title fields.
    WireFields,
    /// Strict title nesting.
    WireNesting,
    /// Strict title work.
    WireWork,
}

impl fmt::Display for BodyTableTitleLimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::InputBytes => "input bytes",
            Self::OutputBytes => "output bytes",
            Self::Entries => "entries",
            Self::EntryBytes => "entry bytes",
            Self::TotalEntryBytes => "total entry bytes",
            Self::PayloadBytes => "payload bytes",
            Self::TotalPayloadBytes => "total payload bytes",
            Self::PayloadObjects => "payload objects",
            Self::PayloadMessages => "payload messages",
            Self::PayloadItems => "payload items",
            Self::PayloadReferences => "payload references",
            Self::WireBytes => "wire bytes",
            Self::WireOutputBytes => "wire output bytes",
            Self::WireFields => "wire fields",
            Self::WireNesting => "wire nesting",
            Self::WireWork => "wire work",
        })
    }
}

/// Failure from a Pages body-table title read or transaction.
#[derive(Debug, Error, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum BodyTableTitleError {
    /// No rooted body table matched the selector.
    #[error("the Pages body has no table matching the requested selector")]
    TableNotFound,
    /// More than one rooted body table matched a name selector.
    #[error("the Pages body has more than one table with the requested name")]
    AmbiguousTableName,
    /// The selector or rooted native graph was ambiguous.
    #[error("the Pages body-table title selector is ambiguous")]
    AmbiguousSelector,
    /// The selected source cannot be edited while preserving exact bytes.
    #[error("the Pages package source does not support exact body-table title editing")]
    UnsupportedSource,
    /// The rooted native graph or selected title payload is malformed.
    #[error("the selected Pages body-table title source is invalid")]
    InvalidSource,
    /// A finite transaction ceiling was exceeded.
    #[error("Pages body-table title {kind} limit exceeded: observed {observed}, maximum {maximum}")]
    LimitExceeded {
        /// Resource category.
        kind: BodyTableTitleLimitKind,
        /// Observed amount.
        observed: u64,
        /// Maximum configured amount.
        maximum: u64,
    },
    /// A fallible bounded allocation failed.
    #[error("could not allocate {amount} units for the Pages body-table title transaction")]
    Allocation {
        /// Requested units.
        amount: usize,
    },
    /// Candidate reopening did not reproduce the requested semantic state.
    #[error("the edited Pages body-table title failed semantic verification")]
    Verification,
    /// The patch was created from another exact package artifact.
    #[error("the Pages body-table title patch does not match the exact source package")]
    PatchConflict,
}

/// Mutable semantic title settings staged against one immutable package.
pub struct BodyTableTitleEdit<'a> {
    source: &'a Package,
    target: table_lock::BodyTableTarget,
    before: Settings,
    settings: Settings,
}

impl fmt::Debug for BodyTableTitleEdit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("BodyTableTitleEdit")
            .field("before", &self.before)
            .field("settings", &self.settings)
            .finish_non_exhaustive()
    }
}

impl BodyTableTitleEdit<'_> {
    /// Return the settings that would be published.
    #[must_use]
    pub const fn settings(&self) -> Settings {
        self.settings
    }

    /// Replace the complete staged lossless title settings.
    #[must_use]
    pub fn set(mut self, settings: Settings) -> Self {
        self.settings = settings;
        self
    }

    /// Validate and publish the staged settings atomically.
    pub fn commit(self) -> Result<BodyTableTitleCommit, BodyTableTitleError> {
        commit_edit(self)
    }
}

/// Exact-source reversible title patch.
#[derive(Clone, PartialEq, Eq)]
pub struct BodyTableTitlePatch {
    source: Arc<[u8]>,
    target: Arc<[u8]>,
    source_fingerprint: u64,
    target_fingerprint: u64,
    proof: table_lock::BodyTableTarget,
    before: Settings,
    after: Settings,
    source_preview_count: usize,
    target_preview_count: usize,
}

impl fmt::Debug for BodyTableTitlePatch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("BodyTableTitlePatch")
            .field("before", &self.before)
            .field("after", &self.after)
            .finish_non_exhaustive()
    }
}

impl BodyTableTitlePatch {
    /// Return the source semantic settings.
    #[must_use]
    pub const fn before(&self) -> Settings {
        self.before
    }

    /// Return the target semantic settings.
    #[must_use]
    pub const fn after(&self) -> Settings {
        self.after
    }

    /// Return the source fingerprint used for exact conflict detection.
    #[must_use]
    pub const fn source_fingerprint(&self) -> u64 {
        self.source_fingerprint
    }

    /// Return the target fingerprint used for exact conflict detection.
    #[must_use]
    pub const fn target_fingerprint(&self) -> u64 {
        self.target_fingerprint
    }

    /// Return whether the semantic and physical artifacts are unchanged.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before == self.after
            && self.source_fingerprint == self.target_fingerprint
            && (Arc::ptr_eq(&self.source, &self.target) || self.source == self.target)
    }

    /// Return the exact target-to-source inverse.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            source: Arc::clone(&self.target),
            target: Arc::clone(&self.source),
            source_fingerprint: self.target_fingerprint,
            target_fingerprint: self.source_fingerprint,
            proof: self.proof.clone(),
            before: self.after,
            after: self.before,
            source_preview_count: self.target_preview_count,
            target_preview_count: self.source_preview_count,
        }
    }
}

/// Content-free diagnostics from one title publication.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct BodyTableTitleDiagnostics {
    changed: bool,
    touched_components: usize,
    deleted_previews: usize,
    full_reparse_performed: bool,
}

impl BodyTableTitleDiagnostics {
    const fn unchanged() -> Self {
        Self {
            changed: false,
            touched_components: 0,
            deleted_previews: 0,
            full_reparse_performed: false,
        }
    }

    const fn published(deleted_previews: usize) -> Self {
        Self {
            changed: true,
            touched_components: 1,
            deleted_previews,
            full_reparse_performed: true,
        }
    }

    /// Whether exact package bytes changed.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }

    /// Number of rewritten native components.
    #[must_use]
    pub const fn touched_components(self) -> usize {
        self.touched_components
    }

    /// Number of root previews removed by the edit.
    #[must_use]
    pub const fn deleted_previews(self) -> usize {
        self.deleted_previews
    }

    /// Whether the complete candidate was reopened before publication.
    #[must_use]
    pub const fn full_reparse_performed(self) -> bool {
        self.full_reparse_performed
    }
}

/// Fully reopened immutable result of one title transaction.
#[must_use = "a body-table title commit contains the validated package snapshot"]
#[derive(Debug)]
pub struct BodyTableTitleCommit {
    package: Package,
    patch: BodyTableTitlePatch,
    diagnostics: BodyTableTitleDiagnostics,
}

impl BodyTableTitleCommit {
    /// Borrow the validated package.
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.package
    }

    /// Consume the commit and return its package.
    #[must_use]
    pub fn into_package(self) -> Package {
        self.package
    }

    /// Borrow the reversible patch.
    #[must_use]
    pub const fn patch(&self) -> &BodyTableTitlePatch {
        &self.patch
    }

    /// Borrow publication diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> &BodyTableTitleDiagnostics {
        &self.diagnostics
    }
}

impl Package {
    /// Read one rooted body's lossless title visibility and outline settings.
    pub fn body_table_title_settings<'table>(
        &self,
        selector: impl Into<BodyTableSelector<'table>>,
    ) -> Result<Settings, BodyTableTitleError> {
        let target = resolve_target(self, selector.into())?;
        settings_at_target(self, &target)
    }

    /// Start a selector-first immutable body-table title edit.
    pub fn edit_body_table_title<'table>(
        &self,
        selector: impl Into<BodyTableSelector<'table>>,
    ) -> Result<BodyTableTitleEdit<'_>, BodyTableTitleError> {
        let target = resolve_target(self, selector.into())?;
        let before = settings_at_target(self, &target)?;
        Ok(BodyTableTitleEdit {
            source: self,
            target,
            before,
            settings: before,
        })
    }

    /// Apply an exact-source-checked reversible title patch.
    pub fn apply_body_table_title(
        &self,
        patch: &BodyTableTitlePatch,
    ) -> Result<BodyTableTitleCommit, BodyTableTitleError> {
        if page_layout::fingerprint(self.source_bytes()) != patch.source_fingerprint
            || self.source_bytes() != patch.source.as_ref()
        {
            return Err(BodyTableTitleError::PatchConflict);
        }
        if settings_at_target(self, &patch.proof)? != patch.before {
            return Err(BodyTableTitleError::PatchConflict);
        }
        if patch.is_noop() {
            if patch.source_preview_count != patch.target_preview_count {
                return Err(BodyTableTitleError::PatchConflict);
            }
            return Ok(BodyTableTitleCommit {
                package: self.snapshot(),
                patch: patch.clone(),
                diagnostics: BodyTableTitleDiagnostics::unchanged(),
            });
        }
        if !self.state.source.source_is_exact()
            || page_layout::preview_count(self) != patch.source_preview_count
            || page_layout::fingerprint(patch.target.as_ref()) != patch.target_fingerprint
        {
            return Err(BodyTableTitleError::PatchConflict);
        }
        let candidate = reopen_target(self, Arc::clone(&patch.target))?;
        if settings_at_target(&candidate, &patch.proof)? != patch.after
            || page_layout::preview_count(&candidate) != patch.target_preview_count
        {
            return Err(BodyTableTitleError::Verification);
        }
        Ok(BodyTableTitleCommit {
            package: candidate,
            patch: patch.clone(),
            diagnostics: BodyTableTitleDiagnostics::published(
                patch
                    .source_preview_count
                    .saturating_sub(patch.target_preview_count),
            ),
        })
    }
}

fn commit_edit(edit: BodyTableTitleEdit<'_>) -> Result<BodyTableTitleCommit, BodyTableTitleError> {
    let source = edit.source;
    let source_bytes: Arc<[u8]> = source.state.source.shared_source();
    let source_fingerprint = page_layout::fingerprint(source_bytes.as_ref());
    let source_preview_count = page_layout::preview_count(source);
    if edit.before == edit.settings {
        return Ok(BodyTableTitleCommit {
            package: source.snapshot(),
            patch: BodyTableTitlePatch {
                source: Arc::clone(&source_bytes),
                target: source_bytes,
                source_fingerprint,
                target_fingerprint: source_fingerprint,
                proof: edit.target,
                before: edit.before,
                after: edit.settings,
                source_preview_count,
                target_preview_count: source_preview_count,
            },
            diagnostics: BodyTableTitleDiagnostics::unchanged(),
        });
    }
    if !source.state.source.source_is_exact() {
        return Err(BodyTableTitleError::UnsupportedSource);
    }
    let package = rewrite_title(source, &edit.target, edit.before, edit.settings)?;
    let target = package.state.source.shared_source();
    let target_fingerprint = page_layout::fingerprint(target.as_ref());
    let target_preview_count = page_layout::preview_count(&package);
    Ok(BodyTableTitleCommit {
        package,
        patch: BodyTableTitlePatch {
            source: source_bytes,
            target,
            source_fingerprint,
            target_fingerprint,
            proof: edit.target,
            before: edit.before,
            after: edit.settings,
            source_preview_count,
            target_preview_count,
        },
        diagnostics: BodyTableTitleDiagnostics::published(
            source_preview_count.saturating_sub(target_preview_count),
        ),
    })
}

fn resolve_target(
    package: &Package,
    selector: BodyTableSelector<'_>,
) -> Result<table_lock::BodyTableTarget, BodyTableTitleError> {
    package.resolve_body_table(selector).map_err(map_lock_error)
}

fn settings_at_target(
    package: &Package,
    target: &table_lock::BodyTableTarget,
) -> Result<Settings, BodyTableTitleError> {
    let mut budget =
        table_lock::WireBudget::new(package.state.source.limits()).map_err(map_lock_error)?;
    settings_at_target_with_budget(package, target, &mut budget)
}

fn settings_at_target_with_budget(
    package: &Package,
    target: &table_lock::BodyTableTarget,
    budget: &mut table_lock::WireBudget,
) -> Result<Settings, BodyTableTitleError> {
    table_lock::validate_body_table_target(package, target, budget).map_err(map_lock_error)?;
    let message = model_message(package, target)?;
    let snapshot = decode_snapshot(&message.data, budget)?;
    Ok(Settings::new(
        snapshot.table_name_enabled(),
        snapshot.table_name_border_enabled(),
    ))
}

fn model_message<'a>(
    package: &'a Package,
    target: &table_lock::BodyTableTarget,
) -> Result<&'a RawMessage, BodyTableTitleError> {
    let component = package
        .state
        .source
        .components()
        .get_index(target.model_component_index)
        .ok_or(BodyTableTitleError::InvalidSource)?;
    let object = component
        .archive()
        .objects
        .get(target.model_object_index)
        .ok_or(BodyTableTitleError::InvalidSource)?;
    if object.archive_info.identifier != Some(target.model_identifier.get()) {
        return Err(BodyTableTitleError::InvalidSource);
    }
    let message = object
        .messages
        .get(target.model_message_index)
        .filter(|message| message.type_ == target.model_message_type)
        .ok_or(BodyTableTitleError::InvalidSource)?;
    object
        .archive_info
        .message_infos
        .get(target.model_message_index)
        .ok_or(BodyTableTitleError::InvalidSource)?;
    Ok(message)
}

fn decode_snapshot(
    source: &[u8],
    budget: &mut table_lock::WireBudget,
) -> Result<TableTitleSettingsSnapshot, BodyTableTitleError> {
    let limits = budget.wire_limits();
    let max = source.len().max(1);
    let options = numbers_table_title_codec::DecodeOptions::new(
        max.min(limits.max_input_bytes()),
        limits.max_fields(),
        limits.max_rewrite_work(),
        u32::try_from(limits.max_nesting()).unwrap_or(u32::MAX),
        budget.maximum_payload_references(),
    );
    let (snapshot, report) =
        numbers_table_title_codec::decode_table_title_settings_with_report(source, options)
            .map_err(map_codec_error)?;
    budget
        .charge_codec_report(
            report.fields(),
            report.work_bytes(),
            report.max_depth(),
            report.references(),
        )
        .map_err(map_lock_error)?;
    Ok(snapshot)
}

fn rewrite_title(
    source: &Package,
    target: &table_lock::BodyTableTarget,
    before: Settings,
    after: Settings,
) -> Result<Package, BodyTableTitleError> {
    let mut budget =
        table_lock::WireBudget::new(source.state.source.limits()).map_err(map_lock_error)?;
    table_lock::validate_body_table_target(source, target, &mut budget).map_err(map_lock_error)?;
    let (_component_name, stream_length) =
        table_lock::preflight_body_table_component(source, target, &mut budget)
            .map_err(map_lock_error)?;
    let component = source
        .state
        .source
        .components()
        .get_index(target.model_component_index)
        .ok_or(BodyTableTitleError::InvalidSource)?;
    let component_name = component.name();
    let (mut archive, archive_limits) =
        page_layout::editable_archive(source, component_name).map_err(map_page_layout_error)?;
    let object = archive
        .objects
        .get_mut(target.model_object_index)
        .ok_or(BodyTableTitleError::InvalidSource)?;
    if object.archive_info.identifier != Some(target.model_identifier.get()) {
        return Err(BodyTableTitleError::InvalidSource);
    }
    page_layout::validate_selected_metadata(object, target.model_message_index)
        .map_err(map_page_layout_error)?;
    let message = object
        .messages
        .get(target.model_message_index)
        .filter(|message| message.type_ == target.model_message_type)
        .ok_or(BodyTableTitleError::InvalidSource)?;
    let decoded = decode_snapshot(&message.data, &mut budget)?;
    let actual = Settings::new(
        decoded.table_name_enabled(),
        decoded.table_name_border_enabled(),
    );
    if actual != before {
        return Err(BodyTableTitleError::InvalidSource);
    }
    validate_visible_prerequisites(source, target, decoded, after, &mut budget)?;
    let limits = title_wire_limits(source)?
        .with_output_bytes(message.data.len().saturating_add(64))
        .map_err(map_wire_error)?
        .with_rewrite_work(message.data.len().saturating_mul(8).max(1))
        .map_err(map_wire_error)?;
    let paths = [[TITLE_VISIBLE_FIELD], [TITLE_OUTLINED_FIELD]];
    let edits = [
        NestedFieldEdit::new(
            &paths[0],
            before.visible().is_some(),
            NestedFieldReplacement::Varint(after.visible().map(u64::from)),
        ),
        NestedFieldEdit::new(
            &paths[1],
            before.outlined().is_some(),
            NestedFieldReplacement::Varint(after.outlined().map(u64::from)),
        ),
    ];
    let rewritten_bound = message
        .data
        .len()
        .checked_add(64)
        .ok_or(BodyTableTitleError::InvalidSource)?;
    budget
        .charge_output_bytes(rewritten_bound)
        .map_err(map_lock_error)?;
    budget
        .charge_payload_bytes(rewritten_bound)
        .map_err(map_lock_error)?;
    budget
        .charge_total_payload_bytes(rewritten_bound)
        .map_err(map_lock_error)?;
    budget
        .charge_payload_work(rewritten_bound)
        .map_err(map_lock_error)?;
    let package_bound = source
        .source_bytes()
        .len()
        .checked_add(stream_length)
        .and_then(|value| value.checked_add(rewritten_bound.saturating_mul(2)))
        .and_then(|value| value.checked_add(1_024))
        .ok_or(BodyTableTitleError::InvalidSource)?;
    budget
        .charge_output_bytes(package_bound)
        .map_err(map_lock_error)?;
    let rewritten = patch_nested_fields_batched_with_limits(&message.data, &edits, limits)
        .map_err(map_wire_error)?;
    let verified = decode_snapshot(&rewritten, &mut budget)?;
    if Settings::new(
        verified.table_name_enabled(),
        verified.table_name_border_enabled(),
    ) != after
    {
        return Err(BodyTableTitleError::Verification);
    }
    object
        .replace_message_preserving_header_with_limits(
            target.model_message_index,
            RawMessage {
                type_: target.model_message_type,
                data: rewritten,
            },
            archive_limits,
        )
        .map_err(map_core_error)?;
    let compressed =
        page_layout::compress_archive(archive, archive_limits).map_err(map_page_layout_error)?;
    let mut deletions = Vec::new();
    deletions
        .try_reserve_exact(PREVIEW_ENTRY_NAMES.len())
        .map_err(|_| BodyTableTitleError::Allocation {
            amount: PREVIEW_ENTRY_NAMES.len(),
        })?;
    for name in PREVIEW_ENTRY_NAMES {
        if source
            .state
            .source
            .package()
            .iter()
            .any(|entry| entry.name() == name)
        {
            deletions.push(name);
        }
    }
    let output = source
        .state
        .source
        .package()
        .reassemble_with_deletions_to_bytes(
            &[EntryEdit::new(component_name, &compressed)],
            &deletions,
            source.state.source.limits(),
        )
        .map_err(map_archive_error)?;
    if output.len() > package_bound {
        return Err(BodyTableTitleError::Verification);
    }
    budget
        .charge_input_source(&output)
        .map_err(map_lock_error)?;
    budget
        .charge_payload_work(output.len())
        .map_err(map_lock_error)?;
    let candidate_source =
        SourceCatalog::from_shared_bytes_with_limits(output.into(), source.state.source.limits())
            .map_err(map_archive_error)?;
    budget
        .charge_source_catalog(&candidate_source)
        .map_err(map_lock_error)?;
    let candidate = Package::from_source_catalog(candidate_source).map_err(map_package_error)?;
    if settings_at_target_with_budget(&candidate, target, &mut budget)? != after {
        return Err(BodyTableTitleError::Verification);
    }
    Ok(candidate)
}

fn validate_visible_prerequisites(
    package: &Package,
    target: &table_lock::BodyTableTarget,
    snapshot: TableTitleSettingsSnapshot,
    after: Settings,
    budget: &mut table_lock::WireBudget,
) -> Result<(), BodyTableTitleError> {
    if !after.is_visible() {
        return Ok(());
    }
    let Some(height_bits) = snapshot.table_name_height_bits() else {
        return Err(BodyTableTitleError::InvalidSource);
    };
    let height = f64::from_bits(height_bits);
    if !height.is_finite() || height < 0.0 {
        return Err(BodyTableTitleError::InvalidSource);
    }
    let Some(style) = snapshot.table_name_style() else {
        return Err(BodyTableTitleError::InvalidSource);
    };
    let Some(shape) = snapshot.table_name_shape_style() else {
        return Err(BodyTableTitleError::InvalidSource);
    };
    if style.identifier() == 0
        || shape.identifier() == 0
        || style.identifier() == shape.identifier()
    {
        return Err(BodyTableTitleError::InvalidSource);
    }
    if style.deprecated_is_external() == Some(true) || shape.deprecated_is_external() == Some(true)
    {
        return Err(BodyTableTitleError::InvalidSource);
    }
    let component = package
        .state
        .source
        .components()
        .get_index(target.model_component_index)
        .ok_or(BodyTableTitleError::InvalidSource)?;
    let model = component
        .archive()
        .objects
        .get(target.model_object_index)
        .ok_or(BodyTableTitleError::InvalidSource)?;
    let info = model
        .archive_info
        .message_infos
        .get(target.model_message_index)
        .ok_or(BodyTableTitleError::InvalidSource)?;
    budget
        .charge_payload_work(
            info.object_references
                .len()
                .saturating_add(info.data_references.len())
                .saturating_add(info.field_infos.len()),
        )
        .map_err(map_lock_error)?;
    let rooted_role_ids = [
        1_u64,
        target.body_identifier.get(),
        target.attachment_identifier.get(),
        target.drawable_identifier.get(),
        target.model_identifier.get(),
    ];
    if rooted_role_ids.contains(&style.identifier())
        || rooted_role_ids.contains(&shape.identifier())
        || info
            .data_references
            .iter()
            .any(|identifier| [style.identifier(), shape.identifier()].contains(identifier))
        || info.field_infos.iter().any(|field| {
            field
                .data_references
                .iter()
                .any(|identifier| [style.identifier(), shape.identifier()].contains(identifier))
        })
    {
        return Err(BodyTableTitleError::InvalidSource);
    }
    if info
        .object_references
        .iter()
        .filter(|identifier| **identifier == style.identifier())
        .count()
        != 1
        || info
            .object_references
            .iter()
            .filter(|identifier| **identifier == shape.identifier())
            .count()
            != 1
    {
        return Err(BodyTableTitleError::InvalidSource);
    }
    for (field_number, identifier) in [(30_u32, style.identifier()), (36_u32, shape.identifier())] {
        let mut declaration = None;
        for field in &info.field_infos {
            if field.path.as_slice() != [field_number] {
                continue;
            }
            if declaration.replace(field).is_some() {
                return Err(BodyTableTitleError::InvalidSource);
            }
        }
        if declaration.is_some_and(|field| {
            field
                .r#type
                .is_some_and(|kind| !matches!(kind, litchi_iwa_core::FieldType::ObjectReference))
                || !field.data_references.is_empty()
                || field.object_references.as_slice() != [identifier]
        }) {
            return Err(BodyTableTitleError::InvalidSource);
        }
    }
    require_style_object(package, style.identifier(), 2_022, budget)?;
    require_style_object(package, shape.identifier(), 2_025, budget)?;
    Ok(())
}

fn require_style_object(
    package: &Package,
    identifier: u64,
    message_type: u32,
    budget: &mut table_lock::WireBudget,
) -> Result<(), BodyTableTitleError> {
    let mut matched = None;
    for component in package.state.source.components().iter() {
        budget.charge_payload_work(1).map_err(map_lock_error)?;
        for object in &component.archive().objects {
            budget.charge_payload_work(1).map_err(map_lock_error)?;
            if object.archive_info.identifier == Some(identifier) {
                if matched.replace(object).is_some() {
                    return Err(BodyTableTitleError::InvalidSource);
                }
            }
        }
    }
    let object = matched.ok_or(BodyTableTitleError::InvalidSource)?;
    if object.messages.len() != 1 || object.archive_info.message_infos.len() != 1 {
        return Err(BodyTableTitleError::InvalidSource);
    }
    if object.messages[0].type_ != message_type {
        return Err(BodyTableTitleError::InvalidSource);
    }
    page_layout::validate_selected_metadata(object, 0).map_err(map_page_layout_error)?;
    Ok(())
}

fn reopen_target(source: &Package, target: Arc<[u8]>) -> Result<Package, BodyTableTitleError> {
    let catalog =
        SourceCatalog::from_shared_bytes_with_limits(target, source.state.source.limits())
            .map_err(map_archive_error)?;
    Package::from_source_catalog(catalog).map_err(map_package_error)
}

fn title_wire_limits(package: &Package) -> Result<WireLimits, BodyTableTitleError> {
    let maximum = package
        .state
        .source
        .limits()
        .effective_archive_limits()
        .map_err(map_archive_error)?
        .max_message_bytes();
    WireLimits::default()
        .with_input_bytes(maximum.max(1))
        .and_then(|limits| limits.with_output_bytes(maximum.max(1)))
        .map_err(map_wire_error)
}

fn map_lock_error(error: table_lock::BodyTableLockError) -> BodyTableTitleError {
    match error {
        table_lock::BodyTableLockError::TableNotFound => BodyTableTitleError::TableNotFound,
        table_lock::BodyTableLockError::AmbiguousTableName => {
            BodyTableTitleError::AmbiguousTableName
        },
        table_lock::BodyTableLockError::AmbiguousSelector => BodyTableTitleError::AmbiguousSelector,
        table_lock::BodyTableLockError::UnsupportedSource => BodyTableTitleError::UnsupportedSource,
        table_lock::BodyTableLockError::InvalidSource => BodyTableTitleError::InvalidSource,
        table_lock::BodyTableLockError::LimitExceeded {
            kind,
            observed,
            maximum,
        } => BodyTableTitleError::LimitExceeded {
            kind: map_lock_limit(kind),
            observed,
            maximum,
        },
        table_lock::BodyTableLockError::Allocation { amount } => {
            BodyTableTitleError::Allocation { amount }
        },
        table_lock::BodyTableLockError::Verification => BodyTableTitleError::Verification,
        table_lock::BodyTableLockError::PatchConflict => BodyTableTitleError::PatchConflict,
    }
}

const fn map_lock_limit(kind: table_lock::BodyTableLockLimitKind) -> BodyTableTitleLimitKind {
    use table_lock::BodyTableLockLimitKind as Lock;
    match kind {
        Lock::InputBytes => BodyTableTitleLimitKind::InputBytes,
        Lock::OutputBytes => BodyTableTitleLimitKind::OutputBytes,
        Lock::Entries => BodyTableTitleLimitKind::Entries,
        Lock::EntryBytes => BodyTableTitleLimitKind::EntryBytes,
        Lock::TotalEntryBytes => BodyTableTitleLimitKind::TotalEntryBytes,
        Lock::PackageBytes => BodyTableTitleLimitKind::PayloadItems,
        Lock::PayloadBytes => BodyTableTitleLimitKind::PayloadBytes,
        Lock::TotalPayloadBytes => BodyTableTitleLimitKind::TotalPayloadBytes,
        Lock::PayloadObjects => BodyTableTitleLimitKind::PayloadObjects,
        Lock::PayloadMessages => BodyTableTitleLimitKind::PayloadMessages,
        Lock::PayloadItems => BodyTableTitleLimitKind::PayloadItems,
        Lock::PayloadReferences => BodyTableTitleLimitKind::PayloadReferences,
        Lock::WireBytes => BodyTableTitleLimitKind::WireBytes,
        Lock::WireFields => BodyTableTitleLimitKind::WireFields,
        Lock::WireNesting => BodyTableTitleLimitKind::WireNesting,
        Lock::WireWork => BodyTableTitleLimitKind::WireWork,
    }
}

fn map_page_layout_error(error: page_layout::PageLayoutError) -> BodyTableTitleError {
    match error {
        page_layout::PageLayoutError::UnsupportedSource => BodyTableTitleError::UnsupportedSource,
        page_layout::PageLayoutError::InvalidSource => BodyTableTitleError::InvalidSource,
        page_layout::PageLayoutError::LimitExceeded {
            kind,
            observed,
            maximum,
        } => BodyTableTitleError::LimitExceeded {
            kind: map_page_limit(kind),
            observed,
            maximum,
        },
        page_layout::PageLayoutError::Allocation { amount } => {
            BodyTableTitleError::Allocation { amount }
        },
        _ => BodyTableTitleError::InvalidSource,
    }
}

const fn map_page_limit(kind: page_layout::PageLayoutLimitKind) -> BodyTableTitleLimitKind {
    use page_layout::PageLayoutLimitKind as Layout;
    match kind {
        Layout::InputBytes => BodyTableTitleLimitKind::InputBytes,
        Layout::OutputBytes => BodyTableTitleLimitKind::OutputBytes,
        Layout::Entries => BodyTableTitleLimitKind::Entries,
        Layout::EntryBytes => BodyTableTitleLimitKind::EntryBytes,
        Layout::TotalEntryBytes => BodyTableTitleLimitKind::TotalEntryBytes,
        Layout::PackageBytes => BodyTableTitleLimitKind::PayloadItems,
        Layout::PayloadBytes => BodyTableTitleLimitKind::PayloadBytes,
        Layout::TotalPayloadBytes => BodyTableTitleLimitKind::TotalPayloadBytes,
        Layout::PayloadObjects => BodyTableTitleLimitKind::PayloadObjects,
        Layout::PayloadMessages => BodyTableTitleLimitKind::PayloadMessages,
        Layout::PayloadItems => BodyTableTitleLimitKind::PayloadItems,
        Layout::WireBytes => BodyTableTitleLimitKind::WireBytes,
        Layout::WireFields => BodyTableTitleLimitKind::WireFields,
        Layout::WireNesting => BodyTableTitleLimitKind::WireNesting,
        Layout::WireWork => BodyTableTitleLimitKind::WireWork,
    }
}

fn map_codec_error(error: numbers_table_title_codec::DecodeError) -> BodyTableTitleError {
    match error.resource_limit() {
        Some(DecodeLimit::Bytes { observed, maximum }) => BodyTableTitleError::LimitExceeded {
            kind: BodyTableTitleLimitKind::WireBytes,
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        },
        Some(DecodeLimit::References { observed, maximum }) => BodyTableTitleError::LimitExceeded {
            kind: BodyTableTitleLimitKind::PayloadReferences,
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        },
        Some(DecodeLimit::Fields { observed, maximum }) => BodyTableTitleError::LimitExceeded {
            kind: BodyTableTitleLimitKind::WireFields,
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        },
        Some(DecodeLimit::Work { observed, maximum }) => BodyTableTitleError::LimitExceeded {
            kind: BodyTableTitleLimitKind::WireWork,
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        },
        Some(DecodeLimit::Nesting { observed, maximum }) => BodyTableTitleError::LimitExceeded {
            kind: BodyTableTitleLimitKind::WireNesting,
            observed: u64::from(observed),
            maximum: u64::from(maximum),
        },
        Some(_) | None => BodyTableTitleError::InvalidSource,
    }
}

fn map_archive_error(error: litchi_iwa_archive::Error) -> BodyTableTitleError {
    match error {
        litchi_iwa_archive::Error::Limit {
            kind,
            observed,
            maximum,
        } => BodyTableTitleError::LimitExceeded {
            kind: match kind {
                litchi_iwa_archive::LimitKind::InputBytes => BodyTableTitleLimitKind::InputBytes,
                litchi_iwa_archive::LimitKind::OutputBytes => BodyTableTitleLimitKind::OutputBytes,
                litchi_iwa_archive::LimitKind::Entries => BodyTableTitleLimitKind::Entries,
                litchi_iwa_archive::LimitKind::MemberNameBytes
                | litchi_iwa_archive::LimitKind::MetadataBytes => {
                    BodyTableTitleLimitKind::PayloadItems
                },
                litchi_iwa_archive::LimitKind::CompressedEntryBytes
                | litchi_iwa_archive::LimitKind::EntryBytes => BodyTableTitleLimitKind::EntryBytes,
                litchi_iwa_archive::LimitKind::TotalBytes => {
                    BodyTableTitleLimitKind::TotalEntryBytes
                },
                litchi_iwa_archive::LimitKind::IwaStreamBytes => {
                    BodyTableTitleLimitKind::PayloadBytes
                },
                litchi_iwa_archive::LimitKind::IwaTotalBytes => {
                    BodyTableTitleLimitKind::TotalPayloadBytes
                },
            },
            observed,
            maximum,
        },
        litchi_iwa_archive::Error::Allocation { amount, .. } => {
            BodyTableTitleError::Allocation { amount }
        },
        litchi_iwa_archive::Error::Reassembly(_) => BodyTableTitleError::UnsupportedSource,
        litchi_iwa_archive::Error::Iwa(error) => map_core_error(error),
        _ => BodyTableTitleError::InvalidSource,
    }
}

fn map_core_error(error: litchi_iwa_core::Error) -> BodyTableTitleError {
    match error {
        litchi_iwa_core::Error::Limit {
            kind,
            observed,
            maximum,
        } => BodyTableTitleError::LimitExceeded {
            kind: match kind {
                litchi_iwa_core::LimitKind::Objects => BodyTableTitleLimitKind::PayloadObjects,
                litchi_iwa_core::LimitKind::Messages
                | litchi_iwa_core::LimitKind::MessagesPerObject => {
                    BodyTableTitleLimitKind::PayloadMessages
                },
                litchi_iwa_core::LimitKind::HeaderFields
                | litchi_iwa_core::LimitKind::MetadataItems
                | litchi_iwa_core::LimitKind::SnappyFrames => BodyTableTitleLimitKind::PayloadItems,
                litchi_iwa_core::LimitKind::HeaderNesting => BodyTableTitleLimitKind::WireNesting,
                _ => BodyTableTitleLimitKind::PayloadBytes,
            },
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        },
        litchi_iwa_core::Error::Allocation { requested, .. } => {
            BodyTableTitleError::Allocation { amount: requested }
        },
        _ => BodyTableTitleError::InvalidSource,
    }
}

fn map_package_error(error: PackageError) -> BodyTableTitleError {
    match error {
        PackageError::Archive(error) => map_archive_error(error),
        PackageError::Allocation { amount } => BodyTableTitleError::Allocation { amount },
        PackageError::ObjectLimit { observed, limit } => BodyTableTitleError::LimitExceeded {
            kind: BodyTableTitleLimitKind::PayloadObjects,
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(limit),
        },
        PackageError::PayloadLimit { observed, limit } => BodyTableTitleError::LimitExceeded {
            kind: BodyTableTitleLimitKind::PayloadBytes,
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(limit),
        },
        _ => BodyTableTitleError::InvalidSource,
    }
}

fn map_wire_error(error: litchi_iwa_common::Error) -> BodyTableTitleError {
    match error {
        litchi_iwa_common::Error::LimitExceeded {
            kind,
            observed,
            limit,
        } => BodyTableTitleError::LimitExceeded {
            kind: match kind {
                litchi_iwa_common::LimitKind::InputBytes => BodyTableTitleLimitKind::WireBytes,
                litchi_iwa_common::LimitKind::OutputBytes => {
                    BodyTableTitleLimitKind::WireOutputBytes
                },
                litchi_iwa_common::LimitKind::Fields => BodyTableTitleLimitKind::WireFields,
                litchi_iwa_common::LimitKind::Nesting => BodyTableTitleLimitKind::WireNesting,
                litchi_iwa_common::LimitKind::RewriteWork => BodyTableTitleLimitKind::WireWork,
                _ => BodyTableTitleLimitKind::PayloadItems,
            },
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(limit),
        },
        litchi_iwa_common::Error::Allocation { amount, .. } => {
            BodyTableTitleError::Allocation { amount }
        },
        _ => BodyTableTitleError::InvalidSource,
    }
}

fn usize_to_u64(value: usize) -> u64 {
    u64::try_from(value).unwrap_or(u64::MAX)
}
