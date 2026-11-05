//! Exact-source transactions for Pages body-table title settings.
//!
//! The rooted body-table graph is resolved by `table_lock`; this module only
//! owns the title payload projection and the immutable transaction around it.
//! The strict title codec validates the selected scalar/reference fields while
//! the original model payload remains the source-preservation authority.

use std::fmt;
use std::sync::Arc;

use litchi_iwa_archive::package::EntryEdit;
use litchi_iwa_common::{
    WireLimits,
    wire::{NestedFieldEdit, NestedFieldReplacement, patch_nested_fields_batched},
};
use litchi_iwa_core::RawMessage;
use litchi_iwa_protos::numbers_table_title_codec::{self, DecodeLimit, TableTitleSettingsSnapshot};
use thiserror::Error as ThisError;

use super::{Package, PackageError, page_layout, table_lock};
use crate::selector::BodyTableSelector;
use crate::table::title::Settings;

const TITLE_VISIBLE_FIELD: u32 = 22;
const TITLE_OUTLINED_FIELD: u32 = 37;
const PREVIEW_ENTRY_NAMES: [&str; 3] = ["preview.jpg", "preview-micro.jpg", "preview-web.jpg"];

/// Finite resource categories enforced by a body-table title transaction.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
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

impl fmt::Display for LimitKind {
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
#[derive(Debug, ThisError, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum Error {
    /// A well-formed request this route refuses; [`Refusal::kind`] classifies it
    /// (record 0351 §2.4, record 0361).
    #[error(transparent)]
    Refused(Refusal),
    /// The rooted native graph or selected title payload is malformed.
    #[error("the selected Pages body-table title source is invalid")]
    InvalidSource,
    /// A finite transaction ceiling was exceeded.
    #[error("Pages body-table title {kind} limit exceeded: observed {observed}, maximum {maximum}")]
    Limit {
        /// Resource category.
        kind: LimitKind,
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

/// A well-formed request this family's route refuses, classified by
/// [`Refusal::kind`] (record 0351 §2.4, record 0361).
#[derive(Debug, ThisError, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum Refusal {
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
}

impl Refusal {
    /// The shared refusal kind litchi-ops classifies this reason by.
    #[must_use]
    pub const fn kind(&self) -> litchi_core::edit::RefusalKind {
        use litchi_core::edit::RefusalKind;
        match self {
            Self::TableNotFound { .. } => RefusalKind::NotFound,
            Self::AmbiguousTableName { .. } | Self::AmbiguousSelector { .. } => {
                RefusalKind::InvalidValue
            },
            Self::UnsupportedSource { .. } => RefusalKind::UnsupportedContent,
        }
    }
}

impl From<Refusal> for Error {
    fn from(refusal: Refusal) -> Self {
        Self::Refused(refusal)
    }
}

/// Mutable semantic title settings staged against one immutable package.
pub struct Edit<'a> {
    source: &'a Package,
    selector: table_lock::OwnedBodyTableSelector,
    target: Option<table_lock::BodyTableTarget>,
    before: Option<Settings>,
    settings: Option<Settings>,
}

impl fmt::Debug for Edit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("resolved", &self.target.is_some())
            .field("settings", &self.settings)
            .finish_non_exhaustive()
    }
}

impl Edit<'_> {
    fn resolve(&mut self) -> Result<(), Error> {
        if self.target.is_none() {
            let target = self.selector.resolve(self.source).map_err(map_lock_error)?;
            let before = settings_at_target(self.source, &target)?;
            if self.settings.is_none() {
                self.settings = Some(before);
            }
            self.before = Some(before);
            self.target = Some(target);
        }
        Ok(())
    }

    /// Return the settings that would be published.
    pub fn settings(&mut self) -> Result<Settings, Error> {
        self.resolve()?;
        self.settings.ok_or(Error::InvalidSource)
    }

    /// Replace the complete staged lossless title settings.
    #[must_use]
    pub fn set(mut self, settings: Settings) -> Self {
        self.settings = Some(settings);
        self
    }

    /// Validate and publish the staged settings atomically.
    pub fn commit(mut self) -> Result<Commit, Error> {
        self.resolve()?;
        commit_edit(self)
    }
}

/// Exact-source reversible title patch.
#[derive(Clone, PartialEq, Eq)]
pub struct Patch {
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

impl fmt::Debug for Patch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Patch")
            .field("before", &self.before)
            .field("after", &self.after)
            .finish_non_exhaustive()
    }
}

impl Patch {
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
pub struct Diagnostics {
    changed: bool,
    touched_components: usize,
    deleted_previews: usize,
    full_reparse_performed: bool,
}

impl Diagnostics {
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
pub struct Commit {
    package: Package,
    patch: Patch,
    diagnostics: Diagnostics,
}

impl Commit {
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

    /// Consume the commit and return its package and patch.
    #[must_use]
    pub fn into_parts(self) -> (Package, Patch) {
        (self.package, self.patch)
    }

    /// Borrow the reversible patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Borrow publication diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> &Diagnostics {
        &self.diagnostics
    }
}

impl Package {
    /// Read one rooted body's lossless title visibility and outline settings.
    pub fn body_table_title_settings<'table>(
        &self,
        selector: impl Into<BodyTableSelector<'table>>,
    ) -> Result<Settings, Error> {
        let target = resolve_target(self, selector.into())?;
        settings_at_target(self, &target)
    }

    /// Start a selector-first immutable body-table title edit.
    pub fn edit_body_table_title<'table>(
        &self,
        selector: impl Into<BodyTableSelector<'table>>,
    ) -> Result<Edit<'_>, Error> {
        Ok(Edit {
            source: self,
            selector: table_lock::OwnedBodyTableSelector::try_from(self, selector.into())
                .map_err(map_lock_error)?,
            target: None,
            before: None,
            settings: None,
        })
    }

    /// Apply an exact-source-checked reversible title patch.
    pub fn apply_body_table_title(&self, patch: &Patch) -> Result<Commit, Error> {
        if page_layout::fingerprint(self.source_bytes()) != patch.source_fingerprint
            || self.source_bytes() != patch.source.as_ref()
        {
            return Err(Error::PatchConflict);
        }
        if settings_at_target(self, &patch.proof)? != patch.before {
            return Err(Error::PatchConflict);
        }
        if patch.is_noop() {
            if patch.source_preview_count != patch.target_preview_count {
                return Err(Error::PatchConflict);
            }
            return Ok(Commit {
                package: self.clone(),
                patch: patch.clone(),
                diagnostics: Diagnostics::unchanged(),
            });
        }
        if !self.state.source.source_is_exact()
            || page_layout::preview_count(self) != patch.source_preview_count
            || page_layout::fingerprint(patch.target.as_ref()) != patch.target_fingerprint
        {
            return Err(Error::PatchConflict);
        }
        let candidate = reopen_target(self, Arc::clone(&patch.target))?;
        if settings_at_target(&candidate, &patch.proof)? != patch.after
            || page_layout::preview_count(&candidate) != patch.target_preview_count
        {
            return Err(Error::Verification);
        }
        Ok(Commit {
            package: candidate,
            patch: patch.clone(),
            diagnostics: Diagnostics::published(
                patch
                    .source_preview_count
                    .saturating_sub(patch.target_preview_count),
            ),
        })
    }
}

fn commit_edit(edit: Edit<'_>) -> Result<Commit, Error> {
    let source = edit.source;
    let target = edit.target.ok_or(Error::InvalidSource)?;
    let before = edit.before.ok_or(Error::InvalidSource)?;
    let settings = edit.settings.ok_or(Error::InvalidSource)?;
    let source_bytes: Arc<[u8]> = source.state.source.shared_source();
    let source_fingerprint = page_layout::fingerprint(source_bytes.as_ref());
    let source_preview_count = page_layout::preview_count(source);
    if before == settings {
        return Ok(Commit {
            package: source.clone(),
            patch: Patch {
                source: Arc::clone(&source_bytes),
                target: source_bytes,
                source_fingerprint,
                target_fingerprint: source_fingerprint,
                proof: target,
                before,
                after: settings,
                source_preview_count,
                target_preview_count: source_preview_count,
            },
            diagnostics: Diagnostics::unchanged(),
        });
    }
    if !source.state.source.source_is_exact() {
        return Err(Error::Refused(Refusal::UnsupportedSource));
    }
    let package = rewrite_title(source, &target, before, settings)?;
    let target_owner = package.state.source.shared_source();
    let target_fingerprint = page_layout::fingerprint(target_owner.as_ref());
    let target_preview_count = page_layout::preview_count(&package);
    Ok(Commit {
        package,
        patch: Patch {
            source: source_bytes,
            target: target_owner,
            source_fingerprint,
            target_fingerprint,
            proof: target,
            before,
            after: settings,
            source_preview_count,
            target_preview_count,
        },
        diagnostics: Diagnostics::published(
            source_preview_count.saturating_sub(target_preview_count),
        ),
    })
}

fn resolve_target(
    package: &Package,
    selector: BodyTableSelector<'_>,
) -> Result<table_lock::BodyTableTarget, Error> {
    package.resolve_body_table(selector).map_err(map_lock_error)
}

fn settings_at_target(
    package: &Package,
    target: &table_lock::BodyTableTarget,
) -> Result<Settings, Error> {
    let mut budget =
        table_lock::WireBudget::new(package.state.source.limits()).map_err(map_lock_error)?;
    settings_at_target_with_budget(package, target, &mut budget)
}

fn settings_at_target_with_budget(
    package: &Package,
    target: &table_lock::BodyTableTarget,
    budget: &mut table_lock::WireBudget,
) -> Result<Settings, Error> {
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
) -> Result<&'a RawMessage, Error> {
    let component = package
        .state
        .source
        .components()
        .get_index(target.model_component_index)
        .ok_or(Error::InvalidSource)?;
    let object = component
        .archive()
        .objects
        .get(target.model_object_index)
        .ok_or(Error::InvalidSource)?;
    if object.archive_info.identifier != Some(target.model_identifier.get()) {
        return Err(Error::InvalidSource);
    }
    let message = object
        .messages
        .get(target.model_message_index)
        .filter(|message| message.type_ == target.model_message_type)
        .ok_or(Error::InvalidSource)?;
    object
        .archive_info
        .message_infos
        .get(target.model_message_index)
        .ok_or(Error::InvalidSource)?;
    Ok(message)
}

fn decode_snapshot(
    source: &[u8],
    budget: &mut table_lock::WireBudget,
) -> Result<TableTitleSettingsSnapshot, Error> {
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
) -> Result<Package, Error> {
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
        .ok_or(Error::InvalidSource)?;
    let component_name = component.name();
    let (mut archive, archive_limits) =
        page_layout::editable_archive(source, component_name).map_err(map_page_layout_error)?;
    let object = archive
        .objects
        .get_mut(target.model_object_index)
        .ok_or(Error::InvalidSource)?;
    if object.archive_info.identifier != Some(target.model_identifier.get()) {
        return Err(Error::InvalidSource);
    }
    page_layout::validate_selected_metadata(object, target.model_message_index)
        .map_err(map_page_layout_error)?;
    let message = object
        .messages
        .get(target.model_message_index)
        .filter(|message| message.type_ == target.model_message_type)
        .ok_or(Error::InvalidSource)?;
    let decoded = decode_snapshot(&message.data, &mut budget)?;
    let actual = Settings::new(
        decoded.table_name_enabled(),
        decoded.table_name_border_enabled(),
    );
    if actual != before {
        return Err(Error::InvalidSource);
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
        .ok_or(Error::InvalidSource)?;
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
        .ok_or(Error::InvalidSource)?;
    budget
        .charge_output_bytes(package_bound)
        .map_err(map_lock_error)?;
    let rewritten =
        patch_nested_fields_batched(&message.data, &edits, limits).map_err(map_wire_error)?;
    let verified = decode_snapshot(&rewritten, &mut budget)?;
    if Settings::new(
        verified.table_name_enabled(),
        verified.table_name_border_enabled(),
    ) != after
    {
        return Err(Error::Verification);
    }
    object
        .replace_message_preserving_header(
            target.model_message_index,
            RawMessage {
                type_: target.model_message_type,
                data: rewritten,
            },
            archive_limits,
        )
        .map_err(map_core_error)?;
    let compressed = page_layout::compress_archive(source, archive, archive_limits)
        .map_err(map_page_layout_error)?;
    let mut deletions = Vec::new();
    deletions
        .try_reserve_exact(PREVIEW_ENTRY_NAMES.len())
        .map_err(|_| Error::Allocation {
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
        return Err(Error::Verification);
    }
    budget
        .charge_input_source(&output)
        .map_err(map_lock_error)?;
    budget
        .charge_payload_work(output.len())
        .map_err(map_lock_error)?;
    let candidate_source = source
        .state
        .source
        .reopen(output.into())
        .map_err(map_archive_error)?;
    budget
        .charge_source_catalog(&candidate_source)
        .map_err(map_lock_error)?;
    let candidate = Package::from_source_catalog(candidate_source).map_err(map_package_error)?;
    if settings_at_target_with_budget(&candidate, target, &mut budget)? != after {
        return Err(Error::Verification);
    }
    Ok(candidate)
}

fn validate_visible_prerequisites(
    package: &Package,
    target: &table_lock::BodyTableTarget,
    snapshot: TableTitleSettingsSnapshot,
    after: Settings,
    budget: &mut table_lock::WireBudget,
) -> Result<(), Error> {
    if !after.is_visible() {
        return Ok(());
    }
    let Some(height_bits) = snapshot.table_name_height_bits() else {
        return Err(Error::InvalidSource);
    };
    let height = f64::from_bits(height_bits);
    if !height.is_finite() || height < 0.0 {
        return Err(Error::InvalidSource);
    }
    let Some(style) = snapshot.table_name_style() else {
        return Err(Error::InvalidSource);
    };
    let Some(shape) = snapshot.table_name_shape_style() else {
        return Err(Error::InvalidSource);
    };
    if style.identifier() == 0
        || shape.identifier() == 0
        || style.identifier() == shape.identifier()
    {
        return Err(Error::InvalidSource);
    }
    if style.deprecated_is_external() == Some(true) || shape.deprecated_is_external() == Some(true)
    {
        return Err(Error::InvalidSource);
    }
    let component = package
        .state
        .source
        .components()
        .get_index(target.model_component_index)
        .ok_or(Error::InvalidSource)?;
    let model = component
        .archive()
        .objects
        .get(target.model_object_index)
        .ok_or(Error::InvalidSource)?;
    let info = model
        .archive_info
        .message_infos
        .get(target.model_message_index)
        .ok_or(Error::InvalidSource)?;
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
        return Err(Error::InvalidSource);
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
        return Err(Error::InvalidSource);
    }
    for (field_number, identifier) in [(30_u32, style.identifier()), (36_u32, shape.identifier())] {
        let mut declaration = None;
        for field in &info.field_infos {
            if field.path.as_slice() != [field_number] {
                continue;
            }
            if declaration.replace(field).is_some() {
                return Err(Error::InvalidSource);
            }
        }
        if declaration.is_some_and(|field| {
            field
                .r#type
                .is_some_and(|kind| !matches!(kind, litchi_iwa_core::FieldType::ObjectReference))
                || !field.data_references.is_empty()
                || field.object_references.as_slice() != [identifier]
        }) {
            return Err(Error::InvalidSource);
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
) -> Result<(), Error> {
    let mut matched = None;
    for component in package.state.source.components().iter() {
        budget.charge_payload_work(1).map_err(map_lock_error)?;
        for object in &component.archive().objects {
            budget.charge_payload_work(1).map_err(map_lock_error)?;
            if object.archive_info.identifier == Some(identifier) {
                if matched.replace(object).is_some() {
                    return Err(Error::InvalidSource);
                }
            }
        }
    }
    let object = matched.ok_or(Error::InvalidSource)?;
    if object.messages.len() != 1 || object.archive_info.message_infos.len() != 1 {
        return Err(Error::InvalidSource);
    }
    if object.messages[0].type_ != message_type {
        return Err(Error::InvalidSource);
    }
    page_layout::validate_selected_metadata(object, 0).map_err(map_page_layout_error)?;
    Ok(())
}

fn reopen_target(source: &Package, target: Arc<[u8]>) -> Result<Package, Error> {
    let catalog = source
        .state
        .source
        .reopen(target)
        .map_err(map_archive_error)?;
    Package::from_source_catalog(catalog).map_err(map_package_error)
}

fn title_wire_limits(package: &Package) -> Result<WireLimits, Error> {
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

fn map_lock_error(error: table_lock::Error) -> Error {
    match error {
        table_lock::Error::Refused(crate::table::lock::Refusal::TableNotFound) => {
            Error::Refused(Refusal::TableNotFound)
        },
        table_lock::Error::Refused(crate::table::lock::Refusal::AmbiguousTableName) => {
            Error::Refused(Refusal::AmbiguousTableName)
        },
        table_lock::Error::Refused(crate::table::lock::Refusal::AmbiguousSelector) => {
            Error::Refused(Refusal::AmbiguousSelector)
        },
        table_lock::Error::Refused(crate::table::lock::Refusal::UnsupportedSource) => {
            Error::Refused(Refusal::UnsupportedSource)
        },
        table_lock::Error::InvalidSource => Error::InvalidSource,
        table_lock::Error::Limit {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: map_lock_limit(kind),
            observed,
            maximum,
        },
        table_lock::Error::Allocation { amount } => Error::Allocation { amount },
        table_lock::Error::Verification => Error::Verification,
        table_lock::Error::PatchConflict => Error::PatchConflict,
    }
}

const fn map_lock_limit(kind: table_lock::LimitKind) -> LimitKind {
    use table_lock::LimitKind as Lock;
    match kind {
        Lock::InputBytes => LimitKind::InputBytes,
        Lock::OutputBytes => LimitKind::OutputBytes,
        Lock::Entries => LimitKind::Entries,
        Lock::EntryBytes => LimitKind::EntryBytes,
        Lock::TotalEntryBytes => LimitKind::TotalEntryBytes,
        Lock::PackageBytes => LimitKind::PayloadItems,
        Lock::PayloadBytes => LimitKind::PayloadBytes,
        Lock::TotalPayloadBytes => LimitKind::TotalPayloadBytes,
        Lock::PayloadObjects => LimitKind::PayloadObjects,
        Lock::PayloadMessages => LimitKind::PayloadMessages,
        Lock::PayloadItems => LimitKind::PayloadItems,
        Lock::PayloadReferences => LimitKind::PayloadReferences,
        Lock::WireBytes => LimitKind::WireBytes,
        Lock::WireFields => LimitKind::WireFields,
        Lock::WireNesting => LimitKind::WireNesting,
        Lock::WireWork => LimitKind::WireWork,
    }
}

fn map_page_layout_error(error: page_layout::Error) -> Error {
    match error {
        page_layout::Error::Refused(
            crate::page_layout::transaction::Refusal::UnsupportedSource,
        ) => Error::Refused(Refusal::UnsupportedSource),
        page_layout::Error::InvalidSource => Error::InvalidSource,
        page_layout::Error::Limit {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: map_page_limit(kind),
            observed,
            maximum,
        },
        page_layout::Error::Allocation { amount } => Error::Allocation { amount },
        _ => Error::InvalidSource,
    }
}

const fn map_page_limit(kind: page_layout::LimitKind) -> LimitKind {
    use page_layout::LimitKind as Layout;
    match kind {
        Layout::InputBytes => LimitKind::InputBytes,
        Layout::OutputBytes => LimitKind::OutputBytes,
        Layout::Entries => LimitKind::Entries,
        Layout::EntryBytes => LimitKind::EntryBytes,
        Layout::TotalEntryBytes => LimitKind::TotalEntryBytes,
        Layout::PackageBytes => LimitKind::PayloadItems,
        Layout::PayloadBytes => LimitKind::PayloadBytes,
        Layout::TotalPayloadBytes => LimitKind::TotalPayloadBytes,
        Layout::PayloadObjects => LimitKind::PayloadObjects,
        Layout::PayloadMessages => LimitKind::PayloadMessages,
        Layout::PayloadItems => LimitKind::PayloadItems,
        Layout::WireBytes => LimitKind::WireBytes,
        Layout::WireFields => LimitKind::WireFields,
        Layout::WireNesting => LimitKind::WireNesting,
        Layout::WireWork => LimitKind::WireWork,
    }
}

fn map_codec_error(error: numbers_table_title_codec::DecodeError) -> Error {
    match error.resource_limit() {
        Some(DecodeLimit::Bytes { observed, maximum }) => Error::Limit {
            kind: LimitKind::WireBytes,
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        },
        Some(DecodeLimit::References { observed, maximum }) => Error::Limit {
            kind: LimitKind::PayloadReferences,
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        },
        Some(DecodeLimit::Fields { observed, maximum }) => Error::Limit {
            kind: LimitKind::WireFields,
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        },
        Some(DecodeLimit::Work { observed, maximum }) => Error::Limit {
            kind: LimitKind::WireWork,
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        },
        Some(DecodeLimit::Nesting { observed, maximum }) => Error::Limit {
            kind: LimitKind::WireNesting,
            observed: u64::from(observed),
            maximum: u64::from(maximum),
        },
        Some(_) | None => Error::InvalidSource,
    }
}

fn map_archive_error(error: litchi_iwa_archive::Error) -> Error {
    match error {
        litchi_iwa_archive::Error::Limit {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: match kind {
                litchi_iwa_archive::LimitKind::InputBytes => LimitKind::InputBytes,
                litchi_iwa_archive::LimitKind::OutputBytes => LimitKind::OutputBytes,
                litchi_iwa_archive::LimitKind::Entries => LimitKind::Entries,
                litchi_iwa_archive::LimitKind::MemberNameBytes
                | litchi_iwa_archive::LimitKind::MetadataBytes => LimitKind::PayloadItems,
                litchi_iwa_archive::LimitKind::CompressedEntryBytes
                | litchi_iwa_archive::LimitKind::EntryBytes => LimitKind::EntryBytes,
                litchi_iwa_archive::LimitKind::TotalBytes => LimitKind::TotalEntryBytes,
                litchi_iwa_archive::LimitKind::IwaStreamBytes => LimitKind::PayloadBytes,
                litchi_iwa_archive::LimitKind::IwaTotalBytes => LimitKind::TotalPayloadBytes,
            },
            observed,
            maximum,
        },
        litchi_iwa_archive::Error::Allocation { amount, .. } => Error::Allocation { amount },
        litchi_iwa_archive::Error::Reassembly(_) => Error::Refused(Refusal::UnsupportedSource),
        litchi_iwa_archive::Error::Iwa(error) => map_core_error(error),
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
                litchi_iwa_core::LimitKind::Objects => LimitKind::PayloadObjects,
                litchi_iwa_core::LimitKind::Messages
                | litchi_iwa_core::LimitKind::MessagesPerObject => LimitKind::PayloadMessages,
                litchi_iwa_core::LimitKind::HeaderFields
                | litchi_iwa_core::LimitKind::MetadataItems
                | litchi_iwa_core::LimitKind::SnappyFrames => LimitKind::PayloadItems,
                litchi_iwa_core::LimitKind::HeaderNesting => LimitKind::WireNesting,
                _ => LimitKind::PayloadBytes,
            },
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        },
        litchi_iwa_core::Error::Allocation { requested, .. } => {
            Error::Allocation { amount: requested }
        },
        _ => Error::InvalidSource,
    }
}

fn map_package_error(error: PackageError) -> Error {
    match error {
        PackageError::Archive(error) => map_archive_error(error),
        PackageError::Allocation { amount } => Error::Allocation { amount },
        PackageError::ObjectLimit { observed, limit } => Error::Limit {
            kind: LimitKind::PayloadObjects,
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(limit),
        },
        PackageError::PayloadLimit { observed, limit } => Error::Limit {
            kind: LimitKind::PayloadBytes,
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(limit),
        },
        _ => Error::InvalidSource,
    }
}

fn map_wire_error(error: litchi_iwa_common::Error) -> Error {
    match error {
        litchi_iwa_common::Error::LimitExceeded {
            kind,
            observed,
            limit,
        } => Error::Limit {
            kind: match kind {
                litchi_iwa_common::LimitKind::InputBytes => LimitKind::WireBytes,
                litchi_iwa_common::LimitKind::OutputBytes => LimitKind::WireOutputBytes,
                litchi_iwa_common::LimitKind::Fields => LimitKind::WireFields,
                litchi_iwa_common::LimitKind::Nesting => LimitKind::WireNesting,
                litchi_iwa_common::LimitKind::RewriteWork => LimitKind::WireWork,
                _ => LimitKind::PayloadItems,
            },
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(limit),
        },
        litchi_iwa_common::Error::Allocation { amount, .. } => Error::Allocation { amount },
        _ => Error::InvalidSource,
    }
}

fn usize_to_u64(value: usize) -> u64 {
    u64::try_from(value).unwrap_or(u64::MAX)
}
