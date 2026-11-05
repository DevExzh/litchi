//! Exact-source transactions for the persisted name of a Keynote slide table.
//!
//! Table names are stored in field 8 of the canonical table-model payload.
//! Selection and authority proof live in [`super::slide_table_core`]; this
//! module only owns the semantic name projection and the field-8 rewrite.

#![allow(
    clippy::map_err_ignore,
    clippy::needless_pass_by_value,
    reason = "the package boundary deliberately redacts native failure detail"
)]

use std::fmt;
use std::sync::Arc;

use litchi_core::Position;
use litchi_iwa_archive::package::{EntryEdit, ExactArtifacts};
use litchi_iwa_core::{Archive, RawMessage, SnappyStream};
use litchi_iwa_protos::table_model_discovery_codec;
use thiserror::Error as ThisError;

use super::slide_table_core as core;
use super::{Package, PayloadLimitKind, PhysicalSource, ReadError, SemanticLimitKind};
use crate::SlideSelector;
use crate::slide::table::{TableSelector, name};

const TABLE_MODEL_MESSAGE_TYPE: u32 = 6_001;

/// Finite resources charged by a slide-table name transaction.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    InputBytes,
    OutputBytes,
    Entries,
    EntryBytes,
    TotalBytes,
    PayloadObjects,
    PayloadMessages,
    References,
    WireFields,
    WireNesting,
    WireWork,
    Allocations,
    Retained,
    Scratch,
    Components,
}

impl fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::InputBytes => "input bytes",
            Self::OutputBytes => "output bytes",
            Self::Entries => "entries",
            Self::EntryBytes => "entry bytes",
            Self::TotalBytes => "total bytes",
            Self::PayloadObjects => "payload objects",
            Self::PayloadMessages => "payload messages",
            Self::References => "references",
            Self::WireFields => "wire fields",
            Self::WireNesting => "wire nesting depth",
            Self::WireWork => "wire work",
            Self::Allocations => "allocations",
            Self::Retained => "retained bytes",
            Self::Scratch => "scratch bytes",
            Self::Components => "components",
        })
    }
}

/// Content-free semantic path for a slide-table name operation.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Path {
    Package,
    Table { slide: Position, table: Position },
}

impl fmt::Display for Path {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Package => formatter.write_str("package"),
            Self::Table { slide, table } => {
                write!(
                    formatter,
                    "slide {} table {} name",
                    slide.get(),
                    table.get()
                )
            },
        }
    }
}

/// Failure from a Keynote slide-table name read or transaction.
#[derive(Debug, Clone, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// A well-formed request this route refuses; [`Refusal::kind`] classifies it
    /// (record 0351 §2.4, record 0361).
    #[error(transparent)]
    Refused(Refusal),
    #[error("the selected Keynote slide-table name source is invalid")]
    InvalidSource,
    #[error("invalid Keynote slide-table name: {0}")]
    InvalidName(#[source] name::Error),
    #[error(
        "Keynote slide-table name {kind} limit exceeded: observed {observed}, maximum {maximum}"
    )]
    Limit {
        kind: LimitKind,
        observed: u64,
        maximum: u64,
    },
    #[error("could not allocate {amount} units for the Keynote slide-table name transaction")]
    Allocation { amount: usize },
    #[error("the edited Keynote slide-table name failed semantic verification")]
    Verification,
    #[error("the Keynote slide-table name patch does not match the exact source package")]
    PatchConflict,
}

/// A well-formed request this family's route refuses, classified by
/// [`Refusal::kind`] (record 0351 §2.4, record 0361).
#[derive(Debug, Clone, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Refusal {
    #[error("this Keynote source does not support physical slide-table name edits")]
    UnsupportedSource,
    #[error("the requested Keynote slide-table name graph is outside the supported scope")]
    UnsupportedDependency,
    #[error("the requested Keynote slide-table name topology is unsupported")]
    UnsupportedTopology,
    #[error("the Keynote slide-table name selector is ambiguous")]
    AmbiguousSelector,
    #[error("the Keynote slide selector name cannot be empty")]
    EmptySlideName,
    #[error("the Keynote show has no slide matching the requested name")]
    SlideNameNotFound,
    #[error("the Keynote show has no slide at position {position:?}")]
    SlidePositionNotFound { position: Position },
    #[error("the selected Keynote slide has no table at position {position:?}")]
    TablePositionNotFound { position: Position },
    #[error("the selected Keynote table is locked")]
    Locked,
}

impl Refusal {
    /// The shared refusal kind litchi-ops classifies this reason by.
    #[must_use]
    pub const fn kind(&self) -> litchi_core::edit::RefusalKind {
        use litchi_core::edit::RefusalKind;
        match self {
            Self::UnsupportedSource { .. }
            | Self::UnsupportedDependency { .. }
            | Self::UnsupportedTopology { .. } => RefusalKind::UnsupportedContent,
            Self::AmbiguousSelector { .. } | Self::EmptySlideName { .. } => {
                RefusalKind::InvalidValue
            },
            Self::SlideNameNotFound { .. }
            | Self::SlidePositionNotFound { .. }
            | Self::TablePositionNotFound { .. } => RefusalKind::NotFound,
            Self::Locked { .. } => RefusalKind::Protected,
        }
    }
}

impl From<Refusal> for Error {
    fn from(refusal: Refusal) -> Self {
        Self::Refused(refusal)
    }
}

/// One mutable name staged against an immutable package snapshot.
pub struct Edit<'a> {
    source: &'a Package,
    selection: NameSelection,
    after: name::Name,
}

impl fmt::Debug for Edit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("path", &self.selection.path())
            .field("before", &self.selection.before)
            .field("after", &self.after)
            .finish_non_exhaustive()
    }
}

impl Edit<'_> {
    #[must_use]
    pub const fn slide_position(&self) -> Position {
        self.selection.target.slide_position
    }

    #[must_use]
    pub const fn table_position(&self) -> Position {
        self.selection.target.table_position
    }

    #[must_use]
    pub const fn path(&self) -> Path {
        self.selection.path()
    }

    #[must_use]
    pub fn before(&self) -> &name::Name {
        &self.selection.before
    }

    #[must_use]
    pub fn name(&self) -> &name::Name {
        &self.after
    }

    #[must_use]
    pub fn after(&self) -> &name::Name {
        &self.after
    }

    #[must_use]
    pub fn set(mut self, value: name::Name) -> Self {
        self.after = value;
        self
    }

    pub fn set_name(self, value: &str) -> Result<Self, Error> {
        Ok(self.set(name::Name::new(value).map_err(Error::InvalidName)?))
    }

    pub fn commit(self) -> Result<Commit, Error> {
        commit_edit(self.source, &self.selection, self.after)
    }
}

/// Exact-source checked reversible slide-table name patch.
#[derive(Clone, PartialEq, Eq)]
pub struct Patch {
    artifacts: ExactArtifacts,
    selection: NameSelection,
    before: name::Name,
    after: name::Name,
    touched_components: usize,
    deleted_previews: usize,
    source_previews_absent: bool,
    target_previews_absent: bool,
}

impl fmt::Debug for Patch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Patch")
            .field("path", &self.selection.path())
            .field("before", &self.before)
            .field("after", &self.after)
            .finish_non_exhaustive()
    }
}

impl Patch {
    #[must_use]
    pub const fn path(&self) -> Path {
        self.selection.path()
    }

    #[must_use]
    pub fn before(&self) -> &name::Name {
        &self.before
    }

    #[must_use]
    pub fn after(&self) -> &name::Name {
        &self.after
    }

    #[must_use]
    pub const fn source_fingerprint(&self) -> u64 {
        self.artifacts.source_fingerprint()
    }

    #[must_use]
    pub const fn target_fingerprint(&self) -> u64 {
        self.artifacts.target_fingerprint()
    }

    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before == self.after && self.artifacts.is_byte_noop()
    }

    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            artifacts: self.artifacts.inverse(),
            selection: self.selection.clone(),
            before: self.after.clone(),
            after: self.before.clone(),
            touched_components: self.touched_components,
            deleted_previews: self.deleted_previews,
            source_previews_absent: self.target_previews_absent,
            target_previews_absent: self.source_previews_absent,
        }
    }
}

/// Compact diagnostics for one published slide-table name transaction.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
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

    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }

    #[must_use]
    pub const fn touched_components(self) -> usize {
        self.touched_components
    }

    #[must_use]
    pub const fn deleted_previews(self) -> usize {
        self.deleted_previews
    }

    #[must_use]
    pub const fn full_reparse_performed(self) -> bool {
        self.full_reparse_performed
    }
}

/// Fully verified result of one slide-table name transaction.
#[must_use = "a Keynote slide-table name commit contains the validated package snapshot"]
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

    /// Consume the commit and return its package and patch.
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
}

#[derive(Clone, PartialEq, Eq)]
struct NameSelection {
    target: core::Target,
    before: name::Name,
    budget: core::Budget,
}

impl fmt::Debug for NameSelection {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("NameSelection")
            .field("path", &self.path())
            .field("before", &self.before)
            .field("locked", &self.target.locked)
            .finish_non_exhaustive()
    }
}

impl NameSelection {
    const fn path(&self) -> Path {
        Path::Table {
            slide: self.target.slide_position,
            table: self.target.table_position,
        }
    }
}

impl Package {
    /// Read one existing slide-table's persisted name.
    pub fn slide_table_name<'slide>(
        &self,
        slide: impl Into<SlideSelector<'slide>>,
        table: impl Into<TableSelector>,
    ) -> Result<name::Name, Error> {
        let mut budget = core::Budget::new(self).map_err(map_core_error)?;
        Ok(select_name(self, slide.into(), table.into(), &mut budget)?.before)
    }

    /// Begin an immutable exact edit of one existing slide-table name.
    pub fn edit_slide_table_name<'slide>(
        &self,
        slide: impl Into<SlideSelector<'slide>>,
        table: impl Into<TableSelector>,
    ) -> Result<Edit<'_>, Error> {
        let mut budget = core::Budget::new(self).map_err(map_core_error)?;
        let selection = select_name(self, slide.into(), table.into(), &mut budget)?;
        let after = selection.before.clone();
        Ok(Edit {
            source: self,
            selection: NameSelection {
                budget,
                ..selection
            },
            after,
        })
    }

    /// Apply an exact-source checked reversible slide-table name patch.
    pub fn apply_slide_table_name(&self, patch: &Patch) -> Result<Commit, Error> {
        let catalog = physical_catalog(self)?;
        if !patch.artifacts.authorizes_source(&catalog.shared_source()) {
            return Err(Error::PatchConflict);
        }
        let mut budget = core::Budget::new(self).map_err(map_core_error)?;
        if previews_absent(self, &mut budget)? != patch.source_previews_absent {
            return Err(Error::PatchConflict);
        }
        let current = select_name(
            self,
            SlideSelector::position(patch.selection.target.slide_position),
            TableSelector::position(patch.selection.target.table_position),
            &mut budget,
        )?;
        if !same_selection(&current, &patch.selection) || current.before != patch.before {
            return Err(Error::PatchConflict);
        }
        if patch.is_noop() {
            self.validate().map_err(map_read_error)?;
            return Ok(Commit {
                package: self.snapshot(),
                patch: patch.clone(),
                diagnostics: Diagnostics::unchanged(),
            });
        }
        reopen_patch(self, patch, budget)
    }
}

fn select_name(
    package: &Package,
    slide: SlideSelector<'_>,
    table: TableSelector,
    budget: &mut core::Budget,
) -> Result<NameSelection, Error> {
    let target = core::select_table(package, slide, table, budget).map_err(map_core_error)?;
    let payload = core::model_payload(package, &target).map_err(map_core_error)?;
    let snapshot = decode_model(payload, package, budget)?;
    let before = owned_name(snapshot.table_name(), budget)?;
    Ok(NameSelection {
        target,
        before,
        budget: *budget,
    })
}

fn commit_edit(
    source: &Package,
    selection: &NameSelection,
    after: name::Name,
) -> Result<Commit, Error> {
    let mut budget = selection.budget;
    budget
        .owned_value(after.as_str().len())
        .map_err(map_core_error)?;
    if selection.before == after {
        let source_previews_absent = previews_absent(source, &mut budget)?;
        let bytes = shared_source_artifact(source, &mut budget)?;
        return Ok(Commit {
            package: source.snapshot(),
            patch: Patch {
                artifacts: ExactArtifacts::new(Arc::clone(&bytes), bytes),
                selection: selection.clone(),
                before: selection.before.clone(),
                after,
                touched_components: 0,
                deleted_previews: 0,
                source_previews_absent,
                target_previews_absent: source_previews_absent,
            },
            diagnostics: Diagnostics::unchanged(),
        });
    }
    if selection.target.locked {
        return Err(Error::Refused(Refusal::Locked));
    }
    let source_previews_absent = previews_absent(source, &mut budget)?;
    let (candidate, deleted_previews) = rewrite_name(source, selection, &after, &mut budget)?;
    if !previews_absent(&candidate, &mut budget)? {
        return Err(Error::Verification);
    }
    let mut reopen_budget = budget;
    let selected = select_name(
        &candidate,
        SlideSelector::position(selection.target.slide_position),
        TableSelector::position(selection.target.table_position),
        &mut reopen_budget,
    )?;
    if !same_selection(&selected, selection) || selected.before != after {
        return Err(Error::Verification);
    }
    core::verify_locality(
        source,
        &candidate,
        &selection.target,
        true,
        &mut reopen_budget,
    )
    .map_err(map_core_error)?;
    let mut artifact_budget = reopen_budget;
    let target = shared_source_artifact(&candidate, &mut artifact_budget)?;
    let source_artifact = shared_source_artifact(source, &mut artifact_budget)?;
    Ok(Commit {
        package: candidate,
        patch: Patch {
            artifacts: ExactArtifacts::new(source_artifact, target),
            selection: selection.clone(),
            before: selection.before.clone(),
            after,
            touched_components: 1,
            deleted_previews,
            source_previews_absent,
            target_previews_absent: true,
        },
        diagnostics: Diagnostics::published(deleted_previews),
    })
}

fn reopen_patch(
    source: &Package,
    patch: &Patch,
    mut budget: core::Budget,
) -> Result<Commit, Error> {
    let target_source = patch.artifacts.target();
    let candidate = parse_candidate(Arc::clone(&target_source), source, &mut budget)?;
    if previews_absent(&candidate, &mut budget)? != patch.target_previews_absent {
        return Err(Error::Verification);
    }
    let selected = select_name(
        &candidate,
        SlideSelector::position(patch.selection.target.slide_position),
        TableSelector::position(patch.selection.target.table_position),
        &mut budget,
    )?;
    if !same_selection(&selected, &patch.selection) || selected.before != patch.after {
        return Err(Error::Verification);
    }
    core::verify_locality(
        source,
        &candidate,
        &patch.selection.target,
        patch.target_previews_absent,
        &mut budget,
    )
    .map_err(map_core_error)?;
    Ok(Commit {
        package: candidate,
        patch: patch.clone(),
        diagnostics: Diagnostics::published(patch.deleted_previews),
    })
}

fn rewrite_name(
    source: &Package,
    selection: &NameSelection,
    after: &name::Name,
    budget: &mut core::Budget,
) -> Result<(Package, usize), Error> {
    let catalog = physical_catalog(source)?;
    let entry = catalog
        .package()
        .iter()
        .find(|entry| entry.name() == selection.target.model.component.as_ref())
        .ok_or(Error::InvalidSource)?;
    if entry.is_opaque() {
        return Err(Error::Refused(Refusal::UnsupportedSource));
    }
    let physical_limits = source.state.options.archive();
    let archive_limits = physical_limits
        .effective_archive_limits()
        .map_err(map_archive_error)?;
    let snappy_limits = physical_limits.snappy_limits().map_err(map_archive_error)?;
    budget.allocations(1).map_err(map_core_error)?;
    budget
        .retained(entry.data().len())
        .map_err(map_core_error)?;
    budget.scratch(entry.data().len()).map_err(map_core_error)?;
    budget
        .physical(entry.data().len())
        .map_err(map_core_error)?;
    let stream =
        SnappyStream::decompress(entry.data(), snappy_limits).map_err(map_core_native_error)?;
    budget.allocations(1).map_err(map_core_error)?;
    budget
        .retained(stream.as_bytes().len())
        .map_err(map_core_error)?;
    budget
        .scratch(stream.as_bytes().len())
        .map_err(map_core_error)?;
    budget
        .physical(stream.as_bytes().len())
        .map_err(map_core_error)?;
    budget
        .allocations(stream.as_bytes().len())
        .map_err(map_core_error)?;
    budget
        .retained(stream.as_bytes().len())
        .map_err(map_core_error)?;
    budget
        .scratch(stream.as_bytes().len())
        .map_err(map_core_error)?;
    let mut archive =
        Archive::parse(stream.as_bytes(), archive_limits).map_err(map_core_native_error)?;
    let original_source = archive
        .object(selection.target.model.identifier)
        .ok_or(Error::InvalidSource)?
        .messages
        .get(selection.target.model_message_index)
        .filter(|message| message.type_ == TABLE_MODEL_MESSAGE_TYPE)
        .ok_or(Error::InvalidSource)?
        .data
        .as_slice();
    budget
        .allocations(original_source.len())
        .map_err(map_core_error)?;
    budget
        .retained(original_source.len())
        .map_err(map_core_error)?;
    let mut original = Vec::new();
    original
        .try_reserve_exact(original_source.len())
        .map_err(|_| Error::Allocation {
            amount: original_source.len(),
        })?;
    original.extend_from_slice(original_source);
    let original_snapshot = decode_model(&original, source, budget)?;
    if original_snapshot.table_name() != selection.before.as_str() {
        return Err(Error::InvalidSource);
    }
    let fingerprint = table_model_discovery_codec::table_model_source_fingerprint(&original);
    let options = budget
        .model_codec_options(source, &original)
        .map_err(map_core_error)?
        .with_max_allocations(budget.remaining_allocations().map_err(map_core_error)?)
        .with_max_retained_bytes(budget.remaining_retained().map_err(map_core_error)?)
        .with_max_scratch_bytes(budget.remaining_scratch().map_err(map_core_error)?);
    let prepared = table_model_discovery_codec::prepare_table_model_name_rewrite(
        &original,
        table_model_discovery_codec::TableModelNameWrite::new(after.as_str())
            .with_fingerprint(fingerprint),
        options,
    )
    .map_err(map_codec_error)?;
    let requirements = prepared.execution_requirements();
    budget
        .name_rewrite_requirements(requirements)
        .map_err(map_core_error)?;
    let rewritten = prepared
        .execute(requirements.exact())
        .map_err(map_codec_error)?
        .into_bytes();
    let verified = decode_model(&rewritten, source, budget)?;
    if verified.table_name() != after.as_str()
        || verified.table_id() != original_snapshot.table_id()
        || verified.rows() != original_snapshot.rows()
        || verified.columns() != original_snapshot.columns()
    {
        return Err(Error::Verification);
    }
    archive
        .object_mut(selection.target.model.identifier)
        .ok_or(Error::InvalidSource)?
        .replace_message_preserving_header(
            selection.target.model_message_index,
            RawMessage {
                type_: TABLE_MODEL_MESSAGE_TYPE,
                data: rewritten,
            },
            archive_limits,
        )
        .map_err(map_core_native_error)?;
    let encoded_bound = archive
        .encoded_len(archive_limits)
        .map_err(map_core_native_error)?;
    let compressed_bound = SnappyStream::maximum_compressed_len(encoded_bound, snappy_limits)
        .map_err(map_core_native_error)?;
    budget
        .output(
            encoded_bound
                .checked_add(compressed_bound)
                .ok_or(core::Error::InvalidSource)
                .map_err(map_core_error)?,
        )
        .map_err(map_core_error)?;
    budget.allocations(encoded_bound).map_err(map_core_error)?;
    budget.retained(encoded_bound).map_err(map_core_error)?;
    budget
        .allocations(compressed_bound)
        .map_err(map_core_error)?;
    budget.retained(compressed_bound).map_err(map_core_error)?;
    let bytes = archive
        .to_bytes(archive_limits)
        .map_err(map_core_native_error)?;
    if bytes.len() != encoded_bound {
        return Err(Error::Verification);
    }
    let compressed =
        SnappyStream::compress(&bytes, snappy_limits).map_err(map_core_native_error)?;
    if compressed.len() > compressed_bound {
        return Err(Error::Verification);
    }
    core::charge_preview_scan(source, budget).map_err(map_core_error)?;
    let previews = super::rendering_invalidation::root_preview_deletions(catalog.package())
        .map_err(|_| Error::InvalidSource)?;
    let edits = [EntryEdit::new(
        selection.target.model.component.as_ref(),
        &compressed,
    )];
    charge_reassembly_prepare(catalog, compressed.len(), previews.names(), budget)?;
    let prepared = catalog
        .prepare_reassembly_with_deletions(&edits, previews.names(), physical_limits)
        .map_err(map_archive_error)?;
    let requirements = prepared.execution_requirements();
    budget.reassembly(requirements).map_err(map_core_error)?;
    let output = prepared
        .execute(requirements.exact_limits())
        .map_err(map_archive_error)?;
    let candidate = parse_candidate(output.into(), source, budget)?;
    Ok((candidate, previews.len()))
}

fn owned_name(value: &str, budget: &mut core::Budget) -> Result<name::Name, Error> {
    budget.owned_value(value.len()).map_err(map_core_error)?;
    name::Name::new(value).map_err(Error::InvalidName)
}

fn shared_source_artifact(
    package: &Package,
    budget: &mut core::Budget,
) -> Result<Arc<[u8]>, Error> {
    let source = physical_catalog(package)?;
    let bytes = source.package().source_bytes().len();
    budget.artifact(bytes, 0).map_err(map_core_error)?;
    Ok(source.shared_source())
}

fn parse_candidate(
    source: Arc<[u8]>,
    original: &Package,
    budget: &mut core::Budget,
) -> Result<Package, Error> {
    let source_bytes = source.len();
    // SourceCatalog construction indexes the ZIP and copies each decoded IWA
    // payload; validation can then materialize the lazy semantic show. Both
    // phases happen before the candidate is published, so precharge their
    // conservative source-sized allocation/retention envelope together.
    budget
        .preflight_candidate(original, source_bytes)
        .map_err(map_core_error)?;
    budget
        .preflight_semantic_scan(original)
        .map_err(map_core_error)?;
    let candidate = Package::from_source_with_options(source, original.state.options)
        .map_err(map_read_error)?;
    candidate.validate().map_err(map_read_error)?;
    Ok(candidate)
}

fn charge_reassembly_prepare(
    catalog: &litchi_iwa_archive::SourceCatalog,
    edited_bytes: usize,
    deleted_names: &[&str],
    budget: &mut core::Budget,
) -> Result<(), Error> {
    budget
        .preflight_reassembly(catalog, edited_bytes, deleted_names.len())
        .map_err(map_core_error)
}

fn decode_model<'source>(
    payload: &'source [u8],
    package: &Package,
    budget: &mut core::Budget,
) -> Result<table_model_discovery_codec::TableModelSnapshot<'source>, Error> {
    let options = budget
        .model_codec_options(package, payload)
        .map_err(map_core_error)?;
    let (snapshot, report) =
        table_model_discovery_codec::decode_table_model_with_report(payload, options)
            .map_err(map_codec_error)?;
    budget.codec_report(report).map_err(map_core_error)?;
    Ok(snapshot)
}

fn same_selection(left: &NameSelection, right: &NameSelection) -> bool {
    core::same_target(&left.target, &right.target)
}

fn physical_catalog(package: &Package) -> Result<&litchi_iwa_archive::SourceCatalog, Error> {
    match &package.state.source {
        PhysicalSource::Package(source) => Ok(source),
        PhysicalSource::Semantic(_) => Err(Error::Refused(Refusal::UnsupportedSource)),
    }
}

fn previews_absent(package: &Package, budget: &mut core::Budget) -> Result<bool, Error> {
    core::charge_preview_scan(package, budget).map_err(map_core_error)?;
    super::rendering_invalidation::root_previews_absent(physical_catalog(package)?.package())
        .map_err(|_| Error::Verification)
}

fn map_core_error(error: core::Error) -> Error {
    match error {
        core::Error::UnsupportedSource => Error::Refused(Refusal::UnsupportedSource),
        core::Error::UnsupportedDependency => Error::Refused(Refusal::UnsupportedDependency),
        core::Error::UnsupportedTopology => Error::Refused(Refusal::UnsupportedTopology),
        core::Error::AmbiguousSelector => Error::Refused(Refusal::AmbiguousSelector),
        core::Error::EmptySlideName => Error::Refused(Refusal::EmptySlideName),
        core::Error::SlideNameNotFound => Error::Refused(Refusal::SlideNameNotFound),
        core::Error::SlidePositionNotFound(position) => {
            Error::Refused(Refusal::SlidePositionNotFound { position })
        },
        core::Error::TablePositionNotFound(position) => {
            Error::Refused(Refusal::TablePositionNotFound { position })
        },
        core::Error::Limit {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: map_limit_kind(kind),
            observed,
            maximum,
        },
        core::Error::Allocation(amount) => Error::Allocation { amount },
        core::Error::InvalidSource
        | core::Error::Read
        | core::Error::Wire
        | core::Error::Codec
        | core::Error::Archive
        | core::Error::Verification => Error::InvalidSource,
    }
}

fn map_limit_kind(kind: core::LimitKind) -> LimitKind {
    match kind {
        core::LimitKind::InputBytes => LimitKind::InputBytes,
        core::LimitKind::OutputBytes => LimitKind::OutputBytes,
        core::LimitKind::Entries => LimitKind::Entries,
        core::LimitKind::EntryBytes => LimitKind::EntryBytes,
        core::LimitKind::TotalBytes => LimitKind::TotalBytes,
        core::LimitKind::PayloadObjects => LimitKind::PayloadObjects,
        core::LimitKind::PayloadMessages => LimitKind::PayloadMessages,
        core::LimitKind::References => LimitKind::References,
        core::LimitKind::WireFields => LimitKind::WireFields,
        core::LimitKind::WireNesting => LimitKind::WireNesting,
        core::LimitKind::WireWork => LimitKind::WireWork,
        core::LimitKind::Allocations => LimitKind::Allocations,
        core::LimitKind::Retained => LimitKind::Retained,
        core::LimitKind::Scratch => LimitKind::Scratch,
        core::LimitKind::Components => LimitKind::Components,
    }
}

fn map_read_error(error: ReadError) -> Error {
    match error {
        ReadError::SemanticLimit {
            kind,
            observed,
            maximum,
            ..
        } => Error::Limit {
            kind: match kind {
                SemanticLimitKind::References => LimitKind::References,
                _ => LimitKind::WireWork,
            },
            observed: observed as u64,
            maximum: maximum as u64,
        },
        ReadError::PayloadLimit {
            kind,
            observed,
            maximum,
            ..
        } => Error::Limit {
            kind: match kind {
                PayloadLimitKind::Bytes => LimitKind::InputBytes,
                PayloadLimitKind::Fields => LimitKind::WireFields,
                PayloadLimitKind::Nesting => LimitKind::WireNesting,
                PayloadLimitKind::Work => LimitKind::WireWork,
            },
            observed: observed as u64,
            maximum: maximum as u64,
        },
        ReadError::Allocation { amount, .. } => Error::Allocation { amount },
        ReadError::Archive(error) => map_archive_error(error),
        _ => Error::InvalidSource,
    }
}

fn map_codec_error(error: table_model_discovery_codec::DecodeError) -> Error {
    let Some(limit) = error.resource_limit() else {
        return Error::InvalidSource;
    };
    let (kind, observed, maximum) = match limit {
        table_model_discovery_codec::DecodeLimit::Bytes { observed, maximum } => {
            (LimitKind::InputBytes, observed as u64, maximum as u64)
        },
        table_model_discovery_codec::DecodeLimit::Fields { observed, maximum } => {
            (LimitKind::WireFields, observed as u64, maximum as u64)
        },
        table_model_discovery_codec::DecodeLimit::Work { observed, maximum } => {
            (LimitKind::WireWork, observed as u64, maximum as u64)
        },
        table_model_discovery_codec::DecodeLimit::Text { observed, maximum } => {
            (LimitKind::WireWork, observed as u64, maximum as u64)
        },
        table_model_discovery_codec::DecodeLimit::Output { observed, maximum } => {
            (LimitKind::OutputBytes, observed as u64, maximum as u64)
        },
        table_model_discovery_codec::DecodeLimit::Allocations { observed, maximum } => {
            (LimitKind::Allocations, observed as u64, maximum as u64)
        },
        table_model_discovery_codec::DecodeLimit::Retained { observed, maximum } => {
            (LimitKind::Retained, observed as u64, maximum as u64)
        },
        table_model_discovery_codec::DecodeLimit::Scratch { observed, maximum } => {
            (LimitKind::Scratch, observed as u64, maximum as u64)
        },
        table_model_discovery_codec::DecodeLimit::Nesting { observed, maximum } => (
            LimitKind::WireNesting,
            u64::from(observed),
            u64::from(maximum),
        ),
        _ => return Error::InvalidSource,
    };
    Error::Limit {
        kind,
        observed,
        maximum,
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
                litchi_iwa_archive::LimitKind::EntryBytes
                | litchi_iwa_archive::LimitKind::CompressedEntryBytes => LimitKind::EntryBytes,
                litchi_iwa_archive::LimitKind::TotalBytes => LimitKind::TotalBytes,
                _ => LimitKind::WireWork,
            },
            observed,
            maximum,
        },
        litchi_iwa_archive::Error::Allocation { amount, .. } => Error::Allocation { amount },
        litchi_iwa_archive::Error::Iwa(error) => map_core_native_error(error),
        _ => Error::InvalidSource,
    }
}

fn map_core_native_error(error: litchi_iwa_core::Error) -> Error {
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
                litchi_iwa_core::LimitKind::HeaderNesting => LimitKind::WireNesting,
                litchi_iwa_core::LimitKind::SnappyCompressedChunkBytes
                | litchi_iwa_core::LimitKind::SnappyCompressedStreamBytes
                | litchi_iwa_core::LimitKind::SnappyFrames => LimitKind::EntryBytes,
                _ => LimitKind::WireWork,
            },
            observed: observed as u64,
            maximum: maximum as u64,
        },
        litchi_iwa_core::Error::Allocation { requested, .. } => {
            Error::Allocation { amount: requested }
        },
        _ => Error::InvalidSource,
    }
}
