//! Selector-first, source-preserving table appearance transactions.
//!
//! This module owns the rooted Numbers table selection and publication
//! boundary.  The table-style wire projection is deliberately kept behind the
//! hidden protos codec; native object identifiers and generated protobuf
//! values never cross the public `table::appearance` module.

use std::fmt;

use litchi_iwa_archive::package::{EntryEdit, OwnedExactArtifacts};
use litchi_iwa_common::WireLimits;
use litchi_iwa_core::archive::{FieldObjectReferenceTransition, ObjectReferenceTransition};
use litchi_iwa_core::{Archive, ArchiveObject, FieldType, RawMessage, SnappyStream};
use litchi_iwa_protos::package_metadata_codec::{
    AdditionSaveTokenBatch, Batch as MetadataBatch, ComponentSelector, ExternalReferenceAddition,
    ObjectUuidAddition, PackageMetadataVisitor, RewriteError as MetadataRewriteError,
    RewriteOptions as MetadataRewriteOptions, SaveTokenBatch, UuidBits,
    inspect_package_metadata_with_visitor, prepare_package_metadata_additions_and_save_tokens,
};
use litchi_iwa_protos::table_appearance_codec as codec;
use thiserror::Error as ThisError;

use super::Package;
use crate::{
    selector::{SheetSelector, TableSelector},
    table::{appearance::Appearance, lock::State as LockState},
};

const TABLE_STYLE_MESSAGE_TYPE: u32 = 6_003;
const STYLESHEET_MESSAGE_TYPE: u32 = 401;
const MAX_INHERITANCE_DEPTH: usize = 64;

/// A native payload requested by the migration-host source-built appearance
/// bridge.
///
/// This type is hidden from the supported Numbers API.  It keeps the bridge
/// selector-free and lets the host retain ownership of its parsed archive
/// cache while this crate owns all appearance interpretation.
#[cfg(feature = "internal-iwork-source")]
#[doc(hidden)]
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SourceBuiltAppearancePayload {
    /// A `TST.TableStyleArchive` payload.
    TableStyle,
    /// A `TST.TableStylePresetArchive` payload.
    TableStylePreset,
    /// A `TST.TableStyleNetworkArchive` payload.
    TableStyleNetwork,
}

#[cfg(feature = "internal-iwork-source")]
const SOURCE_BUILT_MAX_PAYLOAD_BYTES: usize = 8 * 1024 * 1024;
#[cfg(feature = "internal-iwork-source")]
const SOURCE_BUILT_MAX_FIELDS: usize = 1 << 20;
#[cfg(feature = "internal-iwork-source")]
const SOURCE_BUILT_MAX_WORK_BYTES: usize = 64 * 1024 * 1024;
#[cfg(feature = "internal-iwork-source")]
const SOURCE_BUILT_MAX_ALLOCATIONS: usize = 256;

/// A content-free location associated with a table-appearance operation.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Path {
    /// The complete Numbers package.
    Package,
    /// One rooted table at checked zero-based positions.
    Table { sheet: usize, table: usize },
}

/// A finite resource governed by a table-appearance operation.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Complete source package bytes.
    InputBytes,
    /// Complete candidate package bytes.
    OutputBytes,
    /// Physical package members.
    Entries,
    /// Bytes in one physical member.
    EntryBytes,
    /// Aggregate physical member bytes.
    TotalEntryBytes,
    /// Physical package names and container metadata.
    PackageBytes,
    /// Bytes in one decoded native payload.
    PayloadBytes,
    /// Aggregate decoded payload bytes.
    TotalPayloadBytes,
    /// Native objects inspected.
    PayloadObjects,
    /// Native messages inspected.
    PayloadMessages,
    /// Native framing or metadata items inspected.
    PayloadItems,
    /// Native object references inspected.
    PayloadReferences,
    /// Strict wire input bytes.
    WireBytes,
    /// Strict wire output bytes.
    WireOutputBytes,
    /// Strict wire fields.
    WireFields,
    /// Strict wire nesting.
    WireNesting,
    /// Strict wire work.
    WireWork,
    /// Aggregate focused transaction work.
    TransactionWork,
}

impl fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(formatter, "{self:?}")
    }
}

/// A content-redacted table-appearance failure.
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// No rooted sheet matched the selector.
    #[error("the Numbers workbook has no sheet matching the requested selector")]
    SheetNotFound,
    /// No table on the selected sheet matched the selector.
    #[error("the selected Numbers sheet has no table matching the requested selector")]
    TableNotFound,
    /// A changed operation targeted a locked table.
    #[error("the selected Numbers table is locked at {path:?}")]
    TableLocked { path: Path },
    /// A native style dependency is not safe to rewrite in this owner.
    #[error("the selected Numbers table appearance has an unsupported dependency at {path:?}")]
    UnsupportedDependency { path: Path },
    /// The source is not an exact supported native profile.
    #[error("this Numbers source does not support exact table-appearance editing")]
    UnsupportedSource,
    /// Rooted ownership or wire framing is invalid.
    #[error("the Numbers table-appearance source is invalid at {path:?}")]
    InvalidSource { path: Path },
    /// A finite resource ceiling was exceeded.
    #[error(
        "Numbers table-appearance {kind} limit exceeded: observed {observed}, maximum {maximum}"
    )]
    LimitExceeded {
        kind: LimitKind,
        observed: u64,
        maximum: u64,
        path: Path,
    },
    /// A bounded allocation failed before publication.
    #[error("could not allocate {amount} units for the Numbers table-appearance transaction")]
    Allocation { amount: usize, path: Path },
    /// Candidate reopening or locality verification failed.
    #[error("the edited Numbers table appearance failed semantic verification")]
    Verification,
    /// A patch was applied to a package other than its exact source.
    #[error("the table-appearance patch does not match the exact source package")]
    PatchConflict,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct Target {
    native: super::table_headers::Target,
    appearance: Appearance,
    style_identifier: Option<u64>,
}

/// Immutable appearance settings staged against one package snapshot.
pub struct Edit<'a> {
    source: &'a Package,
    target: Target,
    before: Appearance,
    appearance: Appearance,
    budget: TransactionBudget,
}

impl fmt::Debug for Edit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("path", &self.path())
            .field("appearance", &self.appearance)
            .finish_non_exhaustive()
    }
}

impl Edit<'_> {
    /// Return the selected table path.
    #[must_use]
    pub const fn path(&self) -> Path {
        Path::Table {
            sheet: self.target.native.sheet_position,
            table: self.target.native.table_position,
        }
    }

    /// Return the staged appearance.
    #[must_use]
    pub const fn appearance(&self) -> Appearance {
        self.appearance
    }

    /// Replace the staged appearance without touching package bytes.
    #[must_use]
    pub fn set(mut self, appearance: Appearance) -> Self {
        self.appearance = appearance;
        self
    }

    /// Validate and atomically publish the staged appearance.
    pub fn commit(self) -> Result<Commit, Error> {
        commit_edit(self)
    }
}

/// A reversible process-local exact-source appearance patch.
#[derive(Clone, PartialEq, Eq)]
pub struct Patch {
    artifacts: OwnedExactArtifacts,
    target: Target,
    before: Appearance,
    after: Appearance,
    source_previews: usize,
    target_previews: usize,
    touched_components: usize,
}

impl fmt::Debug for Patch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Patch")
            .field("path", &self.path())
            .field("before", &self.before)
            .field("after", &self.after)
            .finish_non_exhaustive()
    }
}

impl Patch {
    /// Return the selected table path.
    #[must_use]
    pub const fn path(&self) -> Path {
        Path::Table {
            sheet: self.target.native.sheet_position,
            table: self.target.native.table_position,
        }
    }

    /// Return exact source appearance settings.
    #[must_use]
    pub const fn before(&self) -> Appearance {
        self.before
    }

    /// Return exact target appearance settings.
    #[must_use]
    pub const fn after(&self) -> Appearance {
        self.after
    }

    /// Return the source diagnostic fingerprint.
    #[must_use]
    pub const fn source_fingerprint(&self) -> u64 {
        self.artifacts.source_fingerprint()
    }

    /// Return the target diagnostic fingerprint.
    #[must_use]
    pub const fn target_fingerprint(&self) -> u64 {
        self.artifacts.target_fingerprint()
    }

    /// Return whether this patch is an exact byte no-op.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before == self.after && self.artifacts.is_byte_noop()
    }

    /// Return the exact target-to-source inverse.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            artifacts: self.artifacts.inverse(),
            target: self.target,
            before: self.after,
            after: self.before,
            source_previews: self.target_previews,
            target_previews: self.source_previews,
            touched_components: self.touched_components,
        }
    }
}

/// Content-free publication diagnostics.
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

    const fn published(deleted_previews: usize, touched_components: usize) -> Self {
        Self {
            changed: true,
            touched_components,
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

    /// Number of canonical root previews deleted in this direction.
    #[must_use]
    pub const fn deleted_previews(self) -> usize {
        self.deleted_previews
    }

    /// Whether a complete candidate package was reopened.
    #[must_use]
    pub const fn full_reparse_performed(self) -> bool {
        self.full_reparse_performed
    }
}

/// One fully validated immutable publication.
#[must_use = "a table-appearance commit contains the validated package snapshot"]
#[derive(Debug)]
pub struct Commit {
    package: Package,
    patch: Patch,
    diagnostics: Diagnostics,
}

impl Commit {
    /// Borrow the validated package snapshot.
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.package
    }

    /// Consume the publication and return its package snapshot.
    #[must_use]
    pub fn into_package(self) -> Package {
        self.package
    }

    /// Borrow the reversible patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Borrow content-free diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> &Diagnostics {
        &self.diagnostics
    }
}

#[cfg(feature = "internal-iwork-source")]
#[derive(Debug, Clone, Copy)]
struct SourceBuiltAppearanceBudget {
    remaining_input_bytes: usize,
    remaining_fields: usize,
    remaining_work: usize,
    remaining_styles: usize,
    remaining_allocations: usize,
}

#[cfg(feature = "internal-iwork-source")]
impl SourceBuiltAppearanceBudget {
    const fn new() -> Self {
        Self {
            remaining_input_bytes: SOURCE_BUILT_MAX_PAYLOAD_BYTES,
            remaining_fields: SOURCE_BUILT_MAX_FIELDS,
            remaining_work: SOURCE_BUILT_MAX_WORK_BYTES,
            remaining_styles: MAX_INHERITANCE_DEPTH,
            remaining_allocations: SOURCE_BUILT_MAX_ALLOCATIONS,
        }
    }

    fn codec_options(&self, source: &[u8]) -> codec::DecodeOptions {
        codec::DecodeOptions::new(
            self.remaining_input_bytes.min(source.len().max(1)),
            self.remaining_input_bytes.max(1),
            self.remaining_fields.max(1),
            self.remaining_work.max(1),
            u32::try_from(WireLimits::MAX_NESTING).unwrap_or(u32::MAX),
            self.remaining_styles.max(1),
        )
        .with_max_allocations(self.remaining_allocations.max(1))
    }

    fn charge_input(&mut self, amount: usize, path: Path) -> Result<(), Error> {
        charge_source_built_budget(
            &mut self.remaining_input_bytes,
            SOURCE_BUILT_MAX_PAYLOAD_BYTES,
            amount,
            LimitKind::WireBytes,
            path,
        )
    }

    fn charge_fields(&mut self, amount: usize, path: Path) -> Result<(), Error> {
        charge_source_built_budget(
            &mut self.remaining_fields,
            SOURCE_BUILT_MAX_FIELDS,
            amount,
            LimitKind::WireFields,
            path,
        )
    }

    fn charge_work(&mut self, amount: usize, path: Path) -> Result<(), Error> {
        charge_source_built_budget(
            &mut self.remaining_work,
            SOURCE_BUILT_MAX_WORK_BYTES,
            amount,
            LimitKind::WireWork,
            path,
        )
    }

    fn charge_style(&mut self, path: Path) -> Result<(), Error> {
        charge_source_built_budget(
            &mut self.remaining_styles,
            MAX_INHERITANCE_DEPTH,
            1,
            LimitKind::PayloadItems,
            path,
        )
    }

    fn consume_report(&mut self, report: codec::DecodeReport, path: Path) -> Result<(), Error> {
        self.charge_input(report.input_bytes(), path)?;
        self.charge_fields(report.fields(), path)?;
        self.charge_work(report.work_bytes(), path)?;
        charge_source_built_budget(
            &mut self.remaining_allocations,
            SOURCE_BUILT_MAX_ALLOCATIONS,
            report.allocations(),
            LimitKind::TransactionWork,
            path,
        )
    }
}

#[cfg(feature = "internal-iwork-source")]
fn charge_source_built_budget(
    remaining: &mut usize,
    maximum: usize,
    amount: usize,
    kind: LimitKind,
    path: Path,
) -> Result<(), Error> {
    if amount > *remaining {
        let observed = maximum.saturating_sub(*remaining).saturating_add(amount);
        return Err(Error::LimitExceeded {
            kind,
            observed: u64::try_from(observed).unwrap_or(u64::MAX),
            maximum: u64::try_from(maximum).unwrap_or(u64::MAX),
            path,
        });
    }
    *remaining -= amount;
    Ok(())
}

#[cfg(feature = "internal-iwork-source")]
fn merge_source_built_overrides(
    target: &mut codec::AppearanceOverrides,
    source: codec::AppearanceOverrides,
) {
    if target.row_banding.is_none() {
        target.row_banding = source.row_banding;
    }
    if target.row_sizing.is_none() {
        target.row_sizing = source.row_sizing;
    }
    if target.body_horizontal.is_none() {
        target.body_horizontal = source.body_horizontal;
    }
    if target.body_vertical.is_none() {
        target.body_vertical = source.body_vertical;
    }
    if target.header_columns_horizontal.is_none() {
        target.header_columns_horizontal = source.header_columns_horizontal;
    }
    if target.header_rows_vertical.is_none() {
        target.header_rows_vertical = source.header_rows_vertical;
    }
    if target.footer_rows_vertical.is_none() {
        target.footer_rows_vertical = source.footer_rows_vertical;
    }
}

#[cfg(feature = "internal-iwork-source")]
const fn source_built_overrides_complete(overrides: codec::AppearanceOverrides) -> bool {
    overrides.row_banding.is_some()
        && overrides.row_sizing.is_some()
        && overrides.body_horizontal.is_some()
        && overrides.body_vertical.is_some()
        && overrides.header_columns_horizontal.is_some()
        && overrides.header_rows_vertical.is_some()
        && overrides.footer_rows_vertical.is_some()
}

impl Package {
    /// Read one rooted table's effective appearance.
    pub fn table_appearance<'sheet, 'table>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        table: impl Into<TableSelector<'table>>,
    ) -> Result<Appearance, Error> {
        let mut budget = TransactionBudget::new(self);
        let result = resolve_target_with_budget(self, sheet, table, &mut budget, true);
        Ok(result?.appearance)
    }

    /// Read a source-built table appearance through a host-owned payload
    /// lookup.
    ///
    /// This is an unstable migration bridge for the legacy iWork coordinator.
    /// The coordinator keeps its parsed archive cache and invokes `lookup`
    /// only for the selected style graph; this owner performs all bounded wire
    /// decoding, preset/network resolution, inheritance, and defaulting. The
    /// callback receives a consumer instead of a returned byte slice so no
    /// native payload needs to be cloned or retained between style hops.
    #[cfg(feature = "internal-iwork-source")]
    #[doc(hidden)]
    pub fn __table_appearance_from_source_built<LookupError>(
        model_style_identifier: u64,
        model_style_preset_identifier: Option<u64>,
        mut lookup: impl FnMut(
            u64,
            SourceBuiltAppearancePayload,
            &mut dyn FnMut(&[u8]),
        ) -> Result<(), LookupError>,
    ) -> Result<Appearance, Error> {
        let path = Path::Package;
        let mut budget = SourceBuiltAppearanceBudget::new();

        // Source-built models created by the historical editor may omit both
        // style edges.  Preserve the established native defaults without
        // routing that compatibility graph through the exact-source owner.
        let style_identifier = if model_style_identifier != 0 {
            model_style_identifier
        } else {
            let Some(preset_identifier) =
                model_style_preset_identifier.filter(|identifier| *identifier != 0)
            else {
                return Ok(Appearance::default());
            };

            let mut decoded_preset = None;
            let mut consume = |payload: &[u8]| {
                decoded_preset = Some((|| {
                    let (preset, report) = codec::decode_table_style_preset_with_report(
                        payload,
                        budget.codec_options(payload),
                    )
                    .map_err(|error| map_codec_error(error, path))?;
                    budget.consume_report(report, path)?;
                    Ok::<_, Error>(preset.style_network_identifier())
                })());
            };
            lookup(
                preset_identifier,
                SourceBuiltAppearancePayload::TableStylePreset,
                &mut consume,
            )
            .map_err(|_| Error::InvalidSource { path })?;
            let network_identifier = decoded_preset
                .transpose()?
                .flatten()
                .filter(|identifier| *identifier != 0)
                .ok_or(Error::InvalidSource { path })?;

            let mut decoded_network = None;
            let mut consume = |payload: &[u8]| {
                decoded_network = Some((|| {
                    let (network, report) = codec::decode_table_style_network_with_report(
                        payload,
                        budget.codec_options(payload),
                    )
                    .map_err(|error| map_codec_error(error, path))?;
                    budget.consume_report(report, path)?;
                    Ok::<_, Error>(network.table_style_identifier())
                })());
            };
            lookup(
                network_identifier,
                SourceBuiltAppearancePayload::TableStyleNetwork,
                &mut consume,
            )
            .map_err(|_| Error::InvalidSource { path })?;
            decoded_network
                .transpose()?
                .filter(|identifier| *identifier != 0)
                .ok_or(Error::InvalidSource { path })?
        };

        let mut overrides = codec::AppearanceOverrides::default();
        let mut visited = [0_u64; MAX_INHERITANCE_DEPTH];
        let mut current = Some(style_identifier);
        for visited_len in 0..MAX_INHERITANCE_DEPTH {
            let Some(identifier) = current else {
                break;
            };
            if visited[..visited_len].contains(&identifier) {
                return Err(Error::InvalidSource { path });
            }
            visited[visited_len] = identifier;
            budget.charge_style(path)?;

            let mut decoded_style = None;
            let mut consume = |payload: &[u8]| {
                decoded_style = Some((|| {
                    let (snapshot, report) = codec::decode_table_style_with_report(
                        payload,
                        budget.codec_options(payload),
                    )
                    .map_err(|error| map_codec_error(error, path))?;
                    budget.consume_report(report, path)?;
                    Ok::<_, Error>((snapshot.parent_identifier(), snapshot.overrides()))
                })());
            };
            lookup(
                identifier,
                SourceBuiltAppearancePayload::TableStyle,
                &mut consume,
            )
            .map_err(|_| Error::InvalidSource { path })?;
            let (parent_identifier, direct_overrides) = decoded_style
                .transpose()?
                .ok_or(Error::InvalidSource { path })?;
            merge_source_built_overrides(&mut overrides, direct_overrides);
            if source_built_overrides_complete(overrides) {
                return Ok(snapshot_appearance(codec::AppearanceSnapshot {
                    row_banding: overrides.row_banding.unwrap_or(false),
                    row_sizing: overrides.row_sizing.unwrap_or(false),
                    body_horizontal: overrides.body_horizontal.unwrap_or(true),
                    body_vertical: overrides.body_vertical.unwrap_or(true),
                    header_columns_horizontal: overrides.header_columns_horizontal.unwrap_or(true),
                    header_rows_vertical: overrides.header_rows_vertical.unwrap_or(true),
                    footer_rows_vertical: overrides.footer_rows_vertical.unwrap_or(true),
                }));
            }
            current = parent_identifier.filter(|identifier| *identifier != 0);
        }
        if current.is_some() {
            return Err(Error::InvalidSource { path });
        }
        Ok(snapshot_appearance(codec::AppearanceSnapshot {
            row_banding: overrides.row_banding.unwrap_or(false),
            row_sizing: overrides.row_sizing.unwrap_or(false),
            body_horizontal: overrides.body_horizontal.unwrap_or(true),
            body_vertical: overrides.body_vertical.unwrap_or(true),
            header_columns_horizontal: overrides.header_columns_horizontal.unwrap_or(true),
            header_rows_vertical: overrides.header_rows_vertical.unwrap_or(true),
            footer_rows_vertical: overrides.footer_rows_vertical.unwrap_or(true),
        }))
    }

    /// Start a selector-first immutable table-appearance edit.
    pub fn edit_table_appearance<'sheet, 'table>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        table: impl Into<TableSelector<'table>>,
    ) -> Result<Edit<'_>, Error> {
        let mut budget = TransactionBudget::new(self);
        let target = resolve_target_with_budget(self, sheet, table, &mut budget, false)?;
        Ok(Edit {
            source: self,
            before: target.appearance,
            appearance: target.appearance,
            target,
            budget,
        })
    }

    /// Apply a reversible exact-source appearance patch.
    pub fn apply_table_appearance(&self, patch: &Patch) -> Result<Commit, Error> {
        let mut budget = TransactionBudget::new(self);
        let source_catalog = physical_source(self)?;
        let source_owner = source_catalog.__source_owner();
        if !patch.artifacts.authorizes_owner(&source_owner) {
            return Err(Error::PatchConflict);
        }
        if patch.is_noop() {
            return Ok(Commit {
                package: self.snapshot(),
                patch: patch.clone(),
                diagnostics: Diagnostics::unchanged(),
            });
        }
        let selected = resolve_at_with_budget(
            self,
            patch.target.native.sheet_position,
            patch.target.native.table_position,
            &mut budget,
            true,
        )?;
        if selected.appearance != patch.before || selected.native != patch.target.native {
            return Err(Error::PatchConflict);
        }
        let target_owner = patch.artifacts.target_owner();
        budget.charge_output_bytes(target_owner.len(), Path::Package)?;
        budget.charge_transaction_work(target_owner.len().saturating_mul(2), Path::Package)?;
        let candidate = Package::from_source_owner_with_options(target_owner, self.state.options)
            .map_err(|_| Error::Verification)?;
        let after = resolve_at_with_budget(
            &candidate,
            patch.target.native.sheet_position,
            patch.target.native.table_position,
            &mut budget,
            true,
        )?;
        if after.appearance != patch.after {
            return Err(Error::Verification);
        }
        verify_candidate_locality(
            self,
            &candidate,
            patch.target,
            selected.style_identifier,
            after.style_identifier,
            patch.source_previews,
            patch.target_previews,
            Path::Package,
        )?;
        Ok(Commit {
            package: candidate,
            patch: patch.clone(),
            diagnostics: Diagnostics::published(
                patch.source_previews.saturating_sub(patch.target_previews),
                patch.touched_components,
            ),
        })
    }
}

fn commit_edit(edit: Edit<'_>) -> Result<Commit, Error> {
    let catalog = physical_source(edit.source)?;
    let source_owner = catalog.__source_owner();
    if edit.before == edit.appearance {
        return Ok(Commit {
            package: edit.source.snapshot(),
            patch: Patch {
                artifacts: OwnedExactArtifacts::new(source_owner.clone(), source_owner),
                target: edit.target,
                before: edit.before,
                after: edit.appearance,
                source_previews: 0,
                target_previews: 0,
                touched_components: 0,
            },
            diagnostics: Diagnostics::unchanged(),
        });
    }
    if !catalog.source_is_exact() {
        return Err(Error::UnsupportedSource);
    }
    if edit.target.native.locked == LockState::Locked {
        return Err(Error::TableLocked { path: edit.path() });
    }
    let mut budget = edit.budget;
    budget.charge_transaction_work(edit.source.source_bytes().len(), edit.path())?;
    let target = edit.target;
    let old_style = target
        .style_identifier
        .ok_or(Error::UnsupportedDependency { path: edit.path() })?;
    let metadata_selector_set = preflight_metadata_for_style(
        edit.source,
        target.native.component_index,
        old_style,
        edit.path(),
        &mut budget,
    )?;
    let (new_style, uuid, native_edits) = rewrite_native_appearance(
        edit.source,
        target,
        old_style,
        edit.appearance,
        edit.path(),
        &mut budget,
    )?;
    let metadata_edit = rewrite_metadata_for_style(
        edit.source,
        target.native.component_index,
        new_style,
        uuid,
        native_edits.style_component,
        metadata_selector_set,
        edit.path(),
        &mut budget,
    )?;
    let previews = super::table_headers::rewrite::root_preview_deletions(catalog)
        .map_err(|_| Error::InvalidSource { path: edit.path() })?;
    let mut edits = native_edits.edits;
    edits.push(metadata_edit);
    let touched_components = edits
        .iter()
        .enumerate()
        .filter(|(index, edit)| {
            edits[..*index]
                .iter()
                .all(|previous| previous.name != edit.name)
        })
        .count();
    let mut entry_edits = Vec::new();
    budget.charge_allocations(edits.len().saturating_add(1), edit.path())?;
    entry_edits
        .try_reserve_exact(edits.len())
        .map_err(|_| Error::Allocation {
            amount: edits.len(),
            path: edit.path(),
        })?;
    for edit in &edits {
        entry_edits.push(EntryEdit::new(edit.name.as_str(), edit.data.as_slice()));
    }
    let physical_limits = catalog.limits();
    let prepared = catalog
        .package()
        .prepare_reassembly_with_deletions(&entry_edits, &previews, physical_limits)
        .map_err(|error| map_archive_error(error, edit.path()))?;
    let requirements = prepared.execution_requirements();
    budget.preflight_reassembly(requirements, edit.path())?;
    let target_bytes = prepared
        .execute(requirements.exact_limits())
        .map_err(|error| map_archive_error(error, edit.path()))?;
    let package = Package::from_owned_bytes_with_options(target_bytes, edit.source.state.options)
        .map_err(|_| Error::Verification)?;
    let after = resolve_at_with_budget(
        &package,
        target.native.sheet_position,
        target.native.table_position,
        &mut budget,
        true,
    )?;
    if after.appearance != edit.appearance {
        return Err(Error::Verification);
    }
    verify_candidate_locality(
        edit.source,
        &package,
        target,
        Some(old_style),
        after.style_identifier,
        previews.len(),
        0,
        edit.path(),
    )?;
    let target_catalog = physical_source(&package)?;
    let target_owner = target_catalog.__source_owner();
    let patch = Patch {
        artifacts: OwnedExactArtifacts::new(source_owner, target_owner),
        target,
        before: edit.before,
        after: edit.appearance,
        source_previews: previews.len(),
        target_previews: 0,
        touched_components,
    };
    Ok(Commit {
        package,
        patch,
        diagnostics: Diagnostics::published(previews.len(), touched_components),
    })
}

#[derive(Debug)]
struct NativeEdit {
    name: String,
    data: Vec<u8>,
}

#[derive(Debug)]
struct NativeOutput {
    edits: Vec<NativeEdit>,
    style_component: usize,
}

#[derive(Debug, Clone, Copy)]
struct TransactionBudget {
    maximum_wire_bytes: usize,
    maximum_output_bytes: usize,
    maximum_fields: usize,
    maximum_work: usize,
    maximum_styles: usize,
    maximum_components: usize,
    maximum_references: usize,
    maximum_additions: usize,
    maximum_allocations: usize,
    maximum_transaction_work: usize,
    remaining_fields: usize,
    remaining_work: usize,
    remaining_output_bytes: usize,
    remaining_styles: usize,
    remaining_components: usize,
    remaining_references: usize,
    remaining_additions: usize,
    remaining_allocations: usize,
    remaining_transaction_work: usize,
}

impl TransactionBudget {
    fn new(source: &Package) -> Self {
        let archive = source.state.options.archive();
        let maximum_wire_bytes = archive.max_iwa_stream_bytes().max(1);
        let maximum_output_bytes = usize::try_from(archive.max_total_bytes())
            .unwrap_or(usize::MAX)
            .max(1);
        let maximum_fields = maximum_wire_bytes.saturating_mul(8).max(1);
        let maximum_work = maximum_wire_bytes.saturating_mul(32).max(1);
        let maximum_styles = maximum_wire_bytes.max(1);
        let maximum_components = maximum_wire_bytes.max(1);
        let maximum_references = source
            .state
            .options
            .semantic()
            .max_references()
            .saturating_mul(8)
            .max(1);
        let maximum_additions = 64;
        let maximum_transaction_work = usize::try_from(archive.max_total_bytes())
            .unwrap_or(usize::MAX)
            .saturating_mul(32)
            .max(1);
        // Strict metadata ownership is inspected alongside the native style
        // graph.  Bound its visitor/codec scratch by the source component
        // cardinality instead of using a fixed allowance that rejects valid
        // multi-component packages before publication.
        let maximum_allocations = source
            .state
            .components
            .catalog()
            .len()
            .saturating_mul(64)
            .saturating_add(128)
            .max(128);
        Self {
            maximum_wire_bytes,
            maximum_output_bytes,
            maximum_fields,
            maximum_work,
            maximum_styles,
            maximum_components,
            maximum_references,
            maximum_additions,
            maximum_allocations,
            maximum_transaction_work,
            remaining_fields: maximum_fields,
            remaining_work: maximum_work,
            remaining_output_bytes: maximum_output_bytes,
            remaining_styles: maximum_styles,
            remaining_components: maximum_components,
            remaining_references: maximum_references,
            remaining_additions: maximum_additions,
            remaining_allocations: maximum_allocations,
            remaining_transaction_work: maximum_transaction_work,
        }
    }

    fn codec_options(&self, source: &[u8]) -> codec::DecodeOptions {
        codec::DecodeOptions::new(
            self.maximum_wire_bytes.min(source.len().max(1)),
            self.remaining_transaction_work,
            self.remaining_fields,
            self.remaining_work,
            u32::try_from(WireLimits::MAX_NESTING).unwrap_or(u32::MAX),
            self.remaining_styles,
        )
        .with_max_allocations(self.remaining_allocations)
    }

    fn charge_fields(&mut self, amount: usize, path: Path) -> Result<(), Error> {
        charge_budget(
            &mut self.remaining_fields,
            self.maximum_fields,
            amount,
            LimitKind::WireFields,
            path,
        )
    }

    fn charge_work(&mut self, amount: usize, path: Path) -> Result<(), Error> {
        charge_budget(
            &mut self.remaining_work,
            self.maximum_work,
            amount,
            LimitKind::WireWork,
            path,
        )
    }

    fn charge_styles(&mut self, amount: usize, path: Path) -> Result<(), Error> {
        charge_budget(
            &mut self.remaining_styles,
            self.maximum_styles,
            amount,
            LimitKind::PayloadItems,
            path,
        )
    }

    fn charge_components(&mut self, amount: usize, path: Path) -> Result<(), Error> {
        charge_budget(
            &mut self.remaining_components,
            self.maximum_components,
            amount,
            LimitKind::PayloadItems,
            path,
        )
    }

    fn charge_references(&mut self, amount: usize, path: Path) -> Result<(), Error> {
        charge_budget(
            &mut self.remaining_references,
            self.maximum_references,
            amount,
            LimitKind::PayloadReferences,
            path,
        )
    }

    fn charge_additions(&mut self, amount: usize, path: Path) -> Result<(), Error> {
        charge_budget(
            &mut self.remaining_additions,
            self.maximum_additions,
            amount,
            LimitKind::PayloadItems,
            path,
        )
    }

    fn charge_allocations(&mut self, amount: usize, path: Path) -> Result<(), Error> {
        charge_budget(
            &mut self.remaining_allocations,
            self.maximum_allocations,
            amount,
            LimitKind::TransactionWork,
            path,
        )
    }

    fn charge_transaction_work(&mut self, amount: usize, path: Path) -> Result<(), Error> {
        charge_budget(
            &mut self.remaining_transaction_work,
            self.maximum_transaction_work,
            amount,
            LimitKind::TransactionWork,
            path,
        )
    }

    fn charge_output_bytes(&mut self, amount: usize, path: Path) -> Result<(), Error> {
        charge_budget(
            &mut self.remaining_output_bytes,
            self.maximum_output_bytes,
            amount,
            LimitKind::OutputBytes,
            path,
        )
    }

    fn preflight_reassembly(
        &mut self,
        requirements: litchi_iwa_archive::package::ReassemblyExecutionRequirements,
        path: Path,
    ) -> Result<(), Error> {
        self.charge_output_bytes(requirements.output_bytes(), path)?;
        self.charge_allocations(requirements.allocations(), path)?;
        self.charge_transaction_work(
            requirements
                .output_bytes()
                .saturating_add(requirements.scratch_bytes())
                .saturating_add(requirements.retained_bytes()),
            path,
        )
    }

    fn consume_report(&mut self, report: codec::DecodeReport, path: Path) -> Result<(), Error> {
        self.charge_fields(report.fields(), path)?;
        self.charge_work(report.work_bytes(), path)?;
        self.charge_allocations(report.allocations(), path)?;
        self.charge_transaction_work(
            report.input_bytes().saturating_add(report.output_bytes()),
            path,
        )
    }

    fn preflight_requirements(
        &mut self,
        requirements: codec::RewriteExecutionRequirements,
        path: Path,
    ) -> Result<(), Error> {
        self.charge_fields(requirements.fields(), path)?;
        self.charge_work(requirements.work_bytes(), path)?;
        self.charge_allocations(requirements.allocations(), path)?;
        self.charge_transaction_work(
            requirements
                .input_bytes()
                .saturating_add(requirements.output_bytes()),
            path,
        )
    }

    fn preflight_metadata_requirements(
        &mut self,
        requirements: litchi_iwa_protos::package_metadata_codec::RewriteExecutionRequirements,
        path: Path,
    ) -> Result<(), Error> {
        self.charge_fields(requirements.fields(), path)?;
        self.charge_work(requirements.work_bytes(), path)?;
        self.charge_components(requirements.components(), path)?;
        self.charge_references(requirements.references(), path)?;
        self.charge_allocations(requirements.allocations(), path)?;
        self.charge_transaction_work(
            requirements
                .output_bytes()
                .saturating_add(requirements.retained_bytes())
                .saturating_add(requirements.scratch_bytes()),
            path,
        )
    }

    fn metadata_options(
        &self,
        source: &Package,
        bytes: usize,
        additions: usize,
    ) -> MetadataRewriteOptions {
        let base = metadata_options(source, bytes, additions);
        MetadataRewriteOptions::new(
            base.max_input_bytes().min(self.remaining_transaction_work),
            base.max_output_bytes().min(self.remaining_transaction_work),
            base.max_fields().min(self.remaining_fields),
            base.max_work_bytes().min(self.remaining_work),
            base.recursion_limit(),
            base.max_components().min(self.remaining_components),
            base.max_references().min(self.remaining_references),
            base.max_additions().min(self.remaining_additions),
        )
    }

    fn consume_metadata_report(
        &mut self,
        report: litchi_iwa_protos::package_metadata_codec::RewriteReport,
        path: Path,
    ) -> Result<(), Error> {
        self.charge_fields(report.fields(), path)?;
        self.charge_work(report.work_bytes(), path)?;
        self.charge_components(report.components_scanned(), path)?;
        self.charge_references(report.references_scanned(), path)?;
        self.charge_additions(report.additions(), path)?;
        self.charge_transaction_work(
            report.input_bytes().saturating_add(report.output_bytes()),
            path,
        )
    }
}

fn charge_budget(
    remaining: &mut usize,
    maximum: usize,
    amount: usize,
    kind: LimitKind,
    path: Path,
) -> Result<(), Error> {
    if amount > *remaining {
        let observed = maximum.saturating_sub(*remaining).saturating_add(amount);
        return Err(Error::LimitExceeded {
            kind,
            observed: u64::try_from(observed).unwrap_or(u64::MAX),
            maximum: u64::try_from(maximum).unwrap_or(u64::MAX),
            path,
        });
    }
    *remaining -= amount;
    Ok(())
}

fn rewrite_native_appearance(
    source: &Package,
    target: Target,
    old_style: u64,
    appearance: Appearance,
    path: Path,
    budget: &mut TransactionBudget,
) -> Result<(u64, UuidBits, NativeOutput), Error> {
    let style_location = resolved_location(source, old_style, path)?;
    let style_payload = style_message_data(source, old_style, path)?;
    let (style, style_report) =
        codec::decode_table_style_with_report(style_payload, budget.codec_options(style_payload))
            .map_err(|error| map_codec_error(error, path))?;
    budget.consume_report(style_report, path)?;
    let stylesheet_id = style
        .stylesheet_identifier()
        .filter(|identifier| *identifier != 0)
        .ok_or(Error::UnsupportedDependency { path })?;
    let stylesheet_location = resolved_location(source, stylesheet_id, path)?;
    let stylesheet_payload = stylesheet_message_data(source, stylesheet_id, path)?;
    let style_options = budget.codec_options(style_payload);
    let stylesheet_options = budget.codec_options(stylesheet_payload);
    let new_style = next_style_identifier(source, path, budget)?;
    let uuid = UuidBits::new(
        new_style ^ 0x9e37_79b9_7f4a_7c15,
        new_style.rotate_left(29) ^ 0xd1b5_4a32_d192_ed03,
    );
    if uuid.lower() == 0 && uuid.upper() == 0 {
        return Err(Error::InvalidSource { path });
    }
    let model_payload = super::table_headers::rewrite::selected_payload(source, target.native)
        .map_err(|_| Error::InvalidSource { path })?;
    let physical = physical_source(source)?;
    let archive_limits = physical
        .limits()
        .effective_archive_limits()
        .map_err(|error| map_archive_error(error, path))?;
    let mut component_indices = vec![target.native.component_index, style_location.0];
    if !component_indices.contains(&stylesheet_location.0) {
        component_indices.push(stylesheet_location.0);
    }
    component_indices.sort_unstable();
    component_indices.dedup();
    let snappy_limits = physical
        .limits()
        .snappy_limits()
        .map_err(|error| map_archive_error(error, path))?;
    let maximum_archive_bytes = archive_limits.max_archive_bytes();
    let maximum_compressed_bytes = SnappyStream::maximum_compressed_len(maximum_archive_bytes)
        .map_err(|error| map_core_error(error, path))?;
    budget.charge_allocations(
        component_indices.len().saturating_mul(4).saturating_add(8),
        path,
    )?;
    budget.charge_transaction_work(
        component_indices
            .len()
            .saturating_mul(maximum_archive_bytes.saturating_add(maximum_compressed_bytes)),
        path,
    )?;
    let mut edits = Vec::new();
    edits
        .try_reserve_exact(component_indices.len())
        .map_err(|_| Error::Allocation {
            amount: component_indices.len(),
            path,
        })?;
    let variation = codec::canonical_table_style_variation(
        codec::TableStyleVariationWrite {
            parent_identifier: old_style,
            stylesheet_identifier: stylesheet_id,
            overrides: appearance_overrides(appearance),
        },
        style_options,
    )
    .map_err(|error| map_codec_error(error, path))?;
    budget.consume_report(variation.report(), path)?;
    let model_plan = codec::prepare_table_model_style_rewrite(
        model_payload,
        old_style,
        new_style,
        budget.codec_options(model_payload),
    )
    .map_err(|error| map_codec_error(error, path))?;
    budget.preflight_requirements(model_plan.execution_requirements(), path)?;
    let (model_rewritten, _) = model_plan
        .execute(model_plan.execution_requirements().exact_limits())
        .map_err(|error| map_codec_error(error, path))?;
    let stylesheet_plan = codec::prepare_stylesheet_append(
        stylesheet_payload,
        codec::StylesheetStyleAppend {
            style_identifier: new_style,
            parent_identifier: Some(old_style),
        },
        stylesheet_options,
    )
    .map_err(|error| map_codec_error(error, path))?;
    budget.preflight_requirements(stylesheet_plan.execution_requirements(), path)?;
    let (stylesheet_rewritten, _) = stylesheet_plan
        .execute(stylesheet_plan.execution_requirements().exact_limits())
        .map_err(|error| map_codec_error(error, path))?;
    for component_index in component_indices {
        let component = source
            .state
            .components
            .catalog()
            .get_index(component_index)
            .ok_or(Error::InvalidSource { path })?;
        let name = component.name().to_owned();
        let entry = physical
            .package()
            .iter()
            .find(|entry| entry.name() == name)
            .ok_or(Error::InvalidSource { path })?;
        if entry.is_opaque() {
            return Err(Error::UnsupportedSource);
        }
        budget.charge_transaction_work(entry.data().len(), path)?;
        let stream = SnappyStream::decompress_with_limits(entry.data(), snappy_limits)
            .map_err(|error| map_core_error(error, path))?;
        let mut archive = Archive::parse_with_limits(stream.as_bytes(), archive_limits)
            .map_err(|error| map_core_error(error, path))?;
        archive
            .validate_canonical_object_framing(stream.as_bytes())
            .map_err(|error| map_core_error(error, path))?;
        if component_index == target.native.component_index {
            replace_model_in_archive(
                &mut archive,
                target.native.model_identifier,
                old_style,
                new_style,
                target.native.message_type,
                model_rewritten.as_slice(),
                archive_limits,
                path,
            )?;
        }
        if component_index == style_location.0 {
            append_style_object(
                &mut archive,
                new_style,
                old_style,
                stylesheet_id,
                variation.bytes(),
                archive_limits,
                path,
            )?;
        }
        if component_index == stylesheet_location.0 {
            replace_stylesheet_in_archive(
                &mut archive,
                stylesheet_id,
                new_style,
                stylesheet_rewritten.as_slice(),
                archive_limits,
                path,
            )?;
        }
        let bytes = archive
            .to_bytes_with_limits(archive_limits)
            .map_err(|error| map_core_error(error, path))?;
        let compressed =
            SnappyStream::compress(&bytes).map_err(|error| map_core_error(error, path))?;
        edits.push(NativeEdit {
            name,
            data: compressed,
        });
    }
    Ok((
        new_style,
        uuid,
        NativeOutput {
            edits,
            style_component: style_location.0,
        },
    ))
}

fn appearance_overrides(appearance: Appearance) -> codec::AppearanceOverrides {
    codec::AppearanceOverrides {
        row_banding: Some(matches!(
            appearance.row_banding,
            crate::table::appearance::Banding::Enabled
        )),
        row_sizing: Some(matches!(
            appearance.row_sizing,
            crate::table::appearance::RowSizing::FitCellContents
        )),
        body_horizontal: Some(matches!(
            appearance.gridlines.body_horizontal,
            crate::table::appearance::GridlineVisibility::Visible
        )),
        body_vertical: Some(matches!(
            appearance.gridlines.body_vertical,
            crate::table::appearance::GridlineVisibility::Visible
        )),
        header_columns_horizontal: Some(matches!(
            appearance.gridlines.header_columns_horizontal,
            crate::table::appearance::GridlineVisibility::Visible
        )),
        header_rows_vertical: Some(matches!(
            appearance.gridlines.header_rows_vertical,
            crate::table::appearance::GridlineVisibility::Visible
        )),
        footer_rows_vertical: Some(matches!(
            appearance.gridlines.footer_rows_vertical,
            crate::table::appearance::GridlineVisibility::Visible
        )),
    }
}

fn next_style_identifier(
    source: &Package,
    path: Path,
    budget: &mut TransactionBudget,
) -> Result<u64, Error> {
    let mut maximum = 0u64;
    for component in source.state.components.catalog().iter() {
        for object in &component.archive().objects {
            maximum = maximum.max(
                object
                    .archive_info
                    .identifier
                    .ok_or(Error::InvalidSource { path })?,
            );
            budget.charge_transaction_work(1, path)?;
        }
    }
    let metadata = metadata_payload(source, path)?;
    let mut visitor = IdentifierCensus::default();
    let inspection = inspect_package_metadata_with_visitor(
        metadata,
        budget.metadata_options(source, metadata.len(), 0),
        &mut visitor,
    )
    .map_err(|error| map_metadata_error(error, path))?;
    budget.consume_metadata_report(inspection.report(), path)?;
    maximum = maximum
        .max(inspection.last_object_identifier())
        .max(visitor.maximum);
    maximum
        .checked_add(1)
        .filter(|identifier| *identifier != 0)
        .ok_or(Error::LimitExceeded {
            kind: LimitKind::PayloadObjects,
            observed: u64::MAX,
            maximum: u64::MAX - 1,
            path,
        })
}

fn resolved_location(
    source: &Package,
    identifier: u64,
    path: Path,
) -> Result<(usize, usize), Error> {
    let resolved = source
        .state
        .index
        .resolve_ref_id(&source.state.components, identifier)
        .map_err(|_| Error::InvalidSource { path })?
        .ok_or(Error::InvalidSource { path })?;
    let object = source
        .state
        .components
        .catalog()
        .get_index(resolved.component_index)
        .and_then(|component| component.archive().objects.get(resolved.object_index))
        .ok_or(Error::InvalidSource { path })?;
    if object.archive_info.identifier != Some(identifier) {
        return Err(Error::InvalidSource { path });
    }
    Ok((resolved.component_index, resolved.object_index))
}

fn stylesheet_message_data(source: &Package, identifier: u64, path: Path) -> Result<&[u8], Error> {
    let (component_index, object_index) = resolved_location(source, identifier, path)?;
    let object = source
        .state
        .components
        .catalog()
        .get_index(component_index)
        .and_then(|component| component.archive().objects.get(object_index))
        .ok_or(Error::InvalidSource { path })?;
    let mut messages = object
        .messages
        .iter()
        .filter(|message| message.type_ == STYLESHEET_MESSAGE_TYPE);
    let message = messages
        .next()
        .ok_or(Error::UnsupportedDependency { path })?;
    if messages.next().is_some() {
        return Err(Error::InvalidSource { path });
    }
    Ok(message.data.as_slice())
}

fn map_codec_error(error: codec::DecodeError, path: Path) -> Error {
    if let Some(amount) = error.allocation_requested() {
        return Error::Allocation { amount, path };
    }
    let Some(limit) = error.resource_limit() else {
        return Error::InvalidSource { path };
    };
    let (kind, observed, maximum) = match limit {
        codec::DecodeLimit::InputBytes { observed, maximum } => {
            (LimitKind::WireBytes, observed, maximum)
        },
        codec::DecodeLimit::OutputBytes { observed, maximum } => {
            (LimitKind::WireOutputBytes, observed, maximum)
        },
        codec::DecodeLimit::Fields { observed, maximum } => {
            (LimitKind::WireFields, observed, maximum)
        },
        codec::DecodeLimit::WorkBytes { observed, maximum } => {
            (LimitKind::WireWork, observed, maximum)
        },
        codec::DecodeLimit::Nesting { observed, maximum } => {
            return Error::LimitExceeded {
                kind: LimitKind::WireNesting,
                observed: u64::from(observed),
                maximum: u64::from(maximum),
                path,
            };
        },
        codec::DecodeLimit::Styles { observed, maximum } => {
            (LimitKind::PayloadItems, observed, maximum)
        },
        codec::DecodeLimit::Allocations { observed, maximum } => {
            (LimitKind::TransactionWork, observed, maximum)
        },
        _ => return Error::InvalidSource { path },
    };
    Error::LimitExceeded {
        kind,
        observed: u64::try_from(observed).unwrap_or(u64::MAX),
        maximum: u64::try_from(maximum).unwrap_or(u64::MAX),
        path,
    }
}

fn snapshot_appearance(snapshot: codec::AppearanceSnapshot) -> Appearance {
    Appearance {
        row_banding: if snapshot.row_banding {
            crate::table::appearance::Banding::Enabled
        } else {
            crate::table::appearance::Banding::Disabled
        },
        row_sizing: if snapshot.row_sizing {
            crate::table::appearance::RowSizing::FitCellContents
        } else {
            crate::table::appearance::RowSizing::Fixed
        },
        gridlines: crate::table::appearance::Gridlines {
            body_horizontal: bool_gridline(snapshot.body_horizontal),
            body_vertical: bool_gridline(snapshot.body_vertical),
            header_columns_horizontal: bool_gridline(snapshot.header_columns_horizontal),
            header_rows_vertical: bool_gridline(snapshot.header_rows_vertical),
            footer_rows_vertical: bool_gridline(snapshot.footer_rows_vertical),
        },
    }
}

const fn bool_gridline(value: bool) -> crate::table::appearance::GridlineVisibility {
    if value {
        crate::table::appearance::GridlineVisibility::Visible
    } else {
        crate::table::appearance::GridlineVisibility::Hidden
    }
}

struct SelectorVisitor<'source> {
    locators: &'source [&'source str],
    identifiers: Vec<Option<u64>>,
    effective_locators: Vec<Option<String>>,
    duplicate: bool,
}

struct ParentReferenceVisitor<'source> {
    source_identifier: u64,
    source_locator: &'source str,
    target_identifier: u64,
    object_identifier: u64,
    exact: usize,
    related: usize,
    versioned: bool,
    wrong_weakness: bool,
    unexpected: bool,
}

struct StyleRegistryVisitor {
    target_identifier: u64,
    expected_component_identifier: u64,
    current_uuid: usize,
    versioned_uuid: bool,
    cross_component_uuid: bool,
    external_reference: bool,
    data_owner: bool,
    ambiguous: bool,
    root_data_map: bool,
}

struct MetadataSelectorSet {
    identifiers: Vec<u64>,
    effective_locators: Vec<String>,
    last_object_identifier: u64,
}

impl MetadataSelectorSet {
    fn selectors(&self, path: Path) -> Result<Vec<ComponentSelector<'_>>, Error> {
        let mut selectors = Vec::new();
        selectors
            .try_reserve_exact(self.identifiers.len())
            .map_err(|_| Error::Allocation {
                amount: self.identifiers.len(),
                path,
            })?;
        selectors.extend(
            self.identifiers
                .iter()
                .zip(self.effective_locators.iter())
                .map(|(identifier, locator)| ComponentSelector::new(*identifier, locator.as_str())),
        );
        Ok(selectors)
    }
}

#[derive(Default)]
struct IdentifierCensus {
    maximum: u64,
}

impl IdentifierCensus {
    fn record(&mut self, identifier: u64) {
        self.maximum = self.maximum.max(identifier);
    }
}

impl PackageMetadataVisitor for IdentifierCensus {
    fn visit_component(
        &mut self,
        component: litchi_iwa_protos::package_metadata_codec::ComponentDescriptor<'_>,
    ) -> Result<(), MetadataRewriteError> {
        self.record(component.identifier());
        Ok(())
    }

    fn visit_object_uuid(
        &mut self,
        binding: litchi_iwa_protos::package_metadata_codec::ObjectUuidDescriptor<'_>,
    ) -> Result<(), MetadataRewriteError> {
        self.record(binding.object_identifier());
        Ok(())
    }

    fn visit_external_reference(
        &mut self,
        reference: litchi_iwa_protos::package_metadata_codec::ExternalReferenceDescriptor<'_>,
    ) -> Result<(), MetadataRewriteError> {
        self.record(reference.target_component_identifier());
        if let Some(identifier) = reference.object_identifier() {
            self.record(identifier);
        }
        Ok(())
    }

    fn visit_data_reference_owner(
        &mut self,
        owner: litchi_iwa_protos::package_metadata_codec::DataReferenceOwnerDescriptor<'_>,
    ) -> Result<(), MetadataRewriteError> {
        self.record(owner.data_identifier());
        self.record(owner.object_identifier());
        Ok(())
    }

    fn visit_ambiguous_object_identifier(
        &mut self,
        _component: litchi_iwa_protos::package_metadata_codec::ComponentDescriptor<'_>,
        identifier: u64,
    ) -> Result<(), MetadataRewriteError> {
        self.record(identifier);
        Ok(())
    }

    fn visit_data_metadata_map(
        &mut self,
        object_identifier: u64,
        _has_unknown_fields: bool,
    ) -> Result<(), MetadataRewriteError> {
        self.record(object_identifier);
        Ok(())
    }
}

impl<'source> SelectorVisitor<'source> {
    fn new(locators: &'source [&'source str], path: Path) -> Result<Self, Error> {
        let mut identifiers = Vec::new();
        identifiers
            .try_reserve_exact(locators.len())
            .map_err(|_| Error::Allocation {
                amount: locators.len(),
                path,
            })?;
        identifiers.resize(locators.len(), None);
        let mut effective_locators = Vec::new();
        effective_locators
            .try_reserve_exact(locators.len())
            .map_err(|_| Error::Allocation {
                amount: locators.len(),
                path,
            })?;
        effective_locators.resize_with(locators.len(), || None);
        Ok(Self {
            locators,
            identifiers,
            effective_locators,
            duplicate: false,
        })
    }
}

impl PackageMetadataVisitor for SelectorVisitor<'_> {
    fn visit_component(
        &mut self,
        component: litchi_iwa_protos::package_metadata_codec::ComponentDescriptor<'_>,
    ) -> Result<(), MetadataRewriteError> {
        if !component.is_current() {
            return Ok(());
        }
        for (index, locator) in self.locators.iter().enumerate() {
            if component.preferred_locator() != *locator
                && component.effective_locator() != *locator
            {
                continue;
            }
            if self.identifiers[index]
                .replace(component.identifier())
                .is_some()
            {
                self.duplicate = true;
            }
            self.effective_locators[index] = Some(component.effective_locator().to_owned());
        }
        Ok(())
    }
}

impl<'source> ParentReferenceVisitor<'source> {
    fn new(
        source: ComponentSelector<'source>,
        target: ComponentSelector<'source>,
        object: u64,
    ) -> Self {
        Self {
            source_identifier: source.identifier(),
            source_locator: source.locator(),
            target_identifier: target.identifier(),
            object_identifier: object,
            exact: 0,
            related: 0,
            versioned: false,
            wrong_weakness: false,
            unexpected: false,
        }
    }
}

impl PackageMetadataVisitor for ParentReferenceVisitor<'_> {
    fn visit_external_reference(
        &mut self,
        reference: litchi_iwa_protos::package_metadata_codec::ExternalReferenceDescriptor<'_>,
    ) -> Result<(), MetadataRewriteError> {
        let source = reference.source();
        let related = source.identifier() == self.source_identifier
            && reference.target_component_identifier() == self.target_identifier
            && reference.object_identifier() == Some(self.object_identifier);
        if !related {
            if reference.object_identifier() == Some(self.object_identifier) {
                self.unexpected = true;
            }
            return Ok(());
        }
        self.related = self.related.saturating_add(1);
        if reference.is_versioned() || !source.is_current() {
            self.versioned = true;
            return Ok(());
        }
        if source.effective_locator() != self.source_locator {
            self.wrong_weakness = true;
            return Ok(());
        }
        if reference.is_weak().is_some() {
            self.wrong_weakness = true;
            return Ok(());
        }
        self.exact = self.exact.saturating_add(1);
        Ok(())
    }
}

impl StyleRegistryVisitor {
    fn new(component: ComponentSelector<'_>, target_identifier: u64) -> Self {
        Self {
            target_identifier,
            expected_component_identifier: component.identifier(),
            current_uuid: 0,
            versioned_uuid: false,
            cross_component_uuid: false,
            external_reference: false,
            data_owner: false,
            ambiguous: false,
            root_data_map: false,
        }
    }
}

impl PackageMetadataVisitor for StyleRegistryVisitor {
    fn visit_object_uuid(
        &mut self,
        binding: litchi_iwa_protos::package_metadata_codec::ObjectUuidDescriptor<'_>,
    ) -> Result<(), MetadataRewriteError> {
        if binding.object_identifier() != self.target_identifier {
            return Ok(());
        }
        if binding.component().is_current() {
            self.current_uuid = self.current_uuid.saturating_add(1);
            if binding.component().identifier() != self.expected_component_identifier {
                self.cross_component_uuid = true;
            }
        } else {
            self.versioned_uuid = true;
        }
        Ok(())
    }

    fn visit_external_reference(
        &mut self,
        reference: litchi_iwa_protos::package_metadata_codec::ExternalReferenceDescriptor<'_>,
    ) -> Result<(), MetadataRewriteError> {
        if reference.object_identifier() == Some(self.target_identifier) {
            self.external_reference = true;
        }
        Ok(())
    }

    fn visit_data_reference_owner(
        &mut self,
        owner: litchi_iwa_protos::package_metadata_codec::DataReferenceOwnerDescriptor<'_>,
    ) -> Result<(), MetadataRewriteError> {
        if owner.object_identifier() == self.target_identifier
            || owner.data_identifier() == self.target_identifier
        {
            self.data_owner = true;
        }
        Ok(())
    }

    fn visit_ambiguous_object_identifier(
        &mut self,
        _component: litchi_iwa_protos::package_metadata_codec::ComponentDescriptor<'_>,
        identifier: u64,
    ) -> Result<(), MetadataRewriteError> {
        self.ambiguous |= identifier == self.target_identifier;
        Ok(())
    }

    fn visit_data_metadata_map(
        &mut self,
        object_identifier: u64,
        _has_unknown_fields: bool,
    ) -> Result<(), MetadataRewriteError> {
        self.root_data_map |= object_identifier == self.target_identifier;
        Ok(())
    }
}

fn metadata_payload(source: &Package, path: Path) -> Result<&[u8], Error> {
    let route =
        super::metadata::unique_message_route(source).ok_or(Error::InvalidSource { path })?;
    source
        .state
        .components
        .catalog()
        .get_index(route.component_index)
        .and_then(|component| component.archive().objects.get(route.object_index))
        .and_then(|object| object.messages.get(route.message_index))
        .filter(|message| message.type_ == super::metadata::MESSAGE_TYPE)
        .map(|message| message.data.as_slice())
        .ok_or(Error::InvalidSource { path })
}

fn metadata_options(source: &Package, bytes: usize, additions: usize) -> MetadataRewriteOptions {
    let physical_max = source.state.options.archive().max_iwa_stream_bytes();
    MetadataRewriteOptions::new(
        bytes.max(1).min(physical_max),
        bytes.saturating_add(4_096).max(1).min(physical_max),
        bytes.saturating_mul(64).clamp(1, WireLimits::MAX_FIELDS),
        bytes
            .saturating_mul(256)
            .saturating_add(additions.saturating_mul(bytes))
            .clamp(1, WireLimits::MAX_REWRITE_WORK),
        64,
        bytes
            .max(source.state.components.catalog().len())
            .saturating_mul(4)
            .max(1),
        source.state.options.semantic().max_references().max(1),
        additions.max(1),
    )
}

fn metadata_selectors(
    source: &Package,
    component_indices: &[usize],
    payload: &[u8],
    path: Path,
    budget: &mut TransactionBudget,
) -> Result<MetadataSelectorSet, Error> {
    let mut locators = Vec::new();
    locators
        .try_reserve_exact(component_indices.len())
        .map_err(|_| Error::Allocation {
            amount: component_indices.len(),
            path,
        })?;
    budget.charge_allocations(
        component_indices.len().saturating_mul(3).saturating_add(4),
        path,
    )?;
    for &index in component_indices {
        let component = source
            .state
            .components
            .catalog()
            .get_index(index)
            .ok_or(Error::InvalidSource { path })?;
        locators.push(super::metadata::normalized_locator(component.name()));
    }
    let locator_refs: Vec<&str> = locators.to_vec();
    let mut visitor = SelectorVisitor::new(&locator_refs, path)?;
    let inspection = inspect_package_metadata_with_visitor(
        payload,
        budget.metadata_options(source, payload.len(), component_indices.len()),
        &mut visitor,
    )
    .map_err(|error| map_metadata_error(error, path))?;
    budget.consume_metadata_report(inspection.report(), path)?;
    if visitor.duplicate || visitor.identifiers.iter().any(Option::is_none) {
        return Err(Error::InvalidSource { path });
    }
    let mut identifiers = Vec::new();
    identifiers
        .try_reserve_exact(visitor.identifiers.len())
        .map_err(|_| Error::Allocation {
            amount: visitor.identifiers.len(),
            path,
        })?;
    let mut effective_locators = Vec::new();
    effective_locators
        .try_reserve_exact(visitor.identifiers.len())
        .map_err(|_| Error::Allocation {
            amount: visitor.identifiers.len(),
            path,
        })?;
    for (identifier, locator) in visitor
        .identifiers
        .into_iter()
        .zip(visitor.effective_locators)
    {
        identifiers.push(identifier.ok_or(Error::InvalidSource { path })?);
        effective_locators.push(locator.ok_or(Error::InvalidSource { path })?);
    }
    Ok(MetadataSelectorSet {
        identifiers,
        effective_locators,
        last_object_identifier: inspection.last_object_identifier(),
    })
}

fn validate_parent_external_reference(
    source: &Package,
    payload: &[u8],
    model_selector: ComponentSelector<'_>,
    style_selector: ComponentSelector<'_>,
    old_style: u64,
    path: Path,
    budget: &mut TransactionBudget,
) -> Result<(), Error> {
    let mut visitor = ParentReferenceVisitor::new(model_selector, style_selector, old_style);
    let inspection = inspect_package_metadata_with_visitor(
        payload,
        budget.metadata_options(source, payload.len(), 0),
        &mut visitor,
    )
    .map_err(|error| map_metadata_error(error, path))?;
    budget.consume_metadata_report(inspection.report(), path)?;
    if visitor.exact != 1
        || visitor.related != 1
        || visitor.versioned
        || visitor.wrong_weakness
        || visitor.unexpected
    {
        return Err(Error::InvalidSource { path });
    }
    Ok(())
}

fn validate_style_registry_entry(
    source: &Package,
    payload: &[u8],
    component: ComponentSelector<'_>,
    identifier: u64,
    allow_external_reference: bool,
    path: Path,
    budget: &mut TransactionBudget,
) -> Result<(), Error> {
    let mut visitor = StyleRegistryVisitor::new(component, identifier);
    let inspection = inspect_package_metadata_with_visitor(
        payload,
        budget.metadata_options(source, payload.len(), 0),
        &mut visitor,
    )
    .map_err(|error| map_metadata_error(error, path))?;
    budget.consume_metadata_report(inspection.report(), path)?;
    if visitor.current_uuid != 1
        || visitor.versioned_uuid
        || visitor.cross_component_uuid
        || visitor.data_owner
        || visitor.ambiguous
        || visitor.root_data_map
        || (!allow_external_reference && visitor.external_reference)
    {
        return Err(Error::InvalidSource { path });
    }
    Ok(())
}

fn map_metadata_error(error: MetadataRewriteError, path: Path) -> Error {
    if let Some(amount) = error.allocation_request() {
        return Error::Allocation { amount, path };
    }
    let Some(limit) = error.resource_limit() else {
        return Error::InvalidSource { path };
    };
    let (kind, observed, maximum) = match limit {
        litchi_iwa_protos::package_metadata_codec::RewriteLimit::InputBytes {
            observed,
            maximum,
        } => (LimitKind::WireBytes, observed, maximum),
        litchi_iwa_protos::package_metadata_codec::RewriteLimit::OutputBytes {
            observed,
            maximum,
        } => (LimitKind::WireOutputBytes, observed, maximum),
        litchi_iwa_protos::package_metadata_codec::RewriteLimit::Fields { observed, maximum } => {
            (LimitKind::WireFields, observed, maximum)
        },
        litchi_iwa_protos::package_metadata_codec::RewriteLimit::Work { observed, maximum } => {
            (LimitKind::WireWork, observed, maximum)
        },
        litchi_iwa_protos::package_metadata_codec::RewriteLimit::Nesting { observed, maximum } => {
            return Error::LimitExceeded {
                kind: LimitKind::WireNesting,
                observed: u64::from(observed),
                maximum: u64::from(maximum),
                path,
            };
        },
        litchi_iwa_protos::package_metadata_codec::RewriteLimit::Components {
            observed,
            maximum,
        } => (LimitKind::PayloadItems, observed, maximum),
        litchi_iwa_protos::package_metadata_codec::RewriteLimit::References {
            observed,
            maximum,
        } => (LimitKind::PayloadReferences, observed, maximum),
        litchi_iwa_protos::package_metadata_codec::RewriteLimit::Additions {
            observed,
            maximum,
        } => (LimitKind::PayloadItems, observed, maximum),
        _ => return Error::InvalidSource { path },
    };
    Error::LimitExceeded {
        kind,
        observed: u64::try_from(observed).unwrap_or(u64::MAX),
        maximum: u64::try_from(maximum).unwrap_or(u64::MAX),
        path,
    }
}

fn preflight_metadata_for_style(
    source: &Package,
    model_component: usize,
    old_style: u64,
    path: Path,
    budget: &mut TransactionBudget,
) -> Result<MetadataSelectorSet, Error> {
    let style_component = resolved_location(source, old_style, path)?.0;
    let payload = metadata_payload(source, path)?;
    let component_indices = if model_component == style_component {
        vec![model_component]
    } else {
        vec![model_component, style_component]
    };
    let selector_set = metadata_selectors(source, &component_indices, payload, path, budget)?;
    budget.charge_allocations(selector_set.identifiers.len(), path)?;
    let selectors = selector_set.selectors(path)?;
    let model_selector = selectors
        .first()
        .copied()
        .ok_or(Error::InvalidSource { path })?;
    let style_selector = if model_component == style_component {
        model_selector
    } else {
        selectors
            .get(1)
            .copied()
            .ok_or(Error::InvalidSource { path })?
    };
    if model_component != style_component {
        validate_parent_external_reference(
            source,
            payload,
            model_selector,
            style_selector,
            old_style,
            path,
            budget,
        )?;
    }
    let style_payload = style_message_data(source, old_style, path)?;
    let (style, style_report) =
        codec::decode_table_style_with_report(style_payload, budget.codec_options(style_payload))
            .map_err(|error| map_codec_error(error, path))?;
    budget.consume_report(style_report, path)?;
    let stylesheet_identifier = style
        .stylesheet_identifier()
        .filter(|identifier| *identifier != 0)
        .ok_or(Error::UnsupportedDependency { path })?;
    let stylesheet_component = resolved_location(source, stylesheet_identifier, path)?.0;
    if stylesheet_component != style_component {
        return Err(Error::UnsupportedDependency { path });
    }
    let stylesheet_selector = style_selector;
    validate_style_registry_entry(
        source,
        payload,
        style_selector,
        old_style,
        model_component != style_component,
        path,
        budget,
    )?;
    // Native Numbers commonly uses the StylesheetArchive object itself as
    // the current component root. ComponentInfo then names that root through
    // `identifier` and does not repeat it in the component UUID map.
    if stylesheet_identifier != stylesheet_selector.identifier() {
        validate_style_registry_entry(
            source,
            payload,
            stylesheet_selector,
            stylesheet_identifier,
            false,
            path,
            budget,
        )?;
    }
    Ok(selector_set)
}

fn rewrite_metadata_for_style(
    source: &Package,
    model_component: usize,
    new_style: u64,
    uuid: UuidBits,
    style_component: usize,
    selector_set: MetadataSelectorSet,
    path: Path,
    budget: &mut TransactionBudget,
) -> Result<NativeEdit, Error> {
    let payload = metadata_payload(source, path)?;
    budget.charge_allocations(selector_set.identifiers.len(), path)?;
    let selectors = selector_set.selectors(path)?;
    let expected_last = selector_set.last_object_identifier;
    let model_selector = selectors
        .first()
        .copied()
        .ok_or(Error::InvalidSource { path })?;
    let style_selector = if model_component == style_component {
        model_selector
    } else {
        selectors
            .get(1)
            .copied()
            .ok_or(Error::InvalidSource { path })?
    };
    budget.charge_allocations(4, path)?;
    let uuid_additions = vec![ObjectUuidAddition::new(style_selector, new_style, uuid)];
    let mut external_additions = Vec::new();
    if model_component != style_component {
        external_additions.push(ExternalReferenceAddition::new(
            model_selector,
            style_selector,
            new_style,
            None,
        ));
    }
    let mut save_selectors = Vec::new();
    save_selectors.push(model_selector);
    if model_component != style_component {
        save_selectors.push(style_selector);
    }
    let metadata_options = budget.metadata_options(
        source,
        payload.len(),
        uuid_additions
            .len()
            .saturating_add(external_additions.len()),
    );
    let batch = AdditionSaveTokenBatch::new(
        MetadataBatch::new(
            expected_last,
            new_style,
            &uuid_additions,
            &external_additions,
        ),
        SaveTokenBatch::new(save_selectors.as_slice()),
    );
    let prepared =
        prepare_package_metadata_additions_and_save_tokens(payload, batch, metadata_options)
            .map_err(|error| map_metadata_error(error, path))?;
    budget.consume_metadata_report(prepared.prepare_report(), path)?;
    let requirements = prepared.execution_requirements();
    budget.preflight_metadata_requirements(requirements, path)?;
    let rewritten = prepared
        .execute(requirements.exact_limits())
        .map_err(|error| map_metadata_error(error, path))?
        .into_bytes();
    rewrite_metadata_entry(source, rewritten, path, budget)
}

fn rewrite_metadata_entry(
    source: &Package,
    payload: Vec<u8>,
    path: Path,
    budget: &mut TransactionBudget,
) -> Result<NativeEdit, Error> {
    let route =
        super::metadata::unique_message_route(source).ok_or(Error::InvalidSource { path })?;
    let physical = physical_source(source)?;
    let entry = physical
        .package()
        .iter()
        .find(|entry| entry.name() == super::metadata::ENTRY_NAME)
        .ok_or(Error::InvalidSource { path })?;
    if entry.is_opaque() {
        return Err(Error::UnsupportedSource);
    }
    budget.charge_transaction_work(entry.data().len(), path)?;
    let archive_limits = physical
        .limits()
        .effective_archive_limits()
        .map_err(|error| map_archive_error(error, path))?;
    let maximum_archive_bytes = archive_limits.max_archive_bytes();
    let maximum_compressed_bytes = SnappyStream::maximum_compressed_len(maximum_archive_bytes)
        .map_err(|error| map_core_error(error, path))?;
    budget.charge_allocations(4, path)?;
    budget.charge_transaction_work(
        maximum_archive_bytes.saturating_add(maximum_compressed_bytes),
        path,
    )?;
    let stream = SnappyStream::decompress_with_limits(
        entry.data(),
        physical
            .limits()
            .snappy_limits()
            .map_err(|error| map_archive_error(error, path))?,
    )
    .map_err(|error| map_core_error(error, path))?;
    let mut archive = Archive::parse_with_limits(stream.as_bytes(), archive_limits)
        .map_err(|error| map_core_error(error, path))?;
    archive
        .validate_canonical_object_framing(stream.as_bytes())
        .map_err(|error| map_core_error(error, path))?;
    let object = archive
        .objects
        .get_mut(route.object_index)
        .ok_or(Error::InvalidSource { path })?;
    let message = object
        .messages
        .get(route.message_index)
        .ok_or(Error::InvalidSource { path })?;
    if message.type_ != super::metadata::MESSAGE_TYPE {
        return Err(Error::InvalidSource { path });
    }
    object
        .replace_message_preserving_header_with_limits(
            route.message_index,
            RawMessage {
                type_: super::metadata::MESSAGE_TYPE,
                data: payload,
            },
            archive_limits,
        )
        .map_err(|error| map_core_error(error, path))?;
    let bytes = archive
        .to_bytes_with_limits(archive_limits)
        .map_err(|error| map_core_error(error, path))?;
    let compressed = SnappyStream::compress(&bytes).map_err(|error| map_core_error(error, path))?;
    Ok(NativeEdit {
        name: super::metadata::ENTRY_NAME.to_owned(),
        data: compressed,
    })
}

fn replace_model_in_archive(
    archive: &mut Archive,
    model_identifier: u64,
    old_style: u64,
    new_style: u64,
    model_type: u32,
    replacement: &[u8],
    limits: litchi_iwa_core::Limits,
    path: Path,
) -> Result<(), Error> {
    let object = archive
        .object_mut(model_identifier)
        .ok_or(Error::InvalidSource { path })?;
    let mut indexes = object
        .messages
        .iter()
        .enumerate()
        .filter(|(_, message)| message.type_ == model_type || message.type_ == 6_000);
    let message_index = indexes
        .next()
        .map(|(index, _)| index)
        .ok_or(Error::InvalidSource { path })?;
    if indexes.next().is_some() {
        return Err(Error::InvalidSource { path });
    }
    let info = object
        .archive_info
        .message_infos
        .get(message_index)
        .ok_or(Error::InvalidSource { path })?;
    let before = info.object_references.clone();
    if before
        .iter()
        .filter(|identifier| **identifier == old_style)
        .count()
        != 1
        || before.contains(&new_style)
    {
        return Err(Error::InvalidSource { path });
    }
    // ArchiveInfo transitions retain surviving aggregate references in their
    // source order and publish newly introduced identifiers as a final suffix.
    // The selected field-local edge still changes directly from old to new.
    let mut after: Vec<u64> = before
        .iter()
        .copied()
        .filter(|identifier| *identifier != old_style)
        .collect();
    after.push(new_style);
    replace_message_with_reference_transition(
        object,
        message_index,
        RawMessage {
            type_: object.messages[message_index].type_,
            data: replacement.to_owned(),
        },
        before.as_slice(),
        after.as_slice(),
        &[3],
        Some((
            std::slice::from_ref(&old_style),
            std::slice::from_ref(&new_style),
        )),
        false,
        limits,
        path,
    )
}

fn append_style_object(
    archive: &mut Archive,
    identifier: u64,
    parent: u64,
    stylesheet: u64,
    payload: &[u8],
    limits: litchi_iwa_core::Limits,
    path: Path,
) -> Result<(), Error> {
    if archive
        .objects
        .iter()
        .any(|object| object.archive_info.identifier == Some(identifier))
    {
        return Err(Error::InvalidSource { path });
    }
    let mut object = ArchiveObject::new_with_limits(
        identifier,
        vec![RawMessage {
            type_: TABLE_STYLE_MESSAGE_TYPE,
            data: payload.to_owned(),
        }],
        limits,
    )
    .map_err(|error| map_core_error(error, path))?;
    object.archive_info.message_infos[0].object_references = vec![parent, stylesheet];
    archive
        .objects
        .try_reserve(1)
        .map_err(|_| Error::Allocation { amount: 1, path })?;
    archive.objects.push(object);
    Ok(())
}

fn replace_stylesheet_in_archive(
    archive: &mut Archive,
    stylesheet_identifier: u64,
    new_style: u64,
    replacement: &[u8],
    limits: litchi_iwa_core::Limits,
    path: Path,
) -> Result<(), Error> {
    let object = archive
        .object_mut(stylesheet_identifier)
        .ok_or(Error::InvalidSource { path })?;
    let mut indexes = object
        .messages
        .iter()
        .enumerate()
        .filter(|(_, message)| message.type_ == STYLESHEET_MESSAGE_TYPE);
    let message_index = indexes
        .next()
        .map(|(index, _)| index)
        .ok_or(Error::UnsupportedDependency { path })?;
    if indexes.next().is_some() {
        return Err(Error::InvalidSource { path });
    }
    let info = object
        .archive_info
        .message_infos
        .get(message_index)
        .ok_or(Error::InvalidSource { path })?;
    let before = info.object_references.clone();
    if before.contains(&new_style) {
        return Err(Error::InvalidSource { path });
    }
    let mut after = before.clone();
    after.push(new_style);
    replace_message_with_reference_transition(
        object,
        message_index,
        RawMessage {
            type_: STYLESHEET_MESSAGE_TYPE,
            data: replacement.to_owned(),
        },
        before.as_slice(),
        after.as_slice(),
        &[1],
        Some((&before, &after)),
        false,
        limits,
        path,
    )
}

fn replace_message_with_reference_transition(
    object: &mut ArchiveObject,
    message_index: usize,
    message: RawMessage,
    before: &[u64],
    after: &[u64],
    field_path: &[u32],
    field_references: Option<(&[u64], &[u64])>,
    required_field_path: bool,
    limits: litchi_iwa_core::Limits,
    path: Path,
) -> Result<(), Error> {
    let field_indices: Vec<usize> = object
        .archive_info
        .message_infos
        .get(message_index)
        .ok_or(Error::InvalidSource { path })?
        .field_infos
        .iter()
        .enumerate()
        .filter(|(_, field)| field.path.as_slice() == field_path)
        .map(|(index, _)| index)
        .collect();
    if field_indices.len() > 1 || (required_field_path && field_indices.len() != 1) {
        return Err(Error::InvalidSource { path });
    }
    let changed_before: Vec<u64> = before
        .iter()
        .copied()
        .filter(|identifier| !after.contains(identifier))
        .collect();
    let changed_after: Vec<u64> = after
        .iter()
        .copied()
        .filter(|identifier| !before.contains(identifier))
        .collect();
    let mut field_before = Vec::new();
    let mut field_after = Vec::new();
    for (index, field) in object.archive_info.message_infos[message_index]
        .field_infos
        .iter()
        .enumerate()
    {
        if field.path.as_slice() != field_path
            && field.object_references.iter().any(|identifier| {
                changed_before.contains(identifier) || changed_after.contains(identifier)
            })
        {
            return Err(Error::InvalidSource { path });
        }
        if !field_indices.contains(&index) {
            continue;
        }
        if field
            .r#type
            .is_some_and(|kind| kind != FieldType::ObjectReference)
        {
            return Err(Error::InvalidSource { path });
        }
        let expected_before = field_references.map(|(before, _)| before).unwrap_or(before);
        let expected_after = field_references.map(|(_, after)| after).unwrap_or(after);
        field_before.push(field.object_references.clone());
        let values = field.object_references.clone();
        if values != expected_before {
            return Err(Error::InvalidSource { path });
        }
        field_after.push(expected_after.to_vec());
    }
    let mut transitions = Vec::new();
    transitions
        .try_reserve_exact(field_indices.len())
        .map_err(|_| Error::Allocation {
            amount: field_indices.len(),
            path,
        })?;
    for index in 0..field_indices.len() {
        transitions.push(FieldObjectReferenceTransition {
            field_info_index: field_indices[index],
            expected_path: field_path,
            before: field_before[index].as_slice(),
            after: field_after[index].as_slice(),
        });
    }
    object
        .replace_message_transitioning_object_references_preserving_header_with_limits(
            message_index,
            message,
            ObjectReferenceTransition {
                aggregate_before: before,
                aggregate_after: after,
                fields: transitions.as_slice(),
            },
            limits,
        )
        .map_err(|error| map_core_error(error, path))?;
    Ok(())
}

fn resolve_target_with_budget<'sheet, 'table>(
    source: &Package,
    sheet: impl Into<SheetSelector<'sheet>>,
    table: impl Into<TableSelector<'table>>,
    budget: &mut TransactionBudget,
    require_metadata: bool,
) -> Result<Target, Error> {
    let selected_sheet = source
        .state
        .document
        .sheet(sheet)
        .map_err(|_| Error::InvalidSource {
            path: Path::Package,
        })?
        .ok_or(Error::SheetNotFound)?;
    let table_position = match table.into() {
        TableSelector::Index(index) => selected_sheet.tables().nth(index).map(|_| index),
        TableSelector::Name(name) => {
            let mut matches = selected_sheet
                .tables()
                .enumerate()
                .filter(|(_, candidate)| candidate.name() == name);
            let first = matches.next().map(|(index, _)| index);
            if matches.next().is_some() {
                return Err(Error::InvalidSource {
                    path: Path::Table {
                        sheet: selected_sheet.index(),
                        table: 0,
                    },
                });
            }
            first
        },
    }
    .ok_or(Error::TableNotFound)?;
    resolve_at_with_budget(
        source,
        selected_sheet.index(),
        table_position,
        budget,
        require_metadata,
    )
}

fn resolve_at_with_budget(
    source: &Package,
    sheet: usize,
    table: usize,
    budget: &mut TransactionBudget,
    require_metadata: bool,
) -> Result<Target, Error> {
    let path = Path::Table { sheet, table };
    let native = super::table_headers::resolve::resolve_target(source, sheet, table)
        .map_err(|_| Error::InvalidSource { path })?;
    ensure_unique_physical_identifier(source, native.model_identifier, path)?;
    budget.charge_transaction_work(source.source_bytes().len(), path)?;
    let payload = super::table_headers::rewrite::selected_payload(source, native)
        .map_err(|_| Error::InvalidSource { path })?;
    let (style_identifier, appearance) = resolve_codec_appearance(
        source,
        payload,
        native.component_index,
        path,
        budget,
        require_metadata,
    )?;
    Ok(Target {
        native,
        appearance,
        style_identifier,
    })
}

fn resolve_codec_appearance(
    source: &Package,
    model_payload: &[u8],
    model_component: usize,
    path: Path,
    budget: &mut TransactionBudget,
    require_metadata: bool,
) -> Result<(Option<u64>, Appearance), Error> {
    // The model projection itself is part of the one aggregate wire budget.
    // DecodeReport is consumed even when the semantic projection is empty.
    let (model, model_report) =
        codec::decode_table_model_with_report(model_payload, budget.codec_options(model_payload))
            .map_err(|error| map_codec_error(error, path))?;
    budget.consume_report(model_report, path)?;
    let style_identifier = model.style_identifier();
    // Native models commonly retain a nonzero preset alongside the direct
    // style edge.  The direct edge is authoritative; a preset-only model is
    // outside this focused COW owner and fails closed.
    if style_identifier == 0 {
        return Err(Error::UnsupportedDependency { path });
    }
    let mut nodes = Vec::new();
    budget.charge_allocations(8, path)?;
    let mut current = Some(style_identifier);
    let mut stylesheet = None;
    for _ in 0..MAX_INHERITANCE_DEPTH {
        let Some(identifier) = current else { break };
        ensure_unique_physical_identifier(source, identifier, path)?;
        budget.charge_styles(1, path)?;
        if nodes
            .iter()
            .any(|node: &codec::TableStyleNode<'_>| node.identifier() == identifier)
        {
            return Err(Error::InvalidSource { path });
        }
        let style_payload = style_message_data(source, identifier, path)?;
        let (style, style_report) = codec::decode_table_style_with_report(
            style_payload,
            budget.codec_options(style_payload),
        )
        .map_err(|error| map_codec_error(error, path))?;
        budget.consume_report(style_report, path)?;
        let style_sheet = style
            .stylesheet_identifier()
            .filter(|identifier| *identifier != 0)
            .ok_or(Error::UnsupportedDependency { path })?;
        if style_sheet == 0 {
            return Err(Error::UnsupportedDependency { path });
        }
        if let Some(previous) = stylesheet {
            if previous != style_sheet {
                return Err(Error::UnsupportedDependency { path });
            }
        } else {
            stylesheet = Some(style_sheet);
        }
        current = style.parent_identifier();
        nodes.try_reserve(1).map_err(|_| Error::Allocation {
            amount: nodes.len().saturating_add(1),
            path,
        })?;
        nodes.push(codec::TableStyleNode::new(identifier, style));
    }
    if current.is_some() {
        return Err(Error::InvalidSource { path });
    }
    let stylesheet_identifier = stylesheet.ok_or(Error::UnsupportedDependency { path })?;
    let style_location = resolved_location(source, style_identifier, path)?;
    let stylesheet_location = resolved_location(source, stylesheet_identifier, path)?;
    ensure_unique_physical_identifier(source, stylesheet_identifier, path)?;
    let _stylesheet_payload = stylesheet_message_data(source, stylesheet_identifier, path)?;
    if style_location.0 != stylesheet_location.0 {
        return Err(Error::UnsupportedDependency { path });
    }
    if require_metadata {
        let _ =
            preflight_metadata_for_style(source, model_component, style_identifier, path, budget)?;
    }
    budget.charge_fields(nodes.len(), path)?;
    budget.charge_work(nodes.len().saturating_mul(model_payload.len()), path)?;
    let effective = codec::resolve_table_style_appearance(
        &nodes,
        style_identifier,
        codec::DecodeOptions::for_source(model_payload),
    )
    .map_err(|error| map_codec_error(error, path))?;
    Ok((Some(style_identifier), snapshot_appearance(effective)))
}

fn style_message_data(source: &Package, identifier: u64, path: Path) -> Result<&[u8], Error> {
    let resolved = source
        .state
        .index
        .resolve_ref_id(&source.state.components, identifier)
        .map_err(|_| Error::InvalidSource { path })?
        .ok_or(Error::InvalidSource { path })?;
    let object = source
        .state
        .components
        .catalog()
        .get_index(resolved.component_index)
        .and_then(|component| component.archive().objects.get(resolved.object_index))
        .ok_or(Error::InvalidSource { path })?;
    if object.archive_info.identifier != Some(identifier) {
        return Err(Error::InvalidSource { path });
    }
    let mut messages = object
        .messages
        .iter()
        .filter(|message| message.type_ == TABLE_STYLE_MESSAGE_TYPE);
    let message = messages
        .next()
        .ok_or(Error::UnsupportedDependency { path })?;
    if messages.next().is_some() {
        return Err(Error::InvalidSource { path });
    }
    Ok(message.data.as_slice())
}

fn ensure_unique_physical_identifier(
    source: &Package,
    identifier: u64,
    path: Path,
) -> Result<(), Error> {
    let matches = source
        .state
        .components
        .catalog()
        .iter()
        .flat_map(|component| component.archive().objects.iter())
        .filter(|object| object.archive_info.identifier == Some(identifier))
        .count();
    if matches != 1 {
        return Err(Error::InvalidSource { path });
    }
    Ok(())
}

fn physical_source(source: &Package) -> Result<&litchi_iwa_archive::SourceCatalog, Error> {
    source
        .state
        .components
        .physical()
        .ok_or(Error::UnsupportedSource)
}

fn verify_candidate_locality(
    source: &Package,
    candidate: &Package,
    target: Target,
    old_style: Option<u64>,
    new_style: Option<u64>,
    source_previews: usize,
    target_previews: usize,
    path: Path,
) -> Result<(), Error> {
    let source_catalog = physical_source(source)?;
    let candidate_catalog = physical_source(candidate)?;
    let source_preview_names =
        super::table_headers::rewrite::root_preview_deletions(source_catalog)
            .map_err(|_| Error::Verification)?;
    let candidate_preview_names =
        super::table_headers::rewrite::root_preview_deletions(candidate_catalog)
            .map_err(|_| Error::Verification)?;
    if source_preview_names.len() != source_previews
        || candidate_preview_names.len() != target_previews
    {
        return Err(Error::Verification);
    }

    let mut component_indices = vec![target.native.component_index];
    if let Some(identifier) = old_style {
        component_indices.push(resolved_location(source, identifier, path)?.0);
    }
    if let Some(identifier) = new_style {
        component_indices.push(resolved_location(candidate, identifier, path)?.0);
    }
    let mut allowed_names = Vec::new();
    for component_index in component_indices {
        let Some(component) = source
            .state
            .components
            .catalog()
            .get_index(component_index)
            .or_else(|| {
                candidate
                    .state
                    .components
                    .catalog()
                    .get_index(component_index)
            })
        else {
            return Err(Error::Verification);
        };
        if !allowed_names
            .iter()
            .any(|name: &String| name == component.name())
        {
            allowed_names.push(component.name().to_owned());
        }
    }
    if !allowed_names
        .iter()
        .any(|name| name == super::metadata::ENTRY_NAME)
    {
        allowed_names.push(super::metadata::ENTRY_NAME.to_owned());
    }

    let mut before_entries = source_catalog
        .package()
        .iter()
        .filter(|entry| !source_preview_names.contains(&entry.name()));
    let mut after_entries = candidate_catalog
        .package()
        .iter()
        .filter(|entry| !candidate_preview_names.contains(&entry.name()));
    loop {
        match (before_entries.next(), after_entries.next()) {
            (Some(before), Some(after)) if before.name() == after.name() => {
                if allowed_names.iter().any(|name| name == before.name()) {
                    if !super::table_headers::rewrite::selected_package_member_preserved(
                        before, after,
                    ) {
                        return Err(Error::Verification);
                    }
                } else if !super::table_headers::rewrite::package_member_preserved(before, after) {
                    return Err(Error::Verification);
                }
            },
            (None, None) => break,
            _ => return Err(Error::Verification),
        }
    }
    Ok(())
}

fn map_archive_error(error: litchi_iwa_archive::Error, path: Path) -> Error {
    match error {
        litchi_iwa_archive::Error::Limit {
            kind,
            observed,
            maximum,
        } => {
            let kind = match kind {
                litchi_iwa_archive::LimitKind::InputBytes => LimitKind::InputBytes,
                litchi_iwa_archive::LimitKind::OutputBytes => LimitKind::OutputBytes,
                litchi_iwa_archive::LimitKind::Entries => LimitKind::Entries,
                litchi_iwa_archive::LimitKind::MemberNameBytes
                | litchi_iwa_archive::LimitKind::MetadataBytes => LimitKind::PackageBytes,
                litchi_iwa_archive::LimitKind::CompressedEntryBytes
                | litchi_iwa_archive::LimitKind::EntryBytes => LimitKind::EntryBytes,
                litchi_iwa_archive::LimitKind::TotalBytes => LimitKind::TotalEntryBytes,
                litchi_iwa_archive::LimitKind::IwaStreamBytes => LimitKind::PayloadBytes,
                litchi_iwa_archive::LimitKind::IwaTotalBytes => LimitKind::TotalPayloadBytes,
            };
            Error::LimitExceeded {
                kind,
                observed,
                maximum,
                path,
            }
        },
        litchi_iwa_archive::Error::Allocation { amount, .. } => Error::Allocation { amount, path },
        litchi_iwa_archive::Error::Iwa(error) => map_core_error(error, path),
        _ => Error::InvalidSource { path },
    }
}

fn map_core_error(error: litchi_iwa_core::Error, path: Path) -> Error {
    match error {
        litchi_iwa_core::Error::Limit {
            kind,
            observed,
            maximum,
        } => {
            let kind = match kind {
                litchi_iwa_core::LimitKind::ArchiveBytes
                | litchi_iwa_core::LimitKind::ObjectBytes
                | litchi_iwa_core::LimitKind::MessageBytes
                | litchi_iwa_core::LimitKind::HeaderBytes
                | litchi_iwa_core::LimitKind::HeaderMemoryBytes
                | litchi_iwa_core::LimitKind::SnappyChunkBytes
                | litchi_iwa_core::LimitKind::SnappyStreamBytes
                | litchi_iwa_core::LimitKind::SnappyCompressedChunkBytes
                | litchi_iwa_core::LimitKind::SnappyCompressedStreamBytes => {
                    LimitKind::PayloadBytes
                },
                litchi_iwa_core::LimitKind::Objects => LimitKind::PayloadObjects,
                litchi_iwa_core::LimitKind::Messages
                | litchi_iwa_core::LimitKind::MessagesPerObject => LimitKind::PayloadMessages,
                litchi_iwa_core::LimitKind::HeaderFields
                | litchi_iwa_core::LimitKind::MetadataItems
                | litchi_iwa_core::LimitKind::SnappyFrames => LimitKind::PayloadItems,
                litchi_iwa_core::LimitKind::HeaderNesting => LimitKind::WireNesting,
            };
            Error::LimitExceeded {
                kind,
                observed: u64::try_from(observed).unwrap_or(u64::MAX),
                maximum: u64::try_from(maximum).unwrap_or(u64::MAX),
                path,
            }
        },
        litchi_iwa_core::Error::Allocation { requested, .. } => Error::Allocation {
            amount: requested,
            path,
        },
        _ => Error::InvalidSource { path },
    }
}

#[cfg(all(test, feature = "internal-iwork-source"))]
mod source_built_tests {
    use super::*;
    use crate::table::appearance::{Banding, GridlineVisibility, Gridlines, RowSizing};

    fn varint(mut value: u64) -> Vec<u8> {
        let mut bytes = Vec::new();
        while value >= 0x80 {
            bytes.push((value as u8) | 0x80);
            value >>= 7;
        }
        bytes.push(value as u8);
        bytes
    }

    fn field_varint(field: u32, value: u64) -> Vec<u8> {
        let mut bytes = varint(u64::from(field) << 3);
        bytes.extend_from_slice(&varint(value));
        bytes
    }

    fn field_bytes(field: u32, payload: &[u8]) -> Vec<u8> {
        let mut bytes = varint((u64::from(field) << 3) | 2);
        bytes.extend_from_slice(&varint(payload.len() as u64));
        bytes.extend_from_slice(payload);
        bytes
    }

    fn reference(identifier: u64) -> Vec<u8> {
        field_varint(1, identifier)
    }

    fn style_payload(parent: Option<u64>, values: [Option<bool>; 7]) -> Vec<u8> {
        let mut style_super = Vec::new();
        if let Some(parent) = parent {
            style_super.extend_from_slice(&field_bytes(3, &reference(parent)));
        }
        let mut properties = Vec::new();
        for (field, value) in [
            (1, values[0]),
            (22, values[1]),
            (33, values[2]),
            (34, values[3]),
            (42, values[4]),
            (43, values[5]),
            (44, values[6]),
        ] {
            if let Some(value) = value {
                properties.extend_from_slice(&field_varint(field, u64::from(value)));
            }
        }
        let mut style = field_bytes(1, &style_super);
        if !properties.is_empty() {
            style.extend_from_slice(&field_varint(10, 7));
            style.extend_from_slice(&field_bytes(11, &properties));
        }
        style
    }

    fn preset_payload(network_identifier: u64) -> Vec<u8> {
        field_bytes(3, &reference(network_identifier))
    }

    fn network_payload(style_identifier: u64) -> Vec<u8> {
        let mut network = Vec::new();
        for field in 1..=8 {
            network.extend_from_slice(&field_bytes(field, &reference(field as u64 + 100)));
        }
        network.extend_from_slice(&field_bytes(9, &reference(style_identifier)));
        network
    }

    fn expected(
        row_banding: Banding,
        row_sizing: RowSizing,
        visible: [GridlineVisibility; 5],
    ) -> Appearance {
        Appearance {
            row_banding,
            row_sizing,
            gridlines: Gridlines {
                body_horizontal: visible[0],
                body_vertical: visible[1],
                header_columns_horizontal: visible[2],
                header_rows_vertical: visible[3],
                footer_rows_vertical: visible[4],
            },
        }
    }

    #[test]
    fn source_built_missing_style_edges_keep_native_defaults() {
        let appearance =
            Package::__table_appearance_from_source_built(0, None, |_, _, _| -> Result<(), ()> {
                panic!("default appearance must not request a payload")
            })
            .expect("missing style edges should use defaults");
        assert_eq!(appearance, Appearance::default());
    }

    #[test]
    fn source_built_inheritance_and_early_completion_are_bounded() {
        let child = style_payload(Some(8), [Some(true), None, None, None, None, None, None]);
        let parent = style_payload(
            None,
            [
                None,
                Some(true),
                Some(false),
                Some(false),
                Some(false),
                Some(true),
                Some(false),
            ],
        );
        let appearance =
            Package::__table_appearance_from_source_built(7, None, |identifier, kind, consume| {
                assert_eq!(kind, SourceBuiltAppearancePayload::TableStyle);
                match identifier {
                    7 => consume(&child),
                    8 => consume(&parent),
                    _ => panic!("unexpected style lookup"),
                }
                Ok::<_, ()>(())
            })
            .expect("parent style should complete inherited appearance");
        assert_eq!(
            appearance,
            expected(
                Banding::Enabled,
                RowSizing::FitCellContents,
                [
                    GridlineVisibility::Hidden,
                    GridlineVisibility::Hidden,
                    GridlineVisibility::Hidden,
                    GridlineVisibility::Visible,
                    GridlineVisibility::Hidden,
                ],
            )
        );

        let complete_child = style_payload(
            None,
            [
                Some(false),
                Some(false),
                Some(true),
                Some(true),
                Some(true),
                Some(true),
                Some(true),
            ],
        );
        let appearance =
            Package::__table_appearance_from_source_built(9, None, |identifier, kind, consume| {
                assert_eq!(identifier, 9);
                assert_eq!(kind, SourceBuiltAppearancePayload::TableStyle);
                consume(&complete_child);
                Ok::<_, ()>(())
            })
            .expect("complete child should not require a dangling parent");
        assert_eq!(
            appearance,
            expected(
                Banding::Disabled,
                RowSizing::Fixed,
                [
                    GridlineVisibility::Visible,
                    GridlineVisibility::Visible,
                    GridlineVisibility::Visible,
                    GridlineVisibility::Visible,
                    GridlineVisibility::Visible,
                ],
            )
        );
    }

    #[test]
    fn source_built_preset_network_resolves_style_without_direct_edge() {
        let preset = preset_payload(17);
        let network = network_payload(9);
        let style = style_payload(
            None,
            [
                Some(true),
                Some(true),
                Some(false),
                Some(true),
                Some(false),
                Some(true),
                Some(false),
            ],
        );
        let appearance = Package::__table_appearance_from_source_built(
            0,
            Some(16),
            |identifier, kind, consume| {
                match (identifier, kind) {
                    (16, SourceBuiltAppearancePayload::TableStylePreset) => consume(&preset),
                    (17, SourceBuiltAppearancePayload::TableStyleNetwork) => consume(&network),
                    (9, SourceBuiltAppearancePayload::TableStyle) => consume(&style),
                    _ => panic!("unexpected preset style lookup"),
                }
                Ok::<_, ()>(())
            },
        )
        .expect("preset network should resolve a table style");
        assert_eq!(
            appearance,
            expected(
                Banding::Enabled,
                RowSizing::FitCellContents,
                [
                    GridlineVisibility::Hidden,
                    GridlineVisibility::Visible,
                    GridlineVisibility::Hidden,
                    GridlineVisibility::Visible,
                    GridlineVisibility::Hidden,
                ],
            )
        );
    }

    #[test]
    fn source_built_cycle_is_rejected() {
        let first = style_payload(Some(2), [None, None, None, None, None, None, None]);
        let second = style_payload(Some(1), [None, None, None, None, None, None, None]);
        let result =
            Package::__table_appearance_from_source_built(1, None, |identifier, kind, consume| {
                assert_eq!(kind, SourceBuiltAppearancePayload::TableStyle);
                match identifier {
                    1 => consume(&first),
                    2 => consume(&second),
                    _ => panic!("unexpected cyclic style lookup"),
                }
                Ok::<_, ()>(())
            });
        assert!(matches!(
            result,
            Err(Error::InvalidSource {
                path: Path::Package
            })
        ));
    }

    #[test]
    fn source_built_sixty_four_empty_styles_are_accepted() {
        let styles = (1_u64..=64)
            .map(|identifier| {
                let parent = (identifier < 64).then_some(identifier + 1);
                style_payload(parent, [None, None, None, None, None, None, None])
            })
            .collect::<Vec<_>>();
        let appearance =
            Package::__table_appearance_from_source_built(1, None, |identifier, kind, consume| {
                assert_eq!(kind, SourceBuiltAppearancePayload::TableStyle);
                let payload = styles
                    .get(usize::try_from(identifier).expect("test identifier") - 1)
                    .expect("bounded style lookup");
                consume(payload);
                Ok::<_, ()>(())
            })
            .expect("64-style inheritance chain should remain within the bound");
        assert_eq!(appearance, Appearance::default());
    }

    #[test]
    fn source_built_sixty_five_style_chain_is_rejected() {
        let styles = (1_u64..=65)
            .map(|identifier| {
                let parent = (identifier < 65).then_some(identifier + 1);
                style_payload(parent, [None, None, None, None, None, None, None])
            })
            .collect::<Vec<_>>();
        let result =
            Package::__table_appearance_from_source_built(1, None, |identifier, kind, consume| {
                assert_eq!(kind, SourceBuiltAppearancePayload::TableStyle);
                let payload = styles
                    .get(usize::try_from(identifier).expect("test identifier") - 1)
                    .expect("bounded style lookup");
                consume(payload);
                Ok::<_, ()>(())
            });
        assert!(matches!(
            result,
            Err(Error::InvalidSource {
                path: Path::Package
            })
        ));
    }

    #[test]
    fn source_built_malformed_style_payload_is_rejected() {
        let result =
            Package::__table_appearance_from_source_built(1, None, |identifier, kind, consume| {
                assert_eq!(identifier, 1);
                assert_eq!(kind, SourceBuiltAppearancePayload::TableStyle);
                consume(&[0x08]);
                Ok::<_, ()>(())
            });
        assert!(result.is_err(), "malformed style payload must be rejected");
    }
}
