//! Exact-source transactions for one rooted Numbers table title.

use std::{
    fmt,
    mem::{size_of, size_of_val},
};

use litchi_iwa_archive::package::OwnedExactArtifacts;
use litchi_iwa_common::{WireLimits, wire::WireView};
use litchi_iwa_protos::{numbers_table_title_codec, table_info_codec};
use thiserror::Error as ThisError;

use super::Package;
use crate::{
    selector::{SheetSelector, TableSelector},
    table::lock::State as LockState,
    table::title::Settings,
};

const PARAGRAPH_STYLE_MESSAGE_TYPE: u32 = 2_022;
const SHAPE_STYLE_MESSAGE_TYPE: u32 = 2_025;

mod rewrite;

/// A content-free semantic location associated with a title transaction.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Path {
    /// The complete Numbers package.
    Package,
    /// One rooted table at checked zero-based positions.
    Table {
        /// Zero-based rooted sheet position.
        sheet: usize,
        /// Zero-based table position within the sheet.
        table: usize,
    },
}

/// A finite resource governed by a table-title transaction.
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
    /// Aggregate member bytes.
    TotalEntryBytes,
    /// Physical container names and metadata.
    PackageBytes,
    /// Bytes in one decoded payload container.
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
    /// Bytes inspected by a strict wire projection.
    WireBytes,
    /// Bytes emitted by a wire rewrite.
    WireOutputBytes,
    /// Protobuf fields inspected.
    WireFields,
    /// Protobuf nesting depth.
    WireNesting,
    /// Aggregate strict wire work.
    WireWork,
    /// Aggregate focused transaction work.
    TransactionWork,
}

impl fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(formatter, "{self:?}")
    }
}

/// A content-redacted table-title transaction failure.
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// No rooted sheet matched the selector.
    #[error("the Numbers workbook has no sheet matching the requested selector")]
    SheetNotFound,
    /// No table on the selected sheet matched the selector.
    #[error("the selected Numbers sheet has no table matching the requested selector")]
    TableNotFound,
    /// A changed edit targeted an effectively locked table.
    #[error("the selected Numbers table is locked at {path:?}")]
    TableLocked {
        /// Selected semantic table.
        path: Path,
    },
    /// A visible title is missing a required native rendering dependency.
    #[error("the selected Numbers table title is missing a required dependency at {path:?}")]
    UnsupportedDependency {
        /// Selected semantic table.
        path: Path,
    },
    /// The exact native profile is not supported for changed publication.
    #[error("this Numbers source does not support exact table-title editing")]
    UnsupportedSource,
    /// Rooted ownership, metadata, or wire framing is invalid.
    #[error("the Numbers table-title source is invalid at {path:?}")]
    InvalidSource {
        /// Content-free semantic failure location.
        path: Path,
    },
    /// A finite transaction resource ceiling was exceeded.
    #[error("Numbers table-title {kind} limit exceeded: observed {observed}, maximum {maximum}")]
    LimitExceeded {
        /// Resource category.
        kind: LimitKind,
        /// Observed amount.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
        /// Content-free semantic failure location.
        path: Path,
    },
    /// A bounded allocation failed before publication.
    #[error("could not allocate {amount} units for the Numbers table-title transaction")]
    Allocation {
        /// Requested bytes or elements.
        amount: usize,
        /// Content-free semantic failure location.
        path: Path,
    },
    /// Candidate reopening or locality verification failed.
    #[error("the edited Numbers table title failed semantic verification")]
    Verification,
    /// The patch was applied to a package other than its exact source.
    #[error("the table-title patch does not match the exact source package")]
    PatchConflict,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct Target {
    native: super::table_headers::Target,
    settings: Settings,
    height_bits: Option<u64>,
    paragraph_style: Option<u64>,
    shape_style: Option<u64>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct ReopenCost {
    work: usize,
    references: usize,
}

/// Immutable settings staged against one package snapshot.
pub struct Edit<'a> {
    source: &'a Package,
    target: Target,
    before: Settings,
    settings: Settings,
    budget: TransactionBudget,
}

impl fmt::Debug for Edit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("path", &self.path())
            .field("settings", &self.settings)
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

    /// Return the settings that would be published.
    #[must_use]
    pub const fn settings(&self) -> Settings {
        self.settings
    }

    /// Replace the staged settings without touching package bytes.
    #[must_use]
    pub fn set(mut self, settings: Settings) -> Self {
        self.settings = settings;
        self
    }

    /// Validate and atomically publish the staged title settings.
    ///
    /// # Errors
    ///
    /// Returns a typed source, lock, resource, or verification error.
    pub fn commit(self) -> Result<Commit, Error> {
        commit_edit(self)
    }
}

/// A reversible, process-local exact-source title patch.
#[derive(Clone, PartialEq, Eq)]
pub struct Patch {
    artifacts: OwnedExactArtifacts,
    target: Target,
    before: Settings,
    after: Settings,
    source_reopen: ReopenCost,
    target_reopen: ReopenCost,
    source_previews: usize,
    target_previews: usize,
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
    /// Return the selected semantic table path.
    #[must_use]
    pub const fn path(&self) -> Path {
        Path::Table {
            sheet: self.target.native.sheet_position,
            table: self.target.native.table_position,
        }
    }

    /// Return exact source settings.
    #[must_use]
    pub const fn before(&self) -> Settings {
        self.before
    }

    /// Return exact target settings.
    #[must_use]
    pub const fn after(&self) -> Settings {
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

    /// Return whether semantics and retained artifacts are unchanged.
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
            source_reopen: self.target_reopen,
            target_reopen: self.source_reopen,
            source_previews: self.target_previews,
            target_previews: self.source_previews,
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

    /// Number of canonical previews deleted in this direction.
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
#[must_use = "a table-title commit contains the validated package snapshot"]
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

impl Package {
    /// Read one rooted table's lossless title settings.
    pub fn table_title_settings<'sheet, 'table>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        table: impl Into<TableSelector<'table>>,
    ) -> Result<Settings, Error> {
        Ok(resolve_title_target(self, sheet, table)?.settings)
    }

    /// Start a selector-first immutable table-title edit.
    pub fn edit_table_title<'sheet, 'table>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        table: impl Into<TableSelector<'table>>,
    ) -> Result<Edit<'_>, Error> {
        let mut budget = TransactionBudget::new(self);
        budget.charge_transaction_work(self.source_bytes().len().saturating_mul(2))?;
        let (sheet_position, table_position) = resolve_title_selectors(self, sheet, table)?;
        let target = resolve_title_at_positions_with_budget(
            self,
            sheet_position,
            table_position,
            &mut budget,
        )?;
        Ok(Edit {
            source: self,
            before: target.settings,
            settings: target.settings,
            target,
            budget,
        })
    }

    /// Apply a reversible exact-source title patch.
    pub fn apply_table_title(&self, patch: &Patch) -> Result<Commit, Error> {
        apply_patch(self, patch)
    }
}

fn resolve_title_target<'sheet, 'table>(
    source: &Package,
    sheet: impl Into<SheetSelector<'sheet>>,
    table: impl Into<TableSelector<'table>>,
) -> Result<Target, Error> {
    let (sheet_position, table_position) = resolve_title_selectors(source, sheet, table)?;
    resolve_title_at_positions(source, sheet_position, table_position)
}

fn resolve_title_selectors<'sheet, 'table>(
    source: &Package,
    sheet: impl Into<SheetSelector<'sheet>>,
    table: impl Into<TableSelector<'table>>,
) -> Result<(usize, usize), Error> {
    let selected_sheet = source
        .state
        .document
        .sheet(sheet)
        .map_err(|_error| Error::InvalidSource {
            path: Path::Package,
        })?
        .ok_or(Error::SheetNotFound)?;
    let table_position = match table.into() {
        TableSelector::Index(index) => selected_sheet.tables().nth(index).map(|_| index),
        TableSelector::Name(name) => selected_sheet
            .tables()
            .position(|candidate| candidate.name() == name),
    }
    .ok_or(Error::TableNotFound)?;
    Ok((selected_sheet.index(), table_position))
}

fn resolve_title_at_positions(
    source: &Package,
    sheet_position: usize,
    table_position: usize,
) -> Result<Target, Error> {
    let native = resolve_native_title_target(source, sheet_position, table_position, None)?;
    let payload = super::table_headers::rewrite::selected_payload(source, native)
        .map_err(map_header_error)?;
    let (decoded, _decode_report) = decode_title_with_report(payload)?;
    Ok(Target {
        native,
        settings: decoded.settings,
        height_bits: decoded.height_bits,
        paragraph_style: decoded.paragraph_style,
        shape_style: decoded.shape_style,
    })
}

fn resolve_title_at_positions_with_budget(
    source: &Package,
    sheet_position: usize,
    table_position: usize,
    budget: &mut TransactionBudget,
) -> Result<Target, Error> {
    let native = resolve_native_title_target(source, sheet_position, table_position, Some(budget))?;
    let payload = super::table_headers::rewrite::selected_payload(source, native)
        .map_err(map_header_error)?;
    let (decoded, decode_report) = decode_title_with_options(payload, budget.codec_options())?;
    budget.consume(decode_report)?;
    Ok(Target {
        native,
        settings: decoded.settings,
        height_bits: decoded.height_bits,
        paragraph_style: decoded.paragraph_style,
        shape_style: decoded.shape_style,
    })
}

fn resolve_native_title_target(
    source: &Package,
    sheet_position: usize,
    table_position: usize,
    mut budget: Option<&mut TransactionBudget>,
) -> Result<super::table_headers::Target, Error> {
    use super::table_headers::resolve as topology;

    let document_object = source
        .state
        .components
        .get_archive("Index/Document.iwa")
        .and_then(|archive| archive.object(1))
        .ok_or_else(invalid_source)?;
    let (_document_index, document_message) =
        topology::unique_message_index(&document_object.messages, super::DOCUMENT_MESSAGE_TYPE)
            .map_err(map_header_error)?
            .ok_or_else(invalid_source)?;
    if let Some(active) = budget.as_deref_mut() {
        active.charge_fields(document_message.data.len())?;
        active.charge_transaction_work(document_message.data.len().saturating_mul(2))?;
    }
    let sheet_payloads =
        topology::repeated_length_payloads(&document_message.data, 1).map_err(map_header_error)?;
    let sheet_identifier = topology::local_reference_identifier(
        sheet_payloads
            .get(sheet_position)
            .copied()
            .ok_or(Error::SheetNotFound)?,
    )
    .map_err(map_header_error)?;
    if let Some(active) = budget.as_deref_mut() {
        active.charge_index_lookup(source)?;
    }
    let sheet = source
        .state
        .index
        .resolve_ref_id(&source.state.components, sheet_identifier)
        .map_err(|_error| invalid_source())?
        .ok_or_else(invalid_source)?;
    let sheet_message_index =
        topology::unique_sheet_message_index(sheet.messages).map_err(map_header_error)?;
    let sheet_message = sheet
        .messages
        .get(sheet_message_index)
        .ok_or_else(invalid_source)?;
    if let Some(active) = budget.as_deref_mut() {
        active.charge_fields(sheet_message.data.len())?;
        active.charge_transaction_work(sheet_message.data.len().saturating_mul(4))?;
    }
    let drawable_payloads =
        topology::sheet_drawable_payloads(sheet_message.type_, &sheet_message.data)
            .map_err(map_header_error)?;
    let mut semantic_table = 0usize;
    for (drawable_position, drawable_payload) in drawable_payloads.iter().enumerate() {
        if let Some(active) = budget.as_deref_mut() {
            active.charge_transaction_work(drawable_payload.len().saturating_add(1))?;
        }
        let drawable_identifier =
            topology::local_reference_identifier(drawable_payload).map_err(map_header_error)?;
        if let Some(active) = budget.as_deref_mut() {
            active.charge_index_lookup(source)?;
        }
        let info = source
            .state
            .index
            .resolve_ref_id(&source.state.components, drawable_identifier)
            .map_err(|_error| invalid_source())?
            .ok_or_else(invalid_source)?;
        let Some((info_message_index, info_message)) =
            topology::unique_table_info(info).map_err(map_header_error)?
        else {
            continue;
        };
        if let Some(active) = budget.as_deref_mut() {
            active.charge_fields(info_message.data.len())?;
            active.charge_work(info_message.data.len().saturating_mul(4))?;
            active.charge_transaction_work(info_message.data.len().saturating_mul(4))?;
        }
        let info_snapshot = table_info_codec::decode_table_info(
            &info_message.data,
            super::table_info_decode_options(&info_message.data),
        )
        .map_err(|_error| invalid_source())?;
        if semantic_table != table_position {
            semantic_table = semantic_table.checked_add(1).ok_or_else(invalid_source)?;
            continue;
        }
        let model_identifier = info_snapshot.table_model().identifier().get();
        if let Some(active) = budget.as_deref_mut() {
            active.charge_index_lookup(source)?;
        }
        let model = source
            .state
            .index
            .resolve_ref_id(&source.state.components, model_identifier)
            .map_err(|_error| invalid_source())?
            .ok_or_else(invalid_source)?;
        let (message_index, message) =
            topology::unique_table_model(model.messages).map_err(map_header_error)?;
        if sheet_identifier == drawable_identifier
            || sheet_identifier == model_identifier
            || drawable_identifier == model_identifier
        {
            return Err(invalid_source());
        }
        return Ok(super::table_headers::Target {
            sheet_position,
            table_position,
            model_identifier,
            sheet_identifier,
            drawable_identifier,
            drawable_position,
            sheet_component_index: sheet.component_index,
            sheet_object_index: sheet.object_index,
            sheet_message_index,
            sheet_message_type: sheet_message.type_,
            info_component_index: info.component_index,
            info_object_index: info.object_index,
            info_message_index,
            info_message_type: info_message.type_,
            component_index: model.component_index,
            object_index: model.object_index,
            message_index,
            message_type: message.type_,
            settings: crate::table::headers::Settings::default(),
            rows: 0,
            columns: 0,
            locked: LockState::from_locked(info_snapshot.locked().unwrap_or(false)),
        });
    }
    Err(Error::TableNotFound)
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct DecodedTitle {
    settings: Settings,
    height_bits: Option<u64>,
    paragraph_style: Option<u64>,
    shape_style: Option<u64>,
}

#[derive(Debug)]
struct TransactionBudget {
    max_message_bytes: usize,
    recursion_limit: u32,
    maximum_fields: usize,
    maximum_work: usize,
    maximum_references: usize,
    maximum_transaction_work: usize,
    remaining_fields: usize,
    remaining_work: usize,
    remaining_references: usize,
    remaining_transaction_work: usize,
}

impl TransactionBudget {
    fn new(source: &Package) -> Self {
        let archive = source.state.options.archive();
        let wire = archive.max_iwa_stream_bytes().max(1);
        let maximum_fields = wire.saturating_mul(5).max(1);
        let maximum_work = wire.saturating_mul(12).max(1);
        let maximum_references = source
            .state
            .options
            .semantic()
            .max_references()
            .saturating_mul(4);
        let maximum_transaction_work = usize::try_from(archive.max_total_bytes())
            .unwrap_or(usize::MAX)
            .saturating_mul(32);
        Self {
            max_message_bytes: wire,
            recursion_limit: u32::try_from(WireLimits::MAX_NESTING).unwrap_or(u32::MAX),
            maximum_fields,
            maximum_work,
            maximum_references,
            maximum_transaction_work,
            remaining_fields: maximum_fields,
            remaining_work: maximum_work,
            remaining_references: maximum_references,
            remaining_transaction_work: maximum_transaction_work,
        }
    }

    const fn codec_options(&self) -> numbers_table_title_codec::DecodeOptions {
        numbers_table_title_codec::DecodeOptions::new(
            self.max_message_bytes,
            self.remaining_fields,
            self.remaining_work,
            self.recursion_limit,
            self.remaining_references,
        )
    }

    fn consume(&mut self, report: numbers_table_title_codec::DecodeReport) -> Result<(), Error> {
        self.charge_fields(report.fields())?;
        self.charge_work(report.work_bytes())?;
        self.charge_references(report.references())?;
        if report.max_depth() > u32::try_from(WireLimits::MAX_NESTING).unwrap_or(u32::MAX) {
            return Err(Error::LimitExceeded {
                kind: LimitKind::WireNesting,
                observed: u64::from(report.max_depth()),
                maximum: u64::try_from(WireLimits::MAX_NESTING).unwrap_or(u64::MAX),
                path: Path::Package,
            });
        }
        Ok(())
    }

    fn charge_fields(&mut self, amount: usize) -> Result<(), Error> {
        charge_remaining(
            &mut self.remaining_fields,
            self.maximum_fields,
            amount,
            LimitKind::WireFields,
        )
    }

    fn charge_work(&mut self, amount: usize) -> Result<(), Error> {
        charge_remaining(
            &mut self.remaining_work,
            self.maximum_work,
            amount,
            LimitKind::WireWork,
        )
    }

    fn charge_references(&mut self, amount: usize) -> Result<(), Error> {
        charge_remaining(
            &mut self.remaining_references,
            self.maximum_references,
            amount,
            LimitKind::PayloadReferences,
        )
    }

    fn charge_transaction_work(&mut self, amount: usize) -> Result<(), Error> {
        charge_remaining(
            &mut self.remaining_transaction_work,
            self.maximum_transaction_work,
            amount,
            LimitKind::TransactionWork,
        )
    }

    fn charge_index_lookup(&mut self, source: &Package) -> Result<(), Error> {
        self.charge_transaction_work(source.state.index.lookup_work())
    }

    fn settle_ownership(
        &mut self,
        reserved_work: usize,
        reserved_references: usize,
        reserved_transaction_work: usize,
        actual: super::table_headers::ownership::OwnershipReport,
    ) -> Result<(), Error> {
        if actual.work > reserved_work || actual.references > reserved_references {
            return Err(invalid_source());
        }
        self.remaining_work = self
            .remaining_work
            .checked_add(reserved_work - actual.work)
            .ok_or_else(invalid_source)?;
        self.remaining_references = self
            .remaining_references
            .checked_add(reserved_references - actual.references)
            .ok_or_else(invalid_source)?;
        if actual.transaction_work > reserved_transaction_work {
            return Err(invalid_source());
        }
        self.remaining_transaction_work = self
            .remaining_transaction_work
            .checked_add(reserved_transaction_work - actual.transaction_work)
            .ok_or_else(invalid_source)?;
        Ok(())
    }

    #[cfg(test)]
    fn test_with_limits(
        max_message_bytes: usize,
        maximum_fields: usize,
        maximum_work: usize,
        maximum_references: usize,
        maximum_transaction_work: usize,
    ) -> Self {
        Self {
            max_message_bytes,
            recursion_limit: u32::try_from(WireLimits::MAX_NESTING).unwrap_or(u32::MAX),
            maximum_fields,
            maximum_work,
            maximum_references,
            maximum_transaction_work,
            remaining_fields: maximum_fields,
            remaining_work: maximum_work,
            remaining_references: maximum_references,
            remaining_transaction_work: maximum_transaction_work,
        }
    }

    #[cfg(test)]
    const fn test_usage(&self) -> (usize, usize, usize, usize) {
        (
            self.maximum_fields - self.remaining_fields,
            self.maximum_work - self.remaining_work,
            self.maximum_references - self.remaining_references,
            self.maximum_transaction_work - self.remaining_transaction_work,
        )
    }
}

fn charge_remaining(
    remaining: &mut usize,
    maximum: usize,
    amount: usize,
    kind: LimitKind,
) -> Result<(), Error> {
    if amount > *remaining {
        let observed = maximum
            .checked_sub(*remaining)
            .and_then(|used| used.checked_add(amount))
            .unwrap_or(usize::MAX);
        return Err(Error::LimitExceeded {
            kind,
            observed: u64::try_from(observed).unwrap_or(u64::MAX),
            maximum: u64::try_from(maximum).unwrap_or(u64::MAX),
            path: Path::Package,
        });
    }
    *remaining -= amount;
    Ok(())
}

fn decode_title_with_report(
    source: &[u8],
) -> Result<(DecodedTitle, numbers_table_title_codec::DecodeReport), Error> {
    decode_title_with_options(
        source,
        numbers_table_title_codec::DecodeOptions::new(
            source.len().max(1),
            source.len().max(1),
            source.len().saturating_mul(4).max(1),
            u32::try_from(WireLimits::MAX_NESTING).unwrap_or(u32::MAX),
            2,
        ),
    )
}

fn decode_title_with_budget(
    source: &[u8],
    budget: &mut TransactionBudget,
) -> Result<DecodedTitle, Error> {
    let (decoded, report) = decode_title_with_options(source, budget.codec_options())?;
    budget.consume(report)?;
    Ok(decoded)
}

fn decode_title_with_options(
    source: &[u8],
    options: numbers_table_title_codec::DecodeOptions,
) -> Result<(DecodedTitle, numbers_table_title_codec::DecodeReport), Error> {
    let (snapshot, report) =
        numbers_table_title_codec::decode_table_title_settings_with_report(source, options)
            .map_err(map_title_codec_error)?;
    let paragraph_style = snapshot.table_name_style();
    let shape_style = snapshot.table_name_shape_style();
    if paragraph_style.is_some_and(|reference| reference.deprecated_is_external() == Some(true))
        || shape_style.is_some_and(|reference| reference.deprecated_is_external() == Some(true))
    {
        return Err(invalid_source());
    }
    Ok((
        DecodedTitle {
            settings: Settings::new(
                snapshot.table_name_enabled(),
                snapshot.table_name_border_enabled(),
            ),
            height_bits: snapshot.table_name_height_bits(),
            paragraph_style: paragraph_style.map(|reference| reference.identifier()),
            shape_style: shape_style.map(|reference| reference.identifier()),
        },
        report,
    ))
}

const fn invalid_source() -> Error {
    Error::InvalidSource {
        path: Path::Package,
    }
}

fn map_header_error(error: super::table_headers::Error) -> Error {
    match error {
        super::table_headers::Error::SheetNotFound => Error::SheetNotFound,
        super::table_headers::Error::TableNotFound => Error::TableNotFound,
        super::table_headers::Error::UnsupportedSource => Error::UnsupportedSource,
        super::table_headers::Error::Allocation { amount, path } => Error::Allocation {
            amount,
            path: map_header_path(path),
        },
        super::table_headers::Error::LimitExceeded {
            kind,
            observed,
            maximum,
            path,
        } => Error::LimitExceeded {
            kind: map_header_limit(kind),
            observed,
            maximum,
            path: map_header_path(path),
        },
        super::table_headers::Error::PatchConflict => Error::PatchConflict,
        super::table_headers::Error::Verification => Error::Verification,
        super::table_headers::Error::TableLocked { path } => Error::TableLocked {
            path: map_header_path(path),
        },
        _ => invalid_source(),
    }
}

const fn map_header_path(path: super::table_headers::Path) -> Path {
    match path {
        super::table_headers::Path::Package => Path::Package,
        super::table_headers::Path::Table { sheet, table } => Path::Table { sheet, table },
    }
}

const fn map_header_limit(kind: super::table_headers::LimitKind) -> LimitKind {
    use super::table_headers::LimitKind as Header;
    match kind {
        Header::InputBytes => LimitKind::InputBytes,
        Header::OutputBytes => LimitKind::OutputBytes,
        Header::Entries => LimitKind::Entries,
        Header::EntryBytes => LimitKind::EntryBytes,
        Header::TotalEntryBytes => LimitKind::TotalEntryBytes,
        Header::PackageBytes => LimitKind::PackageBytes,
        Header::PayloadBytes => LimitKind::PayloadBytes,
        Header::TotalPayloadBytes => LimitKind::TotalPayloadBytes,
        Header::PayloadObjects => LimitKind::PayloadObjects,
        Header::PayloadMessages => LimitKind::PayloadMessages,
        Header::PayloadItems => LimitKind::PayloadItems,
        Header::PayloadReferences => LimitKind::PayloadReferences,
        Header::WireBytes => LimitKind::WireBytes,
        Header::WireOutputBytes => LimitKind::WireOutputBytes,
        Header::WireFields => LimitKind::WireFields,
        Header::WireNesting => LimitKind::WireNesting,
        Header::WireWork => LimitKind::WireWork,
        Header::TransactionWork => LimitKind::TransactionWork,
    }
}

fn map_wire_error(error: litchi_iwa_common::Error) -> Error {
    match error {
        litchi_iwa_common::Error::LimitExceeded {
            kind,
            observed,
            limit,
        } => Error::LimitExceeded {
            kind: match kind {
                litchi_iwa_common::LimitKind::InputBytes => LimitKind::WireBytes,
                litchi_iwa_common::LimitKind::OutputBytes => LimitKind::WireOutputBytes,
                litchi_iwa_common::LimitKind::Fields => LimitKind::WireFields,
                litchi_iwa_common::LimitKind::Nesting => LimitKind::WireNesting,
                litchi_iwa_common::LimitKind::RewriteWork => LimitKind::WireWork,
                _ => LimitKind::PayloadItems,
            },
            observed: u64::try_from(observed).unwrap_or(u64::MAX),
            maximum: u64::try_from(limit).unwrap_or(u64::MAX),
            path: Path::Package,
        },
        litchi_iwa_common::Error::Allocation { amount, .. } => Error::Allocation {
            amount,
            path: Path::Package,
        },
        _ => invalid_source(),
    }
}

fn map_title_codec_error(error: numbers_table_title_codec::DecodeError) -> Error {
    use numbers_table_title_codec::DecodeLimit;

    match error.resource_limit() {
        Some(DecodeLimit::Bytes { observed, maximum }) => Error::LimitExceeded {
            kind: LimitKind::WireBytes,
            observed: u64::try_from(observed).unwrap_or(u64::MAX),
            maximum: u64::try_from(maximum).unwrap_or(u64::MAX),
            path: Path::Package,
        },
        Some(DecodeLimit::References { observed, maximum }) => Error::LimitExceeded {
            kind: LimitKind::PayloadReferences,
            observed: u64::try_from(observed).unwrap_or(u64::MAX),
            maximum: u64::try_from(maximum).unwrap_or(u64::MAX),
            path: Path::Package,
        },
        Some(DecodeLimit::Fields { observed, maximum }) => Error::LimitExceeded {
            kind: LimitKind::WireFields,
            observed: u64::try_from(observed).unwrap_or(u64::MAX),
            maximum: u64::try_from(maximum).unwrap_or(u64::MAX),
            path: Path::Package,
        },
        Some(DecodeLimit::Work { observed, maximum }) => Error::LimitExceeded {
            kind: LimitKind::WireWork,
            observed: u64::try_from(observed).unwrap_or(u64::MAX),
            maximum: u64::try_from(maximum).unwrap_or(u64::MAX),
            path: Path::Package,
        },
        Some(DecodeLimit::Nesting { observed, maximum }) => Error::LimitExceeded {
            kind: LimitKind::WireNesting,
            observed: u64::from(observed),
            maximum: u64::from(maximum),
            path: Path::Package,
        },
        Some(_) | None => invalid_source(),
    }
}

fn validate_changed_source(
    source: &Package,
    target: Target,
    budget: &mut TransactionBudget,
) -> Result<(), Error> {
    let path = Path::Table {
        sheet: target.native.sheet_position,
        table: target.native.table_position,
    };
    let reserved_work = source.state.options.archive().max_iwa_stream_bytes();
    let reserved_references = source.state.options.semantic().max_references();
    let reserved_transaction_work = reserved_work.saturating_mul(8).saturating_add(
        source
            .state
            .index
            .object_count()
            .saturating_mul(2)
            .saturating_add(1)
            .saturating_mul(source.state.index.lookup_work()),
    );
    budget.charge_work(reserved_work)?;
    budget.charge_references(reserved_references)?;
    budget.charge_transaction_work(reserved_transaction_work)?;
    let report = super::table_headers::ownership::validate_selected_ownership_with_report(
        source,
        target.native,
    )
    .map_err(map_header_error)?;
    budget.settle_ownership(
        reserved_work,
        reserved_references,
        reserved_transaction_work,
        report,
    )?;
    if target.native.locked == LockState::Locked {
        return Err(Error::TableLocked { path });
    }
    validate_view_state_dependency(source, budget)?;
    Ok(())
}

fn validate_view_state_dependency(
    source: &Package,
    budget: &mut TransactionBudget,
) -> Result<(), Error> {
    const TABLE_NAME_SELECTION_MESSAGE_TYPE: u32 = 6_284;

    let Some(archive) = source.state.components.get_archive("Index/ViewState.iwa") else {
        return Ok(());
    };
    let mut work = 0usize;
    for object in &archive.objects {
        work = work
            .checked_add(1)
            .and_then(|value| value.checked_add(object.messages.len()))
            .ok_or_else(invalid_source)?;
        if object
            .messages
            .iter()
            .any(|message| message.type_ == TABLE_NAME_SELECTION_MESSAGE_TYPE)
        {
            budget.charge_transaction_work(work)?;
            return Err(Error::UnsupportedSource);
        }
    }
    budget.charge_transaction_work(work)
}

fn validate_visible_prerequisites(
    source: &Package,
    target: Target,
    budget: &mut TransactionBudget,
) -> Result<(), Error> {
    if !target.settings.is_visible() {
        return Ok(());
    }
    let path = Path::Table {
        sheet: target.native.sheet_position,
        table: target.native.table_position,
    };
    let height_bits = target
        .height_bits
        .ok_or(Error::UnsupportedDependency { path })?;
    let height = f64::from_bits(height_bits);
    if !height.is_finite() || height < 0.0 {
        return Err(invalid_source());
    }
    let paragraph_style = target
        .paragraph_style
        .ok_or(Error::UnsupportedDependency { path })?;
    let shape_style = target
        .shape_style
        .ok_or(Error::UnsupportedDependency { path })?;
    if paragraph_style == shape_style
        || [paragraph_style, shape_style]
            .into_iter()
            .any(|identifier| {
                identifier == 1
                    || identifier == target.native.sheet_identifier
                    || identifier == target.native.drawable_identifier
                    || identifier == target.native.model_identifier
            })
    {
        return Err(invalid_source());
    }
    let model_object = target_object(source, target.native)?;
    charge_message_metadata(model_object, target.native.message_index, budget)?;
    let declaration_work = declared_reference_scan_work(model_object, target.native.message_index)?;
    budget.charge_transaction_work(declaration_work.saturating_mul(3))?;
    super::table_headers::resolve::require_declared_reference(
        model_object,
        target.native.message_index,
        paragraph_style,
        &[30],
    )
    .map_err(map_header_error)?;
    super::table_headers::resolve::require_declared_reference(
        model_object,
        target.native.message_index,
        shape_style,
        &[36],
    )
    .map_err(map_header_error)?;
    require_style(
        source,
        paragraph_style,
        PARAGRAPH_STYLE_MESSAGE_TYPE,
        budget,
    )?;
    require_style(source, shape_style, SHAPE_STYLE_MESSAGE_TYPE, budget)
}

fn target_object(
    source: &Package,
    target: super::table_headers::Target,
) -> Result<&litchi_iwa_core::ArchiveObject, Error> {
    source
        .state
        .components
        .catalog()
        .get_index(target.component_index)
        .and_then(|component| component.archive().objects.get(target.object_index))
        .filter(|object| object.archive_info.identifier == Some(target.model_identifier))
        .ok_or_else(invalid_source)
}

fn require_style(
    source: &Package,
    identifier: u64,
    message_type: u32,
    budget: &mut TransactionBudget,
) -> Result<(), Error> {
    budget.charge_index_lookup(source)?;
    let resolved = source
        .state
        .index
        .resolve_ref_id(&source.state.components, identifier)
        .map_err(|_error| invalid_source())?
        .ok_or_else(invalid_source)?;
    let object = super::table_headers::resolve::resolved_object(source, resolved)
        .map_err(map_header_error)?;
    if object.archive_info.identifier != Some(identifier) {
        return Err(invalid_source());
    }
    let (message_index, message) =
        super::table_headers::resolve::unique_message_index(&object.messages, message_type)
            .map_err(map_header_error)?
            .ok_or_else(invalid_source)?;
    charge_message_metadata(object, message_index, budget)?;
    super::table_headers::resolve::validate_message_metadata(object, message_index)
        .map_err(map_header_error)?;
    validate_required_super(&message.data, budget)
}

fn validate_required_super(source: &[u8], budget: &mut TransactionBudget) -> Result<(), Error> {
    budget.charge_fields(source.len())?;
    budget.charge_work(source.len().saturating_mul(2))?;
    let limits = WireLimits::default()
        .with_input_bytes(source.len().max(1))
        .map_err(map_wire_error)?
        .with_fields(source.len().max(1))
        .map_err(map_wire_error)?;
    let outer = WireView::parse_with_limits(source, limits).map_err(map_wire_error)?;
    let mut super_payload = None;
    for field in outer.fields() {
        field.validate_canonical_framing().map_err(map_wire_error)?;
        if field.number() == 1 {
            if field.wire_type() != 2 || super_payload.replace(field.payload()).is_some() {
                return Err(invalid_source());
            }
        }
    }
    let nested_source = super_payload.ok_or_else(invalid_source)?;
    budget.charge_fields(nested_source.len())?;
    budget.charge_work(nested_source.len().saturating_mul(2))?;
    let nested = WireView::parse_with_limits(
        nested_source,
        WireLimits::default()
            .with_input_bytes(source.len().max(1))
            .map_err(map_wire_error)?
            .with_fields(source.len().max(1))
            .map_err(map_wire_error)?,
    )
    .map_err(map_wire_error)?;
    for field in nested.fields() {
        field.validate_canonical_framing().map_err(map_wire_error)?;
    }
    Ok(())
}

fn charge_message_metadata(
    object: &litchi_iwa_core::ArchiveObject,
    message_index: usize,
    budget: &mut TransactionBudget,
) -> Result<(), Error> {
    let info = object
        .archive_info
        .message_infos
        .get(message_index)
        .ok_or_else(invalid_source)?;
    let mut references = info
        .object_references
        .len()
        .checked_add(info.data_references.len())
        .ok_or_else(invalid_source)?;
    let mut work = 1usize
        .checked_add(info.versions.len())
        .and_then(|value| value.checked_add(info.diff_merge_version.len()))
        .and_then(|value| value.checked_add(info.diff_read_version.len()))
        .and_then(|value| value.checked_add(info.fields_to_remove.len()))
        .ok_or_else(invalid_source)?;
    if let Some(path) = &info.diff_field_path {
        work = work
            .checked_add(path.path.len())
            .ok_or_else(invalid_source)?;
    }
    for field in &info.field_infos {
        references = references
            .checked_add(field.object_references.len())
            .and_then(|value| value.checked_add(field.data_references.len()))
            .ok_or_else(invalid_source)?;
        work = work
            .checked_add(1)
            .and_then(|value| value.checked_add(field.path.path.len()))
            .and_then(|value| value.checked_add(field.known_field_version.len()))
            .and_then(|value| {
                value.checked_add(
                    field
                        .known_field_feature_identifier
                        .as_ref()
                        .map_or(0, String::len),
                )
            })
            .ok_or_else(invalid_source)?;
    }
    budget.charge_references(references)?;
    budget.charge_transaction_work(work.saturating_add(references))
}

fn declared_reference_scan_work(
    object: &litchi_iwa_core::ArchiveObject,
    message_index: usize,
) -> Result<usize, Error> {
    let info = object
        .archive_info
        .message_infos
        .get(message_index)
        .ok_or_else(invalid_source)?;
    info.field_infos
        .iter()
        .try_fold(info.object_references.len(), |work, field| {
            work.checked_add(1)
                .and_then(|value| value.checked_add(field.object_references.len()))
                .ok_or_else(invalid_source)
        })
}

fn reopen_cost(source: &Package) -> Result<ReopenCost, Error> {
    let catalog =
        super::table_headers::rewrite::physical_source(source).map_err(map_header_error)?;
    let mut work = source
        .source_bytes()
        .len()
        .checked_add(source.state.index.rebuild_work())
        .ok_or_else(invalid_source)?;
    for entry in catalog.package().iter() {
        work = work
            .checked_add(entry.data().len().saturating_mul(2))
            .and_then(|value| value.checked_add(size_of_val(entry)))
            .and_then(|value| value.checked_add(entry.raw_name().len()))
            .ok_or_else(invalid_source)?;
    }
    for component in catalog.components().iter() {
        work = work
            .checked_add(size_of_val(component))
            .and_then(|value| value.checked_add(component.name().len()))
            .ok_or_else(invalid_source)?;
        let extent = component
            .archive()
            .objects
            .iter()
            .try_fold(0usize, |maximum, object| {
                let end = object
                    .header_offset
                    .checked_add(object.header_length)
                    .and_then(|offset| offset.checked_add(object.data_length))
                    .and_then(|value| usize::try_from(value).ok())
                    .ok_or_else(invalid_source)?;
                Ok::<usize, Error>(maximum.max(end))
            })?;
        work = work
            .checked_add(extent.saturating_mul(2))
            .ok_or_else(invalid_source)?;
    }
    let mut references = 0usize;
    for object in source.state.components.iter_objects() {
        work = work
            .checked_add(size_of::<litchi_iwa_core::ArchiveObject>())
            .and_then(|value| {
                value.checked_add(
                    object
                        .messages
                        .len()
                        .saturating_mul(size_of::<litchi_iwa_core::RawMessage>()),
                )
            })
            .ok_or_else(invalid_source)?;
        for (message_index, message) in object.messages.iter().enumerate() {
            work = work
                .checked_add(message.data.len().saturating_mul(4))
                .ok_or_else(invalid_source)?;
            let info = object
                .archive_info
                .message_infos
                .get(message_index)
                .ok_or_else(invalid_source)?;
            references = references
                .checked_add(info.object_references.len())
                .and_then(|value| value.checked_add(info.data_references.len()))
                .ok_or_else(invalid_source)?;
            work = work
                .checked_add(size_of::<litchi_iwa_core::MessageInfo>())
                .and_then(|value| {
                    value.checked_add(info.versions.len().saturating_mul(size_of::<u32>()))
                })
                .and_then(|value| {
                    value.checked_add(
                        info.diff_merge_version
                            .len()
                            .saturating_mul(size_of::<u32>()),
                    )
                })
                .and_then(|value| {
                    value.checked_add(
                        info.diff_read_version
                            .len()
                            .saturating_mul(size_of::<u32>()),
                    )
                })
                .and_then(|value| {
                    value.checked_add(
                        info.fields_to_remove
                            .len()
                            .saturating_mul(size_of::<litchi_iwa_core::FieldPath>()),
                    )
                })
                .and_then(|value| {
                    value.checked_add(
                        info.object_references
                            .len()
                            .saturating_mul(size_of::<u64>()),
                    )
                })
                .and_then(|value| {
                    value.checked_add(info.data_references.len().saturating_mul(size_of::<u64>()))
                })
                .ok_or_else(invalid_source)?;
            if let Some(path) = &info.diff_field_path {
                work = work
                    .checked_add(size_of::<litchi_iwa_core::FieldPath>())
                    .and_then(|value| {
                        value.checked_add(path.path.len().saturating_mul(size_of::<u32>()))
                    })
                    .ok_or_else(invalid_source)?;
            }
            for path in &info.fields_to_remove {
                work = work
                    .checked_add(path.path.len().saturating_mul(size_of::<u32>()))
                    .ok_or_else(invalid_source)?;
            }
            for field in &info.field_infos {
                references = references
                    .checked_add(field.object_references.len())
                    .and_then(|value| value.checked_add(field.data_references.len()))
                    .ok_or_else(invalid_source)?;
                work = work
                    .checked_add(size_of::<litchi_iwa_core::FieldInfo>())
                    .and_then(|value| {
                        value.checked_add(field.path.path.len().saturating_mul(size_of::<u32>()))
                    })
                    .and_then(|value| {
                        value.checked_add(
                            field
                                .object_references
                                .len()
                                .saturating_mul(size_of::<u64>()),
                        )
                    })
                    .and_then(|value| {
                        value.checked_add(
                            field.data_references.len().saturating_mul(size_of::<u64>()),
                        )
                    })
                    .and_then(|value| {
                        value.checked_add(
                            field
                                .known_field_version
                                .len()
                                .saturating_mul(size_of::<u32>()),
                        )
                    })
                    .and_then(|value| {
                        value.checked_add(
                            field
                                .known_field_feature_identifier
                                .as_ref()
                                .map_or(0, String::len),
                        )
                    })
                    .ok_or_else(invalid_source)?;
            }
        }
    }
    work = work.checked_add(references).ok_or_else(invalid_source)?;
    Ok(ReopenCost { work, references })
}

fn commit_edit(edit: Edit<'_>) -> Result<Commit, Error> {
    let catalog =
        super::table_headers::rewrite::physical_source(edit.source).map_err(map_header_error)?;
    let source = catalog.__source_owner();
    if edit.before == edit.settings {
        return Ok(Commit {
            package: edit.source.snapshot(),
            patch: Patch {
                artifacts: OwnedExactArtifacts::new(source.clone(), source),
                target: edit.target,
                before: edit.before,
                after: edit.settings,
                source_reopen: ReopenCost {
                    work: 0,
                    references: 0,
                },
                target_reopen: ReopenCost {
                    work: 0,
                    references: 0,
                },
                source_previews: 0,
                target_previews: 0,
            },
            diagnostics: Diagnostics::unchanged(),
        });
    }
    if !catalog.source_is_exact() {
        return Err(Error::UnsupportedSource);
    }
    let mut budget = edit.budget;
    budget.charge_transaction_work(source.len().saturating_mul(2))?;
    if edit.target.settings != edit.before {
        return Err(invalid_source());
    }
    validate_changed_source(edit.source, edit.target, &mut budget)?;
    let requested_target = Target {
        settings: edit.settings,
        ..edit.target
    };
    validate_visible_prerequisites(edit.source, requested_target, &mut budget)?;
    let previews =
        super::table_headers::rewrite::root_preview_deletions(catalog).map_err(map_header_error)?;
    let source_reopen = reopen_cost(edit.source)?;
    let package = rewrite::rewrite(
        edit.source,
        edit.target,
        edit.settings,
        &previews,
        &mut budget,
        source_reopen,
    )?;
    let target_payload =
        super::table_headers::rewrite::selected_payload(&package, edit.target.native)
            .map_err(map_header_error)?;
    let target_bytes = super::table_headers::rewrite::physical_source(&package)
        .map_err(map_header_error)?
        .__source_owner();
    let target_reopen = reopen_cost(&package)?;
    budget.charge_transaction_work(source.len().saturating_add(target_bytes.len()))?;
    budget.charge_references(
        source_reopen
            .references
            .saturating_add(target_reopen.references),
    )?;
    budget.charge_transaction_work(source_reopen.work.saturating_add(target_reopen.work))?;
    super::table_headers::rewrite::verify_exact_locality(
        edit.source,
        &package,
        edit.target.native,
        &previews,
        0,
        target_payload,
    )
    .map_err(map_header_error)?;
    Ok(Commit {
        package,
        patch: Patch {
            artifacts: OwnedExactArtifacts::new(source, target_bytes.clone()),
            target: edit.target,
            before: edit.before,
            after: edit.settings,
            source_reopen,
            target_reopen,
            source_previews: previews.len(),
            target_previews: 0,
        },
        diagnostics: Diagnostics::published(previews.len()),
    })
}

fn apply_patch(source: &Package, patch: &Patch) -> Result<Commit, Error> {
    let source_catalog =
        super::table_headers::rewrite::physical_source(source).map_err(map_header_error)?;
    let bytes = source_catalog.__source_owner();
    let mut budget = TransactionBudget::new(source);
    budget.charge_transaction_work(bytes.len())?;
    if !patch.artifacts.authorizes_owner(&bytes) {
        return Err(Error::PatchConflict);
    }
    if patch.is_noop() {
        return Ok(Commit {
            package: source.snapshot(),
            patch: patch.clone(),
            diagnostics: Diagnostics::unchanged(),
        });
    }
    let current_payload =
        super::table_headers::rewrite::selected_payload(source, patch.target.native)
            .map_err(map_header_error)?;
    let current = decode_title_with_budget(current_payload, &mut budget)?;
    if current.settings != patch.before {
        return Err(Error::PatchConflict);
    }
    if !source_catalog.source_is_exact() {
        return Err(Error::PatchConflict);
    }
    let target_bytes = patch.artifacts.target_owner();
    let target_bytes_len = target_bytes.len();
    budget.charge_transaction_work(
        bytes
            .len()
            .saturating_add(target_bytes_len.saturating_mul(2)),
    )?;
    budget.charge_references(patch.target_reopen.references)?;
    budget.charge_transaction_work(patch.target_reopen.work)?;
    let candidate = Package::from_source_owner_with_options(target_bytes, source.state.options)
        .map_err(|_error| Error::Verification)?;
    let selected = resolve_title_at_positions_with_budget(
        &candidate,
        patch.target.native.sheet_position,
        patch.target.native.table_position,
        &mut budget,
    )?;
    if selected.settings != patch.after {
        return Err(Error::Verification);
    }
    let source_previews = super::table_headers::rewrite::root_preview_deletions(source_catalog)
        .map_err(map_header_error)?;
    if source_previews.len() != patch.source_previews {
        return Err(Error::PatchConflict);
    }
    let candidate_payload =
        super::table_headers::rewrite::selected_payload(&candidate, patch.target.native)
            .map_err(map_header_error)?;
    budget.charge_transaction_work(bytes.len().saturating_add(target_bytes_len))?;
    budget.charge_references(
        patch
            .source_reopen
            .references
            .saturating_add(patch.target_reopen.references),
    )?;
    budget.charge_transaction_work(
        patch
            .source_reopen
            .work
            .saturating_add(patch.target_reopen.work),
    )?;
    super::table_headers::rewrite::verify_exact_locality(
        source,
        &candidate,
        patch.target.native,
        &source_previews,
        patch.target_previews,
        candidate_payload,
    )
    .map_err(map_header_error)?;
    Ok(Commit {
        package: candidate,
        patch: patch.clone(),
        diagnostics: Diagnostics::published(
            patch.source_previews.saturating_sub(patch.target_previews),
        ),
    })
}

#[cfg(test)]
mod tests {
    use litchi_iwa_archive::{Limits, package::Catalog};
    use litchi_iwa_common::wire::append_length_delimited_field;
    use litchi_iwa_core::{Archive, ArchiveObject, RawMessage, SnappyStream};
    use litchi_iwa_protos::{tn, tsd, tsp, tst};
    use prost::Message as _;

    use super::{
        Error, Package, TransactionBudget, decode_title_with_budget, reopen_cost,
        resolve_title_at_positions_with_budget, validate_changed_source,
    };

    type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

    const DOCUMENT: u64 = 1;
    const SHEET: u64 = 2;
    const FIRST_INFO: u64 = 10_000;
    const FIRST_MODEL: u64 = 100_000;
    const SIDECAR: u64 = 900_000;

    fn reference(identifier: u64) -> tsp::Reference {
        tsp::Reference {
            identifier,
            ..Default::default()
        }
    }

    fn object(identifier: u64, message_type: u32, data: Vec<u8>) -> TestResult<ArchiveObject> {
        Ok(ArchiveObject::new(
            identifier,
            vec![RawMessage {
                type_: message_type,
                data,
            }],
        )?)
    }

    fn topology_package(table_count: usize) -> TestResult<Vec<u8>> {
        let mut document = object(
            DOCUMENT,
            1,
            tn::DocumentArchive {
                sheets: vec![reference(SHEET)],
                ..Default::default()
            }
            .encode_to_vec(),
        )?;
        document.archive_info.message_infos[0].object_references = vec![SHEET];
        let drawable_identifiers = (0..table_count)
            .map(|index| FIRST_INFO.saturating_add(u64::try_from(index).unwrap_or(u64::MAX)))
            .collect::<Vec<_>>();
        let mut sheet = object(
            SHEET,
            2,
            tn::SheetArchive {
                name: "Scale".to_owned(),
                drawable_infos: drawable_identifiers
                    .iter()
                    .copied()
                    .map(reference)
                    .collect(),
                ..Default::default()
            }
            .encode_to_vec(),
        )?;
        sheet.archive_info.message_infos[0].object_references = drawable_identifiers.clone();
        let document_component = SnappyStream::compress(
            &Archive {
                objects: vec![document, sheet],
            }
            .to_bytes()?,
        )?;

        let mut table_objects = Vec::with_capacity(table_count.saturating_mul(2).saturating_add(1));
        for (position, drawable_identifier) in drawable_identifiers.iter().copied().enumerate() {
            let model_identifier =
                FIRST_MODEL.saturating_add(u64::try_from(position).unwrap_or(u64::MAX));
            let mut info_payload = Vec::new();
            append_length_delimited_field(
                &mut info_payload,
                1,
                &tsd::DrawableArchive::default().encode_to_vec(),
            )?;
            append_length_delimited_field(
                &mut info_payload,
                2,
                &reference(model_identifier).encode_to_vec(),
            )?;
            let mut info = object(drawable_identifier, 6_000, info_payload)?;
            info.archive_info.message_infos[0].object_references = vec![model_identifier];
            table_objects.push(info);
            table_objects.push(object(
                model_identifier,
                6_001,
                tst::TableModelArchive {
                    table_name: format!("Table {position}"),
                    number_of_rows: 1,
                    number_of_columns: 1,
                    base_data_store: tst::DataStore {
                        string_table: reference(SIDECAR),
                        formula_table: reference(SIDECAR),
                        ..Default::default()
                    },
                    ..Default::default()
                }
                .encode_to_vec(),
            )?);
        }
        table_objects.push(ArchiveObject::new(
            SIDECAR,
            [
                tst::table_data_list::ListType::String,
                tst::table_data_list::ListType::Formula,
            ]
            .into_iter()
            .map(|list_type| RawMessage {
                type_: 6_005,
                data: tst::TableDataList {
                    list_type: list_type as i32,
                    next_list_id: 1,
                    ..Default::default()
                }
                .encode_to_vec(),
            })
            .collect(),
        )?);
        let table_component = SnappyStream::compress(
            &Archive {
                objects: table_objects,
            }
            .to_bytes()?,
        )?;
        Ok(litchi_iwa_archive::package::to_bytes(
            [
                ("preview.jpg", b"full".as_slice()),
                ("preview-micro.jpg", b"micro".as_slice()),
                ("preview-web.jpg", b"web".as_slice()),
                ("Index/Document.iwa", document_component.as_slice()),
                ("Index/Tables.iwa", table_component.as_slice()),
            ],
            Limits::default(),
        )?)
    }

    fn topology_usage(table_count: usize) -> TestResult<(usize, usize, usize, usize)> {
        let bytes = topology_package(table_count)?;
        let package = Package::from_bytes(&bytes)?;
        let mut budget = TransactionBudget::new(&package);
        let target = resolve_title_at_positions_with_budget(
            &package,
            0,
            table_count.saturating_sub(1),
            &mut budget,
        )?;
        validate_changed_source(&package, target, &mut budget)?;
        let payload =
            super::super::table_headers::rewrite::selected_payload(&package, target.native)?;
        let _decoded = decode_title_with_budget(payload, &mut budget)?;
        let reopen = reopen_cost(&package)?;
        budget.charge_references(reopen.references)?;
        budget.charge_transaction_work(reopen.work)?;
        let catalog = Catalog::from_bytes(&bytes)?;
        let selected_compressed_len = catalog
            .iter()
            .find(|entry| entry.name() == "Index/Tables.iwa")
            .ok_or(Error::Verification)?
            .data()
            .len();
        let (_bound, reassembly_work) =
            super::rewrite::reassembly_cost(bytes.len(), selected_compressed_len)?;
        budget.charge_transaction_work(reassembly_work)?;
        Ok(budget.test_usage())
    }

    #[test]
    fn production_rooted_topology_counters_scale_linearly() -> TestResult {
        let small = topology_usage(4_096)?;
        let large = topology_usage(8_192)?;
        for (small_counter, large_counter) in [
            (small.0, large.0),
            (small.1, large.1),
            (small.2, large.2),
            (small.3, large.3),
        ] {
            assert!(small_counter != 0);
            assert!(large_counter.saturating_mul(10) <= small_counter.saturating_mul(23));
        }
        assert_eq!(small.2, 2 + 4 * 4_096);
        assert_eq!(large.2, 2 + 4 * 8_192);

        let mut attempted_budget = TransactionBudget::test_with_limits(1, 1, 1, 1, large.3 - 1);
        let attempted = attempted_budget.charge_transaction_work(large.3);
        assert!(matches!(
            attempted,
            Err(Error::LimitExceeded {
                kind: super::LimitKind::TransactionWork,
                ..
            })
        ));
        assert_eq!(attempted_budget.test_usage(), (0, 0, 0, 0));
        Ok(())
    }
}
