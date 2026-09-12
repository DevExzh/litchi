//! Private, strict slide-table graph admission shared by focused owners.
//!
//! This module deliberately owns no public values and no mutation policy.  It
//! resolves the canonical slide -> `TableInfo` -> table-model route, proves
//! the physical metadata authority for that route, and provides the small
//! helpers needed by package transactions.  The table-name owner is the first
//! consumer; keeping this proof here avoids copying the same several-hundred
//! line graph walk into every table property owner.

#![allow(
    clippy::map_err_ignore,
    clippy::needless_pass_by_value,
    clippy::too_many_arguments,
    clippy::wildcard_enum_match_arm,
    reason = "The private graph boundary redacts native failure detail."
)]

use std::collections::{HashMap, HashSet};
use std::mem::size_of;
use std::sync::Arc;

use litchi_core::Position;
use litchi_iwa_archive::SourceCatalog;
use litchi_iwa_common::{
    WireLimits, decode_varint_from_bytes, varint::encoded_len, wire::WireView,
};
use litchi_iwa_core::{
    Archive, ArchiveObject, ArchiveReferenceKind, ArchiveReferenceOccurrence,
    ArchiveReferencePolicy, ArchiveReferenceScope, ArchiveReferenceVisitor, FieldType,
    SnappyStream,
};
use litchi_iwa_protos::{
    numbers_table_cell_dependency_codec as dependency_codec,
    numbers_table_cell_storage_codec as storage_codec, package_metadata_codec,
    table_dimension_codec, table_info_codec, table_model_discovery_codec, table_sort_order_codec,
};

use super::{
    Package, PayloadLimitKind, PhysicalSource, ReadError, SemanticBudget, SemanticLimitKind,
};
use crate::{SlideSelector, slide::table::TableSelector};

const SLIDE_MESSAGE_TYPE: u32 = 5;
const TABLE_INFO_MESSAGE_TYPE: u32 = 6_000;
const TABLE_MODEL_MESSAGE_TYPE: u32 = 6_001;
const HEADER_BUCKET_MESSAGE_TYPE: u32 = 6_006;
const TILE_MESSAGE_TYPE: u32 = 6_002;
const TABLE_DATA_LIST_MESSAGE_TYPE: u32 = 6_005;
const TABLE_DATA_LIST_NATIVE_MESSAGE_TYPE: u32 = 6_201;
const COLUMN_ROW_UID_MAP_LEGACY_MESSAGE_TYPE: u32 = 6_200;
const COLUMN_ROW_UID_MAP_MESSAGE_TYPE: u32 = 6_267;
const STROKE_SIDECAR_MESSAGE_TYPE: u32 = 6_305;
const HIDDEN_STATE_FORMULA_OWNER_MESSAGE_TYPE: u32 = 6_204;
const FILTER_SET_MESSAGE_TYPE: u32 = 6_220;
const CATEGORY_OWNER_REFERENCE_MESSAGE_TYPE: u32 = 6_372;
const GROUP_BY_MESSAGE_TYPE: u32 = 6_373;
const GROUP_NODE_MESSAGE_TYPE: u32 = 6_383;
const PHYSICAL_MUTABLE_MESSAGE_TYPES: [u32; 4] = [
    TILE_MESSAGE_TYPE,
    HEADER_BUCKET_MESSAGE_TYPE,
    COLUMN_ROW_UID_MAP_LEGACY_MESSAGE_TYPE,
    COLUMN_ROW_UID_MAP_MESSAGE_TYPE,
];
const TABLE_STYLE_MESSAGE_TYPE: u32 = 6_003;
const TABLE_STYLE_PRESET_MESSAGE_TYPE: u32 = 6_008;
const TABLE_STYLE_NETWORK_MESSAGE_TYPE: u32 = 6_247;
const STYLESHEET_MESSAGE_TYPE: u32 = 401;
const PACKAGE_METADATA_MESSAGE_TYPE: u32 = 11_006;
const SLIDE_OWNED_DRAWABLES_FIELD: u32 = 7;
const SLIDE_Z_ORDER_FIELD: u32 = 42;
const TABLE_SUPER_FIELD: u32 = 1;
const TABLE_MODEL_FIELD: u32 = 2;
const DRAWABLE_PARENT_FIELD: u32 = 2;
const MODEL_STORAGE_FIELD: u32 = 4;
const MODEL_ROW_BUCKET_PATH: &[u32] = &[MODEL_STORAGE_FIELD, 1, 2];
const MODEL_COLUMN_BUCKET_PATH: &[u32] = &[MODEL_STORAGE_FIELD, 2];
const MODEL_STRING_TABLE_PATH: &[u32] = &[MODEL_STORAGE_FIELD, 4];
const MODEL_STYLE_TABLE_PATH: &[u32] = &[MODEL_STORAGE_FIELD, 5];
const MODEL_FORMULA_TABLE_PATH: &[u32] = &[MODEL_STORAGE_FIELD, 6];
const MODEL_FORMAT_TABLE_PRE_BNC_PATH: &[u32] = &[MODEL_STORAGE_FIELD, 11];
const MODEL_FORMULA_ERROR_TABLE_PATH: &[u32] = &[MODEL_STORAGE_FIELD, 12];
const MODEL_MULTIPLE_CHOICE_TABLE_PATH: &[u32] = &[MODEL_STORAGE_FIELD, 16];
const MODEL_RICH_TEXT_TABLE_PATH: &[u32] = &[MODEL_STORAGE_FIELD, 17];
const MODEL_CONDITIONAL_STYLE_TABLE_PATH: &[u32] = &[MODEL_STORAGE_FIELD, 18];
const MODEL_COMMENT_STORAGE_TABLE_PATH: &[u32] = &[MODEL_STORAGE_FIELD, 19];
const MODEL_IMPORT_WARNING_TABLE_PATH: &[u32] = &[MODEL_STORAGE_FIELD, 20];
const MODEL_CONTROL_CELL_SPEC_TABLE_PATH: &[u32] = &[MODEL_STORAGE_FIELD, 21];
const MODEL_BASE_COLUMN_ROW_UIDS_FIELD: u32 = 46;
const MODEL_STROKE_SIDECAR_FIELD: u32 = 49;
const MODEL_HEADER_ROWS_FIELD: u32 = 9;
const MODEL_HEADER_COLUMNS_FIELD: u32 = 10;
const MODEL_FOOTER_ROWS_FIELD: u32 = 11;
const HEADER_BUCKET_ROWS: u32 = 65_536;

const ROLE_MESSAGE_TYPES: [u32; 10] = [
    TABLE_INFO_MESSAGE_TYPE,
    TABLE_MODEL_MESSAGE_TYPE,
    TILE_MESSAGE_TYPE,
    HEADER_BUCKET_MESSAGE_TYPE,
    COLUMN_ROW_UID_MAP_LEGACY_MESSAGE_TYPE,
    COLUMN_ROW_UID_MAP_MESSAGE_TYPE,
    TABLE_STYLE_MESSAGE_TYPE,
    TABLE_STYLE_PRESET_MESSAGE_TYPE,
    TABLE_STYLE_NETWORK_MESSAGE_TYPE,
    STYLESHEET_MESSAGE_TYPE,
];

/// Private resource categories charged by a complete table operation.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum LimitKind {
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

/// Private graph/codec failure.  The focused owner maps this to its public
/// content-free error enum.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) enum Error {
    UnsupportedSource,
    UnsupportedDependency,
    UnsupportedTopology,
    AmbiguousSelector,
    EmptySlideName,
    SlideNameNotFound,
    SlidePositionNotFound(Position),
    TablePositionNotFound(Position),
    InvalidSource,
    Limit {
        kind: LimitKind,
        observed: u64,
        maximum: u64,
    },
    Allocation(usize),
    Read,
    Wire,
    Codec,
    Archive,
    Verification,
}

pub(crate) type Result<T> = std::result::Result<T, Error>;

/// One source/candidate operation ledger.  The name owner carries this value
/// from admission through rewrite, candidate reopen, locality, and patch
/// application so no phase receives a fresh unlimited budget.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct Budget {
    max_input: usize,
    max_output: usize,
    max_fields: usize,
    max_work: usize,
    max_nesting: usize,
    max_references: usize,
    max_allocations: usize,
    max_retained: usize,
    max_scratch: usize,
    max_entries: usize,
    max_entry_bytes: usize,
    max_total_bytes: usize,
    max_payload_objects: usize,
    max_payload_messages: usize,
    max_components: usize,
    input: usize,
    output: usize,
    fields: usize,
    work: usize,
    nesting: usize,
    references: usize,
    allocations: usize,
    retained: usize,
    scratch: usize,
    entries: usize,
    entry_bytes: usize,
    total_bytes: usize,
    payload_objects: usize,
    payload_messages: usize,
    components: usize,
}

impl Budget {
    pub(crate) fn new(package: &Package) -> Result<Self> {
        let wire = package.wire_limits().map_err(|_| Error::Wire)?;
        let source = usize::try_from(package.state.options.archive().max_input_bytes())
            .map_err(|_| Error::InvalidSource)?;
        let archive = package
            .state
            .options
            .archive()
            .effective_archive_limits()
            .map_err(|_| Error::Archive)?;
        // Admission, rewrite, candidate validation, locality, and patch
        // inverse are all charged against one conservative bounded envelope.
        // The physical package itself remains subject to the original ZIP,
        // Snappy, and IWA limits; this multiplier only covers the repeated
        // logical passes performed by a focused transaction.
        let passes = 8usize;
        let aggregate = source.checked_mul(passes).ok_or(Error::InvalidSource)?;
        let max_entries = package
            .state
            .options
            .archive()
            .max_entries()
            .checked_mul(passes)
            .ok_or(Error::InvalidSource)?;
        let max_total = usize::try_from(package.state.options.archive().max_total_bytes())
            .map_err(|_| Error::InvalidSource)?
            .checked_mul(passes)
            .ok_or(Error::InvalidSource)?;
        let max_components = max_entries;
        let max_objects = archive
            .max_objects()
            .checked_mul(max_components)
            .ok_or(Error::InvalidSource)?;
        let max_messages = archive
            .max_messages()
            .checked_mul(max_components)
            .ok_or(Error::InvalidSource)?;
        Ok(Self {
            max_input: aggregate,
            max_output: aggregate,
            max_fields: wire.max_fields(),
            max_work: wire.max_rewrite_work(),
            max_nesting: wire.max_nesting(),
            max_references: package.semantic_limits().max_references(),
            max_allocations: aggregate,
            max_retained: aggregate,
            max_scratch: aggregate,
            max_entries,
            max_entry_bytes: max_total,
            max_total_bytes: max_total,
            max_payload_objects: max_objects,
            max_payload_messages: max_messages,
            max_components,
            input: 0,
            output: 0,
            fields: 0,
            work: 0,
            nesting: 0,
            references: 0,
            allocations: 0,
            retained: 0,
            scratch: 0,
            entries: 0,
            entry_bytes: 0,
            total_bytes: 0,
            payload_objects: 0,
            payload_messages: 0,
            components: 0,
        })
    }

    fn add(current: &mut usize, amount: usize, maximum: usize, kind: LimitKind) -> Result<()> {
        let observed = current.checked_add(amount).ok_or(Error::InvalidSource)?;
        if observed > maximum {
            return Err(Error::Limit {
                kind,
                observed: observed as u64,
                maximum: maximum as u64,
            });
        }
        *current = observed;
        Ok(())
    }

    pub(crate) fn input(&mut self, amount: usize) -> Result<()> {
        Self::add(
            &mut self.input,
            amount,
            self.max_input,
            LimitKind::InputBytes,
        )
    }

    pub(crate) fn output(&mut self, amount: usize) -> Result<()> {
        Self::add(
            &mut self.output,
            amount,
            self.max_output,
            LimitKind::OutputBytes,
        )
    }

    pub(crate) fn fields(&mut self, amount: usize) -> Result<()> {
        Self::add(
            &mut self.fields,
            amount,
            self.max_fields,
            LimitKind::WireFields,
        )
    }

    pub(crate) fn work(&mut self, amount: usize) -> Result<()> {
        Self::add(&mut self.work, amount, self.max_work, LimitKind::WireWork)
    }

    pub(crate) fn references(&mut self, amount: usize) -> Result<()> {
        Self::add(
            &mut self.references,
            amount,
            self.max_references,
            LimitKind::References,
        )
    }

    pub(crate) fn allocations(&mut self, amount: usize) -> Result<()> {
        Self::add(
            &mut self.allocations,
            amount,
            self.max_allocations,
            LimitKind::Allocations,
        )
    }

    pub(crate) fn retained(&mut self, amount: usize) -> Result<()> {
        Self::add(
            &mut self.retained,
            amount,
            self.max_retained,
            LimitKind::Retained,
        )
    }

    pub(crate) fn scratch(&mut self, amount: usize) -> Result<()> {
        Self::add(
            &mut self.scratch,
            amount,
            self.max_scratch,
            LimitKind::Scratch,
        )
    }

    pub(crate) fn entries(&mut self, amount: usize) -> Result<()> {
        Self::add(
            &mut self.entries,
            amount,
            self.max_entries,
            LimitKind::Entries,
        )
    }

    pub(crate) fn entry_bytes(&mut self, amount: usize) -> Result<()> {
        Self::add(
            &mut self.entry_bytes,
            amount,
            self.max_entry_bytes,
            LimitKind::EntryBytes,
        )
    }

    pub(crate) fn total_bytes(&mut self, amount: usize) -> Result<()> {
        Self::add(
            &mut self.total_bytes,
            amount,
            self.max_total_bytes,
            LimitKind::TotalBytes,
        )
    }

    pub(crate) fn payload_objects(&mut self, amount: usize) -> Result<()> {
        Self::add(
            &mut self.payload_objects,
            amount,
            self.max_payload_objects,
            LimitKind::PayloadObjects,
        )
    }

    pub(crate) fn payload_messages(&mut self, amount: usize) -> Result<()> {
        Self::add(
            &mut self.payload_messages,
            amount,
            self.max_payload_messages,
            LimitKind::PayloadMessages,
        )
    }

    pub(crate) fn components(&mut self, amount: usize) -> Result<()> {
        Self::add(
            &mut self.components,
            amount,
            self.max_components,
            LimitKind::Components,
        )
    }

    pub(crate) fn physical(&mut self, amount: usize) -> Result<()> {
        self.input(amount)?;
        self.entry_bytes(amount)?;
        self.total_bytes(amount)?;
        self.work(amount)
    }

    fn remaining(used: usize, maximum: usize, kind: LimitKind) -> Result<usize> {
        maximum
            .checked_sub(used)
            .filter(|remaining| *remaining > 0)
            .ok_or(Error::Limit {
                kind,
                observed: used
                    .checked_add(1)
                    .map_or(u64::MAX, |observed| observed as u64),
                maximum: maximum as u64,
            })
    }

    pub(crate) fn remaining_input(&self) -> Result<usize> {
        Self::remaining(self.input, self.max_input, LimitKind::InputBytes)
    }

    pub(crate) fn remaining_output(&self) -> Result<usize> {
        Self::remaining(self.output, self.max_output, LimitKind::OutputBytes)
    }

    pub(crate) fn remaining_work(&self) -> Result<usize> {
        Self::remaining(self.work, self.max_work, LimitKind::WireWork)
    }

    pub(crate) fn remaining_allocations(&self) -> Result<usize> {
        Self::remaining(
            self.allocations,
            self.max_allocations,
            LimitKind::Allocations,
        )
    }

    pub(crate) fn remaining_references(&self) -> Result<usize> {
        Self::remaining(self.references, self.max_references, LimitKind::References)
    }

    pub(crate) fn remaining_retained(&self) -> Result<usize> {
        Self::remaining(self.retained, self.max_retained, LimitKind::Retained)
    }

    pub(crate) fn remaining_scratch(&self) -> Result<usize> {
        Self::remaining(self.scratch, self.max_scratch, LimitKind::Scratch)
    }

    /// Build one residual table-model policy for either a read or prepared
    /// rewrite.  Every allocation-bearing ceiling is derived from the same
    /// operation ledger; callers must not substitute a source-sized or
    /// unlimited value for these fields.
    pub(crate) fn model_codec_options(
        &self,
        package: &Package,
        payload: &[u8],
    ) -> Result<table_model_discovery_codec::DecodeOptions> {
        let limits = self.residual(package)?;
        let allocations = self.remaining_allocations()?;
        let retained = self.remaining_retained()?;
        let scratch = self.remaining_scratch()?;
        Ok(
            table_model_discovery_codec::DecodeOptions::for_source(payload)
                .with_max_input_bytes(limits.max_input_bytes().min(payload.len().max(1)))
                .with_max_fields(limits.max_fields())
                .with_max_work_bytes(limits.max_rewrite_work())
                .with_max_text_bytes(limits.max_input_bytes())
                .with_recursion_limit(u32::try_from(limits.max_nesting()).unwrap_or(u32::MAX))
                .with_max_output_bytes(self.remaining_output()?.min(limits.max_output_bytes()))
                .with_max_allocations(allocations)
                .with_max_retained_bytes(retained)
                .with_max_scratch_bytes(scratch),
        )
    }

    pub(crate) fn storage_codec_options(
        &self,
        package: &Package,
    ) -> Result<table_dimension_codec::DecodeOptions> {
        let limits = self.residual(package)?;
        let references =
            Self::remaining(self.references, self.max_references, LimitKind::References)?;
        Ok(table_dimension_codec::DecodeOptions::new(
            limits.max_input_bytes(),
            limits.max_fields(),
            limits.max_rewrite_work(),
            u32::try_from(limits.max_nesting()).unwrap_or(u32::MAX),
            references,
            limits.max_input_bytes(),
        ))
    }

    /// Charge the logical workspaces needed before a native archive is
    /// decompressed and parsed.  The archive parser owns its internal
    /// allocations, so these are conservative source-sized envelopes; the
    /// parsed object/message inventory is charged after successful parsing.
    pub(crate) fn preflight_archive_parse(&mut self, compressed: usize) -> Result<()> {
        self.allocations(1)?;
        self.retained(compressed)?;
        self.scratch(compressed)?;
        self.work(compressed)
    }

    pub(crate) fn preflight_decoded_archive(&mut self, decoded: usize) -> Result<()> {
        self.allocations(1)?;
        self.retained(decoded)?;
        self.scratch(decoded)?;
        self.work(decoded)
    }

    /// Precharge the package-owned source buffer and parser/cache envelope
    /// before handing bytes to `Package::from_source_with_options`.
    pub(crate) fn preflight_candidate(&mut self, source: &Package, bytes: usize) -> Result<()> {
        let (objects, messages) = package_inventory_counts(source)?;
        self.input(bytes)?;
        self.work(
            bytes
                .checked_add(objects)
                .and_then(|value| value.checked_add(messages))
                .ok_or(Error::InvalidSource)?,
        )?;
        self.allocations(
            objects
                .checked_add(messages)
                .and_then(|value| value.checked_add(1))
                .ok_or(Error::InvalidSource)?,
        )?;
        self.retained(bytes)?;
        self.scratch(bytes)
    }

    pub(crate) fn preflight_semantic_scan(&mut self, package: &Package) -> Result<()> {
        let (objects, messages) = package_inventory_counts(package)?;
        let scan = objects
            .checked_add(messages)
            .and_then(|value| value.checked_add(1))
            .ok_or(Error::InvalidSource)?;
        self.work(scan)?;
        self.allocations(1)?;
        self.retained(scan)
    }

    pub(crate) fn preflight_preview_scan(&mut self, entries: usize) -> Result<()> {
        self.work(entries.checked_add(3).ok_or(Error::InvalidSource)?)?;
        self.allocations(1)
    }

    pub(crate) fn preflight_reassembly(
        &mut self,
        catalog: &SourceCatalog,
        edited_bytes: usize,
        deleted: usize,
    ) -> Result<()> {
        let entries = catalog.package().len();
        self.entries(entries)?;
        self.entry_bytes(catalog.source_bytes().len())?;
        self.total_bytes(catalog.source_bytes().len())?;
        self.work(
            catalog
                .source_bytes()
                .len()
                .checked_add(edited_bytes)
                .and_then(|value| value.checked_add(entries))
                .ok_or(Error::InvalidSource)?,
        )?;
        self.allocations(
            entries
                .checked_add(deleted)
                .and_then(|value| value.checked_add(2))
                .ok_or(Error::InvalidSource)?,
        )?;
        self.retained(catalog.source_bytes().len())?;
        self.scratch(
            entries
                .checked_mul(size_of::<usize>())
                .and_then(|value| value.checked_add(edited_bytes))
                .ok_or(Error::InvalidSource)?,
        )
    }

    pub(crate) fn artifact(&mut self, source: usize, target: usize) -> Result<()> {
        self.work(source.checked_add(target).ok_or(Error::InvalidSource)?)?;
        self.retained(source.checked_add(target).ok_or(Error::InvalidSource)?)?;
        self.allocations(1)
    }

    /// Charge an owned source-derived value before constructing its storage.
    pub(crate) fn owned_value(&mut self, bytes: usize) -> Result<()> {
        self.allocations(bytes.max(1))?;
        self.retained(bytes)?;
        self.work(bytes)
    }

    pub(crate) fn nesting(&mut self, value: usize) -> Result<()> {
        self.nesting = self.nesting.max(value);
        if self.nesting > self.max_nesting {
            return Err(Error::Limit {
                kind: LimitKind::WireNesting,
                observed: self.nesting as u64,
                maximum: self.max_nesting as u64,
            });
        }
        Ok(())
    }

    pub(crate) fn codec_report(
        &mut self,
        report: table_model_discovery_codec::DecodeReport,
    ) -> Result<()> {
        self.input(report.input_bytes())?;
        self.fields(report.fields())?;
        self.work(report.work_bytes())?;
        self.allocations(report.allocations())?;
        self.retained(report.retained_bytes())?;
        self.scratch(report.scratch_bytes())?;
        self.nesting(report.max_depth() as usize)
    }

    /// Charge strict persisted table-sort decoding and preparation against
    /// the same operation ledger used by graph admission and archive work.
    pub(crate) fn sort_codec_report(
        &mut self,
        report: table_sort_order_codec::DecodeReport,
    ) -> Result<()> {
        self.input(report.input_bytes())?;
        self.output(report.output_bytes())?;
        self.fields(report.fields())?;
        self.work(report.work_bytes())?;
        self.references(report.rules())?;
        self.allocations(report.allocations())?;
        self.retained(report.retained_bytes())?;
        self.scratch(report.scratch_bytes())?;
        self.nesting(report.max_depth() as usize)
    }

    pub(crate) fn sort_rewrite_requirements(
        &mut self,
        requirements: table_sort_order_codec::RewriteExecutionRequirements,
    ) -> Result<()> {
        self.output(requirements.output_bytes)?;
        self.fields(requirements.fields)?;
        self.work(requirements.work_bytes)?;
        self.references(requirements.rules)?;
        self.allocations(requirements.allocations)?;
        self.retained(requirements.retained_bytes)?;
        self.scratch(requirements.scratch_bytes)?;
        self.nesting(requirements.max_depth as usize)
    }

    pub(crate) fn storage_codec_report(
        &mut self,
        report: table_dimension_codec::DecodeReport,
    ) -> Result<()> {
        self.input(report.source_bytes())?;
        self.fields(report.fields())?;
        self.work(report.work_bytes())?;
        self.references(report.references())?;
        self.nesting(report.max_depth() as usize)
    }

    pub(crate) fn name_rewrite_requirements(
        &mut self,
        requirements: table_model_discovery_codec::TableModelNameRewriteRequirements,
    ) -> Result<()> {
        self.input(requirements.input_bytes())?;
        self.output(requirements.output_bytes())?;
        self.fields(requirements.fields())?;
        self.work(requirements.work_bytes())?;
        self.allocations(requirements.allocations())?;
        self.retained(requirements.retained_bytes())?;
        self.scratch(requirements.scratch_bytes())?;
        self.nesting(requirements.max_depth() as usize)
    }

    pub(crate) fn metadata_report(
        &mut self,
        report: package_metadata_codec::RewriteReport,
    ) -> Result<()> {
        self.input(report.input_bytes())?;
        self.output(report.output_bytes())?;
        self.fields(report.fields())?;
        self.work(report.work_bytes())?;
        self.components(report.components_scanned())?;
        self.references(
            report
                .references_scanned()
                .checked_add(report.source_references_scanned())
                .ok_or(Error::InvalidSource)?,
        )?;
        self.allocations(report.allocations())?;
        self.retained(report.retained_bytes())?;
        self.scratch(report.scratch_bytes())?;
        self.nesting(report.max_depth() as usize)
    }

    pub(crate) fn reassembly(
        &mut self,
        requirements: litchi_iwa_archive::package::ReassemblyExecutionRequirements,
    ) -> Result<()> {
        self.output(requirements.output_bytes())?;
        self.allocations(requirements.allocations())?;
        self.retained(requirements.retained_bytes())?;
        self.scratch(requirements.scratch_bytes())?;
        self.work(requirements.output_bytes())
    }

    pub(crate) fn inventory(&mut self, package: &Package, catalog: &SourceCatalog) -> Result<()> {
        self.entries(catalog.package().len())?;
        self.components(package.state.source.components().len())?;
        self.allocations(
            catalog
                .package()
                .len()
                .checked_add(package.state.source.components().len())
                .ok_or(Error::InvalidSource)?,
        )?;
        for entry in catalog.package().iter() {
            self.entry_bytes(entry.data().len())?;
            self.total_bytes(entry.data().len())?;
            self.work(
                entry
                    .name()
                    .len()
                    .checked_add(entry.data().len())
                    .ok_or(Error::InvalidSource)?,
            )?;
        }
        for component in package.state.source.components().iter() {
            self.payload_objects(component.archive().objects.len())?;
            let messages = component
                .archive()
                .objects
                .iter()
                .map(|object| object.messages.len())
                .try_fold(0usize, |sum, value| sum.checked_add(value))
                .ok_or(Error::InvalidSource)?;
            self.payload_messages(messages)?;
            let mut component_work = 0usize;
            for object in &component.archive().objects {
                component_work = component_work
                    .checked_add(object.messages.len())
                    .ok_or(Error::InvalidSource)?;
                for message in &object.messages {
                    component_work = component_work
                        .checked_add(message.data.len())
                        .ok_or(Error::InvalidSource)?;
                }
            }
            self.work(component_work)?;
        }
        Ok(())
    }

    /// Charge the parsed component/object/message inventory when the ZIP
    /// catalog has already been consumed by a shared read handoff.
    pub(crate) fn inventory_components(&mut self, package: &Package) -> Result<()> {
        let components = package.state.source.components();
        self.components(components.len())?;
        self.allocations(components.len())?;
        for component in components.iter() {
            self.payload_objects(component.archive().objects.len())?;
            let messages = component
                .archive()
                .objects
                .iter()
                .map(|object| object.messages.len())
                .try_fold(0usize, |sum, value| sum.checked_add(value))
                .ok_or(Error::InvalidSource)?;
            self.payload_messages(messages)?;
            let mut component_work = 0usize;
            for object in &component.archive().objects {
                component_work = component_work
                    .checked_add(object.messages.len())
                    .ok_or(Error::InvalidSource)?;
                for message in &object.messages {
                    component_work = component_work
                        .checked_add(message.data.len())
                        .ok_or(Error::InvalidSource)?;
                }
            }
            self.work(component_work)?;
        }
        Ok(())
    }

    pub(crate) fn residual(&self, package: &Package) -> Result<WireLimits> {
        let base = package.wire_limits().map_err(|_| Error::Wire)?;
        let input = self.remaining_input()?;
        let output = self.remaining_output()?;
        let fields = Self::remaining(self.fields, self.max_fields, LimitKind::WireFields)?;
        let work = self.remaining_work()?;
        let limits = base
            .with_input_bytes(base.max_input_bytes().min(input))
            .map_err(|_| Error::Wire)?;
        let limits = limits
            .with_output_bytes(base.max_output_bytes().min(output))
            .map_err(|_| Error::Wire)?;
        let limits = limits
            .with_fields(base.max_fields().min(fields))
            .map_err(|_| Error::Wire)?;
        let limits = limits
            .with_rewrite_work(base.max_rewrite_work().min(work))
            .map_err(|_| Error::Wire)?;
        limits
            .with_nesting(base.max_nesting().min(self.max_nesting))
            .map_err(|_| Error::Wire)
    }
}

fn package_inventory_counts(package: &Package) -> Result<(usize, usize)> {
    let mut objects = 0usize;
    let mut messages = 0usize;
    for component in package.state.source.components().iter() {
        objects = objects
            .checked_add(component.archive().objects.len())
            .ok_or(Error::InvalidSource)?;
        for object in &component.archive().objects {
            messages = messages
                .checked_add(object.messages.len())
                .ok_or(Error::InvalidSource)?;
        }
    }
    Ok((objects, messages))
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct ObjectLocation {
    pub(crate) identifier: u64,
    pub(crate) component: Arc<str>,
    pub(crate) object_index: usize,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum StorageRouteKind {
    RowBucket,
    ColumnBucket,
    StringTable,
    StyleTable,
    FormulaTable,
    FormatTablePreBnc,
}

impl StorageRouteKind {
    pub(crate) const fn path(self) -> &'static [u32] {
        match self {
            Self::RowBucket => MODEL_ROW_BUCKET_PATH,
            Self::ColumnBucket => MODEL_COLUMN_BUCKET_PATH,
            Self::StringTable => MODEL_STRING_TABLE_PATH,
            Self::StyleTable => MODEL_STYLE_TABLE_PATH,
            Self::FormulaTable => MODEL_FORMULA_TABLE_PATH,
            Self::FormatTablePreBnc => MODEL_FORMAT_TABLE_PRE_BNC_PATH,
        }
    }

    pub(crate) const fn is_bucket(self) -> bool {
        matches!(self, Self::RowBucket | Self::ColumnBucket)
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct StorageRoute {
    pub(crate) location: ObjectLocation,
    pub(crate) kind: StorageRouteKind,
}

/// One source-ordered tile entry in a physical table's `TileStorage` root.
///
/// The key is the native tile coordinate. The object location is an internal
/// routing detail and never appears in a public Keynote value or error.
#[derive(Debug, Clone, PartialEq, Eq)]
#[allow(dead_code, reason = "Consumed by the physical table-sort owner.")]
pub(crate) struct PhysicalTileReference {
    pub(crate) tile_id: u32,
    pub(crate) location: ObjectLocation,
}

/// One row-header bucket in source order. Header records are sparse inside a
/// bucket; the bucket index is still part of the topology proof because a
/// moved row must stay within an allocated bucket.
#[derive(Debug, Clone, PartialEq, Eq)]
#[allow(dead_code, reason = "Consumed by the physical table-sort owner.")]
pub(crate) struct PhysicalRowHeaderBucket {
    pub(crate) bucket_index: u32,
    pub(crate) location: ObjectLocation,
    pub(crate) header_count: usize,
}

/// Feature admissions that would require a row-affine rewrite beyond the
/// currently owned physical sort transaction.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
#[allow(
    dead_code,
    reason = "Retained on the private topology for diagnostics."
)]
pub(crate) struct PhysicalUnsupportedFeatures {
    hidden: bool,
    filter: bool,
    grouping: bool,
    pivot: bool,
    spill: bool,
    merge: bool,
    conditional: bool,
    formula_affine: bool,
    imported_data: bool,
}

impl PhysicalUnsupportedFeatures {
    #[must_use]
    pub(crate) const fn any(self) -> bool {
        self.hidden
            || self.filter
            || self.grouping
            || self.pivot
            || self.spill
            || self.merge
            || self.conditional
            || self.formula_affine
            || self.imported_data
    }
}

/// A validated physical storage spine and row-affine admission for a table.
///
/// This value intentionally contains no mutable archive state. It is safe to
/// pass between preparation stages or threads; the source package remains the
/// authority and every writer must revalidate its exact source before
/// publishing a candidate.
#[derive(Debug, Clone, PartialEq, Eq)]
#[allow(dead_code, reason = "Consumed by the physical table-sort owner.")]
pub(crate) struct PhysicalTableTopology {
    tile_size: u32,
    wide_rows: Option<bool>,
    tile_references: Vec<PhysicalTileReference>,
    row_header_buckets: Vec<PhysicalRowHeaderBucket>,
    column_headers: ObjectLocation,
    string_table: ObjectLocation,
    row_uid_map: ObjectLocation,
    stroke_sidecar: Option<ObjectLocation>,
    mutable_component_count: usize,
    has_formula_entries: bool,
    has_formula_error_entries: bool,
    has_conditional_style_entries: bool,
    has_comment_entries: bool,
    has_rich_text_entries: bool,
    header_rows: u32,
    header_columns: u32,
    footer_rows: u32,
    unsupported: PhysicalUnsupportedFeatures,
}

impl PhysicalTableTopology {
    #[must_use]
    pub(crate) const fn tile_size(&self) -> u32 {
        self.tile_size
    }

    #[must_use]
    pub(crate) const fn wide_rows(&self) -> Option<bool> {
        self.wide_rows
    }

    #[must_use]
    pub(crate) fn tile_references(&self) -> &[PhysicalTileReference] {
        &self.tile_references
    }

    #[must_use]
    pub(crate) fn row_header_buckets(&self) -> &[PhysicalRowHeaderBucket] {
        &self.row_header_buckets
    }

    #[must_use]
    pub(crate) const fn column_headers(&self) -> &ObjectLocation {
        &self.column_headers
    }

    #[must_use]
    pub(crate) const fn string_table(&self) -> &ObjectLocation {
        &self.string_table
    }

    #[must_use]
    pub(crate) const fn row_uid_map(&self) -> &ObjectLocation {
        &self.row_uid_map
    }

    #[must_use]
    #[allow(dead_code, reason = "Consumed by future physical table owners.")]
    pub(crate) const fn stroke_sidecar(&self) -> Option<&ObjectLocation> {
        self.stroke_sidecar.as_ref()
    }

    #[must_use]
    pub(crate) const fn mutable_component_count(&self) -> usize {
        self.mutable_component_count
    }

    #[must_use]
    pub(crate) const fn has_formula_entries(&self) -> bool {
        self.has_formula_entries
    }

    #[must_use]
    pub(crate) const fn has_formula_error_entries(&self) -> bool {
        self.has_formula_error_entries
    }

    #[must_use]
    pub(crate) const fn has_conditional_style_entries(&self) -> bool {
        self.has_conditional_style_entries
    }

    #[must_use]
    pub(crate) const fn has_comment_entries(&self) -> bool {
        self.has_comment_entries
    }

    #[must_use]
    pub(crate) const fn has_rich_text_entries(&self) -> bool {
        self.has_rich_text_entries
    }

    #[must_use]
    pub(crate) const fn header_rows(&self) -> u32 {
        self.header_rows
    }

    #[must_use]
    #[allow(dead_code, reason = "Consumed by future physical table owners.")]
    pub(crate) const fn header_columns(&self) -> u32 {
        self.header_columns
    }

    #[must_use]
    pub(crate) const fn footer_rows(&self) -> u32 {
        self.footer_rows
    }

    #[must_use]
    #[allow(dead_code, reason = "Consumed by future physical table owners.")]
    pub(crate) const fn unsupported_features(&self) -> PhysicalUnsupportedFeatures {
        self.unsupported
    }
}

/// Explicit candidate-locality admission for a physical transaction.
///
/// Changed component and preview-entry names are normalized to sorted,
/// duplicate-free lists at construction. This makes locality decisions
/// deterministic even when a caller derives the allowlist from a map or a
/// parallel preparation stage.
#[derive(Debug, Clone, PartialEq, Eq)]
#[allow(dead_code, reason = "Consumed by the physical table-sort owner.")]
pub(crate) struct LocalityAllowlist {
    changed_components: Vec<Arc<str>>,
    root_preview_deletions: Vec<Arc<str>>,
    changed_object_ids: Vec<u64>,
}

impl LocalityAllowlist {
    pub(crate) fn new(
        changed_components: &[&str],
        root_preview_deletions: &[&str],
    ) -> Result<Self> {
        let mut changed = Vec::new();
        changed
            .try_reserve_exact(changed_components.len())
            .map_err(|_| Error::Allocation(changed_components.len()))?;
        for component in changed_components {
            if component.is_empty() {
                return Err(Error::InvalidSource);
            }
            changed.push(Arc::from(*component));
        }
        changed.sort_unstable();
        changed.dedup();

        let mut deletions = Vec::new();
        deletions
            .try_reserve_exact(root_preview_deletions.len())
            .map_err(|_| Error::Allocation(root_preview_deletions.len()))?;
        for entry in root_preview_deletions {
            if entry.is_empty() {
                return Err(Error::InvalidSource);
            }
            deletions.push(Arc::from(*entry));
        }
        deletions.sort_unstable();
        deletions.dedup();
        Ok(Self {
            changed_components: changed,
            root_preview_deletions: deletions,
            changed_object_ids: Vec::new(),
        })
    }

    /// Admit exactly the component members that can contain row-affine
    /// objects for one validated physical table topology.
    ///
    /// Native Keynote packages commonly split the tile, row-header bucket,
    /// and row/column UID map across distinct IWA members.  The object-ID
    /// fence applied by [`Self::with_changed_object_ids`] remains authoritative
    /// inside these component envelopes.
    pub(crate) fn with_physical_topology(
        topology: &PhysicalTableTopology,
        root_preview_deletions: &[&str],
    ) -> Result<Self> {
        let capacity = 1usize
            .checked_add(topology.tile_references.len())
            .and_then(|value| value.checked_add(topology.row_header_buckets.len()))
            .ok_or(Error::InvalidSource)?;
        let mut components = Vec::new();
        components
            .try_reserve_exact(capacity)
            .map_err(|_| Error::Allocation(capacity))?;
        components.push(topology.row_uid_map.component.as_ref());
        components.extend(
            topology
                .tile_references
                .iter()
                .map(|tile| tile.location.component.as_ref()),
        );
        components.extend(
            topology
                .row_header_buckets
                .iter()
                .map(|bucket| bucket.location.component.as_ref()),
        );
        Self::new(&components, root_preview_deletions)
    }

    /// Add the exact object identities a physical transaction is allowed to
    /// mutate.  Component locality remains useful as the ZIP envelope
    /// boundary, while this optional second fence prevents an unrelated tile
    /// or header object in the same component from becoming an accidental
    /// mutation wildcard.
    pub(crate) fn with_changed_object_ids(mut self, identifiers: &[u64]) -> Result<Self> {
        let mut objects = Vec::new();
        objects
            .try_reserve_exact(identifiers.len())
            .map_err(|_| Error::Allocation(identifiers.len()))?;
        for identifier in identifiers {
            if *identifier == 0 {
                return Err(Error::InvalidSource);
            }
            objects.push(*identifier);
        }
        objects.sort_unstable();
        objects.dedup();
        self.changed_object_ids = objects;
        Ok(self)
    }

    #[must_use]
    pub(crate) fn changed_components(&self) -> &[Arc<str>] {
        &self.changed_components
    }

    #[must_use]
    pub(crate) fn root_preview_deletions(&self) -> &[Arc<str>] {
        &self.root_preview_deletions
    }

    #[must_use]
    pub(crate) fn changed_object_ids(&self) -> &[u64] {
        &self.changed_object_ids
    }

    fn allows_component(&self, name: &str) -> bool {
        self.changed_components
            .binary_search_by(|candidate| candidate.as_ref().cmp(name))
            .is_ok()
    }

    fn allows_preview_deletion(&self, name: &str) -> bool {
        self.root_preview_deletions
            .binary_search_by(|candidate| candidate.as_ref().cmp(name))
            .is_ok()
    }

    fn allows_object(&self, identifier: u64) -> bool {
        self.changed_object_ids.binary_search(&identifier).is_ok()
    }
}

/// The proven native route for one selected slide table.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct Target {
    pub(crate) slide_position: Position,
    pub(crate) table_position: Position,
    pub(crate) slide: ObjectLocation,
    pub(crate) table_info: ObjectLocation,
    pub(crate) model: ObjectLocation,
    pub(crate) slide_message_index: usize,
    pub(crate) table_info_message_index: usize,
    pub(crate) model_message_index: usize,
    pub(crate) rows: u32,
    pub(crate) columns: u32,
    storage: Vec<StorageRoute>,
    pub(crate) locked: bool,
}

/// One rooted table display name and its canonical `TableInfo` identity.
///
/// Formula owner records refer to the native `TableInfo` object rather than
/// to a semantic slide/table position.  Keeping this small source-backed
/// projection private lets formula rendering resolve only identities proven
/// through the same slide ownership graph as table selection.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct RootedTableName {
    pub(crate) table_info_identifier: u64,
    pub(crate) name: Box<str>,
}

/// Collect names for all rooted tables in presentation order.
///
/// This is deliberately the same ownership path used by [`select_table`]: a
/// table must occur exactly once in both the slide's owned-drawable list and
/// its z-order, its `TableInfo.super.parent` must name that slide, and its
/// table-model message must be unique and canonical.  The caller uses the
/// resulting names only while resolving formula owner references; native
/// identifiers never cross the focused package API.
pub(crate) fn rooted_table_name_catalog(
    package: &Package,
    budget: &mut Budget,
) -> Result<Vec<RootedTableName>> {
    let maximum = package.semantic_limits().max_slides();
    let mut names = Vec::new();
    let mut seen_table_info = HashSet::new();

    for slide_position in 0..maximum {
        let Some(record) = package
            .slide_record_at(slide_position)
            .map_err(read_error)?
        else {
            break;
        };
        budget.work(1)?;
        let slide = locate_object(package, record.slide_identifier)?;
        let slide_object = object_at(package, &slide)?;
        let (slide_message_index, slide_payload) =
            unique_message(slide_object, SLIDE_MESSAGE_TYPE, budget)?;
        let limits = budget.residual(package)?;
        let owned =
            repeated_references(slide_payload, SLIDE_OWNED_DRAWABLES_FIELD, limits, budget)?;
        let z_order = repeated_references(slide_payload, SLIDE_Z_ORDER_FIELD, limits, budget)?;
        budget.references(
            owned
                .len()
                .checked_add(z_order.len())
                .ok_or(Error::InvalidSource)?,
        )?;
        reject_duplicates(&owned, budget)?;
        reject_duplicates(&z_order, budget)?;
        validate_slide_metadata(slide_object, slide_message_index, &owned, &z_order, budget)?;

        for table_info_identifier in z_order {
            let info_location = locate_object(package, table_info_identifier)?;
            let info_object = object_at(package, &info_location)?;
            let info_count = info_object
                .messages
                .iter()
                .filter(|message| message.type_ == TABLE_INFO_MESSAGE_TYPE)
                .count();
            let has_role_alias = info_object.messages.iter().any(|message| {
                ROLE_MESSAGE_TYPES.contains(&message.type_)
                    && message.type_ != TABLE_INFO_MESSAGE_TYPE
            });
            if info_count == 0 {
                if has_role_alias {
                    return Err(Error::UnsupportedDependency);
                }
                continue;
            }
            if info_count != 1 || has_role_alias {
                return Err(Error::UnsupportedDependency);
            }
            if owned
                .iter()
                .filter(|id| **id == table_info_identifier)
                .count()
                != 1
            {
                return Err(Error::InvalidSource);
            }
            let (info_message_index, info_payload) =
                unique_message(info_object, TABLE_INFO_MESSAGE_TYPE, budget)?;
            let info = decode_table_info(info_payload, package, budget)?;
            if table_parent(info_payload, limits, budget)? != record.slide_identifier {
                return Err(Error::InvalidSource);
            }
            let model_identifier = info.table_model().identifier().get();
            validate_table_info_metadata(
                package,
                info_object,
                info_message_index,
                model_identifier,
                budget,
            )?;
            let model_location = locate_object(package, model_identifier)?;
            let model_object = object_at(package, &model_location)?;
            let model_count = model_object
                .messages
                .iter()
                .filter(|message| message.type_ == TABLE_MODEL_MESSAGE_TYPE)
                .count();
            let has_model_role_alias = model_object.messages.iter().any(|message| {
                ROLE_MESSAGE_TYPES.contains(&message.type_)
                    && message.type_ != TABLE_MODEL_MESSAGE_TYPE
            });
            if model_count != 1 || has_model_role_alias {
                return Err(Error::UnsupportedDependency);
            }
            let (_model_message_index, model_payload) =
                unique_message(model_object, TABLE_MODEL_MESSAGE_TYPE, budget)?;
            let options = budget.model_codec_options(package, model_payload)?;
            let (model, report) =
                table_model_discovery_codec::decode_table_model_with_report(model_payload, options)
                    .map_err(|_| Error::Codec)?;
            budget.codec_report(report)?;
            budget.allocations(1)?;
            budget.retained(size_of::<u64>())?;
            seen_table_info
                .try_reserve(1)
                .map_err(|_| Error::Allocation(1))?;
            if !seen_table_info.insert(table_info_identifier) {
                return Err(Error::UnsupportedDependency);
            }
            budget.allocations(1)?;
            budget.retained(size_of::<RootedTableName>())?;
            let name = model.table_name();
            budget.allocations(1)?;
            budget.retained(name.len())?;
            budget.work(name.len())?;
            let mut owned_name = String::new();
            owned_name
                .try_reserve_exact(name.len())
                .map_err(|_| Error::Allocation(name.len()))?;
            owned_name.push_str(name);
            names.try_reserve(1).map_err(|_| Error::Allocation(1))?;
            names.push(RootedTableName {
                table_info_identifier,
                name: owned_name.into_boxed_str(),
            });
        }
    }

    Ok(names)
}

impl Target {
    #[must_use]
    pub(crate) fn storage_route(&self, kind: StorageRouteKind) -> Option<&StorageRoute> {
        self.storage.iter().find(|route| route.kind == kind)
    }
}

/// Resolve one canonical table and prove its package-wide physical authority.
pub(crate) fn select_table(
    package: &Package,
    slide_selector: SlideSelector<'_>,
    table_selector: TableSelector,
    budget: &mut Budget,
) -> Result<Target> {
    let catalog = physical_catalog(package)?;
    budget.input(package.source_bytes().len())?;
    budget.inventory(package, catalog)?;
    select_table_after_inventory(package, slide_selector, table_selector, budget, false)
}

/// Resolve a read-only table against an already checked component catalog.
///
/// The semantic-only package source intentionally remains unsupported by the
/// ordinary selector above because mutation and source-preserving owners need
/// ZIP provenance.  The metadata-only merge reader is the sole caller of this
/// ingress: it retains a checked [`ComponentCatalog`] and uses the same graph,
/// wire, and ownership proof after charging a component inventory pass.
pub(crate) fn select_table_from_components(
    package: &Package,
    slide_selector: SlideSelector<'_>,
    table_selector: TableSelector,
    budget: &mut Budget,
) -> Result<Target> {
    if !matches!(&package.state.source, PhysicalSource::Semantic(_)) {
        return Err(Error::UnsupportedSource);
    }
    budget.inventory_components(package)?;
    select_table_after_inventory(package, slide_selector, table_selector, budget, true)
}

fn select_table_after_inventory(
    package: &Package,
    slide_selector: SlideSelector<'_>,
    table_selector: TableSelector,
    budget: &mut Budget,
    metadata_slide_names: bool,
) -> Result<Target> {
    let slide_position = if metadata_slide_names {
        resolve_slide_position_from_components(package, slide_selector, budget)?
    } else {
        resolve_slide_position(package, slide_selector, budget)?
    };
    let record = package
        .slide_record_at(slide_position.get())
        .map_err(read_error)?
        .ok_or(Error::SlidePositionNotFound(slide_position))?;
    let slide = locate_object(package, record.slide_identifier)?;
    let slide_object = object_at(package, &slide)?;
    let (slide_message_index, slide_payload) =
        unique_message(slide_object, SLIDE_MESSAGE_TYPE, budget)?;
    let limits = budget.residual(package)?;
    let owned = repeated_references(slide_payload, SLIDE_OWNED_DRAWABLES_FIELD, limits, budget)?;
    let z_order = repeated_references(slide_payload, SLIDE_Z_ORDER_FIELD, limits, budget)?;
    budget.references(
        owned
            .len()
            .checked_add(z_order.len())
            .ok_or(Error::InvalidSource)?,
    )?;
    reject_duplicates(&owned, budget)?;
    reject_duplicates(&z_order, budget)?;
    validate_slide_metadata(slide_object, slide_message_index, &owned, &z_order, budget)?;

    let mut candidates = Vec::new();
    budget.allocations(z_order.len())?;
    candidates
        .try_reserve_exact(z_order.len())
        .map_err(|_| Error::Allocation(z_order.len()))?;
    for table_info_identifier in z_order {
        let info_location = locate_object(package, table_info_identifier)?;
        let info_object = object_at(package, &info_location)?;
        let info_count = info_object
            .messages
            .iter()
            .filter(|message| message.type_ == TABLE_INFO_MESSAGE_TYPE)
            .count();
        let has_role_alias = info_object.messages.iter().any(|message| {
            ROLE_MESSAGE_TYPES.contains(&message.type_) && message.type_ != TABLE_INFO_MESSAGE_TYPE
        });
        if info_count == 0 {
            if has_role_alias {
                return Err(Error::UnsupportedDependency);
            }
            continue;
        }
        if info_count != 1 || has_role_alias {
            return Err(Error::UnsupportedDependency);
        }
        if owned
            .iter()
            .filter(|id| **id == table_info_identifier)
            .count()
            != 1
        {
            return Err(Error::InvalidSource);
        }
        let (info_message_index, info_payload) =
            unique_message(info_object, TABLE_INFO_MESSAGE_TYPE, budget)?;
        let info = decode_table_info(info_payload, package, budget)?;
        let parent = table_parent(info_payload, limits, budget)?;
        if parent != record.slide_identifier {
            return Err(Error::InvalidSource);
        }
        let model_identifier = info.table_model().identifier().get();
        validate_table_info_metadata(
            package,
            info_object,
            info_message_index,
            model_identifier,
            budget,
        )?;
        let model_location = locate_object(package, model_identifier)?;
        let model_object = object_at(package, &model_location)?;
        let model_count = model_object
            .messages
            .iter()
            .filter(|message| message.type_ == TABLE_MODEL_MESSAGE_TYPE)
            .count();
        let has_model_role_alias = model_object.messages.iter().any(|message| {
            ROLE_MESSAGE_TYPES.contains(&message.type_) && message.type_ != TABLE_MODEL_MESSAGE_TYPE
        });
        if model_count != 1 || has_model_role_alias {
            return Err(Error::UnsupportedDependency);
        }
        let (model_message_index, model_payload) =
            unique_message(model_object, TABLE_MODEL_MESSAGE_TYPE, budget)?;
        let (rows, columns) = validate_model_payload(model_payload, package, budget)?;
        candidates.push((
            info_location,
            model_location,
            info_message_index,
            model_message_index,
            rows,
            columns,
            info.locked().unwrap_or(false),
        ));
    }
    let table_position = table_selector.as_position();
    let (table_info, model, table_info_message_index, model_message_index, rows, columns, locked) =
        candidates
            .into_iter()
            .nth(table_position.get())
            .ok_or(Error::TablePositionNotFound(table_position))?;
    ensure_unique_identity(package, slide.identifier, budget)?;
    ensure_unique_identity(package, table_info.identifier, budget)?;
    ensure_unique_identity(package, model.identifier, budget)?;
    let model_object = object_at(package, &model)?;
    let model_payload = model_object
        .messages
        .get(model_message_index)
        .filter(|message| message.type_ == TABLE_MODEL_MESSAGE_TYPE)
        .map(|message| message.data.as_slice())
        .ok_or(Error::InvalidSource)?;
    let storage = validate_model_storage(
        package,
        model.identifier,
        model_message_index,
        model_payload,
        budget,
    )?;
    ensure_unique_table_owner(
        package,
        slide.identifier,
        table_info.identifier,
        model.identifier,
        limits,
        budget,
    )?;
    validate_package_metadata(package, [&slide, &table_info, &model], &storage, budget)?;
    validate_global_inbound_references(
        package,
        slide.identifier,
        slide_message_index,
        table_info.identifier,
        table_info_message_index,
        model.identifier,
        model_message_index,
        &storage,
        budget,
    )?;
    Ok(Target {
        slide_position,
        table_position,
        slide,
        table_info,
        model,
        slide_message_index,
        table_info_message_index,
        model_message_index,
        rows,
        columns,
        storage,
        locked,
    })
}

pub(crate) fn same_target(left: &Target, right: &Target) -> bool {
    left == right
}

fn validate_physical_root_roles(
    package: &Package,
    target: &Target,
    budget: &mut Budget,
) -> Result<()> {
    for (location, message_type) in [
        (&target.slide, SLIDE_MESSAGE_TYPE),
        (&target.table_info, TABLE_INFO_MESSAGE_TYPE),
        (&target.model, TABLE_MODEL_MESSAGE_TYPE),
    ] {
        let object = object_at(package, location)?;
        let _ = exact_message_role(object, message_type, budget)?;
    }

    let table_info = object_at(package, &target.table_info)?;
    let table_info_info = table_info
        .archive_info
        .message_infos
        .first()
        .ok_or(Error::InvalidSource)?;
    require_message_or_local_reference_field(
        table_info_info,
        target.model.identifier,
        &[TABLE_MODEL_FIELD],
    )?;

    // Every selected DataStore route is part of the physical graph's exact
    // authority. A lone aggregate MessageInfo edge is not enough: without a
    // field-local route there is no proof that the decoded route and the
    // persisted object graph name the same role.
    for route in &target.storage {
        validate_optional_model_route(
            package,
            target,
            route.location.identifier,
            route.kind.path(),
            budget,
        )?;
    }
    Ok(())
}

fn exact_message_role<'a>(
    object: &'a ArchiveObject,
    message_type: u32,
    budget: &mut Budget,
) -> Result<(usize, &'a [u8])> {
    if !has_exact_mutable_message_shape(object, &[message_type]) {
        return Err(Error::UnsupportedDependency);
    }
    unique_message(object, message_type, budget)
}

fn require_local_reference_field(
    info: &litchi_iwa_core::MessageInfo,
    identifier: u64,
    path: &[u32],
) -> Result<()> {
    if identifier == 0 {
        return Err(Error::InvalidSource);
    }
    let aggregate = info
        .object_references
        .iter()
        .filter(|candidate| **candidate == identifier)
        .count();
    if aggregate != 1 {
        return Err(Error::UnsupportedDependency);
    }
    let mut local = 0usize;
    for field in &info.field_infos {
        if field.object_references.contains(&identifier) {
            if field.object_references.as_slice() != [identifier]
                || !field.data_references.is_empty()
                || !is_message_reference_field(field)
                || field.path.as_slice() != path
            {
                return Err(Error::UnsupportedDependency);
            }
            local = local.checked_add(1).ok_or(Error::InvalidSource)?;
        }
    }
    if local == 1 {
        Ok(())
    } else {
        Err(Error::UnsupportedDependency)
    }
}

fn require_message_or_local_reference_field(
    info: &litchi_iwa_core::MessageInfo,
    identifier: u64,
    path: &[u32],
) -> Result<()> {
    if identifier == 0 {
        return Err(Error::InvalidSource);
    }
    if info
        .object_references
        .iter()
        .filter(|candidate| **candidate == identifier)
        .count()
        != 1
    {
        return Err(Error::UnsupportedDependency);
    }
    let matching = info
        .field_infos
        .iter()
        .filter(|field| field.object_references.contains(&identifier))
        .collect::<Vec<_>>();
    if matching.is_empty()
        || (matching.len() == 1
            && matching[0].object_references.as_slice() == [identifier]
            && matching[0].data_references.is_empty()
            && is_message_reference_field(matching[0])
            && matching[0].path.as_slice() == path)
    {
        Ok(())
    } else {
        Err(Error::UnsupportedDependency)
    }
}

/// Admit the complete native storage spine required by a physical row-sort
/// transaction.  The persisted sort-order owner intentionally does not call
/// this function: physical sorting has stricter row-affine requirements and
/// is kept behind this separate private boundary until its writer is ready.
pub(crate) fn admit_physical_table(
    package: &Package,
    target: &Target,
    budget: &mut Budget,
) -> Result<PhysicalTableTopology> {
    if target.locked {
        return Err(Error::UnsupportedDependency);
    }
    if target.rows == 0 || target.columns == 0 {
        return Err(Error::InvalidSource);
    }

    let payload = model_payload(package, target)?;
    validate_physical_root_roles(package, target, budget)?;
    let model_options = budget.storage_codec_options(package)?;
    let mut collector = PhysicalStorageCollector::new();
    let (model_and_store, report) = storage_codec::decode_table_model_with_data_store_and_visitor(
        payload,
        model_options,
        &mut collector,
    )
    .map_err(|_| Error::Codec)?;
    budget.storage_codec_report(report)?;
    let model = model_and_store.model();
    let store = model_and_store.data_store();
    if model.number_of_rows() != target.rows || model.number_of_columns() != target.columns {
        return Err(Error::InvalidSource);
    }
    let (header_rows, header_columns, footer_rows, uid_identifier, stroke_identifier, features) =
        parse_physical_model_fields(payload, package, target, budget)?;
    if header_rows
        .checked_add(footer_rows)
        .is_none_or(|body| body > target.rows)
        || header_columns > target.columns
    {
        return Err(Error::InvalidSource);
    }
    if features.any() {
        // Keep the flags on the topology type for future capability
        // expansion, but do not let a caller accidentally proceed with a
        // row-affine graph this owner cannot rewrite yet.
        return Err(Error::UnsupportedDependency);
    }
    let (tile_size, wide_rows) = decode_tile_storage(store.tiles(), package, budget)?;
    if tile_size == 0 {
        return Err(Error::InvalidSource);
    }
    let row_storage_options = budget.storage_codec_options(package)?;
    let (row_storage, row_storage_report) =
        storage_codec::decode_header_storage_with_report(store.row_headers(), row_storage_options)
            .map_err(|_| Error::Codec)?;
    budget.storage_codec_report(row_storage_report)?;
    if row_storage.bucket_hash_function() == 0 {
        return Err(Error::InvalidSource);
    }
    let tile_references = validate_tile_storage(
        package,
        target,
        &collector.tile_references,
        tile_size,
        budget,
    )?;
    let row_header_buckets = validate_row_headers(
        package,
        target,
        &collector.row_header_references,
        row_storage.bucket_hash_function(),
        budget,
    )?;

    let column_headers = target
        .storage_route(StorageRouteKind::ColumnBucket)
        .map(|route| route.location.clone())
        .ok_or(Error::InvalidSource)?;
    let string_table = target
        .storage_route(StorageRouteKind::StringTable)
        .map(|route| route.location.clone())
        .ok_or(Error::InvalidSource)?;
    let formula_table = target
        .storage_route(StorageRouteKind::FormulaTable)
        .map(|route| route.location.clone())
        .ok_or(Error::InvalidSource)?;
    let style_table = target
        .storage_route(StorageRouteKind::StyleTable)
        .map(|route| route.location.clone())
        .ok_or(Error::InvalidSource)?;
    let format_table = target
        .storage_route(StorageRouteKind::FormatTablePreBnc)
        .map(|route| route.location.clone())
        .ok_or(Error::InvalidSource)?;
    validate_column_headers(package, &column_headers, target.columns, budget)?;
    let _ = validate_data_list_state(package, &string_table, 1, budget)?;
    // Style and pre-BNC format lists are not rewritten by the physical
    // permutation, but they remain part of the selected storage spine. A
    // list segment or an unknown root/entry field there could carry
    // row-affine state that a byte-preserving untouched-object check would
    // otherwise never interpret.
    let _ = validate_data_list_state(package, &style_table, 4, budget)?;
    let has_formula_entries = validate_data_list_state(package, &formula_table, 3, budget)?;
    let _ = validate_data_list_state(package, &format_table, 2, budget)?;

    for (reference, path, list_type) in [
        (
            store.deprecated_custom_format_table(),
            &[MODEL_STORAGE_FIELD, 15][..],
            6,
        ),
        (store.format_table(), &[MODEL_STORAGE_FIELD, 22][..], 2),
    ] {
        if let Some(reference) = reference {
            validate_optional_model_route(package, target, reference.identifier(), path, budget)?;
            let location = locate_object(package, reference.identifier())?;
            ensure_unique_identity(package, reference.identifier(), budget)?;
            let _ = validate_data_list_state(package, &location, list_type, budget)?;
        }
    }

    let row_uid_identifier = uid_identifier.ok_or(Error::UnsupportedDependency)?;
    validate_optional_model_route(
        package,
        target,
        row_uid_identifier,
        &[MODEL_BASE_COLUMN_ROW_UIDS_FIELD],
        budget,
    )?;
    validate_exclusive_model_reference(
        package,
        target,
        row_uid_identifier,
        &[MODEL_BASE_COLUMN_ROW_UIDS_FIELD],
        budget,
    )?;
    let row_uid_map = locate_object(package, row_uid_identifier)?;
    ensure_unique_identity(package, row_uid_identifier, budget)?;
    validate_row_uid_map(package, &row_uid_map, target.rows, target.columns, budget)?;
    validate_mutable_root_aliases(
        package,
        target,
        &tile_references,
        &row_header_buckets,
        &column_headers,
        row_uid_identifier,
        budget,
    )?;

    let stroke_sidecar = if let Some(identifier) = stroke_identifier {
        let location = locate_object(package, identifier)?;
        ensure_unique_identity(package, identifier, budget)?;
        validate_empty_stroke_sidecar(package, &location, target, budget)?;
        Some(location)
    } else {
        None
    };

    let comment_storage = optional_storage_route(
        package,
        store.comment_storage_table(),
        &target.storage,
        budget,
    )?;
    let has_comment_entries = if let Some(location) = &comment_storage {
        validate_optional_model_route(
            package,
            target,
            location.identifier,
            MODEL_COMMENT_STORAGE_TABLE_PATH,
            budget,
        )?;
        validate_comment_storage(package, location, budget)?
    } else {
        false
    };
    let rich_text_table =
        optional_storage_route(package, store.rich_text_table(), &target.storage, budget)?;
    let has_rich_text_entries = if let Some(location) = &rich_text_table {
        validate_optional_model_route(
            package,
            target,
            location.identifier,
            MODEL_RICH_TEXT_TABLE_PATH,
            budget,
        )?;
        validate_data_list_state(package, location, 8, budget)?
    } else {
        false
    };

    let formula_error_table = optional_storage_route(
        package,
        store.formula_error_table(),
        &target.storage,
        budget,
    )?;
    let has_formula_error_entries = if let Some(location) = &formula_error_table {
        validate_optional_model_route(
            package,
            target,
            location.identifier,
            MODEL_FORMULA_ERROR_TABLE_PATH,
            budget,
        )?;
        validate_data_list_state(package, location, 5, budget)?
    } else {
        false
    };

    let conditional_style_table = optional_storage_route(
        package,
        store.conditional_style_table(),
        &target.storage,
        budget,
    )?;
    let has_conditional_style_entries = if let Some(location) = &conditional_style_table {
        validate_optional_model_route(
            package,
            target,
            location.identifier,
            MODEL_CONDITIONAL_STYLE_TABLE_PATH,
            budget,
        )?;
        validate_data_list_state(package, location, 9, budget)?
    } else {
        false
    };

    for (reference, path, list_type) in [
        (
            store.multiple_choice_list_format_table(),
            MODEL_MULTIPLE_CHOICE_TABLE_PATH,
            7,
        ),
        (
            store.import_warning_set_table(),
            MODEL_IMPORT_WARNING_TABLE_PATH,
            11,
        ),
        (
            store.control_cell_spec_table(),
            MODEL_CONTROL_CELL_SPEC_TABLE_PATH,
            12,
        ),
    ] {
        let location = optional_storage_route(package, reference, &target.storage, budget)?;
        if let Some(location) = &location {
            validate_optional_model_route(package, target, location.identifier, path, budget)?;
            if validate_data_list_state(package, location, list_type, budget)? {
                return Err(Error::UnsupportedDependency);
            }
        }
    }

    // A merge-region map is coordinate-affine even if a producer emits an
    // apparently empty root. The focused owner has no merge-map rewrite.
    if store.merge_region_map().is_some() {
        return Err(Error::UnsupportedDependency);
    }

    validate_formula_owner_dependencies(package, target, header_rows, footer_rows, budget)?;

    let mutable_component_count =
        mutable_component_count(&row_uid_map, &tile_references, &row_header_buckets, budget)?;

    Ok(PhysicalTableTopology {
        tile_size,
        wide_rows,
        tile_references,
        row_header_buckets,
        column_headers,
        string_table,
        row_uid_map,
        stroke_sidecar,
        mutable_component_count,
        has_formula_entries,
        has_formula_error_entries,
        has_conditional_style_entries,
        has_comment_entries,
        has_rich_text_entries,
        header_rows,
        header_columns,
        footer_rows,
        unsupported: features,
    })
}

fn mutable_component_count(
    row_uid_map: &ObjectLocation,
    tiles: &[PhysicalTileReference],
    headers: &[PhysicalRowHeaderBucket],
    budget: &mut Budget,
) -> Result<usize> {
    let capacity = 1usize
        .checked_add(tiles.len())
        .and_then(|value| value.checked_add(headers.len()))
        .ok_or(Error::InvalidSource)?;
    budget.allocations(capacity)?;
    budget.retained(
        capacity
            .checked_mul(size_of::<&str>())
            .ok_or(Error::InvalidSource)?,
    )?;
    let mut components = HashSet::new();
    components
        .try_reserve(capacity)
        .map_err(|_| Error::Allocation(capacity))?;
    components.insert(row_uid_map.component.as_ref());
    for tile in tiles {
        components.insert(tile.location.component.as_ref());
    }
    for header in headers {
        components.insert(header.location.component.as_ref());
    }
    Ok(components.len())
}

struct PhysicalStorageCollector {
    tile_references: Vec<(u32, u64)>,
    row_header_references: Vec<u64>,
}

impl PhysicalStorageCollector {
    fn new() -> Self {
        Self {
            tile_references: Vec::new(),
            row_header_references: Vec::new(),
        }
    }
}

impl storage_codec::StorageVisitor for PhysicalStorageCollector {
    fn visit_tile_reference(
        &mut self,
        record: storage_codec::TileReferenceRecord<'_>,
    ) -> std::result::Result<(), storage_codec::DecodeError> {
        self.tile_references
            .try_reserve(1)
            .map_err(|_| storage_codec::DecodeError::allocation(1))?;
        self.tile_references
            .push((record.tile_id(), record.reference().identifier()));
        Ok(())
    }

    fn visit_header_bucket(
        &mut self,
        reference: storage_codec::ReferenceRecord<'_>,
    ) -> std::result::Result<(), storage_codec::DecodeError> {
        self.row_header_references
            .try_reserve(1)
            .map_err(|_| storage_codec::DecodeError::allocation(1))?;
        self.row_header_references
            .push(reference.reference().identifier());
        Ok(())
    }
}

#[derive(Debug, Clone, Copy)]
struct PhysicalTileRow {
    index: u32,
    cell_count: u32,
}

struct PhysicalTileRowCollector {
    columns: u32,
    rows: Vec<PhysicalTileRow>,
}

impl PhysicalTileRowCollector {
    fn new(columns: u32) -> Self {
        Self {
            columns,
            rows: Vec::new(),
        }
    }
}

impl storage_codec::StorageVisitor for PhysicalTileRowCollector {
    fn visit_tile_row(
        &mut self,
        row: storage_codec::TileRowInfoSnapshot<'_>,
    ) -> std::result::Result<(), storage_codec::DecodeError> {
        let offsets = row
            .cell_offsets()
            .unwrap_or_else(|| row.cell_offsets_pre_bnc());
        if !offsets.len().is_multiple_of(2) {
            return Err(storage_codec::DecodeError::invalid_visitor_result());
        }
        let columns = usize::try_from(self.columns)
            .map_err(|_| storage_codec::DecodeError::invalid_visitor_result())?;
        let slots = offsets.len() / 2;
        if slots < columns {
            return Err(storage_codec::DecodeError::invalid_visitor_result());
        }
        let mut actual = 0u32;
        for (slot, bytes) in offsets.chunks_exact(2).enumerate() {
            let value = u16::from_le_bytes([bytes[0], bytes[1]]);
            if slot >= columns {
                if value != u16::MAX {
                    return Err(storage_codec::DecodeError::invalid_visitor_result());
                }
            } else if value != u16::MAX {
                actual = actual
                    .checked_add(1)
                    .ok_or_else(storage_codec::DecodeError::invalid_visitor_result)?;
            }
        }
        if actual != row.cell_count() {
            return Err(storage_codec::DecodeError::invalid_visitor_result());
        }
        self.rows
            .try_reserve(1)
            .map_err(|_| storage_codec::DecodeError::allocation(1))?;
        self.rows.push(PhysicalTileRow {
            index: row.tile_row_index(),
            cell_count: actual,
        });
        Ok(())
    }
}

struct PhysicalHeaderCollector {
    indices: Vec<u32>,
}

impl PhysicalHeaderCollector {
    fn new() -> Self {
        Self {
            indices: Vec::new(),
        }
    }
}

impl storage_codec::StorageVisitor for PhysicalHeaderCollector {
    fn visit_header(
        &mut self,
        header: storage_codec::HeaderSnapshot,
    ) -> std::result::Result<(), storage_codec::DecodeError> {
        self.indices
            .try_reserve(1)
            .map_err(|_| storage_codec::DecodeError::allocation(1))?;
        self.indices.push(header.index());
        Ok(())
    }
}

fn parse_physical_model_fields(
    payload: &[u8],
    package: &Package,
    target: &Target,
    budget: &mut Budget,
) -> Result<(
    u32,
    u32,
    u32,
    Option<u64>,
    Option<u64>,
    PhysicalUnsupportedFeatures,
)> {
    let limits = budget.residual(package)?;
    let fields = WireView::parse_with_limits(payload, limits).map_err(|_| Error::Wire)?;
    let mut seen = [false; 94];
    let mut header_rows = 0;
    let mut header_columns = 0;
    let mut footer_rows = 0;
    let mut row_uid_map = None;
    let mut stroke_sidecar = None;
    let mut features = PhysicalUnsupportedFeatures::default();

    for field in fields.fields() {
        budget.fields(1)?;
        field
            .validate_canonical_framing()
            .map_err(|_| Error::Wire)?;
        let number = field.number();
        if number < seen.len() as u32
            && !matches!(number, 90..=92)
            && std::mem::replace(&mut seen[number as usize], true)
        {
            return Err(Error::InvalidSource);
        }
        match number {
            1 | 8 => {
                std::str::from_utf8(field_bytes(field)?).map_err(|_| Error::InvalidSource)?;
            },
            3 | 18..=21 | 24..=27 | 30 | 36 | 48 => {
                let _ = strict_reference(field_bytes(field)?, limits, budget)?;
            },
            4 => {
                validate_physical_data_store_wire(field_bytes(field)?, package, budget)?;
            },
            5 | 23 | 43 => {
                // Providers, deprecated origin coordinates, and copied-table
                // provenance can carry dependencies outside the selected
                // physical table spine. The focused owner has no rewrite for
                // those graphs.
                return Err(Error::UnsupportedDependency);
            },
            6 | 7 | 28 => {
                let _ = canonical_field_u32(field)?;
            },
            MODEL_HEADER_ROWS_FIELD => {
                header_rows = canonical_field_u32(field)?;
            },
            MODEL_HEADER_COLUMNS_FIELD => {
                header_columns = canonical_field_u32(field)?;
            },
            MODEL_FOOTER_ROWS_FIELD => {
                footer_rows = canonical_field_u32(field)?;
            },
            12 | 13 | 22 | 29 | 31 | 32 | 37 | 50 | 51 => {
                let _ = canonical_field_bool(field)?;
            },
            14 | 15 | 40 | 41 | 42 => {
                if canonical_field_u32(field)? != 0 {
                    features.hidden = true;
                }
            },
            16 | 17 | 33 => {
                let _ = canonical_field_f64(field)?;
            },
            34 | 35 => {
                let identifier = strict_reference(field_bytes(field)?, limits, budget)?;
                validate_optional_model_route(package, target, identifier, &[number], budget)?;
                validate_dormant_hidden_formula_owner(package, identifier, budget)?;
            },
            38 => {
                let _ = strict_reference(field_bytes(field)?, limits, budget)?;
                features.filter = true;
            },
            39 => {
                // This inline CFUUID identifies the conditional-style
                // CalculationEngine owner. Its coordinate-affine dependency
                // graph is reached through UUID maps rather than ordinary
                // ArchiveInfo object references, so field presence remains
                // unsupported until that complete closure has a writer.
                validate_cfuuid(field_bytes(field)?, package, budget)?;
                return Err(Error::UnsupportedDependency);
            },
            44 => {
                if field_bytes(field)?.is_empty() {
                    return Err(Error::InvalidSource);
                }
            },
            45 => {
                // Sort-rule reference trackers can carry row-relative rule
                // state.  Keep this admission closed until their full graph
                // is owned by the physical transaction.
                return Err(Error::UnsupportedDependency);
            },
            MODEL_BASE_COLUMN_ROW_UIDS_FIELD => {
                let reference = strict_reference(field_bytes(field)?, limits, budget)?;
                if row_uid_map.replace(reference).is_some() {
                    return Err(Error::InvalidSource);
                }
            },
            47 => {
                validate_dormant_merge_owner(field_bytes(field)?, package, budget)?;
            },
            52 => {
                if field.wire_type() != 2 {
                    return Err(Error::InvalidSource);
                }
                // `StructuredTextImportRecord` carries an imported region,
                // source dimensions, and source bytes.  Those values are
                // row-affine even when the record happens to be empty, so a
                // physical row permutation must reject its mere presence.
                let _ = field_bytes(field)?;
                features.imported_data = true;
            },
            MODEL_STROKE_SIDECAR_FIELD => {
                let reference = strict_reference(field_bytes(field)?, limits, budget)?;
                // An empty stroke sidecar is retained as an opaque physical
                // root, but its model edge still has to be an exact
                // field-local role.  Accepting an aggregate-only edge here
                // would make a later row-carried BNC identifier impossible
                // to distinguish from an unrelated sidecar alias.
                validate_optional_model_route(
                    package,
                    target,
                    reference,
                    &[MODEL_STROKE_SIDECAR_FIELD],
                    budget,
                )?;
                if stroke_sidecar.replace(reference).is_some() {
                    return Err(Error::InvalidSource);
                }
            },
            60..=69 | 71..=80 | 87..=89 => {
                let _ = strict_reference(field_bytes(field)?, limits, budget)?;
            },
            70 => {
                validate_dormant_hidden_states_owner(field_bytes(field)?, package, target, budget)?;
            },
            81 => {
                validate_dormant_deprecated_category_owner(field_bytes(field)?, package, budget)?;
            },
            82 => {
                validate_dormant_pencil_owner(field_bytes(field)?, package, budget)?;
            },
            83 => {
                if !field_bytes(field)?.is_empty() {
                    features.grouping = true;
                }
            },
            84 => {
                // Hidden-state owners are coordinate-affine and are not part
                // of the native physical-sort rewrite closure.
                return Err(Error::UnsupportedDependency);
            },
            85 => {
                let _ = strict_reference(field_bytes(field)?, limits, budget)?;
                features.pivot = true;
            },
            86 => {
                let identifier = strict_reference(field_bytes(field)?, limits, budget)?;
                validate_optional_model_route(package, target, identifier, &[number], budget)?;
                validate_dormant_category_owner_reference(package, identifier, budget)?;
            },
            90..=92 => {
                let _ = canonical_field_u32(field)?;
                features.pivot = true;
            },
            93 => {
                // Spill owners encode a row-affine dependency graph.  Do not
                // infer safety from a UUID-shaped payload alone.
                return Err(Error::UnsupportedDependency);
            },
            _ => return Err(Error::UnsupportedDependency),
        }
    }
    budget.work(payload.len())?;
    Ok((
        header_rows,
        header_columns,
        footer_rows,
        row_uid_map,
        stroke_sidecar,
        features,
    ))
}

fn validated_wire_view<'a>(
    payload: &'a [u8],
    package: &Package,
    budget: &mut Budget,
) -> Result<WireView<'a>> {
    let limits = budget.residual(package)?;
    let fields = WireView::parse_with_limits(payload, limits).map_err(|_| Error::Wire)?;
    budget.fields(fields.fields().count())?;
    budget.work(payload.len())?;
    for field in fields.fields() {
        field
            .validate_canonical_framing()
            .map_err(|_| Error::Wire)?;
    }
    Ok(fields)
}

/// Mark one singular field while validating a physical storage envelope.
///
/// The generated storage projections intentionally skip unknown fields so
/// they can remain forward-compatible.  Physical sorting has a narrower
/// contract: an unknown field in any mutable storage envelope could carry
/// row-affine state that the row permutation does not move.  These helpers
/// therefore reject unknown keys and duplicate singular keys before the
/// strict Buffa/handwritten codec is entered.
fn mark_physical_singular<const N: usize>(
    seen: &mut [bool; N],
    field: litchi_iwa_common::wire::WireFieldView<'_>,
) -> Result<()> {
    let index = usize::try_from(field.number()).map_err(|_| Error::InvalidSource)?;
    let slot = seen.get_mut(index).ok_or(Error::UnsupportedDependency)?;
    if std::mem::replace(slot, true) {
        return Err(Error::InvalidSource);
    }
    Ok(())
}

fn require_physical_field(
    field: litchi_iwa_common::wire::WireFieldView<'_>,
    allowed: &[u32],
) -> Result<()> {
    if allowed.contains(&field.number()) {
        Ok(())
    } else {
        Err(Error::UnsupportedDependency)
    }
}

fn physical_varint_u32(field: litchi_iwa_common::wire::WireFieldView<'_>) -> Result<u32> {
    canonical_field_u32(field)
}

fn physical_varint_bool(field: litchi_iwa_common::wire::WireFieldView<'_>) -> Result<bool> {
    canonical_field_bool(field)
}

fn physical_bytes<'a>(field: litchi_iwa_common::wire::WireFieldView<'a>) -> Result<&'a [u8]> {
    field_bytes(field)
}

fn physical_reference(payload: &[u8], package: &Package, budget: &mut Budget) -> Result<u64> {
    let limits = budget.residual(package)?;
    strict_reference(payload, limits, budget)
}

fn validate_physical_data_store_wire(
    payload: &[u8],
    package: &Package,
    budget: &mut Budget,
) -> Result<()> {
    let fields = validated_wire_view(payload, package, budget)?;
    let mut seen = [false; 23];
    for field in fields.fields() {
        if !(1..=22).contains(&field.number()) {
            return Err(Error::UnsupportedDependency);
        }
        match field.number() {
            1 => {
                mark_physical_singular(&mut seen, field)?;
                validate_physical_header_storage_wire(physical_bytes(field)?, package, budget)?;
            },
            2 | 4..=6 | 11 | 12 | 13 | 15..=22 => {
                mark_physical_singular(&mut seen, field)?;
                let _ = physical_reference(physical_bytes(field)?, package, budget)?;
            },
            3 => {
                mark_physical_singular(&mut seen, field)?;
                validate_physical_tile_storage_wire(physical_bytes(field)?, package, budget)?;
            },
            7 | 8 | 14 => {
                mark_physical_singular(&mut seen, field)?;
                let _ = physical_varint_u32(field)?;
            },
            9 | 10 => {
                mark_physical_singular(&mut seen, field)?;
                validate_physical_table_tree_wire(physical_bytes(field)?, package, budget)?;
            },
            _ => return Err(Error::UnsupportedDependency),
        }
    }
    Ok(())
}

fn validate_model_storage_envelope_wire(
    payload: &[u8],
    package: &Package,
    budget: &mut Budget,
) -> Result<()> {
    let fields = validated_wire_view(payload, package, budget)?;
    let mut found = false;
    for field in fields.fields() {
        if field.number() != MODEL_STORAGE_FIELD {
            continue;
        }
        if std::mem::replace(&mut found, true) {
            return Err(Error::InvalidSource);
        }
        validate_physical_data_store_wire(physical_bytes(field)?, package, budget)?;
    }
    if found {
        Ok(())
    } else {
        Err(Error::InvalidSource)
    }
}

fn validate_physical_table_tree_wire(
    payload: &[u8],
    package: &Package,
    budget: &mut Budget,
) -> Result<()> {
    let fields = validated_wire_view(payload, package, budget)?;
    for field in fields.fields() {
        require_physical_field(field, &[1])?;
        if field.wire_type() != 2 {
            return Err(Error::InvalidSource);
        }
        let node = validated_wire_view(field.payload(), package, budget)?;
        let mut seen = [false; 3];
        for node_field in node.fields() {
            require_physical_field(node_field, &[1, 2])?;
            mark_physical_singular(&mut seen, node_field)?;
            let _ = physical_varint_u32(node_field)?;
        }
    }
    Ok(())
}

fn validate_physical_tile_storage_wire(
    payload: &[u8],
    package: &Package,
    budget: &mut Budget,
) -> Result<()> {
    let fields = validated_wire_view(payload, package, budget)?;
    let mut seen = [false; 23];
    for field in fields.fields() {
        require_physical_field(field, &[1, 2, 3])?;
        match field.number() {
            1 => {
                if field.wire_type() != 2 {
                    return Err(Error::InvalidSource);
                }
                let record = validated_wire_view(field.payload(), package, budget)?;
                let mut record_seen = [false; 3];
                for record_field in record.fields() {
                    require_physical_field(record_field, &[1, 2])?;
                    mark_physical_singular(&mut record_seen, record_field)?;
                    match record_field.number() {
                        1 => {
                            let _ = physical_varint_u32(record_field)?;
                        },
                        2 => {
                            let _ =
                                physical_reference(physical_bytes(record_field)?, package, budget)?;
                        },
                        _ => return Err(Error::UnsupportedDependency),
                    }
                }
            },
            2 => {
                mark_physical_singular(&mut seen, field)?;
                let _ = physical_varint_u32(field)?;
            },
            3 => {
                mark_physical_singular(&mut seen, field)?;
                let _ = physical_varint_bool(field)?;
            },
            _ => return Err(Error::UnsupportedDependency),
        }
    }
    Ok(())
}

fn validate_physical_tile_wire(
    payload: &[u8],
    package: &Package,
    budget: &mut Budget,
) -> Result<()> {
    let fields = validated_wire_view(payload, package, budget)?;
    let mut seen = [false; 23];
    for field in fields.fields() {
        require_physical_field(field, &[1, 2, 3, 4, 5, 6, 7, 8])?;
        match field.number() {
            1..=4 | 6 => {
                mark_physical_singular(&mut seen, field)?;
                let _ = physical_varint_u32(field)?;
            },
            5 => {
                if field.wire_type() != 2 {
                    return Err(Error::InvalidSource);
                }
                validate_physical_tile_row_wire(field.payload(), package, budget)?;
            },
            7 | 8 => {
                mark_physical_singular(&mut seen, field)?;
                let _ = physical_varint_bool(field)?;
            },
            _ => return Err(Error::UnsupportedDependency),
        }
    }
    Ok(())
}

fn validate_physical_tile_row_wire(
    payload: &[u8],
    package: &Package,
    budget: &mut Budget,
) -> Result<()> {
    let fields = validated_wire_view(payload, package, budget)?;
    let mut seen = [false; 23];
    for field in fields.fields() {
        require_physical_field(field, &[1, 2, 3, 4, 5, 6, 7, 8])?;
        match field.number() {
            1 | 2 | 5 => {
                mark_physical_singular(&mut seen, field)?;
                let _ = physical_varint_u32(field)?;
            },
            3 | 4 | 6 | 7 => {
                mark_physical_singular(&mut seen, field)?;
                let _ = physical_bytes(field)?;
            },
            8 => {
                mark_physical_singular(&mut seen, field)?;
                let _ = physical_varint_bool(field)?;
            },
            _ => return Err(Error::UnsupportedDependency),
        }
    }
    Ok(())
}

fn validate_physical_header_storage_wire(
    payload: &[u8],
    package: &Package,
    budget: &mut Budget,
) -> Result<()> {
    let fields = validated_wire_view(payload, package, budget)?;
    let mut seen = [false; 23];
    for field in fields.fields() {
        require_physical_field(field, &[1, 2])?;
        match field.number() {
            1 => {
                mark_physical_singular(&mut seen, field)?;
                let _ = physical_varint_u32(field)?;
            },
            2 => {
                let reference = physical_bytes(field)?;
                let _ = physical_reference(reference, package, budget)?;
            },
            _ => return Err(Error::UnsupportedDependency),
        }
    }
    Ok(())
}

fn validate_physical_header_bucket_wire(
    payload: &[u8],
    package: &Package,
    budget: &mut Budget,
) -> Result<()> {
    let fields = validated_wire_view(payload, package, budget)?;
    let mut seen = [false; 23];
    for field in fields.fields() {
        require_physical_field(field, &[1, 2])?;
        match field.number() {
            1 => {
                mark_physical_singular(&mut seen, field)?;
                let _ = physical_varint_u32(field)?;
            },
            2 => {
                if field.wire_type() != 2 {
                    return Err(Error::InvalidSource);
                }
                validate_physical_header_wire(field.payload(), package, budget)?;
            },
            _ => return Err(Error::UnsupportedDependency),
        }
    }
    Ok(())
}

fn validate_physical_header_wire(
    payload: &[u8],
    package: &Package,
    budget: &mut Budget,
) -> Result<()> {
    let fields = validated_wire_view(payload, package, budget)?;
    let mut seen = [false; 7];
    for field in fields.fields() {
        require_physical_field(field, &[1, 2, 3, 4, 5, 6])?;
        let index = usize::try_from(field.number()).map_err(|_| Error::InvalidSource)?;
        let slot = seen.get_mut(index).ok_or(Error::UnsupportedDependency)?;
        if std::mem::replace(slot, true) {
            return Err(Error::InvalidSource);
        }
        match field.number() {
            1 | 3 | 4 => {
                let _ = physical_varint_u32(field)?;
            },
            2 => {
                if field.wire_type() != 5 || field.payload().len() != 4 {
                    return Err(Error::InvalidSource);
                }
            },
            5 | 6 => {
                let _ = physical_reference(physical_bytes(field)?, package, budget)?;
            },
            _ => return Err(Error::UnsupportedDependency),
        }
    }
    Ok(())
}

fn validate_physical_uuid_wire(
    payload: &[u8],
    package: &Package,
    budget: &mut Budget,
) -> Result<()> {
    let fields = validated_wire_view(payload, package, budget)?;
    let mut seen = [false; 3];
    for field in fields.fields() {
        require_physical_field(field, &[1, 2])?;
        let index = usize::try_from(field.number()).map_err(|_| Error::InvalidSource)?;
        let slot = seen.get_mut(index).ok_or(Error::UnsupportedDependency)?;
        if std::mem::replace(slot, true) {
            return Err(Error::InvalidSource);
        }
        let _ = canonical_field_u64(field)?;
    }
    Ok(())
}

fn validate_physical_repeated_u32(field: litchi_iwa_common::wire::WireFieldView<'_>) -> Result<()> {
    match field.wire_type() {
        0 => {
            let _ = canonical_field_u32(field)?;
        },
        2 => {
            let mut payload = field.payload();
            while !payload.is_empty() {
                let (value, width) =
                    decode_varint_from_bytes(payload).map_err(|_| Error::InvalidSource)?;
                if width != encoded_len(value) || u32::try_from(value).is_err() {
                    return Err(Error::InvalidSource);
                }
                payload = &payload[width..];
            }
        },
        _ => return Err(Error::InvalidSource),
    }
    Ok(())
}

fn validate_physical_uid_map_wire(
    payload: &[u8],
    package: &Package,
    budget: &mut Budget,
) -> Result<()> {
    let fields = validated_wire_view(payload, package, budget)?;
    for field in fields.fields() {
        require_physical_field(field, &[1, 2, 3, 4, 5, 6])?;
        match field.number() {
            1 | 4 => validate_physical_uuid_wire(physical_bytes(field)?, package, budget)?,
            2 | 3 | 5 | 6 => validate_physical_repeated_u32(field)?,
            _ => return Err(Error::UnsupportedDependency),
        }
    }
    Ok(())
}

fn validate_physical_data_list_entry_wire(
    payload: &[u8],
    package: &Package,
    budget: &mut Budget,
) -> Result<()> {
    let fields = validated_wire_view(payload, package, budget)?;
    let mut seen = [false; 13];
    for field in fields.fields() {
        require_physical_field(field, &[1, 2, 3, 4, 5, 6, 8, 9, 10, 11, 12])?;
        match field.number() {
            1 | 2 => {
                mark_physical_singular(&mut seen, field)?;
                let _ = physical_varint_u32(field)?;
            },
            3 => {
                mark_physical_singular(&mut seen, field)?;
                let value = std::str::from_utf8(physical_bytes(field)?)
                    .map_err(|_| Error::InvalidSource)?;
                let _ = value;
            },
            4 | 9 | 10 => {
                mark_physical_singular(&mut seen, field)?;
                let _ = physical_reference(physical_bytes(field)?, package, budget)?;
            },
            5 | 6 | 8 | 11 | 12 => {
                mark_physical_singular(&mut seen, field)?;
                let _ = physical_bytes(field)?;
            },
            _ => return Err(Error::UnsupportedDependency),
        }
    }
    Ok(())
}

fn validate_physical_data_list_wire(
    payload: &[u8],
    package: &Package,
    budget: &mut Budget,
) -> Result<()> {
    let fields = validated_wire_view(payload, package, budget)?;
    let mut seen = [false; 6];
    for field in fields.fields() {
        require_physical_field(field, &[1, 2, 3, 4, 5])?;
        match field.number() {
            1 | 2 => {
                mark_physical_singular(&mut seen, field)?;
                let _ = physical_varint_u32(field)?;
            },
            3 => {
                if field.wire_type() != 2 {
                    return Err(Error::InvalidSource);
                }
                validate_physical_data_list_entry_wire(field.payload(), package, budget)?;
            },
            4 => {
                let _ = physical_reference(physical_bytes(field)?, package, budget)?;
            },
            5 => {
                mark_physical_singular(&mut seen, field)?;
                let _ = physical_varint_bool(field)?;
            },
            _ => return Err(Error::UnsupportedDependency),
        }
    }
    Ok(())
}

fn canonical_field_u64(field: litchi_iwa_common::wire::WireFieldView<'_>) -> Result<u64> {
    if field.wire_type() != 0 {
        return Err(Error::InvalidSource);
    }
    let (value, width) =
        decode_varint_from_bytes(field.payload()).map_err(|_| Error::InvalidSource)?;
    if width != encoded_len(value) {
        return Err(Error::InvalidSource);
    }
    Ok(value)
}

fn canonical_field_bool(field: litchi_iwa_common::wire::WireFieldView<'_>) -> Result<bool> {
    match canonical_field_u64(field)? {
        0 => Ok(false),
        1 => Ok(true),
        _ => Err(Error::InvalidSource),
    }
}

fn canonical_field_f64(field: litchi_iwa_common::wire::WireFieldView<'_>) -> Result<f64> {
    if field.wire_type() != 1 || field.payload().len() != size_of::<f64>() {
        return Err(Error::InvalidSource);
    }
    let bytes: [u8; size_of::<f64>()] = field
        .payload()
        .try_into()
        .map_err(|_| Error::InvalidSource)?;
    let value = f64::from_le_bytes(bytes);
    value
        .is_finite()
        .then_some(value)
        .ok_or(Error::InvalidSource)
}

#[allow(
    dead_code,
    reason = "Field 45 is fail-closed until its dependency closure is owned."
)]
/// Accept either one canonical 16-byte CFUUID or all four canonical word
/// fields, but never a mixed, partial, duplicate, or all-zero identity.
fn validate_cfuuid(payload: &[u8], package: &Package, budget: &mut Budget) -> Result<()> {
    let fields = validated_wire_view(payload, package, budget)?;
    let mut bytes = None;
    let mut words = [None; 4];
    for field in fields.fields() {
        match field.number() {
            1 => {
                if field.wire_type() != 2 || bytes.replace(field.payload()).is_some() {
                    return Err(Error::InvalidSource);
                }
            },
            2..=5 => {
                let index =
                    usize::try_from(field.number() - 2).map_err(|_| Error::InvalidSource)?;
                if words[index].replace(canonical_field_u32(field)?).is_some() {
                    return Err(Error::InvalidSource);
                }
            },
            _ => return Err(Error::InvalidSource),
        }
    }
    match bytes {
        Some(value) => {
            if words.iter().any(Option::is_some)
                || value.len() != 16
                || value.iter().all(|byte| *byte == 0)
            {
                return Err(Error::InvalidSource);
            }
        },
        None => {
            if words.iter().any(Option::is_none) || words.iter().all(|value| value == &Some(0)) {
                return Err(Error::InvalidSource);
            }
        },
    }
    Ok(())
}

#[allow(
    dead_code,
    reason = "Fields 84 and 93 are fail-closed until their dependency closures are owned."
)]
fn validate_dormant_formula_store(
    payload: &[u8],
    package: &Package,
    budget: &mut Budget,
) -> Result<()> {
    let fields = validated_wire_view(payload, package, budget)?;
    let mut next_index = None;
    for field in fields.fields() {
        match field.number() {
            2 if next_index.is_none() => next_index = Some(canonical_field_u32(field)?),
            3 if field.wire_type() == 2 => return Err(Error::UnsupportedDependency),
            _ => return Err(Error::InvalidSource),
        }
    }
    if next_index == Some(0) {
        Ok(())
    } else {
        Err(Error::UnsupportedDependency)
    }
}

fn validate_dormant_merge_owner(
    payload: &[u8],
    package: &Package,
    budget: &mut Budget,
) -> Result<()> {
    validate_dormant_uuid_formula_owner(payload, package, budget, false)
}

fn validate_dormant_pencil_owner(
    payload: &[u8],
    package: &Package,
    budget: &mut Budget,
) -> Result<()> {
    validate_dormant_uuid_formula_owner(payload, package, budget, true)
}

fn validate_dormant_uuid_formula_owner(
    payload: &[u8],
    package: &Package,
    budget: &mut Budget,
    reject_annotations: bool,
) -> Result<()> {
    let fields = validated_wire_view(payload, package, budget)?;
    let mut owner_seen = false;
    let mut store_seen = false;
    for field in fields.fields() {
        match field.number() {
            1 if !owner_seen && field.wire_type() == 2 => {
                owner_seen = true;
                validate_cfuuid(field.payload(), package, budget)?;
            },
            2 if !store_seen && field.wire_type() == 2 => {
                store_seen = true;
                validate_dormant_formula_store(field.payload(), package, budget)?;
            },
            3 if reject_annotations && field.wire_type() == 2 => {
                return Err(Error::UnsupportedDependency);
            },
            _ => return Err(Error::InvalidSource),
        }
    }
    owner_seen.then_some(()).ok_or(Error::InvalidSource)
}

fn strict_single_message<'a>(
    package: &'a Package,
    identifier: u64,
    message_type: u32,
    budget: &mut Budget,
) -> Result<(&'a ArchiveObject, usize, &'a [u8])> {
    let location = locate_object(package, identifier)?;
    ensure_unique_identity(package, identifier, budget)?;
    let object = object_at(package, &location)?;
    if object.messages.len() != 1 || object.archive_info.message_infos.len() != 1 {
        return Err(Error::UnsupportedDependency);
    }
    let (index, payload) = unique_message(object, message_type, budget)?;
    Ok((object, index, payload))
}

fn validate_declared_references(
    object: &ArchiveObject,
    message_index: usize,
    expected: &[u64],
    accepted_paths: &[&[u32]],
    budget: &mut Budget,
) -> Result<()> {
    let info = object
        .archive_info
        .message_infos
        .get(message_index)
        .ok_or(Error::InvalidSource)?;
    budget.references(
        info.object_references
            .len()
            .checked_add(info.data_references.len())
            .and_then(|count| {
                info.field_infos.iter().try_fold(count, |total, field| {
                    total
                        .checked_add(field.object_references.len())
                        .and_then(|value| value.checked_add(field.data_references.len()))
                })
            })
            .ok_or(Error::InvalidSource)?,
    )?;
    if !info.data_references.is_empty()
        || info.object_references.len() != expected.len()
        || expected.iter().any(|identifier| {
            info.object_references
                .iter()
                .filter(|candidate| *candidate == identifier)
                .count()
                != 1
        })
    {
        return Err(Error::UnsupportedDependency);
    }
    for field in &info.field_infos {
        if !field.data_references.is_empty() {
            return Err(Error::UnsupportedDependency);
        }
        if field.object_references.is_empty() {
            continue;
        }
        if !field
            .r#type
            .is_none_or(|kind| matches!(kind, FieldType::ObjectReference | FieldType::Message))
            || !accepted_paths
                .iter()
                .any(|path| field.path.as_slice() == *path)
            || field.object_references.iter().any(|identifier| {
                expected
                    .iter()
                    .filter(|candidate| *candidate == identifier)
                    .count()
                    != 1
            })
        {
            return Err(Error::UnsupportedDependency);
        }
    }
    Ok(())
}

fn validate_selected_model_references(
    package: &Package,
    target: &Target,
    identifiers: &[u64],
    accepted_path: &[u32],
    budget: &mut Budget,
) -> Result<()> {
    let model = object_at(package, &target.model)?;
    let info = model
        .archive_info
        .message_infos
        .get(target.model_message_index)
        .ok_or(Error::InvalidSource)?;
    budget.work(
        identifiers
            .len()
            .checked_add(info.field_infos.len())
            .ok_or(Error::InvalidSource)?,
    )?;
    budget.references(
        info.object_references
            .len()
            .checked_add(info.data_references.len())
            .and_then(|count| {
                info.field_infos.iter().try_fold(count, |total, field| {
                    total
                        .checked_add(field.object_references.len())
                        .and_then(|value| value.checked_add(field.data_references.len()))
                })
            })
            .ok_or(Error::InvalidSource)?,
    )?;
    for identifier in identifiers {
        if info
            .object_references
            .iter()
            .filter(|candidate| *candidate == identifier)
            .count()
            != 1
        {
            return Err(Error::UnsupportedDependency);
        }
    }
    for field in &info.field_infos {
        let contains_selected = field
            .object_references
            .iter()
            .any(|identifier| identifiers.contains(identifier));
        if contains_selected
            && (!field.data_references.is_empty()
                || !field.r#type.is_none_or(|kind| {
                    matches!(kind, FieldType::ObjectReference | FieldType::Message)
                })
                || field.path.as_slice() != accepted_path
                || field
                    .object_references
                    .iter()
                    .any(|identifier| !identifiers.contains(identifier)))
        {
            return Err(Error::UnsupportedDependency);
        }
    }
    Ok(())
}

fn validate_dormant_hidden_formula_owner(
    package: &Package,
    identifier: u64,
    budget: &mut Budget,
) -> Result<()> {
    let (object, index, payload) = strict_single_message(
        package,
        identifier,
        HIDDEN_STATE_FORMULA_OWNER_MESSAGE_TYPE,
        budget,
    )?;
    validate_declared_references(object, index, &[], &[], budget)?;
    let fields = validated_wire_view(payload, package, budget)?;
    let mut owner_seen = false;
    let mut import_seen = false;
    for field in fields.fields() {
        match field.number() {
            1 if !owner_seen && field.wire_type() == 2 => {
                owner_seen = true;
                validate_cfuuid(field.payload(), package, budget)?;
            },
            2 if field.wire_type() == 2 => return Err(Error::UnsupportedDependency),
            3 if !import_seen => {
                import_seen = true;
                if canonical_field_bool(field)? {
                    return Err(Error::UnsupportedDependency);
                }
            },
            _ => return Err(Error::InvalidSource),
        }
    }
    owner_seen.then_some(()).ok_or(Error::InvalidSource)
}

fn validate_dormant_hidden_states_owner(
    payload: &[u8],
    package: &Package,
    target: &Target,
    budget: &mut Budget,
) -> Result<()> {
    let fields = validated_wire_view(payload, package, budget)?;
    let limits = budget.residual(package)?;
    let mut owner = None;
    let mut state = None;
    for field in fields.fields() {
        match field.number() {
            1 if owner.is_none() && field.wire_type() == 2 => owner = Some(field.payload()),
            2 if state.is_none() && field.wire_type() == 2 => state = Some(field.payload()),
            _ => return Err(Error::UnsupportedDependency),
        }
    }
    let owner_uid = parse_uuid(owner.ok_or(Error::InvalidSource)?, limits, budget)?;
    let filter_ids = validate_dormant_hidden_state(
        state.ok_or(Error::InvalidSource)?,
        owner_uid,
        package,
        budget,
    )?;
    if filter_ids[0] == filter_ids[1] {
        return Err(Error::UnsupportedDependency);
    }
    validate_selected_model_references(package, target, &filter_ids, &[70], budget)?;
    for identifier in filter_ids {
        validate_dormant_filter_set(package, identifier, budget)?;
    }
    Ok(())
}

fn validate_dormant_hidden_state(
    payload: &[u8],
    owner_uid: (u64, u64),
    package: &Package,
    budget: &mut Budget,
) -> Result<[u64; 2]> {
    let fields = validated_wire_view(payload, package, budget)?;
    let limits = budget.residual(package)?;
    let mut uid = None;
    let mut column = None;
    let mut row = None;
    for field in fields.fields() {
        match field.number() {
            1 if uid.is_none() && field.wire_type() == 2 => uid = Some(field.payload()),
            2 if column.is_none() && field.wire_type() == 2 => column = Some(field.payload()),
            3 if row.is_none() && field.wire_type() == 2 => row = Some(field.payload()),
            _ => return Err(Error::UnsupportedDependency),
        }
    }
    if parse_uuid(uid.ok_or(Error::InvalidSource)?, limits, budget)? != owner_uid {
        return Err(Error::UnsupportedDependency);
    }
    Ok([
        validate_dormant_hidden_extent(column.ok_or(Error::InvalidSource)?, 0, package, budget)?,
        validate_dormant_hidden_extent(row.ok_or(Error::InvalidSource)?, 1, package, budget)?,
    ])
}

fn validate_dormant_hidden_extent(
    payload: &[u8],
    expected_direction: u32,
    package: &Package,
    budget: &mut Budget,
) -> Result<u64> {
    let fields = validated_wire_view(payload, package, budget)?;
    let limits = budget.residual(package)?;
    let mut uid = None;
    let mut direction = None;
    let mut import_seen = false;
    let mut filter = None;
    for field in fields.fields() {
        match field.number() {
            1 if uid.is_none() && field.wire_type() == 2 => uid = Some(field.payload()),
            3 if direction.is_none() => direction = Some(canonical_field_u32(field)?),
            6 if !import_seen => {
                import_seen = true;
                if canonical_field_bool(field)? {
                    return Err(Error::UnsupportedDependency);
                }
            },
            8 if filter.is_none() && field.wire_type() == 2 => {
                filter = Some(strict_reference(field.payload(), limits, budget)?);
            },
            2 | 5 | 7 | 9..=12 => return Err(Error::UnsupportedDependency),
            _ => return Err(Error::InvalidSource),
        }
    }
    let _ = parse_uuid(uid.ok_or(Error::InvalidSource)?, limits, budget)?;
    if direction != Some(expected_direction) {
        return Err(Error::UnsupportedDependency);
    }
    filter.ok_or(Error::InvalidSource)
}

fn validate_dormant_filter_set(
    package: &Package,
    identifier: u64,
    budget: &mut Budget,
) -> Result<()> {
    let (object, index, payload) =
        strict_single_message(package, identifier, FILTER_SET_MESSAGE_TYPE, budget)?;
    validate_declared_references(object, index, &[], &[], budget)?;
    let fields = validated_wire_view(payload, package, budget)?;
    let mut kind = None;
    let mut enabled = None;
    let mut import = None;
    let mut offset_count = 0usize;
    for field in fields.fields() {
        match field.number() {
            1 if kind.is_none() => kind = Some(canonical_field_u32(field)?),
            2 if enabled.is_none() => enabled = Some(canonical_field_bool(field)?),
            4 if import.is_none() => import = Some(canonical_field_bool(field)?),
            5 => {
                if field.wire_type() != 0 || canonical_field_u32(field)? != 0 {
                    return Err(Error::UnsupportedDependency);
                }
                offset_count = offset_count.checked_add(1).ok_or(Error::InvalidSource)?;
            },
            3 | 6 | 7 => return Err(Error::UnsupportedDependency),
            _ => return Err(Error::InvalidSource),
        }
    }
    if kind == Some(0) && enabled == Some(false) && import == Some(false) && offset_count == 1 {
        Ok(())
    } else {
        Err(Error::UnsupportedDependency)
    }
}

fn validate_dormant_deprecated_category_owner(
    payload: &[u8],
    package: &Package,
    budget: &mut Budget,
) -> Result<()> {
    let fields = validated_wire_view(payload, package, budget)?;
    let mut owner = None;
    let mut group_count = 0usize;
    for field in fields.fields() {
        match field.number() {
            1 if owner.is_none() && field.wire_type() == 2 => owner = Some(field.payload()),
            2 if field.wire_type() == 2 => {
                group_count = group_count.checked_add(1).ok_or(Error::InvalidSource)?;
                if validate_dormant_group_by(field.payload(), false, package, budget)?.is_some() {
                    return Err(Error::UnsupportedDependency);
                }
            },
            _ => return Err(Error::InvalidSource),
        }
    }
    let owner = parse_uuid_allow_zero(owner.ok_or(Error::InvalidSource)?, package, budget)?;
    if owner == (0, 0) && group_count == 1 {
        Ok(())
    } else {
        Err(Error::UnsupportedDependency)
    }
}

fn parse_uuid_allow_zero(
    payload: &[u8],
    package: &Package,
    budget: &mut Budget,
) -> Result<(u64, u64)> {
    let fields = validated_wire_view(payload, package, budget)?;
    let mut lower = None;
    let mut upper = None;
    for field in fields.fields() {
        match field.number() {
            1 if lower.is_none() => lower = Some(canonical_field_u64(field)?),
            2 if upper.is_none() => upper = Some(canonical_field_u64(field)?),
            _ => return Err(Error::InvalidSource),
        }
    }
    Ok((
        lower.ok_or(Error::InvalidSource)?,
        upper.ok_or(Error::InvalidSource)?,
    ))
}

fn validate_dormant_category_owner_reference(
    package: &Package,
    identifier: u64,
    budget: &mut Budget,
) -> Result<()> {
    let limits = budget.residual(package)?;
    let (owner_object, owner_index, owner_payload) = strict_single_message(
        package,
        identifier,
        CATEGORY_OWNER_REFERENCE_MESSAGE_TYPE,
        budget,
    )?;
    let owner_fields = validated_wire_view(owner_payload, package, budget)?;
    let mut group_identifier = None;
    for field in owner_fields.fields() {
        if field.number() != 1 || field.wire_type() != 2 || group_identifier.is_some() {
            return Err(Error::UnsupportedDependency);
        }
        group_identifier = Some(strict_reference(field.payload(), limits, budget)?);
    }
    let group_identifier = group_identifier.ok_or(Error::InvalidSource)?;
    validate_declared_references(
        owner_object,
        owner_index,
        &[group_identifier],
        &[&[1]],
        budget,
    )?;

    let (group_object, group_index, group_payload) =
        strict_single_message(package, group_identifier, GROUP_BY_MESSAGE_TYPE, budget)?;
    let root_identifier = validate_dormant_group_by(group_payload, true, package, budget)?
        .ok_or(Error::InvalidSource)?;
    validate_declared_references(
        group_object,
        group_index,
        &[root_identifier],
        &[&[17], &[18]],
        budget,
    )?;
    let (root_object, root_index, root_payload) =
        strict_single_message(package, root_identifier, GROUP_NODE_MESSAGE_TYPE, budget)?;
    validate_declared_references(root_object, root_index, &[], &[], budget)?;
    validate_inert_group_node(root_payload, package, budget)
}

fn validate_dormant_group_by(
    payload: &[u8],
    require_root_reference: bool,
    package: &Package,
    budget: &mut Budget,
) -> Result<Option<u64>> {
    let fields = validated_wire_view(payload, package, budget)?;
    let limits = budget.residual(package)?;
    let mut uid = None;
    let mut root = None;
    let mut enabled = None;
    let mut owner_index = None;
    let mut coordinates = [false; 8];
    let mut embedded_root = false;
    for field in fields.fields() {
        match field.number() {
            1 if uid.is_none() && field.wire_type() == 2 => uid = Some(field.payload()),
            3 if !embedded_root && field.wire_type() == 2 => {
                embedded_root = true;
                validate_inert_group_node(field.payload(), package, budget)?;
            },
            6 if enabled.is_none() => enabled = Some(canonical_field_bool(field)?),
            7..=13 | 16 if field.wire_type() == 2 => {
                let (slot, expected_column) = match field.number() {
                    7 => (0, 0),
                    8 => (1, 1),
                    9 => (2, 3),
                    10 => (3, 2),
                    11 => (4, 4),
                    12 => (5, 5),
                    13 => (6, 6),
                    16 => (7, 7),
                    _ => unreachable!(),
                };
                if std::mem::replace(&mut coordinates[slot], true) {
                    return Err(Error::InvalidSource);
                }
                validate_inert_coordinate(field.payload(), expected_column, package, budget)?;
            },
            14 if owner_index.is_none() => owner_index = Some(canonical_field_u32(field)?),
            18 if root.is_none() && field.wire_type() == 2 => {
                root = Some(strict_reference(field.payload(), limits, budget)?);
            },
            2 | 4 | 5 | 15 | 17 => return Err(Error::UnsupportedDependency),
            _ => return Err(Error::InvalidSource),
        }
    }
    let _ = parse_uuid(uid.ok_or(Error::InvalidSource)?, limits, budget)?;
    if enabled != Some(false)
        || owner_index != Some(8)
        || !embedded_root
        || coordinates.iter().any(|seen| !seen)
        || require_root_reference != root.is_some()
    {
        return Err(Error::UnsupportedDependency);
    }
    Ok(root)
}

fn validate_inert_group_node(payload: &[u8], package: &Package, budget: &mut Budget) -> Result<()> {
    let fields = validated_wire_view(payload, package, budget)?;
    let limits = budget.residual(package)?;
    let mut uid = None;
    let mut format_manager_seen = false;
    for field in fields.fields() {
        match field.number() {
            1 if uid.is_none() && field.wire_type() == 2 => uid = Some(field.payload()),
            6 if !format_manager_seen && field.wire_type() == 2 => {
                format_manager_seen = true;
                if !field.payload().is_empty() {
                    return Err(Error::UnsupportedDependency);
                }
            },
            3..=5 | 7..=10 => return Err(Error::UnsupportedDependency),
            _ => return Err(Error::InvalidSource),
        }
    }
    if parse_uuid(uid.ok_or(Error::InvalidSource)?, limits, budget)? == (1, 0) {
        Ok(())
    } else {
        Err(Error::UnsupportedDependency)
    }
}

fn validate_inert_coordinate(
    payload: &[u8],
    expected_column: u32,
    package: &Package,
    budget: &mut Budget,
) -> Result<()> {
    let fields = validated_wire_view(payload, package, budget)?;
    let mut column = None;
    let mut row = None;
    for field in fields.fields() {
        match field.number() {
            2 if column.is_none() => column = Some(canonical_field_u32(field)?),
            3 if row.is_none() => row = Some(canonical_field_u32(field)?),
            _ => return Err(Error::InvalidSource),
        }
    }
    if column == Some(expected_column) && row == Some(0) {
        Ok(())
    } else {
        Err(Error::UnsupportedDependency)
    }
}

fn field_bytes<'a>(field: litchi_iwa_common::wire::WireFieldView<'a>) -> Result<&'a [u8]> {
    if field.wire_type() != 2 {
        return Err(Error::InvalidSource);
    }
    Ok(field.payload())
}

fn canonical_field_u32(field: litchi_iwa_common::wire::WireFieldView<'_>) -> Result<u32> {
    if field.wire_type() != 0 {
        return Err(Error::InvalidSource);
    }
    let raw = field.payload();
    let (value, width) = decode_varint_from_bytes(raw).map_err(|_| Error::InvalidSource)?;
    if width != encoded_len(value) {
        return Err(Error::InvalidSource);
    }
    u32::try_from(value).map_err(|_| Error::InvalidSource)
}

fn decode_tile_storage(
    payload: &[u8],
    package: &Package,
    budget: &mut Budget,
) -> Result<(u32, Option<bool>)> {
    let options = budget.storage_codec_options(package)?;
    let (snapshot, report) = storage_codec::decode_tile_storage_with_report(payload, options)
        .map_err(|_| Error::Codec)?;
    budget.storage_codec_report(report)?;
    Ok((
        snapshot.tile_size().ok_or(Error::InvalidSource)?,
        snapshot.should_use_wide_rows(),
    ))
}

fn validate_tile_storage(
    package: &Package,
    target: &Target,
    references: &[(u32, u64)],
    tile_size: u32,
    budget: &mut Budget,
) -> Result<Vec<PhysicalTileReference>> {
    if references.is_empty() {
        return Err(Error::UnsupportedDependency);
    }
    budget.allocations(
        references
            .len()
            .checked_mul(2)
            .ok_or(Error::InvalidSource)?,
    )?;
    let mut seen_keys = HashSet::new();
    let mut seen_identifiers = HashSet::new();
    seen_keys
        .try_reserve(references.len())
        .map_err(|_| Error::Allocation(references.len()))?;
    seen_identifiers
        .try_reserve(references.len())
        .map_err(|_| Error::Allocation(references.len()))?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(references.len())
        .map_err(|_| Error::Allocation(references.len()))?;
    for (tile_id, identifier) in references.iter().copied() {
        if !seen_keys.insert(tile_id) || identifier == 0 || !seen_identifiers.insert(identifier) {
            return Err(Error::InvalidSource);
        }
        let location = locate_object(package, identifier)?;
        ensure_unique_identity(package, identifier, budget)?;
        let object = object_at(package, &location)?;
        let (message_index, tile_payload) =
            unique_mutable_message(object, &[TILE_MESSAGE_TYPE], budget)?;
        let _ = message_index;
        validate_physical_tile_wire(tile_payload, package, budget)?;
        let mut rows = PhysicalTileRowCollector::new(target.columns);
        let options = budget.storage_codec_options(package)?;
        let (tile, report) =
            storage_codec::decode_tile_with_visitor(tile_payload, options, &mut rows)
                .map_err(|_| Error::Codec)?;
        budget.storage_codec_report(report)?;
        if tile.max_column() >= target.columns
            || tile.max_row() >= target.rows
            || tile.num_rows() > tile_size
        {
            return Err(Error::InvalidSource);
        }
        let mut row_indices = HashSet::new();
        row_indices
            .try_reserve(rows.rows.len())
            .map_err(|_| Error::Allocation(rows.rows.len()))?;
        let mut maximum = 0u32;
        let mut cell_total = 0u32;
        for row in rows.rows {
            if row.index >= tile_size || !row_indices.insert(row.index) {
                return Err(Error::InvalidSource);
            }
            let row_end = row.index.checked_add(1).ok_or(Error::InvalidSource)?;
            maximum = maximum.max(row_end);
            let global = tile_id
                .checked_mul(tile_size)
                .and_then(|base| base.checked_add(row.index))
                .ok_or(Error::InvalidSource)?;
            if global >= target.rows {
                return Err(Error::InvalidSource);
            }
            cell_total = cell_total
                .checked_add(row.cell_count)
                .ok_or(Error::InvalidSource)?;
        }
        if maximum != tile.num_rows() || cell_total != tile.num_cells() {
            return Err(Error::InvalidSource);
        }
        output.push(PhysicalTileReference { tile_id, location });
    }
    Ok(output)
}

fn validate_row_headers(
    package: &Package,
    target: &Target,
    references: &[u64],
    expected_hash: u32,
    budget: &mut Budget,
) -> Result<Vec<PhysicalRowHeaderBucket>> {
    let expected = target.rows.div_ceil(HEADER_BUCKET_ROWS);
    if references.len() != usize::try_from(expected).map_err(|_| Error::InvalidSource)? {
        return Err(Error::InvalidSource);
    }
    let mut seen = HashSet::new();
    seen.try_reserve(references.len())
        .map_err(|_| Error::Allocation(references.len()))?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(references.len())
        .map_err(|_| Error::Allocation(references.len()))?;
    for (bucket_index, identifier) in references.iter().copied().enumerate() {
        if identifier == 0 || !seen.insert(identifier) {
            return Err(Error::InvalidSource);
        }
        let location = locate_object(package, identifier)?;
        ensure_unique_identity(package, identifier, budget)?;
        let object = object_at(package, &location)?;
        let (_, payload) = unique_mutable_message(object, &[HEADER_BUCKET_MESSAGE_TYPE], budget)?;
        validate_physical_header_bucket_wire(payload, package, budget)?;
        let mut headers = PhysicalHeaderCollector::new();
        let options = budget.storage_codec_options(package)?;
        let (bucket, report) = storage_codec::decode_header_storage_bucket_with_visitor(
            payload,
            options,
            &mut headers,
        )
        .map_err(|_| Error::Codec)?;
        budget.storage_codec_report(report)?;
        if bucket.bucket_hash_function() != expected_hash {
            return Err(Error::InvalidSource);
        }
        let bucket_start = u32::try_from(bucket_index)
            .ok()
            .and_then(|index| index.checked_mul(HEADER_BUCKET_ROWS))
            .ok_or(Error::InvalidSource)?;
        let bucket_end = bucket_start
            .checked_add(HEADER_BUCKET_ROWS)
            .map_or(target.rows, |end| end.min(target.rows));
        let mut indices = HashSet::new();
        indices
            .try_reserve(headers.indices.len())
            .map_err(|_| Error::Allocation(headers.indices.len()))?;
        for index in headers.indices.iter().copied() {
            if index < bucket_start || index >= bucket_end || !indices.insert(index) {
                return Err(Error::InvalidSource);
            }
        }
        output.push(PhysicalRowHeaderBucket {
            bucket_index: u32::try_from(bucket_index).map_err(|_| Error::InvalidSource)?,
            location,
            header_count: headers.indices.len(),
        });
    }
    Ok(output)
}

fn validate_column_headers(
    package: &Package,
    location: &ObjectLocation,
    columns: u32,
    budget: &mut Budget,
) -> Result<()> {
    let object = object_at(package, location)?;
    let (_, payload) = unique_mutable_message(object, &[HEADER_BUCKET_MESSAGE_TYPE], budget)?;
    validate_physical_header_bucket_wire(payload, package, budget)?;
    let mut headers = PhysicalHeaderCollector::new();
    let options = budget.storage_codec_options(package)?;
    let (bucket, report) =
        storage_codec::decode_header_storage_bucket_with_visitor(payload, options, &mut headers)
            .map_err(|_| Error::Codec)?;
    budget.storage_codec_report(report)?;
    if bucket.bucket_hash_function() == 0 {
        return Err(Error::InvalidSource);
    }
    let mut seen = HashSet::new();
    seen.try_reserve(headers.indices.len())
        .map_err(|_| Error::Allocation(headers.indices.len()))?;
    for index in headers.indices {
        if index >= columns || !seen.insert(index) {
            return Err(Error::InvalidSource);
        }
    }
    Ok(())
}

fn validate_row_uid_map(
    package: &Package,
    location: &ObjectLocation,
    rows: u32,
    columns: u32,
    budget: &mut Budget,
) -> Result<()> {
    let object = object_at(package, location)?;
    let (_, payload) = unique_mutable_message(
        object,
        &[
            COLUMN_ROW_UID_MAP_LEGACY_MESSAGE_TYPE,
            COLUMN_ROW_UID_MAP_MESSAGE_TYPE,
        ],
        budget,
    )?;
    validate_physical_uid_map_wire(payload, package, budget)?;
    let limits = budget.residual(package)?;
    let fields = WireView::parse_with_limits(payload, limits).map_err(|_| Error::Wire)?;
    let mut sorted_columns = Vec::new();
    let mut column_index_for_uid = Vec::new();
    let mut column_uid_for_index = Vec::new();
    let mut sorted_rows = Vec::new();
    let mut row_index_for_uid = Vec::new();
    let mut row_uid_for_index = Vec::new();
    let mut saw = [false; 7];
    for field in fields.fields() {
        budget.fields(1)?;
        field
            .validate_canonical_framing()
            .map_err(|_| Error::Wire)?;
        match field.number() {
            1 | 4 => {
                let slot = field.number() as usize;
                if field.wire_type() != 2 {
                    return Err(Error::InvalidSource);
                }
                let uuid = parse_uuid(field_bytes(field)?, limits, budget)?;
                if (slot == 1 && !saw[1]) || (slot == 4 && !saw[4]) {
                    saw[slot] = true;
                }
                if slot == 1 {
                    sorted_columns
                        .try_reserve(1)
                        .map_err(|_| Error::Allocation(1))?;
                    sorted_columns.push(uuid);
                } else {
                    sorted_rows
                        .try_reserve(1)
                        .map_err(|_| Error::Allocation(1))?;
                    sorted_rows.push(uuid);
                }
            },
            2 | 3 | 5 | 6 => {
                let slot = field.number() as usize;
                saw[slot] = true;
                let values = decode_repeated_u32(field)?;
                let destination = match slot {
                    2 => &mut column_index_for_uid,
                    3 => &mut column_uid_for_index,
                    5 => &mut row_index_for_uid,
                    6 => &mut row_uid_for_index,
                    _ => unreachable!(),
                };
                destination
                    .try_reserve(values.len())
                    .map_err(|_| Error::Allocation(values.len()))?;
                destination.extend(values);
            },
            _ => return Err(Error::UnsupportedDependency),
        }
    }
    budget.work(payload.len())?;
    let expected_rows = usize::try_from(rows).map_err(|_| Error::InvalidSource)?;
    let expected_columns = usize::try_from(columns).map_err(|_| Error::InvalidSource)?;
    if !saw[1]
        || !saw[2]
        || !saw[3]
        || !saw[4]
        || !saw[5]
        || !saw[6]
        || sorted_columns.len() != expected_columns
        || column_index_for_uid.len() != expected_columns
        || column_uid_for_index.len() != expected_columns
        || sorted_rows.len() != expected_rows
        || row_index_for_uid.len() != expected_rows
        || row_uid_for_index.len() != expected_rows
    {
        return Err(Error::InvalidSource);
    }
    validate_uid_axis(
        &sorted_columns,
        &column_index_for_uid,
        &column_uid_for_index,
    )?;
    validate_uid_axis(&sorted_rows, &row_index_for_uid, &row_uid_for_index)
}

/// Prove that every mutable physical root belongs exclusively to the
/// selected model.  The archive header is the authoritative package-wide
/// edge census; the strict model/DataStore projection closes the native
/// payload routes that may be omitted from `ArchiveInfo` in older packages.
/// A second model that names one of these roots would otherwise let a row
/// permutation mutate data outside the selected table's ownership boundary.
fn validate_mutable_root_aliases(
    package: &Package,
    target: &Target,
    tiles: &[PhysicalTileReference],
    headers: &[PhysicalRowHeaderBucket],
    column_headers: &ObjectLocation,
    row_uid_identifier: u64,
    budget: &mut Budget,
) -> Result<()> {
    let capacity = tiles
        .len()
        .checked_add(headers.len())
        .and_then(|value| value.checked_add(2))
        .ok_or(Error::InvalidSource)?;
    budget.allocations(capacity.checked_mul(2).ok_or(Error::InvalidSource)?)?;
    budget.retained(
        capacity
            .checked_mul(size_of::<u64>())
            .and_then(|value| value.checked_mul(2))
            .ok_or(Error::InvalidSource)?,
    )?;
    let mut mutable_identifiers = HashSet::new();
    mutable_identifiers
        .try_reserve(capacity)
        .map_err(|_| Error::Allocation(capacity))?;
    for identifier in tiles
        .iter()
        .map(|tile| tile.location.identifier)
        .chain(headers.iter().map(|header| header.location.identifier))
        .chain([column_headers.identifier, row_uid_identifier])
    {
        if identifier == 0 || !mutable_identifiers.insert(identifier) {
            // A physical object cannot safely serve two mutable roles: the
            // writer would have no single message-role transformation for it.
            return Err(Error::UnsupportedDependency);
        }
    }

    let model_object = object_at(package, &target.model)?;
    let model_info = model_object
        .archive_info
        .message_infos
        .get(target.model_message_index)
        .ok_or(Error::InvalidSource)?;
    let mut expected = HashMap::new();
    expected
        .try_reserve(mutable_identifiers.len())
        .map_err(|_| Error::Allocation(mutable_identifiers.len()))?;
    let mut allowed_fields = HashMap::new();
    allowed_fields
        .try_reserve(mutable_identifiers.len())
        .map_err(|_| Error::Allocation(mutable_identifiers.len()))?;
    for identifier in mutable_identifiers.iter().copied() {
        let aggregate_count = model_info
            .object_references
            .iter()
            .filter(|candidate| **candidate == identifier)
            .count();
        if aggregate_count != 1 {
            return Err(Error::UnsupportedDependency);
        }
        let mut local_field = None;
        for (field_index, field) in model_info.field_infos.iter().enumerate() {
            if !field.object_references.contains(&identifier) {
                continue;
            }
            if field.object_references.as_slice() != [identifier]
                || !field.data_references.is_empty()
                || !is_message_reference_field(field)
                || local_field.replace(field_index).is_some()
            {
                return Err(Error::UnsupportedDependency);
            }
        }
        let expected_count = if let Some(field_index) = local_field {
            if allowed_fields.insert(identifier, field_index).is_some() {
                return Err(Error::UnsupportedDependency);
            }
            // ArchiveInfo exposes the same authoritative route once in the
            // aggregate list and once in the field-local list. The census
            // counts only the field-local occurrence for this exact role.
            1
        } else {
            aggregate_count
        };
        if expected.insert(identifier, expected_count).is_some() {
            return Err(Error::UnsupportedDependency);
        }
    }

    let archive_limits = package
        .state
        .options
        .archive()
        .effective_archive_limits()
        .map_err(|_| Error::Archive)?;
    let mut census = MutableRootInboundReferenceCensus {
        expected: &expected,
        allowed_fields: &allowed_fields,
        owner_identifier: target.model.identifier,
        owner_message_index: target.model_message_index,
        observed: HashMap::new(),
        invalid: false,
    };

    for component in package.state.source.components().iter() {
        budget.components(1)?;
        for object in &component.archive().objects {
            budget.payload_objects(1)?;
            if object.archive_info.identifier.is_none()
                || object.messages.len() != object.archive_info.message_infos.len()
            {
                return Err(Error::InvalidSource);
            }
            let field_count = object
                .archive_info
                .message_infos
                .iter()
                .map(|info| info.field_infos.len())
                .try_fold(0usize, |sum, count| sum.checked_add(count))
                .ok_or(Error::InvalidSource)?;
            let reference_count =
                object
                    .archive_info
                    .message_infos
                    .iter()
                    .try_fold(0usize, |sum, info| {
                        let nested = info.field_infos.iter().try_fold(0usize, |sum, field| {
                            sum.checked_add(field.object_references.len())
                                .and_then(|value| value.checked_add(field.data_references.len()))
                                .ok_or(Error::InvalidSource)
                        })?;
                        sum.checked_add(info.object_references.len())
                            .and_then(|value| value.checked_add(info.data_references.len()))
                            .and_then(|value| value.checked_add(nested))
                            .ok_or(Error::InvalidSource)
                    })?;
            let message_bytes = object
                .messages
                .iter()
                .map(|message| message.data.len())
                .try_fold(0usize, |sum, bytes| sum.checked_add(bytes))
                .ok_or(Error::InvalidSource)?;
            budget.payload_messages(object.messages.len())?;
            budget.fields(field_count)?;
            budget.references(reference_count)?;
            budget.allocations(
                object
                    .messages
                    .len()
                    .checked_add(field_count)
                    .and_then(|value| value.checked_add(reference_count))
                    .and_then(|value| value.checked_add(1))
                    .ok_or(Error::InvalidSource)?,
            )?;
            budget.work(
                object
                    .messages
                    .len()
                    .checked_add(message_bytes)
                    .and_then(|value| value.checked_add(field_count))
                    .and_then(|value| value.checked_add(reference_count))
                    .ok_or(Error::InvalidSource)?,
            )?;
            object
                .inspect_references_with_policy_and_limits(
                    &mut census,
                    ArchiveReferencePolicy::RejectUnknownMetadata,
                    archive_limits,
                )
                .map_err(|_| Error::Archive)?;

            for (message_index, message) in object.messages.iter().enumerate() {
                validate_message_header(object, message_index)?;
                if message.type_ != TABLE_MODEL_MESSAGE_TYPE
                    || (object.archive_info.identifier == Some(target.model.identifier)
                        && message_index == target.model_message_index)
                {
                    continue;
                }

                let options = budget.storage_codec_options(package)?;
                let mut collector = PhysicalStorageCollector::new();
                validate_model_storage_envelope_wire(&message.data, package, budget)?;
                let (projection, report) =
                    storage_codec::decode_table_model_with_data_store_and_visitor(
                        &message.data,
                        options,
                        &mut collector,
                    )
                    .map_err(|_| Error::Codec)?;
                budget.storage_codec_report(report)?;
                let store = projection.data_store();
                let optional_alias = [
                    store.formula_error_table(),
                    store.merge_region_map(),
                    store.deprecated_custom_format_table(),
                    store.multiple_choice_list_format_table(),
                    store.rich_text_table(),
                    store.conditional_style_table(),
                    store.comment_storage_table(),
                    store.import_warning_set_table(),
                    store.control_cell_spec_table(),
                    store.format_table(),
                ]
                .into_iter()
                .flatten()
                .any(|reference| mutable_identifiers.contains(&reference.identifier()));
                if collector
                    .tile_references
                    .iter()
                    .any(|(_, identifier)| mutable_identifiers.contains(identifier))
                    || collector
                        .row_header_references
                        .iter()
                        .any(|identifier| mutable_identifiers.contains(identifier))
                    || mutable_identifiers.contains(&store.column_headers().identifier())
                    || mutable_identifiers.contains(&store.string_table().identifier())
                    || mutable_identifiers.contains(&store.style_table().identifier())
                    || mutable_identifiers.contains(&store.formula_table().identifier())
                    || mutable_identifiers.contains(&store.format_table_pre_bnc().identifier())
                    || optional_alias
                    || parse_model_uid_reference(&message.data, package, budget)?
                        .is_some_and(|identifier| mutable_identifiers.contains(&identifier))
                {
                    return Err(Error::UnsupportedDependency);
                }
            }
        }
    }

    let complete = expected
        .iter()
        .all(|(identifier, count)| census.observed.get(identifier).copied() == Some(*count));
    if census.invalid || !complete {
        Err(Error::UnsupportedDependency)
    } else {
        Ok(())
    }
}

fn parse_model_uid_reference(
    payload: &[u8],
    package: &Package,
    budget: &mut Budget,
) -> Result<Option<u64>> {
    let limits = budget.residual(package)?;
    let fields = WireView::parse_with_limits(payload, limits).map_err(|_| Error::Wire)?;
    let mut identifier = None;
    for field in fields.fields() {
        budget.fields(1)?;
        field
            .validate_canonical_framing()
            .map_err(|_| Error::Wire)?;
        if field.number() == MODEL_BASE_COLUMN_ROW_UIDS_FIELD {
            let value = strict_reference(field_bytes(field)?, limits, budget)?;
            if identifier.replace(value).is_some() {
                return Err(Error::InvalidSource);
            }
        }
    }
    budget.work(payload.len())?;
    Ok(identifier)
}

fn parse_uuid(payload: &[u8], limits: WireLimits, budget: &mut Budget) -> Result<(u64, u64)> {
    let fields = WireView::parse_with_limits(payload, limits).map_err(|_| Error::Wire)?;
    let mut lower = None;
    let mut upper = None;
    for field in fields.fields() {
        budget.fields(1)?;
        field
            .validate_canonical_framing()
            .map_err(|_| Error::Wire)?;
        let value = match field.number() {
            1 | 2 => {
                if field.wire_type() != 0 {
                    return Err(Error::InvalidSource);
                }
                let (value, width) =
                    decode_varint_from_bytes(field.payload()).map_err(|_| Error::InvalidSource)?;
                if width != encoded_len(value) {
                    return Err(Error::InvalidSource);
                }
                value
            },
            _ => return Err(Error::InvalidSource),
        };
        if field.number() == 1 {
            if lower.replace(value).is_some() {
                return Err(Error::InvalidSource);
            }
        } else if upper.replace(value).is_some() {
            return Err(Error::InvalidSource);
        }
    }
    let lower = lower.ok_or(Error::InvalidSource)?;
    let upper = upper.ok_or(Error::InvalidSource)?;
    if lower == 0 && upper == 0 {
        return Err(Error::InvalidSource);
    }
    Ok((lower, upper))
}

fn decode_repeated_u32(field: litchi_iwa_common::wire::WireFieldView<'_>) -> Result<Vec<u32>> {
    let payload = field.payload();
    let mut values = Vec::new();
    match field.wire_type() {
        0 => values.push(canonical_varint_u32(payload)?),
        2 => {
            let mut remaining = payload;
            while !remaining.is_empty() {
                let (value, width) =
                    decode_varint_from_bytes(remaining).map_err(|_| Error::InvalidSource)?;
                if width != encoded_len(value) {
                    return Err(Error::InvalidSource);
                }
                values.push(u32::try_from(value).map_err(|_| Error::InvalidSource)?);
                remaining = &remaining[width..];
            }
        },
        _ => return Err(Error::InvalidSource),
    }
    Ok(values)
}

fn canonical_varint_u32(payload: &[u8]) -> Result<u32> {
    let (value, width) = decode_varint_from_bytes(payload).map_err(|_| Error::InvalidSource)?;
    if width != encoded_len(value) {
        return Err(Error::InvalidSource);
    }
    u32::try_from(value).map_err(|_| Error::InvalidSource)
}

fn validate_uid_axis(
    sorted_uids: &[(u64, u64)],
    index_for_uid: &[u32],
    uid_for_index: &[u32],
) -> Result<()> {
    if sorted_uids
        .windows(2)
        .any(|pair| (pair[0].1, pair[0].0) >= (pair[1].1, pair[1].0))
    {
        return Err(Error::InvalidSource);
    }
    let length = sorted_uids.len();
    let mut seen_index = HashSet::new();
    let mut seen_uid = HashSet::new();
    seen_index
        .try_reserve(length)
        .map_err(|_| Error::Allocation(length))?;
    seen_uid
        .try_reserve(length)
        .map_err(|_| Error::Allocation(length))?;
    for index in index_for_uid.iter().copied() {
        if usize::try_from(index)
            .ok()
            .filter(|index| *index < length)
            .is_none()
            || !seen_index.insert(index)
        {
            return Err(Error::InvalidSource);
        }
    }
    for uid in uid_for_index.iter().copied() {
        if usize::try_from(uid)
            .ok()
            .filter(|uid| *uid < length)
            .is_none()
            || !seen_uid.insert(uid)
        {
            return Err(Error::InvalidSource);
        }
    }
    for (physical, uid) in uid_for_index.iter().copied().enumerate() {
        let sorted = usize::try_from(uid).map_err(|_| Error::InvalidSource)?;
        let mapped = usize::try_from(index_for_uid[sorted]).map_err(|_| Error::InvalidSource)?;
        if mapped != physical {
            return Err(Error::InvalidSource);
        }
    }
    Ok(())
}

fn validate_empty_stroke_sidecar(
    package: &Package,
    location: &ObjectLocation,
    target: &Target,
    budget: &mut Budget,
) -> Result<()> {
    let object = object_at(package, location)?;
    let (_, payload) = unique_mutable_message(object, &[STROKE_SIDECAR_MESSAGE_TYPE], budget)?;
    let limits = budget.residual(package)?;
    let fields = WireView::parse_with_limits(payload, limits).map_err(|_| Error::Wire)?;
    let mut dimensions = [None; 3];
    let mut nonempty_layers = false;
    for field in fields.fields() {
        budget.fields(1)?;
        field
            .validate_canonical_framing()
            .map_err(|_| Error::Wire)?;
        match field.number() {
            1..=3 => {
                let slot = usize::try_from(field.number() - 1).map_err(|_| Error::InvalidSource)?;
                if dimensions[slot]
                    .replace(canonical_field_u32(field)?)
                    .is_some()
                {
                    return Err(Error::InvalidSource);
                }
            },
            4..=7 => {
                if field.wire_type() != 2 {
                    return Err(Error::InvalidSource);
                }
                let reference = strict_reference(field_bytes(field)?, limits, budget)?;
                if reference == 0 {
                    return Err(Error::InvalidSource);
                }
                nonempty_layers = true;
            },
            _ => return Err(Error::UnsupportedDependency),
        }
    }
    if dimensions[1].is_some_and(|columns| columns != target.columns)
        || dimensions[2].is_some_and(|rows| rows != target.rows)
        || nonempty_layers
    {
        return Err(Error::UnsupportedDependency);
    }
    Ok(())
}

fn optional_storage_route(
    package: &Package,
    reference: Option<storage_codec::ReferenceSnapshot>,
    required: &[StorageRoute],
    budget: &mut Budget,
) -> Result<Option<ObjectLocation>> {
    let Some(reference) = reference else {
        return Ok(None);
    };
    let identifier = reference.identifier();
    if identifier == 0
        || required
            .iter()
            .any(|route| route.location.identifier == identifier)
    {
        return Err(Error::InvalidSource);
    }
    let location = locate_object(package, identifier)?;
    ensure_unique_identity(package, identifier, budget)?;
    Ok(Some(location))
}

/// Prove an optional DataStore route against the model message's exact graph
/// metadata.  The decoded DataStore reference is not sufficient on its own:
/// archive metadata is the authority used by the IWA object graph, and an
/// opaque row-carried identifier is safe to retain only when that authority
/// names the same object at the canonical nested field path.
pub(crate) fn validate_optional_model_route(
    package: &Package,
    target: &Target,
    identifier: u64,
    path: &[u32],
    budget: &mut Budget,
) -> Result<()> {
    if identifier == 0 {
        return Err(Error::InvalidSource);
    }
    let model = object_at(package, &target.model)?;
    let info = model
        .archive_info
        .message_infos
        .get(target.model_message_index)
        .ok_or(Error::InvalidSource)?;

    let aggregate_count = info
        .object_references
        .iter()
        .filter(|candidate| **candidate == identifier)
        .count();
    budget.references(aggregate_count)?;
    if aggregate_count != 1 {
        return Err(Error::UnsupportedDependency);
    }

    let mut field_count = 0usize;
    for field in &info.field_infos {
        let references = field
            .object_references
            .iter()
            .filter(|candidate| **candidate == identifier)
            .count();
        budget.references(
            references
                .checked_add(field.data_references.len())
                .ok_or(Error::InvalidSource)?,
        )?;
        if references == 0 {
            continue;
        }
        if references != 1
            || field.object_references.as_slice() != [identifier]
            || !field.data_references.is_empty()
            || !is_message_reference_field(field)
            || field.path.as_slice() != path
        {
            return Err(Error::UnsupportedDependency);
        }
        field_count = field_count.checked_add(1).ok_or(Error::InvalidSource)?;
    }
    // Physical row movement cannot safely infer a route from the aggregate
    // MessageInfo list. Require the one exact field-local role even for
    // producers that also retain the aggregate edge for graph traversal.
    if field_count > 1 {
        return Err(Error::UnsupportedDependency);
    }
    Ok(())
}

fn validate_exclusive_model_reference(
    package: &Package,
    target: &Target,
    identifier: u64,
    path: &[u32],
    budget: &mut Budget,
) -> Result<()> {
    let model = object_at(package, &target.model)?;
    let info = model
        .archive_info
        .message_infos
        .get(target.model_message_index)
        .ok_or(Error::InvalidSource)?;
    require_local_reference_field(info, identifier, path)?;
    let allowed_field_index = info
        .field_infos
        .iter()
        .enumerate()
        .filter(|(_, field)| {
            field.path.as_slice() == path
                && field.object_references.as_slice() == [identifier]
                && field.data_references.is_empty()
                && is_message_reference_field(field)
        })
        .map(|(index, _)| index)
        .next()
        .ok_or(Error::UnsupportedDependency)?;
    let archive_limits = package
        .state
        .options
        .archive()
        .effective_archive_limits()
        .map_err(|_| Error::Archive)?;
    let wire_limits = budget.residual(package)?;
    let mut census = ExclusiveInboundReferenceCensus {
        identifier,
        owner_identifier: target.model.identifier,
        owner_message_index: target.model_message_index,
        allowed_field_index: Some(allowed_field_index),
        observed: 0,
        invalid: false,
    };
    let mut payload_owners = 0usize;
    for component in package.state.source.components().iter() {
        budget.components(1)?;
        for object in &component.archive().objects {
            budget.payload_objects(1)?;
            if object.archive_info.identifier.is_none()
                || object.messages.len() != object.archive_info.message_infos.len()
            {
                return Err(Error::InvalidSource);
            }
            budget.payload_messages(object.messages.len())?;
            for (message_index, message) in object.messages.iter().enumerate() {
                validate_message_header(object, message_index)?;
                budget.work(
                    message
                        .data
                        .len()
                        .checked_add(1)
                        .ok_or(Error::InvalidSource)?,
                )?;
                if message.type_ != TABLE_MODEL_MESSAGE_TYPE {
                    continue;
                }
                let view = WireView::parse_with_limits(&message.data, wire_limits)
                    .map_err(|_| Error::Wire)?;
                for field in view.fields() {
                    if field.number() != MODEL_BASE_COLUMN_ROW_UIDS_FIELD {
                        continue;
                    }
                    field
                        .validate_canonical_framing()
                        .map_err(|_| Error::Wire)?;
                    let referenced = strict_reference(field_bytes(field)?, wire_limits, budget)?;
                    if referenced == identifier {
                        payload_owners =
                            payload_owners.checked_add(1).ok_or(Error::InvalidSource)?;
                        if object.archive_info.identifier != Some(target.model.identifier)
                            || message_index != target.model_message_index
                        {
                            return Err(Error::UnsupportedDependency);
                        }
                    }
                }
            }
            let field_count = object
                .archive_info
                .message_infos
                .iter()
                .map(|message| message.field_infos.len())
                .try_fold(0usize, |sum, count| sum.checked_add(count))
                .ok_or(Error::InvalidSource)?;
            let reference_count =
                object
                    .archive_info
                    .message_infos
                    .iter()
                    .try_fold(0usize, |sum, message| {
                        let nested =
                            message.field_infos.iter().try_fold(0usize, |sum, field| {
                                sum.checked_add(field.object_references.len())
                                    .and_then(|value| {
                                        value.checked_add(field.data_references.len())
                                    })
                                    .ok_or(Error::InvalidSource)
                            })?;
                        sum.checked_add(message.object_references.len())
                            .and_then(|value| value.checked_add(message.data_references.len()))
                            .and_then(|value| value.checked_add(nested))
                            .ok_or(Error::InvalidSource)
                    })?;
            budget.fields(field_count)?;
            budget.references(reference_count)?;
            budget.work(
                field_count
                    .checked_add(reference_count)
                    .ok_or(Error::InvalidSource)?,
            )?;
            object
                .inspect_references_with_policy_and_limits(
                    &mut census,
                    ArchiveReferencePolicy::RejectUnknownMetadata,
                    archive_limits,
                )
                .map_err(|_| Error::Archive)?;
        }
    }
    if payload_owners == 1 && !census.invalid && census.observed == 1 {
        Ok(())
    } else {
        Err(Error::UnsupportedDependency)
    }
}

struct MutableRootInboundReferenceCensus<'a> {
    expected: &'a HashMap<u64, usize>,
    allowed_fields: &'a HashMap<u64, usize>,
    owner_identifier: u64,
    owner_message_index: usize,
    observed: HashMap<u64, usize>,
    invalid: bool,
}

impl ArchiveReferenceVisitor for MutableRootInboundReferenceCensus<'_> {
    fn visit_reference(
        &mut self,
        occurrence: ArchiveReferenceOccurrence,
    ) -> litchi_iwa_core::Result<()> {
        let Some(expected) = self.expected.get(&occurrence.referenced_identifier) else {
            return Ok(());
        };
        if occurrence.kind != ArchiveReferenceKind::Object
            || occurrence.object_identifier != self.owner_identifier
            || occurrence.message_index != self.owner_message_index
        {
            self.invalid = true;
            return Ok(());
        }
        let allowed_field = self.allowed_fields.get(&occurrence.referenced_identifier);
        match occurrence.scope {
            ArchiveReferenceScope::Message if allowed_field.is_some() => {
                // The aggregate edge is retained for the archive graph, but
                // it is not a row-affine role proof. Count only its exact
                // field-local counterpart below.
                return Ok(());
            },
            ArchiveReferenceScope::Field { field_index }
                if allowed_field.is_some_and(|expected| *expected == field_index) => {},
            ArchiveReferenceScope::Message => {},
            ArchiveReferenceScope::Field { .. } => {
                self.invalid = true;
                return Ok(());
            },
        }
        let count = self
            .observed
            .entry(occurrence.referenced_identifier)
            .or_insert(0);
        if let Some(next) = count.checked_add(1) {
            *count = next;
            if next > *expected {
                self.invalid = true;
            }
        } else {
            self.invalid = true;
        }
        Ok(())
    }
}

struct ExclusiveInboundReferenceCensus {
    identifier: u64,
    owner_identifier: u64,
    owner_message_index: usize,
    allowed_field_index: Option<usize>,
    observed: usize,
    invalid: bool,
}

impl ArchiveReferenceVisitor for ExclusiveInboundReferenceCensus {
    fn visit_reference(
        &mut self,
        occurrence: ArchiveReferenceOccurrence,
    ) -> litchi_iwa_core::Result<()> {
        if occurrence.referenced_identifier != self.identifier {
            return Ok(());
        }
        let allowed_scope = match occurrence.scope {
            ArchiveReferenceScope::Message => {
                // Keep the aggregate edge as graph metadata, but require
                // the selected field-local role to prove exclusivity.
                self.allowed_field_index.is_some()
            },
            ArchiveReferenceScope::Field { field_index } => {
                self.allowed_field_index == Some(field_index)
            },
        };
        if occurrence.kind != ArchiveReferenceKind::Object
            || occurrence.object_identifier != self.owner_identifier
            || occurrence.message_index != self.owner_message_index
            || !allowed_scope
        {
            self.invalid = true;
        } else if matches!(occurrence.scope, ArchiveReferenceScope::Message) {
            // The aggregate occurrence is expected, but only the exact
            // field-local occurrence contributes to the exclusivity count.
        } else if let Some(next) = self.observed.checked_add(1) {
            self.observed = next;
        } else {
            self.invalid = true;
        }
        Ok(())
    }
}

fn validate_data_list_state(
    package: &Package,
    location: &ObjectLocation,
    expected_list_type: i32,
    budget: &mut Budget,
) -> Result<bool> {
    let object = object_at(package, location)?;
    if !has_exact_mutable_message_shape(
        object,
        &[
            TABLE_DATA_LIST_MESSAGE_TYPE,
            TABLE_DATA_LIST_NATIVE_MESSAGE_TYPE,
        ],
    ) {
        return Err(Error::UnsupportedDependency);
    }
    let (_, payload) = unique_message_any(
        object,
        &[
            TABLE_DATA_LIST_MESSAGE_TYPE,
            TABLE_DATA_LIST_NATIVE_MESSAGE_TYPE,
        ],
        budget,
    )?;
    validate_physical_data_list_wire(payload, package, budget)?;
    let options = budget.storage_codec_options(package)?;
    let mut collector = PhysicalListPresenceCollector::default();
    let (snapshot, report) =
        storage_codec::decode_table_data_list_with_visitor(payload, options, &mut collector)
            .map_err(|_| Error::Codec)?;
    budget.storage_codec_report(report)?;
    if snapshot.list_type() != expected_list_type || collector.has_segments {
        return Err(Error::UnsupportedDependency);
    }
    Ok(collector.has_entries)
}

#[derive(Default)]
struct PhysicalDependencyPresence {
    any: bool,
}

impl dependency_codec::DependencyVisitor for PhysicalDependencyPresence {
    fn visit_formula_owner_dependency(
        &mut self,
        _reference: dependency_codec::ReferenceRecord<'_>,
    ) -> std::result::Result<(), dependency_codec::DecodeError> {
        self.any = true;
        Ok(())
    }

    fn visit_tiled_cell_dependency(
        &mut self,
        _reference: dependency_codec::ReferenceRecord<'_>,
    ) -> std::result::Result<(), dependency_codec::DecodeError> {
        self.any = true;
        Ok(())
    }

    fn visit_tiled_range_dependency(
        &mut self,
        _reference: dependency_codec::ReferenceRecord<'_>,
    ) -> std::result::Result<(), dependency_codec::DecodeError> {
        self.any = true;
        Ok(())
    }

    fn visit_cell_record(
        &mut self,
        _record: dependency_codec::CellRecordSnapshot<'_>,
    ) -> std::result::Result<(), dependency_codec::DecodeError> {
        self.any = true;
        Ok(())
    }

    fn visit_range_back_dependency(
        &mut self,
        _record: dependency_codec::RangeBackDependencySnapshot<'_>,
    ) -> std::result::Result<(), dependency_codec::DecodeError> {
        self.any = true;
        Ok(())
    }

    fn visit_from_to_range(
        &mut self,
        _record: dependency_codec::FromToRangeSnapshot<'_>,
    ) -> std::result::Result<(), dependency_codec::DecodeError> {
        self.any = true;
        Ok(())
    }

    fn visit_expanded_edge_component(
        &mut self,
        _component: dependency_codec::ExpandedEdgeComponent,
    ) -> std::result::Result<(), dependency_codec::DecodeError> {
        self.any = true;
        Ok(())
    }
}

fn validate_formula_owner_dependencies(
    package: &Package,
    target: &Target,
    header_rows: u32,
    footer_rows: u32,
    budget: &mut Budget,
) -> Result<()> {
    const FORMULA_OWNER_DEPENDENCIES_MESSAGE_TYPE: u32 = 4_008;

    let mut selected_owners = 0usize;
    for component in package.state.source.components().iter() {
        for object in &component.archive().objects {
            let count = object
                .messages
                .iter()
                .filter(|message| message.type_ == FORMULA_OWNER_DEPENDENCIES_MESSAGE_TYPE)
                .count();
            if count == 0 {
                continue;
            }
            if count != 1 || object.messages.len() != 1 {
                return Err(Error::UnsupportedDependency);
            }
            let (_, payload) =
                unique_message(object, FORMULA_OWNER_DEPENDENCIES_MESSAGE_TYPE, budget)?;
            let outer = validated_wire_view(payload, package, budget)?;
            if outer
                .fields()
                .any(|field| !(1..=16).contains(&field.number()))
            {
                return Err(Error::UnsupportedDependency);
            }
            let options = budget.storage_codec_options(package)?;
            let mut presence = PhysicalDependencyPresence::default();
            let (snapshot, report) =
                dependency_codec::decode_formula_owner_dependencies_with_visitor(
                    payload,
                    options,
                    &mut presence,
                )
                .map_err(|_| Error::Codec)?;
            budget.storage_codec_report(report)?;
            let selected = snapshot
                .formula_owner()
                .is_some_and(|reference| reference.identifier() == target.table_info.identifier);
            if !selected {
                continue;
            }
            selected_owners = selected_owners.checked_add(1).ok_or(Error::InvalidSource)?;
            if presence.any
                || snapshot.internal_formula_owner_id() == 0
                || snapshot.formula_owner_uid().lower() == 0
                    && snapshot.formula_owner_uid().upper() == 0
                || snapshot
                    .base_owner_uid()
                    .is_some_and(|uid| uid.lower() == 0 && uid.upper() == 0)
            {
                return Err(Error::UnsupportedDependency);
            }
            for payload in [
                snapshot.cell_dependencies(),
                snapshot.range_dependencies(),
                snapshot.cell_errors(),
                snapshot.tiled_cell_dependencies(),
                snapshot.uuid_references(),
                snapshot.tiled_range_dependencies(),
                snapshot.spill_range_sizes(),
            ]
            .into_iter()
            .flatten()
            {
                if validated_wire_view(payload, package, budget)?
                    .fields()
                    .next()
                    .is_some()
                {
                    return Err(Error::UnsupportedDependency);
                }
            }
            if let Some(payload) = snapshot.volatile_dependencies() {
                validate_empty_volatile_dependencies(payload, package, budget)?;
            }
            if let Some(payload) = snapshot.whole_owner_dependencies() {
                validate_empty_whole_owner_dependencies(payload, package, budget)?;
            }

            if snapshot.owner_kind() != Some(1) || snapshot.base_owner_uid().is_some() {
                return Err(Error::UnsupportedDependency);
            }
            for payload in [
                snapshot.spanning_column_dependencies(),
                snapshot.spanning_row_dependencies(),
            ]
            .into_iter()
            .flatten()
            {
                validate_inert_spanning_dependencies(
                    payload,
                    Some((target, header_rows, footer_rows)),
                    package,
                    budget,
                )?;
            }
        }
    }
    if selected_owners == 1 {
        Ok(())
    } else {
        Err(Error::UnsupportedDependency)
    }
}

fn validate_empty_volatile_dependencies(
    payload: &[u8],
    package: &Package,
    budget: &mut Budget,
) -> Result<()> {
    let fields = validated_wire_view(payload, package, budget)?;
    let mut seen = [false; 8];
    for field in fields.fields() {
        let index = usize::try_from(field.number()).map_err(|_| Error::InvalidSource)?;
        if !matches!(index, 1..=5 | 7)
            || std::mem::replace(&mut seen[index], true)
            || field.wire_type() != 2
            || validated_wire_view(field.payload(), package, budget)?
                .fields()
                .next()
                .is_some()
        {
            return Err(Error::UnsupportedDependency);
        }
    }
    Ok(())
}

fn validate_empty_whole_owner_dependencies(
    payload: &[u8],
    package: &Package,
    budget: &mut Budget,
) -> Result<()> {
    let fields = validated_wire_view(payload, package, budget)?;
    let mut seen = false;
    for field in fields.fields() {
        if field.number() != 1
            || std::mem::replace(&mut seen, true)
            || field.wire_type() != 2
            || validated_wire_view(field.payload(), package, budget)?
                .fields()
                .next()
                .is_some()
        {
            return Err(Error::UnsupportedDependency);
        }
    }
    Ok(())
}

fn validate_inert_spanning_dependencies(
    payload: &[u8],
    selected: Option<(&Target, u32, u32)>,
    package: &Package,
    budget: &mut Budget,
) -> Result<()> {
    let fields = validated_wire_view(payload, package, budget)?;
    let mut total = None;
    let mut body = None;
    for field in fields.fields() {
        match field.number() {
            1 => return Err(Error::UnsupportedDependency),
            2 if total.is_none() && field.wire_type() == 2 => {
                total = Some(validate_range_coordinate(field.payload(), package, budget)?);
            },
            3 if body.is_none() && field.wire_type() == 2 => {
                body = Some(validate_range_coordinate(field.payload(), package, budget)?);
            },
            _ => return Err(Error::UnsupportedDependency),
        }
    }
    if let Some((target, header_rows, footer_rows)) = selected {
        let last_column = target.columns.checked_sub(1).ok_or(Error::InvalidSource)?;
        let last_row = target.rows.checked_sub(1).ok_or(Error::InvalidSource)?;
        let body_last_row = target
            .rows
            .checked_sub(footer_rows)
            .and_then(|rows| rows.checked_sub(1))
            .ok_or(Error::InvalidSource)?;
        if total != Some((0, 0, last_column, last_row))
            || body != Some((0, header_rows, last_column, body_last_row))
        {
            return Err(Error::UnsupportedDependency);
        }
    }
    Ok(())
}

fn validate_range_coordinate(
    payload: &[u8],
    package: &Package,
    budget: &mut Budget,
) -> Result<(u32, u32, u32, u32)> {
    let fields = validated_wire_view(payload, package, budget)?;
    let mut values = [None; 4];
    for field in fields.fields() {
        let index = usize::try_from(field.number())
            .ok()
            .and_then(|number| number.checked_sub(1))
            .filter(|index| *index < values.len())
            .ok_or(Error::InvalidSource)?;
        if values[index].replace(canonical_field_u32(field)?).is_some() {
            return Err(Error::InvalidSource);
        }
    }
    Ok((
        values[0].ok_or(Error::InvalidSource)?,
        values[1].ok_or(Error::InvalidSource)?,
        values[2].ok_or(Error::InvalidSource)?,
        values[3].ok_or(Error::InvalidSource)?,
    ))
}

#[derive(Default)]
struct PhysicalListPresenceCollector {
    has_entries: bool,
    has_segments: bool,
}

impl storage_codec::StorageVisitor for PhysicalListPresenceCollector {
    fn visit_list_entry(
        &mut self,
        _entry: storage_codec::TableDataListEntrySnapshot<'_>,
    ) -> std::result::Result<(), storage_codec::DecodeError> {
        self.has_entries = true;
        Ok(())
    }

    fn visit_list_segment(
        &mut self,
        _reference: storage_codec::ReferenceRecord<'_>,
    ) -> std::result::Result<(), storage_codec::DecodeError> {
        self.has_segments = true;
        Ok(())
    }
}

fn validate_comment_storage(
    package: &Package,
    location: &ObjectLocation,
    budget: &mut Budget,
) -> Result<bool> {
    let object = object_at(package, location)?;
    if !has_exact_mutable_message_shape(
        object,
        &[
            TABLE_DATA_LIST_MESSAGE_TYPE,
            TABLE_DATA_LIST_NATIVE_MESSAGE_TYPE,
        ],
    ) {
        return Err(Error::UnsupportedDependency);
    }
    let (_, payload) = unique_message_any(
        object,
        &[
            TABLE_DATA_LIST_MESSAGE_TYPE,
            TABLE_DATA_LIST_NATIVE_MESSAGE_TYPE,
        ],
        budget,
    )?;
    validate_physical_data_list_wire(payload, package, budget)?;
    let options = budget.storage_codec_options(package)?;
    let mut collector = PhysicalCommentStorageCollector::new();
    let (snapshot, report) =
        storage_codec::decode_table_data_list_with_visitor(payload, options, &mut collector)
            .map_err(|_| Error::Codec)?;
    budget.storage_codec_report(report)?;
    // TST's comment-storage root is list type 10. It is intentionally kept
    // opaque during row movement; BNC cell references travel with their raw
    // row envelopes and no coordinate rewrite is attempted here.
    if snapshot.list_type() != 10 || collector.has_segments || collector.invalid_entry {
        return Err(Error::UnsupportedDependency);
    }
    let entry_count = collector.keys.len();
    if entry_count != collector.identifiers.len() {
        return Err(Error::InvalidSource);
    }
    budget.allocations(entry_count.checked_mul(2).ok_or(Error::InvalidSource)?)?;
    budget.retained(
        entry_count
            .checked_mul(size_of::<u32>())
            .and_then(|bytes| bytes.checked_add(entry_count.checked_mul(size_of::<u64>())?))
            .ok_or(Error::InvalidSource)?,
    )?;
    let mut keys = collector.keys;
    let mut identifiers = collector.identifiers;
    keys.sort_unstable();
    identifiers.sort_unstable();
    if keys.windows(2).any(|pair| pair[0] == pair[1])
        || identifiers.windows(2).any(|pair| pair[0] == pair[1])
    {
        return Err(Error::UnsupportedDependency);
    }
    for identifier in identifiers {
        // The list entry's BNC reference is a row-carried opaque edge, but it
        // must still resolve to exactly one physical object before the owner
        // can move its enclosing row without consulting comment semantics.
        locate_object(package, identifier)?;
        ensure_unique_identity(package, identifier, budget)?;
    }
    Ok(entry_count != 0)
}

/// Fallible, rollback-free staging for the comment-list shape proof.  The
/// storage codec deliberately invokes visitors before its enclosing decode
/// returns, so this collector publishes nothing; `validate_comment_storage`
/// checks every staged value only after strict wire and Buffa parity succeed.
struct PhysicalCommentStorageCollector {
    keys: Vec<u32>,
    identifiers: Vec<u64>,
    has_segments: bool,
    invalid_entry: bool,
}

impl PhysicalCommentStorageCollector {
    fn new() -> Self {
        Self {
            keys: Vec::new(),
            identifiers: Vec::new(),
            has_segments: false,
            invalid_entry: false,
        }
    }
}

impl storage_codec::StorageVisitor for PhysicalCommentStorageCollector {
    fn visit_list_entry(
        &mut self,
        entry: storage_codec::TableDataListEntrySnapshot<'_>,
    ) -> std::result::Result<(), storage_codec::DecodeError> {
        let Some(reference) = entry.comment_storage() else {
            self.invalid_entry = true;
            return Ok(());
        };
        if entry.string_value().is_some()
            || entry.reference().is_some()
            || entry.formula().is_some()
            || entry.format().is_some()
            || entry.custom_format().is_some()
            || entry.rich_text_payload().is_some()
            || entry.import_warning_set().is_some()
            || entry.cell_spec().is_some()
        {
            self.invalid_entry = true;
            return Ok(());
        }
        self.keys
            .try_reserve(1)
            .map_err(|_| storage_codec::DecodeError::allocation(1))?;
        self.identifiers
            .try_reserve(1)
            .map_err(|_| storage_codec::DecodeError::allocation(1))?;
        self.keys.push(entry.key());
        self.identifiers.push(reference.identifier());
        Ok(())
    }

    fn visit_list_segment(
        &mut self,
        _reference: storage_codec::ReferenceRecord<'_>,
    ) -> std::result::Result<(), storage_codec::DecodeError> {
        self.has_segments = true;
        Ok(())
    }
}

pub(crate) fn physical_catalog(package: &Package) -> Result<&SourceCatalog> {
    match &package.state.source {
        PhysicalSource::Package(source) => Ok(source),
        PhysicalSource::Semantic(_) => Err(Error::UnsupportedSource),
    }
}

/// Resolve one parsed physical component by its canonical package-member
/// name.  Keeping this lookup beside [`locate_object`] gives physical owners
/// one authority for component/object routing and prevents a caller from
/// falling back to a second name scan with subtly different semantics.
pub(crate) fn component<'a>(
    package: &'a Package,
    name: &str,
) -> Result<&'a litchi_iwa_archive::Component> {
    physical_catalog(package)?
        .components()
        .get(name)
        .ok_or(Error::InvalidSource)
}

/// Charge the package-entry scan performed by root preview invalidation.
///
/// The rendering helper owns its small fixed preview plan, while this shared
/// boundary owns the residual transaction ledger used by every focused table
/// property owner.
pub(crate) fn charge_preview_scan(package: &Package, budget: &mut Budget) -> Result<()> {
    let catalog = physical_catalog(package)?;
    budget.entries(catalog.package().len())?;
    budget.preflight_preview_scan(catalog.package().len())?;
    for entry in catalog.package().iter() {
        budget.work(entry_comparison_work(entry)?)?;
    }
    Ok(())
}

pub(crate) fn object_at<'a>(
    package: &'a Package,
    location: &ObjectLocation,
) -> Result<&'a ArchiveObject> {
    let index = package
        .state
        .object_index
        .binary_search_by_key(&location.identifier, |locator| locator.identifier)
        .map_err(|_| Error::InvalidSource)?;
    let locator = package.state.object_index[index];
    let component = package
        .state
        .source
        .components()
        .get_index(locator.component)
        .ok_or(Error::InvalidSource)?;
    if component.name() != location.component.as_ref() || locator.object != location.object_index {
        return Err(Error::InvalidSource);
    }
    let object = component
        .archive()
        .objects
        .get(location.object_index)
        .ok_or(Error::InvalidSource)?;
    if object.archive_info.identifier != Some(location.identifier) {
        return Err(Error::InvalidSource);
    }
    Ok(object)
}

/// Locate one physical object with its component name without exposing native
/// object identifiers to semantic callers. The package object index is sorted
/// once at ingress, so this hot path is logarithmic rather than a full scan.
pub(crate) fn object_with_component(
    package: &Package,
    identifier: u64,
) -> Result<(&str, &ArchiveObject)> {
    let index = package
        .state
        .object_index
        .binary_search_by_key(&identifier, |locator| locator.identifier)
        .map_err(|_| Error::InvalidSource)?;
    let locator = package.state.object_index[index];
    let component = package
        .state
        .source
        .components()
        .get_index(locator.component)
        .ok_or(Error::InvalidSource)?;
    let object = component
        .archive()
        .objects
        .get(locator.object)
        .filter(|object| object.archive_info.identifier == Some(identifier))
        .ok_or(Error::InvalidSource)?;
    Ok((component.name(), object))
}

pub(crate) fn model_payload<'a>(package: &'a Package, target: &Target) -> Result<&'a [u8]> {
    let object = object_at(package, &target.model)?;
    validate_message_header(object, target.model_message_index)?;
    object
        .messages
        .get(target.model_message_index)
        .filter(|message| message.type_ == TABLE_MODEL_MESSAGE_TYPE)
        .map(|message| message.data.as_slice())
        .ok_or(Error::InvalidSource)
}

pub(crate) fn component_archive(
    package: &Package,
    name: &str,
    budget: &mut Budget,
) -> Result<Archive> {
    let entry = physical_catalog(package)?
        .package()
        .iter()
        .find(|entry| entry.name() == name)
        .ok_or(Error::InvalidSource)?;
    if entry.is_opaque() {
        return Err(Error::UnsupportedSource);
    }
    // Snappy decompression owns a stream buffer and archive parsing owns
    // message/header vectors before either operation returns. Reserve the
    // conservative source-sized envelope first; the parsed inventory below
    // charges the exact object/message/reference counts as well.
    budget.preflight_archive_parse(entry.data().len())?;
    budget.physical(entry.data().len())?;
    let snappy = package
        .state
        .options
        .archive()
        .snappy_limits()
        .map_err(|_| Error::Archive)?;
    let limits = package
        .state
        .options
        .archive()
        .effective_archive_limits()
        .map_err(|_| Error::Archive)?;
    let stream =
        SnappyStream::decompress_with_limits(entry.data(), snappy).map_err(|_| Error::Archive)?;
    budget.preflight_decoded_archive(stream.as_bytes().len())?;
    budget.physical(stream.as_bytes().len())?;
    let archive =
        Archive::parse_with_limits(stream.as_bytes(), limits).map_err(|_| Error::Archive)?;
    archive
        .validate_canonical_object_framing(stream.as_bytes())
        .map_err(|_| Error::Archive)?;
    charge_archive_inventory(&archive, budget)?;
    Ok(archive)
}

pub(crate) fn verify_locality(
    source: &Package,
    candidate: &Package,
    target: &Target,
    target_previews_absent: bool,
    budget: &mut Budget,
) -> Result<()> {
    let source_catalog = physical_catalog(source)?;
    let candidate_catalog = physical_catalog(candidate)?;
    let source_entry_count = source_catalog.package().len();
    let candidate_entry_count = candidate_catalog.package().len();
    budget.entries(
        source_entry_count
            .checked_add(candidate_entry_count)
            .ok_or(Error::InvalidSource)?,
    )?;
    let entry_map_capacity = source_entry_count
        .checked_add(candidate_entry_count)
        .ok_or(Error::InvalidSource)?;
    budget.allocations(entry_map_capacity)?;
    budget.retained(
        entry_map_capacity
            .checked_mul(size_of::<usize>())
            .ok_or(Error::InvalidSource)?,
    )?;
    let mut source_names = HashSet::new();
    source_names
        .try_reserve(source_entry_count)
        .map_err(|_| Error::Allocation(source_entry_count))?;
    let mut candidate_entries = HashMap::new();
    candidate_entries
        .try_reserve(candidate_entry_count)
        .map_err(|_| Error::Allocation(candidate_entry_count))?;
    for entry in source_catalog.package().iter() {
        budget.work(entry_comparison_work(entry)?)?;
        if !source_names.insert(entry.name()) {
            return Err(Error::Verification);
        }
    }
    for entry in candidate_catalog.package().iter() {
        budget.work(entry_comparison_work(entry)?)?;
        if candidate_entries.insert(entry.name(), entry).is_some() {
            return Err(Error::Verification);
        }
    }
    budget.preflight_preview_scan(source_entry_count)?;
    let previews = super::rendering_invalidation::root_preview_deletions(source_catalog.package())
        .map_err(|_| Error::Verification)?;
    let restored_previews = if target_previews_absent {
        None
    } else {
        budget.preflight_preview_scan(candidate_entry_count)?;
        Some(
            super::rendering_invalidation::root_preview_deletions(candidate_catalog.package())
                .map_err(|_| Error::Verification)?,
        )
    };
    for entry in source_catalog.package().iter() {
        let candidate_entry = candidate_entries.get(entry.name()).copied();
        if target_previews_absent && previews.names().contains(&entry.name()) {
            if candidate_entry.is_some() {
                return Err(Error::Verification);
            }
            continue;
        }
        let other = candidate_entry.ok_or(Error::Verification)?;
        if entry.name() != target.model.component.as_ref()
            && (entry.data() != other.data()
                || entry.metadata() != other.metadata()
                || entry.raw_record().local_record() != other.raw_record().local_record()
                || !same_central_directory_record(
                    entry.raw_record().central_directory_record(),
                    other.raw_record().central_directory_record(),
                ))
        {
            return Err(Error::Verification);
        }
    }
    for name in candidate_entries.keys() {
        if !source_names.contains(name)
            && restored_previews
                .as_ref()
                .is_none_or(|previews| !previews.names().contains(name))
        {
            return Err(Error::Verification);
        }
    }
    let source_archive = component_archive(source, target.model.component.as_ref(), budget)?;
    let candidate_archive = component_archive(candidate, target.model.component.as_ref(), budget)?;
    if source_archive.objects.len() != candidate_archive.objects.len() {
        return Err(Error::Verification);
    }
    let limits = source
        .state
        .options
        .archive()
        .effective_archive_limits()
        .map_err(|_| Error::Archive)?;
    for source_object in &source_archive.objects {
        let identifier = source_object
            .archive_info
            .identifier
            .ok_or(Error::Verification)?;
        let candidate_object = candidate_archive
            .object(identifier)
            .ok_or(Error::Verification)?;
        if identifier == target.model.identifier {
            let candidate_message = candidate_object
                .messages
                .get(target.model_message_index)
                .ok_or(Error::Verification)?;
            let clone_cost = archive_object_cost(source_object)?;
            budget.allocations(clone_cost)?;
            budget.retained(clone_cost)?;
            let mut expected = source_object.clone();
            expected
                .replace_message_preserving_header_with_limits(
                    target.model_message_index,
                    candidate_message.clone(),
                    limits,
                )
                .map_err(|_| Error::Archive)?;
            expected.header_length = candidate_object.header_length;
            expected.data_length = candidate_object.data_length;
            if !expected.same_content_ignoring_offsets(candidate_object) {
                return Err(Error::Verification);
            }
        } else if !source_object.same_content_ignoring_offsets(candidate_object) {
            return Err(Error::Verification);
        }
    }
    Ok(())
}

/// Verify that a physical candidate differs only in the explicitly admitted
/// component members and, optionally, root preview deletions.
///
/// Physical sorting rewrites several objects in one component (the model,
/// tile rows, row headers, and UID map), so the persisted field-44 locality
/// helper above is intentionally too narrow. This helper performs the same
/// source/candidate ZIP checks but accepts a deterministic allowlist owned by
/// the physical transaction.
pub(crate) fn verify_physical_locality(
    source: &Package,
    candidate: &Package,
    allowlist: &LocalityAllowlist,
    budget: &mut Budget,
) -> Result<()> {
    let archive_limits = source
        .state
        .options
        .archive()
        .effective_archive_limits()
        .map_err(|_| Error::Archive)?;
    let source_catalog = physical_catalog(source)?;
    let candidate_catalog = physical_catalog(candidate)?;
    let source_entries = source_catalog.package();
    let candidate_entries = candidate_catalog.package();
    budget.entries(
        source_entries
            .len()
            .checked_add(candidate_entries.len())
            .ok_or(Error::InvalidSource)?,
    )?;
    budget.allocations(
        source_entries
            .len()
            .checked_add(candidate_entries.len())
            .ok_or(Error::InvalidSource)?,
    )?;
    let mut source_names = HashSet::new();
    source_names
        .try_reserve(source_entries.len())
        .map_err(|_| Error::Allocation(source_entries.len()))?;
    let mut candidate_by_name = HashMap::new();
    candidate_by_name
        .try_reserve(candidate_entries.len())
        .map_err(|_| Error::Allocation(candidate_entries.len()))?;
    for entry in source_entries.iter() {
        budget.work(entry_comparison_work(entry)?)?;
        if !source_names.insert(entry.name()) {
            return Err(Error::Verification);
        }
    }
    for entry in candidate_entries.iter() {
        budget.work(entry_comparison_work(entry)?)?;
        if candidate_by_name.insert(entry.name(), entry).is_some() {
            return Err(Error::Verification);
        }
    }

    // A deletion is admissible only when the source renderer identifies it as
    // a root preview. Never let a caller turn this into an arbitrary package
    // member deletion by supplying a hand-written name.
    budget.preflight_preview_scan(source_entries.len())?;
    let source_previews = super::rendering_invalidation::root_preview_deletions(source_entries)
        .map_err(|_| Error::Verification)?;
    if allowlist.root_preview_deletions().iter().any(|name| {
        !source_previews
            .names()
            .iter()
            .any(|preview| *preview == name.as_ref())
    }) {
        return Err(Error::Verification);
    }

    for entry in source_entries.iter() {
        let name = entry.name();
        if allowlist.allows_preview_deletion(name) {
            if candidate_by_name.contains_key(name) {
                return Err(Error::Verification);
            }
            continue;
        }
        let other = candidate_by_name
            .get(name)
            .copied()
            .ok_or(Error::Verification)?;
        if !allowlist.allows_component(name)
            && (entry.data() != other.data()
                || entry.metadata() != other.metadata()
                || entry.raw_record().local_record() != other.raw_record().local_record()
                || !same_central_directory_record(
                    entry.raw_record().central_directory_record(),
                    other.raw_record().central_directory_record(),
                ))
        {
            return Err(Error::Verification);
        }
    }
    for name in candidate_by_name.keys() {
        if !source_names.contains(name) {
            return Err(Error::Verification);
        }
    }

    // Every allowlisted component must exist on both sides. This catches a
    // typo or a non-deterministic component derivation before object-level
    // checks are attempted.
    for component_name in allowlist.changed_components() {
        let name = component_name.as_ref();
        if !source_names.contains(name) || !candidate_by_name.contains_key(name) {
            return Err(Error::Verification);
        }
    }

    // Object identities are a second, optional fence inside an allowlisted
    // component.  Reject an identity routed through another component (or a
    // candidate that relocates it) before comparing any payload bytes.
    for identifier in allowlist.changed_object_ids() {
        let (source_component_name, _) =
            object_with_component(source, *identifier).map_err(|_| Error::Verification)?;
        let (candidate_component_name, _) =
            object_with_component(candidate, *identifier).map_err(|_| Error::Verification)?;
        if source_component_name != candidate_component_name
            || !allowlist.allows_component(source_component_name)
        {
            return Err(Error::Verification);
        }
    }

    for component_name in allowlist.changed_components() {
        let source_component =
            component(source, component_name.as_ref()).map_err(|_| Error::Verification)?;
        let candidate_component =
            component(candidate, component_name.as_ref()).map_err(|_| Error::Verification)?;
        let source_archive = source_component.archive();
        let candidate_archive = candidate_component.archive();
        if source_archive.objects.len() != candidate_archive.objects.len() {
            return Err(Error::Verification);
        }
        for source_object in &source_archive.objects {
            let identifier = source_object
                .archive_info
                .identifier
                .ok_or(Error::Verification)?;
            let candidate_object = candidate_archive
                .object(identifier)
                .ok_or(Error::Verification)?;
            verify_physical_object_delta(
                source_object,
                candidate_object,
                allowlist,
                archive_limits,
                budget,
            )?;
            budget.work(archive_object_cost(source_object)?)?;
            budget.work(archive_object_cost(candidate_object)?)?;
        }
    }
    Ok(())
}

/// Check one object in an allowlisted component without making the component
/// itself an unchecked mutation wildcard.  Physical sorting may alter only
/// the payload bytes of the native tile, header-bucket, and row/column UID-map
/// messages.  Every other message and all object/reference metadata stay
/// source-authoritative, including objects unrelated to the selected table
/// that happen to share the component.
fn verify_physical_object_delta(
    source: &ArchiveObject,
    candidate: &ArchiveObject,
    allowlist: &LocalityAllowlist,
    limits: litchi_iwa_core::Limits,
    budget: &mut Budget,
) -> Result<()> {
    let identifier = source.archive_info.identifier.ok_or(Error::Verification)?;
    if candidate.archive_info.identifier != Some(identifier)
        || source.messages.len() != source.archive_info.message_infos.len()
        || candidate.messages.len() != candidate.archive_info.message_infos.len()
        || source.messages.len() != candidate.messages.len()
        || source.archive_info.should_merge != candidate.archive_info.should_merge
        || source.archive_info.message_infos.len() != candidate.archive_info.message_infos.len()
    {
        return Err(Error::Verification);
    }

    let mut changed_message = false;
    for (index, (source_message, candidate_message)) in
        source.messages.iter().zip(&candidate.messages).enumerate()
    {
        let source_info = source
            .archive_info
            .message_infos
            .get(index)
            .ok_or(Error::Verification)?;
        let candidate_info = candidate
            .archive_info
            .message_infos
            .get(index)
            .ok_or(Error::Verification)?;
        if source_info.type_ != source_message.type_
            || candidate_info.type_ != candidate_message.type_
            || source_info.type_ != candidate_info.type_
            || source_info.versions != candidate_info.versions
            || source_info.field_infos != candidate_info.field_infos
            || source_info.object_references != candidate_info.object_references
            || source_info.data_references != candidate_info.data_references
            || source_info.base_message_index != candidate_info.base_message_index
            || source_info.diff_merge_version != candidate_info.diff_merge_version
            || source_info.diff_field_path != candidate_info.diff_field_path
            || source_info.fields_to_remove != candidate_info.fields_to_remove
            || source_info.diff_read_version != candidate_info.diff_read_version
            || u32::try_from(source_message.data.len()).ok() != Some(source_info.length)
            || u32::try_from(candidate_message.data.len()).ok() != Some(candidate_info.length)
        {
            return Err(Error::Verification);
        }
        if source_message.data == candidate_message.data {
            if source_info.length != candidate_info.length {
                return Err(Error::Verification);
            }
            continue;
        }
        if !PHYSICAL_MUTABLE_MESSAGE_TYPES.contains(&source_message.type_)
            || (!allowlist.changed_object_ids().is_empty() && !allowlist.allows_object(identifier))
        {
            return Err(Error::Verification);
        }
        changed_message = true;
    }

    if changed_message {
        // A changed payload is admissible only for one exact physical role.
        // In particular, a multi-message object cannot smuggle a second
        // message header or an unrelated payload through the component
        // allowlist alongside a rewritten tile/header/UID root.
        if source.messages.len() != 1
            || candidate.messages.len() != 1
            || !PHYSICAL_MUTABLE_MESSAGE_TYPES.contains(&source.messages[0].type_)
        {
            return Err(Error::Verification);
        }

        // Reconstruct the candidate from the source object using the core's
        // header-preserving primitive.  This is a source-bound proof that the
        // candidate changed only the exact payload bytes: retained unknown
        // ArchiveInfo fields, message framing, reference metadata, and raw
        // header bytes must all match the primitive's deterministic rewrite.
        // The two recorded lengths are updated only because a payload length
        // change legitimately changes the enclosing physical framing.
        let mut expected = source.clone();
        for (index, (source_message, candidate_message)) in
            source.messages.iter().zip(&candidate.messages).enumerate()
        {
            if source_message.data != candidate_message.data {
                expected
                    .replace_message_preserving_header_with_limits(
                        index,
                        candidate_message.clone(),
                        limits,
                    )
                    .map_err(|_| Error::Verification)?;
            }
        }

        // `header_length` and `data_length` are parsed source framing facts,
        // not free-form provenance fields.  Derive both values from the
        // source-preserving replacement and compare them with the candidate
        // before copying anything into the expected object.  In particular,
        // this rejects a candidate that changes the object-length varint
        // width, retains a stale payload length, or otherwise moves the
        // physical object boundary outside the exact rewritten message.
        let ((expected_header_length, expected_data_length), mut expected) =
            physical_object_framing(expected, limits, budget)?;
        if candidate.header_length != expected_header_length
            || candidate.data_length != expected_data_length
        {
            return Err(Error::Verification);
        }
        expected.header_length = expected_header_length;
        expected.data_length = expected_data_length;
        if !expected.same_content_ignoring_offsets(candidate) {
            return Err(Error::Verification);
        }
    } else if !source.same_content_ignoring_offsets(candidate) {
        // A candidate object that retained identical payloads must still be
        // byte/content-identical apart from offsets.
        return Err(Error::Verification);
    }
    Ok(())
}

/// Return the exact serialized object framing implied by its source-preserved
/// ArchiveInfo header and current message payloads.
///
/// `Archive::encoded_len_with_limits` follows the same retained-header rule
/// used by publication, including the canonical object-length varint.
/// Subtracting the checked payload total therefore exposes the exact framed
/// header length without accessing ArchiveObject's private raw-header fields.
fn physical_object_framing(
    object: ArchiveObject,
    limits: litchi_iwa_core::Limits,
    budget: &mut Budget,
) -> Result<((u64, u64), ArchiveObject)> {
    let payload_length = object.messages.iter().try_fold(0usize, |total, message| {
        total
            .checked_add(message.data.len())
            .ok_or(Error::Verification)
    })?;
    budget.allocations(1)?;
    budget.retained(size_of::<ArchiveObject>())?;
    let mut archive = Archive::new();
    archive
        .objects
        .try_reserve_exact(1)
        .map_err(|_| Error::Allocation(1))?;
    archive.objects.push(object);
    let encoded_length = archive
        .encoded_len_with_limits(limits)
        .map_err(|_| Error::Verification)?;
    let header_length = encoded_length
        .checked_sub(payload_length)
        .ok_or(Error::Verification)?;
    let object = archive.objects.pop().ok_or(Error::Verification)?;
    Ok((
        (
            u64::try_from(header_length).map_err(|_| Error::Verification)?,
            u64::try_from(payload_length).map_err(|_| Error::Verification)?,
        ),
        object,
    ))
}

fn entry_comparison_work(entry: &litchi_iwa_archive::package::Entry) -> Result<usize> {
    entry
        .name()
        .len()
        .checked_add(entry.data().len())
        .and_then(|size| size.checked_add(entry.raw_record().local_record().len()))
        .and_then(|size| size.checked_add(entry.raw_record().central_directory_record().len()))
        .ok_or(Error::InvalidSource)
}

fn same_central_directory_record(left: &[u8], right: &[u8]) -> bool {
    const LOCAL_HEADER_OFFSET: std::ops::Range<usize> = 42..46;

    left.len() == right.len()
        && left.len() >= LOCAL_HEADER_OFFSET.end
        && left[..LOCAL_HEADER_OFFSET.start] == right[..LOCAL_HEADER_OFFSET.start]
        && left[LOCAL_HEADER_OFFSET.end..] == right[LOCAL_HEADER_OFFSET.end..]
}

fn archive_object_cost(object: &ArchiveObject) -> Result<usize> {
    let mut cost = object.messages.len();
    cost = cost
        .checked_add(object.archive_info.message_infos.len())
        .ok_or(Error::InvalidSource)?;
    for message in &object.messages {
        cost = cost
            .checked_add(message.data.len())
            .ok_or(Error::InvalidSource)?;
    }
    for info in &object.archive_info.message_infos {
        cost = cost
            .checked_add(
                info.object_references
                    .len()
                    .checked_add(info.data_references.len())
                    .and_then(|size| size.checked_add(info.field_infos.len()))
                    .ok_or(Error::InvalidSource)?,
            )
            .ok_or(Error::InvalidSource)?;
        for field in &info.field_infos {
            cost = cost
                .checked_add(field.path.path.len())
                .and_then(|size| size.checked_add(field.object_references.len()))
                .and_then(|size| size.checked_add(field.data_references.len()))
                .ok_or(Error::InvalidSource)?;
        }
    }
    Ok(cost)
}

fn resolve_slide_position(
    package: &Package,
    selector: SlideSelector<'_>,
    budget: &mut Budget,
) -> Result<Position> {
    match selector {
        SlideSelector::Position(position) => Ok(position),
        SlideSelector::Name(name) => {
            if name.is_empty() {
                return Err(Error::EmptySlideName);
            }
            let show = package.show().map_err(read_error)?;
            budget.work(
                show.slide_count()
                    .checked_add(name.len())
                    .ok_or(Error::InvalidSource)?,
            )?;
            // A duplicate-name diagnostic owns a copy of the requested name.
            // Charge that possible allocation before delegating selection;
            // successful selectors simply leave the conservative charge in
            // the operation ledger.
            budget.allocations(name.len().max(1))?;
            show.select_slide(selector)
                .map_err(|_| Error::AmbiguousSelector)?
                .map(|slide| Position::new(slide.index()))
                .ok_or(Error::SlideNameNotFound)
        },
    }
}

fn resolve_slide_position_from_components(
    package: &Package,
    selector: SlideSelector<'_>,
    budget: &mut Budget,
) -> Result<Position> {
    match selector {
        SlideSelector::Position(position) => Ok(position),
        SlideSelector::Name(name) => {
            if name.is_empty() {
                return Err(Error::EmptySlideName);
            }
            // Name diagnostics retain only the requested selector, while the
            // candidate names remain borrowed directly from their IWA payload.
            budget.allocations(name.len().max(1))?;
            let maximum = package.semantic_limits().max_slides();
            let mut semantic_budget = SemanticBudget::new(package.semantic_limits());
            let mut selected = None;
            for index in 0..maximum {
                let references_before = semantic_budget.references_charged();
                let projection = package
                    .slide_record_at_with_budget(index, &mut semantic_budget)
                    .map_err(read_error)?;
                let references = semantic_budget
                    .references_charged()
                    .checked_sub(references_before)
                    .ok_or(Error::InvalidSource)?;
                budget.references(references)?;
                budget.work(projection.work)?;
                let Some(record) = projection.record else {
                    break;
                };
                let (candidate, candidate_work) = package
                    .slide_name_for_identifier_with_work(record.slide_identifier, index)
                    .map_err(read_error)?;
                budget.work(candidate_work)?;
                let Some(candidate) = candidate else {
                    continue;
                };
                budget.work(candidate.len().checked_add(1).ok_or(Error::InvalidSource)?)?;
                if candidate == name {
                    if selected.replace(Position::new(index)).is_some() {
                        return Err(Error::AmbiguousSelector);
                    }
                }
            }
            selected.ok_or(Error::SlideNameNotFound)
        },
    }
}

fn read_error(error: ReadError) -> Error {
    match error {
        ReadError::SemanticLimit {
            kind,
            observed,
            maximum,
            ..
        } => Error::Limit {
            kind: match kind {
                SemanticLimitKind::Objects => LimitKind::PayloadObjects,
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
        ReadError::Allocation { amount, .. } => Error::Allocation(amount),
        ReadError::Archive(error) => match error {
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
            litchi_iwa_archive::Error::Allocation { amount, .. } => Error::Allocation(amount),
            _ => Error::Read,
        },
        _ => Error::Read,
    }
}

pub(crate) fn locate_object(package: &Package, identifier: u64) -> Result<ObjectLocation> {
    let locator = package
        .state
        .object_index
        .binary_search_by_key(&identifier, |locator| locator.identifier)
        .ok()
        .map(|index| package.state.object_index[index])
        .ok_or(Error::InvalidSource)?;
    let component = package
        .state
        .source
        .components()
        .get_index(locator.component)
        .ok_or(Error::InvalidSource)?;
    let object = component
        .archive()
        .objects
        .get(locator.object)
        .ok_or(Error::InvalidSource)?;
    if object.archive_info.identifier != Some(identifier) {
        return Err(Error::InvalidSource);
    }
    Ok(ObjectLocation {
        identifier,
        component: Arc::from(component.name()),
        object_index: locator.object,
    })
}

pub(crate) fn unique_message<'a>(
    object: &'a ArchiveObject,
    message_type: u32,
    budget: &mut Budget,
) -> Result<(usize, &'a [u8])> {
    if object.messages.len() != object.archive_info.message_infos.len() {
        return Err(Error::InvalidSource);
    }
    let mut selected = None;
    for (index, message) in object.messages.iter().enumerate() {
        budget.work(
            message
                .data
                .len()
                .checked_add(1)
                .ok_or(Error::InvalidSource)?,
        )?;
        validate_message_header(object, index)?;
        if ROLE_MESSAGE_TYPES.contains(&message.type_) && message.type_ != message_type {
            return Err(Error::UnsupportedDependency);
        }
        if message.type_ == message_type
            && selected.replace((index, message.data.as_slice())).is_some()
        {
            return Err(Error::UnsupportedDependency);
        }
    }
    selected.ok_or(Error::InvalidSource)
}

/// Strictly locate one message whose type is in a small compatibility set.
/// This is used for type-renumbered native UID maps and data-list roots while
/// retaining the same canonical message-header and role-alias proof as the
/// focused persisted owners.
pub(crate) fn unique_message_any<'a>(
    object: &'a ArchiveObject,
    message_types: &[u32],
    budget: &mut Budget,
) -> Result<(usize, &'a [u8])> {
    if object.messages.len() != object.archive_info.message_infos.len() {
        return Err(Error::InvalidSource);
    }
    let mut selected = None;
    for (index, message) in object.messages.iter().enumerate() {
        budget.work(
            message
                .data
                .len()
                .checked_add(1)
                .ok_or(Error::InvalidSource)?,
        )?;
        validate_message_header(object, index)?;
        if ROLE_MESSAGE_TYPES.contains(&message.type_) && !message_types.contains(&message.type_) {
            return Err(Error::UnsupportedDependency);
        }
        if message_types.contains(&message.type_)
            && selected.replace((index, message.data.as_slice())).is_some()
        {
            return Err(Error::UnsupportedDependency);
        }
    }
    selected.ok_or(Error::InvalidSource)
}

/// Locate one mutable physical-root message and reject any co-located
/// payload, including an otherwise-unknown future role.  Physical sorting
/// rewrites these objects in place, so preserving an additional message
/// beside the selected role would make the mutation boundary ambiguous.
fn unique_mutable_message<'a>(
    object: &'a ArchiveObject,
    message_types: &[u32],
    budget: &mut Budget,
) -> Result<(usize, &'a [u8])> {
    if !has_exact_mutable_message_shape(object, message_types) {
        return Err(Error::UnsupportedDependency);
    }
    let (index, payload) = unique_message_any(object, message_types, budget)?;
    if index != 0 {
        return Err(Error::UnsupportedDependency);
    }
    Ok((index, payload))
}

fn has_exact_mutable_message_shape(object: &ArchiveObject, message_types: &[u32]) -> bool {
    object.messages.len() == 1
        && object.archive_info.message_infos.len() == 1
        && object
            .messages
            .first()
            .is_some_and(|message| message_types.contains(&message.type_))
}

fn validate_message_header(object: &ArchiveObject, index: usize) -> Result<()> {
    let message = object.messages.get(index).ok_or(Error::InvalidSource)?;
    let info = object
        .archive_info
        .message_infos
        .get(index)
        .ok_or(Error::InvalidSource)?;
    if message.type_ != info.type_
        || usize::try_from(info.length).ok() != Some(message.data.len())
        || object.archive_info.should_merge == Some(true)
        || info.base_message_index.is_some()
        || !info.diff_merge_version.is_empty()
        || info.diff_field_path.is_some()
        || !info.fields_to_remove.is_empty()
        || !info.diff_read_version.is_empty()
    {
        return Err(Error::InvalidSource);
    }
    Ok(())
}

fn repeated_references(
    payload: &[u8],
    number: u32,
    limits: WireLimits,
    budget: &mut Budget,
) -> Result<Vec<u64>> {
    budget.input(payload.len())?;
    let fields = WireView::parse_with_limits(payload, limits).map_err(|_| Error::Wire)?;
    let all_fields = fields.fields().count();
    let count = fields
        .fields()
        .filter(|field| field.number() == number)
        .count();
    budget.fields(all_fields)?;
    budget.work(payload.len())?;
    budget.allocations(count)?;
    budget.references(count)?;
    let mut values = Vec::new();
    values
        .try_reserve_exact(count)
        .map_err(|_| Error::Allocation(count))?;
    for field in fields.fields().filter(|field| field.number() == number) {
        field
            .validate_canonical_framing()
            .map_err(|_| Error::Wire)?;
        if field.wire_type() != 2 {
            return Err(Error::InvalidSource);
        }
        values.push(strict_reference(field.payload(), limits, budget)?);
    }
    Ok(values)
}

fn table_parent(payload: &[u8], limits: WireLimits, budget: &mut Budget) -> Result<u64> {
    budget.input(payload.len())?;
    let fields = WireView::parse_with_limits(payload, limits).map_err(|_| Error::Wire)?;
    let mut super_field = None;
    for field in fields.fields() {
        budget.fields(1)?;
        field
            .validate_canonical_framing()
            .map_err(|_| Error::Wire)?;
        if field.number() == TABLE_SUPER_FIELD {
            if super_field.replace(field).is_some() || field.wire_type() != 2 {
                return Err(Error::InvalidSource);
            }
        }
    }
    let super_field = super_field.ok_or(Error::InvalidSource)?;
    let drawable =
        WireView::parse_with_limits(super_field.payload(), limits).map_err(|_| Error::Wire)?;
    let mut parent_field = None;
    for field in drawable.fields() {
        budget.fields(1)?;
        field
            .validate_canonical_framing()
            .map_err(|_| Error::Wire)?;
        if field.number() == DRAWABLE_PARENT_FIELD {
            if parent_field.replace(field).is_some() || field.wire_type() != 2 {
                return Err(Error::InvalidSource);
            }
        }
    }
    let parent_field = parent_field.ok_or(Error::InvalidSource)?;
    strict_reference(parent_field.payload(), limits, budget)
}

fn strict_reference(payload: &[u8], limits: WireLimits, budget: &mut Budget) -> Result<u64> {
    let fields = WireView::parse_with_limits(payload, limits).map_err(|_| Error::Wire)?;
    let mut identifier = None;
    for field in fields.fields() {
        budget.fields(1)?;
        field
            .validate_canonical_framing()
            .map_err(|_| Error::Wire)?;
        match field.number() {
            1 => {
                if field.wire_type() != 0 || identifier.is_some() {
                    return Err(Error::InvalidSource);
                }
                let (value, width) =
                    decode_varint_from_bytes(field.payload()).map_err(|_| Error::InvalidSource)?;
                if value == 0 || width != encoded_len(value) {
                    return Err(Error::InvalidSource);
                }
                identifier = Some(value);
            },
            2 | 3 => return Err(Error::UnsupportedDependency),
            _ => return Err(Error::InvalidSource),
        }
    }
    identifier.ok_or(Error::InvalidSource)
}

fn reject_duplicates(values: &[u64], budget: &mut Budget) -> Result<()> {
    budget.allocations(values.len())?;
    budget.work(values.len())?;
    let mut seen = HashSet::new();
    seen.try_reserve(values.len())
        .map_err(|_| Error::Allocation(values.len()))?;
    for value in values {
        if !seen.insert(*value) {
            return Err(Error::InvalidSource);
        }
    }
    Ok(())
}

fn validate_slide_metadata(
    object: &ArchiveObject,
    message_index: usize,
    owned: &[u64],
    z_order: &[u64],
    budget: &mut Budget,
) -> Result<()> {
    validate_message_header(object, message_index)?;
    let info = object
        .archive_info
        .message_infos
        .get(message_index)
        .ok_or(Error::InvalidSource)?;
    if info.object_references.is_empty() {
        return Err(Error::InvalidSource);
    }
    if !info.data_references.is_empty() {
        return Err(Error::UnsupportedDependency);
    }
    reject_duplicates(&info.object_references, budget)?;
    budget.fields(info.field_infos.len())?;
    budget.references(
        info.object_references
            .len()
            .checked_add(info.data_references.len())
            .ok_or(Error::InvalidSource)?,
    )?;
    let expected_capacity = owned
        .len()
        .checked_add(z_order.len())
        .ok_or(Error::InvalidSource)?;
    budget.allocations(expected_capacity)?;
    let mut expected = HashMap::new();
    expected
        .try_reserve(expected_capacity)
        .map_err(|_| Error::Allocation(expected_capacity))?;
    for identifier in owned.iter().chain(z_order) {
        expected.insert(*identifier, 0usize);
    }
    let mut saw_owned = false;
    let mut saw_z_order = false;
    for field in &info.field_infos {
        budget.references(
            field
                .object_references
                .len()
                .checked_add(field.data_references.len())
                .ok_or(Error::InvalidSource)?,
        )?;
        if field.path.as_slice() == [SLIDE_OWNED_DRAWABLES_FIELD] {
            if saw_owned
                || !is_message_reference_field(field)
                || !field.data_references.is_empty()
                || field.object_references.as_slice() != owned
            {
                return Err(Error::InvalidSource);
            }
            saw_owned = true;
        } else if field.path.as_slice() == [SLIDE_Z_ORDER_FIELD] {
            if saw_z_order
                || !is_message_reference_field(field)
                || !field.data_references.is_empty()
                || field.object_references.as_slice() != z_order
            {
                return Err(Error::InvalidSource);
            }
            saw_z_order = true;
        } else if !field.object_references.is_empty() || !field.data_references.is_empty() {
            return Err(Error::UnsupportedTopology);
        }
    }
    if saw_owned != saw_z_order {
        return Err(Error::InvalidSource);
    }
    for identifier in info.object_references.iter() {
        if let Some(count) = expected.get_mut(identifier) {
            *count = count.checked_add(1).ok_or(Error::InvalidSource)?;
        }
    }
    if owned.iter().chain(z_order).any(|identifier| {
        info.object_references
            .iter()
            .filter(|id| *id == identifier)
            .count()
            != 1
    }) {
        return Err(Error::InvalidSource);
    }
    for identifier in owned.iter().chain(z_order) {
        if let Some(count) = expected.get_mut(identifier) {
            *count = count.checked_add(1).ok_or(Error::InvalidSource)?;
        }
    }
    if expected.values().any(|count| *count == 0 || *count > 3) {
        return Err(Error::InvalidSource);
    }
    Ok(())
}

fn validate_table_info_metadata(
    package: &Package,
    object: &ArchiveObject,
    message_index: usize,
    model: u64,
    budget: &mut Budget,
) -> Result<()> {
    validate_message_header(object, message_index)?;
    let info = object
        .archive_info
        .message_infos
        .get(message_index)
        .ok_or(Error::InvalidSource)?;
    if !info.data_references.is_empty() {
        return Err(Error::UnsupportedDependency);
    }
    budget.references(info.object_references.len())?;
    budget.fields(info.field_infos.len())?;
    budget.allocations(info.object_references.len())?;
    budget.retained(
        info.object_references
            .len()
            .checked_mul(size_of::<u64>())
            .ok_or(Error::InvalidSource)?,
    )?;
    budget.work(info.object_references.len())?;
    let mut aggregate_references = HashSet::new();
    aggregate_references
        .try_reserve(info.object_references.len())
        .map_err(|_| Error::Allocation(info.object_references.len()))?;
    for identifier in &info.object_references {
        if !aggregate_references.insert(*identifier) {
            return Err(Error::InvalidSource);
        }
    }
    if !aggregate_references.contains(&model) {
        return Err(Error::InvalidSource);
    }
    // Native producers may include caption, title, summary, and category
    // objects beside the selected model in this aggregate. Focused edits
    // preserve those edges, but every one must resolve physically.
    for identifier in &info.object_references {
        locate_object(package, *identifier)?;
    }
    let mut saw_model = false;
    for field in &info.field_infos {
        budget.references(
            field
                .object_references
                .len()
                .checked_add(field.data_references.len())
                .ok_or(Error::InvalidSource)?,
        )?;
        if !field.data_references.is_empty() {
            return Err(Error::UnsupportedDependency);
        }
        if field.object_references.is_empty() {
            continue;
        }
        if !is_message_reference_field(field)
            || field
                .object_references
                .iter()
                .any(|identifier| !aggregate_references.contains(identifier))
        {
            return Err(Error::UnsupportedDependency);
        }
        if field.object_references.contains(&model) {
            if saw_model
                || field.path.as_slice() != [TABLE_MODEL_FIELD]
                || field.object_references.as_slice() != [model]
            {
                return Err(Error::InvalidSource);
            }
            saw_model = true;
        } else if field.path.as_slice() == [TABLE_MODEL_FIELD] {
            return Err(Error::InvalidSource);
        }
    }
    // Native producers may omit the field-local model route while retaining
    // the exact aggregate edge. If present, the loop above proves that the
    // field-local route is unique and canonical; the aggregate remains the
    // mandatory authority in both producer shapes.
    Ok(())
}

fn is_message_reference_field(field: &litchi_iwa_core::FieldInfo) -> bool {
    field
        .r#type
        .is_none_or(|kind| kind == FieldType::ObjectReference)
}

fn decode_table_info(
    payload: &[u8],
    package: &Package,
    budget: &mut Budget,
) -> Result<table_info_codec::TableInfoSnapshot> {
    let limits = budget.residual(package)?;
    budget.input(payload.len())?;
    let options = table_info_codec::DecodeOptions::new(
        payload.len().max(1).min(limits.max_input_bytes()),
        limits.max_fields(),
        limits.max_rewrite_work(),
        u32::try_from(limits.max_nesting()).unwrap_or(u32::MAX),
    )
    .with_max_output_bytes(limits.max_output_bytes())
    .with_max_allocations(budget.remaining_allocations()?)
    .with_max_retained_bytes(budget.remaining_retained()?)
    .with_max_scratch_bytes(budget.remaining_scratch()?);
    let snapshot =
        table_info_codec::decode_table_info(payload, options).map_err(|_| Error::Codec)?;
    let view = WireView::parse_with_limits(payload, limits).map_err(|_| Error::Wire)?;
    budget.fields(view.fields().count())?;
    budget.work(payload.len())?;
    Ok(snapshot)
}

fn validate_model_payload(
    payload: &[u8],
    package: &Package,
    budget: &mut Budget,
) -> Result<(u32, u32)> {
    let options = budget.model_codec_options(package, payload)?;
    let (snapshot, report) =
        table_model_discovery_codec::decode_table_model_with_report(payload, options)
            .map_err(|_| Error::Codec)?;
    budget.codec_report(report)?;
    Ok((snapshot.rows(), snapshot.columns()))
}

struct RowStorageCollector {
    identifiers: Vec<u64>,
}

impl table_dimension_codec::StorageVisitor for RowStorageCollector {
    fn visit_header_bucket(
        &mut self,
        reference: table_dimension_codec::ReferenceRecord<'_>,
    ) -> std::result::Result<(), table_dimension_codec::DecodeError> {
        self.identifiers
            .try_reserve(1)
            .map_err(|_| table_dimension_codec::DecodeError::allocation(1))?;
        self.identifiers.push(reference.reference().identifier());
        Ok(())
    }
}

fn validate_model_storage(
    package: &Package,
    model_identifier: u64,
    model_message_index: usize,
    payload: &[u8],
    budget: &mut Budget,
) -> Result<Vec<StorageRoute>> {
    let model_options = budget.storage_codec_options(package)?;
    let (model, model_report) =
        table_dimension_codec::decode_table_model_with_report(payload, model_options)
            .map_err(|_| Error::Codec)?;
    budget.storage_codec_report(model_report)?;
    let store_options = budget.storage_codec_options(package)?;
    let (store, store_report) = table_dimension_codec::decode_data_store_with_report(
        model.base_data_store(),
        store_options,
    )
    .map_err(|_| Error::Codec)?;
    budget.storage_codec_report(store_report)?;

    let row_limits = budget.residual(package)?;
    let row_view =
        WireView::parse_with_limits(store.row_headers(), row_limits).map_err(|_| Error::Wire)?;
    let row_reference_count = row_view
        .fields()
        .filter(|field| field.number() == 2)
        .count();
    budget.fields(row_view.fields().count())?;
    budget.work(store.row_headers().len())?;
    budget.allocations(row_reference_count)?;
    budget.retained(
        row_reference_count
            .checked_mul(size_of::<u64>())
            .ok_or(Error::InvalidSource)?,
    )?;
    let mut row_identifiers = Vec::new();
    row_identifiers
        .try_reserve_exact(row_reference_count)
        .map_err(|_| Error::Allocation(row_reference_count))?;
    let mut rows = RowStorageCollector {
        identifiers: row_identifiers,
    };
    let row_options = budget.storage_codec_options(package)?;
    let (_, row_report) = table_dimension_codec::decode_header_storage_with_visitor(
        store.row_headers(),
        row_options,
        &mut rows,
    )
    .map_err(|_| Error::Codec)?;
    budget.storage_codec_report(row_report)?;
    if rows.identifiers.is_empty() || rows.identifiers.len() != row_reference_count {
        return Err(Error::UnsupportedDependency);
    }

    let route_count = rows
        .identifiers
        .len()
        .checked_add(5)
        .ok_or(Error::InvalidSource)?;
    budget.allocations(route_count.checked_mul(2).ok_or(Error::InvalidSource)?)?;
    budget.retained(
        route_count
            .checked_mul(size_of::<StorageRoute>())
            .ok_or(Error::InvalidSource)?,
    )?;
    let mut routes = Vec::new();
    routes
        .try_reserve_exact(route_count)
        .map_err(|_| Error::Allocation(route_count))?;
    let mut seen = HashSet::new();
    seen.try_reserve(route_count)
        .map_err(|_| Error::Allocation(route_count))?;
    let mut push_route = |identifier: u64, kind: StorageRouteKind| -> Result<()> {
        if identifier == 0 || !seen.insert(identifier) {
            return Err(Error::UnsupportedDependency);
        }
        let location = locate_object(package, identifier)?;
        ensure_unique_identity(package, identifier, budget)?;
        if kind.is_bucket() {
            let object = object_at(package, &location)?;
            let (message_index, _) = unique_message(object, HEADER_BUCKET_MESSAGE_TYPE, budget)?;
            let info = object
                .archive_info
                .message_infos
                .get(message_index)
                .ok_or(Error::InvalidSource)?;
            if !info.object_references.is_empty()
                || !info.data_references.is_empty()
                || !info.field_infos.is_empty()
            {
                return Err(Error::UnsupportedDependency);
            }
        } else {
            let object = object_at(package, &location)?;
            if object.messages.len() != object.archive_info.message_infos.len() {
                return Err(Error::InvalidSource);
            }
            budget.payload_messages(object.messages.len())?;
            for index in 0..object.messages.len() {
                budget.work(
                    object.messages[index]
                        .data
                        .len()
                        .checked_add(1)
                        .ok_or(Error::InvalidSource)?,
                )?;
                validate_message_header(object, index)?;
                if ROLE_MESSAGE_TYPES.contains(&object.messages[index].type_) {
                    return Err(Error::UnsupportedDependency);
                }
            }
        }
        routes.push(StorageRoute { location, kind });
        Ok(())
    };
    for identifier in rows.identifiers {
        push_route(identifier, StorageRouteKind::RowBucket)?;
    }
    push_route(
        store.column_headers().identifier(),
        StorageRouteKind::ColumnBucket,
    )?;
    push_route(
        store.string_table().identifier(),
        StorageRouteKind::StringTable,
    )?;
    push_route(
        store.style_table().identifier(),
        StorageRouteKind::StyleTable,
    )?;
    push_route(
        store.formula_table().identifier(),
        StorageRouteKind::FormulaTable,
    )?;
    push_route(
        store.format_table_pre_bnc().identifier(),
        StorageRouteKind::FormatTablePreBnc,
    )?;

    validate_model_storage_metadata(
        package,
        model_identifier,
        model_message_index,
        &routes,
        budget,
    )?;
    Ok(routes)
}

fn validate_model_storage_metadata(
    package: &Package,
    model_identifier: u64,
    model_message_index: usize,
    routes: &[StorageRoute],
    budget: &mut Budget,
) -> Result<()> {
    let model = object_at(package, &locate_object(package, model_identifier)?)?;
    let info = model
        .archive_info
        .message_infos
        .get(model_message_index)
        .ok_or(Error::InvalidSource)?;
    budget.fields(info.field_infos.len())?;
    let nested_references = info.field_infos.iter().try_fold(0usize, |sum, field| {
        sum.checked_add(field.object_references.len())
            .and_then(|value| value.checked_add(field.data_references.len()))
            .ok_or(Error::InvalidSource)
    })?;
    budget.references(
        info.object_references
            .len()
            .checked_add(info.data_references.len())
            .and_then(|value| value.checked_add(nested_references))
            .ok_or(Error::InvalidSource)?,
    )?;
    if !info.data_references.is_empty() {
        return Err(Error::UnsupportedDependency);
    }
    budget.allocations(info.object_references.len())?;
    budget.retained(
        info.object_references
            .len()
            .checked_mul(size_of::<u64>())
            .ok_or(Error::InvalidSource)?,
    )?;
    budget.work(info.object_references.len())?;
    let mut aggregate_references = HashSet::new();
    aggregate_references
        .try_reserve(info.object_references.len())
        .map_err(|_| Error::Allocation(info.object_references.len()))?;
    for identifier in &info.object_references {
        if !aggregate_references.insert(*identifier) {
            return Err(Error::UnsupportedDependency);
        }
    }
    for route in routes {
        if !aggregate_references.contains(&route.location.identifier) {
            return Err(Error::UnsupportedDependency);
        }
    }
    // References unrelated to the selected storage spine are preserved by
    // focused edits. Prove that every such aggregate edge resolves to one
    // physical object so an opaque/dangling dependency cannot hide there.
    for identifier in &info.object_references {
        if routes
            .iter()
            .all(|route| route.location.identifier != *identifier)
        {
            locate_object(package, *identifier)?;
        }
    }
    budget.allocations(routes.len())?;
    budget.retained(
        routes
            .len()
            .checked_mul(size_of::<u64>())
            .ok_or(Error::InvalidSource)?,
    )?;
    let mut seen_fields = HashSet::new();
    seen_fields
        .try_reserve(routes.len())
        .map_err(|_| Error::Allocation(routes.len()))?;
    for field in &info.field_infos {
        if !field.data_references.is_empty() {
            return Err(Error::UnsupportedDependency);
        }
        if field.object_references.is_empty() {
            continue;
        }
        if field
            .object_references
            .iter()
            .any(|identifier| !aggregate_references.contains(identifier))
        {
            return Err(Error::UnsupportedDependency);
        }
        let mut selected_route = None;
        for route in routes {
            if field.object_references.contains(&route.location.identifier)
                && selected_route.replace(route).is_some()
            {
                return Err(Error::UnsupportedDependency);
            }
        }
        let Some(route) = selected_route else {
            if !field
                .r#type
                .is_none_or(|kind| matches!(kind, FieldType::ObjectReference | FieldType::Message))
            {
                return Err(Error::UnsupportedDependency);
            }
            continue;
        };
        let [identifier] = field.object_references.as_slice() else {
            return Err(Error::UnsupportedDependency);
        };
        if !is_message_reference_field(field)
            || field.path.as_slice() != route.kind.path()
            || !seen_fields.insert(*identifier)
        {
            return Err(Error::UnsupportedDependency);
        }
    }
    if !seen_fields.is_empty()
        && routes
            .iter()
            .any(|route| !seen_fields.contains(&route.location.identifier))
    {
        return Err(Error::UnsupportedDependency);
    }
    Ok(())
}

pub(crate) fn ensure_unique_identity(
    package: &Package,
    identifier: u64,
    budget: &mut Budget,
) -> Result<()> {
    let mut count = 0usize;
    for component in package.state.source.components().iter() {
        budget.components(1)?;
        for object in &component.archive().objects {
            budget.payload_objects(1)?;
            budget.work(1)?;
            if object.archive_info.identifier == Some(identifier) {
                count = count.checked_add(1).ok_or(Error::InvalidSource)?;
            }
        }
    }
    if count == 1 {
        Ok(())
    } else {
        Err(Error::UnsupportedDependency)
    }
}

fn ensure_unique_table_owner(
    package: &Package,
    slide_identifier: u64,
    table_info_identifier: u64,
    model_identifier: u64,
    limits: WireLimits,
    budget: &mut Budget,
) -> Result<()> {
    let mut owned = 0usize;
    let mut z_order = 0usize;
    let mut selected_slide = false;
    let mut model_owners = 0usize;
    for component in package.state.source.components().iter() {
        budget.components(1)?;
        for object in &component.archive().objects {
            budget.payload_objects(1)?;
            for message in &object.messages {
                budget.payload_messages(1)?;
                budget.work(
                    message
                        .data
                        .len()
                        .checked_add(1)
                        .ok_or(Error::InvalidSource)?,
                )?;
                if message.type_ == SLIDE_MESSAGE_TYPE {
                    let owned_refs = repeated_references(
                        &message.data,
                        SLIDE_OWNED_DRAWABLES_FIELD,
                        limits,
                        budget,
                    )?;
                    let z_refs =
                        repeated_references(&message.data, SLIDE_Z_ORDER_FIELD, limits, budget)?;
                    let owned_hits = owned_refs
                        .iter()
                        .filter(|id| **id == table_info_identifier)
                        .count();
                    let z_hits = z_refs
                        .iter()
                        .filter(|id| **id == table_info_identifier)
                        .count();
                    owned = owned.checked_add(owned_hits).ok_or(Error::InvalidSource)?;
                    z_order = z_order.checked_add(z_hits).ok_or(Error::InvalidSource)?;
                    if object.archive_info.identifier == Some(slide_identifier)
                        && owned_hits == 1
                        && z_hits == 1
                    {
                        selected_slide = true;
                    }
                } else if message.type_ == TABLE_INFO_MESSAGE_TYPE {
                    let info = decode_table_info(&message.data, package, budget)?;
                    if info.table_model().identifier().get() == model_identifier {
                        model_owners = model_owners.checked_add(1).ok_or(Error::InvalidSource)?;
                    }
                }
            }
        }
    }
    if owned == 1 && z_order == 1 && selected_slide && model_owners == 1 {
        Ok(())
    } else {
        Err(Error::UnsupportedDependency)
    }
}

#[derive(Clone, Copy)]
struct MetadataTarget<'a> {
    identifier: u64,
    component_name: &'a str,
    requires_uuid: bool,
    storage: bool,
}

fn validate_package_metadata(
    package: &Package,
    targets: [&ObjectLocation; 3],
    storage: &[StorageRoute],
    budget: &mut Budget,
) -> Result<()> {
    let physical_capacity = package
        .state
        .source
        .components()
        .iter()
        .map(|component| component.archive().objects.len())
        .try_fold(0usize, |sum, count| sum.checked_add(count))
        .ok_or(Error::InvalidSource)?;
    budget.allocations(physical_capacity)?;
    budget.retained(
        physical_capacity
            .checked_mul(size_of::<u64>())
            .ok_or(Error::InvalidSource)?,
    )?;
    budget.work(physical_capacity)?;
    let mut physical_identifiers = HashSet::new();
    physical_identifiers
        .try_reserve(physical_capacity)
        .map_err(|_| Error::Allocation(physical_capacity))?;
    let mut physical_maximum = 0u64;
    let mut data_reference_capacity = 0usize;
    for component in package.state.source.components().iter() {
        budget.components(1)?;
        for object in &component.archive().objects {
            budget.payload_objects(1)?;
            let identifier = object.archive_info.identifier.ok_or(Error::InvalidSource)?;
            if !physical_identifiers.insert(identifier) {
                return Err(Error::UnsupportedDependency);
            }
            physical_maximum = physical_maximum.max(identifier);
            for info in &object.archive_info.message_infos {
                data_reference_capacity = data_reference_capacity
                    .checked_add(info.data_references.len())
                    .ok_or(Error::InvalidSource)?;
            }
        }
    }
    budget.references(data_reference_capacity)?;
    budget.work(data_reference_capacity)?;
    budget.allocations(data_reference_capacity)?;
    budget.retained(
        data_reference_capacity
            .checked_mul(size_of::<((u64, u64), usize)>())
            .ok_or(Error::InvalidSource)?,
    )?;
    let mut physical_data_references = HashMap::new();
    physical_data_references
        .try_reserve(data_reference_capacity)
        .map_err(|_| Error::Allocation(data_reference_capacity))?;
    for component in package.state.source.components().iter() {
        budget.components(1)?;
        for object in &component.archive().objects {
            budget.payload_objects(1)?;
            let identifier = object.archive_info.identifier.ok_or(Error::InvalidSource)?;
            for info in &object.archive_info.message_infos {
                for data_identifier in &info.data_references {
                    let count = physical_data_references
                        .entry((identifier, *data_identifier))
                        .or_insert(0usize);
                    *count = count.checked_add(1).ok_or(Error::InvalidSource)?;
                }
            }
        }
    }
    let mut selected_targets = Vec::new();
    let target_count = targets
        .len()
        .checked_add(storage.len())
        .ok_or(Error::InvalidSource)?;
    budget.allocations(target_count)?;
    budget.retained(
        target_count
            .checked_mul(size_of::<MetadataTarget<'_>>())
            .ok_or(Error::InvalidSource)?,
    )?;
    selected_targets
        .try_reserve_exact(target_count)
        .map_err(|_| Error::Allocation(target_count))?;
    for target in targets {
        selected_targets.push(MetadataTarget {
            identifier: target.identifier,
            component_name: target.component.as_ref(),
            requires_uuid: true,
            storage: false,
        });
    }
    for route in storage {
        selected_targets.push(MetadataTarget {
            identifier: route.location.identifier,
            component_name: route.location.component.as_ref(),
            requires_uuid: false,
            storage: true,
        });
    }
    let mut payload = None;
    for component in package.state.source.components().iter() {
        budget.components(1)?;
        for object in &component.archive().objects {
            budget.payload_objects(1)?;
            budget.work(object.messages.len())?;
            for (index, message) in object.messages.iter().enumerate() {
                validate_message_header(object, index)?;
                budget.payload_messages(1)?;
                budget.work(message.data.len())?;
                if message.type_ == PACKAGE_METADATA_MESSAGE_TYPE
                    && payload.replace(message.data.as_slice()).is_some()
                {
                    return Err(Error::InvalidSource);
                }
            }
        }
    }
    let payload = payload.ok_or(Error::InvalidSource)?;
    let limits = budget.residual(package)?;
    let remaining_components = Budget::remaining(
        budget.components,
        budget.max_components,
        LimitKind::Components,
    )?;
    let remaining_references = Budget::remaining(
        budget.references,
        budget.max_references,
        LimitKind::References,
    )?;
    let metadata_map_capacity = physical_capacity
        .checked_mul(7)
        .ok_or(Error::InvalidSource)?;
    budget.allocations(metadata_map_capacity)?;
    budget.retained(
        metadata_map_capacity
            .checked_mul(size_of::<u64>())
            .ok_or(Error::InvalidSource)?,
    )?;
    let options = package_metadata_codec::RewriteOptions::new(
        payload.len().min(limits.max_input_bytes()),
        limits.max_output_bytes(),
        limits.max_fields(),
        limits.max_rewrite_work(),
        u32::try_from(limits.max_nesting()).unwrap_or(u32::MAX),
        remaining_components,
        remaining_references,
        remaining_references,
    );
    let mut visitor = StrictPackageMetadataVisitor::new(
        &physical_identifiers,
        &physical_data_references,
        &selected_targets,
    )
    .map_err(|_| Error::Allocation(selected_targets.len()))?;
    let inspection = package_metadata_codec::inspect_package_metadata_with_visitor(
        payload,
        options,
        &mut visitor,
    )
    .map_err(|_| Error::Codec)?;
    budget.metadata_report(inspection.report())?;
    let storage_targets = selected_targets
        .iter()
        .filter(|target| target.storage)
        .count();
    let bound_storage_targets = selected_targets
        .iter()
        .filter(|target| {
            target.storage && visitor.object_components.contains_key(&target.identifier)
        })
        .count();
    let partial_storage_authority =
        bound_storage_targets != 0 && bound_storage_targets != storage_targets;
    let incomplete_data_reference = visitor
        .active_data_reference
        .is_some_and(|reference| reference.remaining_owners != 0);
    if visitor.unknown
        || visitor
            .external_component_identifiers
            .iter()
            .any(|id| !visitor.component_identifiers.contains(id))
        || visitor
            .external_object_identifiers
            .iter()
            .any(|id| !physical_identifiers.contains(id))
        || inspection.last_object_identifier() < physical_maximum
        || visitor.authority_invalid
        || visitor.duplicate_uuid
        || visitor.duplicate_component
        || visitor.selected_mismatch
        || partial_storage_authority
        || incomplete_data_reference
        || visitor.seen_data_owners.len() != physical_data_references.len()
        || selected_targets.iter().any(|target| {
            target.requires_uuid && !visitor.object_components.contains_key(&target.identifier)
        })
    {
        return Err(Error::UnsupportedDependency);
    }
    Ok(())
}

fn metadata_component_matches_physical(metadata: &str, physical: &str) -> bool {
    fn canonical(name: &str) -> &str {
        let name = name.strip_suffix(".iwa").unwrap_or(name);
        // PackageMetadata addresses components relative to the `Index/`
        // archive directory, while SourceCatalog retains that directory in
        // its physical entry name.  Normalize only that fixed prefix and the
        // native suffix; never collapse arbitrary path components.
        name.strip_prefix("Index/").unwrap_or(name)
    }
    canonical(metadata) == canonical(physical)
}

#[derive(Clone, Copy)]
struct ActiveDataReference {
    component_identifier: u64,
    data_identifier: u64,
    remaining_owners: usize,
}

struct StrictPackageMetadataVisitor<'a> {
    physical_identifiers: &'a HashSet<u64>,
    physical_data_references: &'a HashMap<(u64, u64), usize>,
    selected_targets: &'a [MetadataTarget<'a>],
    unknown: bool,
    authority_invalid: bool,
    duplicate_uuid: bool,
    duplicate_component: bool,
    selected_mismatch: bool,
    component_identifiers: HashSet<u64>,
    external_component_identifiers: HashSet<u64>,
    external_object_identifiers: HashSet<u64>,
    uuid_pairs: HashSet<(u64, u64)>,
    object_bindings: HashMap<(u64, u64), (u64, u64)>,
    object_components: HashMap<u64, u64>,
    active_data_reference: Option<ActiveDataReference>,
    seen_data_owners: HashSet<(u64, u64)>,
    saw_data_metadata_map: bool,
}

impl<'a> StrictPackageMetadataVisitor<'a> {
    fn new(
        physical_identifiers: &'a HashSet<u64>,
        physical_data_references: &'a HashMap<(u64, u64), usize>,
        selected_targets: &'a [MetadataTarget<'a>],
    ) -> std::result::Result<Self, ()> {
        let capacity = physical_identifiers.len();
        let mut component_identifiers = HashSet::new();
        component_identifiers
            .try_reserve(capacity)
            .map_err(|_| ())?;
        let mut external_component_identifiers = HashSet::new();
        external_component_identifiers
            .try_reserve(capacity)
            .map_err(|_| ())?;
        let mut external_object_identifiers = HashSet::new();
        external_object_identifiers
            .try_reserve(capacity)
            .map_err(|_| ())?;
        let mut uuid_pairs = HashSet::new();
        uuid_pairs.try_reserve(capacity).map_err(|_| ())?;
        let mut object_bindings = HashMap::new();
        object_bindings.try_reserve(capacity).map_err(|_| ())?;
        let mut object_components = HashMap::new();
        object_components.try_reserve(capacity).map_err(|_| ())?;
        let mut seen_data_owners = HashSet::new();
        seen_data_owners
            .try_reserve(physical_data_references.len())
            .map_err(|_| ())?;
        Ok(Self {
            physical_identifiers,
            physical_data_references,
            selected_targets,
            unknown: false,
            authority_invalid: false,
            duplicate_uuid: false,
            duplicate_component: false,
            selected_mismatch: false,
            component_identifiers,
            external_component_identifiers,
            external_object_identifiers,
            uuid_pairs,
            object_bindings,
            object_components,
            active_data_reference: None,
            seen_data_owners,
            saw_data_metadata_map: false,
        })
    }
}

impl package_metadata_codec::PackageMetadataVisitor for StrictPackageMetadataVisitor<'_> {
    fn visit_unknown_field(
        &mut self,
    ) -> std::result::Result<(), package_metadata_codec::RewriteError> {
        self.unknown = true;
        Ok(())
    }

    fn visit_component(
        &mut self,
        component: package_metadata_codec::ComponentDescriptor<'_>,
    ) -> std::result::Result<(), package_metadata_codec::RewriteError> {
        if !component.is_current() {
            self.authority_invalid = true;
        }
        self.component_identifiers
            .try_reserve(1)
            .map_err(|_| package_metadata_codec::RewriteError::allocation(1))?;
        if !self.component_identifiers.insert(component.identifier()) {
            self.duplicate_component = true;
        }
        Ok(())
    }

    fn visit_object_uuid(
        &mut self,
        binding: package_metadata_codec::ObjectUuidDescriptor<'_>,
    ) -> std::result::Result<(), package_metadata_codec::RewriteError> {
        if !binding.component().is_current()
            || !self
                .physical_identifiers
                .contains(&binding.object_identifier())
        {
            self.authority_invalid = true;
        }
        let uuid = binding.uuid();
        let component_identifier = binding.component().identifier();
        let object_identifier = binding.object_identifier();
        let pair = (uuid.lower(), uuid.upper());
        self.object_bindings
            .try_reserve(1)
            .map_err(|_| package_metadata_codec::RewriteError::allocation(1))?;
        self.object_components
            .try_reserve(1)
            .map_err(|_| package_metadata_codec::RewriteError::allocation(1))?;
        self.uuid_pairs
            .try_reserve(1)
            .map_err(|_| package_metadata_codec::RewriteError::allocation(1))?;
        if pair == (0, 0) || !self.uuid_pairs.insert(pair) {
            self.duplicate_uuid = true;
        }
        if self
            .object_bindings
            .insert((component_identifier, object_identifier), pair)
            .is_some()
        {
            self.duplicate_uuid = true;
        }
        if self
            .object_components
            .insert(object_identifier, component_identifier)
            .is_some()
        {
            self.duplicate_uuid = true;
        }
        for target in self.selected_targets {
            if target.identifier == object_identifier
                && !metadata_component_matches_physical(
                    binding.component().effective_locator(),
                    target.component_name,
                )
            {
                self.selected_mismatch = true;
            }
        }
        Ok(())
    }

    fn visit_external_reference(
        &mut self,
        reference: package_metadata_codec::ExternalReferenceDescriptor<'_>,
    ) -> std::result::Result<(), package_metadata_codec::RewriteError> {
        // Cross-component edges are a normal part of current Keynote table
        // graphs. They remain admissible only when the source is current,
        // unversioned, and the target component/object is proven below to be
        // part of this physical package.
        if !reference.source().is_current() || reference.is_versioned() {
            self.authority_invalid = true;
        }
        self.external_component_identifiers
            .try_reserve(1)
            .map_err(|_| package_metadata_codec::RewriteError::allocation(1))?;
        self.external_component_identifiers
            .insert(reference.target_component_identifier());
        if let Some(identifier) = reference.object_identifier() {
            self.external_object_identifiers
                .try_reserve(1)
                .map_err(|_| package_metadata_codec::RewriteError::allocation(1))?;
            self.external_object_identifiers.insert(identifier);
        }
        Ok(())
    }

    fn visit_data_reference(
        &mut self,
        reference: package_metadata_codec::DataReferenceDescriptor<'_>,
    ) -> std::result::Result<(), package_metadata_codec::RewriteError> {
        if self
            .active_data_reference
            .is_some_and(|active| active.remaining_owners != 0)
            || !reference.component().is_current()
            || reference.has_unknown_fields()
            || reference.owner_count() == 0
        {
            self.authority_invalid = true;
        }
        self.active_data_reference = Some(ActiveDataReference {
            component_identifier: reference.component().identifier(),
            data_identifier: reference.data_identifier(),
            remaining_owners: reference.owner_count(),
        });
        Ok(())
    }

    fn visit_data_reference_owner(
        &mut self,
        owner: package_metadata_codec::DataReferenceOwnerDescriptor<'_>,
    ) -> std::result::Result<(), package_metadata_codec::RewriteError> {
        let Some(active) = self.active_data_reference else {
            self.authority_invalid = true;
            return Ok(());
        };
        let key = (owner.object_identifier(), owner.data_identifier());
        let reported_count = usize::try_from(owner.count()).ok();
        let selected_owner = self
            .selected_targets
            .iter()
            .any(|target| target.identifier == owner.object_identifier());
        let valid = owner.component().is_current()
            && !owner.has_unknown_fields()
            && owner.component().identifier() == active.component_identifier
            && owner.data_identifier() == active.data_identifier
            && active.remaining_owners != 0
            && self
                .physical_identifiers
                .contains(&owner.object_identifier())
            && !selected_owner
            && reported_count
                .is_some_and(|count| self.physical_data_references.get(&key) == Some(&count))
            && !self.seen_data_owners.contains(&key);
        if !valid {
            self.authority_invalid = true;
        } else if !self.seen_data_owners.insert(key) {
            // The membership check above and insert are intentionally kept
            // together: the visitor is single-threaded, so disagreement can
            // only indicate an internal invariant violation.
            self.authority_invalid = true;
        } else if let Some(active) = self.active_data_reference.as_mut() {
            active.remaining_owners -= 1;
        } else {
            self.authority_invalid = true;
        }
        Ok(())
    }

    fn visit_ambiguous_object_identifier(
        &mut self,
        _component: package_metadata_codec::ComponentDescriptor<'_>,
        _identifier: u64,
    ) -> std::result::Result<(), package_metadata_codec::RewriteError> {
        self.authority_invalid = true;
        Ok(())
    }

    fn visit_data_metadata_map(
        &mut self,
        object_identifier: u64,
        has_unknown_fields: bool,
    ) -> std::result::Result<(), package_metadata_codec::RewriteError> {
        if self.saw_data_metadata_map
            || has_unknown_fields
            || !self.physical_identifiers.contains(&object_identifier)
            || self
                .selected_targets
                .iter()
                .any(|target| target.identifier == object_identifier)
        {
            self.authority_invalid = true;
        }
        self.saw_data_metadata_map = true;
        Ok(())
    }
}

fn validate_global_inbound_references(
    package: &Package,
    slide_identifier: u64,
    slide_message_index: usize,
    table_info_identifier: u64,
    table_info_message_index: usize,
    model_identifier: u64,
    model_message_index: usize,
    storage: &[StorageRoute],
    budget: &mut Budget,
) -> Result<()> {
    let limits = package
        .state
        .options
        .archive()
        .effective_archive_limits()
        .map_err(|_| Error::Archive)?;
    let slide = object_at(package, &locate_object(package, slide_identifier)?)?;
    let slide_info = slide
        .archive_info
        .message_infos
        .get(slide_message_index)
        .ok_or(Error::InvalidSource)?;
    let table_info = object_at(package, &locate_object(package, table_info_identifier)?)?;
    let table_info_info = table_info
        .archive_info
        .message_infos
        .get(table_info_message_index)
        .ok_or(Error::InvalidSource)?;
    let model = object_at(package, &locate_object(package, model_identifier)?)?;
    let model_info = model
        .archive_info
        .message_infos
        .get(model_message_index)
        .ok_or(Error::InvalidSource)?;
    let expected_info_edges = reference_occurrence_count(slide_info, table_info_identifier)?;
    let expected_model_edges = reference_occurrence_count(table_info_info, model_identifier)?;
    if expected_info_edges == 0 || expected_model_edges == 0 {
        return Err(Error::UnsupportedDependency);
    }
    budget.fields(
        slide_info
            .field_infos
            .len()
            .checked_add(table_info_info.field_infos.len())
            .and_then(|size| size.checked_add(model_info.field_infos.len()))
            .ok_or(Error::InvalidSource)?,
    )?;
    budget.references(
        slide_info
            .object_references
            .len()
            .checked_add(table_info_info.object_references.len())
            .and_then(|size| size.checked_add(model_info.object_references.len()))
            .ok_or(Error::InvalidSource)?,
    )?;
    let mut model_targets = HashMap::new();
    budget.allocations(storage.len())?;
    budget.retained(
        storage
            .len()
            .checked_mul(size_of::<(u64, usize)>())
            .ok_or(Error::InvalidSource)?,
    )?;
    model_targets
        .try_reserve(storage.len())
        .map_err(|_| Error::Allocation(storage.len()))?;
    for route in storage {
        let expected = reference_occurrence_count(model_info, route.location.identifier)?;
        if expected == 0 {
            return Err(Error::UnsupportedDependency);
        }
        if model_targets
            .insert(route.location.identifier, expected)
            .is_some()
        {
            return Err(Error::UnsupportedDependency);
        }
    }
    let mut model_route_fields = HashSet::new();
    budget.allocations(model_info.field_infos.len())?;
    budget.retained(
        model_info
            .field_infos
            .len()
            .checked_mul(size_of::<usize>())
            .ok_or(Error::InvalidSource)?,
    )?;
    model_route_fields
        .try_reserve(model_info.field_infos.len())
        .map_err(|_| Error::Allocation(model_info.field_infos.len()))?;
    for (field_index, field) in model_info.field_infos.iter().enumerate() {
        if let [identifier] = field.object_references.as_slice()
            && let Some(route) = storage
                .iter()
                .find(|route| route.location.identifier == *identifier)
            && field.path.as_slice() == route.kind.path()
        {
            model_route_fields.insert(field_index);
        }
    }
    if model_targets.is_empty()
        || (!model_route_fields.is_empty() && model_route_fields.len() != storage.len())
    {
        return Err(Error::UnsupportedDependency);
    }
    let mut model_edges = HashMap::new();
    budget.allocations(model_targets.len())?;
    budget.retained(
        model_targets
            .len()
            .checked_mul(size_of::<(u64, usize)>())
            .ok_or(Error::InvalidSource)?,
    )?;
    model_edges
        .try_reserve(model_targets.len())
        .map_err(|_| Error::Allocation(model_targets.len()))?;
    let mut census = InboundReferenceCensus {
        slide_identifier,
        slide_message_index,
        table_info_identifier,
        table_info_message_index,
        model_identifier,
        slide_field_count: slide_info.field_infos.len(),
        table_info_field_count: table_info_info.field_infos.len(),
        slide_route_fields: route_field_indices(
            slide_info,
            [&[SLIDE_OWNED_DRAWABLES_FIELD], &[SLIDE_Z_ORDER_FIELD]],
        ),
        table_info_route_fields: route_field_indices(
            table_info_info,
            [
                &[TABLE_MODEL_FIELD],
                &[TABLE_SUPER_FIELD, DRAWABLE_PARENT_FIELD],
            ],
        ),
        expected_info_edges,
        expected_model_edges,
        model_targets,
        model_route_fields,
        model_field_count: model_info.field_infos.len(),
        model_message_index,
        model_edges,
        info_edges: 0,
        model_info_edges: 0,
        invalid: false,
    };
    for component in package.state.source.components().iter() {
        budget.components(1)?;
        for object in &component.archive().objects {
            budget.payload_objects(1)?;
            if object.archive_info.identifier.is_none()
                || object.messages.len() != object.archive_info.message_infos.len()
            {
                return Err(Error::InvalidSource);
            }
            let field_count = object
                .archive_info
                .message_infos
                .iter()
                .map(|info| info.field_infos.len())
                .try_fold(0usize, |sum, value| sum.checked_add(value))
                .ok_or(Error::InvalidSource)?;
            let reference_count =
                object
                    .archive_info
                    .message_infos
                    .iter()
                    .try_fold(0usize, |sum, info| {
                        let nested = info.field_infos.iter().try_fold(0usize, |sum, field| {
                            sum.checked_add(field.object_references.len())
                                .and_then(|value| value.checked_add(field.data_references.len()))
                                .ok_or(Error::InvalidSource)
                        })?;
                        sum.checked_add(info.object_references.len())
                            .and_then(|value| value.checked_add(info.data_references.len()))
                            .and_then(|value| value.checked_add(nested))
                            .ok_or(Error::InvalidSource)
                    })?;
            let message_bytes = object
                .messages
                .iter()
                .map(|message| message.data.len())
                .try_fold(0usize, |sum, value| sum.checked_add(value))
                .ok_or(Error::InvalidSource)?;
            budget.payload_messages(object.messages.len())?;
            budget.fields(field_count)?;
            budget.references(reference_count)?;
            budget.allocations(
                object
                    .messages
                    .len()
                    .checked_add(field_count)
                    .and_then(|value| value.checked_add(reference_count))
                    .and_then(|value| value.checked_add(1))
                    .ok_or(Error::InvalidSource)?,
            )?;
            budget.work(
                object
                    .messages
                    .len()
                    .checked_add(message_bytes)
                    .and_then(|value| value.checked_add(field_count))
                    .ok_or(Error::InvalidSource)?,
            )?;
            object
                .inspect_references_with_policy_and_limits(
                    &mut census,
                    ArchiveReferencePolicy::RejectUnknownMetadata,
                    limits,
                )
                .map_err(|_| Error::Archive)?;
        }
    }
    let model_edges_valid = census.model_targets.iter().all(|(identifier, expected)| {
        census.model_edges.get(identifier).copied() == Some(*expected)
    });
    if census.invalid
        || census.info_edges != census.expected_info_edges
        || census.model_info_edges != census.expected_model_edges
        || !model_edges_valid
    {
        return Err(Error::UnsupportedDependency);
    }
    Ok(())
}

fn reference_occurrence_count(
    info: &litchi_iwa_core::MessageInfo,
    identifier: u64,
) -> Result<usize> {
    let direct = info
        .object_references
        .iter()
        .filter(|value| **value == identifier)
        .count();
    info.field_infos.iter().try_fold(direct, |sum, field| {
        sum.checked_add(
            field
                .object_references
                .iter()
                .filter(|value| **value == identifier)
                .count(),
        )
        .ok_or(Error::InvalidSource)
    })
}

fn route_field_indices(
    info: &litchi_iwa_core::MessageInfo,
    paths: [&[u32]; 2],
) -> [Option<usize>; 2] {
    [
        info.field_infos
            .iter()
            .position(|field| field.path.as_slice() == paths[0]),
        info.field_infos
            .iter()
            .position(|field| field.path.as_slice() == paths[1]),
    ]
}

struct InboundReferenceCensus {
    slide_identifier: u64,
    slide_message_index: usize,
    table_info_identifier: u64,
    table_info_message_index: usize,
    model_identifier: u64,
    slide_field_count: usize,
    table_info_field_count: usize,
    slide_route_fields: [Option<usize>; 2],
    table_info_route_fields: [Option<usize>; 2],
    expected_info_edges: usize,
    expected_model_edges: usize,
    model_targets: HashMap<u64, usize>,
    model_route_fields: HashSet<usize>,
    model_field_count: usize,
    model_message_index: usize,
    info_edges: usize,
    model_info_edges: usize,
    model_edges: HashMap<u64, usize>,
    invalid: bool,
}

impl ArchiveReferenceVisitor for InboundReferenceCensus {
    fn visit_reference(
        &mut self,
        occurrence: ArchiveReferenceOccurrence,
    ) -> litchi_iwa_core::Result<()> {
        if occurrence.referenced_identifier == self.table_info_identifier {
            let allowed = match occurrence.scope {
                ArchiveReferenceScope::Message => true,
                ArchiveReferenceScope::Field { field_index } => {
                    field_index < self.slide_field_count
                        && self.slide_route_fields.contains(&Some(field_index))
                },
            };
            if occurrence.kind != ArchiveReferenceKind::Object
                || occurrence.object_identifier != self.slide_identifier
                || occurrence.message_index != self.slide_message_index
                || !allowed
            {
                self.invalid = true;
            } else {
                if let Some(next) = self.info_edges.checked_add(1) {
                    self.info_edges = next;
                } else {
                    self.invalid = true;
                }
            }
        }
        if occurrence.referenced_identifier == self.model_identifier {
            let allowed = match occurrence.scope {
                ArchiveReferenceScope::Message => true,
                ArchiveReferenceScope::Field { field_index } => {
                    field_index < self.table_info_field_count
                        && self.table_info_route_fields.contains(&Some(field_index))
                },
            };
            if occurrence.kind != ArchiveReferenceKind::Object
                || occurrence.object_identifier != self.table_info_identifier
                || occurrence.message_index != self.table_info_message_index
                || !allowed
            {
                self.invalid = true;
            } else {
                if let Some(next) = self.model_info_edges.checked_add(1) {
                    self.model_info_edges = next;
                } else {
                    self.invalid = true;
                }
            }
        }
        if let Some(expected) = self.model_targets.get(&occurrence.referenced_identifier) {
            let scope_is_allowed = match occurrence.scope {
                ArchiveReferenceScope::Message => true,
                ArchiveReferenceScope::Field { field_index } => {
                    field_index < self.model_field_count
                        && self.model_route_fields.contains(&field_index)
                },
            };
            if occurrence.kind != ArchiveReferenceKind::Object
                || occurrence.object_identifier != self.model_identifier
                || occurrence.message_index != self.model_message_index
                || !scope_is_allowed
            {
                self.invalid = true;
            } else {
                let count = self
                    .model_edges
                    .entry(occurrence.referenced_identifier)
                    .or_insert(0);
                if let Some(next) = count.checked_add(1) {
                    *count = next;
                    if next > *expected {
                        self.invalid = true;
                    }
                } else {
                    self.invalid = true;
                }
            }
        }
        Ok(())
    }
}

fn charge_archive_inventory(archive: &Archive, budget: &mut Budget) -> Result<()> {
    let objects = archive.objects.len();
    let messages = archive
        .objects
        .iter()
        .map(|object| object.messages.len())
        .try_fold(0usize, |sum, value| sum.checked_add(value))
        .ok_or(Error::InvalidSource)?;
    let fields = archive
        .objects
        .iter()
        .flat_map(|object| object.archive_info.message_infos.iter())
        .map(|info| info.field_infos.len())
        .try_fold(0usize, |sum, value| sum.checked_add(value))
        .ok_or(Error::InvalidSource)?;
    let references = archive
        .objects
        .iter()
        .flat_map(|object| object.archive_info.message_infos.iter())
        .try_fold(0usize, |sum, info| {
            let nested = info.field_infos.iter().try_fold(0usize, |sum, field| {
                sum.checked_add(field.object_references.len())
                    .and_then(|value| value.checked_add(field.data_references.len()))
                    .ok_or(Error::InvalidSource)
            })?;
            sum.checked_add(info.object_references.len())
                .and_then(|value| value.checked_add(info.data_references.len()))
                .and_then(|value| value.checked_add(nested))
                .ok_or(Error::InvalidSource)
        })?;
    let message_bytes = archive
        .objects
        .iter()
        .flat_map(|object| object.messages.iter())
        .map(|message| message.data.len())
        .try_fold(0usize, |sum, value| sum.checked_add(value))
        .ok_or(Error::InvalidSource)?;
    budget.payload_objects(objects)?;
    budget.payload_messages(messages)?;
    budget.fields(fields)?;
    budget.references(references)?;
    budget.allocations(
        objects
            .checked_add(messages)
            .and_then(|value| value.checked_add(fields))
            .and_then(|value| value.checked_add(references))
            .ok_or(Error::InvalidSource)?,
    )?;
    budget.work(
        objects
            .checked_add(messages)
            .and_then(|value| value.checked_add(fields))
            .and_then(|value| value.checked_add(references))
            .and_then(|value| value.checked_add(message_bytes))
            .ok_or(Error::InvalidSource)?,
    )?;
    // Archive parsing copies each message payload (and may retain a
    // non-canonical header) into the returned object graph. Count that
    // retained payload before callers use the archive for further scans.
    budget.retained(message_bytes)?;
    Ok(())
}

#[cfg(test)]
mod tests {
    use litchi_iwa_core::RawMessage;

    use super::*;

    #[test]
    fn mutable_roots_require_one_known_message_shape() {
        let tile = ArchiveObject::new(
            1,
            vec![RawMessage {
                type_: TILE_MESSAGE_TYPE,
                data: Vec::new(),
            }],
        )
        .expect("canonical test object");
        assert!(has_exact_mutable_message_shape(&tile, &[TILE_MESSAGE_TYPE]));

        let wrong_role = ArchiveObject::new(
            2,
            vec![RawMessage {
                type_: HEADER_BUCKET_MESSAGE_TYPE,
                data: Vec::new(),
            }],
        )
        .expect("canonical test object");
        assert!(!has_exact_mutable_message_shape(
            &wrong_role,
            &[TILE_MESSAGE_TYPE]
        ));

        let co_located = ArchiveObject::new(
            3,
            vec![
                RawMessage {
                    type_: TILE_MESSAGE_TYPE,
                    data: Vec::new(),
                },
                RawMessage {
                    type_: TILE_MESSAGE_TYPE,
                    data: Vec::new(),
                },
            ],
        )
        .expect("canonical test object");
        assert!(!has_exact_mutable_message_shape(
            &co_located,
            &[TILE_MESSAGE_TYPE]
        ));
    }

    #[test]
    fn metadata_component_matching_preserves_full_component_path() {
        assert!(metadata_component_matches_physical(
            "Tables/Tile",
            "Index/Tables/Tile.iwa"
        ));
        assert!(metadata_component_matches_physical(
            "Index/Tables/Tile.iwa",
            "Tables/Tile"
        ));
        assert!(!metadata_component_matches_physical(
            "A/Tables/Tile",
            "B/Tables/Tile.iwa"
        ));
        assert!(!metadata_component_matches_physical(
            "Tile",
            "Index/Tables/Tile.iwa"
        ));
    }

    #[test]
    fn mutable_root_census_rejects_non_owner_edges() {
        let expected = HashMap::from([(42, 1)]);
        let allowed_fields = HashMap::new();
        let mut census = MutableRootInboundReferenceCensus {
            expected: &expected,
            allowed_fields: &allowed_fields,
            owner_identifier: 7,
            owner_message_index: 0,
            observed: HashMap::new(),
            invalid: false,
        };
        census
            .visit_reference(ArchiveReferenceOccurrence {
                object_identifier: 7,
                message_index: 0,
                scope: ArchiveReferenceScope::Message,
                kind: ArchiveReferenceKind::Object,
                referenced_identifier: 42,
            })
            .expect("reference visitor accepts a valid owner edge");
        assert!(!census.invalid);
        assert_eq!(census.observed.get(&42), Some(&1));

        census
            .visit_reference(ArchiveReferenceOccurrence {
                object_identifier: 9,
                message_index: 0,
                scope: ArchiveReferenceScope::Message,
                kind: ArchiveReferenceKind::Object,
                referenced_identifier: 42,
            })
            .expect("reference visitor reports aliases without leaking errors");
        assert!(census.invalid);
    }
}
