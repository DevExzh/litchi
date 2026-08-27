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
    package_metadata_codec, table_dimension_codec, table_info_codec, table_model_discovery_codec,
};

use super::{Package, PayloadLimitKind, PhysicalSource, ReadError, SemanticLimitKind};
use crate::{SlideSelector, slide::table::TableSelector};

const SLIDE_MESSAGE_TYPE: u32 = 5;
const TABLE_INFO_MESSAGE_TYPE: u32 = 6_000;
const TABLE_MODEL_MESSAGE_TYPE: u32 = 6_001;
const HEADER_BUCKET_MESSAGE_TYPE: u32 = 6_006;
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

const ROLE_MESSAGE_TYPES: [u32; 7] = [
    TABLE_INFO_MESSAGE_TYPE,
    TABLE_MODEL_MESSAGE_TYPE,
    HEADER_BUCKET_MESSAGE_TYPE,
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
            .unwrap_or(usize::MAX);
        let max_total = usize::try_from(package.state.options.archive().max_total_bytes())
            .map_err(|_| Error::InvalidSource)?
            .checked_mul(passes)
            .unwrap_or(usize::MAX);
        let max_components = max_entries;
        let max_objects = archive
            .max_objects()
            .checked_mul(max_components)
            .unwrap_or(usize::MAX);
        let max_messages = archive
            .max_messages()
            .checked_mul(max_components)
            .unwrap_or(usize::MAX);
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
                observed: used.saturating_add(1) as u64,
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

    fn storage_codec_options(
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

    fn storage_codec_report(&mut self, report: table_dimension_codec::DecodeReport) -> Result<()> {
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
enum StorageRouteKind {
    RowBucket,
    ColumnBucket,
    StringTable,
    StyleTable,
    FormulaTable,
    FormatTablePreBnc,
}

impl StorageRouteKind {
    const fn path(self) -> &'static [u32] {
        match self {
            Self::RowBucket => MODEL_ROW_BUCKET_PATH,
            Self::ColumnBucket => MODEL_COLUMN_BUCKET_PATH,
            Self::StringTable => MODEL_STRING_TABLE_PATH,
            Self::StyleTable => MODEL_STYLE_TABLE_PATH,
            Self::FormulaTable => MODEL_FORMULA_TABLE_PATH,
            Self::FormatTablePreBnc => MODEL_FORMAT_TABLE_PRE_BNC_PATH,
        }
    }

    const fn is_bucket(self) -> bool {
        matches!(self, Self::RowBucket | Self::ColumnBucket)
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct StorageRoute {
    location: ObjectLocation,
    kind: StorageRouteKind,
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
    storage: Vec<StorageRoute>,
    pub(crate) locked: bool,
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
    let slide_position = resolve_slide_position(package, slide_selector, budget)?;
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
            info_object,
            info_message_index,
            parent,
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
        validate_model_payload(model_payload, package, budget)?;
        candidates.push((
            info_location,
            model_location,
            info_message_index,
            model_message_index,
            info.locked().unwrap_or(false),
        ));
    }
    let table_position = table_selector.as_position();
    let (table_info, model, table_info_message_index, model_message_index, locked) = candidates
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
        storage,
        locked,
    })
}

pub(crate) fn same_target(left: &Target, right: &Target) -> bool {
    left == right
}

pub(crate) fn physical_catalog(package: &Package) -> Result<&SourceCatalog> {
    match &package.state.source {
        PhysicalSource::Package(source) => Ok(source),
        PhysicalSource::Semantic(_) => Err(Error::UnsupportedSource),
    }
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
    let object = package
        .state
        .source
        .components()
        .get_index(
            package
                .state
                .object_index
                .iter()
                .find(|locator| locator.identifier == location.identifier)
                .map(|locator| locator.component)
                .ok_or(Error::InvalidSource)?,
        )
        .and_then(|component| component.archive().objects.get(location.object_index))
        .ok_or(Error::InvalidSource)?;
    if object.archive_info.identifier != Some(location.identifier) {
        return Err(Error::InvalidSource);
    }
    Ok(object)
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
    let catalog = physical_catalog(package)?;
    let entry = catalog
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

fn locate_object(package: &Package, identifier: u64) -> Result<ObjectLocation> {
    let locator = package
        .state
        .object_index
        .iter()
        .find(|locator| locator.identifier == identifier)
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

fn unique_message<'a>(
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
    let expected_capacity = owned.len().saturating_add(z_order.len());
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
    object: &ArchiveObject,
    message_index: usize,
    parent: u64,
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
    let direct_is_canonical = info.object_references.as_slice() == [model]
        || (info.object_references.len() == 2
            && info.object_references.contains(&parent)
            && info.object_references.contains(&model));
    if !direct_is_canonical {
        return Err(Error::InvalidSource);
    }
    if !info.object_references.is_empty() {
        reject_duplicates(&info.object_references, budget)?;
    }
    budget.references(info.object_references.len())?;
    budget.fields(info.field_infos.len())?;
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
        match field.path.as_slice() {
            [TABLE_MODEL_FIELD] => {
                if saw_model
                    || !is_message_reference_field(field)
                    || field.object_references.as_slice() != [model]
                {
                    return Err(Error::InvalidSource);
                }
                saw_model = true;
            },
            _ if !field.object_references.is_empty() => return Err(Error::UnsupportedTopology),
            _ => {},
        }
    }
    if !saw_model {
        return Err(Error::InvalidSource);
    }
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

fn validate_model_payload(payload: &[u8], package: &Package, budget: &mut Budget) -> Result<()> {
    let options = budget.model_codec_options(package, payload)?;
    let (_, report) = table_model_discovery_codec::decode_table_model_with_report(payload, options)
        .map_err(|_| Error::Codec)?;
    budget.codec_report(report)
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
    if !info.data_references.is_empty()
        || info.object_references.len() != routes.len()
        || !info
            .object_references
            .iter()
            .copied()
            .eq(routes.iter().map(|route| route.location.identifier))
    {
        return Err(Error::UnsupportedDependency);
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
        let [identifier] = field.object_references.as_slice() else {
            return Err(Error::UnsupportedDependency);
        };
        let Some(route) = routes
            .iter()
            .find(|route| route.location.identifier == *identifier)
        else {
            return Err(Error::UnsupportedDependency);
        };
        if field.path.as_slice() != route.kind.path() || !seen_fields.insert(*identifier) {
            return Err(Error::UnsupportedDependency);
        }
    }
    if routes
        .iter()
        .any(|route| !seen_fields.contains(&route.location.identifier))
    {
        return Err(Error::UnsupportedDependency);
    }
    Ok(())
}

fn ensure_unique_identity(package: &Package, identifier: u64, budget: &mut Budget) -> Result<()> {
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
                budget.work(message.data.len().saturating_add(1))?;
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
                    owned = owned.saturating_add(owned_hits);
                    z_order = z_order.saturating_add(z_hits);
                    if object.archive_info.identifier == Some(slide_identifier)
                        && owned_hits == 1
                        && z_hits == 1
                    {
                        selected_slide = true;
                    }
                } else if message.type_ == TABLE_INFO_MESSAGE_TYPE {
                    let info = decode_table_info(&message.data, package, budget)?;
                    if info.table_model().identifier().get() == model_identifier {
                        model_owners = model_owners.saturating_add(1);
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
    for component in package.state.source.components().iter() {
        budget.components(1)?;
        for object in &component.archive().objects {
            budget.payload_objects(1)?;
            let identifier = object.archive_info.identifier.ok_or(Error::InvalidSource)?;
            if !physical_identifiers.insert(identifier) {
                return Err(Error::UnsupportedDependency);
            }
            physical_maximum = physical_maximum.max(identifier);
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
        });
    }
    for route in storage {
        selected_targets.push(MetadataTarget {
            identifier: route.location.identifier,
            component_name: route.location.component.as_ref(),
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
    let remaining_components = budget
        .max_components
        .saturating_sub(budget.components)
        .max(1);
    let remaining_references = budget
        .max_references
        .saturating_sub(budget.references)
        .max(1);
    let metadata_map_capacity = physical_capacity
        .checked_mul(6)
        .ok_or(Error::InvalidSource)?;
    budget.allocations(metadata_map_capacity)?;
    budget.retained(
        metadata_map_capacity
            .checked_mul(size_of::<u64>())
            .ok_or(Error::InvalidSource)?,
    )?;
    let options = package_metadata_codec::RewriteOptions::new(
        payload.len().max(1).min(limits.max_input_bytes()),
        limits.max_output_bytes(),
        limits.max_fields(),
        limits.max_rewrite_work(),
        u32::try_from(limits.max_nesting()).unwrap_or(u32::MAX),
        remaining_components,
        remaining_references,
        remaining_references,
    );
    let mut visitor = StrictPackageMetadataVisitor::new(&physical_identifiers, &selected_targets)
        .map_err(|_| Error::Allocation(selected_targets.len()))?;
    let inspection = package_metadata_codec::inspect_package_metadata_with_visitor(
        payload,
        options,
        &mut visitor,
    )
    .map_err(|_| Error::Codec)?;
    budget.metadata_report(inspection.report())?;
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
        || selected_targets
            .iter()
            .any(|target| !visitor.object_components.contains_key(&target.identifier))
    {
        return Err(Error::UnsupportedDependency);
    }
    Ok(())
}

fn metadata_component_matches_physical(metadata: &str, physical: &str) -> bool {
    fn basename(name: &str) -> &str {
        name.rsplit('/').next().unwrap_or(name)
    }
    fn without_iwa(name: &str) -> &str {
        name.strip_suffix(".iwa").unwrap_or(name)
    }
    without_iwa(basename(metadata)) == without_iwa(basename(physical))
}

struct StrictPackageMetadataVisitor<'a> {
    physical_identifiers: &'a HashSet<u64>,
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
}

impl<'a> StrictPackageMetadataVisitor<'a> {
    fn new(
        physical_identifiers: &'a HashSet<u64>,
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
        Ok(Self {
            physical_identifiers,
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
        // The focused owner rewrites a rooted native object and therefore
        // admits only the current package object/UUID authority.  Any
        // external edge would make that authority ambiguous, even when the
        // target component happens to be current and locally present.
        self.authority_invalid = true;
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
        _reference: package_metadata_codec::DataReferenceDescriptor<'_>,
    ) -> std::result::Result<(), package_metadata_codec::RewriteError> {
        self.authority_invalid = true;
        Ok(())
    }

    fn visit_data_reference_owner(
        &mut self,
        _owner: package_metadata_codec::DataReferenceOwnerDescriptor<'_>,
    ) -> std::result::Result<(), package_metadata_codec::RewriteError> {
        self.authority_invalid = true;
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
        _object_identifier: u64,
        _has_unknown_fields: bool,
    ) -> std::result::Result<(), package_metadata_codec::RewriteError> {
        self.authority_invalid = true;
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
        if model_targets
            .insert(route.location.identifier, 2usize)
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
    if model_targets.is_empty() || model_route_fields.len() != storage.len() {
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
                self.info_edges = self.info_edges.saturating_add(1);
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
                self.model_info_edges = self.model_info_edges.saturating_add(1);
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
                *count = count.saturating_add(1);
                if *count > *expected {
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
