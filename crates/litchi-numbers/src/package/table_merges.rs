//! Selector-first, bounded readback of native Numbers merged-cell geometry.
//!
//! The package adapter owns rooted sheet/table selection and the small native
//! ownership proof required to locate a table model.  The selected model's
//! merge-owner path is decoded by the shared borrowed Numbers wire reader.
//! Native object identifiers, archive members, and generated protobuf values
//! never cross this module's public boundary.

mod reader;

pub use reader::MergeReader;

use std::fmt;

use litchi_iwa_common::{
    LimitKind as CommonLimitKind, WireLimits, decode_varint_from_bytes,
    varint::encoded_len,
    wire::{RawWireField, RawWireFields, RawWireLimits},
};
use litchi_iwa_core::{ArchiveObject, RawMessage};
use litchi_numbers_wire::table_merges::{
    self as merge_wire, FORMULA_INDEX_BYTES, PAIR_REFERENCE_BYTES, REGION_BYTES,
};
use thiserror::Error as ThisError;

use super::{
    Components, Error as PackageError, Index, Package, ReadOptions, Resolved, SemanticLimitKind,
};
use crate::table::merge::Region;
use crate::{SheetSelector, TableSelector};

const DOCUMENT_MESSAGE_TYPE: u32 = 1;
const SHEET_MESSAGE_TYPE: u32 = 2;
const FORM_BASED_SHEET_MESSAGE_TYPE: u32 = 3;
const TABLE_INFO_MESSAGE_TYPE: u32 = 6_000;
const TABLE_MODEL_MESSAGE_TYPE: u32 = 6_001;
const LEGACY_TABLE_INFO_MESSAGE_TYPE: u32 = 6_003;
const MERGE_READER_ALLOCATIONS_PER_REGION: usize = 3;
const MERGE_READER_SCRATCH_PER_REGION: usize = PAIR_REFERENCE_BYTES + FORMULA_INDEX_BYTES;
const MIN_SIGN_EXTENDED_I32: u64 = u64::MAX - 2_147_483_647;

/// A finite resource governed by one focused Numbers merged-cell read.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum TableMergesLimitKind {
    /// Complete package input bytes.
    InputBytes,
    /// Complete candidate package output bytes (unused by this read).
    OutputBytes,
    /// Package entries visited by the operation.
    Entries,
    /// Bytes retained by one package entry.
    EntryBytes,
    /// Aggregate package-entry bytes.
    TotalEntryBytes,
    /// Package metadata bytes.
    PackageBytes,
    /// Bytes in one decoded native payload.
    PayloadBytes,
    /// Aggregate decoded native payload bytes.
    TotalPayloadBytes,
    /// Native payload objects inspected.
    PayloadObjects,
    /// Native payload messages inspected.
    PayloadMessages,
    /// Native payload metadata items inspected.
    PayloadItems,
    /// Native object references inspected.
    PayloadReferences,
    /// Rooted sheets visited while resolving selectors.
    Sheets,
    /// Rooted tables visited while resolving selectors.
    Tables,
    /// Bytes inspected by the selected merge wire path.
    WireBytes,
    /// Bytes produced by a merge wire rewrite (unused by this read).
    WireOutputBytes,
    /// Protobuf fields inspected by the operation.
    WireFields,
    /// Maximum nested wire depth inspected.
    WireNesting,
    /// Aggregate wire traversal and formula work.
    WireWork,
    /// Fallible collections admitted by the operation.
    Allocations,
    /// Bytes retained by the semantic result.
    Retained,
    /// Temporary bytes retained during decoding.
    Scratch,
    /// Number of merged regions returned or staged.
    Regions,
}

impl fmt::Display for TableMergesLimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::InputBytes => "input bytes",
            Self::OutputBytes => "output bytes",
            Self::Entries => "entries",
            Self::EntryBytes => "entry bytes",
            Self::TotalEntryBytes => "total entry bytes",
            Self::PackageBytes => "package bytes",
            Self::PayloadBytes => "payload bytes",
            Self::TotalPayloadBytes => "total payload bytes",
            Self::PayloadObjects => "payload objects",
            Self::PayloadMessages => "payload messages",
            Self::PayloadItems => "payload items",
            Self::PayloadReferences => "payload references",
            Self::Sheets => "sheets",
            Self::Tables => "tables",
            Self::WireBytes => "wire bytes",
            Self::WireOutputBytes => "wire output bytes",
            Self::WireFields => "wire fields",
            Self::WireNesting => "wire nesting depth",
            Self::WireWork => "wire work",
            Self::Allocations => "allocations",
            Self::Retained => "retained bytes",
            Self::Scratch => "scratch bytes",
            Self::Regions => "regions",
        })
    }
}

/// Failure while reading one rooted Numbers table's merged-cell geometry.
#[derive(Debug, ThisError, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum TableMergesError {
    /// No rooted sheet matched the selector.
    #[error("the Numbers workbook has no sheet matching the requested selector")]
    SheetNotFound,
    /// No rooted table matched the selector.
    #[error("the selected Numbers sheet has no table matching the requested selector")]
    TableNotFound,
    /// A name selector matched more than one semantic sheet or table.
    #[error("the Numbers source has more than one sheet or table with the requested name")]
    AmbiguousSelector,
    /// The package source cannot be inspected through this exact profile.
    #[error("the Numbers package source does not support focused table-merge reads")]
    UnsupportedSource,
    /// The rooted graph, metadata, or merge formula is malformed.
    #[error("the selected Numbers table-merge source is invalid")]
    InvalidSource,
    /// A finite read resource ceiling was exceeded.
    #[error("Numbers table merges {kind} limit exceeded: observed {observed}, maximum {maximum}")]
    LimitExceeded {
        /// Resource category.
        kind: TableMergesLimitKind,
        /// Observed or requested amount.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
    },
    /// A bounded result or temporary allocation failed.
    #[error("could not allocate {amount} units for Numbers table merges")]
    Allocation {
        /// Requested elements or bytes.
        amount: usize,
    },
}

impl Package {
    /// Read all validated merged-cell regions for one rooted Numbers table.
    ///
    /// Sheet and table selectors use exact visible names, checked zero-based
    /// positions, or typed positions.  Selection follows the immutable
    /// semantic workbook order, while native ownership is proven through the
    /// rooted document graph.  An absent merge owner is a valid empty result.
    /// The read never changes the package source and never exposes native
    /// identifiers or generated protobuf values.
    ///
    /// # Errors
    ///
    /// Returns [`TableMergesError::InvalidSource`] when the rooted graph or a
    /// selected merge formula is malformed, and a typed limit error when any
    /// finite package, wire, work, result, or allocation ceiling is exceeded.
    pub fn table_merges<'sheet, 'table>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        table: impl Into<TableSelector<'table>>,
    ) -> Result<Vec<Region>, TableMergesError> {
        let mut budget = Budget::new(self)?;
        let (sheet_position, table_position) =
            select_semantic_positions(self, sheet.into(), table.into(), &mut budget)?;
        let model = resolve_model_payload(
            &self.state.components,
            &self.state.index,
            sheet_position,
            table_position,
            &mut budget,
        )?;
        decode_model_merges(model, &mut budget)
    }
}

fn decode_model_merges(model: &[u8], budget: &mut Budget) -> Result<Vec<Region>, TableMergesError> {
    let max_regions = merge_region_capacity(
        budget.remaining_allocations()?,
        budget.remaining_scratch()?,
        budget.remaining_retained()?,
        budget.remaining_regions()?,
    );
    let limits = merge_wire_limits(budget, max_regions)?;
    let read = merge_wire::read_table_merges(model, limits).map_err(|error| {
        charge_attempted(budget, &error);
        map_merge_error(error)
    })?;
    budget.charge_wire_report(
        read.report.input_bytes(),
        read.report.fields(),
        read.report.work(),
    )?;
    budget.charge_regions(read.regions.len())?;
    budget.charge_allocations(
        read.regions
            .len()
            .checked_mul(MERGE_READER_ALLOCATIONS_PER_REGION)
            .ok_or(TableMergesError::InvalidSource)?,
    )?;
    budget.charge_scratch(
        read.regions
            .len()
            .checked_mul(MERGE_READER_SCRATCH_PER_REGION)
            .ok_or(TableMergesError::InvalidSource)?,
    )?;
    budget.charge_retained(
        read.regions
            .len()
            .checked_mul(REGION_BYTES)
            .ok_or(TableMergesError::InvalidSource)?,
    )?;
    Ok(read.regions)
}

#[derive(Debug, Clone, Copy)]
struct Budget {
    wire_max_input: usize,
    wire_max_fields: usize,
    wire_max_work: usize,
    wire_max_nesting: usize,
    max_objects: usize,
    max_messages: usize,
    max_items: usize,
    max_references: usize,
    max_sheets: usize,
    max_tables: usize,
    max_allocations: usize,
    max_retained: usize,
    max_scratch: usize,
    max_regions: usize,
    input: usize,
    fields: usize,
    work: usize,
    objects: usize,
    messages: usize,
    items: usize,
    references: usize,
    sheets: usize,
    tables: usize,
    allocations: usize,
    retained: usize,
    scratch: usize,
    regions: usize,
}

impl Budget {
    fn new(package: &Package) -> Result<Self, TableMergesError> {
        Self::from_options(package.state.options)
    }

    fn from_options(options: ReadOptions) -> Result<Self, TableMergesError> {
        let physical = options
            .archive()
            .effective_archive_limits()
            .map_err(map_archive_error)?;
        let semantic = options.semantic();
        let stream = physical.max_archive_bytes().max(1);
        let wire_max_input = stream
            .saturating_mul(8)
            .clamp(1, WireLimits::MAX_INPUT_BYTES);
        let wire_max_fields = physical
            .max_header_fields()
            .saturating_mul(8)
            .clamp(1, WireLimits::MAX_FIELDS);
        let wire_max_work = stream
            .saturating_mul(32)
            .clamp(1, WireLimits::MAX_REWRITE_WORK);
        let max_messages = physical
            .max_messages()
            .saturating_mul(8)
            .clamp(1, WireLimits::MAX_REWRITE_WORK);
        let max_items = physical
            .max_metadata_items()
            .saturating_mul(8)
            .clamp(1, WireLimits::MAX_REWRITE_WORK);
        let max_regions = semantic
            .max_references()
            .min(WireLimits::MAX_FIELDS)
            .min(stream / REGION_BYTES.max(1));
        let max_allocations = semantic
            .max_references()
            .saturating_mul(MERGE_READER_ALLOCATIONS_PER_REGION)
            .min(WireLimits::MAX_REWRITE_WORK);
        Ok(Self {
            wire_max_input,
            wire_max_fields,
            wire_max_work,
            wire_max_nesting: physical.max_header_nesting().min(WireLimits::MAX_NESTING),
            max_objects: semantic.max_objects(),
            max_messages,
            max_items,
            max_references: semantic.max_references(),
            max_sheets: semantic.max_sheets(),
            max_tables: semantic.max_tables(),
            max_allocations,
            max_retained: stream.min(WireLimits::MAX_OUTPUT_BYTES),
            max_scratch: stream.min(WireLimits::MAX_INPUT_BYTES),
            max_regions,
            input: 0,
            fields: 0,
            work: 0,
            objects: 0,
            messages: 0,
            items: 0,
            references: 0,
            sheets: 0,
            tables: 0,
            allocations: 0,
            retained: 0,
            scratch: 0,
            regions: 0,
        })
    }

    fn add(
        current: &mut usize,
        amount: usize,
        maximum: usize,
        kind: TableMergesLimitKind,
    ) -> Result<(), TableMergesError> {
        let observed = current
            .checked_add(amount)
            .ok_or(TableMergesError::InvalidSource)?;
        if observed > maximum {
            return Err(TableMergesError::LimitExceeded {
                kind,
                observed: observed as u64,
                maximum: maximum as u64,
            });
        }
        *current = observed;
        Ok(())
    }

    fn charge_input(&mut self, amount: usize) -> Result<(), TableMergesError> {
        Self::add(
            &mut self.input,
            amount,
            self.wire_max_input,
            TableMergesLimitKind::WireBytes,
        )
    }

    fn charge_fields(&mut self, amount: usize) -> Result<(), TableMergesError> {
        Self::add(
            &mut self.fields,
            amount,
            self.wire_max_fields,
            TableMergesLimitKind::WireFields,
        )
    }

    fn charge_work(&mut self, amount: usize) -> Result<(), TableMergesError> {
        Self::add(
            &mut self.work,
            amount,
            self.wire_max_work,
            TableMergesLimitKind::WireWork,
        )
    }

    fn charge_wire_report(
        &mut self,
        input: usize,
        fields: usize,
        work: usize,
    ) -> Result<(), TableMergesError> {
        self.charge_input(input)?;
        self.charge_fields(fields)?;
        self.charge_work(work)
    }

    fn charge_objects(&mut self, amount: usize) -> Result<(), TableMergesError> {
        Self::add(
            &mut self.objects,
            amount,
            self.max_objects,
            TableMergesLimitKind::PayloadObjects,
        )
    }

    fn charge_messages(&mut self, amount: usize) -> Result<(), TableMergesError> {
        // The message scan is linear even when the caller ultimately selects
        // no message of this type.  Account for that traversal before the
        // collection is inspected.
        self.charge_work(amount)?;
        Self::add(
            &mut self.messages,
            amount,
            self.max_messages,
            TableMergesLimitKind::PayloadMessages,
        )
    }

    fn charge_items(&mut self, amount: usize) -> Result<(), TableMergesError> {
        // Metadata and selector collections are traversed to validate the
        // rooted projection.  Charge their traversal together with the
        // semantic item ceiling so a large deferred vector cannot evade the
        // aggregate work budget.
        self.charge_work(amount)?;
        Self::add(
            &mut self.items,
            amount,
            self.max_items,
            TableMergesLimitKind::PayloadItems,
        )
    }

    fn charge_references(&mut self, amount: usize) -> Result<(), TableMergesError> {
        Self::add(
            &mut self.references,
            amount,
            self.max_references,
            TableMergesLimitKind::PayloadReferences,
        )
    }

    fn charge_sheets(&mut self, amount: usize) -> Result<(), TableMergesError> {
        Self::add(
            &mut self.sheets,
            amount,
            self.max_sheets,
            TableMergesLimitKind::Sheets,
        )
    }

    fn charge_tables(&mut self, amount: usize) -> Result<(), TableMergesError> {
        Self::add(
            &mut self.tables,
            amount,
            self.max_tables,
            TableMergesLimitKind::Tables,
        )
    }

    fn charge_allocations(&mut self, amount: usize) -> Result<(), TableMergesError> {
        Self::add(
            &mut self.allocations,
            amount,
            self.max_allocations,
            TableMergesLimitKind::Allocations,
        )
    }

    fn charge_retained(&mut self, amount: usize) -> Result<(), TableMergesError> {
        Self::add(
            &mut self.retained,
            amount,
            self.max_retained,
            TableMergesLimitKind::Retained,
        )
    }

    fn charge_scratch(&mut self, amount: usize) -> Result<(), TableMergesError> {
        Self::add(
            &mut self.scratch,
            amount,
            self.max_scratch,
            TableMergesLimitKind::Scratch,
        )
    }

    fn charge_regions(&mut self, amount: usize) -> Result<(), TableMergesError> {
        Self::add(
            &mut self.regions,
            amount,
            self.max_regions,
            TableMergesLimitKind::Regions,
        )
    }

    fn remaining(
        used: usize,
        maximum: usize,
        kind: TableMergesLimitKind,
    ) -> Result<usize, TableMergesError> {
        maximum
            .checked_sub(used)
            .ok_or(TableMergesError::LimitExceeded {
                kind,
                observed: maximum.saturating_add(1) as u64,
                maximum: maximum as u64,
            })
    }

    fn remaining_input(&self) -> Result<usize, TableMergesError> {
        Self::remaining(
            self.input,
            self.wire_max_input,
            TableMergesLimitKind::WireBytes,
        )
    }

    fn remaining_fields(&self) -> Result<usize, TableMergesError> {
        Self::remaining(
            self.fields,
            self.wire_max_fields,
            TableMergesLimitKind::WireFields,
        )
    }

    fn remaining_work(&self) -> Result<usize, TableMergesError> {
        Self::remaining(
            self.work,
            self.wire_max_work,
            TableMergesLimitKind::WireWork,
        )
    }

    fn remaining_allocations(&self) -> Result<usize, TableMergesError> {
        Self::remaining(
            self.allocations,
            self.max_allocations,
            TableMergesLimitKind::Allocations,
        )
    }

    fn remaining_retained(&self) -> Result<usize, TableMergesError> {
        Self::remaining(
            self.retained,
            self.max_retained,
            TableMergesLimitKind::Retained,
        )
    }

    fn remaining_scratch(&self) -> Result<usize, TableMergesError> {
        Self::remaining(
            self.scratch,
            self.max_scratch,
            TableMergesLimitKind::Scratch,
        )
    }

    fn remaining_regions(&self) -> Result<usize, TableMergesError> {
        Self::remaining(
            self.regions,
            self.max_regions,
            TableMergesLimitKind::Regions,
        )
    }

    fn residual_wire(&self) -> Result<WireLimits, TableMergesError> {
        let input = self.remaining_input()?;
        let fields = self.remaining_fields()?;
        let work = self.remaining_work()?;
        if input == 0 {
            return Err(limit(
                TableMergesLimitKind::WireBytes,
                self.wire_max_input.saturating_add(1),
                self.wire_max_input,
            ));
        }
        if fields == 0 {
            return Err(limit(
                TableMergesLimitKind::WireFields,
                self.wire_max_fields.saturating_add(1),
                self.wire_max_fields,
            ));
        }
        if work == 0 {
            return Err(limit(
                TableMergesLimitKind::WireWork,
                self.wire_max_work.saturating_add(1),
                self.wire_max_work,
            ));
        }
        WireLimits::default()
            .with_input_bytes(input)
            .and_then(|limits| limits.with_fields(fields))
            .and_then(|limits| limits.with_nesting(self.wire_max_nesting))
            .and_then(|limits| limits.with_rewrite_work(work))
            .map_err(map_common_error)
    }
}

fn select_semantic_positions(
    package: &Package,
    sheet_selector: SheetSelector<'_>,
    table_selector: TableSelector<'_>,
    budget: &mut Budget,
) -> Result<(usize, usize), TableMergesError> {
    let sheet = match sheet_selector {
        SheetSelector::Name(name) => {
            let mut selected = None;
            for candidate in package.state.document.sheets() {
                budget.charge_sheets(1)?;
                budget.charge_work(candidate.name().len().saturating_add(1))?;
                budget.charge_items(1)?;
                if candidate.name() == name {
                    if selected.is_some() {
                        return Err(TableMergesError::AmbiguousSelector);
                    }
                    selected = Some(candidate);
                }
            }
            selected.ok_or(TableMergesError::SheetNotFound)?
        },
        SheetSelector::Index(index) => {
            budget.charge_sheets(1)?;
            budget.charge_work(1)?;
            package
                .state
                .document
                .sheets()
                .get(index)
                .ok_or(TableMergesError::SheetNotFound)?
        },
    };
    let table_position = match table_selector {
        TableSelector::Index(index) => {
            budget.charge_tables(1)?;
            budget.charge_work(1)?;
            if sheet.tables().nth(index).is_none() {
                return Err(TableMergesError::TableNotFound);
            }
            index
        },
        TableSelector::Name(name) => {
            let mut first = None;
            for (index, candidate) in sheet.tables().enumerate() {
                budget.charge_tables(1)?;
                budget.charge_work(candidate.name().len().saturating_add(1))?;
                budget.charge_items(1)?;
                if candidate.name() != name {
                    continue;
                }
                if first.is_some() {
                    return Err(TableMergesError::AmbiguousSelector);
                }
                first = Some(index);
            }
            let Some(index) = first else {
                return Err(TableMergesError::TableNotFound);
            };
            index
        },
    };
    Ok((sheet.index(), table_position))
}

fn scan_fields<'source, F>(
    source: &'source [u8],
    depth: usize,
    budget: &mut Budget,
    mut visitor: F,
) -> Result<(), TableMergesError>
where
    F: FnMut(RawWireField<'source>, &mut Budget) -> Result<(), TableMergesError>,
{
    if depth > budget.wire_max_nesting {
        return Err(limit(
            TableMergesLimitKind::WireNesting,
            depth,
            budget.wire_max_nesting,
        ));
    }
    budget.charge_input(source.len())?;
    if source.is_empty() {
        return Ok(());
    }

    // Parse one top-level field at a time.  A single RawWireFields iterator
    // snapshots its limits, while visitors below may consume work and fields
    // recursively (for references, nested envelopes, and table metadata).
    // Rebuilding the bounded view from the remaining slice makes every next
    // parse observe the current aggregate residual ceilings.  Input bytes are
    // charged once for this source above; the per-field view therefore uses
    // only the remaining slice length as its local framing bound.
    let mut offset = 0usize;
    while offset < source.len() {
        let remaining = source
            .get(offset..)
            .ok_or(TableMergesError::InvalidSource)?;
        let fields = budget.remaining_fields()?;
        if fields == 0 {
            return Err(limit(
                TableMergesLimitKind::WireFields,
                budget.wire_max_fields.saturating_add(1),
                budget.wire_max_fields,
            ));
        }
        let work = budget.remaining_work()?;
        if work == 0 {
            return Err(limit(
                TableMergesLimitKind::WireWork,
                budget.wire_max_work.saturating_add(1),
                budget.wire_max_work,
            ));
        }
        let limits = RawWireLimits::new(
            remaining.len(),
            fields,
            budget.wire_max_nesting.saturating_sub(depth),
            work,
        )
        .map_err(map_common_error)?;
        let mut parser = RawWireFields::with_limits(remaining, limits);
        let before_fields = parser.fields();
        let before_work = parser.work();
        let field = parser
            .next()
            .map_err(map_common_error)?
            .ok_or(TableMergesError::InvalidSource)?;
        budget.charge_fields(parser.fields().saturating_sub(before_fields))?;
        budget.charge_work(parser.work().saturating_sub(before_work))?;
        let consumed = field.end();
        if consumed == 0 || consumed > remaining.len() {
            return Err(TableMergesError::InvalidSource);
        }
        offset = offset
            .checked_add(consumed)
            .ok_or(TableMergesError::InvalidSource)?;
        visitor(field, budget)?;
    }
    Ok(())
}

fn length_payload<'source>(
    field: RawWireField<'source>,
) -> Result<&'source [u8], TableMergesError> {
    if field.wire_type() != 2 || !field.key_is_canonical() || !field.length_is_canonical() {
        return Err(TableMergesError::InvalidSource);
    }
    Ok(field.payload())
}

fn canonical_varint(field: RawWireField<'_>) -> Result<u64, TableMergesError> {
    if field.wire_type() != 0 || !field.key_is_canonical() || !field.value_is_canonical() {
        return Err(TableMergesError::InvalidSource);
    }
    let (value, width) =
        decode_varint_from_bytes(field.payload()).map_err(|_| TableMergesError::InvalidSource)?;
    if width != field.payload().len() || encoded_len(value) != width {
        return Err(TableMergesError::InvalidSource);
    }
    Ok(value)
}

fn parse_reference(
    source: &[u8],
    depth: usize,
    budget: &mut Budget,
) -> Result<u64, TableMergesError> {
    let mut identifier = None;
    let mut deprecated_type = None;
    let mut external = None;
    scan_fields(source, depth, budget, |field, _budget| {
        match field.number() {
            1 if identifier.is_none() && field.wire_type() == 0 => {
                identifier = Some(canonical_varint(field)?);
            },
            2 if deprecated_type.is_none() && field.wire_type() == 0 => {
                let value = canonical_varint(field)?;
                if value > u64::from(i32::MAX.unsigned_abs()) && value < MIN_SIGN_EXTENDED_I32 {
                    return Err(TableMergesError::InvalidSource);
                }
                deprecated_type = Some(value);
            },
            3 if external.is_none() && field.wire_type() == 0 => {
                let value = canonical_varint(field)?;
                if value > 1 {
                    return Err(TableMergesError::InvalidSource);
                }
                external = Some(value != 0);
            },
            1..=3 => return Err(TableMergesError::InvalidSource),
            _ => {},
        }
        Ok(())
    })?;
    let identifier = identifier.ok_or(TableMergesError::InvalidSource)?;
    if identifier == 0 || external == Some(true) {
        return Err(TableMergesError::InvalidSource);
    }
    Ok(identifier)
}

fn unique_message<'source>(
    object: &'source ArchiveObject,
    message_type: u32,
    budget: &mut Budget,
) -> Result<Option<(usize, &'source RawMessage)>, TableMergesError> {
    budget.charge_messages(object.messages.len())?;
    let mut found = None;
    for (index, message) in object.messages.iter().enumerate() {
        if message.type_ != message_type {
            continue;
        }
        if found.is_some() {
            return Err(TableMergesError::InvalidSource);
        }
        found = Some((index, message));
    }
    Ok(found)
}

fn unique_sheet_message<'source>(
    object: &'source ArchiveObject,
    budget: &mut Budget,
) -> Result<(usize, &'source RawMessage), TableMergesError> {
    let canonical = unique_message(object, SHEET_MESSAGE_TYPE, budget)?;
    let form = unique_message(object, FORM_BASED_SHEET_MESSAGE_TYPE, budget)?;
    match (canonical, form) {
        (Some(_), Some(_)) | (None, None) => Err(TableMergesError::InvalidSource),
        (Some(message), None) | (None, Some(message)) => Ok(message),
    }
}

fn unique_table_info<'source>(
    object: &'source ArchiveObject,
    budget: &mut Budget,
) -> Result<Option<(usize, &'source RawMessage)>, TableMergesError> {
    let canonical = unique_message(object, TABLE_INFO_MESSAGE_TYPE, budget)?;
    let legacy = unique_message(object, LEGACY_TABLE_INFO_MESSAGE_TYPE, budget)?;
    match (canonical, legacy) {
        (Some(_), Some(_)) => Err(TableMergesError::InvalidSource),
        (Some(message), None) | (None, Some(message)) => Ok(Some(message)),
        (None, None) => Ok(None),
    }
}

fn unique_table_model<'source>(
    object: &'source ArchiveObject,
    budget: &mut Budget,
) -> Result<(usize, &'source RawMessage), TableMergesError> {
    let canonical = unique_message(object, TABLE_MODEL_MESSAGE_TYPE, budget)?;
    let legacy = unique_message(object, TABLE_INFO_MESSAGE_TYPE, budget)?;
    match (canonical, legacy) {
        (Some(canonical), Some(legacy)) => {
            // Match the rooted Numbers projection: a historical alias may
            // accompany the authoritative model only with identical bytes.
            budget.charge_work(canonical.1.data.len().min(legacy.1.data.len()))?;
            if canonical.1.data == legacy.1.data {
                Ok(canonical)
            } else {
                Err(TableMergesError::InvalidSource)
            }
        },
        (None, None) => Err(TableMergesError::InvalidSource),
        (Some(message), None) | (None, Some(message)) => Ok(message),
    }
}

fn resolve_model_payload<'source>(
    components: &'source Components,
    index: &'source Index,
    sheet_position: usize,
    table_position: usize,
    budget: &mut Budget,
) -> Result<&'source [u8], TableMergesError> {
    let document_archive = components
        .get_archive("Index/Document.iwa")
        .ok_or(TableMergesError::UnsupportedSource)?;
    // Archive::object performs a linear identifier scan.  Precharge the
    // complete candidate set before entering it so malformed archives cannot
    // spend unbounded lookup work before a finite refusal.
    budget.charge_work(document_archive.objects.len())?;
    let document = document_archive
        .object(1)
        .ok_or(TableMergesError::InvalidSource)?;
    if document.archive_info.identifier != Some(1) {
        return Err(TableMergesError::InvalidSource);
    }
    budget.charge_objects(1)?;
    let (document_index, document_message) =
        unique_message(document, DOCUMENT_MESSAGE_TYPE, budget)?
            .ok_or(TableMergesError::InvalidSource)?;
    validate_message_metadata(document, document_index, budget)?;

    let mut selected_sheet = None;
    let mut sheet_count = 0usize;
    scan_fields(&document_message.data, 0, budget, |field, budget| {
        if field.number() != 1 {
            return Ok(());
        }
        let payload = length_payload(field)?;
        let identifier = read_reference(payload, 1, budget)?;
        if sheet_count == sheet_position {
            selected_sheet = Some(identifier);
        }
        sheet_count = sheet_count
            .checked_add(1)
            .ok_or(TableMergesError::InvalidSource)?;
        Ok(())
    })?;
    let sheet_identifier = selected_sheet.ok_or(TableMergesError::InvalidSource)?;
    require_declared_reference(document, document_index, sheet_identifier, &[1], budget)?;
    let sheet = resolve_object(components, index, sheet_identifier, budget)?;
    let (sheet_message_index, sheet_message) = unique_sheet_message(sheet, budget)?;
    validate_message_metadata(sheet, sheet_message_index, budget)?;

    let mut semantic_table = 0usize;
    let mut selected_model = None;
    match sheet_message.type_ {
        SHEET_MESSAGE_TYPE => scan_fields(&sheet_message.data, 0, budget, |field, budget| {
            if selected_model.is_some() || field.number() != 2 {
                return Ok(());
            }
            let payload = length_payload(field)?;
            let drawable_identifier = read_reference(payload, 1, budget)?;
            selected_model = inspect_drawable(
                components,
                index,
                sheet_identifier,
                drawable_identifier,
                table_position,
                &mut semantic_table,
                budget,
            )?;
            if selected_model.is_some() {
                require_declared_reference(
                    sheet,
                    sheet_message_index,
                    drawable_identifier,
                    &[2],
                    budget,
                )?;
            }
            Ok(())
        })?,
        FORM_BASED_SHEET_MESSAGE_TYPE => {
            let mut super_payload = None;
            scan_fields(&sheet_message.data, 0, budget, |field, _budget| {
                if field.number() != 1 {
                    return Ok(());
                }
                if super_payload.is_some() {
                    return Err(TableMergesError::InvalidSource);
                }
                super_payload = Some(length_payload(field)?);
                Ok(())
            })?;
            let super_payload = super_payload.ok_or(TableMergesError::InvalidSource)?;
            scan_fields(super_payload, 1, budget, |field, budget| {
                if selected_model.is_some() || field.number() != 2 {
                    return Ok(());
                }
                let payload = length_payload(field)?;
                let drawable_identifier = read_reference(payload, 2, budget)?;
                selected_model = inspect_drawable(
                    components,
                    index,
                    sheet_identifier,
                    drawable_identifier,
                    table_position,
                    &mut semantic_table,
                    budget,
                )?;
                if selected_model.is_some() {
                    require_declared_reference(
                        sheet,
                        sheet_message_index,
                        drawable_identifier,
                        &[1, 2],
                        budget,
                    )?;
                }
                Ok(())
            })?;
        },
        _ => return Err(TableMergesError::InvalidSource),
    }
    selected_model.ok_or(TableMergesError::TableNotFound)
}

fn inspect_drawable<'source>(
    components: &'source Components,
    index: &'source Index,
    sheet_identifier: u64,
    drawable_identifier: u64,
    table_position: usize,
    semantic_table: &mut usize,
    budget: &mut Budget,
) -> Result<Option<&'source [u8]>, TableMergesError> {
    let info_object = resolve_object(components, index, drawable_identifier, budget)?;
    let Some((info_index, info_message)) = unique_table_info(info_object, budget)? else {
        return Ok(None);
    };
    validate_message_metadata(info_object, info_index, budget)?;
    if info_object.archive_info.identifier != Some(drawable_identifier) {
        return Err(TableMergesError::InvalidSource);
    }
    if *semantic_table != table_position {
        *semantic_table = semantic_table
            .checked_add(1)
            .ok_or(TableMergesError::InvalidSource)?;
        return Ok(None);
    }
    let model_identifier = table_model_identifier(&info_message.data, budget)?;
    require_declared_reference(info_object, info_index, model_identifier, &[2], budget)?;
    if sheet_identifier == drawable_identifier
        || sheet_identifier == model_identifier
        || drawable_identifier == model_identifier
    {
        return Err(TableMergesError::InvalidSource);
    }
    let model_object = resolve_object(components, index, model_identifier, budget)?;
    let (model_index, model_message) = unique_table_model(model_object, budget)?;
    validate_message_metadata(model_object, model_index, budget)?;
    if model_object.archive_info.identifier != Some(model_identifier) {
        return Err(TableMergesError::InvalidSource);
    }
    Ok(Some(model_message.data.as_slice()))
}

fn resolve_object<'source>(
    components: &'source Components,
    index: &'source Index,
    identifier: u64,
    budget: &mut Budget,
) -> Result<&'source ArchiveObject, TableMergesError> {
    budget.charge_objects(1)?;
    budget.charge_work(index.lookup_work())?;
    let resolved: Resolved<'source> = index
        .resolve_ref_id(components, identifier)
        .map_err(map_package_error)?
        .ok_or(TableMergesError::InvalidSource)?;
    components
        .catalog()
        .get_index(resolved.component_index)
        .and_then(|component| component.archive().objects.get(resolved.object_index))
        .filter(|object| object.archive_info.identifier == Some(identifier))
        .ok_or(TableMergesError::InvalidSource)
}

fn read_reference(
    source: &[u8],
    depth: usize,
    budget: &mut Budget,
) -> Result<u64, TableMergesError> {
    budget.charge_references(1)?;
    parse_reference(source, depth, budget)
}

fn require_declared_reference(
    object: &ArchiveObject,
    message_index: usize,
    identifier: u64,
    accepted_path: &[u32],
    budget: &mut Budget,
) -> Result<(), TableMergesError> {
    let info = object
        .archive_info
        .message_infos
        .get(message_index)
        .ok_or(TableMergesError::InvalidSource)?;
    // Both the aggregate reference list and every FieldInfo list are
    // linearly searched.  Charge the complete traversal before looking at a
    // candidate so a zero residual work budget cannot enter either loop.
    budget.charge_work(info.object_references.len())?;
    if info
        .object_references
        .iter()
        .filter(|candidate| **candidate == identifier)
        .count()
        != 1
    {
        return Err(TableMergesError::InvalidSource);
    }
    budget.charge_work(info.field_infos.len())?;
    let mut field_occurrence = false;
    for field in &info.field_infos {
        budget.charge_work(field.object_references.len())?;
        let count = field
            .object_references
            .iter()
            .filter(|candidate| **candidate == identifier)
            .count();
        if count == 0 {
            continue;
        }
        budget.charge_work(field.path.path.len())?;
        if count != 1 || field_occurrence || field.path.path.as_slice() != accepted_path {
            return Err(TableMergesError::InvalidSource);
        }
        field_occurrence = true;
    }
    // Some Apple objects declare rooted edges only in the aggregate object
    // reference list while retaining unrelated FieldInfo records.  The
    // aggregate occurrence above is authoritative in that representation.
    budget.charge_work(1)
}

fn table_model_identifier(source: &[u8], budget: &mut Budget) -> Result<u64, TableMergesError> {
    let mut super_payload = None;
    let mut model_payload = None;
    scan_fields(source, 0, budget, |field, _budget| {
        match field.number() {
            1 => {
                if super_payload.is_some() {
                    return Err(TableMergesError::InvalidSource);
                }
                super_payload = Some(length_payload(field)?);
            },
            2 => {
                if model_payload.is_some() {
                    return Err(TableMergesError::InvalidSource);
                }
                model_payload = Some(length_payload(field)?);
            },
            _ => {},
        }
        Ok(())
    })?;
    let super_payload = super_payload.ok_or(TableMergesError::InvalidSource)?;
    validate_table_super(super_payload, budget)?;
    let model_payload = model_payload.ok_or(TableMergesError::InvalidSource)?;
    read_reference(model_payload, 1, budget)
}

fn validate_table_super(source: &[u8], budget: &mut Budget) -> Result<(), TableMergesError> {
    let mut locked = None;
    scan_fields(source, 1, budget, |field, _budget| {
        if field.number() != 5 {
            return Ok(());
        }
        if locked.is_some() {
            return Err(TableMergesError::InvalidSource);
        }
        let value = canonical_varint(field)?;
        if value > 1 {
            return Err(TableMergesError::InvalidSource);
        }
        locked = Some(value != 0);
        Ok(())
    })?;
    Ok(())
}

fn validate_message_metadata(
    object: &ArchiveObject,
    message_index: usize,
    budget: &mut Budget,
) -> Result<(), TableMergesError> {
    if object.archive_info.message_infos.len() != object.messages.len() {
        return Err(TableMergesError::InvalidSource);
    }
    let info = object
        .archive_info
        .message_infos
        .get(message_index)
        .ok_or(TableMergesError::InvalidSource)?;
    let message = object
        .messages
        .get(message_index)
        .ok_or(TableMergesError::InvalidSource)?;
    charge_metadata(object, budget)?;
    if info.type_ != message.type_
        || object.archive_info.should_merge == Some(true)
        || info.base_message_index.is_some()
        || !info.diff_merge_version.is_empty()
        || info.diff_field_path.is_some()
        || !info.fields_to_remove.is_empty()
        || !info.diff_read_version.is_empty()
    {
        return Err(TableMergesError::InvalidSource);
    }
    Ok(())
}

fn charge_metadata(object: &ArchiveObject, budget: &mut Budget) -> Result<(), TableMergesError> {
    // Charge each deferred collection before walking it.  Archive metadata is
    // already resident in the package model, but its vectors can still be
    // attacker-controlled in a malformed fixture.  Incremental accounting
    // keeps a later oversized collection from being traversed under a stale
    // aggregate snapshot and makes refusal happen before that work starts.
    budget.charge_items(object.archive_info.message_infos.len())?;
    for info in &object.archive_info.message_infos {
        budget.charge_items(info.versions.len())?;
        budget.charge_items(info.field_infos.len())?;
        budget.charge_items(info.diff_merge_version.len())?;
        budget.charge_items(info.fields_to_remove.len())?;
        budget.charge_items(info.diff_read_version.len())?;
        budget.charge_references(info.object_references.len())?;
        budget.charge_references(info.data_references.len())?;
        budget.charge_items(info.object_references.len())?;
        budget.charge_items(info.data_references.len())?;
        if let Some(path) = &info.diff_field_path {
            budget.charge_items(1)?;
            budget.charge_items(path.path.len())?;
        }
        for field_info in &info.field_infos {
            budget.charge_items(1)?;
            budget.charge_items(field_info.path.path.len())?;
            budget.charge_items(field_info.object_references.len())?;
            budget.charge_items(field_info.data_references.len())?;
            budget.charge_items(field_info.known_field_version.len())?;
            budget.charge_items(usize::from(
                field_info.known_field_feature_identifier.is_some(),
            ))?;
            budget.charge_references(field_info.object_references.len())?;
            budget.charge_references(field_info.data_references.len())?;
        }
        for path in &info.fields_to_remove {
            budget.charge_items(path.path.len())?;
        }
    }
    Ok(())
}

fn merge_region_capacity(
    allocations: usize,
    scratch: usize,
    retained: usize,
    regions: usize,
) -> usize {
    allocations
        .checked_div(MERGE_READER_ALLOCATIONS_PER_REGION)
        .unwrap_or(0)
        .min(
            scratch
                .checked_div(MERGE_READER_SCRATCH_PER_REGION)
                .unwrap_or(0),
        )
        .min(retained.checked_div(REGION_BYTES).unwrap_or(0))
        .min(regions)
        .min(WireLimits::MAX_FIELDS)
}

fn merge_wire_limits(
    budget: &Budget,
    max_regions: usize,
) -> Result<merge_wire::ReadLimits, TableMergesError> {
    let wire = budget.residual_wire()?;
    Ok(merge_wire::ReadLimits {
        wire,
        max_regions,
        max_overlap_checks: budget.remaining_work()?,
    })
}

fn charge_attempted(budget: &mut Budget, error: &merge_wire::MergeReadError) {
    let attempted = error.attempted();
    let _ = budget.charge_input(attempted.input_bytes);
    let _ = budget.charge_fields(attempted.fields);
    let _ = budget.charge_work(attempted.work);
}

fn map_merge_error(error: merge_wire::MergeReadError) -> TableMergesError {
    match error.error() {
        litchi_iwa_common::Error::LimitExceeded {
            kind,
            observed,
            limit,
        } => TableMergesError::LimitExceeded {
            kind: map_common_limit(*kind),
            observed: *observed as u64,
            maximum: *limit as u64,
        },
        litchi_iwa_common::Error::Allocation { amount, .. } => {
            TableMergesError::Allocation { amount: *amount }
        },
        litchi_iwa_common::Error::InvalidFormat(_)
        | litchi_iwa_common::Error::InvalidLimit { .. } => TableMergesError::InvalidSource,
    }
}

fn map_common_limit(kind: CommonLimitKind) -> TableMergesLimitKind {
    match kind {
        CommonLimitKind::InputBytes => TableMergesLimitKind::WireBytes,
        CommonLimitKind::Fields => TableMergesLimitKind::WireFields,
        CommonLimitKind::OutputBytes => TableMergesLimitKind::WireOutputBytes,
        CommonLimitKind::Nesting => TableMergesLimitKind::WireNesting,
        CommonLimitKind::RewriteWork => TableMergesLimitKind::WireWork,
        CommonLimitKind::TableRows
        | CommonLimitKind::TableColumns
        | CommonLimitKind::TableCells
        | CommonLimitKind::MaterializedCells => TableMergesLimitKind::Regions,
    }
}

fn map_common_error(error: litchi_iwa_common::Error) -> TableMergesError {
    match error {
        litchi_iwa_common::Error::LimitExceeded {
            kind,
            observed,
            limit,
        } => TableMergesError::LimitExceeded {
            kind: map_common_limit(kind),
            observed: observed as u64,
            maximum: limit as u64,
        },
        litchi_iwa_common::Error::Allocation { amount, .. } => {
            TableMergesError::Allocation { amount }
        },
        litchi_iwa_common::Error::InvalidFormat(_)
        | litchi_iwa_common::Error::InvalidLimit { .. } => TableMergesError::InvalidSource,
    }
}

fn map_archive_error(error: litchi_iwa_archive::Error) -> TableMergesError {
    match error {
        litchi_iwa_archive::Error::Limit {
            kind,
            observed,
            maximum,
        } => TableMergesError::LimitExceeded {
            kind: match kind {
                litchi_iwa_archive::LimitKind::InputBytes => TableMergesLimitKind::InputBytes,
                litchi_iwa_archive::LimitKind::OutputBytes => TableMergesLimitKind::OutputBytes,
                litchi_iwa_archive::LimitKind::Entries => TableMergesLimitKind::Entries,
                litchi_iwa_archive::LimitKind::MemberNameBytes
                | litchi_iwa_archive::LimitKind::MetadataBytes => {
                    TableMergesLimitKind::PackageBytes
                },
                litchi_iwa_archive::LimitKind::CompressedEntryBytes
                | litchi_iwa_archive::LimitKind::EntryBytes => TableMergesLimitKind::EntryBytes,
                litchi_iwa_archive::LimitKind::TotalBytes => TableMergesLimitKind::TotalEntryBytes,
                litchi_iwa_archive::LimitKind::IwaStreamBytes => TableMergesLimitKind::PayloadBytes,
                litchi_iwa_archive::LimitKind::IwaTotalBytes => {
                    TableMergesLimitKind::TotalPayloadBytes
                },
            },
            observed,
            maximum,
        },
        litchi_iwa_archive::Error::Allocation { amount, .. } => {
            TableMergesError::Allocation { amount }
        },
        litchi_iwa_archive::Error::Iwa(error) => map_core_error(error),
        litchi_iwa_archive::Error::Reassembly(_) => TableMergesError::UnsupportedSource,
        _ => TableMergesError::InvalidSource,
    }
}

fn map_core_error(error: litchi_iwa_core::Error) -> TableMergesError {
    match error {
        litchi_iwa_core::Error::Limit {
            kind,
            observed,
            maximum,
        } => TableMergesError::LimitExceeded {
            kind: match kind {
                litchi_iwa_core::LimitKind::Objects => TableMergesLimitKind::PayloadObjects,
                litchi_iwa_core::LimitKind::Messages
                | litchi_iwa_core::LimitKind::MessagesPerObject => {
                    TableMergesLimitKind::PayloadMessages
                },
                litchi_iwa_core::LimitKind::HeaderNesting => TableMergesLimitKind::WireNesting,
                litchi_iwa_core::LimitKind::HeaderFields
                | litchi_iwa_core::LimitKind::MetadataItems
                | litchi_iwa_core::LimitKind::SnappyFrames => TableMergesLimitKind::PayloadItems,
                _ => TableMergesLimitKind::PayloadBytes,
            },
            observed: observed as u64,
            maximum: maximum as u64,
        },
        litchi_iwa_core::Error::Allocation { requested, .. } => {
            TableMergesError::Allocation { amount: requested }
        },
        _ => TableMergesError::InvalidSource,
    }
}

fn map_package_error(error: PackageError) -> TableMergesError {
    match error {
        PackageError::Archive(error) => map_archive_error(error),
        PackageError::Common(error) => map_common_error(error),
        PackageError::SemanticLimit {
            kind,
            observed,
            maximum,
            ..
        } => TableMergesError::LimitExceeded {
            kind: match kind {
                SemanticLimitKind::Objects => TableMergesLimitKind::PayloadObjects,
                SemanticLimitKind::References => TableMergesLimitKind::PayloadReferences,
                SemanticLimitKind::OutputTextBytes
                | SemanticLimitKind::FormulaWireBytes
                | SemanticLimitKind::TextBytes => TableMergesLimitKind::PayloadBytes,
                SemanticLimitKind::FormulaRenderDepth | SemanticLimitKind::FormulaDepth => {
                    TableMergesLimitKind::WireNesting
                },
                SemanticLimitKind::FormulaRenderWork | SemanticLimitKind::FormulaWork => {
                    TableMergesLimitKind::WireWork
                },
                SemanticLimitKind::Sheets => TableMergesLimitKind::Sheets,
                SemanticLimitKind::Tables => TableMergesLimitKind::Tables,
                SemanticLimitKind::MaterializedCells => TableMergesLimitKind::PayloadItems,
            },
            observed: observed as u64,
            maximum: maximum as u64,
        },
        PackageError::InputTooLarge { observed, maximum } => TableMergesError::LimitExceeded {
            kind: TableMergesLimitKind::InputBytes,
            observed,
            maximum,
        },
        PackageError::NotNumbers => TableMergesError::UnsupportedSource,
        PackageError::Io(_)
        | PackageError::Detection(_)
        | PackageError::MalformedPayload { .. }
        | PackageError::InvalidFormat(_)
        | PackageError::ParseError(_)
        | PackageError::Semantic(_) => TableMergesError::InvalidSource,
        #[allow(unreachable_patterns)]
        _ => TableMergesError::InvalidSource,
    }
}

fn limit(kind: TableMergesLimitKind, observed: usize, maximum: usize) -> TableMergesError {
    TableMergesError::LimitExceeded {
        kind,
        observed: observed as u64,
        maximum: maximum as u64,
    }
}

#[cfg(test)]
#[path = "table_merges/budget_tests.rs"]
mod budget_tests;

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn malformed_merge_wire_maps_to_content_free_source_error() {
        let error = merge_wire::read_table_merges(&[0x2f], merge_wire::ReadLimits::default())
            .expect_err("truncated wire must fail");
        assert_eq!(map_merge_error(error), TableMergesError::InvalidSource);
    }

    #[test]
    fn result_capacity_is_bounded_by_each_result_resource() {
        assert_eq!(
            merge_region_capacity(2, usize::MAX, usize::MAX, usize::MAX),
            0
        );
        assert_eq!(
            merge_region_capacity(
                MERGE_READER_ALLOCATIONS_PER_REGION * 4,
                MERGE_READER_SCRATCH_PER_REGION * 4,
                REGION_BYTES * 4,
                4,
            ),
            4
        );
    }

    #[test]
    fn malformed_reference_is_rejected_without_allocating_a_view() {
        let mut budget = Budget {
            wire_max_input: 64,
            wire_max_fields: 64,
            wire_max_work: 256,
            wire_max_nesting: 4,
            max_objects: 1,
            max_sheets: 1,
            max_tables: 1,
            max_messages: 1,
            max_items: 1,
            max_references: 1,
            max_allocations: 1,
            max_retained: 1,
            max_scratch: 1,
            max_regions: 1,
            input: 0,
            fields: 0,
            work: 0,
            objects: 0,
            sheets: 0,
            tables: 0,
            messages: 0,
            items: 0,
            references: 0,
            allocations: 0,
            retained: 0,
            scratch: 0,
            regions: 0,
        };
        assert_eq!(
            parse_reference(&[0x08, 0x00], 0, &mut budget),
            Err(TableMergesError::InvalidSource)
        );
    }
}
