//! Bounded borrowed reader for native Numbers merged-cell storage.
//!
//! The decoder owns only the selected `TableModelArchive.merge_owner` wire
//! path. It returns archive-free common geometry after validating the native
//! merge formula shape; callers retain package identity and transaction state.

use core::fmt;
use core::mem::size_of;
use std::str;

use litchi_iwa_common::table::merge::Region;
use litchi_iwa_common::wire::{RawWireField, RawWireFields, RawWireLimits};
use litchi_iwa_common::{Error, LimitKind, WireLimits};
use litchi_iwa_protos::numbers_formula_codec::{
    self, FormulaRenderCfuuid, FormulaRenderColonTract, FormulaRenderEvent, FormulaRenderVisitor,
};

const TABLE_MODEL_TABLE_ID_FIELD: u32 = 1;
const TABLE_MODEL_ROWS_FIELD: u32 = 6;
const TABLE_MODEL_COLUMNS_FIELD: u32 = 7;
const TABLE_MODEL_MERGE_OWNER_FIELD: u32 = 47;
const MERGE_OWNER_FORMULA_STORE_FIELD: u32 = 2;
const FORMULA_STORE_NEXT_INDEX_FIELD: u32 = 2;
const FORMULA_STORE_FORMULAS_FIELD: u32 = 3;
const FORMULA_PAIR_INDEX_FIELD: u32 = 1;
const FORMULA_PAIR_FORMULA_FIELD: u32 = 2;
const NATIVE_MERGE_FUNCTION_INDEX: u32 = 168;
const NATIVE_MERGE_FUNCTION_ARGUMENTS: u32 = 1;

/// Temporary borrowed pair-reference storage required by a merge read.
pub const PAIR_REFERENCE_BYTES: usize = size_of::<&[u8]>();
/// Temporary formula-index storage required by a merge read.
pub const FORMULA_INDEX_BYTES: usize = size_of::<u32>();
/// Retained archive-free region storage required by a merge read.
pub const REGION_BYTES: usize = size_of::<Region>();

/// Read limits for one selected table-model merge path.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ReadLimits {
    /// Aggregate common wire limits for selected envelope scans.
    pub wire: WireLimits,
    /// Maximum number of staged merged regions.
    pub max_regions: usize,
    /// Maximum pairwise overlap checks.
    pub max_overlap_checks: usize,
}

impl Default for ReadLimits {
    fn default() -> Self {
        Self {
            wire: WireLimits::default(),
            max_regions: WireLimits::MAX_FIELDS,
            max_overlap_checks: WireLimits::MAX_REWRITE_WORK,
        }
    }
}

/// Aggregate resource totals from one successful merge read.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct ReadReport {
    input_bytes: usize,
    fields: usize,
    work: usize,
}

impl ReadReport {
    /// Total selected-message bytes scanned, including formula decodes.
    #[must_use]
    pub const fn input_bytes(self) -> usize {
        self.input_bytes
    }

    /// Total selected wire and formula fields inspected.
    #[must_use]
    pub const fn fields(self) -> usize {
        self.fields
    }

    /// Total bounded scanner and formula work charged.
    #[must_use]
    pub const fn work(self) -> usize {
        self.work
    }
}

/// Successful archive-free merge projection.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct MergeRead {
    /// Validated non-overlapping merged-cell rectangles in source order.
    pub regions: Vec<Region>,
    /// Cumulative selected-wire resource report.
    pub report: ReadReport,
}

/// Cost observed before a merge read succeeds or refuses.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct AttemptedCost {
    /// Selected-message bytes scanned.
    pub input_bytes: usize,
    /// Selected wire and formula fields inspected.
    pub fields: usize,
    /// Bounded scanner and formula work charged.
    pub work: usize,
}

impl From<ReadReport> for AttemptedCost {
    fn from(report: ReadReport) -> Self {
        Self {
            input_bytes: report.input_bytes,
            fields: report.fields,
            work: report.work,
        }
    }
}

/// Failure while reading native merge storage.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct MergeReadError {
    error: Error,
    attempted: AttemptedCost,
}

impl MergeReadError {
    /// Return the underlying common wire failure.
    #[must_use]
    pub const fn error(&self) -> &Error {
        &self.error
    }

    /// Return work observed before refusal.
    #[must_use]
    pub const fn attempted(&self) -> AttemptedCost {
        self.attempted
    }
}

impl fmt::Display for MergeReadError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        self.error.fmt(formatter)
    }
}

impl std::error::Error for MergeReadError {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        Some(&self.error)
    }
}

/// Read and validate the native Numbers merge-owner path.
///
/// The table UUID is read only when a formula store is present. Its native
/// formula-owner byte swap is kept private to this wire boundary, while the
/// returned rectangles use the archive-free common geometry type.
pub fn read_table_merges(source: &[u8], limits: ReadLimits) -> Result<MergeRead, MergeReadError> {
    let mut decoder = Decoder::new(limits);
    if let Err(error) = decoder.validate_limits() {
        return Err(decoder.failure(error));
    }

    let mut model = ModelFields::default();
    if let Err(error) = decoder.scan(source, 0, |field| model.visit(field)) {
        return Err(decoder.failure(error));
    }
    let Some(owner) = model.merge_owner else {
        return Ok(decoder.success());
    };
    let store = match decoder.scan_owner(owner, 1) {
        Ok(store) => store,
        Err(error) => return Err(decoder.failure(error)),
    };
    let Some(store) = store else {
        return Ok(decoder.success());
    };
    let table_uuid = match model.table_id {
        Some(table_id) => match parse_table_uuid(table_id) {
            Ok(table_uuid) => table_uuid,
            Err(error) => return Err(decoder.failure(error)),
        },
        None => return Err(decoder.failure(invalid("table model merge owner is missing table_id"))),
    };
    if let Err(error) = decoder.scan_store(
        store,
        table_uuid.owner_cfuuid_words(),
        model.rows,
        model.columns,
        2,
    ) {
        return Err(decoder.failure(error));
    }
    Ok(decoder.success())
}

#[derive(Debug, Default)]
struct ModelFields<'source> {
    table_id: Option<&'source [u8]>,
    rows: Option<u32>,
    columns: Option<u32>,
    merge_owner: Option<&'source [u8]>,
}

impl<'source> ModelFields<'source> {
    fn visit(&mut self, field: RawWireField<'source>) -> Result<(), Error> {
        match field.number() {
            TABLE_MODEL_TABLE_ID_FIELD => {
                if self.table_id.is_some() {
                    return Err(duplicate("TableModelArchive.table_id"));
                }
                let table_id = length_payload(field, "table_id")?;
                str::from_utf8(table_id)
                    .map_err(|_| invalid("table model table_id is not UTF-8"))?;
                self.table_id = Some(table_id);
            },
            TABLE_MODEL_ROWS_FIELD => {
                if self.rows.is_some() {
                    return Err(duplicate("TableModelArchive.number_of_rows"));
                }
                self.rows = Some(varint_u32(field, "number_of_rows")?);
            },
            TABLE_MODEL_COLUMNS_FIELD => {
                if self.columns.is_some() {
                    return Err(duplicate("TableModelArchive.number_of_columns"));
                }
                self.columns = Some(varint_u32(field, "number_of_columns")?);
            },
            TABLE_MODEL_MERGE_OWNER_FIELD => {
                if self.merge_owner.is_some() {
                    return Err(duplicate("TableModelArchive.merge_owner"));
                }
                self.merge_owner = Some(length_payload(field, "merge_owner")?);
            },
            _ => {},
        }
        Ok(())
    }
}

#[derive(Debug, Clone, Copy)]
struct TableUuid {
    lower: u64,
    upper: u64,
}

impl TableUuid {
    const fn owner_cfuuid_words(self) -> [u32; 4] {
        let owner_lower = self.upper.swap_bytes();
        let owner_upper = self.lower.swap_bytes();
        [
            owner_lower as u32,
            (owner_lower >> 32) as u32,
            owner_upper as u32,
            (owner_upper >> 32) as u32,
        ]
    }
}

fn parse_table_uuid(source: &[u8]) -> Result<TableUuid, Error> {
    let mut value = 0u128;
    let mut digits = 0usize;
    for byte in source.iter().copied() {
        if byte == b'-' {
            continue;
        }
        let nibble = match byte {
            b'0'..=b'9' => u128::from(byte - b'0'),
            b'a'..=b'f' => u128::from(byte - b'a' + 10),
            b'A'..=b'F' => u128::from(byte - b'A' + 10),
            _ => return Err(invalid("iWork table UUID is malformed")),
        };
        digits = digits
            .checked_add(1)
            .ok_or_else(|| invalid("iWork table UUID digit count overflow"))?;
        if digits > 32 {
            return Err(invalid("iWork table UUID is malformed"));
        }
        value = (value << 4) | nibble;
    }
    if digits != 32 {
        return Err(invalid("iWork table UUID is malformed"));
    }
    Ok(TableUuid {
        lower: value as u64,
        upper: (value >> 64) as u64,
    })
}

struct Decoder {
    limits: ReadLimits,
    report: ReadReport,
    regions: Vec<Region>,
    indexes: Vec<u32>,
    overlap_checks: usize,
}

impl Decoder {
    fn new(limits: ReadLimits) -> Self {
        Self {
            limits,
            report: ReadReport::default(),
            regions: Vec::new(),
            indexes: Vec::new(),
            overlap_checks: 0,
        }
    }

    fn validate_limits(&self) -> Result<(), Error> {
        if self.limits.max_regions > WireLimits::MAX_FIELDS {
            return Err(Error::InvalidLimit {
                field: "table merge regions",
                value: self.limits.max_regions,
                maximum: WireLimits::MAX_FIELDS,
            });
        }
        if self.limits.max_overlap_checks > WireLimits::MAX_REWRITE_WORK {
            return Err(Error::InvalidLimit {
                field: "table merge overlap checks",
                value: self.limits.max_overlap_checks,
                maximum: WireLimits::MAX_REWRITE_WORK,
            });
        }
        Ok(())
    }

    fn success(self) -> MergeRead {
        MergeRead {
            regions: self.regions,
            report: self.report,
        }
    }

    fn failure(&self, error: Error) -> MergeReadError {
        MergeReadError {
            error,
            attempted: self.report.into(),
        }
    }

    fn scan<'source, F>(
        &mut self,
        source: &'source [u8],
        depth: usize,
        mut visitor: F,
    ) -> Result<(), Error>
    where
        F: FnMut(RawWireField<'source>) -> Result<(), Error>,
    {
        let remaining_nesting = self.remaining_nesting(depth)?;
        self.charge_input(source.len())?;
        let remaining_fields = self
            .limits
            .wire
            .max_fields()
            .checked_sub(self.report.fields)
            .ok_or_else(|| {
                limit(
                    LimitKind::Fields,
                    self.report.fields,
                    self.limits.wire.max_fields(),
                )
            })?;
        let remaining_work = self
            .limits
            .wire
            .max_rewrite_work()
            .checked_sub(self.report.work)
            .ok_or_else(|| {
                limit(
                    LimitKind::RewriteWork,
                    self.report.work,
                    self.limits.wire.max_rewrite_work(),
                )
            })?;
        let raw_limits = RawWireLimits::new(
            source.len().max(1),
            remaining_fields.max(1),
            remaining_nesting,
            remaining_work.max(1),
        )?;
        let mut fields = RawWireFields::with_limits(source, raw_limits);
        loop {
            let before_fields = fields.fields();
            let before_work = fields.work();
            let next = fields.next();
            self.charge_fields(fields.fields().saturating_sub(before_fields))?;
            self.charge_work(fields.work().saturating_sub(before_work))?;
            let Some(field) = next? else {
                break;
            };
            visitor(field)?;
        }
        Ok(())
    }

    fn charge_input(&mut self, amount: usize) -> Result<(), Error> {
        let observed = self
            .report
            .input_bytes
            .checked_add(amount)
            .ok_or_else(|| invalid("table merge input-byte count overflow"))?;
        if observed > self.limits.wire.max_input_bytes() {
            return Err(limit(
                LimitKind::InputBytes,
                observed,
                self.limits.wire.max_input_bytes(),
            ));
        }
        self.report.input_bytes = observed;
        Ok(())
    }

    fn remaining_nesting(&self, depth: usize) -> Result<usize, Error> {
        self.limits
            .wire
            .max_nesting()
            .checked_sub(depth)
            .ok_or_else(|| limit(LimitKind::Nesting, depth, self.limits.wire.max_nesting()))
    }

    fn charge_fields(&mut self, amount: usize) -> Result<(), Error> {
        let observed = self
            .report
            .fields
            .checked_add(amount)
            .ok_or_else(|| invalid("table merge field count overflow"))?;
        if observed > self.limits.wire.max_fields() {
            return Err(limit(
                LimitKind::Fields,
                observed,
                self.limits.wire.max_fields(),
            ));
        }
        self.report.fields = observed;
        Ok(())
    }

    fn charge_work(&mut self, amount: usize) -> Result<(), Error> {
        let observed = self
            .report
            .work
            .checked_add(amount)
            .ok_or_else(|| invalid("table merge work count overflow"))?;
        if observed > self.limits.wire.max_rewrite_work() {
            return Err(limit(
                LimitKind::RewriteWork,
                observed,
                self.limits.wire.max_rewrite_work(),
            ));
        }
        self.report.work = observed;
        Ok(())
    }

    fn scan_owner<'source>(
        &mut self,
        source: &'source [u8],
        depth: usize,
    ) -> Result<Option<&'source [u8]>, Error> {
        let mut store = None;
        let mut owner_uid = None;
        self.scan(source, depth, |field| {
            match field.number() {
                1 => {
                    if owner_uid.is_some() {
                        return Err(duplicate("MergeOwnerArchive.formula_owner_uid"));
                    }
                    owner_uid = Some(length_payload(field, "formula_owner_uid")?);
                },
                MERGE_OWNER_FORMULA_STORE_FIELD => {
                    if store.is_some() {
                        return Err(duplicate("MergeOwnerArchive.formula_store"));
                    }
                    store = Some(length_payload(field, "formula_store")?);
                },
                _ => {},
            }
            Ok(())
        })?;
        let owner_uid =
            owner_uid.ok_or_else(|| invalid("MergeOwnerArchive is missing required owner_id"))?;
        self.scan_cfuuid(
            owner_uid,
            depth
                .checked_add(1)
                .ok_or_else(|| invalid("table merge owner UUID depth overflow"))?,
        )?;
        Ok(store)
    }

    fn scan_cfuuid(&mut self, source: &[u8], depth: usize) -> Result<(), Error> {
        let mut seen = [false; 6];
        self.scan(source, depth, |field| match field.number() {
            1 => {
                singular(&mut seen, 1, "CFUUIDArchive.uuid_bytes")?;
                length_payload(field, "CFUUIDArchive.uuid_bytes")?;
                Ok(())
            },
            2..=5 => {
                let slot = field.number() as usize;
                singular(&mut seen, slot, "CFUUIDArchive.word")?;
                varint_u32(field, "CFUUIDArchive.word")?;
                Ok(())
            },
            _ => Ok(()),
        })
    }

    fn scan_store(
        &mut self,
        source: &[u8],
        expected_table: [u32; 4],
        rows: Option<u32>,
        columns: Option<u32>,
        depth: usize,
    ) -> Result<(), Error> {
        let mut next_index = None;
        let mut pair_count = 0usize;
        let max_regions = self.limits.max_regions;
        self.scan(source, depth, |field| match field.number() {
            FORMULA_STORE_NEXT_INDEX_FIELD => {
                if next_index.is_some() {
                    return Err(duplicate("FormulaStoreArchive.next_formula_index"));
                }
                next_index = Some(varint_u32(field, "next_formula_index")?);
                Ok(())
            },
            FORMULA_STORE_FORMULAS_FIELD => {
                length_payload(field, "formula pair")?;
                pair_count = pair_count
                    .checked_add(1)
                    .ok_or_else(|| invalid("table merge formula-pair count overflow"))?;
                if pair_count > max_regions {
                    return Err(limit(LimitKind::TableCells, pair_count, max_regions));
                }
                Ok(())
            },
            _ => Ok(()),
        })?;

        let next_index = next_index
            .ok_or_else(|| invalid("FormulaStoreArchive is missing next_formula_index"))?;
        if pair_count == 0 {
            return Ok(());
        }
        let rows = rows.ok_or_else(|| invalid("table model is missing number_of_rows"))?;
        let columns = columns.ok_or_else(|| invalid("table model is missing number_of_columns"))?;

        self.indexes
            .try_reserve_exact(pair_count)
            .map_err(|_| Error::Allocation {
                resource: "table merge formula indexes",
                amount: self.indexes.len().saturating_add(pair_count),
            })?;
        self.regions
            .try_reserve_exact(pair_count)
            .map_err(|_| Error::Allocation {
                resource: "table merge regions",
                amount: self.regions.len().saturating_add(pair_count),
            })?;
        let mut pairs = Vec::new();
        pairs
            .try_reserve_exact(pair_count)
            .map_err(|_| Error::Allocation {
                resource: "table merge pair references",
                amount: pair_count,
            })?;
        let mut second_pair_count = 0usize;
        let mut second_next_index = None;
        self.scan(source, depth, |field| match field.number() {
            FORMULA_STORE_NEXT_INDEX_FIELD => {
                if second_next_index.is_some() {
                    return Err(duplicate("FormulaStoreArchive.next_formula_index"));
                }
                second_next_index = Some(varint_u32(field, "next_formula_index")?);
                Ok(())
            },
            FORMULA_STORE_FORMULAS_FIELD => {
                let payload = length_payload(field, "formula pair")?;
                if second_pair_count >= pair_count {
                    return Err(invalid("table merge formula-pair count changed"));
                }
                second_pair_count += 1;
                pairs.push(payload);
                Ok(())
            },
            _ => Ok(()),
        })?;
        if second_next_index != Some(next_index) || second_pair_count != pair_count {
            return Err(invalid("table merge formula-pair count changed"));
        }
        for pair in pairs {
            self.decode_pair(pair, next_index, expected_table, rows, columns, depth + 1)?;
        }
        Ok(())
    }

    fn decode_pair(
        &mut self,
        source: &[u8],
        next_index: u32,
        expected_table: [u32; 4],
        rows: u32,
        columns: u32,
        depth: usize,
    ) -> Result<(), Error> {
        let mut index = None;
        let mut formula = None;
        self.scan(source, depth, |field| match field.number() {
            FORMULA_PAIR_INDEX_FIELD => {
                if index.is_some() {
                    return Err(duplicate("FormulaStorePair.formula_index"));
                }
                index = Some(varint_u32(field, "formula_index")?);
                Ok(())
            },
            FORMULA_PAIR_FORMULA_FIELD => {
                if formula.is_some() {
                    return Err(duplicate("FormulaStorePair.formula"));
                }
                formula = Some(length_payload(field, "formula")?);
                Ok(())
            },
            _ => Ok(()),
        })?;
        let index = index.ok_or_else(|| invalid("FormulaStorePair is missing formula_index"))?;
        let formula = formula.ok_or_else(|| invalid("FormulaStorePair is missing formula"))?;
        if index >= next_index {
            return Err(invalid("iWork merge formula index reaches next index"));
        }
        for previous_index in 0..self.indexes.len() {
            self.charge_work(1)?;
            if self.indexes[previous_index] == index {
                return Err(invalid("iWork merge formula index is duplicated"));
            }
        }
        if self.regions.len() >= self.limits.max_regions {
            return Err(limit(
                LimitKind::TableCells,
                self.regions.len().saturating_add(1),
                self.limits.max_regions,
            ));
        }
        let region = self.decode_formula(formula, expected_table, depth + 1)?;
        if region.end_row() >= rows || region.end_column() >= columns {
            return Err(invalid("iWork merge region exceeds table dimensions"));
        }
        for previous_region in 0..self.regions.len() {
            self.charge_overlap_check()?;
            if self.regions[previous_region].overlaps(region) {
                return Err(invalid("iWork merge regions overlap"));
            }
        }
        self.indexes.push(index);
        self.regions.push(region);
        Ok(())
    }

    fn charge_overlap_check(&mut self) -> Result<(), Error> {
        let observed = self
            .overlap_checks
            .checked_add(1)
            .ok_or_else(|| invalid("table merge overlap-check count overflow"))?;
        if observed > self.limits.max_overlap_checks {
            return Err(limit(
                LimitKind::RewriteWork,
                observed,
                self.limits.max_overlap_checks,
            ));
        }
        self.overlap_checks = observed;
        self.charge_work(1)
    }

    fn decode_formula(
        &mut self,
        source: &[u8],
        expected_table: [u32; 4],
        depth: usize,
    ) -> Result<Region, Error> {
        let remaining_nesting = self.remaining_nesting(depth)?;
        if remaining_nesting == 0 {
            return Err(limit(
                LimitKind::Nesting,
                depth.saturating_add(1),
                self.limits.wire.max_nesting(),
            ));
        }
        let remaining_input = self
            .limits
            .wire
            .max_input_bytes()
            .checked_sub(self.report.input_bytes)
            .ok_or_else(|| {
                limit(
                    LimitKind::InputBytes,
                    self.report.input_bytes,
                    self.limits.wire.max_input_bytes(),
                )
            })?;
        let remaining_fields = self
            .limits
            .wire
            .max_fields()
            .checked_sub(self.report.fields)
            .ok_or_else(|| {
                limit(
                    LimitKind::Fields,
                    self.report.fields,
                    self.limits.wire.max_fields(),
                )
            })?;
        let remaining_work = self
            .limits
            .wire
            .max_rewrite_work()
            .checked_sub(self.report.work)
            .ok_or_else(|| {
                limit(
                    LimitKind::RewriteWork,
                    self.report.work,
                    self.limits.wire.max_rewrite_work(),
                )
            })?;
        if source.len() > remaining_input {
            return Err(limit(
                LimitKind::InputBytes,
                self.report.input_bytes.saturating_add(source.len()),
                self.limits.wire.max_input_bytes(),
            ));
        }
        if source.is_empty() {
            return Err(invalid("iWork merge formula is empty"));
        }
        // Account for the nested FormulaArchive before entering the decoder.
        // This keeps attempted-cost accounting monotonic even when Buffa
        // rejects the formula before it can return a successful report.
        self.charge_input(source.len())?;
        let max_depth = u32::try_from(remaining_nesting).unwrap_or(u32::MAX);
        let options = numbers_formula_codec::DecodeOptions::new(
            source.len(),
            remaining_fields,
            remaining_work,
            max_depth,
            WireLimits::MAX_FIELDS,
            source.len(),
        )
        .with_opaque_unknown_fields(true)
        .with_render_recursion_limit(max_depth);
        let context = numbers_formula_codec::FormulaContext::new(1, 0, 0, 1, 1);
        let mut visitor = MergeFormulaVisitor::default();
        let report = numbers_formula_codec::decode_formula_archive_for_render(
            source,
            context,
            options,
            &mut visitor,
        )
        .map_err(map_formula_error)?;
        debug_assert_eq!(report.bytes(), source.len());
        self.charge_fields(report.fields())?;
        self.charge_work(report.work())?;
        visitor.finish(expected_table)
    }
}

#[derive(Debug, Default)]
struct MergeFormulaVisitor {
    step: u8,
    range: Option<FormulaRenderColonTract>,
    function: Option<(u32, u32)>,
}

impl FormulaRenderVisitor for MergeFormulaVisitor {
    fn visit(
        &mut self,
        event: FormulaRenderEvent<'_>,
    ) -> Result<(), numbers_formula_codec::DecodeError> {
        match (self.step, event) {
            (0, FormulaRenderEvent::BeginArray { depth: 1 }) => self.step = 1,
            (1, FormulaRenderEvent::ColonTract(range)) => {
                self.range = Some(range);
                self.step = 2;
            },
            (
                2,
                FormulaRenderEvent::Function {
                    identifier,
                    argument_count,
                },
            ) => {
                self.function = Some((identifier, argument_count));
                self.step = 3;
            },
            (3, FormulaRenderEvent::EndArray) => self.step = 4,
            _ => self.step = u8::MAX,
        }
        Ok(())
    }
}

impl MergeFormulaVisitor {
    fn finish(self, expected_table: [u32; 4]) -> Result<Region, Error> {
        if self.step != 4
            || self.function != Some((NATIVE_MERGE_FUNCTION_INDEX, NATIVE_MERGE_FUNCTION_ARGUMENTS))
        {
            return Err(invalid("iWork table contains an unsupported merge formula"));
        }
        let range = self
            .range
            .ok_or_else(|| invalid("iWork table contains an unsupported merge formula"))?;
        if range.cross_table_extra.map(|extra| extra.table_id)
            != Some(expected_cfuuid(expected_table))
            || !range.sticky.begin_row_is_absolute
            || !range.sticky.begin_column_is_absolute
            || !range.sticky.end_row_is_absolute
            || !range.sticky.end_column_is_absolute
            || range.relative_column.count != 0
            || range.relative_row.count != 0
            || range.absolute_column.count != 1
            || range.absolute_row.count != 1
            || !range.preserve_rectangular
        {
            return Err(invalid("iWork table contains an unsupported merge formula"));
        }
        let begin_column = range
            .absolute_column
            .first_begin
            .and_then(|value| u32::try_from(value).ok())
            .ok_or_else(|| invalid("iWork table contains an unsupported merge formula"))?;
        let end_column = range
            .absolute_column
            .first_end
            .unwrap_or(i64::from(begin_column));
        let end_column = u32::try_from(end_column)
            .map_err(|_| invalid("iWork table contains an unsupported merge formula"))?;
        let begin_row = range
            .absolute_row
            .first_begin
            .and_then(|value| u32::try_from(value).ok())
            .ok_or_else(|| invalid("iWork table contains an unsupported merge formula"))?;
        let end_row = range.absolute_row.first_end.unwrap_or(i64::from(begin_row));
        let end_row = u32::try_from(end_row)
            .map_err(|_| invalid("iWork table contains an unsupported merge formula"))?;
        let column_count = end_column
            .checked_sub(begin_column)
            .and_then(|difference| difference.checked_add(1))
            .ok_or_else(|| invalid("iWork table contains an unsupported merge formula"))?;
        let row_count = end_row
            .checked_sub(begin_row)
            .and_then(|difference| difference.checked_add(1))
            .ok_or_else(|| invalid("iWork table contains an unsupported merge formula"))?;
        Region::new(begin_row, begin_column, row_count, column_count)
            .map_err(|_| invalid("iWork table contains an unsupported merge formula"))
    }
}

fn expected_cfuuid(words: [u32; 4]) -> FormulaRenderCfuuid {
    FormulaRenderCfuuid {
        has_uuid_bytes: false,
        word0: Some(words[0]),
        word1: Some(words[1]),
        word2: Some(words[2]),
        word3: Some(words[3]),
    }
}

fn length_payload<'source>(
    field: RawWireField<'source>,
    name: &str,
) -> Result<&'source [u8], Error> {
    if field.wire_type() != 2 || !field.key_is_canonical() || !field.length_is_canonical() {
        return Err(invalid(format!(
            "protobuf {name} field has invalid wire framing"
        )));
    }
    Ok(field.payload())
}

fn varint_u32(field: RawWireField<'_>, name: &str) -> Result<u32, Error> {
    if field.wire_type() != 0 || !field.key_is_canonical() || !field.value_is_canonical() {
        return Err(invalid(format!(
            "protobuf {name} field has invalid wire framing"
        )));
    }
    let (value, width) = litchi_iwa_common::decode_varint_from_bytes(field.payload())
        .map_err(|error| invalid(format!("protobuf {name} field has invalid value: {error}")))?;
    if width != field.payload().len() {
        return Err(invalid(format!("protobuf {name} field has invalid value")));
    }
    u32::try_from(value).map_err(|_| invalid(format!("protobuf {name} field exceeds u32")))
}

fn singular(seen: &mut [bool], index: usize, name: &str) -> Result<(), Error> {
    let Some(slot) = seen.get_mut(index) else {
        return Err(invalid(format!("protobuf {name} field number is invalid")));
    };
    if *slot {
        return Err(duplicate(name));
    }
    *slot = true;
    Ok(())
}

fn duplicate(name: &str) -> Error {
    invalid(format!("protobuf {name} field is duplicated"))
}

fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}

fn limit(kind: LimitKind, observed: usize, maximum: usize) -> Error {
    Error::LimitExceeded {
        kind,
        observed,
        limit: maximum,
    }
}

fn map_formula_error(error: numbers_formula_codec::DecodeError) -> Error {
    use numbers_formula_codec::DecodeLimit;
    match error.resource_limit() {
        Some(DecodeLimit::Bytes { observed, maximum }) => {
            limit(LimitKind::InputBytes, observed, maximum)
        },
        Some(DecodeLimit::Fields { observed, maximum }) => {
            limit(LimitKind::Fields, observed, maximum)
        },
        Some(DecodeLimit::Work { observed, maximum }) => {
            limit(LimitKind::RewriteWork, observed, maximum)
        },
        Some(DecodeLimit::Nesting { observed, maximum }) => {
            limit(LimitKind::Nesting, observed as usize, maximum as usize)
        },
        Some(DecodeLimit::Nodes { observed, maximum }) => {
            limit(LimitKind::Fields, observed, maximum)
        },
        Some(DecodeLimit::Text { observed, maximum }) => {
            limit(LimitKind::InputBytes, observed, maximum)
        },
        Some(DecodeLimit::Allocation { requested }) => Error::Allocation {
            resource: "table merge formula",
            amount: requested,
        },
        None => invalid("iWork table merge formula is malformed"),
    }
}
