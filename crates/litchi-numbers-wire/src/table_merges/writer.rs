//! Source-preserving writer for the native Numbers merge-owner path.
//!
//! The writer deliberately lives below [`super`] so that its public surface is
//! limited to one semantic operation.  It scans only the selected table-model
//! closure, copies every retained wire field verbatim, and emits canonical
//! bytes for fields that are newly created. Formula construction delegates to
//! the bounded Buffa codec in the protobuf owner; the final candidate is
//! checked by the borrowed Buffa reader in the parent module.

use super::{
    Decoder, FORMULA_PAIR_FORMULA_FIELD, FORMULA_PAIR_INDEX_FIELD, FORMULA_STORE_FORMULAS_FIELD,
    FORMULA_STORE_NEXT_INDEX_FIELD, MERGE_OWNER_FORMULA_STORE_FIELD, MergeReadError, ReadLimits,
    ReadReport, TABLE_MODEL_COLUMNS_FIELD, TABLE_MODEL_MERGE_OWNER_FIELD, TABLE_MODEL_ROWS_FIELD,
    TABLE_MODEL_TABLE_ID_FIELD, invalid, length_payload, limit, varint_u32,
};
use litchi_iwa_common::table::merge::Region;
use litchi_iwa_common::varint;
use litchi_iwa_common::wire::RawWireField;
use litchi_iwa_common::{Error, LimitKind, WireLimits};
use litchi_iwa_protos::table_merge_formula_codec::{
    self as merge_formula_codec, EncodeError, EncodeLimit, EncodeOptions,
};

// The native formula codec has a fixed schema shape. Every u32 value in that
// shape is at most five bytes, which bounds one fresh formula at 88 bytes and
// its enclosing FormulaStorePair at 96 bytes. Keeping these bounds here lets
// the writer reserve the later copy/framing work before asking the Buffa codec
// to allocate any staging buffers.
const MAX_MERGE_FORMULA_BYTES: usize = 88;
const MAX_MERGE_PAIR_BYTES: usize = 96;
const MERGE_FORMULA_FOLLOWUP_WORK: usize =
    MAX_MERGE_FORMULA_BYTES.saturating_add(MAX_MERGE_PAIR_BYTES);

/// The source-preserving result of one desired merge-set rewrite.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct MergeWrite {
    /// Rewritten table-model bytes.
    pub data: Vec<u8>,
    /// Aggregate source, rewrite, and candidate-verification accounting.
    pub report: ReadReport,
    /// Whether the returned bytes differ from the source bytes.
    pub changed: bool,
}

#[derive(Debug, Clone, Copy)]
struct FieldRef {
    payload_start: usize,
    payload_end: usize,
}

impl FieldRef {
    fn from_field(field: RawWireField<'_>) -> Self {
        Self {
            payload_start: field.payload_start(),
            payload_end: field.payload_end(),
        }
    }

    fn payload(self, source: &[u8]) -> &[u8] {
        &source[self.payload_start..self.payload_end]
    }
}

#[derive(Debug, Default)]
struct Root<'source> {
    table_id: Option<&'source [u8]>,
    rows: Option<u32>,
    columns: Option<u32>,
    merge_owner: Option<FieldRef>,
}

#[derive(Debug)]
struct Owner<'source> {
    source: &'source [u8],
    store: Option<Store<'source>>,
}

#[derive(Debug)]
struct Store<'source> {
    source: &'source [u8],
    next_index: u32,
    pair_count: usize,
}

#[derive(Debug)]
struct Source<'source> {
    root: Root<'source>,
    owner: Option<Owner<'source>>,
}

/// Rewrite the native table-model merge set.
///
/// `regions` is the desired set. Existing regions are retained in source
/// order, including their formula indexes and raw pair/formula bytes; regions
/// absent from the source are appended in caller order. An unchanged set is
/// an exact byte-preserving no-op, regardless of request order. `owner_uuid`
/// is used only when a new merge owner must be created.
pub fn rewrite_table_merges(
    source: &[u8],
    regions: &[Region],
    owner_uuid: [u8; 16],
    limits: ReadLimits,
) -> Result<MergeWrite, MergeReadError> {
    let current = super::read_table_merges(source, limits)?;
    let mut decoder = Decoder::new(limits);
    decoder.report = current.report;
    decoder.overlap_checks = match pair_count(current.regions.len()) {
        Ok(value) => value,
        Err(error) => return Err(decoder.failure(error)),
    };
    if let Err(error) = decoder.validate_limits() {
        return Err(decoder.failure(error));
    }
    if regions.len() > limits.max_regions {
        return Err(decoder.failure(limit(
            LimitKind::TableCells,
            regions.len(),
            limits.max_regions,
        )));
    }
    if let Err(error) = validate_desired_regions(&mut decoder, regions) {
        return Err(decoder.failure(error));
    }
    let same_set = match same_region_set(&mut decoder, &current.regions, regions) {
        Ok(value) => value,
        Err(error) => return Err(decoder.failure(error)),
    };
    if same_set {
        if let Err(error) = decoder.charge_work(source.len()) {
            return Err(decoder.failure(error));
        }
        let data = match clone_output(source, limits.wire) {
            Ok(data) => data,
            Err(error) => return Err(decoder.failure(error)),
        };
        return Ok(MergeWrite {
            data,
            report: decoder.report,
            changed: false,
        });
    }

    let parsed = match gather_source(source, &mut decoder) {
        Ok(parsed) => parsed,
        Err(error) => return Err(decoder.failure(error)),
    };
    let rows = match parsed.root.rows {
        Some(rows) => rows,
        None => return Err(decoder.failure(invalid("table model is missing number_of_rows"))),
    };
    let columns = match parsed.root.columns {
        Some(columns) => columns,
        None => return Err(decoder.failure(invalid("table model is missing number_of_columns"))),
    };
    if let Err(error) = validate_bounds_and_set(&mut decoder, regions, rows, columns) {
        return Err(decoder.failure(error));
    }

    let owner = parsed.owner.as_ref();
    let current_regions = &current.regions;
    let mut keep = match allocate_keep(current_regions.len()) {
        Ok(keep) => keep,
        Err(error) => return Err(decoder.failure(error)),
    };
    let mut new_regions = Vec::new();
    if new_regions.try_reserve_exact(regions.len()).is_err() {
        return Err(decoder.failure(Error::Allocation {
            resource: "table merge new regions",
            amount: regions.len(),
        }));
    }
    let mut retained_count = 0usize;
    for &desired in regions {
        let mut retained = None;
        for (index, &current) in current_regions.iter().enumerate() {
            if let Err(error) = decoder.charge_work(1) {
                return Err(decoder.failure(error));
            }
            if current == desired {
                retained = Some(index);
                break;
            }
        }
        if let Some(index) = retained {
            keep[index] = true;
            retained_count = retained_count.saturating_add(1);
        } else {
            new_regions.push(desired);
        }
    }
    if owner.is_none() && new_regions.is_empty() {
        return Err(decoder.failure(invalid(
            "table merge source has no owner for a retained region",
        )));
    }
    let table_uuid = match parsed.root.table_id {
        Some(table_id) => match super::parse_table_uuid(table_id) {
            Ok(value) => value,
            Err(error) => return Err(decoder.failure(error)),
        },
        None => return Err(decoder.failure(invalid("table model merge owner is missing table_id"))),
    };

    let next_index = owner
        .and_then(|owner| owner.store.as_ref())
        .map_or(0, |store| store.next_index);
    let appended_count = match u32::try_from(new_regions.len()) {
        Ok(value) => value,
        Err(_) => {
            return Err(decoder.failure(invalid("table merge formula index count overflows u32")));
        },
    };
    let next_after = match next_index.checked_add(appended_count) {
        Some(value) => value,
        None => {
            return Err(decoder.failure(invalid("table merge next_formula_index overflows u32")));
        },
    };
    let mut new_pairs = Vec::new();
    if new_pairs.try_reserve_exact(new_regions.len()).is_err() {
        return Err(decoder.failure(Error::Allocation {
            resource: "table merge formula pairs",
            amount: new_regions.len(),
        }));
    }
    for (offset, &region) in new_regions.iter().enumerate() {
        let offset = match u32::try_from(offset) {
            Ok(value) => value,
            Err(_) => {
                return Err(decoder.failure(invalid(
                    "table merge formula index overflows next_formula_index",
                )));
            },
        };
        let index = match next_index.checked_add(offset) {
            Some(value) => value,
            None => {
                return Err(decoder.failure(invalid(
                    "table merge formula index overflows next_formula_index",
                )));
            },
        };
        let formula =
            match encode_merge_formula(&mut decoder, region, table_uuid.owner_cfuuid_words()) {
                Ok(formula) => formula,
                Err(error) => return Err(decoder.failure(error)),
            };
        if let Err(error) = decoder.charge_work(formula.len()) {
            return Err(decoder.failure(error));
        }
        let pair = match encode_pair(
            &mut decoder,
            index,
            &formula,
            limits.wire.max_output_bytes(),
        ) {
            Ok(pair) => pair,
            Err(error) => return Err(decoder.failure(error)),
        };
        new_pairs.push(pair);
    }

    if regions.is_empty() {
        // Host behavior removes the owner only when an existing non-empty set
        // is actually reduced to empty. The equal-set fast path above keeps an
        // empty owner/store byte-for-byte intact.
        if current_regions.is_empty() {
            return Err(decoder.failure(invalid(
                "table merge empty-set transition reached an invalid source state",
            )));
        }
        let data = match rewrite_root(
            source,
            &mut decoder,
            &parsed.root,
            RootAction::Remove,
            limits.wire.max_output_bytes(),
        ) {
            Ok(data) => data,
            Err(error) => return Err(decoder.failure(error)),
        };
        return finish_candidate(data, current_regions, regions, decoder, limits);
    }

    let mut store_payload = None;
    if let Some(owner) = owner {
        let existing_store = owner.store.as_ref();
        let remove_any = existing_store.is_some() && retained_count < current_regions.len();
        let add_any = !new_pairs.is_empty();
        if let Some(store) = existing_store {
            if remove_any || add_any {
                let payload = match rewrite_store(
                    &mut decoder,
                    store,
                    &keep,
                    if add_any { Some(next_after) } else { None },
                    &new_pairs,
                    limits.wire.max_output_bytes(),
                ) {
                    Ok(payload) => payload,
                    Err(error) => return Err(decoder.failure(error)),
                };
                store_payload = Some(payload);
            }
        } else {
            let payload = match encode_store(
                &mut decoder,
                next_after,
                &new_pairs,
                limits.wire.max_output_bytes(),
            ) {
                Ok(payload) => payload,
                Err(error) => return Err(decoder.failure(error)),
            };
            store_payload = Some(payload);
        }
        let owner_payload = match store_payload.as_deref() {
            Some(store_payload) => match rewrite_owner(
                &mut decoder,
                owner,
                Some(store_payload),
                limits.wire.max_output_bytes(),
            ) {
                Ok(payload) => payload,
                Err(error) => return Err(decoder.failure(error)),
            },
            None => return Err(decoder.failure(invalid("table merge owner rewrite lost store"))),
        };
        let data = match rewrite_root(
            source,
            &mut decoder,
            &parsed.root,
            RootAction::Replace(&owner_payload),
            limits.wire.max_output_bytes(),
        ) {
            Ok(data) => data,
            Err(error) => return Err(decoder.failure(error)),
        };
        return finish_candidate(data, current_regions, regions, decoder, limits);
    }

    let owner_id = match encode_owner_id(&mut decoder, owner_uuid, limits.wire.max_output_bytes()) {
        Ok(owner_id) => owner_id,
        Err(error) => return Err(decoder.failure(error)),
    };
    let store = match encode_store(
        &mut decoder,
        next_after,
        &new_pairs,
        limits.wire.max_output_bytes(),
    ) {
        Ok(payload) => payload,
        Err(error) => return Err(decoder.failure(error)),
    };
    let new_owner = match encode_new_owner(
        &mut decoder,
        &owner_id,
        &store,
        limits.wire.max_output_bytes(),
    ) {
        Ok(payload) => payload,
        Err(error) => return Err(decoder.failure(error)),
    };
    let data = match rewrite_root(
        source,
        &mut decoder,
        &parsed.root,
        RootAction::Append(&new_owner),
        limits.wire.max_output_bytes(),
    ) {
        Ok(data) => data,
        Err(error) => return Err(decoder.failure(error)),
    };
    finish_candidate(data, current_regions, regions, decoder, limits)
}

fn validate_desired_regions(decoder: &mut Decoder, regions: &[Region]) -> Result<(), Error> {
    for (index, &region) in regions.iter().enumerate() {
        for &previous in &regions[..index] {
            decoder.charge_overlap_check()?;
            if previous == region {
                return Err(invalid("desired table merge regions contain a duplicate"));
            }
            if previous.overlaps(region) {
                return Err(invalid("desired table merge regions overlap"));
            }
        }
    }
    Ok(())
}

fn validate_bounds_and_set(
    decoder: &mut Decoder,
    regions: &[Region],
    rows: u32,
    columns: u32,
) -> Result<(), Error> {
    for &region in regions {
        if region.end_row() >= rows || region.end_column() >= columns {
            return Err(invalid(
                "desired table merge region exceeds table dimensions",
            ));
        }
        decoder.charge_work(1)?;
    }
    Ok(())
}

fn same_region_set(
    decoder: &mut Decoder,
    current: &[Region],
    desired: &[Region],
) -> Result<bool, Error> {
    if current.len() != desired.len() {
        return Ok(false);
    }
    for &wanted in desired {
        let mut found = false;
        for &candidate in current {
            decoder.charge_work(1)?;
            if candidate == wanted {
                found = true;
                break;
            }
        }
        if !found {
            return Ok(false);
        }
    }
    Ok(true)
}

fn pair_count(length: usize) -> Result<usize, Error> {
    length
        .checked_mul(length.saturating_sub(1))
        .and_then(|value| value.checked_div(2))
        .ok_or_else(|| invalid("table merge overlap-check count overflows usize"))
}

fn allocate_keep(length: usize) -> Result<Vec<bool>, Error> {
    let mut keep = Vec::new();
    keep.try_reserve_exact(length)
        .map_err(|_| Error::Allocation {
            resource: "table merge retention flags",
            amount: length,
        })?;
    keep.resize(length, false);
    Ok(keep)
}

fn gather_source<'source>(
    source: &'source [u8],
    decoder: &mut Decoder,
) -> Result<Source<'source>, Error> {
    let mut model = Root::default();
    decoder.scan(source, 0, |field| {
        match field.number() {
            TABLE_MODEL_TABLE_ID_FIELD => {
                if model.table_id.is_some() {
                    return Err(super::duplicate("TableModelArchive.table_id"));
                }
                let value = length_payload(field, "table_id")?;
                core::str::from_utf8(value)
                    .map_err(|_| invalid("table model table_id is not UTF-8"))?;
                model.table_id = Some(value);
            },
            TABLE_MODEL_ROWS_FIELD => {
                if model.rows.is_some() {
                    return Err(super::duplicate("TableModelArchive.number_of_rows"));
                }
                model.rows = Some(varint_u32(field, "number_of_rows")?);
            },
            TABLE_MODEL_COLUMNS_FIELD => {
                if model.columns.is_some() {
                    return Err(super::duplicate("TableModelArchive.number_of_columns"));
                }
                model.columns = Some(varint_u32(field, "number_of_columns")?);
            },
            TABLE_MODEL_MERGE_OWNER_FIELD => {
                if model.merge_owner.is_some() {
                    return Err(super::duplicate("TableModelArchive.merge_owner"));
                }
                length_payload(field, "merge_owner")?;
                model.merge_owner = Some(FieldRef::from_field(field));
            },
            _ => {},
        }
        Ok(())
    })?;
    let owner = match model.merge_owner {
        Some(owner_field) => Some(gather_owner(source, decoder, owner_field)?),
        None => None,
    };
    Ok(Source { root: model, owner })
}

fn gather_owner<'source>(
    source: &'source [u8],
    decoder: &mut Decoder,
    owner_field: FieldRef,
) -> Result<Owner<'source>, Error> {
    let owner_source = owner_field.payload(source);
    let mut owner_id = None;
    let mut store = None;
    decoder.scan(owner_source, 1, |field| {
        match field.number() {
            1 => {
                if owner_id.is_some() {
                    return Err(super::duplicate("MergeOwnerArchive.formula_owner_uid"));
                }
                owner_id = Some(length_payload(field, "formula_owner_uid")?);
            },
            MERGE_OWNER_FORMULA_STORE_FIELD => {
                if store.is_some() {
                    return Err(super::duplicate("MergeOwnerArchive.formula_store"));
                }
                length_payload(field, "formula_store")?;
                store = Some(FieldRef::from_field(field));
            },
            _ => {},
        }
        Ok(())
    })?;
    let owner_id =
        owner_id.ok_or_else(|| invalid("MergeOwnerArchive is missing required owner_id"))?;
    decoder.scan_cfuuid(owner_id, 2)?;
    let store = match store {
        Some(store_field) => Some(gather_store(owner_source, decoder, store_field, 2)?),
        None => None,
    };
    Ok(Owner {
        source: owner_source,
        store,
    })
}

fn gather_store<'source>(
    source: &'source [u8],
    decoder: &mut Decoder,
    store_field: FieldRef,
    depth: usize,
) -> Result<Store<'source>, Error> {
    let store_source = store_field.payload(source);
    let mut next_index = None;
    let mut pair_fields = Vec::new();
    let max_regions = decoder.limits.max_regions;
    decoder.scan(store_source, depth, |field| {
        match field.number() {
            FORMULA_STORE_NEXT_INDEX_FIELD => {
                if next_index.is_some() {
                    return Err(super::duplicate("FormulaStoreArchive.next_formula_index"));
                }
                next_index = Some(varint_u32(field, "next_formula_index")?);
            },
            FORMULA_STORE_FORMULAS_FIELD => {
                length_payload(field, "formula pair")?;
                if pair_fields.len() >= max_regions {
                    return Err(limit(
                        LimitKind::TableCells,
                        pair_fields.len().saturating_add(1),
                        max_regions,
                    ));
                }
                pair_fields.try_reserve(1).map_err(|_| Error::Allocation {
                    resource: "table merge pair fields",
                    amount: pair_fields.len().saturating_add(1),
                })?;
                pair_fields.push(FieldRef::from_field(field));
            },
            _ => {},
        }
        Ok(())
    })?;
    let next_index =
        next_index.ok_or_else(|| invalid("FormulaStoreArchive is missing next_formula_index"))?;
    let mut pair_count = 0usize;
    for pair_field in pair_fields {
        let pair_source = pair_field.payload(store_source);
        let mut index = None;
        let mut formula = None;
        decoder.scan(pair_source, depth + 1, |field| {
            match field.number() {
                FORMULA_PAIR_INDEX_FIELD => {
                    if index.is_some() {
                        return Err(super::duplicate("FormulaStorePair.formula_index"));
                    }
                    index = Some(varint_u32(field, "formula_index")?);
                },
                FORMULA_PAIR_FORMULA_FIELD => {
                    if formula.is_some() {
                        return Err(super::duplicate("FormulaStorePair.formula"));
                    }
                    length_payload(field, "formula")?;
                    formula = Some(FieldRef::from_field(field));
                },
                _ => {},
            }
            Ok(())
        })?;
        index.ok_or_else(|| invalid("FormulaStorePair is missing formula_index"))?;
        formula.ok_or_else(|| invalid("FormulaStorePair is missing formula"))?;
        pair_count = pair_count
            .checked_add(1)
            .ok_or_else(|| invalid("table merge pair count overflows usize"))?;
    }
    Ok(Store {
        source: store_source,
        next_index,
        pair_count,
    })
}

#[derive(Debug, Clone, Copy)]
enum RootAction<'payload> {
    Remove,
    Replace(&'payload [u8]),
    Append(&'payload [u8]),
}

fn rewrite_root<'source>(
    source: &'source [u8],
    decoder: &mut Decoder,
    _root: &Root<'source>,
    action: RootAction<'_>,
    max_output: usize,
) -> Result<Vec<u8>, Error> {
    let mut measure = Sink::measure(max_output);
    emit_root_pass(source, decoder, action, &mut measure)?;
    let length = measure.len;
    decoder.charge_work(length)?;
    let mut output = Sink::materialize(length, max_output)?;
    emit_root_pass(source, decoder, action, &mut output)?;
    output.finish(length)
}

fn emit_root_pass(
    source: &[u8],
    decoder: &mut Decoder,
    action: RootAction<'_>,
    sink: &mut Sink,
) -> Result<(), Error> {
    let mut seen_owner = false;
    decoder.scan(source, 0, |field| {
        if field.number() != TABLE_MODEL_MERGE_OWNER_FIELD {
            return sink.raw(field.raw());
        }
        seen_owner = true;
        match action {
            RootAction::Remove => Ok(()),
            RootAction::Replace(payload) => sink.with_key(field.key(), payload),
            RootAction::Append(_) => sink.raw(field.raw()),
        }
    })?;
    if !seen_owner {
        if let RootAction::Append(payload) = action {
            sink.bytes_field(TABLE_MODEL_MERGE_OWNER_FIELD, payload)?;
        }
    }
    Ok(())
}

fn rewrite_owner<'source>(
    decoder: &mut Decoder,
    owner: &Owner<'source>,
    replacement_store: Option<&[u8]>,
    max_output: usize,
) -> Result<Vec<u8>, Error> {
    let owner_source = owner.source;
    let mut measure = Sink::measure(max_output);
    emit_owner_pass(owner_source, decoder, replacement_store, &mut measure)?;
    let length = measure.len;
    decoder.charge_work(length)?;
    let mut output = Sink::materialize(length, max_output)?;
    emit_owner_pass(owner_source, decoder, replacement_store, &mut output)?;
    output.finish(length)
}

fn emit_owner_pass(
    source: &[u8],
    decoder: &mut Decoder,
    replacement_store: Option<&[u8]>,
    sink: &mut Sink,
) -> Result<(), Error> {
    let mut seen_store = false;
    decoder.scan(source, 1, |field| {
        if field.number() != MERGE_OWNER_FORMULA_STORE_FIELD {
            return sink.raw(field.raw());
        }
        seen_store = true;
        match replacement_store {
            Some(payload) => sink.with_key(field.key(), payload),
            None => sink.raw(field.raw()),
        }
    })?;
    if !seen_store {
        if let Some(payload) = replacement_store {
            sink.bytes_field(MERGE_OWNER_FORMULA_STORE_FIELD, payload)?;
        }
    }
    Ok(())
}

fn rewrite_store<'source>(
    decoder: &mut Decoder,
    store: &Store<'source>,
    keep: &[bool],
    replacement_next: Option<u32>,
    appended_pairs: &[Vec<u8>],
    max_output: usize,
) -> Result<Vec<u8>, Error> {
    let store_source = store.source;
    let mut measure = Sink::measure(max_output);
    emit_store_pass(
        store_source,
        decoder,
        store,
        keep,
        replacement_next,
        appended_pairs,
        &mut measure,
    )?;
    let length = measure.len;
    decoder.charge_work(length)?;
    let mut output = Sink::materialize(length, max_output)?;
    emit_store_pass(
        store_source,
        decoder,
        store,
        keep,
        replacement_next,
        appended_pairs,
        &mut output,
    )?;
    output.finish(length)
}

fn emit_store_pass(
    source: &[u8],
    decoder: &mut Decoder,
    store: &Store<'_>,
    keep: &[bool],
    replacement_next: Option<u32>,
    appended_pairs: &[Vec<u8>],
    sink: &mut Sink,
) -> Result<(), Error> {
    let mut pair_position = 0usize;
    let mut seen_next = false;
    decoder.scan(source, 2, |field| match field.number() {
        FORMULA_STORE_NEXT_INDEX_FIELD => {
            if seen_next {
                return Err(super::duplicate("FormulaStoreArchive.next_formula_index"));
            }
            seen_next = true;
            match replacement_next {
                Some(value) if value != store.next_index => {
                    sink.with_key_varint(field.key(), value)
                },
                _ => sink.raw(field.raw()),
            }
        },
        FORMULA_STORE_FORMULAS_FIELD => {
            let retain = keep
                .get(pair_position)
                .copied()
                .ok_or_else(|| invalid("table merge formula-pair count changed during rewrite"))?;
            pair_position = pair_position
                .checked_add(1)
                .ok_or_else(|| invalid("table merge formula-pair position overflow"))?;
            if retain {
                sink.raw(field.raw())
            } else {
                Ok(())
            }
        },
        _ => sink.raw(field.raw()),
    })?;
    if !seen_next || pair_position != store.pair_count {
        return Err(invalid(
            "table merge formula-pair count changed during rewrite",
        ));
    }
    for pair in appended_pairs {
        sink.bytes_field(FORMULA_STORE_FORMULAS_FIELD, pair)?;
    }
    Ok(())
}

fn encode_store(
    decoder: &mut Decoder,
    next_index: u32,
    pairs: &[Vec<u8>],
    max_output: usize,
) -> Result<Vec<u8>, Error> {
    let mut measure = Sink::measure(max_output);
    measure.varint_field(FORMULA_STORE_NEXT_INDEX_FIELD, u64::from(next_index))?;
    for pair in pairs {
        measure.bytes_field(FORMULA_STORE_FORMULAS_FIELD, pair)?;
    }
    let length = measure.len;
    decoder.charge_work(length)?;
    let mut output = Sink::materialize(length, max_output)?;
    output.varint_field(FORMULA_STORE_NEXT_INDEX_FIELD, u64::from(next_index))?;
    for pair in pairs {
        output.bytes_field(FORMULA_STORE_FORMULAS_FIELD, pair)?;
    }
    output.finish(length)
}

fn encode_new_owner(
    decoder: &mut Decoder,
    owner_id: &[u8],
    store: &[u8],
    max_output: usize,
) -> Result<Vec<u8>, Error> {
    let mut measure = Sink::measure(max_output);
    measure.bytes_field(1, owner_id)?;
    measure.bytes_field(MERGE_OWNER_FORMULA_STORE_FIELD, store)?;
    let length = measure.len;
    decoder.charge_work(length)?;
    let mut output = Sink::materialize(length, max_output)?;
    output.bytes_field(1, owner_id)?;
    output.bytes_field(MERGE_OWNER_FORMULA_STORE_FIELD, store)?;
    output.finish(length)
}

fn encode_owner_id(
    decoder: &mut Decoder,
    owner_uuid: [u8; 16],
    max_output: usize,
) -> Result<Vec<u8>, Error> {
    let lower = u64::from_le_bytes(owner_uuid[..8].try_into().expect("owner UUID has 16 bytes"));
    let upper = u64::from_le_bytes(owner_uuid[8..].try_into().expect("owner UUID has 16 bytes"));
    let words = [
        lower as u32,
        (lower >> 32) as u32,
        upper as u32,
        (upper >> 32) as u32,
    ];
    let mut sink = Sink::measure(max_output);
    for (index, word) in words.into_iter().enumerate() {
        sink.varint_field((index as u32) + 2, u64::from(word))?;
    }
    let length = sink.len;
    decoder.charge_work(length)?;
    let mut output = Sink::materialize(length, max_output)?;
    for (index, word) in words.into_iter().enumerate() {
        output.varint_field((index as u32) + 2, u64::from(word))?;
    }
    output.finish(length)
}

fn encode_pair(
    decoder: &mut Decoder,
    index: u32,
    formula: &[u8],
    max_output: usize,
) -> Result<Vec<u8>, Error> {
    let mut sink = Sink::measure(max_output);
    sink.varint_field(FORMULA_PAIR_INDEX_FIELD, u64::from(index))?;
    sink.bytes_field(FORMULA_PAIR_FORMULA_FIELD, formula)?;
    let length = sink.len;
    decoder.charge_work(length)?;
    let mut output = Sink::materialize(length, max_output)?;
    output.varint_field(FORMULA_PAIR_INDEX_FIELD, u64::from(index))?;
    output.bytes_field(FORMULA_PAIR_FORMULA_FIELD, formula)?;
    output.finish(length)
}

fn encode_merge_formula(
    decoder: &mut Decoder,
    region: Region,
    table_words: [u32; 4],
) -> Result<Vec<u8>, Error> {
    ensure_work(decoder, MERGE_FORMULA_FOLLOWUP_WORK)?;
    let wire = decoder.limits.wire;
    let remaining_work = wire
        .max_rewrite_work()
        .saturating_sub(decoder.report.work)
        .checked_sub(MERGE_FORMULA_FOLLOWUP_WORK)
        .ok_or_else(|| {
            limit(
                LimitKind::RewriteWork,
                wire.max_rewrite_work(),
                wire.max_rewrite_work(),
            )
        })?;
    let options = EncodeOptions::new(
        wire.max_output_bytes(),
        wire.max_fields().saturating_sub(decoder.report.fields),
        remaining_work,
    );
    let (formula, report) = merge_formula_codec::encode_merge_formula(
        table_words,
        region.row(),
        region.end_row(),
        region.column(),
        region.end_column(),
        options,
    )
    .map_err(map_formula_encode_error)?;
    if report.output_bytes() != formula.len() {
        return Err(invalid(
            "table merge formula codec returned an inconsistent output length",
        ));
    }
    decoder.charge_fields(report.fields())?;
    decoder.charge_work(report.work_bytes())?;
    Ok(formula)
}

fn ensure_work(decoder: &Decoder, amount: usize) -> Result<(), Error> {
    let observed = decoder
        .report
        .work
        .checked_add(amount)
        .ok_or_else(|| invalid("table merge work count overflows usize"))?;
    if observed > decoder.limits.wire.max_rewrite_work() {
        return Err(limit(
            LimitKind::RewriteWork,
            observed,
            decoder.limits.wire.max_rewrite_work(),
        ));
    }
    Ok(())
}

fn map_formula_encode_error(error: EncodeError) -> Error {
    match error {
        EncodeError::InvalidRange => invalid("table merge formula range is invalid"),
        EncodeError::Limit(limit_kind) => match limit_kind {
            EncodeLimit::OutputBytes { observed, maximum } => {
                limit(LimitKind::OutputBytes, observed, maximum)
            },
            EncodeLimit::Fields { observed, maximum } => {
                limit(LimitKind::Fields, observed, maximum)
            },
            EncodeLimit::WorkBytes { observed, maximum } => {
                limit(LimitKind::RewriteWork, observed, maximum)
            },
            _ => invalid("table merge formula codec exceeded an unsupported limit"),
        },
        EncodeError::Allocation { requested } => Error::Allocation {
            resource: "table merge formula",
            amount: requested,
        },
        other => invalid(format!("table merge formula codec failed: {other}")),
    }
}

fn clone_output(source: &[u8], limits: WireLimits) -> Result<Vec<u8>, Error> {
    if source.len() > limits.max_output_bytes() {
        return Err(Error::LimitExceeded {
            kind: LimitKind::OutputBytes,
            observed: source.len(),
            limit: limits.max_output_bytes(),
        });
    }
    let mut output = Vec::new();
    output
        .try_reserve_exact(source.len())
        .map_err(|_| Error::Allocation {
            resource: "table merge output",
            amount: source.len(),
        })?;
    output.extend_from_slice(source);
    Ok(output)
}

fn finish_candidate(
    data: Vec<u8>,
    current: &[Region],
    desired: &[Region],
    mut decoder: Decoder,
    limits: ReadLimits,
) -> Result<MergeWrite, MergeReadError> {
    if data.len() > limits.wire.max_output_bytes() {
        return Err(decoder.failure(Error::LimitExceeded {
            kind: LimitKind::OutputBytes,
            observed: data.len(),
            limit: limits.wire.max_output_bytes(),
        }));
    }
    let candidate_limits = match residual_read_limits(limits, &decoder) {
        Ok(value) => value,
        Err(error) => return Err(decoder.failure(error)),
    };
    let candidate = match super::read_table_merges(&data, candidate_limits) {
        Ok(value) => value,
        Err(error) => {
            let attempted = error.attempted();
            add_attempted(&mut decoder, attempted);
            return Err(decoder.failure(error.error.clone()));
        },
    };
    if let Err(error) = add_report(&mut decoder, candidate.report) {
        return Err(decoder.failure(error));
    }
    let mut expected = Vec::new();
    if expected.try_reserve_exact(desired.len()).is_err() {
        return Err(decoder.failure(Error::Allocation {
            resource: "table merge candidate regions",
            amount: desired.len(),
        }));
    }
    for &region in current {
        let mut retained = false;
        for &candidate in desired {
            if let Err(error) = decoder.charge_work(1) {
                return Err(decoder.failure(error));
            }
            if candidate == region {
                retained = true;
                break;
            }
        }
        if retained {
            expected.push(region);
        }
    }
    for &region in desired {
        let mut retained = false;
        for &candidate in current {
            if let Err(error) = decoder.charge_work(1) {
                return Err(decoder.failure(error));
            }
            if candidate == region {
                retained = true;
                break;
            }
        }
        if !retained {
            expected.push(region);
        }
    }
    if candidate.regions.len() != expected.len() {
        return Err(decoder.failure(invalid(
            "table merge rewrite candidate changed region topology",
        )));
    }
    for (actual, expected) in candidate.regions.iter().zip(expected.iter()) {
        if let Err(error) = decoder.charge_work(1) {
            return Err(decoder.failure(error));
        }
        if actual != expected {
            return Err(decoder.failure(invalid(
                "table merge rewrite candidate changed region topology",
            )));
        }
    }
    Ok(MergeWrite {
        changed: true,
        data,
        report: decoder.report,
    })
}

fn residual_read_limits(limits: ReadLimits, decoder: &Decoder) -> Result<ReadLimits, Error> {
    let wire = WireLimits::default()
        .with_input_bytes(
            limits
                .wire
                .max_input_bytes()
                .saturating_sub(decoder.report.input_bytes)
                .max(1),
        )?
        .with_fields(
            limits
                .wire
                .max_fields()
                .saturating_sub(decoder.report.fields)
                .max(1),
        )?
        .with_output_bytes(limits.wire.max_output_bytes())?
        .with_nesting(limits.wire.max_nesting())?
        .with_rewrite_work(
            limits
                .wire
                .max_rewrite_work()
                .saturating_sub(decoder.report.work)
                .max(1),
        )?;
    Ok(ReadLimits {
        wire,
        max_regions: limits.max_regions,
        max_overlap_checks: limits
            .max_overlap_checks
            .saturating_sub(decoder.overlap_checks),
    })
}

fn add_report(decoder: &mut Decoder, report: ReadReport) -> Result<(), Error> {
    add_counter(
        &mut decoder.report.input_bytes,
        report.input_bytes(),
        decoder.limits.wire.max_input_bytes(),
        LimitKind::InputBytes,
    )?;
    add_counter(
        &mut decoder.report.fields,
        report.fields(),
        decoder.limits.wire.max_fields(),
        LimitKind::Fields,
    )?;
    add_counter(
        &mut decoder.report.work,
        report.work(),
        decoder.limits.wire.max_rewrite_work(),
        LimitKind::RewriteWork,
    )?;
    Ok(())
}

fn add_attempted(decoder: &mut Decoder, attempted: super::AttemptedCost) {
    decoder.report.input_bytes = decoder
        .report
        .input_bytes
        .saturating_add(attempted.input_bytes);
    decoder.report.fields = decoder.report.fields.saturating_add(attempted.fields);
    decoder.report.work = decoder.report.work.saturating_add(attempted.work);
}

fn add_counter(
    current: &mut usize,
    amount: usize,
    maximum: usize,
    kind: LimitKind,
) -> Result<(), Error> {
    let observed = current.saturating_add(amount);
    if observed > maximum {
        *current = observed;
        return Err(Error::LimitExceeded {
            kind,
            observed,
            limit: maximum,
        });
    }
    *current = observed;
    Ok(())
}

struct Sink {
    output: Option<Vec<u8>>,
    len: usize,
    max: usize,
}

impl Sink {
    const fn measure(max: usize) -> Self {
        Self {
            output: None,
            len: 0,
            max,
        }
    }

    fn materialize(length: usize, max: usize) -> Result<Self, Error> {
        if length > max {
            return Err(Error::LimitExceeded {
                kind: LimitKind::OutputBytes,
                observed: length,
                limit: max,
            });
        }
        let mut output = Vec::new();
        output
            .try_reserve_exact(length)
            .map_err(|_| Error::Allocation {
                resource: "table merge output",
                amount: length,
            })?;
        Ok(Self {
            output: Some(output),
            len: 0,
            max,
        })
    }

    fn finish(mut self, expected: usize) -> Result<Vec<u8>, Error> {
        if self.len != expected {
            return Err(invalid("table merge output length changed during emission"));
        }
        let output = self
            .output
            .take()
            .ok_or_else(|| invalid("table merge output was not materialized"))?;
        if output.len() != expected {
            return Err(invalid("table merge output allocation length changed"));
        }
        Ok(output)
    }

    fn reserve(&mut self, amount: usize) -> Result<(), Error> {
        let length = self
            .len
            .checked_add(amount)
            .ok_or_else(|| invalid("table merge output length overflows usize"))?;
        if length > self.max {
            return Err(Error::LimitExceeded {
                kind: LimitKind::OutputBytes,
                observed: length,
                limit: self.max,
            });
        }
        self.len = length;
        Ok(())
    }

    fn raw(&mut self, bytes: &[u8]) -> Result<(), Error> {
        self.reserve(bytes.len())?;
        if let Some(output) = self.output.as_mut() {
            output.extend_from_slice(bytes);
        }
        Ok(())
    }

    fn key(&mut self, field: u32, wire_type: u8) -> Result<(), Error> {
        let key = (u64::from(field) << 3) | u64::from(wire_type);
        self.varint(key)
    }

    fn varint(&mut self, value: u64) -> Result<(), Error> {
        let length = varint::encoded_len(value);
        self.reserve(length)?;
        if let Some(output) = self.output.as_mut() {
            varint::encode_varint_into(output, value);
        }
        Ok(())
    }

    fn bytes_field(&mut self, field: u32, payload: &[u8]) -> Result<(), Error> {
        self.key(field, 2)?;
        self.varint(
            u64::try_from(payload.len())
                .map_err(|_| invalid("table merge payload length exceeds u64"))?,
        )?;
        self.raw(payload)
    }

    fn varint_field(&mut self, field: u32, value: u64) -> Result<(), Error> {
        self.key(field, 0)?;
        self.varint(value)
    }

    fn with_key(&mut self, key: &[u8], payload: &[u8]) -> Result<(), Error> {
        self.raw(key)?;
        self.varint(
            u64::try_from(payload.len())
                .map_err(|_| invalid("table merge payload length exceeds u64"))?,
        )?;
        self.raw(payload)
    }

    fn with_key_varint(&mut self, key: &[u8], value: u32) -> Result<(), Error> {
        self.raw(key)?;
        self.varint(u64::from(value))
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::table_merges::read_table_merges;

    const TABLE_ID: &[u8] = b"00112233445566778899aabbccddeeff";
    const TABLE_WORDS: [u32; 4] = [0x3322_1100, 0x7766_5544, 0xbbaa_9988, 0xffee_ddcc];

    fn varint(mut value: u64) -> Vec<u8> {
        let mut output = Vec::new();
        loop {
            let mut byte = (value & 0x7f) as u8;
            value >>= 7;
            if value != 0 {
                byte |= 0x80;
            }
            output.push(byte);
            if value == 0 {
                return output;
            }
        }
    }

    fn bytes_field(field: u32, payload: &[u8]) -> Vec<u8> {
        let mut output = varint((u64::from(field) << 3) | 2);
        output.extend(varint(payload.len() as u64));
        output.extend_from_slice(payload);
        output
    }

    fn varint_field(field: u32, value: u64) -> Vec<u8> {
        let mut output = varint(u64::from(field) << 3);
        output.extend(varint(value));
        output
    }

    fn cfuuid(words: [u32; 4]) -> Vec<u8> {
        let mut output = Vec::new();
        for (index, word) in words.into_iter().enumerate() {
            output.extend(varint_field((index as u32) + 2, u64::from(word)));
        }
        output
    }

    fn encode_test_merge_formula(region: Region, words: [u32; 4]) -> Vec<u8> {
        merge_formula_codec::encode_merge_formula(
            words,
            region.row(),
            region.end_row(),
            region.column(),
            region.end_column(),
            EncodeOptions::default(),
        )
        .expect("formula")
        .0
    }

    fn model_without_owner() -> Vec<u8> {
        let mut output = bytes_field(1, TABLE_ID);
        output.extend(varint_field(6, 16));
        output.extend(varint_field(7, 16));
        output.extend(bytes_field(100, &[0xde, 0xad, 0xbe, 0xef]));
        output
    }

    fn model_without_owner_with_noncanonical_unknown() -> Vec<u8> {
        let mut output = bytes_field(1, TABLE_ID);
        output.extend(varint_field(6, 16));
        output.extend(varint_field(7, 16));
        // Field 100 and its length deliberately use overlong varints. The
        // selected reader treats this unknown field as opaque and the writer
        // must copy the exact framing when it appends a new owner.
        output.extend([0xa2, 0x86, 0x00, 0x84, 0x00, 0xde, 0xad, 0xbe, 0xef]);
        output
    }

    fn model_with_empty_store(next_index: u32) -> Vec<u8> {
        let mut store = varint_field(2, u64::from(next_index));
        store.extend(bytes_field(100, &[0x66]));
        let mut owner = bytes_field(1, &cfuuid(TABLE_WORDS));
        owner.extend(bytes_field(2, &store));
        let mut model = model_without_owner();
        model.extend(bytes_field(TABLE_MODEL_MERGE_OWNER_FIELD, &owner));
        model
    }

    fn model_with_owner(region: Region) -> Vec<u8> {
        let formula = encode_test_merge_formula(region, TABLE_WORDS);
        let pair = {
            let mut value = varint_field(1, 0);
            value.extend(bytes_field(2, &formula));
            value.extend(bytes_field(100, &[0xa1, 0xb2, 0xc3]));
            value
        };
        let store = {
            let mut value = varint_field(2, 1);
            value.extend(bytes_field(3, &pair));
            value.extend(bytes_field(100, &[0x91, 0x92]));
            value
        };
        let owner = {
            let mut value = bytes_field(1, &cfuuid(TABLE_WORDS));
            value.extend(bytes_field(2, &store));
            value.extend(bytes_field(100, &[0x81]));
            value
        };
        let mut model = model_without_owner();
        model.extend(bytes_field(TABLE_MODEL_MERGE_OWNER_FIELD, &owner));
        model
    }

    #[test]
    fn creates_owner_and_round_trips_formula() {
        let region = Region::new(1, 2, 2, 3).expect("region");
        let result = rewrite_table_merges(
            &model_without_owner(),
            &[region],
            [1; 16],
            ReadLimits::default(),
        )
        .expect("write");
        assert!(result.changed);
        assert_eq!(
            read_table_merges(&result.data, ReadLimits::default())
                .unwrap()
                .regions,
            [region]
        );
    }

    #[test]
    fn fixed_formula_bounds_cover_u32_domain() {
        let region = Region::new(1, 1, u32::MAX, u32::MAX).expect("region");
        let formula = encode_test_merge_formula(region, [u32::MAX; 4]);
        assert!(formula.len() <= MAX_MERGE_FORMULA_BYTES);

        let mut decoder = Decoder::new(ReadLimits::default());
        let pair = encode_pair(
            &mut decoder,
            u32::MAX,
            &formula,
            WireLimits::MAX_OUTPUT_BYTES,
        )
        .expect("pair");
        assert!(pair.len() <= MAX_MERGE_PAIR_BYTES);
    }

    #[test]
    fn retains_unknowns_and_raw_pair_when_appending() {
        let first = Region::new(1, 2, 2, 3).expect("region");
        let second = Region::new(8, 0, 1, 2).expect("region");
        let source = model_with_owner(first);
        let result =
            rewrite_table_merges(&source, &[first, second], [2; 16], ReadLimits::default())
                .expect("write");
        assert!(result.changed);
        assert!(
            result
                .data
                .windows(4)
                .any(|window| window == [0xde, 0xad, 0xbe, 0xef])
        );
        assert!(
            result
                .data
                .windows(3)
                .any(|window| window == [0xa1, 0xb2, 0xc3])
        );
        let formula = encode_test_merge_formula(first, TABLE_WORDS);
        assert!(
            result
                .data
                .windows(formula.len())
                .any(|window| window == formula.as_slice())
        );
        assert_eq!(
            read_table_merges(&result.data, ReadLimits::default())
                .unwrap()
                .regions,
            [first, second]
        );
    }

    #[test]
    fn unchanged_set_is_exact_noop_and_last_remove_drops_owner() {
        let first = Region::new(1, 2, 2, 3).expect("region");
        let source = model_with_owner(first);
        let reordered =
            rewrite_table_merges(&source, &[first], [3; 16], ReadLimits::default()).expect("no-op");
        assert!(!reordered.changed);
        assert_eq!(reordered.data, source);
        let removed =
            rewrite_table_merges(&source, &[], [4; 16], ReadLimits::default()).expect("remove");
        assert!(removed.changed);
        assert!(
            read_table_merges(&removed.data, ReadLimits::default())
                .unwrap()
                .regions
                .is_empty()
        );
        assert!(!removed.data.windows(2).any(|window| window == [0xf8, 0x02]));
    }

    #[test]
    fn empty_owner_is_preserved_by_empty_set_noop() {
        let owner = {
            let mut value = bytes_field(1, &cfuuid(TABLE_WORDS));
            value.extend(bytes_field(100, &[0x77]));
            value
        };
        let mut source = model_without_owner();
        source.extend(bytes_field(TABLE_MODEL_MERGE_OWNER_FIELD, &owner));
        let result =
            rewrite_table_merges(&source, &[], [5; 16], ReadLimits::default()).expect("no-op");
        assert!(!result.changed);
        assert_eq!(result.data, source);
    }

    #[test]
    fn copies_noncanonical_unknown_framing_exactly() {
        let region = Region::new(4, 4, 2, 2).expect("region");
        let source = model_without_owner_with_noncanonical_unknown();
        let result = rewrite_table_merges(&source, &[region], [6; 16], ReadLimits::default())
            .expect("write");
        assert!(
            result
                .data
                .windows(9)
                .any(|window| window == [0xa2, 0x86, 0x00, 0x84, 0x00, 0xde, 0xad, 0xbe, 0xef])
        );
    }

    #[test]
    fn rejects_duplicate_desired_regions_before_mutation() {
        let region = Region::new(2, 2, 2, 2).expect("region");
        let failure = rewrite_table_merges(
            &model_without_owner(),
            &[region, region],
            [7; 16],
            ReadLimits::default(),
        )
        .expect_err("duplicate desired regions must fail");
        assert!(matches!(failure.error(), Error::InvalidFormat(_)));
    }

    #[test]
    fn rejects_next_index_overflow_without_output() {
        let region = Region::new(2, 2, 2, 2).expect("region");
        let failure = rewrite_table_merges(
            &model_with_empty_store(u32::MAX),
            &[region],
            [8; 16],
            ReadLimits::default(),
        )
        .expect_err("next index overflow must fail");
        assert!(matches!(failure.error(), Error::InvalidFormat(_)));
    }

    #[test]
    fn output_limit_is_checked_before_candidate_is_materialized() {
        let region = Region::new(2, 2, 2, 2).expect("region");
        let source = model_without_owner();
        let wire = WireLimits::default()
            .with_output_bytes(source.len())
            .expect("output limit");
        let failure = rewrite_table_merges(
            &source,
            &[region],
            [9; 16],
            ReadLimits {
                wire,
                ..ReadLimits::default()
            },
        )
        .expect_err("rewritten owner exceeds output limit");
        assert!(matches!(
            failure.error(),
            Error::LimitExceeded {
                kind: LimitKind::OutputBytes,
                ..
            }
        ));
    }

    #[test]
    fn no_op_work_budget_is_exact_and_ledgered() {
        let region = Region::new(2, 2, 2, 2).expect("region");
        let source = model_with_owner(region);
        let baseline = rewrite_table_merges(&source, &[region], [10; 16], ReadLimits::default())
            .expect("baseline no-op");
        let exact_work = baseline.report.work();
        let exact_wire = WireLimits::default()
            .with_rewrite_work(exact_work)
            .expect("baseline work is nonzero");
        let exact = rewrite_table_merges(
            &source,
            &[region],
            [11; 16],
            ReadLimits {
                wire: exact_wire,
                ..ReadLimits::default()
            },
        )
        .expect("exact no-op budget");
        assert_eq!(exact.report.work(), exact_work);
        let below_wire = WireLimits::default()
            .with_rewrite_work(exact_work - 1)
            .expect("one below baseline work is valid");
        let failure = rewrite_table_merges(
            &source,
            &[region],
            [12; 16],
            ReadLimits {
                wire: below_wire,
                ..ReadLimits::default()
            },
        )
        .expect_err("one below no-op work must fail");
        assert!(matches!(
            failure.error(),
            Error::LimitExceeded {
                kind: LimitKind::RewriteWork,
                ..
            }
        ));
    }
}
