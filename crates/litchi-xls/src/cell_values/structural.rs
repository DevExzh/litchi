use super::{
    BLANK, BOOL_ERR, Error, FORMULA, LABEL_SST, MUL_RK, NUMBER, RK, Reference, Result, Storage,
    StructuralChange, StyleIndex, Value,
};
use crate::writer::row_blocks::{RowBlockLayoutPlan, RowBlockLayoutRow};
use litchi_biff::Records;
use litchi_core::binary;
use std::collections::BTreeMap;

const BOUND_SHEET: u16 = 0x0085;
const CONTINUE: u16 = 0x003c;
const INDEX: u16 = 0x020b;
const ROW: u16 = 0x0208;
const DBCELL: u16 = 0x00d7;
const DIMENSIONS: u16 = 0x0200;
const DEF_COL_WIDTH: u16 = 0x0055;
const TABLE: u16 = 0x0236;
const SHARED_FORMULA: u16 = 0x04bc;
const ARRAY: u16 = 0x0221;
const STRING: u16 = 0x0207;
const LABEL: u16 = 0x0204;
const RSTRING: u16 = 0x00d6;
const MUL_BLANK: u16 = 0x00be;

#[derive(Clone)]
struct RawRecord {
    kind: u16,
    start: usize,
    end: usize,
}

#[derive(Clone)]
struct BoundSheet {
    record: RawRecord,
    position: u32,
}

struct RowData {
    row_record: Vec<u8>,
    cell_records: Vec<Vec<u8>>,
}

pub(super) fn apply(
    mut workbook: Vec<u8>,
    source: &super::Snapshot,
    changes: &[StructuralChange],
) -> Result<Vec<u8>> {
    for change in changes {
        if let StructuralChange::RenameSheet {
            sheet,
            before,
            after,
        } = change
        {
            rename_sheet(&mut workbook, source, *sheet, before, after)?;
        }
    }

    let mut by_sheet: BTreeMap<usize, Vec<&StructuralChange>> = BTreeMap::new();
    for change in changes {
        if matches!(change, StructuralChange::RenameSheet { .. }) {
            continue;
        }
        let sheet = operation_sheet(change);
        by_sheet.entry(sheet).or_default().push(change);
    }
    // Rewriting later streams first keeps earlier absolute offsets stable and
    // minimizes the number of BoundSheet position adjustments.
    for (sheet, operations) in by_sheet.into_iter().rev() {
        rewrite_worksheet(&mut workbook, source, sheet, &operations)?;
    }
    Ok(workbook)
}

pub(super) fn validate_sheet_name(name: &str) -> Result<()> {
    let units = name.encode_utf16().count();
    if name.is_empty() || units > 31 {
        return Err(Error::InvalidData(
            "BIFF8 worksheet names must contain 1 through 31 UTF-16 units".into(),
        ));
    }
    if name
        .chars()
        .any(|ch| matches!(ch, ':' | '\\' | '/' | '?' | '*' | '[' | ']'))
    {
        return Err(Error::InvalidData(
            "BIFF8 worksheet name contains a forbidden character".into(),
        ));
    }
    if name.starts_with('\'') || name.ends_with('\'') || name.chars().any(|ch| ch == '\0') {
        return Err(Error::InvalidData(
            "BIFF8 worksheet name has a forbidden quote or NUL".into(),
        ));
    }
    Ok(())
}

fn operation_sheet(change: &StructuralChange) -> usize {
    match change {
        StructuralChange::Cell { sheet, .. }
        | StructuralChange::Rows { sheet, .. }
        | StructuralChange::Columns { sheet, .. }
        | StructuralChange::RenameSheet { sheet, .. } => *sheet,
    }
}

fn rename_sheet(
    workbook: &mut Vec<u8>,
    source: &super::Snapshot,
    sheet: usize,
    before: &str,
    after: &str,
) -> Result<()> {
    validate_sheet_name(after)?;
    let workbook_index = source
        .inner
        .sheets
        .get(sheet)
        .ok_or_else(|| Error::UnsafeEdit("worksheet rename index is stale".into()))?
        .workbook_index;
    let bounds = bound_sheets(workbook)?;
    let bound = bounds
        .get(workbook_index)
        .ok_or_else(|| Error::UnsafeEdit("worksheet BoundSheet disappeared".into()))?;
    let payload = &workbook[bound.record.start + 4..bound.record.end];
    let decoded = decode_bound_name(payload)?;
    if decoded != before {
        return Err(Error::UnsafeEdit(
            "worksheet rename precondition is stale".into(),
        ));
    }
    let encoded = encode_bound_name(after)?;
    let payload_len = 7usize
        .checked_add(encoded.len())
        .ok_or_else(|| Error::InvalidData("BoundSheet length overflow".into()))?;
    let payload_len = u16::try_from(payload_len)
        .map_err(|_error| Error::InvalidData("BoundSheet payload exceeds u16".into()))?;
    let mut replacement = Vec::with_capacity(4 + usize::from(payload_len));
    replacement.extend_from_slice(&BOUND_SHEET.to_le_bytes());
    replacement.extend_from_slice(&payload_len.to_le_bytes());
    replacement.extend_from_slice(&payload[..6]);
    replacement.push(
        u8::try_from(after.encode_utf16().count())
            .map_err(|_error| Error::InvalidData("worksheet name length exceeds u8".into()))?,
    );
    replacement.extend_from_slice(&encoded);
    replace_range_and_adjust_bounds(workbook, bound.record.start, bound.record.end, &replacement)
}

fn rewrite_worksheet(
    workbook: &mut Vec<u8>,
    source: &super::Snapshot,
    sheet: usize,
    operations: &[&StructuralChange],
) -> Result<()> {
    if operations.iter().any(|operation| {
        matches!(
            operation,
            StructuralChange::Rows { .. } | StructuralChange::Columns { .. }
        )
    }) {
        certify_workbook_shift(workbook)?;
    }
    let workbook_index = source
        .inner
        .sheets
        .get(sheet)
        .ok_or_else(|| Error::UnsafeEdit("worksheet operation index is stale".into()))?
        .workbook_index;
    let bounds = bound_sheets(workbook)?;
    let bound = bounds
        .get(workbook_index)
        .ok_or_else(|| Error::UnsafeEdit("worksheet BoundSheet disappeared".into()))?;
    let start = usize::try_from(bound.position)
        .map_err(|_error| Error::InvalidData("worksheet position exceeds usize".into()))?;
    let end = bounds
        .iter()
        .filter_map(|candidate| {
            let position = usize::try_from(candidate.position).ok()?;
            (position > start).then_some(position)
        })
        .min()
        .unwrap_or(workbook.len());
    let worksheet = workbook
        .get(start..end)
        .ok_or_else(|| Error::InvalidData("worksheet range is outside Workbook".into()))?;
    let replacement = rebuild_sheet(worksheet, start, source, operations)?;
    replace_range_and_adjust_bounds(workbook, start, end, &replacement)
}

fn certify_workbook_shift(workbook: &[u8]) -> Result<()> {
    // These records can retain formulas or coordinates into any worksheet.
    // Without a complete token/range rewriter, moving coordinates would make
    // them stale even when the selected worksheet itself has no Formula cell.
    const GLOBAL_OR_CROSS_SHEET_DEPENDENCIES: &[u16] = &[
        FORMULA,
        TABLE,
        SHARED_FORMULA,
        ARRAY,
        0x0018,
        0x0023,
        0x1051,
        0x01b0,
        0x0879,
        0x087a,
        0x01be,
        0x01b8,
        0x00e5,
        0x009e,
    ];
    if let Some(record) = raw_records(workbook)?
        .iter()
        .find(|record| GLOBAL_OR_CROSS_SHEET_DEPENDENCIES.contains(&record.kind))
    {
        return Err(Error::UnsafeEdit(format!(
            "row/column movement cannot close workbook dependency record 0x{:04X}",
            record.kind
        )));
    }
    Ok(())
}

fn rebuild_sheet(
    worksheet: &[u8],
    absolute_start: usize,
    source: &super::Snapshot,
    operations: &[&StructuralChange],
) -> Result<Vec<u8>> {
    let records = raw_records(worksheet)?;
    let index = unique_kind(&records, INDEX, "INDEX")?;
    let dimensions = unique_kind(&records, DIMENSIONS, "DIMENSIONS")?;
    let def_col = unique_kind(&records, DEF_COL_WIDTH, "DEFCOLWIDTH")?;
    let first_row = records
        .iter()
        .position(|record| record.kind == ROW)
        .ok_or_else(|| {
            Error::UnsafeEdit("structural edit requires an existing BIFF8 row table".into())
        })?;
    let last_dbcell = records
        .iter()
        .rposition(|record| record.kind == DBCELL)
        .ok_or_else(|| Error::UnsafeEdit("structural edit requires an existing DBCELL".into()))?;
    if first_row <= index || last_dbcell < first_row {
        return Err(Error::InvalidData(
            "worksheet row-block order is invalid".into(),
        ));
    }
    let shifting = operations.iter().any(|operation| {
        matches!(
            operation,
            StructuralChange::Rows { .. } | StructuralChange::Columns { .. }
        )
    });
    if shifting {
        certify_coordinate_shift(&records)?;
    }

    let mut rows = collect_rows(worksheet, &records[first_row..=last_dbcell])?;
    for operation in operations {
        match operation {
            StructuralChange::Cell {
                reference,
                before,
                after,
                ..
            } => {
                apply_cell_change(
                    &mut rows,
                    *reference,
                    before.as_ref(),
                    after.as_ref(),
                    source,
                )?;
            },
            StructuralChange::Rows {
                start,
                count,
                insert,
                ..
            } => {
                shift_rows(&mut rows, *start, *count, *insert)?;
            },
            StructuralChange::Columns {
                start,
                count,
                insert,
                ..
            } => {
                shift_columns(&mut rows, *start, *count, *insert)?;
            },
            StructuralChange::RenameSheet { .. } => {},
        }
    }
    patch_row_extents(&mut rows)?;

    let row_table_start = records[first_row].start;
    let row_table_end = records[last_dbcell].end;
    let index_record = &records[index];
    let between = worksheet
        .get(index_record.end..row_table_start)
        .ok_or_else(|| Error::InvalidData("INDEX-to-row range is invalid".into()))?;
    let def_offset = records[def_col]
        .start
        .checked_sub(index_record.end)
        .ok_or_else(|| Error::InvalidData("DEFCOLWIDTH precedes INDEX".into()))?;
    let layout_rows = rows
        .into_iter()
        .map(|(row, data)| {
            let mut cells = Vec::new();
            for record in data.cell_records {
                cells.extend_from_slice(&record);
            }
            RowBlockLayoutRow::new(row, data.row_record, cells)
        })
        .collect();
    let absolute_index = absolute_start
        .checked_add(index_record.start)
        .ok_or_else(|| Error::InvalidData("absolute INDEX offset overflow".into()))?;
    let plan = RowBlockLayoutPlan::generate(
        u64::try_from(absolute_index)
            .map_err(|_error| Error::InvalidData("INDEX offset exceeds u64".into()))?,
        u64::try_from(between.len())
            .map_err(|_error| Error::InvalidData("INDEX gap exceeds u64".into()))?,
        u64::try_from(def_offset)
            .map_err(|_error| Error::InvalidData("DEFCOLWIDTH offset exceeds u64".into()))?,
        layout_rows,
    )
    .map_err(|error| Error::UnsafeEdit(format!("BIFF8 row-block regeneration: {error}")))?;
    let (new_index, new_rows) = plan.into_records();

    let (first_used_row, last_used_row, first_used_col, last_used_col) =
        dimensions_from_row_table(&new_rows)?;
    let mut new_dimensions = Vec::new();
    crate::writer::biff::write_dimensions(
        &mut new_dimensions,
        first_used_row,
        last_used_row,
        first_used_col,
        last_used_col,
    )?;
    if new_dimensions.len() != records[dimensions].end - records[dimensions].start {
        return Err(Error::UnsafeEdit(
            "DIMENSIONS encoding changed fixed record width".into(),
        ));
    }
    let mut between = between.to_vec();
    let dimensions_relative = records[dimensions]
        .start
        .checked_sub(index_record.end)
        .ok_or_else(|| Error::InvalidData("DIMENSIONS precedes INDEX".into()))?;
    let dimensions_end = dimensions_relative
        .checked_add(new_dimensions.len())
        .ok_or_else(|| Error::InvalidData("DIMENSIONS range overflow".into()))?;
    between
        .get_mut(dimensions_relative..dimensions_end)
        .ok_or_else(|| Error::InvalidData("DIMENSIONS is outside INDEX gap".into()))?
        .copy_from_slice(&new_dimensions);

    let mut result = Vec::new();
    result.extend_from_slice(&worksheet[..index_record.start]);
    result.extend_from_slice(&new_index);
    result.extend_from_slice(&between);
    result.extend_from_slice(&new_rows);
    result.extend_from_slice(&worksheet[row_table_end..]);
    patch_selection_records(&mut result, operations)?;
    Ok(result)
}

fn patch_selection_records(worksheet: &mut [u8], operations: &[&StructuralChange]) -> Result<()> {
    let records = raw_records(worksheet)?;
    for operation in operations {
        for record in records.iter().filter(|record| record.kind == 0x001d) {
            let payload = worksheet
                .get_mut(record.start + 4..record.end)
                .ok_or_else(|| Error::InvalidData("SELECTION payload is truncated".into()))?;
            patch_selection(payload, operation)?;
        }
    }
    Ok(())
}

fn patch_selection(payload: &mut [u8], operation: &StructuralChange) -> Result<()> {
    if payload.len() < 9 {
        return Err(Error::InvalidData(
            "SELECTION payload is shorter than nine bytes".into(),
        ));
    }
    let range_count = usize::from(binary::read_u16_le_at(payload, 7)?);
    let expected = 9usize
        .checked_add(
            range_count
                .checked_mul(6)
                .ok_or_else(|| Error::InvalidData("SELECTION range size overflow".into()))?,
        )
        .ok_or_else(|| Error::InvalidData("SELECTION payload size overflow".into()))?;
    if payload.len() != expected {
        return Err(Error::InvalidData(
            "SELECTION range count does not match its payload".into(),
        ));
    }
    match operation {
        StructuralChange::Rows {
            start,
            count,
            insert,
            ..
        } => {
            let active = binary::read_u16_le_at(payload, 1)?;
            let active = shift_axis_point(active, *start, *count, *insert, 65_536)?;
            payload[1..3].copy_from_slice(&active.to_le_bytes());
            for range in payload[9..].chunks_exact_mut(6) {
                let first = binary::read_u16_le_at(range, 0)?;
                let last = binary::read_u16_le_at(range, 2)?;
                let (first, last) = shift_axis_range(first, last, *start, *count, *insert, 65_536)?;
                range[0..2].copy_from_slice(&first.to_le_bytes());
                range[2..4].copy_from_slice(&last.to_le_bytes());
            }
        },
        StructuralChange::Columns {
            start,
            count,
            insert,
            ..
        } => {
            let active = binary::read_u16_le_at(payload, 3)?;
            let active =
                shift_axis_point(active, u16::from(*start), u16::from(*count), *insert, 256)?;
            payload[3..5].copy_from_slice(&active.to_le_bytes());
            for range in payload[9..].chunks_exact_mut(6) {
                let (first, last) = shift_axis_range(
                    u16::from(range[4]),
                    u16::from(range[5]),
                    u16::from(*start),
                    u16::from(*count),
                    *insert,
                    256,
                )?;
                range[4] = u8::try_from(first).map_err(|_error| {
                    Error::InvalidData("shifted SELECTION column exceeds u8".into())
                })?;
                range[5] = u8::try_from(last).map_err(|_error| {
                    Error::InvalidData("shifted SELECTION column exceeds u8".into())
                })?;
            }
        },
        StructuralChange::Cell { .. } | StructuralChange::RenameSheet { .. } => {},
    }
    Ok(())
}

fn shift_axis_point(value: u16, start: u16, count: u16, insert: bool, limit: u32) -> Result<u16> {
    let value = u32::from(value);
    let start = u32::from(start);
    let count = u32::from(count);
    let shifted = if insert {
        if value < start {
            value
        } else {
            value
                .checked_add(count)
                .ok_or_else(|| Error::InvalidData("SELECTION coordinate shift overflows".into()))?
        }
    } else {
        let end = start.saturating_add(count).min(limit);
        let removed = end - start;
        if value < start {
            value
        } else if value >= end {
            value - removed
        } else {
            start.min(limit - 1)
        }
    };
    if shifted >= limit {
        return Err(Error::UnsafeEdit(
            "row/column insertion moves a SELECTION coordinate outside BIFF8".into(),
        ));
    }
    u16::try_from(shifted)
        .map_err(|_error| Error::InvalidData("SELECTION coordinate exceeds u16".into()))
}

fn shift_axis_range(
    first: u16,
    last: u16,
    start: u16,
    count: u16,
    insert: bool,
    limit: u32,
) -> Result<(u16, u16)> {
    if first > last {
        return Err(Error::InvalidData("SELECTION range is inverted".into()));
    }
    if insert && first < start && last >= start {
        let shifted_last = shift_axis_point(last, start, count, true, limit)?;
        return Ok((first, shifted_last));
    }
    let first = shift_axis_point(first, start, count, insert, limit)?;
    let last = shift_axis_point(last, start, count, insert, limit)?;
    Ok((first.min(last), first.max(last)))
}

fn collect_rows(worksheet: &[u8], records: &[RawRecord]) -> Result<BTreeMap<u16, RowData>> {
    let mut rows = BTreeMap::new();
    let mut last_cell_row = None;
    for record in records {
        let bytes = worksheet[record.start..record.end].to_vec();
        match record.kind {
            ROW => {
                let row = binary::read_u16_le_at(&bytes, 4)?;
                if rows
                    .insert(
                        row,
                        RowData {
                            row_record: bytes,
                            cell_records: Vec::new(),
                        },
                    )
                    .is_some()
                {
                    return Err(Error::UnsafeEdit(
                        "duplicate ROW record in row table".into(),
                    ));
                }
            },
            DBCELL => {},
            TABLE | SHARED_FORMULA | ARRAY => {
                let row = last_cell_row.ok_or_else(|| {
                    Error::InvalidData("formula continuation has no preceding cell".into())
                })?;
                rows.get_mut(&row)
                    .ok_or_else(|| Error::InvalidData("formula continuation row is absent".into()))?
                    .cell_records
                    .push(bytes);
            },
            STRING | CONTINUE => {
                return Err(Error::UnsafeEdit(
                    "formula string caches and continued row records are not yet structurally rewritable".into(),
                ));
            },
            kind if is_cell_record(kind) => {
                let row = binary::read_u16_le_at(&bytes, 4)?;
                rows.get_mut(&row)
                    .ok_or_else(|| Error::InvalidData("cell record has no ROW owner".into()))?
                    .cell_records
                    .push(bytes);
                last_cell_row = Some(row);
            },
            _ => {
                return Err(Error::UnsafeEdit(format!(
                    "record 0x{:04X} inside the BIFF8 row table is not losslessly classified",
                    record.kind
                )));
            },
        }
    }
    Ok(rows)
}

fn apply_cell_change(
    rows: &mut BTreeMap<u16, RowData>,
    reference: Reference,
    before: Option<&(Storage, Value, StyleIndex)>,
    after: Option<&(Storage, Value, StyleIndex)>,
    source: &super::Snapshot,
) -> Result<()> {
    let mut found = None;
    if let Some(row) = rows.get_mut(&reference.row()) {
        for (index, record) in row.cell_records.iter().enumerate() {
            if record_contains_reference(record, reference)? && found.replace(index).is_some() {
                return Err(Error::UnsafeEdit(
                    "structural cell target is ambiguous".into(),
                ));
            }
        }
    }
    match (before, after, found) {
        (None, Some(after), None) => {
            if let std::collections::btree_map::Entry::Vacant(row) = rows.entry(reference.row()) {
                let mut row_record = Vec::new();
                crate::writer::biff::write_row(
                    &mut row_record,
                    u32::from(reference.row()),
                    u16::from(reference.column()),
                    u16::from(reference.column()) + 1,
                    255,
                    false,
                )?;
                row.insert(RowData {
                    row_record,
                    cell_records: Vec::new(),
                });
            }
            let record = encode_cell(reference, after, source)?;
            rows.get_mut(&reference.row())
                .ok_or_else(|| Error::InvalidData("inserted ROW disappeared".into()))?
                .cell_records
                .push(record);
        },
        (Some(expected), None, Some(index)) => {
            verify_record_state(
                &rows[&reference.row()].cell_records[index],
                reference,
                expected,
                source,
            )?;
            rows.get_mut(&reference.row())
                .ok_or_else(|| Error::InvalidData("removed ROW disappeared".into()))?
                .cell_records
                .remove(index);
        },
        (Some(expected), Some(after), Some(index)) => {
            let row = rows
                .get_mut(&reference.row())
                .ok_or_else(|| Error::InvalidData("styled ROW disappeared".into()))?;
            verify_record_state(&row.cell_records[index], reference, expected, source)?;
            if expected.0 != after.0 || !super::values_equal(&expected.1, &after.1) {
                return Err(Error::UnsafeEdit(
                    "structural replacement may only change an existing cell style".into(),
                ));
            }
            patch_record_style(&mut row.cell_records[index], reference, after.2)?;
        },
        (None, Some(_), Some(_)) => {
            return Err(Error::UnsafeEdit(
                "cell insertion precondition is stale".into(),
            ));
        },
        (Some(_), _, None) => {
            return Err(Error::UnsafeEdit(
                "structural cell precondition is stale".into(),
            ));
        },
        (None, None, _) => {
            return Err(Error::InvalidData(
                "structural cell operation has no outcome".into(),
            ));
        },
    }
    Ok(())
}

fn shift_rows(
    rows: &mut BTreeMap<u16, RowData>,
    start: u16,
    count: u16,
    insert: bool,
) -> Result<()> {
    let mut shifted = BTreeMap::new();
    let delete_end = start.saturating_add(count);
    for (row, mut data) in std::mem::take(rows) {
        let target = if insert {
            if row < start {
                row
            } else {
                row.checked_add(count).ok_or_else(|| {
                    Error::UnsafeEdit("row insertion moves content outside BIFF8".into())
                })?
            }
        } else if row < start {
            row
        } else if row < delete_end {
            continue;
        } else {
            row - count
        };
        patch_row_number(&mut data, target)?;
        shifted.insert(target, data);
    }
    *rows = shifted;
    Ok(())
}

fn shift_columns(
    rows: &mut BTreeMap<u16, RowData>,
    start: u8,
    count: u8,
    insert: bool,
) -> Result<()> {
    let delete_end = start.saturating_add(count);
    for data in rows.values_mut() {
        let mut retained = Vec::new();
        for mut record in std::mem::take(&mut data.cell_records) {
            if record_kind(&record)? == MUL_RK || record_kind(&record)? == MUL_BLANK {
                return Err(Error::UnsafeEdit(
                    "column movement across packed cell records is refused".into(),
                ));
            }
            let column = cell_column(&record)?;
            let target = if insert {
                if column < start {
                    column
                } else {
                    column.checked_add(count).ok_or_else(|| {
                        Error::UnsafeEdit("column insertion moves content outside BIFF8".into())
                    })?
                }
            } else if column < start {
                column
            } else if column < delete_end {
                continue;
            } else {
                column - count
            };
            record[6..8].copy_from_slice(&u16::from(target).to_le_bytes());
            retained.push(record);
        }
        data.cell_records = retained;
    }
    Ok(())
}

fn patch_row_number(data: &mut RowData, row: u16) -> Result<()> {
    data.row_record
        .get_mut(4..6)
        .ok_or_else(|| Error::InvalidData("ROW record is truncated".into()))?
        .copy_from_slice(&row.to_le_bytes());
    for record in &mut data.cell_records {
        if matches!(record_kind(record)?, TABLE | SHARED_FORMULA | ARRAY) {
            return Err(Error::UnsafeEdit(
                "row movement of formula ownership records is refused".into(),
            ));
        }
        record
            .get_mut(4..6)
            .ok_or_else(|| Error::InvalidData("cell record is truncated".into()))?
            .copy_from_slice(&row.to_le_bytes());
    }
    Ok(())
}

fn patch_row_extents(rows: &mut BTreeMap<u16, RowData>) -> Result<()> {
    for data in rows.values_mut() {
        let mut first = u16::MAX;
        let mut last = 0_u16;
        for record in &data.cell_records {
            let kind = record_kind(record)?;
            if matches!(kind, TABLE | SHARED_FORMULA | ARRAY) {
                continue;
            }
            let start = u16::from(cell_column(record)?);
            let end = if matches!(kind, MUL_RK | MUL_BLANK) {
                let payload_len = usize::from(u16::from_le_bytes([record[2], record[3]]));
                binary::read_u16_le_at(record, 4 + payload_len - 2)?
            } else {
                start
            };
            first = first.min(start);
            last = last.max(end);
        }
        let (first, last_plus_one) = if first == u16::MAX {
            (0, 0)
        } else {
            (first, last + 1)
        };
        data.row_record
            .get_mut(6..8)
            .ok_or_else(|| Error::InvalidData("ROW colMic is truncated".into()))?
            .copy_from_slice(&first.to_le_bytes());
        data.row_record
            .get_mut(8..10)
            .ok_or_else(|| Error::InvalidData("ROW colMac is truncated".into()))?
            .copy_from_slice(&last_plus_one.to_le_bytes());
    }
    Ok(())
}

fn dimensions_from_row_table(bytes: &[u8]) -> Result<(u32, u32, u16, u16)> {
    let records = raw_records(bytes)?;
    let mut first_row = u32::MAX;
    let mut last_row = 0_u32;
    let mut first_col = u16::MAX;
    let mut last_col = 0_u16;
    for record in records {
        if !is_cell_record(record.kind) {
            continue;
        }
        let raw = &bytes[record.start..record.end];
        let row = u32::from(binary::read_u16_le_at(raw, 4)?);
        let col = u16::from(cell_column(raw)?);
        let final_col = if matches!(record.kind, MUL_RK | MUL_BLANK) {
            binary::read_u16_le_at(raw, raw.len() - 2)?
        } else {
            col
        };
        first_row = first_row.min(row);
        last_row = last_row.max(row + 1);
        first_col = first_col.min(col);
        last_col = last_col.max(final_col + 1);
    }
    if first_row == u32::MAX {
        Ok((0, 0, 0, 0))
    } else {
        Ok((first_row, last_row, first_col, last_col))
    }
}

fn encode_cell(
    reference: Reference,
    state: &(Storage, Value, StyleIndex),
    source: &super::Snapshot,
) -> Result<Vec<u8>> {
    let (storage, value, style) = state;
    if usize::from(style.get()) >= source.inner.xf_records.len() {
        return Err(Error::UnsafeEdit(
            "inserted cell XF resource is stale".into(),
        ));
    }
    let mut record = Vec::new();
    match (storage, value) {
        (Storage::Number, Value::Number(number)) => crate::writer::biff::write_number(
            &mut record,
            u32::from(reference.row()),
            u16::from(reference.column()),
            style.get(),
            *number,
        )?,
        (Storage::LabelSst, Value::Text(text)) => {
            let index = source
                .inner
                .shared_strings
                .iter()
                .position(|candidate| candidate == text)
                .ok_or_else(|| Error::UnsafeEdit("inserted text is absent from SST".into()))?;
            crate::writer::biff::write_labelsst(
                &mut record,
                u32::from(reference.row()),
                u16::from(reference.column()),
                style.get(),
                u32::try_from(index)
                    .map_err(|_error| Error::InvalidData("SST index exceeds u32".into()))?,
            )?;
        },
        (Storage::BoolErr, Value::Boolean(value)) => crate::writer::biff::write_boolerr(
            &mut record,
            u32::from(reference.row()),
            u16::from(reference.column()),
            style.get(),
            *value,
        )?,
        (Storage::BoolErr, Value::Error(error)) => {
            record = scalar_record(BOOL_ERR, reference, style, &[error.code(), 1]);
        },
        (Storage::Blank, Value::Blank) => {
            record = scalar_record(BLANK, reference, style, &[]);
        },
        _ => {
            return Err(Error::UnsafeEdit(
                "inserted cell value has no standalone BIFF8 encoding".into(),
            ));
        },
    }
    Ok(record)
}

fn scalar_record(kind: u16, reference: Reference, style: &StyleIndex, tail: &[u8]) -> Vec<u8> {
    let length = 6_u16 + u16::try_from(tail.len()).unwrap_or(0);
    let mut record = Vec::with_capacity(4 + usize::from(length));
    record.extend_from_slice(&kind.to_le_bytes());
    record.extend_from_slice(&length.to_le_bytes());
    record.extend_from_slice(&reference.row().to_le_bytes());
    record.extend_from_slice(&u16::from(reference.column()).to_le_bytes());
    record.extend_from_slice(&style.get().to_le_bytes());
    record.extend_from_slice(tail);
    record
}

fn verify_record_state(
    record: &[u8],
    reference: Reference,
    expected: &(Storage, Value, StyleIndex),
    source: &super::Snapshot,
) -> Result<()> {
    let kind = record_kind(record)?;
    if kind != super::storage_record_kind(expected.0)
        || record_style(record, reference)? != expected.2
    {
        return Err(Error::UnsafeEdit("structural cell state is stale".into()));
    }
    // Snapshot parsing already validated the semantic value. Reconfirm text
    // resources because the structural path can change Workbook offsets.
    if let Value::Text(text) = &expected.1
        && !source
            .inner
            .shared_strings
            .iter()
            .any(|candidate| candidate == text)
    {
        return Err(Error::UnsafeEdit(
            "structural cell SST dependency is stale".into(),
        ));
    }
    Ok(())
}

fn patch_record_style(record: &mut [u8], reference: Reference, style: StyleIndex) -> Result<()> {
    if record_kind(record)? == MUL_RK {
        let first = binary::read_u16_le_at(record, 6)?;
        let index = usize::from(
            u16::from(reference.column())
                .checked_sub(first)
                .ok_or_else(|| {
                    Error::UnsafeEdit("MulRk style target precedes packed range".into())
                })?,
        );
        let offset = 8usize
            .checked_add(
                index
                    .checked_mul(6)
                    .ok_or_else(|| Error::InvalidData("MulRk style offset overflow".into()))?,
            )
            .ok_or_else(|| Error::InvalidData("MulRk style offset overflow".into()))?;
        record
            .get_mut(offset..offset + 2)
            .ok_or_else(|| Error::InvalidData("MulRk style field is truncated".into()))?
            .copy_from_slice(&style.get().to_le_bytes());
    } else {
        record
            .get_mut(8..10)
            .ok_or_else(|| Error::InvalidData("cell XF field is truncated".into()))?
            .copy_from_slice(&style.get().to_le_bytes());
    }
    Ok(())
}

fn record_style(record: &[u8], reference: Reference) -> Result<StyleIndex> {
    let index = if record_kind(record)? == MUL_RK {
        let first = binary::read_u16_le_at(record, 6)?;
        let index = usize::from(
            u16::from(reference.column())
                .checked_sub(first)
                .ok_or_else(|| {
                    Error::UnsafeEdit("MulRk style target precedes packed range".into())
                })?,
        );
        binary::read_u16_le_at(record, 8 + index * 6)?
    } else {
        binary::read_u16_le_at(record, 8)?
    };
    Ok(StyleIndex(index))
}

fn record_contains_reference(record: &[u8], reference: Reference) -> Result<bool> {
    let kind = record_kind(record)?;
    if !is_cell_record(kind) {
        return Ok(false);
    }
    if binary::read_u16_le_at(record, 4)? != reference.row() {
        return Ok(false);
    }
    let first = binary::read_u16_le_at(record, 6)?;
    let column = u16::from(reference.column());
    if matches!(kind, MUL_RK | MUL_BLANK) {
        let last = binary::read_u16_le_at(record, record.len() - 2)?;
        Ok((first..=last).contains(&column))
    } else {
        Ok(first == column)
    }
}

fn cell_column(record: &[u8]) -> Result<u8> {
    u8::try_from(binary::read_u16_le_at(record, 6)?)
        .map_err(|_error| Error::InvalidData("cell column exceeds BIFF8".into()))
}

fn certify_coordinate_shift(records: &[RawRecord]) -> Result<()> {
    const DEPENDENT: &[u16] = &[
        FORMULA,
        TABLE,
        SHARED_FORMULA,
        ARRAY,
        STRING,
        0x00e5,
        0x01b8,
        0x01be,
        0x01b2,
        0x01b0,
        0x0879,
        0x087a,
        0x001c,
        0x005d,
        0x00ec,
        0x007d,
        0x009e,
        0x0041,
        0x001a,
        0x001b,
        0x0090,
    ];
    if let Some(record) = records
        .iter()
        .find(|record| DEPENDENT.contains(&record.kind))
    {
        return Err(Error::UnsafeEdit(format!(
            "row/column movement cannot close dependency record 0x{:04X}",
            record.kind
        )));
    }
    Ok(())
}

fn replace_range_and_adjust_bounds(
    workbook: &mut Vec<u8>,
    start: usize,
    end: usize,
    replacement: &[u8],
) -> Result<()> {
    let old_len = end
        .checked_sub(start)
        .ok_or_else(|| Error::InvalidData("replacement range is reversed".into()))?;
    let delta = i64::try_from(replacement.len())
        .ok()
        .and_then(|new| {
            i64::try_from(old_len)
                .ok()
                .and_then(|old| new.checked_sub(old))
        })
        .ok_or_else(|| Error::InvalidData("Workbook replacement delta overflow".into()))?;
    let old_bounds = bound_sheets(workbook)?;
    workbook.splice(start..end, replacement.iter().copied());
    if delta == 0 {
        return Ok(());
    }
    let threshold = u32::try_from(end)
        .map_err(|_error| Error::InvalidData("Workbook replacement offset exceeds u32".into()))?;
    for bound in old_bounds {
        if bound.position < threshold {
            continue;
        }
        let record_start = adjusted_offset(bound.record.start, start, end, delta)?;
        let position = i64::from(bound.position)
            .checked_add(delta)
            .and_then(|value| u32::try_from(value).ok())
            .ok_or_else(|| Error::UnsafeEdit("BoundSheet position adjustment overflows".into()))?;
        workbook
            .get_mut(record_start + 4..record_start + 8)
            .ok_or_else(|| Error::InvalidData("BoundSheet position field is truncated".into()))?
            .copy_from_slice(&position.to_le_bytes());
    }
    Ok(())
}

fn adjusted_offset(offset: usize, start: usize, end: usize, delta: i64) -> Result<usize> {
    if offset < start {
        return Ok(offset);
    }
    if offset >= end {
        return usize::try_from(
            i64::try_from(offset)
                .map_err(|_error| Error::InvalidData("record offset exceeds i64".into()))?
                .checked_add(delta)
                .ok_or_else(|| Error::InvalidData("record offset adjustment overflows".into()))?,
        )
        .map_err(|_error| Error::InvalidData("adjusted record offset exceeds usize".into()));
    }
    // The only supported overlap is replacement of that exact BoundSheet
    // record during a sheet rename; its rebuilt record starts at `start`.
    Ok(start)
}

fn bound_sheets(workbook: &[u8]) -> Result<Vec<BoundSheet>> {
    let records = raw_records(workbook)?;
    let mut bounds = Vec::new();
    for record in records {
        if record.kind == 0x000a {
            break;
        }
        if record.kind != BOUND_SHEET {
            continue;
        }
        let payload = workbook
            .get(record.start + 4..record.end)
            .ok_or_else(|| Error::InvalidData("BoundSheet payload is truncated".into()))?;
        if payload.len() < 8 {
            return Err(Error::InvalidData("BoundSheet payload is truncated".into()));
        }
        bounds.push(BoundSheet {
            position: binary::read_u32_le_at(payload, 0)?,
            record,
        });
    }
    Ok(bounds)
}

fn decode_bound_name(payload: &[u8]) -> Result<String> {
    if payload.len() < 8 {
        return Err(Error::InvalidData("BoundSheet name is truncated".into()));
    }
    let units = usize::from(payload[6]);
    let wide = payload[7] & 1 != 0;
    if wide {
        let byte_len = units
            .checked_mul(2)
            .ok_or_else(|| Error::InvalidData("BoundSheet name length overflow".into()))?;
        let bytes = payload
            .get(8..8 + byte_len)
            .ok_or_else(|| Error::InvalidData("wide BoundSheet name is truncated".into()))?;
        let utf16: Vec<u16> = bytes
            .chunks_exact(2)
            .map(|pair| u16::from_le_bytes([pair[0], pair[1]]))
            .collect();
        String::from_utf16(&utf16)
            .map_err(|error| Error::InvalidData(format!("invalid BoundSheet UTF-16: {error}")))
    } else {
        let bytes = payload
            .get(8..8 + units)
            .ok_or_else(|| Error::InvalidData("compressed BoundSheet name is truncated".into()))?;
        Ok(bytes.iter().map(|byte| char::from(*byte)).collect())
    }
}

fn encode_bound_name(name: &str) -> Result<Vec<u8>> {
    let units: Vec<u16> = name.encode_utf16().collect();
    let compressed = units.iter().all(|unit| *unit <= 0xff);
    let mut encoded = Vec::new();
    encoded.push(u8::from(!compressed));
    if compressed {
        for unit in units {
            encoded.push(u8::try_from(unit).map_err(|_error| {
                Error::InvalidData("compressed BoundSheet character exceeds u8".into())
            })?);
        }
    } else {
        for unit in units {
            encoded.extend_from_slice(&unit.to_le_bytes());
        }
    }
    Ok(encoded)
}

fn raw_records(bytes: &[u8]) -> Result<Vec<RawRecord>> {
    let mut records = Vec::new();
    for item in Records::new(bytes) {
        let record = item?;
        let start = record.offset();
        let end = start
            .checked_add(4)
            .and_then(|value| value.checked_add(record.payload().len()))
            .ok_or_else(|| Error::InvalidData("BIFF record range overflow".into()))?;
        records.push(RawRecord {
            kind: record.kind().get(),
            start,
            end,
        });
    }
    Ok(records)
}

fn unique_kind(records: &[RawRecord], kind: u16, label: &str) -> Result<usize> {
    let mut matches = records
        .iter()
        .enumerate()
        .filter(|(_, record)| record.kind == kind);
    let index = matches
        .next()
        .map(|(index, _)| index)
        .ok_or_else(|| Error::UnsafeEdit(format!("worksheet has no {label} record")))?;
    if matches.next().is_some() {
        return Err(Error::UnsafeEdit(format!(
            "worksheet has duplicate {label} records"
        )));
    }
    Ok(index)
}

fn record_kind(record: &[u8]) -> Result<u16> {
    Ok(binary::read_u16_le_at(record, 0)?)
}

fn is_cell_record(kind: u16) -> bool {
    matches!(
        kind,
        FORMULA | BLANK | NUMBER | LABEL | BOOL_ERR | RK | MUL_RK | MUL_BLANK | RSTRING | LABEL_SST
    )
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn selection_coordinates_follow_row_insertion_and_inverse_deletion() {
        let mut payload = vec![3, 3, 0, 2, 0, 0, 0, 1, 0, 3, 0, 4, 0, 2, 5];
        let insert = StructuralChange::Rows {
            sheet: 0,
            start: 2,
            count: 2,
            insert: true,
        };
        patch_selection(&mut payload, &insert).unwrap();
        assert_eq!(binary::read_u16_le_at(&payload, 1).unwrap(), 5);
        assert_eq!(binary::read_u16_le_at(&payload, 9).unwrap(), 5);
        assert_eq!(binary::read_u16_le_at(&payload, 11).unwrap(), 6);

        let delete = StructuralChange::Rows {
            sheet: 0,
            start: 2,
            count: 2,
            insert: false,
        };
        patch_selection(&mut payload, &delete).unwrap();
        assert_eq!(payload, vec![3, 3, 0, 2, 0, 0, 0, 1, 0, 3, 0, 4, 0, 2, 5]);
    }
}
