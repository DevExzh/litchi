use std::io::{self, ErrorKind};

const INDEX_RECORD_TYPE: u16 = 0x020b;
const ROW_RECORD_TYPE: u16 = 0x0208;
const TABLE_RECORD_TYPE: u16 = 0x0236;
const DBCELL_RECORD_TYPE: u16 = 0x00d7;
const MAX_ROWS_PER_BLOCK: usize = 32;
const MAX_ROW_BLOCKS: usize = 2048;
const BIFF8_MAX_RECORD_PAYLOAD: usize = 8_224;

/// An encoded BIFF8 `ROW` record and the encoded cell records belonging to it.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct RowBlockLayoutRow {
    row: u16,
    row_record: Vec<u8>,
    cell_records: Vec<u8>,
}

impl RowBlockLayoutRow {
    pub fn new(row: u16, row_record: Vec<u8>, cell_records: Vec<u8>) -> Self {
        Self {
            row,
            row_record,
            cell_records,
        }
    }

    pub fn row(&self) -> u16 {
        self.row
    }
    pub fn row_record(&self) -> &[u8] {
        &self.row_record
    }
    pub fn cell_records(&self) -> &[u8] {
        &self.cell_records
    }
}

/// A fully checked INDEX/DBCELL layout for one serialized worksheet stream.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct RowBlockLayoutPlan {
    index_record_position: u32,
    row_table_position: u32,
    default_column_width_position: u32,
    dbcell_positions: Vec<u32>,
    index_record: Vec<u8>,
    row_table: Vec<u8>,
}

impl RowBlockLayoutPlan {
    /// Rebuilds a staged `ROW* CELL*` table into BIFF8 row blocks.
    ///
    /// The staged bytes must contain all ROW records first, followed by cell
    /// records. Every cell row must have a corresponding ROW record.
    pub fn generate_from_staged(
        index_record_position: u64,
        records_between_index_and_rows: u64,
        default_column_width_offset: u64,
        staged_row_table: &[u8],
    ) -> io::Result<Self> {
        Self::generate(
            index_record_position,
            records_between_index_and_rows,
            default_column_width_offset,
            decode_staged_rows(staged_row_table)?,
        )
    }

    /// Builds a layout without writing or patching caller-owned bytes.
    ///
    /// `records_between_index_and_rows` is the final byte length of all records
    /// after INDEX and before the first ROW. `default_column_width_offset` is
    /// the offset of DEFCOLWIDTH within those records.
    pub fn generate(
        index_record_position: u64,
        records_between_index_and_rows: u64,
        default_column_width_offset: u64,
        rows: Vec<RowBlockLayoutRow>,
    ) -> io::Result<Self> {
        if default_column_width_offset >= records_between_index_and_rows {
            return Err(invalid(
                "DEFCOLWIDTH is outside the records preceding the row table",
            ));
        }
        validate_rows(&rows)?;
        let blocks = logical_blocks(&rows)?;
        if blocks.len() > MAX_ROW_BLOCKS {
            return Err(invalid("worksheet has more than 2048 BIFF8 row blocks"));
        }

        let index_payload_len = 16usize
            .checked_add(blocks.len().checked_mul(4).ok_or_else(overflow)?)
            .ok_or_else(overflow)?;
        if index_payload_len > BIFF8_MAX_RECORD_PAYLOAD {
            return Err(invalid("INDEX record exceeds the BIFF8 record-size limit"));
        }
        let index_record_len = 4usize.checked_add(index_payload_len).ok_or_else(overflow)?;
        let row_table_position_u64 = index_record_position
            .checked_add(u64::try_from(index_record_len).map_err(|_| overflow())?)
            .and_then(|value| value.checked_add(records_between_index_and_rows))
            .ok_or_else(overflow)?;
        let default_column_width_position_u64 = index_record_position
            .checked_add(u64::try_from(index_record_len).map_err(|_| overflow())?)
            .and_then(|value| value.checked_add(default_column_width_offset))
            .ok_or_else(overflow)?;
        let index_record_position = checked_u32(index_record_position, "INDEX position")?;
        let row_table_position = checked_u32(row_table_position_u64, "row-table position")?;
        let default_column_width_position =
            checked_u32(default_column_width_position_u64, "DEFCOLWIDTH position")?;

        let mut row_table = Vec::new();
        let mut dbcell_positions = Vec::with_capacity(blocks.len());
        for block in blocks {
            append_block(
                &mut row_table,
                &mut dbcell_positions,
                row_table_position_u64,
                &rows[block],
            )?;
        }
        let planned_end = row_table_position_u64
            .checked_add(u64::try_from(row_table.len()).map_err(|_| overflow())?)
            .ok_or_else(overflow)?;
        checked_u32(planned_end, "worksheet stream end")?;

        let first_row = rows.first().map_or(0, |row| u32::from(row.row));
        let last_row_plus_one = rows.last().map_or(0, |row| u32::from(row.row) + 1);
        let index_record = encode_index(
            first_row,
            last_row_plus_one,
            default_column_width_position,
            &dbcell_positions,
        )?;
        Ok(Self {
            index_record_position,
            row_table_position,
            default_column_width_position,
            dbcell_positions,
            index_record,
            row_table,
        })
    }

    pub fn index_record_position(&self) -> u32 {
        self.index_record_position
    }
    pub fn row_table_position(&self) -> u32 {
        self.row_table_position
    }
    pub fn default_column_width_position(&self) -> u32 {
        self.default_column_width_position
    }
    pub fn dbcell_positions(&self) -> &[u32] {
        &self.dbcell_positions
    }
    pub fn index_record(&self) -> &[u8] {
        &self.index_record
    }
    pub fn row_table(&self) -> &[u8] {
        &self.row_table
    }
    pub fn into_records(self) -> (Vec<u8>, Vec<u8>) {
        (self.index_record, self.row_table)
    }
}

fn decode_staged_rows(bytes: &[u8]) -> io::Result<Vec<RowBlockLayoutRow>> {
    let mut rows = Vec::new();
    let mut offset = 0usize;
    let mut reached_cells = false;
    let mut last_cell_row: Option<usize> = None;
    while offset < bytes.len() {
        let header_end = offset.checked_add(4).ok_or_else(overflow)?;
        if header_end > bytes.len() {
            return Err(invalid("staged row table ends inside a BIFF header"));
        }
        let record_type = u16::from_le_bytes([bytes[offset], bytes[offset + 1]]);
        let payload_len = usize::from(u16::from_le_bytes([bytes[offset + 2], bytes[offset + 3]]));
        if payload_len > BIFF8_MAX_RECORD_PAYLOAD {
            return Err(invalid(
                "staged row-table record exceeds the BIFF8 record-size limit",
            ));
        }
        let record_end = header_end.checked_add(payload_len).ok_or_else(overflow)?;
        if record_end > bytes.len() {
            return Err(invalid("staged row table ends inside a BIFF payload"));
        }
        if record_type == ROW_RECORD_TYPE {
            if reached_cells {
                return Err(invalid(
                    "ROW record follows a cell record in the staged row table",
                ));
            }
            if payload_len != 16 {
                return Err(invalid("staged ROW record does not have a 16-byte payload"));
            }
            let row = u16::from_le_bytes([bytes[header_end], bytes[header_end + 1]]);
            rows.push(RowBlockLayoutRow::new(
                row,
                bytes[offset..record_end].to_vec(),
                Vec::new(),
            ));
        } else if record_type == TABLE_RECORD_TYPE {
            // A Table record follows its anchor Formula and shares that
            // cell's layout row; its own payload does not start with the
            // row coordinate of the anchor cell.
            reached_cells = true;
            let row_index = last_cell_row
                .ok_or_else(|| invalid("staged Table record has no preceding cell record"))?;
            rows[row_index]
                .cell_records
                .extend_from_slice(&bytes[offset..record_end]);
        } else {
            reached_cells = true;
            if !is_row_addressed_cell_record(record_type) || payload_len < 2 {
                return Err(invalid("staged row table contains a non-cell record"));
            }
            let row = u16::from_le_bytes([bytes[header_end], bytes[header_end + 1]]);
            let row_index = rows
                .binary_search_by_key(&row, RowBlockLayoutRow::row)
                .map_err(|_| invalid("staged cell record has no corresponding ROW record"))?;
            rows[row_index]
                .cell_records
                .extend_from_slice(&bytes[offset..record_end]);
            last_cell_row = Some(row_index);
        }
        offset = record_end;
    }
    Ok(rows)
}

fn validate_rows(rows: &[RowBlockLayoutRow]) -> io::Result<()> {
    let mut previous = None;
    for row in rows {
        if previous.is_some_and(|value| row.row <= value) {
            return Err(invalid("layout rows are not strictly increasing"));
        }
        validate_row_record(row)?;
        validate_cell_records(row)?;
        previous = Some(row.row);
    }
    Ok(())
}

fn logical_blocks(rows: &[RowBlockLayoutRow]) -> io::Result<Vec<std::ops::Range<usize>>> {
    let Some(first) = rows.first() else {
        return Ok(Vec::new());
    };
    let last = rows.last().ok_or_else(overflow)?;
    let span = usize::from(last.row)
        .checked_sub(usize::from(first.row))
        .and_then(|value| value.checked_add(1))
        .ok_or_else(overflow)?;
    let block_count = span
        .checked_add(MAX_ROWS_PER_BLOCK - 1)
        .ok_or_else(overflow)?
        / MAX_ROWS_PER_BLOCK;
    if block_count > MAX_ROW_BLOCKS {
        return Err(invalid("worksheet has more than 2048 BIFF8 row blocks"));
    }

    let mut blocks = Vec::with_capacity(block_count);
    let mut start = 0usize;
    for block_index in 0..block_count {
        let exclusive_row = usize::from(first.row)
            .checked_add(
                block_index
                    .checked_add(1)
                    .and_then(|value| value.checked_mul(MAX_ROWS_PER_BLOCK))
                    .ok_or_else(overflow)?,
            )
            .ok_or_else(overflow)?;
        let mut end = start;
        while end < rows.len() && usize::from(rows[end].row) < exclusive_row {
            end += 1;
        }
        blocks.push(start..end);
        start = end;
    }
    Ok(blocks)
}

fn validate_row_record(row: &RowBlockLayoutRow) -> io::Result<()> {
    if row.row_record.len() != 20 {
        return Err(invalid("ROW record must have a 16-byte BIFF8 payload"));
    }
    let record_type = u16::from_le_bytes([row.row_record[0], row.row_record[1]]);
    let payload_len = u16::from_le_bytes([row.row_record[2], row.row_record[3]]);
    let encoded_row = u16::from_le_bytes([row.row_record[4], row.row_record[5]]);
    if record_type != ROW_RECORD_TYPE || payload_len != 16 {
        return Err(invalid("row bytes are not one complete BIFF8 ROW record"));
    }
    if encoded_row != row.row {
        return Err(invalid(
            "ROW record coordinate does not match its layout row",
        ));
    }
    Ok(())
}

fn validate_cell_records(row: &RowBlockLayoutRow) -> io::Result<()> {
    let mut offset = 0usize;
    while offset < row.cell_records.len() {
        let header_end = offset.checked_add(4).ok_or_else(overflow)?;
        if header_end > row.cell_records.len() {
            return Err(invalid("cell record buffer ends inside a BIFF header"));
        }
        let record_type =
            u16::from_le_bytes([row.cell_records[offset], row.cell_records[offset + 1]]);
        let payload_len = usize::from(u16::from_le_bytes([
            row.cell_records[offset + 2],
            row.cell_records[offset + 3],
        ]));
        if payload_len > BIFF8_MAX_RECORD_PAYLOAD {
            return Err(invalid("cell record exceeds the BIFF8 record-size limit"));
        }
        let record_end = header_end.checked_add(payload_len).ok_or_else(overflow)?;
        if record_end > row.cell_records.len() {
            return Err(invalid("cell record buffer ends inside a BIFF payload"));
        }
        if record_type == TABLE_RECORD_TYPE {
            // Table records carry the table range, not the anchor cell's row.
            offset = record_end;
            continue;
        }
        if !is_row_addressed_cell_record(record_type) {
            return Err(invalid("record is not a supported BIFF8 cell-table record"));
        }
        if payload_len < 2 {
            return Err(invalid(
                "cell record is too short to contain a row coordinate",
            ));
        }
        let encoded_row = u16::from_le_bytes([
            row.cell_records[header_end],
            row.cell_records[header_end + 1],
        ]);
        if encoded_row != row.row {
            return Err(invalid(
                "cell record coordinate does not match its layout row",
            ));
        }
        offset = record_end;
    }
    Ok(())
}

fn append_block(
    row_table: &mut Vec<u8>,
    dbcell_positions: &mut Vec<u32>,
    row_table_position: u64,
    rows: &[RowBlockLayoutRow],
) -> io::Result<()> {
    let block_start = row_table.len();
    for row in rows {
        row_table.extend_from_slice(&row.row_record);
    }
    let mut cell_offsets = Vec::with_capacity(rows.len());
    let first_row_end = block_start
        .checked_add(rows.first().map_or(0, |row| row.row_record.len()))
        .ok_or_else(overflow)?;
    let mut previous_first_cell = None;
    for row in rows {
        if !row.cell_records.is_empty() {
            let first_cell = row_table.len();
            let base = previous_first_cell.unwrap_or(first_row_end);
            let offset = first_cell.checked_sub(base).ok_or_else(overflow)?;
            cell_offsets.push(checked_u16(offset, "DBCELL cell offset")?);
            previous_first_cell = Some(first_cell);
        }
        row_table.extend_from_slice(&row.cell_records);
    }
    let dbcell_relative = row_table.len();
    let dbcell_position = row_table_position
        .checked_add(u64::try_from(dbcell_relative).map_err(|_| overflow())?)
        .ok_or_else(overflow)?;
    dbcell_positions.push(checked_u32(dbcell_position, "DBCELL position")?);
    let block_data_len = row_table
        .len()
        .checked_sub(block_start)
        .ok_or_else(overflow)?;
    let row_offset = if rows.iter().all(|row| row.cell_records.is_empty()) {
        0
    } else {
        checked_u32(
            u64::try_from(block_data_len).map_err(|_| overflow())?,
            "DBCELL row offset",
        )?
    };
    let payload_len = 4usize
        .checked_add(cell_offsets.len().checked_mul(2).ok_or_else(overflow)?)
        .ok_or_else(overflow)?;
    push_header(row_table, DBCELL_RECORD_TYPE, payload_len)?;
    row_table.extend_from_slice(&row_offset.to_le_bytes());
    for offset in cell_offsets {
        row_table.extend_from_slice(&offset.to_le_bytes());
    }
    Ok(())
}

fn encode_index(
    first_row: u32,
    last_row_plus_one: u32,
    default_column_width_position: u32,
    dbcell_positions: &[u32],
) -> io::Result<Vec<u8>> {
    let payload_len = 16usize
        .checked_add(dbcell_positions.len().checked_mul(4).ok_or_else(overflow)?)
        .ok_or_else(overflow)?;
    let mut bytes = Vec::with_capacity(payload_len + 4);
    push_header(&mut bytes, INDEX_RECORD_TYPE, payload_len)?;
    bytes.extend_from_slice(&0u32.to_le_bytes());
    bytes.extend_from_slice(&first_row.to_le_bytes());
    bytes.extend_from_slice(&last_row_plus_one.to_le_bytes());
    bytes.extend_from_slice(&default_column_width_position.to_le_bytes());
    for position in dbcell_positions {
        bytes.extend_from_slice(&position.to_le_bytes());
    }
    Ok(bytes)
}

fn push_header(bytes: &mut Vec<u8>, record_type: u16, payload_len: usize) -> io::Result<()> {
    let payload_len = checked_u16(payload_len, "BIFF record payload length")?;
    bytes.extend_from_slice(&record_type.to_le_bytes());
    bytes.extend_from_slice(&payload_len.to_le_bytes());
    Ok(())
}

fn is_row_addressed_cell_record(record_type: u16) -> bool {
    matches!(
        record_type,
        0x0006 | 0x0201 | 0x0203 | 0x0204 | 0x0205 | 0x027e | 0x00bd | 0x00be | 0x00d6 | 0x00fd
    )
}

fn checked_u16(value: usize, name: &str) -> io::Result<u16> {
    u16::try_from(value).map_err(|_| invalid(&format!("{name} does not fit in 16 bits")))
}

fn checked_u32(value: u64, name: &str) -> io::Result<u32> {
    u32::try_from(value)
        .map_err(|_| invalid(&format!("{name} does not fit in a BIFF8 stream pointer")))
}

fn invalid(message: &str) -> io::Error {
    io::Error::new(ErrorKind::InvalidInput, message)
}
fn overflow() -> io::Error {
    invalid("BIFF8 worksheet layout arithmetic overflow")
}

#[cfg(test)]
mod tests {
    use super::{RowBlockLayoutPlan, RowBlockLayoutRow};

    fn row_record(row: u16) -> Vec<u8> {
        let mut bytes = vec![0x08, 0x02, 0x10, 0x00];
        bytes.extend_from_slice(&row.to_le_bytes());
        bytes.extend_from_slice(&[0, 0, 1, 0, 0xff, 0, 0, 0, 0, 0, 0, 0, 0, 0]);
        bytes
    }

    fn number_record(row: u16, column: u16) -> Vec<u8> {
        let mut bytes = vec![0x03, 0x02, 0x0e, 0x00];
        bytes.extend_from_slice(&row.to_le_bytes());
        bytes.extend_from_slice(&column.to_le_bytes());
        bytes.extend_from_slice(&0u16.to_le_bytes());
        bytes.extend_from_slice(&1.0f64.to_le_bytes());
        bytes
    }

    #[test]
    fn generates_exact_single_block_bytes() {
        let rows = vec![
            RowBlockLayoutRow::new(0, row_record(0), number_record(0, 0)),
            RowBlockLayoutRow::new(1, row_record(1), number_record(1, 0)),
        ];
        let plan = RowBlockLayoutPlan::generate(100, 40, 8, rows).unwrap();
        assert_eq!(plan.index_record_position(), 100);
        assert_eq!(plan.row_table_position(), 164);
        assert_eq!(plan.default_column_width_position(), 132);
        assert_eq!(plan.dbcell_positions(), &[240]);
        assert_eq!(
            plan.index_record(),
            &[
                0x0b, 0x02, 0x14, 0x00, 0, 0, 0, 0, 0, 0, 0, 0, 2, 0, 0, 0, 132, 0, 0, 0, 240, 0,
                0, 0
            ]
        );
        assert_eq!(
            &plan.row_table()[76..],
            &[0xd7, 0x00, 0x08, 0x00, 76, 0, 0, 0, 20, 0, 18, 0]
        );
    }

    #[test]
    fn splits_at_the_32_row_boundary() {
        let rows = (0..33)
            .map(|row| RowBlockLayoutRow::new(row, row_record(row), number_record(row, 0)))
            .collect();
        let plan = RowBlockLayoutPlan::generate(0, 8, 0, rows).unwrap();
        assert_eq!(plan.dbcell_positions().len(), 2);
        assert_eq!(plan.index_record().len(), 28);
    }

    #[test]
    fn emits_an_empty_dbcell_for_a_sparse_logical_block() {
        let rows = vec![
            RowBlockLayoutRow::new(1, row_record(1), number_record(1, 0)),
            RowBlockLayoutRow::new(76, row_record(76), number_record(76, 0)),
        ];
        let plan = RowBlockLayoutPlan::generate(0, 8, 0, rows).unwrap();
        assert_eq!(plan.dbcell_positions().len(), 3);
        let first_dbcell_end = 48;
        assert_eq!(
            &plan.row_table()[first_dbcell_end..first_dbcell_end + 8],
            &[0xd7, 0x00, 0x04, 0x00, 0, 0, 0, 0]
        );
    }

    #[test]
    fn rejects_malformed_layouts_before_serializing() {
        let wrong_row = RowBlockLayoutRow::new(2, row_record(1), number_record(2, 0));
        assert!(RowBlockLayoutPlan::generate(0, 8, 0, vec![wrong_row]).is_err());
        let truncated_cell = RowBlockLayoutRow::new(0, row_record(0), vec![3, 2, 14]);
        assert!(RowBlockLayoutPlan::generate(0, 8, 0, vec![truncated_cell]).is_err());
        let unordered = vec![
            RowBlockLayoutRow::new(1, row_record(1), Vec::new()),
            RowBlockLayoutRow::new(0, row_record(0), Vec::new()),
        ];
        assert!(RowBlockLayoutPlan::generate(0, 8, 0, unordered).is_err());
    }

    #[test]
    fn rejects_stream_pointer_overflow() {
        let rows = vec![RowBlockLayoutRow::new(
            0,
            row_record(0),
            number_record(0, 0),
        )];
        assert!(RowBlockLayoutPlan::generate(u64::from(u32::MAX), 8, 0, rows).is_err());
    }

    #[test]
    fn regenerates_a_staged_table_into_the_same_checked_layout() {
        let first = RowBlockLayoutRow::new(0, row_record(0), number_record(0, 0));
        let second = RowBlockLayoutRow::new(1, row_record(1), number_record(1, 0));
        let expected =
            RowBlockLayoutPlan::generate(100, 40, 8, vec![first.clone(), second.clone()]).unwrap();
        let mut staged = Vec::new();
        staged.extend_from_slice(first.row_record());
        staged.extend_from_slice(second.row_record());
        staged.extend_from_slice(first.cell_records());
        staged.extend_from_slice(second.cell_records());
        let regenerated = RowBlockLayoutPlan::generate_from_staged(100, 40, 8, &staged).unwrap();
        assert_eq!(regenerated, expected);
    }

    #[test]
    fn omits_rgdb_entries_for_rows_without_cells() {
        let rows = vec![
            RowBlockLayoutRow::new(0, row_record(0), Vec::new()),
            RowBlockLayoutRow::new(1, row_record(1), number_record(1, 0)),
        ];
        let plan = RowBlockLayoutPlan::generate(0, 8, 0, rows).unwrap();
        assert_eq!(
            &plan.row_table()[58..],
            &[0xd7, 0x00, 0x06, 0x00, 58, 0, 0, 0, 20, 0]
        );
    }
}
