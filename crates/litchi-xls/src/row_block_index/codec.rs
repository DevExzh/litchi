//! BIFF8 `INDEX` and `DBCELL` record codecs and worksheet collector.

use std::collections::BTreeMap;

use super::model::{DbCellRecord, IndexedRow, RowBlock, RowBlockIndex, WorksheetIndexRecord};
use super::{
    BOF_RECORD_TYPE, DBCELL_RECORD_TYPE, DEF_COL_WIDTH_RECORD_TYPE, EOF_RECORD_TYPE,
    INDEX_FIXED_LEN, INDEX_RECORD_TYPE, MAX_DBCELL_PAYLOAD_LEN, MAX_INDEX_PAYLOAD_LEN,
    MAX_ROW_BLOCKS, MAX_ROWS_PER_BLOCK, ROW_RECORD_TYPE,
};
use crate::{Error, Result};

fn invalid(record_type: u16, message: impl Into<String>) -> Error {
    Error::InvalidRecord {
        record_type,
        message: message.into(),
    }
}

fn read_u16(data: &[u8], offset: usize, record_type: u16, field: &str) -> Result<u16> {
    data.get(offset..offset + 2)
        .map(|bytes| u16::from_le_bytes([bytes[0], bytes[1]]))
        .ok_or_else(|| invalid(record_type, format!("truncated {field}")))
}

fn read_u32(data: &[u8], offset: usize, record_type: u16, field: &str) -> Result<u32> {
    data.get(offset..offset + 4)
        .map(|bytes| u32::from_le_bytes(bytes.try_into().unwrap()))
        .ok_or_else(|| invalid(record_type, format!("truncated {field}")))
}

fn with_record_header(record_type: u16, payload: Vec<u8>) -> Result<Vec<u8>> {
    let length = u16::try_from(payload.len())
        .map_err(|_error| invalid(record_type, "payload length exceeds BIFF u16"))?;
    let mut record = Vec::with_capacity(4 + payload.len());
    record.extend_from_slice(&record_type.to_le_bytes());
    record.extend_from_slice(&length.to_le_bytes());
    record.extend_from_slice(&payload);
    Ok(record)
}

impl WorksheetIndexRecord {
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn parse_payload(data: &[u8], workbook_stream_len: u64) -> Result<Self> {
        if !(INDEX_FIXED_LEN..=MAX_INDEX_PAYLOAD_LEN).contains(&data.len()) {
            return Err(invalid(
                INDEX_RECORD_TYPE,
                format!(
                    "INDEX payload must be 16..={MAX_INDEX_PAYLOAD_LEN} bytes, got {}",
                    data.len()
                ),
            ));
        }
        if !(data.len() - INDEX_FIXED_LEN).is_multiple_of(4) {
            return Err(invalid(
                INDEX_RECORD_TYPE,
                "INDEX has a partial DBCell pointer",
            ));
        }
        if read_u32(data, 0, INDEX_RECORD_TYPE, "INDEX.reserved")? != 0 {
            return Err(invalid(
                INDEX_RECORD_TYPE,
                "INDEX reserved field must be zero",
            ));
        }
        let first_data_row = read_u32(data, 4, INDEX_RECORD_TYPE, "INDEX.rwMic")?;
        let last_data_row_exclusive = read_u32(data, 8, INDEX_RECORD_TYPE, "INDEX.rwMac")?;
        if first_data_row > 65_535 || last_data_row_exclusive > 65_536 {
            return Err(invalid(
                INDEX_RECORD_TYPE,
                "INDEX row bounds exceed BIFF8 limits",
            ));
        }
        if last_data_row_exclusive == 0 && first_data_row != 0 {
            return Err(invalid(
                INDEX_RECORD_TYPE,
                "empty INDEX row bounds must both be zero",
            ));
        }
        if last_data_row_exclusive != 0 && last_data_row_exclusive <= first_data_row {
            return Err(invalid(
                INDEX_RECORD_TYPE,
                "INDEX.rwMac must be greater than rwMic",
            ));
        }
        let default_column_width_position = read_u32(data, 12, INDEX_RECORD_TYPE, "INDEX.ibXF")?;
        if u64::from(default_column_width_position) >= workbook_stream_len {
            return Err(invalid(
                INDEX_RECORD_TYPE,
                "INDEX.ibXF is outside the workbook stream",
            ));
        }
        let count = (data.len() - INDEX_FIXED_LEN) / 4;
        if count > MAX_ROW_BLOCKS {
            return Err(invalid(
                INDEX_RECORD_TYPE,
                "INDEX contains more than 2048 row blocks",
            ));
        }
        if last_data_row_exclusive == 0 && count != 0 {
            return Err(invalid(
                INDEX_RECORD_TYPE,
                "empty INDEX must not reference DBCELL records",
            ));
        }
        let mut dbcell_positions = Vec::with_capacity(count);
        for offset in (INDEX_FIXED_LEN..data.len()).step_by(4) {
            let position = read_u32(data, offset, INDEX_RECORD_TYPE, "INDEX.rgibRw")?;
            if u64::from(position) >= workbook_stream_len {
                return Err(invalid(
                    INDEX_RECORD_TYPE,
                    "INDEX DBCell pointer is outside the workbook stream",
                ));
            }
            if dbcell_positions
                .last()
                .is_some_and(|previous| position <= *previous)
            {
                return Err(invalid(
                    INDEX_RECORD_TYPE,
                    "INDEX DBCell pointers must be strictly increasing",
                ));
            }
            dbcell_positions.push(position);
        }
        Ok(Self {
            first_data_row,
            last_data_row_exclusive,
            default_column_width_position,
            dbcell_positions,
        })
    }

    fn to_payload(&self) -> Vec<u8> {
        let mut data = Vec::with_capacity(INDEX_FIXED_LEN + self.dbcell_positions.len() * 4);
        data.extend_from_slice(&0u32.to_le_bytes());
        data.extend_from_slice(&self.first_data_row.to_le_bytes());
        data.extend_from_slice(&self.last_data_row_exclusive.to_le_bytes());
        data.extend_from_slice(&self.default_column_width_position.to_le_bytes());
        for position in &self.dbcell_positions {
            data.extend_from_slice(&position.to_le_bytes());
        }
        data
    }
}

impl DbCellRecord {
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn parse_payload(
        record_position: u32,
        workbook_stream_len: u64,
        data: &[u8],
    ) -> Result<Self> {
        if !(4..=MAX_DBCELL_PAYLOAD_LEN).contains(&data.len()) {
            return Err(invalid(
                DBCELL_RECORD_TYPE,
                format!(
                    "DBCELL payload must be 4..={MAX_DBCELL_PAYLOAD_LEN} bytes, got {}",
                    data.len()
                ),
            ));
        }
        if !(data.len() - 4).is_multiple_of(2) {
            return Err(invalid(
                DBCELL_RECORD_TYPE,
                "DBCELL has a partial rgdb offset",
            ));
        }
        if u64::from(record_position) >= workbook_stream_len {
            return Err(invalid(
                DBCELL_RECORD_TYPE,
                "DBCELL position is outside the workbook stream",
            ));
        }
        let row_offset = read_u32(data, 0, DBCELL_RECORD_TYPE, "DBCELL.dbRtrw")?;
        let first_row_position = if row_offset == 0 {
            None
        } else {
            Some(record_position.checked_sub(row_offset).ok_or_else(|| {
                invalid(
                    DBCELL_RECORD_TYPE,
                    "DBCELL.dbRtrw underflows the workbook stream",
                )
            })?)
        };
        let mut cell_offsets = Vec::with_capacity((data.len() - 4) / 2);
        for offset in (4..data.len()).step_by(2) {
            cell_offsets.push(read_u16(data, offset, DBCELL_RECORD_TYPE, "DBCELL.rgdb")?);
        }
        Ok(Self {
            record_position,
            first_row_position,
            cell_offsets,
        })
    }

    /// Resolve the chained `rgdb` offsets from the end of the first ROW record.
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn resolve_cell_positions(&self, first_row_end_position: u32) -> Result<Vec<u32>> {
        if !self.cell_offsets.is_empty() && self.first_row_position.is_none() {
            return Err(invalid(
                DBCELL_RECORD_TYPE,
                "DBCELL has cell offsets but dbRtrw is zero",
            ));
        }
        let mut positions = Vec::with_capacity(self.cell_offsets.len());
        let mut base = first_row_end_position;
        for offset in &self.cell_offsets {
            let position = base
                .checked_add(u32::from(*offset))
                .ok_or_else(|| invalid(DBCELL_RECORD_TYPE, "DBCELL cell offset overflows"))?;
            if position >= self.record_position {
                return Err(invalid(
                    DBCELL_RECORD_TYPE,
                    "DBCELL cell pointer does not precede DBCELL",
                ));
            }
            positions.push(position);
            base = position;
        }
        Ok(positions)
    }

    fn to_payload(&self) -> Vec<u8> {
        let row_offset = self
            .first_row_position
            .map_or(0, |position| self.record_position - position);
        let mut data = Vec::with_capacity(4 + self.cell_offsets.len() * 2);
        data.extend_from_slice(&row_offset.to_le_bytes());
        for offset in &self.cell_offsets {
            data.extend_from_slice(&offset.to_le_bytes());
        }
        data
    }
}

impl RowBlock {
    /// Reproduce this cross-validated `DBCELL` record at its original position.
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn to_record_bytes(&self) -> Result<Vec<u8>> {
        with_record_header(DBCELL_RECORD_TYPE, self.dbcell.to_payload())
    }
}

impl RowBlockIndex {
    /// Reproduce this cross-validated `INDEX` record at its original position.
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn to_index_record_bytes(&self) -> Result<Vec<u8>> {
        with_record_header(INDEX_RECORD_TYPE, self.index.to_payload())
    }
}

#[derive(Debug, Clone, Copy)]
struct ObservedRow {
    row: u16,
    position: u32,
    end_position: u32,
}

pub(crate) struct RowBlockIndexCollector {
    stream_len: u64,
    sheet_start: Option<u32>,
    sheet_end: Option<u32>,
    pub(super) index: Option<(u32, WorksheetIndexRecord)>,
    default_column_width_positions: Vec<u32>,
    pending_rows: Vec<ObservedRow>,
    pending_first_cells: BTreeMap<u16, u32>,
    blocks: Vec<RowBlock>,
    error: Option<Error>,
}

impl RowBlockIndexCollector {
    pub(crate) fn new(stream_len: u64, sheet_start: u64) -> Self {
        let start = u32::try_from(sheet_start).ok();
        Self {
            stream_len,
            sheet_start: start,
            sheet_end: None,
            index: None,
            default_column_width_positions: Vec::new(),
            pending_rows: Vec::new(),
            pending_first_cells: BTreeMap::new(),
            blocks: Vec::new(),
            error: start.is_none().then(|| {
                invalid(
                    INDEX_RECORD_TYPE,
                    "worksheet start exceeds BIFF FilePointer range",
                )
            }),
        }
    }

    pub(crate) fn feed_record(&mut self, record_position: u64, record_type: u16, data: &[u8]) {
        if self.error.is_some() {
            return;
        }
        if let Err(error) = self.feed_record_checked(record_position, record_type, data) {
            self.error = Some(error);
        }
    }

    fn feed_record_checked(
        &mut self,
        record_position: u64,
        record_type: u16,
        data: &[u8],
    ) -> Result<()> {
        let position = u32::try_from(record_position).map_err(|_error| {
            invalid(
                record_type,
                "record position exceeds BIFF FilePointer range",
            )
        })?;
        match record_type {
            BOF_RECORD_TYPE => self.sheet_start = Some(position),
            EOF_RECORD_TYPE => {
                let size = u32::try_from(4usize + data.len())
                    .map_err(|_error| invalid(record_type, "EOF size overflows"))?;
                self.sheet_end = position.checked_add(size);
            },
            INDEX_RECORD_TYPE => {
                if self.index.is_some() {
                    return Err(invalid(record_type, "duplicate worksheet INDEX record"));
                }
                self.index = Some((
                    position,
                    WorksheetIndexRecord::parse_payload(data, self.stream_len)?,
                ));
            },
            DEF_COL_WIDTH_RECORD_TYPE => self.default_column_width_positions.push(position),
            ROW_RECORD_TYPE => {
                if data.len() != 16 {
                    return Err(invalid(
                        record_type,
                        "ROW payload must be exactly 16 bytes for index validation",
                    ));
                }
                let row = read_u16(data, 0, record_type, "ROW.rw")?;
                if self
                    .pending_rows
                    .last()
                    .is_some_and(|previous| row <= previous.row)
                {
                    return Err(invalid(
                        record_type,
                        "ROW records in a row block must be strictly ordered",
                    ));
                }
                let end_position = position
                    .checked_add(20)
                    .ok_or_else(|| invalid(record_type, "ROW end position overflows"))?;
                self.pending_rows.push(ObservedRow {
                    row,
                    position,
                    end_position,
                });
            },
            DBCELL_RECORD_TYPE => self.finish_block(position, data)?,
            _ if is_cell_record(record_type) => {
                if data.len() < 2 {
                    return Err(invalid(
                        record_type,
                        "cell record is truncated before row index",
                    ));
                }
                let row = read_u16(data, 0, record_type, "cell row")?;
                self.pending_first_cells.entry(row).or_insert(position);
            },
            _ => {},
        }
        Ok(())
    }

    fn finish_block(&mut self, position: u32, data: &[u8]) -> Result<()> {
        let dbcell = DbCellRecord::parse_payload(position, self.stream_len, data)?;
        if self.pending_rows.is_empty() {
            if data.len() != 4
                || dbcell.first_row_position.is_some()
                || !dbcell.cell_offsets.is_empty()
                || !self.pending_first_cells.is_empty()
            {
                return Err(invalid(
                    DBCELL_RECORD_TYPE,
                    "DBCELL has no preceding ROW records",
                ));
            }
            let (_, index) = self.index.as_ref().ok_or_else(|| {
                invalid(
                    DBCELL_RECORD_TYPE,
                    "empty DBCELL block has no worksheet INDEX",
                )
            })?;
            let block_offset = u32::try_from(self.blocks.len())
                .ok()
                .and_then(|value| {
                    value.checked_mul(crate::utils::truncate_usize_to_u32(MAX_ROWS_PER_BLOCK))
                })
                .ok_or_else(|| invalid(DBCELL_RECORD_TYPE, "empty DBCELL block range overflow"))?;
            let first_row = index
                .first_data_row()
                .checked_add(block_offset)
                .ok_or_else(|| invalid(DBCELL_RECORD_TYPE, "empty DBCELL block range overflow"))?;
            let last_row_exclusive = first_row
                .checked_add(crate::utils::truncate_usize_to_u32(MAX_ROWS_PER_BLOCK))
                .map(|value| value.min(index.last_data_row_exclusive()))
                .ok_or_else(|| invalid(DBCELL_RECORD_TYPE, "empty DBCELL block range overflow"))?;
            if first_row >= last_row_exclusive {
                return Err(invalid(
                    DBCELL_RECORD_TYPE,
                    "empty DBCELL lies outside INDEX row bounds",
                ));
            }
            self.blocks.push(RowBlock {
                first_row: u16::try_from(first_row).map_err(|_error| {
                    invalid(
                        DBCELL_RECORD_TYPE,
                        "empty DBCELL first row exceeds BIFF8 limit",
                    )
                })?,
                last_row: u16::try_from(last_row_exclusive - 1).map_err(|_error| {
                    invalid(
                        DBCELL_RECORD_TYPE,
                        "empty DBCELL last row exceeds BIFF8 limit",
                    )
                })?,
                dbcell,
                indexed_rows: Vec::new(),
            });
            return Ok(());
        }
        if self.pending_rows.len() > MAX_ROWS_PER_BLOCK {
            return Err(invalid(
                DBCELL_RECORD_TYPE,
                "DBCELL row block exceeds 32 rows",
            ));
        }
        let row_positions = self
            .pending_rows
            .iter()
            .map(|row| (row.row, row.position))
            .collect::<BTreeMap<_, _>>();
        for row in self.pending_first_cells.keys() {
            if !row_positions.contains_key(row) {
                return Err(invalid(
                    DBCELL_RECORD_TYPE,
                    "DBCELL cell row has no ROW record in its block",
                ));
            }
        }

        let expected_cells = self
            .pending_rows
            .iter()
            .filter_map(|row| {
                self.pending_first_cells
                    .get(&row.row)
                    .map(|position| (row, *position))
            })
            .collect::<Vec<_>>();
        if expected_cells.is_empty() {
            if dbcell.first_row_position.is_some() || !dbcell.cell_offsets.is_empty() {
                return Err(invalid(
                    DBCELL_RECORD_TYPE,
                    "empty DBCELL block must have zero dbRtrw and no rgdb",
                ));
            }
        } else {
            if dbcell.first_row_position != Some(self.pending_rows[0].position) {
                return Err(invalid(
                    DBCELL_RECORD_TYPE,
                    "DBCELL.dbRtrw does not reference the first ROW record",
                ));
            }
            if dbcell.cell_offsets.len() != expected_cells.len() {
                return Err(invalid(
                    DBCELL_RECORD_TYPE,
                    "DBCELL rgdb count does not match rows containing cells",
                ));
            }
            let resolved = dbcell.resolve_cell_positions(self.pending_rows[0].end_position)?;
            let expected_positions = expected_cells
                .iter()
                .map(|(_, position)| *position)
                .collect::<Vec<_>>();
            if resolved != expected_positions {
                return Err(invalid(
                    DBCELL_RECORD_TYPE,
                    "DBCELL chained cell offsets do not match record boundaries",
                ));
            }
        }

        let indexed_rows = expected_cells
            .iter()
            .map(|(row, first_cell_position)| IndexedRow {
                row: row.row,
                row_record_position: row.position,
                first_cell_position: *first_cell_position,
            })
            .collect();
        self.blocks.push(RowBlock {
            first_row: self.pending_rows[0].row,
            last_row: self.pending_rows.last().unwrap().row,
            dbcell,
            indexed_rows,
        });
        self.pending_rows.clear();
        self.pending_first_cells.clear();
        Ok(())
    }

    pub(crate) fn finish(self) -> Result<Option<RowBlockIndex>> {
        if let Some(error) = self.error {
            return Err(error);
        }
        let Some((index_record_position, index)) = self.index else {
            if self.blocks.is_empty() {
                return Ok(None);
            }
            return Err(invalid(
                DBCELL_RECORD_TYPE,
                "DBCELL records exist without worksheet INDEX",
            ));
        };
        if !self.pending_rows.is_empty() || !self.pending_first_cells.is_empty() {
            return Err(invalid(
                INDEX_RECORD_TYPE,
                "INDEX row table ends without a DBCELL record",
            ));
        }
        let sheet_start = self
            .sheet_start
            .ok_or_else(|| invalid(INDEX_RECORD_TYPE, "worksheet INDEX has no BOF boundary"))?;
        let sheet_end = self
            .sheet_end
            .ok_or_else(|| invalid(INDEX_RECORD_TYPE, "worksheet INDEX has no EOF boundary"))?;
        if !(sheet_start..sheet_end).contains(&index_record_position) {
            return Err(invalid(
                INDEX_RECORD_TYPE,
                "INDEX record is outside worksheet bounds",
            ));
        }
        if !self
            .default_column_width_positions
            .contains(&index.default_column_width_position)
        {
            return Err(invalid(
                INDEX_RECORD_TYPE,
                "INDEX.ibXF does not reference DEFCOLWIDTH",
            ));
        }
        let actual_dbcell_positions = self
            .blocks
            .iter()
            .map(|block| block.dbcell.record_position)
            .collect::<Vec<_>>();
        if index.dbcell_positions != actual_dbcell_positions {
            return Err(invalid(
                INDEX_RECORD_TYPE,
                "INDEX DBCell pointers do not match observed DBCELL records",
            ));
        }
        if actual_dbcell_positions
            .iter()
            .any(|position| !(*position >= sheet_start && *position < sheet_end))
        {
            return Err(invalid(
                INDEX_RECORD_TYPE,
                "INDEX DBCell pointer is outside worksheet bounds",
            ));
        }

        let first_data_row = self
            .blocks
            .iter()
            .flat_map(|block| block.indexed_rows.iter())
            .map(|row| u32::from(row.row))
            .min();
        let last_data_row = self
            .blocks
            .iter()
            .flat_map(|block| block.indexed_rows.iter())
            .map(|row| u32::from(row.row))
            .max();
        if let (Some(first), Some(last)) = (first_data_row, last_data_row)
            && (index.first_data_row, index.last_data_row_exclusive) != (first, last + 1)
        {
            return Err(invalid(
                INDEX_RECORD_TYPE,
                "INDEX row bounds do not match indexed cell rows",
            ));
        }
        Ok(Some(RowBlockIndex {
            index_record_position,
            index,
            blocks: self.blocks,
        }))
    }
}

fn is_cell_record(record_type: u16) -> bool {
    matches!(
        record_type,
        0x0006 // Formula
            | 0x00BD // MulRk
            | 0x00BE // MulBlank
            | 0x00D6 // RString
            | 0x00FD // LabelSst
            | 0x0201 // Blank
            | 0x0203 // Number
            | 0x0204 // Label
            | 0x0205 // BoolErr
            | 0x027E // RK
    )
}
