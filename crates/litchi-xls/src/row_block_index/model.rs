//! Semantic BIFF8 worksheet row-block values.

/// Parsed worksheet `INDEX` record.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct WorksheetIndexRecord {
    pub(super) first_data_row: u32,
    pub(super) last_data_row_exclusive: u32,
    pub(super) default_column_width_position: u32,
    pub(super) dbcell_positions: Vec<u32>,
}

impl WorksheetIndexRecord {
    #[must_use]
    pub fn first_data_row(&self) -> u32 {
        self.first_data_row
    }

    #[must_use]
    pub fn last_data_row_exclusive(&self) -> u32 {
        self.last_data_row_exclusive
    }

    #[must_use]
    pub fn default_column_width_position(&self) -> u32 {
        self.default_column_width_position
    }

    #[must_use]
    pub fn dbcell_positions(&self) -> &[u32] {
        &self.dbcell_positions
    }
}

/// Parsed worksheet `DBCELL` record with its absolute stream position.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DbCellRecord {
    pub(super) record_position: u32,
    pub(super) first_row_position: Option<u32>,
    pub(super) cell_offsets: Vec<u16>,
}

impl DbCellRecord {
    #[must_use]
    pub fn record_position(&self) -> u32 {
        self.record_position
    }

    #[must_use]
    pub fn first_row_position(&self) -> Option<u32> {
        self.first_row_position
    }

    #[must_use]
    pub fn cell_offsets(&self) -> &[u16] {
        &self.cell_offsets
    }
}

/// First cell-record pointer for one indexed worksheet row.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct IndexedRow {
    pub(super) row: u16,
    pub(super) row_record_position: u32,
    pub(super) first_cell_position: u32,
}

impl IndexedRow {
    #[must_use]
    pub fn row(self) -> u16 {
        self.row
    }

    #[must_use]
    pub fn row_record_position(self) -> u32 {
        self.row_record_position
    }

    #[must_use]
    pub fn first_cell_position(self) -> u32 {
        self.first_cell_position
    }
}

/// One validated row block and its `DBCELL` pointer chain.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RowBlock {
    pub(super) first_row: u16,
    pub(super) last_row: u16,
    pub(super) dbcell: DbCellRecord,
    pub(super) indexed_rows: Vec<IndexedRow>,
}

impl RowBlock {
    #[must_use]
    pub fn first_row(&self) -> u16 {
        self.first_row
    }

    #[must_use]
    pub fn last_row(&self) -> u16 {
        self.last_row
    }

    #[must_use]
    pub fn dbcell(&self) -> &DbCellRecord {
        &self.dbcell
    }

    #[must_use]
    pub fn indexed_rows(&self) -> &[IndexedRow] {
        &self.indexed_rows
    }
}

/// Cross-validated worksheet `INDEX` plus all referenced row blocks.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RowBlockIndex {
    pub(super) index_record_position: u32,
    pub(super) index: WorksheetIndexRecord,
    pub(super) blocks: Vec<RowBlock>,
}

impl RowBlockIndex {
    #[must_use]
    pub fn index_record_position(&self) -> u32 {
        self.index_record_position
    }

    #[must_use]
    pub fn index_record(&self) -> &WorksheetIndexRecord {
        &self.index
    }

    #[must_use]
    pub fn blocks(&self) -> &[RowBlock] {
        &self.blocks
    }

    #[must_use]
    pub fn block_for_row(&self, row: u32) -> Option<&RowBlock> {
        let row = u16::try_from(row).ok()?;
        self.blocks
            .iter()
            .find(|block| block.first_row <= row && row <= block.last_row)
    }

    #[must_use]
    pub fn first_cell_position(&self, row: u32) -> Option<u32> {
        let row = u16::try_from(row).ok()?;
        self.block_for_row(u32::from(row))?
            .indexed_rows
            .iter()
            .find(|entry| entry.row == row)
            .map(|entry| entry.first_cell_position)
    }
}
