//! Typed, source-bound access to an XLSB worksheet binary index.
//!
//! A binary index is only a directory of offsets.  The worksheet stream is
//! authoritative: [`BinaryIndex::bind`] checks every indexed offset
//! against actual row and cell record boundaries before exposing a lookup.
//! This keeps a corrupt or stale index from turning an arbitrary payload byte
//! into a cell value.

use std::collections::{BTreeMap, BTreeSet};
use std::fmt;
use std::sync::Arc;

use crate::package::error::{Error, Result};
use crate::raw::{self, kind};

/// The maximum row index in an XLSB worksheet.
pub const MAX_ROW_INDEX: u32 = 1_048_575;
/// The maximum column index in an XLSB worksheet.
pub const MAX_COLUMN_INDEX: u32 = 16_383;
/// The number of rows covered by one `BrtIndexBlock`.
pub const INDEX_BLOCK_ROWS: u32 = 32;
/// The number of columns in one row column block.
pub const INDEX_COLUMN_BLOCK_WIDTH: u32 = 1_024;
const INDEX_COLUMN_BLOCKS: u32 = 16;
const INDEX_BLOCK_UNUSED_BYTES: usize = 8;
const MAX_WORKSHEET_SOURCE_BYTES_DEFAULT: usize = 512 * 1024 * 1024;
const MAX_INDEX_RECORDS_DEFAULT: usize = 1_000_000;
const MAX_INDEX_BLOCKS_DEFAULT: usize = 32_768;
const MAX_INDEX_CELLS_DEFAULT: usize = 1_000_000;
pub(crate) const MAX_INDEX_STREAM_BYTES: usize = 64 * 1024 * 1024;

pub(crate) const BINARY_INDEX_RELATIONSHIP: &str =
    "http://schemas.microsoft.com/office/2006/relationships/xlBinaryIndex";
pub(crate) const BINARY_INDEX_CONTENT_TYPE: &str = "application/vnd.ms-excel.binIndexWs";

/// Finite resource limits for one binary-index operation.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Limits {
    /// Raw BIFF12 payload and string ceilings used while framing the streams.
    pub raw: raw::Limits,
    /// Maximum worksheet-part bytes scanned while binding or generating an index.
    pub source_bytes: usize,
    /// Maximum number of records accepted in either stream.
    pub max_records: usize,
    /// Maximum number of index blocks accepted.
    pub max_blocks: usize,
    /// Maximum number of worksheet cell anchors retained by a binding scan.
    pub max_cells: usize,
}

impl Limits {
    /// Safe limits for ordinary XLSB worksheets.
    pub const DEFAULT: Self = Self {
        raw: raw::Limits::DEFAULT,
        source_bytes: MAX_WORKSHEET_SOURCE_BYTES_DEFAULT,
        max_records: MAX_INDEX_RECORDS_DEFAULT,
        max_blocks: MAX_INDEX_BLOCKS_DEFAULT,
        max_cells: MAX_INDEX_CELLS_DEFAULT,
    };

    /// Construct explicit index limits.
    #[must_use]
    pub const fn new(
        raw: raw::Limits,
        max_records: usize,
        max_blocks: usize,
        max_cells: usize,
    ) -> Self {
        Self {
            raw,
            source_bytes: MAX_WORKSHEET_SOURCE_BYTES_DEFAULT,
            max_records,
            max_blocks,
            max_cells,
        }
    }

    /// Construct explicit limits including the worksheet-part byte ceiling.
    #[must_use]
    pub const fn new_with_source_bytes(
        raw: raw::Limits,
        source_bytes: usize,
        max_records: usize,
        max_blocks: usize,
        max_cells: usize,
    ) -> Self {
        Self {
            raw,
            source_bytes,
            max_records,
            max_blocks,
            max_cells,
        }
    }

    /// Construct the publication profile used by the worksheet cell owner.
    pub(crate) const fn publication_default() -> Self {
        // Every BIFF12 record has at least one kind byte and one size byte.
        // This explicit record ceiling admits every framing sequence within
        // the existing worksheet owner's source-byte ceiling.
        const MIN_RECORD_BYTES: usize = 2;
        let worksheet = crate::cell_values::Limits::DEFAULT;
        Self::new_with_source_bytes(
            worksheet.raw(),
            worksheet.source_bytes(),
            worksheet.source_bytes() / MIN_RECORD_BYTES,
            Self::DEFAULT.max_blocks,
            worksheet.cells(),
        )
    }

    pub(crate) fn validate(self) -> Result<Self> {
        self.raw.validate()?;
        if self.max_records == 0 {
            return Err(invalid("max_records must be nonzero"));
        }
        if self.max_blocks == 0 {
            return Err(invalid("max_blocks must be nonzero"));
        }
        if self.max_cells == 0 {
            return Err(invalid("max_cells must be nonzero"));
        }
        if self.source_bytes == 0 {
            return Err(invalid("source_bytes must be nonzero"));
        }
        Ok(self)
    }
}

impl Default for Limits {
    fn default() -> Self {
        Self::DEFAULT
    }
}

/// Descriptive alias for callers that use several XLSB limit profiles.
pub type BinaryIndexLimits = Limits;

/// A first-cell anchor in the worksheet stream.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord)]
pub struct CellOffset {
    /// Zero-based row containing the cell.
    pub row: u32,
    /// Zero-based column containing the cell.
    pub column: u32,
    /// Byte offset of the complete cell record in the worksheet part.
    pub offset: u64,
    /// BIFF12 record kind of the cell record.
    pub record_kind: u16,
}

/// One indexed row/column block.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ColumnBlock {
    /// Zero-based row covered by this anchor.
    pub row: u32,
    /// Zero-based 1024-column block covered by this anchor.
    pub column_block: u8,
    /// Absolute worksheet-part offset of the first cell in the block.
    pub offset: u64,
}

/// The typed contents of one `BrtIndexRowBlock`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RowBlock {
    /// Bit mask of rows in the containing block that have indexed cells.
    pub row_mask: u32,
    /// Absolute worksheet-part offset of the first indexed cell.
    pub base_offset: u64,
    /// First-cell anchors in row-major order.
    pub columns: Vec<ColumnBlock>,
}

/// One row range and its immediately following row-block directory.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct IndexBlock {
    /// Zero-based inclusive first row.
    pub row_start: u32,
    /// One-based exclusive last row.
    pub row_end: u32,
    /// The sparse row/column offsets for this range.
    pub rows: RowBlock,
}

/// A parsed worksheet binary index independent of a worksheet stream.
#[derive(Debug, Clone)]
pub struct BinaryIndex {
    source: Arc<Vec<u8>>,
    blocks: Vec<IndexBlock>,
}

impl BinaryIndex {
    /// Parse a bounded binary-index record stream.
    pub fn parse(bytes: &[u8], limits: Limits) -> Result<Self> {
        let limits = limits.validate()?;
        if bytes.len() > MAX_INDEX_STREAM_BYTES {
            return Err(limit(
                "binary-index bytes",
                bytes.len(),
                MAX_INDEX_STREAM_BYTES,
            ));
        }
        if bytes.len() > limits.source_bytes {
            return Err(limit(
                "binary-index bytes",
                bytes.len(),
                limits.source_bytes,
            ));
        }
        let source = Arc::new(copy_bytes(bytes, "binary-index bytes")?);
        Self::parse_shared(source, limits)
    }

    /// Bind the parsed index to one worksheet stream.
    pub fn bind(&self, worksheet: &[u8], limits: Limits) -> Result<WorksheetBinaryIndex> {
        let limits = limits.validate()?;
        let scan = scan_worksheet(worksheet, limits)?;
        bind_anchors(self.clone(), scan, worksheet.len(), limits, true)
    }

    /// Return the parsed row ranges in source order.
    #[must_use]
    pub fn blocks(&self) -> &[IndexBlock] {
        &self.blocks
    }

    /// Return the exact source bytes retained by this parsed index.
    #[must_use]
    pub fn source_bytes(&self) -> &[u8] {
        self.source.as_slice()
    }

    fn parse_shared(source: Arc<Vec<u8>>, limits: Limits) -> Result<Self> {
        let mut blocks = Vec::new();
        blocks
            .try_reserve(1)
            .map_err(|source| allocation("binary-index blocks", source))?;
        let mut pending = None;
        let mut ended = false;
        let mut end_count = 0usize;
        let mut record_count = 0usize;
        let mut entry_count = 0usize;
        let mut previous_start = None;
        let mut previous_end = None;

        for record_result in raw::Records::try_with_limits(source.as_slice(), limits.raw)? {
            let record = record_result?;
            record_count = record_count.checked_add(1).ok_or(Error::CapacityOverflow {
                resource: "binary-index record count",
            })?;
            if record_count > limits.max_records {
                return Err(limit(
                    "binary-index records",
                    record_count,
                    limits.max_records,
                ));
            }
            if ended && record.kind() != kind::INDEX_PART_END {
                return Err(invalid("record follows BrtIndexPartEnd"));
            }

            match record.kind() {
                kind::INDEX_BLOCK => {
                    if let Some((row_start, row_end)) = pending.take() {
                        push_index_block(
                            &mut blocks,
                            limits,
                            IndexBlock {
                                row_start,
                                row_end,
                                rows: empty_row_block(),
                            },
                        )?;
                    }
                    let (row_start, row_end) = parse_index_block(record.payload())?;
                    if let Some(previous) = previous_end
                        && row_start < previous
                    {
                        return Err(invalid("index row ranges overlap"));
                    }
                    if let Some(previous) = previous_start
                        && row_start <= previous
                    {
                        return Err(invalid("index row ranges are not strictly increasing"));
                    }
                    previous_start = Some(row_start);
                    previous_end = Some(row_end);
                    pending = Some((row_start, row_end));
                },
                kind::INDEX_ROW_BLOCK => {
                    let Some((row_start, row_end)) = pending.take() else {
                        return Err(invalid("BrtIndexRowBlock has no preceding BrtIndexBlock"));
                    };
                    let rows =
                        parse_row_block(record.payload(), row_start, row_end, limits.max_cells)?;
                    entry_count = entry_count.checked_add(rows.columns.len()).ok_or(
                        Error::CapacityOverflow {
                            resource: "binary-index entry count",
                        },
                    )?;
                    if entry_count > limits.max_cells {
                        return Err(limit("binary-index entries", entry_count, limits.max_cells));
                    }
                    push_index_block(
                        &mut blocks,
                        limits,
                        IndexBlock {
                            row_start,
                            row_end,
                            rows,
                        },
                    )?;
                },
                kind::INDEX_PART_END => {
                    if let Some((row_start, row_end)) = pending.take() {
                        push_index_block(
                            &mut blocks,
                            limits,
                            IndexBlock {
                                row_start,
                                row_end,
                                rows: empty_row_block(),
                            },
                        )?;
                    }
                    if !record.payload().is_empty() {
                        return Err(invalid("BrtIndexPartEnd has a nonempty payload"));
                    }
                    end_count = end_count.checked_add(1).ok_or(Error::CapacityOverflow {
                        resource: "binary-index end-marker count",
                    })?;
                    if end_count > 2 {
                        return Err(invalid(
                            "binary index has more than two BrtIndexPartEnd records",
                        ));
                    }
                    ended = true;
                },
                _ => {
                    return Err(invalid("binary index contains a record outside SHEETINDEX"));
                },
            }
        }
        if blocks.is_empty() {
            return Err(invalid("binary index is missing BrtIndexBlock"));
        }
        if !ended {
            return Err(invalid("binary index is missing BrtIndexPartEnd"));
        }

        Ok(Self { source, blocks })
    }
}

fn empty_row_block() -> RowBlock {
    RowBlock {
        row_mask: 0,
        base_offset: 0,
        columns: Vec::new(),
    }
}

fn push_index_block(blocks: &mut Vec<IndexBlock>, limits: Limits, block: IndexBlock) -> Result<()> {
    if blocks.len() >= limits.max_blocks {
        return Err(limit(
            "binary-index blocks",
            blocks.len() + 1,
            limits.max_blocks,
        ));
    }
    blocks
        .try_reserve(1)
        .map_err(|source| allocation("binary-index blocks", source))?;
    blocks.push(block);
    Ok(())
}

/// A typed binary index proven against one worksheet's row and cell framing.
#[derive(Debug, Clone)]
pub struct WorksheetBinaryIndex {
    index: BinaryIndex,
    /// One verified first-cell anchor per nonempty row/column block.
    buckets: Vec<CellOffset>,
    worksheet_len: usize,
    limits: Limits,
}

#[derive(Debug)]
struct WorksheetScan {
    anchors: Vec<CellOffset>,
    row_headers: Vec<u32>,
}

impl WorksheetBinaryIndex {
    /// Parse and bind a binary index in one bounded operation.
    pub fn from_parts(index: &[u8], worksheet: &[u8], limits: Limits) -> Result<Self> {
        let parsed = BinaryIndex::parse(index, limits)?;
        parsed.bind(worksheet, limits)
    }

    /// Parse and bind using safe defaults.
    pub fn from_parts_default(index: &[u8], worksheet: &[u8]) -> Result<Self> {
        Self::from_parts(index, worksheet, Limits::DEFAULT)
    }

    /// Return the validated standalone index.
    #[must_use]
    pub fn index(&self) -> &BinaryIndex {
        &self.index
    }

    /// Return the actual worksheet-part length used during binding.
    #[must_use]
    pub fn worksheet_len(&self) -> usize {
        self.worksheet_len
    }

    /// Return the validated cell payload for an internal source-backed owner.
    ///
    /// The returned slice borrows the caller's retained worksheet allocation;
    /// this method never detaches or copies an OPC-managed part.
    pub(crate) fn lookup_record<'a>(
        &self,
        worksheet: &'a [u8],
        row: u32,
        column: u32,
    ) -> Result<Option<(CellOffset, &'a [u8])>> {
        validate_coordinate(row, column)?;
        let Some(start) = self.start_offset(row, column)? else {
            return Ok(None);
        };
        self.lookup_from_offset(worksheet, start, row, column)
    }

    /// Return the first cell record in the row/column block selected by the
    /// binary index.  Callers that need a value can scan forward from this
    /// offset until the requested column is reached.
    pub fn start_offset(&self, row: u32, column: u32) -> Result<Option<CellOffset>> {
        validate_coordinate(row, column)?;
        let column_block = column / INDEX_COLUMN_BLOCK_WIDTH;
        Ok(self
            .buckets
            .binary_search_by_key(&(row, column_block), |anchor| {
                (anchor.row, anchor.column / INDEX_COLUMN_BLOCK_WIDTH)
            })
            .ok()
            .map(|index| self.buckets[index]))
    }

    /// Iterate compact worksheet anchors retained by the binding scan.
    #[must_use]
    pub fn anchors(&self) -> &[CellOffset] {
        &self.buckets
    }

    fn lookup_from_offset<'a>(
        &self,
        worksheet: &'a [u8],
        start: CellOffset,
        row: u32,
        column: u32,
    ) -> Result<Option<(CellOffset, &'a [u8])>> {
        let start_offset = usize::try_from(start.offset)
            .map_err(|_error| invalid("indexed worksheet offset exceeds platform size"))?;
        let tail = worksheet
            .get(start_offset..)
            .ok_or_else(|| invalid("indexed worksheet offset is outside the stream"))?;
        let mut current_row = row;
        let mut first = true;
        let mut record_count = 0usize;
        for record_result in raw::Records::try_with_limits(tail, self.limits.raw)? {
            let record = record_result?;
            record_count = record_count.checked_add(1).ok_or(Error::CapacityOverflow {
                resource: "binary-index lookup records",
            })?;
            if record_count > self.limits.max_records {
                return Err(limit(
                    "binary-index lookup records",
                    record_count,
                    self.limits.max_records,
                ));
            }
            if first && !is_cell_kind(record.kind()) {
                return Err(invalid("indexed anchor does not begin at a cell record"));
            }
            first = false;
            match record.kind() {
                kind::ROW_HDR => {
                    let next_row = read_u32(record.payload(), 0)?;
                    if next_row > row {
                        break;
                    }
                    current_row = next_row;
                },
                cell_kind if is_cell_kind(cell_kind) => {
                    let cell_row = current_row;
                    let cell_column = read_u32(record.payload(), 0)?;
                    if cell_row != row {
                        break;
                    }
                    if cell_column == column {
                        let relative = start_offset.checked_add(record.offset()).ok_or(
                            Error::CapacityOverflow {
                                resource: "binary-index lookup offset",
                            },
                        )?;
                        let offset = CellOffset {
                            row,
                            column,
                            offset: u64::try_from(relative)
                                .map_err(|_error| invalid("lookup offset overflows"))?,
                            record_kind: cell_kind.get(),
                        };
                        return Ok(Some((offset, record.payload())));
                    }
                    if cell_column > column {
                        break;
                    }
                },
                _ => {},
            }
        }
        Ok(None)
    }
}

/// Parse, bind, and return the exact source index for a worksheet edit.
///
/// A byte-identical worksheet is a no-op and returns the original index bytes
/// without parsing them.  When row/column geometry is unchanged, the original
/// bytes are retained exactly, including unparsed unknown or malformed index
/// bytes when row headers and bucket keys/offsets are unchanged. Strict indexed
/// lookup still rejects such an index. If only known offsets moved, the helper patches
/// those fields in place. A topology change is regenerated from the modeled
/// grammar; bytes outside that grammar are rejected before publication.
pub(crate) fn maintain_index(
    source_index: Option<&[u8]>,
    before_worksheet: &[u8],
    after_worksheet: &[u8],
    limits: Limits,
) -> Result<Option<Vec<u8>>> {
    let limits = limits.validate()?;
    if let Some(bytes) = source_index {
        let maximum = MAX_INDEX_STREAM_BYTES.min(limits.source_bytes);
        if bytes.len() > maximum {
            return Err(limit("binary-index bytes", bytes.len(), maximum));
        }
    }
    if before_worksheet == after_worksheet {
        return source_index
            .map(|bytes| copy_bytes(bytes, "binary-index bytes"))
            .transpose();
    }

    let before_scan = scan_worksheet(before_worksheet, limits)?;
    let after_scan = scan_worksheet(after_worksheet, limits)?;
    let before_geometry = geometry(&before_scan.anchors);
    let after_geometry = geometry(&after_scan.anchors);
    let row_headers_unchanged = before_scan.row_headers == after_scan.row_headers;
    let Some(source_index) = source_index else {
        return if after_scan.anchors.is_empty() && after_scan.row_headers.is_empty() {
            Ok(None)
        } else {
            encode_anchors(&after_scan, limits).map(Some)
        };
    };

    if before_geometry == after_geometry && row_headers_unchanged {
        // The same source anchors remain valid even when a value or record
        // kind changed in place. Preserve the complete index byte-for-byte
        // without interpreting extension records that are irrelevant to the
        // unchanged index geometry.
        return Ok(Some(copy_bytes(source_index, "binary-index bytes")?));
    }

    let parsed = BinaryIndex::parse(source_index, limits)?;
    let _before_bound = bind_anchors(
        parsed.clone(),
        before_scan,
        before_worksheet.len(),
        limits,
        true,
    )?;

    if row_headers_unchanged && before_geometry.keys().eq(after_geometry.keys()) {
        // The source index may use a native block partition or contain
        // reserved bytes that the canonical writer does not model.  When the
        // row list and compact cell-bucket keys are unchanged, bind that
        // already-validated source topology to the new worksheet and patch
        // only the known offsets in place.
        let after_bound = bind_anchors(parsed, after_scan, after_worksheet.len(), limits, false)?;
        return patch_offsets(source_index, &after_bound, limits).map(Some);
    }

    // A row-header or cell-bucket topology change must be regenerated from
    // the worksheet framing, so row-header-only rows and zero column masks
    // remain represented in the new index.
    Ok(Some(encode_anchors(&after_scan, limits)?))
}

/// Generate a canonical worksheet binary index from an already serialized
/// worksheet part. Empty worksheets receive one empty `BrtIndexBlock` followed
/// by `BrtIndexPartEnd`, as required by `SHEETINDEX`.
pub(crate) fn encode_for_worksheet(worksheet: &[u8], limits: Limits) -> Result<Vec<u8>> {
    let limits = limits.validate()?;
    let scan = scan_worksheet(worksheet, limits)?;
    encode_anchors(&scan, limits)
}

fn bind_anchors(
    index: BinaryIndex,
    scan: WorksheetScan,
    worksheet_len: usize,
    limits: Limits,
    verify_offsets: bool,
) -> Result<WorksheetBinaryIndex> {
    let actual = geometry(&scan.anchors);
    let mut indexed = BTreeSet::new();
    let mut indexed_rows = BTreeSet::new();
    let mut expected_offsets = BTreeMap::new();
    for block in &index.blocks {
        for row_offset in 0..INDEX_BLOCK_ROWS {
            if block.rows.row_mask & (1_u32 << row_offset) == 0 {
                continue;
            }
            let row = block
                .row_start
                .checked_add(row_offset)
                .ok_or_else(|| invalid("row-block row overflows"))?;
            if row >= block.row_end {
                return Err(invalid("row mask exceeds its index block"));
            }
            indexed_rows.insert(row);
        }
        for entry in &block.rows.columns {
            if entry.row < block.row_start {
                return Err(invalid("row-block anchor precedes its index block"));
            }
            let row_offset = entry.row - block.row_start;
            if row_offset >= INDEX_BLOCK_ROWS || block.rows.row_mask & (1_u32 << row_offset) == 0 {
                return Err(invalid("too many row-block column masks"));
            }
            let row = entry.row;
            if row >= block.row_end || u32::from(entry.column_block) >= INDEX_COLUMN_BLOCKS {
                return Err(invalid("row-block anchor lies outside its index block"));
            }
            let key = (row, u32::from(entry.column_block));
            if indexed.insert(key) {
                expected_offsets.insert(key, entry.offset);
            } else {
                return Err(invalid("duplicate row/column index anchor"));
            }
        }
        let expected_rows =
            usize::try_from(block.rows.row_mask.count_ones()).map_err(|_error| {
                Error::CapacityOverflow {
                    resource: "binary-index row mask",
                }
            })?;
        let row_masks = block
            .rows
            .columns
            .iter()
            .map(|entry| entry.row)
            .collect::<BTreeSet<_>>();
        if row_masks.len() > expected_rows {
            return Err(invalid("row-block column-mask count exceeds row mask"));
        }
        let _ = usize::try_from(block.rows.base_offset)
            .map_err(|_error| invalid("index base offset exceeds platform size"))?;
        for entry in &block.rows.columns {
            if entry.offset < block.rows.base_offset {
                return Err(invalid("index sub-offset underflows its base offset"));
            }
            let relative = entry.offset - block.rows.base_offset;
            if relative > u64::from(u32::MAX) {
                return Err(invalid("index sub-offset exceeds u32"));
            }
        }
    }
    let worksheet_size = u64::try_from(worksheet_len)
        .map_err(|_error| invalid("worksheet length overflows platform size"))?;
    if scan
        .anchors
        .iter()
        .any(|anchor| anchor.offset >= worksheet_size)
    {
        return Err(invalid("cell anchor lies outside worksheet stream"));
    }

    if verify_offsets {
        for (key, offset) in &expected_offsets {
            if actual.get(key) != Some(offset) {
                return Err(invalid(
                    "index offset is not an exact first-cell record boundary",
                ));
            }
        }
    }
    let actual_keys = actual.keys().copied().collect::<BTreeSet<_>>();
    if actual_keys != indexed {
        return Err(invalid(
            "binary index does not cover worksheet cell blocks exactly",
        ));
    }
    let actual_rows = scan.row_headers.iter().copied().collect::<BTreeSet<_>>();
    if actual_rows != indexed_rows {
        return Err(invalid(
            "binary index row mask does not cover worksheet row headers",
        ));
    }

    Ok(WorksheetBinaryIndex {
        index,
        buckets: scan.anchors,
        worksheet_len,
        limits,
    })
}

fn parse_index_block(payload: &[u8]) -> Result<(u32, u32)> {
    if payload.len() < 16 {
        return Err(invalid("BrtIndexBlock payload is truncated"));
    }
    let row_start = read_u32(payload, 0)?;
    let row_end = read_u32(payload, 4)?;
    if row_end <= row_start || row_end - row_start > INDEX_BLOCK_ROWS || row_end > MAX_ROW_INDEX + 1
    {
        return Err(invalid("BrtIndexBlock row range is invalid"));
    }
    let range =
        usize::try_from(row_end - row_start).map_err(|_error| invalid("row range overflow"))?;
    let expected_unused = ((range
        + usize::try_from(INDEX_BLOCK_ROWS).map_err(|_error| invalid("row block overflow"))?)
        / usize::try_from(INDEX_BLOCK_ROWS).map_err(|_error| invalid("row block overflow"))?)
    .checked_mul(4)
    .ok_or(Error::CapacityOverflow {
        resource: "BrtIndexBlock unused bytes",
    })?;
    let expected = 16usize
        .checked_add(expected_unused)
        .ok_or(Error::CapacityOverflow {
            resource: "BrtIndexBlock payload",
        })?;
    if payload.len() != expected {
        return Err(invalid("BrtIndexBlock unused field length is invalid"));
    }
    Ok((row_start, row_end))
}

fn parse_row_block(
    payload: &[u8],
    row_start: u32,
    row_end: u32,
    max_entries: usize,
) -> Result<RowBlock> {
    if payload.len() < 12 {
        return Err(invalid("BrtIndexRowBlock payload is truncated"));
    }
    let row_mask = read_u32(payload, 0)?;
    let base_offset = read_u64(payload, 4)?;
    let row_count =
        usize::try_from(row_mask.count_ones()).map_err(|_error| Error::CapacityOverflow {
            resource: "BrtIndexRowBlock row count",
        })?;
    let masks_len = row_count.checked_mul(2).ok_or(Error::CapacityOverflow {
        resource: "BrtIndexRowBlock column masks",
    })?;
    let masks_end = 12usize
        .checked_add(masks_len)
        .ok_or(Error::CapacityOverflow {
            resource: "BrtIndexRowBlock payload",
        })?;
    if payload.len() < masks_end {
        return Err(invalid("BrtIndexRowBlock column masks are truncated"));
    }
    let mut columns = Vec::new();
    let mut total_columns = 0usize;
    let mut mask_index = 0usize;
    for row_offset in 0..INDEX_BLOCK_ROWS {
        if row_mask & (1_u32 << row_offset) == 0 {
            continue;
        }
        let row = row_start + row_offset;
        if row >= row_end {
            return Err(invalid("BrtIndexRowBlock row mask exceeds its range"));
        }
        let row_index = mask_index;
        mask_index = mask_index.checked_add(1).ok_or(Error::CapacityOverflow {
            resource: "BrtIndexRowBlock mask index",
        })?;
        let mask_offset = 12usize
            .checked_add(row_index.checked_mul(2).ok_or(Error::CapacityOverflow {
                resource: "BrtIndexRowBlock mask offset",
            })?)
            .ok_or(Error::CapacityOverflow {
                resource: "BrtIndexRowBlock mask offset",
            })?;
        let mask = read_u16(payload, mask_offset)?;
        total_columns = total_columns
            .checked_add(
                usize::try_from(mask.count_ones())
                    .map_err(|_error| invalid("column count overflow"))?,
            )
            .ok_or(Error::CapacityOverflow {
                resource: "BrtIndexRowBlock column count",
            })?;
    }
    let offsets_len = total_columns
        .checked_mul(4)
        .ok_or(Error::CapacityOverflow {
            resource: "BrtIndexRowBlock offsets",
        })?;
    let expected = masks_end
        .checked_add(offsets_len)
        .ok_or(Error::CapacityOverflow {
            resource: "BrtIndexRowBlock payload",
        })?;
    if payload.len() != expected {
        return Err(invalid("BrtIndexRowBlock offset array length is invalid"));
    }
    if total_columns > max_entries {
        return Err(limit("binary-index entries", total_columns, max_entries));
    }
    columns
        .try_reserve_exact(total_columns)
        .map_err(|source| allocation("BrtIndexRowBlock columns", source))?;
    let mut offset_cursor = masks_end;
    mask_index = 0;
    for row_offset in 0..INDEX_BLOCK_ROWS {
        if row_mask & (1_u32 << row_offset) == 0 {
            continue;
        }
        let row = row_start + row_offset;
        let row_index = mask_index;
        mask_index = mask_index.checked_add(1).ok_or(Error::CapacityOverflow {
            resource: "BrtIndexRowBlock mask index",
        })?;
        let mask_offset = 12usize
            .checked_add(row_index.checked_mul(2).ok_or(Error::CapacityOverflow {
                resource: "BrtIndexRowBlock mask offset",
            })?)
            .ok_or(Error::CapacityOverflow {
                resource: "BrtIndexRowBlock mask offset",
            })?;
        let mask = read_u16(payload, mask_offset)?;
        for column_block in 0..16_u32 {
            if mask & (1_u16 << column_block) == 0 {
                continue;
            }
            let offset = read_u32(payload, offset_cursor)?;
            offset_cursor = offset_cursor
                .checked_add(4)
                .ok_or(Error::CapacityOverflow {
                    resource: "BrtIndexRowBlock offset cursor",
                })?;
            let offset = base_offset
                .checked_add(u64::from(offset))
                .ok_or_else(|| invalid("BrtIndexRowBlock offset overflows"))?;
            columns.push(ColumnBlock {
                row,
                column_block: u8::try_from(column_block)
                    .map_err(|_error| invalid("column block conversion failed"))?,
                offset,
            });
        }
    }
    Ok(RowBlock {
        row_mask,
        base_offset,
        columns,
    })
}

fn scan_worksheet(worksheet: &[u8], limits: Limits) -> Result<WorksheetScan> {
    if worksheet.len() > limits.source_bytes {
        return Err(limit(
            "binary-index worksheet bytes",
            worksheet.len(),
            limits.source_bytes,
        ));
    }
    let mut buckets = Vec::new();
    let mut row_headers = Vec::new();
    let mut in_sheet_data = false;
    let mut seen_sheet_data = false;
    let mut current_row = None;
    let mut previous_row = None;
    let mut previous_column = None;
    let mut cell_count = 0usize;
    let mut record_count = 0usize;
    for record_result in raw::Records::try_with_limits(worksheet, limits.raw)? {
        let record = record_result?;
        record_count = record_count.checked_add(1).ok_or(Error::CapacityOverflow {
            resource: "worksheet record count",
        })?;
        if record_count > limits.max_records {
            return Err(limit("worksheet records", record_count, limits.max_records));
        }
        match record.kind() {
            kind::BEGIN_SHEET_DATA => {
                if in_sheet_data || seen_sheet_data {
                    return Err(invalid("worksheet has duplicate or nested sheet data"));
                }
                if !record.payload().is_empty() {
                    return Err(invalid("BrtBeginSheetData has a nonempty payload"));
                }
                in_sheet_data = true;
                seen_sheet_data = true;
                current_row = None;
                previous_column = None;
            },
            kind::END_SHEET_DATA => {
                if !in_sheet_data || !record.payload().is_empty() {
                    return Err(invalid("invalid BrtEndSheetData framing"));
                }
                in_sheet_data = false;
                current_row = None;
                previous_column = None;
            },
            kind::ROW_HDR if in_sheet_data => {
                if record.payload().len() < 4 {
                    return Err(invalid("BrtRowHdr payload is truncated"));
                }
                let row = read_u32(record.payload(), 0)?;
                validate_coordinate(row, 0)?;
                if let Some(previous) = previous_row
                    && row <= previous
                {
                    return Err(invalid("worksheet rows are not strictly increasing"));
                }
                previous_row = Some(row);
                current_row = Some(row);
                previous_column = None;
                row_headers
                    .try_reserve(1)
                    .map_err(|source| allocation("worksheet index row headers", source))?;
                row_headers.push(row);
            },
            cell_kind if in_sheet_data && is_cell_kind(cell_kind) => {
                let row = current_row.ok_or_else(|| invalid("cell record has no row header"))?;
                if record.payload().len() < 8 {
                    return Err(invalid("cell record payload is truncated"));
                }
                let column = read_u32(record.payload(), 0)?;
                validate_coordinate(row, column)?;
                if let Some(previous) = previous_column
                    && column <= previous
                {
                    return Err(invalid("worksheet columns are not strictly increasing"));
                }
                previous_column = Some(column);
                cell_count = cell_count.checked_add(1).ok_or(Error::CapacityOverflow {
                    resource: "worksheet cell count",
                })?;
                if cell_count > limits.max_cells {
                    return Err(limit("worksheet cells", cell_count, limits.max_cells));
                }
                let anchor = CellOffset {
                    row,
                    column,
                    offset: u64::try_from(record.offset())
                        .map_err(|_error| invalid("worksheet record offset overflows"))?,
                    record_kind: cell_kind.get(),
                };
                let bucket = (row, column / INDEX_COLUMN_BLOCK_WIDTH);
                let duplicate = buckets.last().is_some_and(|last: &CellOffset| {
                    (last.row, last.column / INDEX_COLUMN_BLOCK_WIDTH) == bucket
                });
                if !duplicate {
                    buckets
                        .try_reserve(1)
                        .map_err(|source| allocation("worksheet index buckets", source))?;
                    buckets.push(anchor);
                }
            },
            _ => {},
        }
    }
    if !seen_sheet_data || in_sheet_data {
        return Err(invalid("worksheet sheet-data framing is incomplete"));
    }
    Ok(WorksheetScan {
        anchors: buckets,
        row_headers,
    })
}

fn is_cell_kind(kind_value: raw::Kind) -> bool {
    (kind::CELL_BLANK.get()..=kind::FMLA_ERROR.get()).contains(&kind_value.get())
        || kind_value == kind::CELL_R_STRING
}

/// Decode only the cached scalar carried by an already verified cell record.
///
/// Formula tokens are deliberately left opaque; this helper validates the
/// cached-result field and returns the existing source-bound cell-value model.
pub(crate) fn decode_cached_value(
    kind_value: raw::Kind,
    payload: &[u8],
    limits: raw::Limits,
) -> Result<crate::cell_values::Value> {
    limits.validate()?;
    if payload.len() > limits.payload() {
        return Err(limit(
            "cached cell payload",
            payload.len(),
            limits.payload(),
        ));
    }
    let cell_string_limits = raw::Limits::new(limits.payload(), limits.string_units().min(32_767));
    if payload.len() < 8 {
        return Err(invalid("cell payload is shorter than its Cell header"));
    }
    if payload[7] & 0xfe != 0 {
        return Err(invalid("cell payload has reserved flag bits"));
    }
    match kind_value {
        kind::CELL_BLANK => {
            require_cached_exact(payload, 8)?;
            Ok(crate::cell_values::Value::Blank)
        },
        kind::CELL_RK => {
            require_cached_exact(payload, 12)?;
            let mut cursor = raw::Cursor::with_limits(&payload[8..], "BrtCellRk", limits);
            let value = cursor.read_rk()?;
            validate_cached_number(value)?;
            Ok(crate::cell_values::Value::RkNumber(value))
        },
        kind::CELL_ERROR => {
            require_cached_exact(payload, 9)?;
            Ok(crate::cell_values::Value::Error(
                crate::cell_values::CellError::from_code(payload[8])?,
            ))
        },
        kind::CELL_BOOL => {
            require_cached_exact(payload, 9)?;
            Ok(crate::cell_values::Value::Boolean(read_cached_bool(
                payload[8],
            )?))
        },
        kind::CELL_REAL => {
            require_cached_exact(payload, 16)?;
            let value = f64::from_le_bytes(
                payload[8..16]
                    .try_into()
                    .map_err(|_error| invalid("BrtCellReal value conversion"))?,
            );
            validate_cached_number(value)?;
            Ok(crate::cell_values::Value::Number(value))
        },
        kind::CELL_ST => {
            let mut cursor =
                raw::Cursor::with_limits(&payload[8..], "BrtCellSt", cell_string_limits);
            let value = cursor.read_wide_string()?;
            cursor.finish()?;
            Ok(crate::cell_values::Value::InlineString(value))
        },
        kind::CELL_ISST => {
            require_cached_exact(payload, 12)?;
            Ok(crate::cell_values::Value::SharedStringIndex(read_u32(
                payload, 8,
            )?))
        },
        kind::CELL_R_STRING => {
            require_cached_at_least(payload, 13)?;
            preflight_cached_rich_string(&payload[8..], limits)?;
            Ok(crate::cell_values::Value::RichString(
                crate::package::SharedString::parse(&payload[8..])?,
            ))
        },
        kind::FMLA_STRING => {
            require_cached_at_least(payload, 14)?;
            let mut cursor =
                raw::Cursor::with_limits(&payload[8..], "BrtFmlaString", cell_string_limits);
            let value = cursor.read_wide_string()?;
            if cursor.remaining() < 10 {
                return Err(invalid("BrtFmlaString is missing formula metadata"));
            }
            Ok(crate::cell_values::Value::FormulaStringCache(value))
        },
        kind::FMLA_NUM => {
            require_cached_at_least(payload, 26)?;
            let value = f64::from_le_bytes(
                payload[8..16]
                    .try_into()
                    .map_err(|_error| invalid("BrtFmlaNum value conversion"))?,
            );
            validate_cached_number(value)?;
            Ok(crate::cell_values::Value::FormulaNumberCache(value))
        },
        kind::FMLA_BOOL => {
            require_cached_at_least(payload, 19)?;
            Ok(crate::cell_values::Value::FormulaBooleanCache(
                read_cached_bool(payload[8])?,
            ))
        },
        kind::FMLA_ERROR => {
            require_cached_at_least(payload, 19)?;
            Ok(crate::cell_values::Value::FormulaErrorCache(
                crate::cell_values::CellError::from_code(payload[8])?,
            ))
        },
        _ => Err(Error::InvalidRecordType(kind_value.get())),
    }
}

fn require_cached_exact(payload: &[u8], expected: usize) -> Result<()> {
    if payload.len() != expected {
        return Err(Error::InvalidLength {
            expected,
            found: payload.len(),
        });
    }
    Ok(())
}

fn require_cached_at_least(payload: &[u8], expected: usize) -> Result<()> {
    if payload.len() < expected {
        return Err(Error::InvalidLength {
            expected,
            found: payload.len(),
        });
    }
    Ok(())
}

fn read_cached_bool(value: u8) -> Result<bool> {
    match value {
        0 => Ok(false),
        1 => Ok(true),
        _ => Err(invalid("cell Boolean value is neither zero nor one")),
    }
}

fn validate_cached_number(value: f64) -> Result<()> {
    if value.is_normal() || value.to_bits() == 0 {
        Ok(())
    } else {
        Err(invalid(
            "Xnum requires a normalized number or positive zero",
        ))
    }
}

fn preflight_cached_rich_string(bytes: &[u8], limits: raw::Limits) -> Result<()> {
    let mut cursor = raw::Cursor::with_limits(bytes, "indexed RichStr", limits);
    let flags = cursor.read_u8()?;
    preflight_cached_string(&mut cursor, limits.string_units().min(32_767))?;
    if flags & 1 != 0 {
        let count = usize::try_from(cursor.read_u32()?)
            .map_err(|_error| invalid("rich-string run count exceeds platform size"))?;
        let bytes = count.checked_mul(4).ok_or(Error::CapacityOverflow {
            resource: "indexed rich-string runs",
        })?;
        cursor.skip(bytes)?;
    }
    if flags & 2 != 0 {
        preflight_cached_string(&mut cursor, limits.string_units())?;
    }
    Ok(())
}

fn preflight_cached_string(cursor: &mut raw::Cursor<'_>, maximum: usize) -> Result<()> {
    let units = usize::try_from(cursor.read_u32()?)
        .map_err(|_error| invalid("rich-string length exceeds platform size"))?;
    if units > maximum {
        return Err(limit("indexed rich-string code units", units, maximum));
    }
    let bytes = units.checked_mul(2).ok_or(Error::CapacityOverflow {
        resource: "indexed rich-string bytes",
    })?;
    cursor.skip(bytes)?;
    Ok(())
}

fn encode_anchors(scan: &WorksheetScan, limits: Limits) -> Result<Vec<u8>> {
    let mut entries_by_block = BTreeMap::<u32, BTreeMap<(u32, u32), u64>>::new();
    let mut rows_by_block = BTreeMap::<u32, BTreeSet<u32>>::new();
    for &row in &scan.row_headers {
        let block_start = (row / INDEX_BLOCK_ROWS) * INDEX_BLOCK_ROWS;
        rows_by_block.entry(block_start).or_default().insert(row);
    }
    for anchor in &scan.anchors {
        let block_start = (anchor.row / INDEX_BLOCK_ROWS) * INDEX_BLOCK_ROWS;
        let key = (anchor.row, anchor.column / INDEX_COLUMN_BLOCK_WIDTH);
        rows_by_block
            .entry(block_start)
            .or_default()
            .insert(anchor.row);
        let entries = entries_by_block.entry(block_start).or_default();
        let slot = entries.entry(key).or_insert(anchor.offset);
        if anchor.offset < *slot {
            *slot = anchor.offset;
        }
    }

    let mut block_starts = rows_by_block.keys().copied().collect::<BTreeSet<_>>();
    block_starts.extend(entries_by_block.keys().copied());
    if block_starts.len() > limits.max_blocks {
        return Err(limit(
            "binary-index blocks",
            block_starts.len(),
            limits.max_blocks,
        ));
    }
    let mut output = Vec::new();
    if block_starts.is_empty() {
        write_empty_index_block(&mut output, limits)?;
        write_record(&mut output, kind::INDEX_PART_END, &[], limits)?;
        return Ok(output);
    }

    for block_start in block_starts {
        let entries = entries_by_block.remove(&block_start).unwrap_or_default();
        let mut rows = rows_by_block.remove(&block_start).unwrap_or_default();
        for &(row, _) in entries.keys() {
            rows.insert(row);
        }
        if rows.is_empty() {
            return Err(invalid("index block has no worksheet rows"));
        }
        let mut row_mask = 0_u32;
        let mut column_masks = [0_u16; 32];
        for &row in &rows {
            let row_offset = row
                .checked_sub(block_start)
                .ok_or_else(|| invalid("row precedes index block"))?;
            if row_offset >= INDEX_BLOCK_ROWS {
                return Err(invalid("row exceeds index block"));
            }
            row_mask |= 1_u32 << row_offset;
        }
        for &(row, column_block) in entries.keys() {
            let row_offset = usize::try_from(
                row.checked_sub(block_start)
                    .ok_or_else(|| invalid("row precedes index block"))?,
            )
            .map_err(|_error| invalid("row offset overflow"))?;
            let column_block = u16::try_from(column_block)
                .map_err(|_error| invalid("column block conversion failed"))?;
            column_masks[row_offset] |= 1_u16 << column_block;
        }
        let base_offset = entries.values().next().copied().unwrap_or(0);
        let mut payload = Vec::new();
        payload
            .try_reserve(24)
            .map_err(|source| allocation("BrtIndexBlock payload", source))?;
        payload.extend_from_slice(&block_start.to_le_bytes());
        payload.extend_from_slice(&(block_start + INDEX_BLOCK_ROWS).to_le_bytes());
        payload.extend_from_slice(&[0; 8]);
        payload.extend_from_slice(&[0; INDEX_BLOCK_UNUSED_BYTES]);
        write_record(&mut output, kind::INDEX_BLOCK, &payload, limits)?;

        let row_count = usize::try_from(row_mask.count_ones())
            .map_err(|_error| invalid("row count overflow"))?;
        let offsets_bytes = entries
            .len()
            .checked_mul(4)
            .ok_or(Error::CapacityOverflow {
                resource: "BrtIndexRowBlock offsets",
            })?;
        let row_block_capacity = 12usize
            .checked_add(row_count.checked_mul(2).ok_or(Error::CapacityOverflow {
                resource: "BrtIndexRowBlock masks",
            })?)
            .and_then(|value| value.checked_add(offsets_bytes))
            .ok_or(Error::CapacityOverflow {
                resource: "BrtIndexRowBlock payload",
            })?;
        let mut row_block = Vec::new();
        row_block
            .try_reserve(row_block_capacity)
            .map_err(|source| allocation("BrtIndexRowBlock payload", source))?;
        row_block.extend_from_slice(&row_mask.to_le_bytes());
        row_block.extend_from_slice(&base_offset.to_le_bytes());
        for row_offset in 0..INDEX_BLOCK_ROWS {
            let row_offset =
                usize::try_from(row_offset).map_err(|_error| invalid("row offset overflow"))?;
            if row_mask & (1_u32 << row_offset) != 0 {
                row_block.extend_from_slice(&column_masks[row_offset].to_le_bytes());
            }
        }
        for &offset in entries.values() {
            let relative = offset
                .checked_sub(base_offset)
                .ok_or_else(|| invalid("index offset precedes base offset"))?;
            let relative = u32::try_from(relative)
                .map_err(|_error| invalid("index sub-offset exceeds u32"))?;
            row_block.extend_from_slice(&relative.to_le_bytes());
        }
        write_record(&mut output, kind::INDEX_ROW_BLOCK, &row_block, limits)?;
    }
    write_record(&mut output, kind::INDEX_PART_END, &[], limits)?;
    Ok(output)
}

fn write_empty_index_block(output: &mut Vec<u8>, limits: Limits) -> Result<()> {
    let mut payload = Vec::new();
    payload
        .try_reserve(24)
        .map_err(|source| allocation("empty BrtIndexBlock payload", source))?;
    payload.extend_from_slice(&0_u32.to_le_bytes());
    payload.extend_from_slice(&INDEX_BLOCK_ROWS.to_le_bytes());
    payload.extend_from_slice(&[0; 8]);
    payload.extend_from_slice(&[0; INDEX_BLOCK_UNUSED_BYTES]);
    write_record(output, kind::INDEX_BLOCK, &payload, limits)
}

fn patch_offsets(source: &[u8], after: &WorksheetBinaryIndex, limits: Limits) -> Result<Vec<u8>> {
    let mut output = Vec::new();
    output
        .try_reserve(source.len())
        .map_err(|source| allocation("patched binary-index bytes", source))?;
    output.extend_from_slice(source);
    let after_offsets = after
        .buckets
        .iter()
        .map(|anchor| {
            (
                (anchor.row, anchor.column / INDEX_COLUMN_BLOCK_WIDTH),
                anchor.offset,
            )
        })
        .collect::<BTreeMap<_, _>>();
    let mut block_index = 0usize;
    let mut current_block = None;
    for record_result in raw::Records::try_with_limits(source, limits.raw)? {
        let record = record_result?;
        match record.kind() {
            kind::INDEX_BLOCK => {
                current_block = Some(block_index);
                block_index = block_index.checked_add(1).ok_or(Error::CapacityOverflow {
                    resource: "patched index block count",
                })?;
                continue;
            },
            kind::INDEX_ROW_BLOCK => {
                let source_block = current_block
                    .take()
                    .ok_or_else(|| invalid("patched row block has no preceding block"))?;
                let block = after
                    .index
                    .blocks
                    .get(source_block)
                    .ok_or_else(|| invalid("patched index block count changed"))?;
                if block.rows.columns.is_empty() {
                    continue;
                }
                let row_count = usize::try_from(block.rows.row_mask.count_ones())
                    .map_err(|_error| invalid("row count overflow"))?;
                let masks_end = 12usize
                    .checked_add(row_count.checked_mul(2).ok_or(Error::CapacityOverflow {
                        resource: "patched index mask offset",
                    })?)
                    .ok_or(Error::CapacityOverflow {
                        resource: "patched index mask offset",
                    })?;
                // The index record kind is 40 (< 0x80) and every generated source
                // record has a one-byte kind header. For source records with a
                // two-byte kind, derive the payload offset from Header::parse below.
                let (_, header_len) = raw::Header::parse(
                    source
                        .get(record.offset()..)
                        .ok_or_else(|| invalid("patched index record offset"))?,
                    limits.raw,
                )?;
                let payload_offset =
                    record
                        .offset()
                        .checked_add(header_len)
                        .ok_or(Error::CapacityOverflow {
                            resource: "patched index payload offset",
                        })?;
                let first_entry = block
                    .rows
                    .columns
                    .first()
                    .ok_or_else(|| invalid("cannot patch an empty index row block"))?;
                let first_key = (first_entry.row, u32::from(first_entry.column_block));
                let base_offset = *after_offsets
                    .get(&first_key)
                    .ok_or_else(|| invalid("patched index anchor is missing from worksheet"))?;
                let base_end = payload_offset
                    .checked_add(12)
                    .ok_or(Error::CapacityOverflow {
                        resource: "patched index base field",
                    })?;
                output
                    .get_mut(payload_offset + 4..base_end)
                    .ok_or_else(|| invalid("patched index base field outside source"))?
                    .copy_from_slice(&base_offset.to_le_bytes());
                let mut cursor =
                    payload_offset
                        .checked_add(masks_end)
                        .ok_or(Error::CapacityOverflow {
                            resource: "patched index offset array",
                        })?;
                for entry in &block.rows.columns {
                    let key = (entry.row, u32::from(entry.column_block));
                    let offset = *after_offsets
                        .get(&key)
                        .ok_or_else(|| invalid("patched index anchor is missing from worksheet"))?;
                    let relative = offset
                        .checked_sub(base_offset)
                        .ok_or_else(|| invalid("patched index offset precedes base"))?;
                    let relative = u32::try_from(relative)
                        .map_err(|_error| invalid("patched index sub-offset exceeds u32"))?;
                    let end = cursor.checked_add(4).ok_or(Error::CapacityOverflow {
                        resource: "patched index offset field",
                    })?;
                    output
                        .get_mut(cursor..end)
                        .ok_or_else(|| invalid("patched index offset outside source"))?
                        .copy_from_slice(&relative.to_le_bytes());
                    cursor = end;
                }
            },
            _ => {},
        }
    }
    if block_index != after.index.blocks.len() {
        return Err(invalid("patched index block count changed"));
    }
    Ok(output)
}

fn geometry(anchors: &[CellOffset]) -> BTreeMap<(u32, u32), u64> {
    let mut result = BTreeMap::new();
    for anchor in anchors {
        result
            .entry((anchor.row, anchor.column / INDEX_COLUMN_BLOCK_WIDTH))
            .or_insert(anchor.offset);
    }
    result
}

fn validate_coordinate(row: u32, column: u32) -> Result<()> {
    if row > MAX_ROW_INDEX {
        return Err(invalid("worksheet row exceeds 1,048,575"));
    }
    if column > MAX_COLUMN_INDEX {
        return Err(invalid("worksheet column exceeds 16,383"));
    }
    Ok(())
}

fn read_u16(bytes: &[u8], offset: usize) -> Result<u16> {
    let end = offset.checked_add(2).ok_or(Error::CapacityOverflow {
        resource: "binary-index field offset",
    })?;
    let field = bytes
        .get(offset..end)
        .ok_or_else(|| invalid("binary-index field is truncated"))?;
    Ok(u16::from_le_bytes(field.try_into().map_err(|_error| {
        invalid("binary-index u16 field conversion")
    })?))
}

fn read_u32(bytes: &[u8], offset: usize) -> Result<u32> {
    let end = offset.checked_add(4).ok_or(Error::CapacityOverflow {
        resource: "binary-index field offset",
    })?;
    let field = bytes
        .get(offset..end)
        .ok_or_else(|| invalid("binary-index field is truncated"))?;
    Ok(u32::from_le_bytes(field.try_into().map_err(|_error| {
        invalid("binary-index u32 field conversion")
    })?))
}

fn read_u64(bytes: &[u8], offset: usize) -> Result<u64> {
    let end = offset.checked_add(8).ok_or(Error::CapacityOverflow {
        resource: "binary-index field offset",
    })?;
    let field = bytes
        .get(offset..end)
        .ok_or_else(|| invalid("binary-index field is truncated"))?;
    Ok(u64::from_le_bytes(field.try_into().map_err(|_error| {
        invalid("binary-index u64 field conversion")
    })?))
}

fn write_record(
    output: &mut Vec<u8>,
    kind_value: raw::Kind,
    payload: &[u8],
    limits: Limits,
) -> Result<()> {
    let record_len = record_wire_len(kind_value, payload.len())?;
    let maximum = MAX_INDEX_STREAM_BYTES.min(limits.source_bytes);
    let new_len = output
        .len()
        .checked_add(record_len)
        .ok_or(Error::CapacityOverflow {
            resource: "binary-index output bytes",
        })?;
    if new_len > maximum {
        return Err(limit("binary-index output bytes", new_len, maximum));
    }
    output
        .try_reserve(record_len)
        .map_err(|source| allocation("binary-index output bytes", source))?;
    let mut writer = raw::Writer::with_limits(output, limits.raw);
    writer.write_record(kind_value, payload)?;
    Ok(())
}

fn record_wire_len(kind_value: raw::Kind, payload_len: usize) -> Result<usize> {
    let kind_len: usize = if kind_value.get() < 0x80 { 1 } else { 2 };
    let mut length = payload_len;
    let mut length_len = 1usize;
    while length >= 0x80 {
        length >>= 7;
        length_len = length_len.checked_add(1).ok_or(Error::CapacityOverflow {
            resource: "binary-index record length",
        })?;
    }
    kind_len
        .checked_add(length_len)
        .and_then(|value| value.checked_add(payload_len))
        .ok_or(Error::CapacityOverflow {
            resource: "binary-index record length",
        })
}

fn invalid(message: &str) -> Error {
    Error::InvalidFormat(format!("worksheet binary index: {message}"))
}

fn limit(resource: &'static str, actual: usize, maximum: usize) -> Error {
    Error::LimitExceeded {
        resource,
        actual,
        maximum,
    }
}

fn allocation(resource: &'static str, source: std::collections::TryReserveError) -> Error {
    Error::Allocation { resource, source }
}

fn copy_bytes(bytes: &[u8], resource: &'static str) -> Result<Vec<u8>> {
    let mut copy = Vec::new();
    copy.try_reserve(bytes.len())
        .map_err(|source| allocation(resource, source))?;
    copy.extend_from_slice(bytes);
    Ok(copy)
}

impl fmt::Display for CellOffset {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(
            formatter,
            "cell({}, {}) at byte {} (record {})",
            self.row, self.column, self.offset, self.record_kind
        )
    }
}

#[cfg(test)]
#[allow(
    clippy::expect_used,
    reason = "unit fixtures use checked, local record construction and panic on an invalid test setup"
)]
mod tests {
    use super::*;

    fn worksheet_with_cells(cells: &[(u32, u32)]) -> Vec<u8> {
        let mut bytes = Vec::new();
        let mut writer = raw::Writer::new(&mut bytes);
        writer
            .write_record(kind::BEGIN_SHEET_DATA, &[])
            .expect("begin");
        let mut row = None;
        for &(row_value, column) in cells {
            if row != Some(row_value) {
                let mut payload = Vec::new();
                payload.extend_from_slice(&row_value.to_le_bytes());
                payload.extend_from_slice(&[0; 16]);
                writer.write_record(kind::ROW_HDR, &payload).expect("row");
                row = Some(row_value);
            }
            let mut payload = Vec::new();
            payload.extend_from_slice(&column.to_le_bytes());
            payload.extend_from_slice(&[0; 4]);
            writer
                .write_record(kind::CELL_BLANK, &payload)
                .expect("cell");
        }
        writer.write_record(kind::END_SHEET_DATA, &[]).expect("end");
        bytes
    }

    #[test]
    fn generated_index_binds_and_looks_up_exact_cells() {
        let worksheet = worksheet_with_cells(&[(0, 0), (0, 1024), (33, 2)]);
        let index = encode_for_worksheet(&worksheet, Limits::DEFAULT).expect("index");
        let bound = WorksheetBinaryIndex::from_parts_default(&index, &worksheet).expect("bind");
        assert_eq!(
            bound
                .lookup_record(&worksheet, 0, 0)
                .expect("lookup")
                .expect("cell")
                .0
                .column,
            0
        );
        assert_eq!(
            bound
                .lookup_record(&worksheet, 0, 1024)
                .expect("lookup")
                .expect("cell")
                .0
                .column,
            1024
        );
        assert!(
            bound
                .lookup_record(&worksheet, 0, 1)
                .expect("lookup")
                .is_none()
        );
        assert_eq!(
            bound
                .lookup_record(&worksheet, 33, 2)
                .expect("lookup")
                .expect("cell")
                .0
                .row,
            33
        );
    }

    #[test]
    fn forged_payload_offset_is_rejected() {
        let worksheet = worksheet_with_cells(&[(0, 0)]);
        let mut index = encode_for_worksheet(&worksheet, Limits::DEFAULT).expect("index");
        // BrtIndexRowBlock payload starts after the two-byte record header;
        // its base offset is at payload+4 and must point to the cell header.
        let base = 2 + 4;
        index[base..base + 8].copy_from_slice(&1_u64.to_le_bytes());
        assert!(WorksheetBinaryIndex::from_parts_default(&index, &worksheet).is_err());
    }

    #[test]
    fn shifted_known_offsets_are_patched_in_place() {
        let worksheet = worksheet_with_cells(&[(0, 0), (0, 1)]);
        let index = encode_for_worksheet(&worksheet, Limits::DEFAULT).expect("index");
        let mut shifted = Vec::new();
        let mut writer = raw::Writer::new(&mut shifted);
        writer.write_record(kind::FRT_BEGIN, &[0; 4]).expect("frt");
        shifted.extend_from_slice(&worksheet);

        let maintained = maintain_index(Some(&index), &worksheet, &shifted, Limits::DEFAULT)
            .expect("maintain")
            .expect("existing index");
        assert_ne!(maintained, index);
        let bound = WorksheetBinaryIndex::from_parts_default(&maintained, &shifted)
            .expect("patched index binds");
        assert_eq!(
            bound
                .lookup_record(&shifted, 0, 1)
                .expect("lookup")
                .expect("cell")
                .0
                .column,
            1
        );
    }

    #[test]
    fn non_aligned_source_partition_is_patched_without_canonicalization() {
        let worksheet = worksheet_with_cells(&[(2, 0), (3, 1_024)]);
        let scan = scan_worksheet(&worksheet, Limits::DEFAULT).expect("worksheet scan");
        let first_offset = scan
            .anchors
            .iter()
            .find(|anchor| anchor.row == 2 && anchor.column / INDEX_COLUMN_BLOCK_WIDTH == 0)
            .expect("first anchor")
            .offset;
        let second_offset = scan
            .anchors
            .iter()
            .find(|anchor| anchor.row == 3 && anchor.column / INDEX_COLUMN_BLOCK_WIDTH == 1)
            .expect("second anchor")
            .offset;
        let mut index = Vec::new();
        let mut writer = raw::Writer::new(&mut index);
        let mut block = Vec::new();
        block.extend_from_slice(&2_u32.to_le_bytes());
        block.extend_from_slice(&4_u32.to_le_bytes());
        block.extend_from_slice(&[0; 8]);
        block.extend_from_slice(&[0; 4]);
        writer
            .write_record(kind::INDEX_BLOCK, &block)
            .expect("native block");
        let mut rows = Vec::new();
        rows.extend_from_slice(&3_u32.to_le_bytes());
        rows.extend_from_slice(&first_offset.to_le_bytes());
        rows.extend_from_slice(&1_u16.to_le_bytes());
        rows.extend_from_slice(&2_u16.to_le_bytes());
        rows.extend_from_slice(&0_u32.to_le_bytes());
        rows.extend_from_slice(
            &u32::try_from(second_offset - first_offset)
                .expect("relative offset")
                .to_le_bytes(),
        );
        writer
            .write_record(kind::INDEX_ROW_BLOCK, &rows)
            .expect("native rows");
        writer
            .write_record(kind::INDEX_PART_END, &[])
            .expect("index end");
        drop(writer);

        let mut shifted = Vec::new();
        raw::Writer::new(&mut shifted)
            .write_record(kind::FRT_BEGIN, &[0; 4])
            .expect("prefix");
        shifted.extend_from_slice(&worksheet);
        let maintained = maintain_index(Some(&index), &worksheet, &shifted, Limits::DEFAULT)
            .expect("maintain")
            .expect("existing index");
        let first = raw::Records::new(&maintained)
            .next()
            .expect("index block record")
            .expect("index block framing");
        assert_eq!(first.kind(), kind::INDEX_BLOCK);
        assert_eq!(read_u32(first.payload(), 0).expect("row start"), 2);
        assert_eq!(read_u32(first.payload(), 4).expect("row end"), 4);
        WorksheetBinaryIndex::from_parts_default(&maintained, &shifted)
            .expect("patched native partition binds");
    }

    #[test]
    fn unchanged_geometry_retains_opaque_index_bytes() {
        let worksheet = worksheet_with_cells(&[(0, 0)]);
        let canonical = encode_for_worksheet(&worksheet, Limits::DEFAULT).expect("index");
        let mut opaque = Vec::new();
        raw::Writer::new(&mut opaque)
            .write_record(
                raw::Kind::new(0x1234).expect("opaque kind"),
                b"extension bytes",
            )
            .expect("opaque record");
        opaque.extend_from_slice(&canonical);

        let first_cell = raw::Records::new(&worksheet)
            .find_map(|record| {
                let record = record.expect("record");
                (is_cell_kind(record.kind())).then_some(record)
            })
            .expect("cell record");
        let (_, header_len) =
            raw::Header::parse(&worksheet[first_cell.offset()..], raw::Limits::DEFAULT)
                .expect("cell header");
        let mut changed = worksheet.clone();
        changed[first_cell.offset() + header_len + 4] ^= 1;

        let maintained = maintain_index(Some(&opaque), &worksheet, &changed, Limits::DEFAULT)
            .expect("maintain")
            .expect("existing index");
        assert_eq!(maintained, opaque);
    }

    #[test]
    fn empty_sheet_has_one_empty_block_and_end_marker() {
        let worksheet = worksheet_with_cells(&[]);
        let index = encode_for_worksheet(&worksheet, Limits::DEFAULT).expect("index");
        assert_eq!(
            index,
            vec![
                0x2a, 0x18, 0, 0, 0, 0, 32, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0,
                0, 0x95, 0x02, 0,
            ]
        );
        let bound = WorksheetBinaryIndex::from_parts_default(&index, &worksheet).expect("bind");
        assert!(
            bound
                .lookup_record(&worksheet, 0, 0)
                .expect("lookup")
                .is_none()
        );
    }
}

#[cfg(test)]
#[allow(
    clippy::expect_used,
    reason = "decoder fixtures use checked, local record construction and panic on an invalid test setup"
)]
mod cached_value_tests {
    use super::{decode_cached_value, kind, raw};

    #[test]
    fn preservation_copy_is_bounded_before_noop_or_geometry_scan() {
        let limits = super::Limits::new_with_source_bytes(raw::Limits::DEFAULT, 8, 100, 100, 100);
        for after in [b"a".as_slice(), b"b".as_slice()] {
            let result = super::maintain_index(Some(b"123456789"), b"a", after, limits);
            assert!(matches!(
                result,
                Err(super::Error::LimitExceeded {
                    resource: "binary-index bytes",
                    actual: 9,
                    maximum: 8,
                })
            ));
        }
    }

    #[test]
    fn xnum_cache_rejects_negative_zero_denormals_and_nonfinite_values() {
        for (record_kind, length) in [(kind::CELL_REAL, 16), (kind::FMLA_NUM, 26)] {
            for value in [-0.0, f64::from_bits(1), f64::INFINITY, f64::NAN] {
                let mut payload = vec![0; length];
                payload[8..16].copy_from_slice(&value.to_le_bytes());
                assert!(decode_cached_value(record_kind, &payload, raw::Limits::DEFAULT).is_err());
            }
            for value in [0.0_f64, 1.0, -1.0] {
                let mut payload = vec![0; length];
                payload[8..16].copy_from_slice(&value.to_le_bytes());
                assert!(decode_cached_value(record_kind, &payload, raw::Limits::DEFAULT).is_ok());
            }
        }
    }

    #[test]
    fn rk_cache_obeys_xnum_after_decoding() {
        for encoded in [0x8000_0000_u32, 4, 0x7ff0_0000] {
            let mut payload = vec![0; 8];
            payload.extend_from_slice(&encoded.to_le_bytes());
            assert!(decode_cached_value(kind::CELL_RK, &payload, raw::Limits::DEFAULT).is_err());
        }
        for encoded in [2_u32, 6] {
            let mut payload = vec![0; 8];
            payload.extend_from_slice(&encoded.to_le_bytes());
            assert!(decode_cached_value(kind::CELL_RK, &payload, raw::Limits::DEFAULT).is_ok());
        }
    }

    #[test]
    fn inline_and_formula_strings_obey_the_format_character_ceiling() {
        for record_kind in [kind::CELL_ST, kind::FMLA_STRING] {
            for units in [32_767_u32, 32_768] {
                let mut payload = vec![0; 8];
                payload.extend_from_slice(&units.to_le_bytes());
                for _ in 0..units {
                    payload.extend_from_slice(&[b'x', 0]);
                }
                if record_kind == kind::FMLA_STRING {
                    payload.extend_from_slice(&[0; 10]);
                }
                let result = decode_cached_value(record_kind, &payload, raw::Limits::DEFAULT);
                assert_eq!(result.is_ok(), units == 32_767);
            }
        }
    }

    #[test]
    fn rich_string_text_is_bounded_before_semantic_parsing() {
        let mut payload = vec![0; 9];
        payload.extend_from_slice(&3_u32.to_le_bytes());
        payload.extend_from_slice(&[b'a', 0, b'b', 0, b'c', 0]);
        assert!(
            decode_cached_value(kind::CELL_R_STRING, &payload, raw::Limits::new(100, 2)).is_err()
        );
        assert!(
            decode_cached_value(kind::CELL_R_STRING, &payload, raw::Limits::new(100, 3)).is_ok()
        );

        let mut phonetic = vec![0; 8];
        phonetic.push(2);
        phonetic.extend_from_slice(&1_u32.to_le_bytes());
        phonetic.extend_from_slice(&[b'x', 0]);
        phonetic.extend_from_slice(&3_u32.to_le_bytes());
        phonetic.extend_from_slice(&[b'a', 0, b'b', 0, b'c', 0]);
        phonetic.extend_from_slice(&[0; 8]);
        assert!(
            decode_cached_value(kind::CELL_R_STRING, &phonetic, raw::Limits::new(100, 2)).is_err()
        );
    }
}
