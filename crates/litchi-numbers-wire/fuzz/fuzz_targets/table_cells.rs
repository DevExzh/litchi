#![no_main]

use std::hint::black_box;

use libfuzzer_sys::fuzz_target;
use litchi_iwa_protos::numbers_table_cell_storage_codec::{DecodeOptions, DecodeReport};
use litchi_numbers_wire::BncCell;
use litchi_numbers_wire::table_cells::{
    AllocationTarget, CellSource, CellValueSink, Message, TILE_MESSAGE_KIND, TableCellIssue,
    TableCellReadBudget, TableDimensions, TileReference, read_table_cells,
};

const MAX_INPUT_BYTES: usize = 1_024;
const MAX_FIELDS: usize = 1_024;
const MAX_WORK_BYTES: usize = 16 * 1_024;
const MAX_REFERENCES: usize = 64;
const MAX_TEXT_BYTES: usize = 1_024;
const MAX_NESTING: u32 = 32;
const MAX_MATERIALIZED_CELLS: usize = 64;
const MAX_CELL_SOURCE_BYTES: usize = 4 * MAX_INPUT_BYTES;
const MAX_ROW_ALLOCATIONS: usize = 64;

/// A small checked-in wire fixture with one actual BNC cell row.
///
/// Keeping a valid payload beside the arbitrary source path ensures that the
/// fuzzer reaches row-span iteration and value classification even when the
/// input is not a valid protobuf tile.
const ACTUAL_TILE_FIXTURE: &[u8] = &[
    0x08, 0x01, // max_column
    0x10, 0x00, // max_row
    0x18, 0x01, // num_cells
    0x20, 0x01, // num_rows
    0x2a, 0x16, // one 22-byte row
    0x08, 0x00, // tile_row_index
    0x10, 0x01, // cell_count
    0x1a, 0x0c, // one minimal BNC cell
    0x05, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x22,
    0x02, // narrow offsets
    0x00, 0x00,
];

#[derive(Debug)]
enum FuzzError {
    Budget,
    Issue(TableCellIssue),
}

#[derive(Default)]
struct FuzzBudget {
    storage_work: usize,
    materialized_cells: usize,
    cell_source_bytes: usize,
    row_allocations: usize,
}

impl TableCellReadBudget for FuzzBudget {
    type Error = FuzzError;

    fn storage_options(&mut self, source: &[u8]) -> Result<DecodeOptions, Self::Error> {
        Ok(DecodeOptions::new(
            source.len().max(1),
            MAX_FIELDS,
            MAX_WORK_BYTES,
            MAX_NESTING,
            MAX_REFERENCES,
            MAX_TEXT_BYTES,
        ))
    }

    fn charge_storage_report(&mut self, report: DecodeReport) -> Result<(), Self::Error> {
        let Some(work) = self.storage_work.checked_add(report.work_bytes()) else {
            return Err(FuzzError::Budget);
        };
        if work > MAX_WORK_BYTES {
            return Err(FuzzError::Budget);
        }
        self.storage_work = work;
        Ok(())
    }

    fn check_materialized_cells(&mut self, observed: usize) -> Result<(), Self::Error> {
        if observed > MAX_MATERIALIZED_CELLS {
            return Err(FuzzError::Budget);
        }
        self.materialized_cells = observed;
        Ok(())
    }

    fn charge_cell_source(&mut self, bytes: usize) -> Result<(), Self::Error> {
        let Some(total) = self.cell_source_bytes.checked_add(bytes) else {
            return Err(FuzzError::Budget);
        };
        if total > MAX_CELL_SOURCE_BYTES {
            return Err(FuzzError::Budget);
        }
        self.cell_source_bytes = total;
        Ok(())
    }

    fn charge_allocation(
        &mut self,
        _target: AllocationTarget,
        amount: usize,
    ) -> Result<(), Self::Error> {
        let Some(total) = self.row_allocations.checked_add(amount) else {
            return Err(FuzzError::Budget);
        };
        if total > MAX_ROW_ALLOCATIONS {
            return Err(FuzzError::Budget);
        }
        self.row_allocations = total;
        Ok(())
    }

    fn map_issue(&mut self, issue: TableCellIssue) -> Self::Error {
        FuzzError::Issue(issue)
    }
}

struct FuzzSink {
    source_start: usize,
    source_end: usize,
    cells: usize,
}

impl FuzzSink {
    fn new(source: &[u8]) -> Self {
        let source_start = source.as_ptr() as usize;
        let source_end = source_start.saturating_add(source.len());
        Self {
            source_start,
            source_end,
            cells: 0,
        }
    }

    fn assert_borrowed(&self, bytes: &[u8]) {
        if bytes.is_empty() {
            return;
        }
        let start = bytes.as_ptr() as usize;
        let end = start.saturating_add(bytes.len());
        assert!(start >= self.source_start && end <= self.source_end);
    }
}

impl CellValueSink<FuzzBudget> for FuzzSink {
    fn visit_cell(
        &mut self,
        cell: CellSource<'_>,
        _budget: &mut FuzzBudget,
    ) -> Result<(), FuzzError> {
        self.assert_borrowed(cell.bytes());
        self.cells = self.cells.checked_add(1).ok_or(FuzzError::Budget)?;
        black_box((cell.row(), cell.column(), cell.value(), cell.bytes().len()));
        Ok(())
    }
}

fn varint(mut value: u64, output: &mut Vec<u8>) {
    loop {
        let byte = u8::try_from(value & 0x7f).unwrap_or(0);
        value >>= 7;
        if value == 0 {
            output.push(byte);
            return;
        }
        output.push(byte | 0x80);
    }
}

fn field_varint(number: u32, value: u64, output: &mut Vec<u8>) {
    varint(u64::from(number) << 3, output);
    varint(value, output);
}

fn field_bytes(number: u32, payload: &[u8], output: &mut Vec<u8>) {
    varint((u64::from(number) << 3) | 2, output);
    varint(u64::try_from(payload.len()).unwrap_or(u64::MAX), output);
    output.extend_from_slice(payload);
}

fn pre_bnc_empty() -> Vec<u8> {
    vec![4, 0, 0xa5, 0x5a, 0, 0, 0, 0, 0x19, 0x91, 0x71, 0x17]
}

fn bnc_number(seed: u8) -> Vec<u8> {
    let mut cell = BncCell::minimal();
    let value = f64::from(seed) + 0.5;
    if cell.set_plain_number(value).is_ok() {
        cell.encode()
    } else {
        BncCell::minimal().encode()
    }
}

fn offsets(starts: &[Option<usize>]) -> Vec<u8> {
    let mut output = Vec::with_capacity(starts.len().saturating_mul(2));
    for start in starts {
        let encoded = start
            .and_then(|value| u16::try_from(value).ok())
            .unwrap_or(u16::MAX);
        output.extend_from_slice(&encoded.to_le_bytes());
    }
    output
}

fn make_row(row_index: u32, seed: u8) -> Vec<u8> {
    let cell_count = usize::from(seed % 4) + 1;
    let mut pre_storage = Vec::new();
    let mut pre_starts = vec![None; 4];
    let mut modern_storage = Vec::new();
    let mut modern_starts = vec![None; 4];
    for column in 0..cell_count {
        let cell_seed = seed.wrapping_add(u8::try_from(column).unwrap_or(0));
        let pre_cell = if cell_seed & 1 == 0 {
            pre_bnc_empty()
        } else {
            bnc_number(cell_seed)
        };
        pre_starts[column] = Some(pre_storage.len());
        pre_storage.extend_from_slice(&pre_cell);

        let modern_cell = if cell_seed & 2 == 0 {
            BncCell::minimal().encode()
        } else {
            bnc_number(cell_seed)
        };
        modern_starts[column] = Some(modern_storage.len());
        modern_storage.extend_from_slice(&modern_cell);
    }

    let mut output = Vec::new();
    field_varint(1, u64::from(row_index), &mut output);
    field_varint(
        2,
        u64::try_from(cell_count).unwrap_or(u64::MAX),
        &mut output,
    );
    field_bytes(3, &pre_storage, &mut output);
    field_bytes(4, &offsets(&pre_starts), &mut output);
    if seed & 4 != 0 {
        field_bytes(6, &modern_storage, &mut output);
        field_bytes(7, &offsets(&modern_starts), &mut output);
    }
    output
}

fn make_tile(data: &[u8]) -> Vec<u8> {
    let row_count = usize::from(data.first().copied().unwrap_or(0) % 4) + 1;
    let mut rows = Vec::with_capacity(row_count);
    let mut cell_count = 0usize;
    for row in 0..row_count {
        let seed = data
            .get(row.saturating_add(1))
            .copied()
            .unwrap_or_else(|| data.first().copied().unwrap_or(0));
        let row_data = make_row(u32::try_from(row).unwrap_or(0), seed);
        cell_count = cell_count.saturating_add(usize::from(seed % 4) + 1);
        rows.push(row_data);
    }

    let mut output = Vec::new();
    let bogus_metadata = data.first().is_some_and(|value| value & 0x80 != 0);
    field_varint(
        1,
        if bogus_metadata {
            u64::from(u32::MAX)
        } else {
            3
        },
        &mut output,
    );
    field_varint(
        2,
        if bogus_metadata {
            u64::from(u32::MAX)
        } else {
            u64::try_from(row_count.saturating_sub(1)).unwrap_or(u64::MAX)
        },
        &mut output,
    );
    field_varint(
        3,
        if bogus_metadata {
            u64::from(u32::MAX)
        } else {
            u64::try_from(cell_count).unwrap_or(u64::MAX)
        },
        &mut output,
    );
    field_varint(
        4,
        if bogus_metadata {
            u64::from(u32::MAX)
        } else {
            u64::try_from(row_count).unwrap_or(u64::MAX)
        },
        &mut output,
    );
    for row in rows {
        field_bytes(5, &row, &mut output);
    }
    output
}

fn run_route(source: &[u8], route: u8) {
    let before = source.to_vec();
    let mut budget = FuzzBudget::default();
    let mut sink = FuzzSink::new(source);
    let result = read_table_cells(
        TableDimensions::new(8, 4, 4),
        [TileReference::new(0, 7)],
        |object_id, _budget| {
            let messages = match route % 4 {
                0 => vec![Message::new(TILE_MESSAGE_KIND, source)],
                1 => vec![Message::new(TILE_MESSAGE_KIND.wrapping_add(1), source)],
                2 => Vec::new(),
                _ => vec![
                    Message::new(TILE_MESSAGE_KIND, source),
                    Message::new(TILE_MESSAGE_KIND, source),
                ],
            };
            Ok((object_id == 7).then_some(messages))
        },
        &mut budget,
        &mut sink,
    );
    let _ = black_box(
        result
            .as_ref()
            .map(|report| (report.tiles(), report.rows(), report.cells())),
    );
    if let Some(error) = result.as_ref().err() {
        match error {
            FuzzError::Budget => {
                black_box(0_u8);
            },
            FuzzError::Issue(issue) => {
                black_box(format!("{issue:?}"));
            },
        };
    }
    black_box(sink.cells);
    assert_eq!(source, before.as_slice());
}

fuzz_target!(|data: &[u8]| {
    if data.len() > MAX_INPUT_BYTES {
        return;
    }

    run_route(data, data.first().copied().unwrap_or(0));

    let generated = make_tile(data);
    run_route(&generated, data.get(1).copied().unwrap_or(0));
    run_route(ACTUAL_TILE_FIXTURE, data.get(2).copied().unwrap_or(0));

    if let Some(first) = data.first().copied() {
        let mut mutated = generated;
        if let Some(byte) = mutated.get_mut(0) {
            *byte ^= first;
        }
        run_route(&mutated, first.rotate_left(1));
    }
});
