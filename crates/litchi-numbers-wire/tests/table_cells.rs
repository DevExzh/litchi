//! Independent wire-contract tests for sparse Numbers table-cell rows.
//!
//! The tile and row envelopes in this file are handwritten protobuf messages.
//! Cell payloads are also built from the public BNC and pre-BNC wire helpers,
//! so the tests exercise the strict storage codec with source data that does
//! not come from its generated fixtures.  The selected-table reader consumes
//! this same borrowed row surface: every test keeps the distinction between a
//! missing slot and a present cell, validates the complete offset table before
//! exposing a span, and stages callbacks until the enclosing decode succeeds.

use litchi_iwa_protos::numbers_table_cell_storage_codec::{
    CellSpan, DecodeError, DecodeLimit, DecodeOptions, StorageVisitor, TileRowInfoSnapshot,
    decode_tile_row_info, decode_tile_row_info_with_report, decode_tile_with_report,
    decode_tile_with_visitor,
};
use litchi_numbers_wire::cell_value::{CellValueSource, ValueSource};
use litchi_numbers_wire::pre_bnc::PreBncCellView;
use litchi_numbers_wire::table_cells::{
    AllocationTarget, CellSource, CellValueSink, Message, TILE_MESSAGE_KIND, TableCellIssue,
    TableCellReadBudget, TableCellReadReport, TableDimensions, TileReference, read_table_cells,
};
use litchi_numbers_wire::{BncCell, BncCellView, StoredValue};

fn varint(mut value: u64) -> Vec<u8> {
    let mut output = Vec::new();
    loop {
        let byte = u8::try_from(value & 0x7f).expect("a varint chunk fits in a byte");
        value >>= 7;
        if value == 0 {
            output.push(byte);
            return output;
        }
        output.push(byte | 0x80);
    }
}

fn field_key(number: u32, wire_type: u8) -> Vec<u8> {
    varint((u64::from(number) << 3) | u64::from(wire_type))
}

fn field_varint(number: u32, value: u64) -> Vec<u8> {
    let mut output = field_key(number, 0);
    output.extend(varint(value));
    output
}

fn field_bytes(number: u32, payload: &[u8]) -> Vec<u8> {
    let mut output = field_key(number, 2);
    output.extend(varint(
        u64::try_from(payload.len()).expect("test payload length fits in u64"),
    ));
    output.extend_from_slice(payload);
    output
}

fn options(source: &[u8]) -> DecodeOptions {
    DecodeOptions::new(
        source.len().saturating_add(64),
        usize::MAX,
        usize::MAX,
        64,
        usize::MAX,
        usize::MAX,
    )
}

fn narrow_offsets(slots: &[Option<u16>]) -> Vec<u8> {
    let mut output = Vec::with_capacity(slots.len().saturating_mul(2));
    for slot in slots {
        output.extend_from_slice(&slot.unwrap_or(u16::MAX).to_le_bytes());
    }
    output
}

fn row(
    tile_row_index: u32,
    cell_count: u32,
    pre_storage: &[u8],
    pre_offsets: &[u8],
    modern: Option<(&[u8], &[u8])>,
    wide_offsets: Option<bool>,
) -> Vec<u8> {
    let mut output = field_varint(1, u64::from(tile_row_index));
    output.extend(field_varint(2, u64::from(cell_count)));
    output.extend(field_bytes(3, pre_storage));
    output.extend(field_bytes(4, pre_offsets));
    if let Some((storage, offsets)) = modern {
        output.extend(field_bytes(6, storage));
        output.extend(field_bytes(7, offsets));
    }
    if let Some(wide) = wide_offsets {
        output.extend(field_varint(8, u64::from(wide)));
    }
    output
}

fn tile(max_column: u32, max_row: u32, num_cells: u32, num_rows: u32, rows: &[Vec<u8>]) -> Vec<u8> {
    let mut output = field_varint(1, u64::from(max_column));
    output.extend(field_varint(2, u64::from(max_row)));
    output.extend(field_varint(3, u64::from(num_cells)));
    output.extend(field_varint(4, u64::from(num_rows)));
    for row in rows {
        output.extend(field_bytes(5, row));
    }
    output
}

fn bnc_empty() -> Vec<u8> {
    BncCell::minimal().encode()
}

fn bnc_number(value: f64) -> Vec<u8> {
    let mut cell = BncCell::minimal();
    cell.set_plain_number(value)
        .expect("finite test number is encodable");
    cell.encode()
}

fn pre_bnc_empty(version: u8) -> Vec<u8> {
    let header_length = if version <= 1 { 8 } else { 12 };
    let mut output = vec![0; header_length];
    output[0] = version;
    if version == 4 {
        output[1] = 0;
        output[2..4].copy_from_slice(&[0xa5, 0x5a]);
    } else {
        output[1] = 0x5a;
        output[2] = 0;
        output[3] = 0xa5;
    }
    if version <= 1 {
        output[4..6].copy_from_slice(&[0, 0]);
        output[6..8].copy_from_slice(&[0xc3, 0x3d]);
    } else {
        output[4..8].copy_from_slice(&[0, 0, 0, 0]);
        output[8..12].copy_from_slice(&[0x19, 0x91, 0x71, 0x17]);
    }
    output
}

fn pre_bnc_number(value: f64) -> Vec<u8> {
    const NUMBER_FLAG: u32 = 0x0000_0020;
    let mut output = pre_bnc_empty(4);
    output[1] = 2;
    output[4..8].copy_from_slice(&NUMBER_FLAG.to_le_bytes());
    output.extend_from_slice(&value.to_le_bytes());
    output
}

#[derive(Debug, Default, PartialEq, Eq)]
struct RowFact {
    index: u32,
    cell_count: u32,
    storage_ptr: usize,
    storage_len: usize,
    offsets_ptr: usize,
    offsets_len: usize,
    modern_storage_ptr: Option<usize>,
    modern_storage_len: Option<usize>,
    modern_offsets_ptr: Option<usize>,
    modern_offsets_len: Option<usize>,
}

impl RowFact {
    fn from_row(row: TileRowInfoSnapshot<'_>) -> Self {
        let (storage, offsets) = row.cell_storage_and_offsets();
        Self {
            index: row.tile_row_index(),
            cell_count: row.cell_count(),
            storage_ptr: storage.as_ptr() as usize,
            storage_len: storage.len(),
            offsets_ptr: offsets.as_ptr() as usize,
            offsets_len: offsets.len(),
            modern_storage_ptr: row
                .cell_storage_buffer()
                .map(|value| value.as_ptr() as usize),
            modern_storage_len: row.cell_storage_buffer().map(<[u8]>::len),
            modern_offsets_ptr: row.cell_offsets().map(|value| value.as_ptr() as usize),
            modern_offsets_len: row.cell_offsets().map(<[u8]>::len),
        }
    }
}

#[derive(Debug, Default)]
struct RowCollector {
    rows: Vec<RowFact>,
}

impl StorageVisitor for RowCollector {
    fn visit_tile_row(&mut self, row: TileRowInfoSnapshot<'_>) -> Result<(), DecodeError> {
        self.rows.push(RowFact::from_row(row));
        Ok(())
    }
}

fn contains_range(source: &[u8], pointer: usize, length: usize) -> bool {
    let begin = source.as_ptr() as usize;
    let end = begin.saturating_add(source.len());
    pointer >= begin && pointer.saturating_add(length) <= end
}

fn assert_borrowed(source: &[u8], row: &RowFact) {
    assert!(contains_range(source, row.storage_ptr, row.storage_len));
    assert!(contains_range(source, row.offsets_ptr, row.offsets_len));
    if let (Some(pointer), Some(length)) = (row.modern_storage_ptr, row.modern_storage_len) {
        assert!(contains_range(source, pointer, length));
    }
    if let (Some(pointer), Some(length)) = (row.modern_offsets_ptr, row.modern_offsets_len) {
        assert!(contains_range(source, pointer, length));
    }
}

fn span_bytes(span: CellSpan, storage: &[u8]) -> &[u8] {
    span.bytes(storage).expect("validated span is in storage")
}

#[derive(Debug, PartialEq)]
enum HarnessError {
    Issue(TableCellIssue),
    StorageBudget { observed: usize, maximum: usize },
    CellBudget { observed: usize, maximum: usize },
    AllocationBudget { observed: usize, maximum: usize },
}

#[derive(Debug, Default)]
struct HarnessBudget {
    storage_work_limit: usize,
    storage_work: usize,
    cell_count_limit: usize,
    cell_count: usize,
    cell_source_limit: usize,
    cell_source_bytes: usize,
    allocation_limit: usize,
    allocations: usize,
    storage_reports: Vec<usize>,
    mapped_issues: Vec<TableCellIssue>,
    events: Vec<&'static str>,
}

impl HarnessBudget {
    fn permissive() -> Self {
        Self {
            storage_work_limit: usize::MAX,
            cell_count_limit: usize::MAX,
            cell_source_limit: usize::MAX,
            allocation_limit: usize::MAX,
            ..Self::default()
        }
    }
}

impl TableCellReadBudget for HarnessBudget {
    type Error = HarnessError;

    fn storage_options(&mut self, source: &[u8]) -> Result<DecodeOptions, Self::Error> {
        self.events.push("options");
        Ok(DecodeOptions::new(
            source.len().max(1),
            usize::MAX,
            usize::MAX,
            64,
            usize::MAX,
            usize::MAX,
        ))
    }

    fn charge_storage_report(
        &mut self,
        report: litchi_iwa_protos::numbers_table_cell_storage_codec::DecodeReport,
    ) -> Result<(), Self::Error> {
        self.events.push("storage");
        let observed = self.storage_work.saturating_add(report.work_bytes());
        if observed > self.storage_work_limit {
            return Err(HarnessError::StorageBudget {
                observed,
                maximum: self.storage_work_limit,
            });
        }
        self.storage_work = observed;
        self.storage_reports.push(report.work_bytes());
        Ok(())
    }

    fn check_materialized_cells(&mut self, observed: usize) -> Result<(), Self::Error> {
        self.events.push("cells");
        if observed > self.cell_count_limit {
            return Err(HarnessError::CellBudget {
                observed,
                maximum: self.cell_count_limit,
            });
        }
        self.cell_count = observed;
        Ok(())
    }

    fn charge_cell_source(&mut self, bytes: usize) -> Result<(), Self::Error> {
        self.events.push("source");
        let observed = self.cell_source_bytes.saturating_add(bytes);
        if observed > self.cell_source_limit {
            return Err(HarnessError::CellBudget {
                observed,
                maximum: self.cell_source_limit,
            });
        }
        self.cell_source_bytes = observed;
        Ok(())
    }

    fn charge_allocation(
        &mut self,
        _target: AllocationTarget,
        _amount: usize,
    ) -> Result<(), Self::Error> {
        self.events.push("allocation");
        let observed = self.allocations.saturating_add(1);
        if observed > self.allocation_limit {
            return Err(HarnessError::AllocationBudget {
                observed,
                maximum: self.allocation_limit,
            });
        }
        self.allocations = observed;
        Ok(())
    }

    fn map_issue(&mut self, issue: TableCellIssue) -> Self::Error {
        self.mapped_issues.push(issue.clone());
        HarnessError::Issue(issue)
    }
}

#[derive(Debug, Clone, PartialEq)]
struct CellFact {
    row: u32,
    column: u32,
    pointer: usize,
    length: usize,
    value: CellValueSource,
}

#[derive(Debug, Default)]
struct CellSink {
    cells: Vec<CellFact>,
}

impl<Budget> CellValueSink<Budget> for CellSink
where
    Budget: TableCellReadBudget<Error = HarnessError>,
{
    fn visit_cell(
        &mut self,
        cell: CellSource<'_>,
        _budget: &mut Budget,
    ) -> Result<(), HarnessError> {
        self.cells.push(CellFact {
            row: cell.row(),
            column: cell.column(),
            pointer: cell.bytes().as_ptr() as usize,
            length: cell.bytes().len(),
            value: cell.value(),
        });
        Ok(())
    }
}

fn read_one_tile(
    dimensions: TableDimensions,
    tile_index: u32,
    object_id: u64,
    source: &[u8],
    budget: &mut HarnessBudget,
    sink: &mut CellSink,
) -> Result<TableCellReadReport, HarnessError> {
    read_table_cells(
        dimensions,
        [TileReference::new(tile_index, object_id)],
        |resolved_id, _budget| {
            Ok((resolved_id == object_id).then(|| vec![Message::new(TILE_MESSAGE_KIND, source)]))
        },
        budget,
        sink,
    )
}

fn assert_cell_source_in(source: &[u8], fact: &CellFact) {
    assert!(contains_range(source, fact.pointer, fact.length));
}

fn report_facts(report: TableCellReadReport) -> (usize, usize, usize) {
    (report.tiles(), report.rows(), report.cells())
}

#[test]
fn selected_table_read_projects_bnc_and_pre_bnc_rows_without_copying_sources() {
    let empty = bnc_empty();
    let first_number = bnc_number(42.5);
    let mut first_storage = empty.clone();
    first_storage.extend_from_slice(&first_number);
    let first_offsets = narrow_offsets(&[
        Some(0),
        None,
        Some(u16::try_from(empty.len()).expect("small BNC cell")),
        None,
    ]);
    let first_row = row(0, 2, &first_storage, &first_offsets, None, Some(false));

    let second_number = pre_bnc_number(7.25);
    let second_offsets = narrow_offsets(&[Some(0), None, None, None]);
    let second_row = row(1, 1, &second_number, &second_offsets, None, Some(false));
    let first_tile = tile(3, 1, 3, 2, &[first_row, second_row]);

    let third_number = bnc_number(-3.0);
    let third_offsets = narrow_offsets(&[Some(0), None, None, None]);
    let third_row = row(0, 1, &third_number, &third_offsets, None, Some(false));
    let second_tile = tile(3, 3, 1, 1, &[third_row]);

    let mut budget = HarnessBudget::permissive();
    let mut sink = CellSink::default();
    let report = read_table_cells(
        TableDimensions::new(4, 4, 2),
        [TileReference::new(0, 41), TileReference::new(1, 42)],
        |object_id, _budget| {
            Ok(match object_id {
                41 => Some(vec![Message::new(TILE_MESSAGE_KIND, &first_tile)]),
                42 => Some(vec![Message::new(TILE_MESSAGE_KIND, &second_tile)]),
                _ => None,
            })
        },
        &mut budget,
        &mut sink,
    )
    .expect("selected BNC and pre-BNC rows are valid");

    assert_eq!(report_facts(report), (2, 3, 4));
    assert_eq!(
        sink.cells
            .iter()
            .map(|cell| (cell.row, cell.column))
            .collect::<Vec<_>>(),
        vec![(0, 0), (0, 2), (1, 0), (2, 0)]
    );
    assert_eq!(sink.cells[0].value.value, ValueSource::Empty);
    assert_eq!(
        sink.cells[1].value.value,
        ValueSource::Number(
            litchi_iwa_common::formula::FiniteF64::new(42.5).expect("finite test value")
        )
    );
    assert_eq!(
        sink.cells[2].value.value,
        ValueSource::Number(
            litchi_iwa_common::formula::FiniteF64::new(7.25).expect("finite test value")
        )
    );
    assert_eq!(
        sink.cells[3].value.value,
        ValueSource::Number(
            litchi_iwa_common::formula::FiniteF64::new(-3.0).expect("finite test value")
        )
    );
    for (cell, source) in sink.cells[..2].iter().zip([&first_tile, &first_tile]) {
        assert_cell_source_in(source, cell);
    }
    for cell in &sink.cells[2..3] {
        assert_cell_source_in(&first_tile, cell);
    }
    assert_cell_source_in(&second_tile, &sink.cells[3]);
    assert_eq!(sink.cells.iter().filter(|cell| cell.column == 1).count(), 0);
}

#[test]
fn duplicate_payloads_and_duplicate_rows_are_structural_failures() {
    let storage = bnc_number(1.0);
    let offsets = narrow_offsets(&[Some(0), None]);
    let one_row = row(0, 1, &storage, &offsets, None, Some(false));
    let source = tile(1, 1, 1, 1, std::slice::from_ref(&one_row));

    let mut budget = HarnessBudget::permissive();
    let mut sink = CellSink::default();
    let duplicate_payload = read_table_cells(
        TableDimensions::new(2, 2, 2),
        [TileReference::new(0, 90)],
        |_, _budget| {
            Ok(Some(vec![
                Message::new(TILE_MESSAGE_KIND, &source),
                Message::new(TILE_MESSAGE_KIND, &source),
            ]))
        },
        &mut budget,
        &mut sink,
    )
    .expect_err("duplicate tile payloads must be rejected");
    assert_eq!(
        duplicate_payload,
        HarnessError::Issue(TableCellIssue::DuplicateTilePayload { object_id: 90 })
    );
    assert!(sink.cells.is_empty());

    let duplicate_row = tile(1, 1, 1, 2, &[one_row.clone(), one_row]);
    let mut budget = HarnessBudget::permissive();
    let mut sink = CellSink::default();
    let error = read_one_tile(
        TableDimensions::new(2, 2, 2),
        0,
        91,
        &duplicate_row,
        &mut budget,
        &mut sink,
    )
    .expect_err("duplicate local row must be rejected");
    assert_eq!(
        error,
        HarnessError::Issue(TableCellIssue::DuplicateTileRow {
            tile_index: 0,
            row_index: 0,
        })
    );
    // The first callback is only a staged prefix. The selected-table owner
    // must discard it when the read returns the structural error.
    assert_eq!(sink.cells.len(), 1);
}

#[test]
fn unresolved_tiles_and_objects_without_tile_payloads_fail_before_row_decode() {
    let source = tile(1, 0, 0, 0, &[]);
    let mut budget = HarnessBudget::permissive();
    let mut sink = CellSink::default();
    let missing = read_table_cells(
        TableDimensions::new(1, 1, 1),
        [TileReference::new(0, 70)],
        |_, _budget| Ok(None::<Vec<Message<'_>>>),
        &mut budget,
        &mut sink,
    )
    .expect_err("an unresolved selected tile must fail");
    assert_eq!(
        missing,
        HarnessError::Issue(TableCellIssue::MissingTile { object_id: 70 })
    );
    assert!(sink.cells.is_empty());

    let mut budget = HarnessBudget::permissive();
    let mut sink = CellSink::default();
    let missing_payload = read_table_cells(
        TableDimensions::new(1, 1, 1),
        [TileReference::new(0, 71)],
        |_, _budget| Ok(Some(vec![Message::new(9999, &source)])),
        &mut budget,
        &mut sink,
    )
    .expect_err("an object without a type-6002 payload must fail");
    assert_eq!(
        missing_payload,
        HarnessError::Issue(TableCellIssue::MissingTilePayload { object_id: 71 })
    );
    assert!(sink.cells.is_empty());
}

#[test]
fn selected_tile_and_row_bounds_fail_with_typed_issues() {
    let storage = bnc_number(1.0);
    let offsets = narrow_offsets(&[Some(0), None]);
    let valid_row = row(0, 1, &storage, &offsets, None, Some(false));
    let valid_tile = tile(1, 0, 1, 1, std::slice::from_ref(&valid_row));

    let mut budget = HarnessBudget::permissive();
    let mut sink = CellSink::default();
    assert_eq!(
        read_one_tile(
            TableDimensions::new(2, 2, 2),
            1,
            1,
            &valid_tile,
            &mut budget,
            &mut sink,
        )
        .unwrap_err(),
        HarnessError::Issue(TableCellIssue::TileIndexOutOfBounds {
            tile_index: 1,
            tile_count: 1,
        })
    );

    let mut budget = HarnessBudget::permissive();
    let mut sink = CellSink::default();
    let report = read_one_tile(
        TableDimensions::new(2, 2, 2),
        0,
        2,
        &valid_tile,
        &mut budget,
        &mut sink,
    )
    .expect("valid tile is admitted");
    assert_eq!(report_facts(report), (1, 1, 1));

    let out_of_tile_row = row(2, 1, &storage, &offsets, None, Some(false));
    let out_of_tile = tile(1, 1, 1, 1, std::slice::from_ref(&out_of_tile_row));
    let mut budget = HarnessBudget::permissive();
    let mut sink = CellSink::default();
    assert_eq!(
        read_one_tile(
            TableDimensions::new(2, 2, 2),
            0,
            3,
            &out_of_tile,
            &mut budget,
            &mut sink,
        )
        .unwrap_err(),
        HarnessError::Issue(TableCellIssue::TileRowOutOfBounds {
            tile_index: 0,
            row_index: 2,
            tile_size: 2,
        })
    );
    assert!(sink.cells.is_empty());

    let out_of_table_row = row(1, 1, &storage, &offsets, None, Some(false));
    let out_of_table = tile(1, 1, 1, 1, std::slice::from_ref(&out_of_table_row));
    let mut budget = HarnessBudget::permissive();
    let mut sink = CellSink::default();
    assert_eq!(
        read_one_tile(
            TableDimensions::new(1, 2, 2),
            0,
            4,
            &out_of_table,
            &mut budget,
            &mut sink,
        )
        .unwrap_err(),
        HarnessError::Issue(TableCellIssue::TableRowOutOfBounds { row: 1, rows: 1 })
    );

    // Tile summary counters and maxima are producer metadata. The selected
    // row records are the authoritative sparse projection, so stale summary
    // values remain readable when all admitted rows and spans are valid.
    let stale_summary = tile(u32::MAX, u32::MAX, 99, 99, &[valid_row]);
    let mut budget = HarnessBudget::permissive();
    let mut sink = CellSink::default();
    let report = read_one_tile(
        TableDimensions::new(2, 2, 2),
        0,
        5,
        &stale_summary,
        &mut budget,
        &mut sink,
    )
    .expect("stale tile summary does not invalidate valid row records");
    assert_eq!(report_facts(report), (1, 1, 1));
}

#[test]
fn malformed_later_row_and_cell_do_not_become_successful_prefixes() {
    let storage = bnc_number(1.0);
    let offsets = narrow_offsets(&[Some(0), None]);
    let first = row(0, 1, &storage, &offsets, None, Some(false));
    let mut malformed_row = row(1, 1, &storage, &offsets, None, Some(false));
    malformed_row.extend(field_varint(1, 99));
    let source = tile(1, 1, 2, 2, &[first, malformed_row]);
    let before = source.clone();
    let mut budget = HarnessBudget::permissive();
    let mut sink = CellSink::default();
    let error = read_one_tile(
        TableDimensions::new(2, 2, 2),
        0,
        51,
        &source,
        &mut budget,
        &mut sink,
    )
    .expect_err("later malformed row must fail the selected read");
    assert!(matches!(
        error,
        HarnessError::Issue(TableCellIssue::StorageDecode(_))
    ));
    assert_eq!(source, before);
    assert_eq!(sink.cells.len(), 1);

    let malformed_cell_storage = vec![0xff];
    let malformed_cell_offsets = narrow_offsets(&[Some(0), None]);
    let malformed_cell_row = row(
        0,
        1,
        &malformed_cell_storage,
        &malformed_cell_offsets,
        None,
        Some(false),
    );
    let malformed_cell_tile = tile(1, 0, 1, 1, &[malformed_cell_row]);
    let mut budget = HarnessBudget::permissive();
    let mut sink = CellSink::default();
    let error = read_one_tile(
        TableDimensions::new(2, 2, 2),
        0,
        52,
        &malformed_cell_tile,
        &mut budget,
        &mut sink,
    )
    .expect_err("malformed cell payload must fail the selected read");
    assert!(matches!(
        error,
        HarnessError::Issue(TableCellIssue::CellValueDecode {
            row: 0,
            column: 0,
            ..
        })
    ));
    assert!(sink.cells.is_empty());
}

#[test]
fn aggregate_limits_precede_cell_callbacks_and_allocation() {
    let storage = bnc_number(5.0);
    let offsets = narrow_offsets(&[Some(0), None]);
    let row_source = row(0, 1, &storage, &offsets, None, Some(false));
    let source = tile(1, 0, 1, 1, std::slice::from_ref(&row_source));

    let mut permissive_budget = HarnessBudget::permissive();
    let mut permissive_sink = CellSink::default();
    let baseline = read_one_tile(
        TableDimensions::new(1, 2, 1),
        0,
        60,
        &source,
        &mut permissive_budget,
        &mut permissive_sink,
    )
    .expect("baseline selected read");
    assert_eq!(report_facts(baseline), (1, 1, 1));
    assert_eq!(permissive_sink.cells.len(), 1);
    assert!(permissive_budget.storage_work > 0);

    let mut limited_budget = HarnessBudget::permissive();
    limited_budget.cell_count_limit = 0;
    let mut limited_sink = CellSink::default();
    let error = read_one_tile(
        TableDimensions::new(1, 2, 1),
        0,
        60,
        &source,
        &mut limited_budget,
        &mut limited_sink,
    )
    .expect_err("cell count limit must precede source and sink callbacks");
    assert!(matches!(
        error,
        HarnessError::CellBudget {
            observed: 1,
            maximum: 0
        }
    ));
    assert!(limited_sink.cells.is_empty());
    let cells_event = limited_budget
        .events
        .iter()
        .position(|event| *event == "cells")
        .expect("cell ledger event");
    assert!(!limited_budget.events[cells_event..].contains(&"source"));
    assert!(!limited_budget.events[cells_event..].contains(&"allocation"));

    let mut limited_budget = HarnessBudget::permissive();
    limited_budget.allocation_limit = 0;
    let mut limited_sink = CellSink::default();
    let error = read_one_tile(
        TableDimensions::new(1, 2, 1),
        0,
        60,
        &source,
        &mut limited_budget,
        &mut limited_sink,
    )
    .expect_err("row identity allocation must be charged before it is attempted");
    assert_eq!(
        error,
        HarnessError::AllocationBudget {
            observed: 1,
            maximum: 0,
        }
    );
    assert!(limited_sink.cells.is_empty());
    assert_eq!(limited_budget.events, ["options", "allocation", "storage"]);

    let mut limited_budget = HarnessBudget::permissive();
    limited_budget.cell_source_limit = 0;
    let mut limited_sink = CellSink::default();
    let error = read_one_tile(
        TableDimensions::new(1, 2, 1),
        0,
        60,
        &source,
        &mut limited_budget,
        &mut limited_sink,
    )
    .expect_err("cell payload source charge must be authoritative");
    assert!(matches!(error, HarnessError::CellBudget { observed, maximum: 0 } if observed > 0));
    assert!(limited_sink.cells.is_empty());
}

#[test]
fn materialized_cell_cap_is_cumulative_across_tiles() {
    let first_storage = bnc_number(1.0);
    let first_offsets = narrow_offsets(&[Some(0), None]);
    let first_row = row(0, 1, &first_storage, &first_offsets, None, Some(false));
    let first_tile = tile(0, 0, 1, 1, std::slice::from_ref(&first_row));

    let second_storage = bnc_number(2.0);
    let second_offsets = narrow_offsets(&[Some(0), None]);
    let second_row = row(0, 1, &second_storage, &second_offsets, None, Some(false));
    let second_tile = tile(0, 0, 1, 1, std::slice::from_ref(&second_row));

    let mut budget = HarnessBudget::permissive();
    budget.cell_count_limit = 1;
    let mut sink = CellSink::default();
    let error = read_table_cells(
        TableDimensions::new(2, 1, 1),
        [TileReference::new(0, 101), TileReference::new(1, 102)],
        |object_id, _budget| {
            Ok(match object_id {
                101 => Some(vec![Message::new(TILE_MESSAGE_KIND, &first_tile)]),
                102 => Some(vec![Message::new(TILE_MESSAGE_KIND, &second_tile)]),
                _ => None,
            })
        },
        &mut budget,
        &mut sink,
    )
    .expect_err("the materialized-cell cap spans all selected tiles");

    assert_eq!(
        error,
        HarnessError::CellBudget {
            observed: 2,
            maximum: 1,
        }
    );
    assert_eq!(sink.cells.len(), 1);
    assert_eq!(sink.cells[0].row, 0);
    assert_eq!(budget.cell_count, 1);
}

#[test]
fn declared_row_cell_count_is_admitted_before_sparse_spans_or_callbacks() {
    let storage = bnc_number(3.0);
    let offsets = narrow_offsets(&[Some(0), None]);
    // Only one span is physically present, but the row declares two cells.
    // The declared count is the conservative admission charge and must be
    // checked before the reader asks the offset codec for spans.
    let row_source = row(0, 2, &storage, &offsets, None, Some(false));
    let source = tile(0, 0, 2, 1, std::slice::from_ref(&row_source));

    let mut budget = HarnessBudget::permissive();
    budget.cell_count_limit = 1;
    let mut sink = CellSink::default();
    let error = read_one_tile(
        TableDimensions::new(1, 1, 1),
        0,
        103,
        &source,
        &mut budget,
        &mut sink,
    )
    .expect_err("declared row cells must be admitted before sparse decoding");

    assert_eq!(
        error,
        HarnessError::CellBudget {
            observed: 2,
            maximum: 1,
        }
    );
    assert!(sink.cells.is_empty());
    let cells_event = budget
        .events
        .iter()
        .position(|event| *event == "cells")
        .expect("declared-cell admission event");
    assert!(!budget.events[cells_event..].contains(&"source"));
    assert_eq!(budget.events, ["options", "allocation", "cells", "storage"]);
}

#[test]
fn bnc_and_pre_bnc_rows_keep_sparse_missing_and_stored_empty_distinct() {
    let bnc_empty = bnc_empty();
    let bnc_number = bnc_number(42.5);
    let mut bnc_storage = bnc_empty.clone();
    bnc_storage.extend_from_slice(&bnc_number);
    let bnc_offsets = narrow_offsets(&[
        Some(0),
        None,
        Some(u16::try_from(bnc_empty.len()).expect("small BNC cell")),
        None,
    ]);
    let bnc_row = row(0, 2, &bnc_storage, &bnc_offsets, None, Some(false));
    let (bnc_snapshot, _) = decode_tile_row_info_with_report(&bnc_row, options(&bnc_row))
        .expect("handwritten BNC row is valid");
    let (bnc_spans, _) = bnc_snapshot
        .cell_spans(4, options(&bnc_row))
        .expect("BNC sparse offsets are valid");
    assert_eq!(bnc_spans.len(), 2);
    assert!(bnc_spans.get(1).is_none(), "missing slot must stay absent");
    let bnc_storage_active = bnc_snapshot.cell_storage_and_offsets().0;
    let spans: Vec<_> = bnc_spans.iter().collect();
    assert_eq!(spans[0].column(), 0);
    assert_eq!(spans[2 - 1].column(), 2);
    assert_eq!(
        BncCellView::parse(span_bytes(spans[0], bnc_storage_active))
            .expect("empty BNC payload")
            .stored_value(),
        StoredValue::Empty
    );
    assert_eq!(
        BncCellView::parse(span_bytes(spans[1], bnc_storage_active))
            .expect("number BNC payload")
            .stored_value(),
        StoredValue::Number
    );

    let pre_empty = pre_bnc_empty(4);
    let pre_offsets = narrow_offsets(&[Some(0), None]);
    let pre_row = row(1, 1, &pre_empty, &pre_offsets, None, Some(false));
    let pre_snapshot = decode_tile_row_info(&pre_row, options(&pre_row))
        .expect("handwritten pre-BNC row is valid");
    let (pre_spans, _) = pre_snapshot
        .cell_spans(2, options(&pre_row))
        .expect("pre-BNC sparse offsets are valid");
    let pre_storage_active = pre_snapshot.cell_storage_and_offsets().0;
    assert_eq!(pre_spans.len(), 1);
    assert_eq!(pre_spans.get(1), None);
    assert_eq!(
        PreBncCellView::parse(span_bytes(
            pre_spans.get(0).expect("column zero"),
            pre_storage_active
        ))
        .expect("empty pre-BNC payload")
        .cell_type(),
        0
    );
}

#[test]
fn modern_pair_is_selected_only_when_both_buffers_are_present() {
    let pre = pre_bnc_empty(4);
    let pre_offsets = narrow_offsets(&[Some(0), None]);
    let modern = bnc_number(7.25);
    let modern_offsets = narrow_offsets(&[Some(0), None]);

    let complete = row(
        3,
        1,
        &pre,
        &pre_offsets,
        Some((&modern, &modern_offsets)),
        Some(false),
    );
    let complete_snapshot =
        decode_tile_row_info(&complete, options(&complete)).expect("complete modern pair is valid");
    let (storage, offsets) = complete_snapshot.cell_storage_and_offsets();
    assert_eq!(storage, modern.as_slice());
    assert_eq!(offsets, modern_offsets.as_slice());
    let span = complete_snapshot
        .cell_spans(2, options(&complete))
        .expect("complete pair spans")
        .0
        .get(0)
        .expect("column zero");
    assert_eq!(
        BncCellView::parse(span_bytes(span, storage))
            .expect("modern BNC number")
            .stored_value(),
        StoredValue::Number
    );

    let only_storage = {
        let mut output = field_varint(1, 4);
        output.extend(field_varint(2, 1));
        output.extend(field_bytes(3, &pre));
        output.extend(field_bytes(4, &pre_offsets));
        output.extend(field_bytes(6, &modern));
        output
    };
    let fallback = decode_tile_row_info(&only_storage, options(&only_storage))
        .expect("one-sided modern pair falls back to pre-BNC");
    let (storage, offsets) = fallback.cell_storage_and_offsets();
    assert_eq!(storage, pre.as_slice());
    assert_eq!(offsets, pre_offsets.as_slice());
    let span = fallback
        .cell_spans(2, options(&only_storage))
        .expect("fallback spans")
        .0
        .get(0)
        .expect("column zero");
    assert_eq!(
        PreBncCellView::parse(span_bytes(span, storage))
            .expect("pre-BNC empty")
            .version(),
        4
    );
}

#[test]
fn tile_rows_are_borrowed_and_duplicate_indices_remain_visible_to_selection() {
    let first_storage = bnc_number(1.0);
    let first_offsets = narrow_offsets(&[Some(0), None]);
    let second_storage = bnc_number(2.0);
    let second_offsets = narrow_offsets(&[Some(0), None]);
    let first = row(2, 1, &first_storage, &first_offsets, None, Some(false));
    let second = row(2, 1, &second_storage, &second_offsets, None, Some(false));
    let source = tile(2, 4, 2, 4, &[first, second]);
    let before = source.clone();
    let mut collector = RowCollector::default();
    let (snapshot, report) = decode_tile_with_visitor(&source, options(&source), &mut collector)
        .expect("strict tile envelope is valid");
    assert_eq!(source, before);
    assert_eq!((snapshot.max_column(), snapshot.max_row()), (2, 4));
    assert_eq!(collector.rows.len(), 2);
    assert_eq!(collector.rows[0].index, collector.rows[1].index);
    assert_eq!(collector.rows[0].cell_count, 1);
    assert!(report.fields() > 0);
    for fact in &collector.rows {
        assert_borrowed(&source, fact);
    }
    // The strict codec preserves source order and does not silently deduplicate
    // row indices. The selected-table coordinator must reject this duplicate
    // while still receiving both fully validated candidates.
}

#[test]
fn malformed_later_row_invalidates_the_tile_after_a_valid_callback_prefix() {
    let storage = bnc_number(1.0);
    let offsets = narrow_offsets(&[Some(0), None]);
    let first = row(0, 1, &storage, &offsets, None, Some(false));
    let mut malformed_later = row(1, 1, &storage, &offsets, None, Some(false));
    malformed_later.extend(field_varint(1, 99));
    let source = tile(2, 2, 2, 2, &[first, malformed_later]);
    let before = source.clone();
    let mut collector = RowCollector::default();
    assert!(decode_tile_with_visitor(&source, options(&source), &mut collector).is_err());
    assert_eq!(source, before);
    assert_eq!(collector.rows.len(), 1);
    assert_eq!(collector.rows[0].index, 0);
}

#[test]
fn sparse_offset_validation_rejects_later_bad_slots_before_iteration() {
    let storage = b"payload";
    let valid_prefix = narrow_offsets(&[Some(0), None]);
    let (spans, _) = litchi_iwa_protos::numbers_table_cell_storage_codec::CellSpans::parse(
        &valid_prefix,
        storage.len(),
        false,
        1,
        2,
        options(&valid_prefix),
    )
    .expect("valid sparse offsets");
    assert_eq!(
        spans.iter().next().expect("one span").range(),
        0..storage.len()
    );

    let cases: [(&str, Vec<u8>, usize, usize, usize); 6] = [
        ("odd byte count", vec![0, 0, 0], storage.len(), 1, 2),
        (
            "present padding",
            narrow_offsets(&[Some(0), None, Some(1)]),
            2,
            2,
            2,
        ),
        ("unsorted", narrow_offsets(&[Some(2), Some(1)]), 4, 2, 2),
        (
            "duplicate start",
            narrow_offsets(&[Some(0), Some(0)]),
            4,
            2,
            2,
        ),
        (
            "out of bounds",
            narrow_offsets(&[Some(7), None]),
            storage.len(),
            1,
            2,
        ),
        ("wrong expected count", valid_prefix, storage.len(), 0, 2),
    ];
    for (label, offsets, storage_length, expected, columns) in cases {
        let result = litchi_iwa_protos::numbers_table_cell_storage_codec::CellSpans::parse(
            &offsets,
            storage_length,
            false,
            expected,
            columns,
            options(&offsets),
        );
        assert!(result.is_err(), "{label} unexpectedly decoded");
    }

    let wide_out_of_bounds = narrow_offsets(&[Some(2), None]);
    assert!(
        litchi_iwa_protos::numbers_table_cell_storage_codec::CellSpans::parse(
            &wide_out_of_bounds,
            8,
            true,
            1,
            2,
            options(&wide_out_of_bounds),
        )
        .is_err()
    );
}

#[test]
fn row_and_tile_bounds_are_checked_without_changing_the_source() {
    let storage = bnc_number(3.0);
    let offsets = narrow_offsets(&[Some(0), None]);
    let valid = row(4, 1, &storage, &offsets, None, Some(false));
    let mut duplicate_index = valid.clone();
    duplicate_index.extend(field_varint(1, 5));
    assert!(decode_tile_row_info(&duplicate_index, options(&duplicate_index)).is_err());

    let too_many_cells = row(0, 2, &storage, &offsets, None, Some(false));
    let before = too_many_cells.clone();
    let result = decode_tile_row_info(&too_many_cells, options(&too_many_cells));
    assert!(result.is_ok(), "row envelope does not own table width");
    assert_eq!(too_many_cells, before);

    let bad_tile = tile(1, 1, 1, 1, &[valid]);
    let before = bad_tile.clone();
    let mut collector = RowCollector::default();
    decode_tile_with_visitor(&bad_tile, options(&bad_tile), &mut collector)
        .expect("tile bounds belong to the selected-table layer");
    assert_eq!(bad_tile, before);
    assert_eq!(collector.rows[0].index, 4);
}

#[test]
fn aggregate_work_limit_is_checked_at_the_strict_decode_boundary() {
    let storage = bnc_number(5.0);
    let offsets = narrow_offsets(&[Some(0), None]);
    let source = tile(
        2,
        2,
        1,
        2,
        &[row(0, 1, &storage, &offsets, None, Some(false))],
    );
    let (_, report) =
        decode_tile_with_report(&source, options(&source)).expect("valid tile report");
    let exact = DecodeOptions::new(
        source.len(),
        report.fields(),
        report.work_bytes(),
        report.max_depth(),
        report.references(),
        report.text_bytes(),
    );
    decode_tile_with_visitor(&source, exact, &mut RowCollector::default())
        .expect("inclusive exact aggregate budget");

    let mut collector = RowCollector::default();
    let error = decode_tile_with_visitor(
        &source,
        DecodeOptions::new(
            source.len(),
            report.fields(),
            report.work_bytes().saturating_sub(1),
            report.max_depth(),
            report.references(),
            report.text_bytes(),
        ),
        &mut collector,
    )
    .expect_err("one less work byte must fail");
    assert!(matches!(
        error.resource_limit(),
        Some(DecodeLimit::Work { observed, maximum })
            if observed > maximum && maximum + 1 == report.work_bytes()
    ));
    // No successful projection may publish a partial tile after the aggregate
    // strict pass fails. A streaming codec can observe a prefix internally,
    // but its caller must stage it until this result is Ok.
}
