//! Differential coverage for the source-backed XLS worksheet occurrence index.
//!
//! These tests deliberately observe the public query behavior through a
//! positional source.  The cache is an optional optimization: values,
//! duplicate order, refusal identity, freshness fences, and cancellation must
//! remain the same when admission is disabled, refused, or evicted.

#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "fixture construction and test assertions intentionally panic"
)]

use litchi_cfb::{OleWriter, SharedOleFile};
use litchi_core::sheet::CellValue;
use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionError, ExecutionLimits, Limits,
    OwnedSource, ReadAt, Resource, SourceVersion,
};
use litchi_xls::{SourceBackedError, SourceBackedLimits, SourceBackedWorkbook};
use std::io::{self, Cursor};
use std::num::{NonZeroU64, NonZeroUsize};
use std::sync::atomic::{AtomicU64, Ordering};
use std::sync::{Arc, Mutex};

fn fixture(name: &str) -> Vec<u8> {
    std::fs::read(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/poi/test-data/spreadsheet")
            .join(name),
    )
    .unwrap()
}

fn workbook_stream(bytes: Vec<u8>) -> Vec<u8> {
    SharedOleFile::open(Arc::new(OwnedSource::new(bytes)))
        .unwrap()
        .open_stream(&["Workbook"])
        .unwrap()
}

fn cfb_with_workbook(stream: &[u8]) -> Vec<u8> {
    let mut writer = OleWriter::new();
    writer.create_stream(&["Workbook"], stream).unwrap();
    let mut output = Vec::new();
    writer.write_to(&mut Cursor::new(&mut output)).unwrap();
    output
}

fn frame_bytes(kind: u16, payload: &[u8]) -> Vec<u8> {
    let mut output = Vec::with_capacity(4 + payload.len());
    output.extend_from_slice(&kind.to_le_bytes());
    output.extend_from_slice(&(u16::try_from(payload.len()).unwrap()).to_le_bytes());
    output.extend_from_slice(payload);
    output
}

fn number_frame(row: u16, column: u16, value: f64) -> Vec<u8> {
    let mut payload = Vec::with_capacity(14);
    payload.extend_from_slice(&row.to_le_bytes());
    payload.extend_from_slice(&column.to_le_bytes());
    payload.extend_from_slice(&0_u16.to_le_bytes());
    payload.extend_from_slice(&value.to_le_bytes());
    frame_bytes(0x0203, &payload)
}

fn label_sst_frame(row: u16, column: u16, string_index: u32) -> Vec<u8> {
    let mut payload = Vec::with_capacity(10);
    payload.extend_from_slice(&row.to_le_bytes());
    payload.extend_from_slice(&column.to_le_bytes());
    payload.extend_from_slice(&0_u16.to_le_bytes());
    payload.extend_from_slice(&string_index.to_le_bytes());
    frame_bytes(0x00FD, &payload)
}

fn first_sheet_offset(stream: &[u8]) -> usize {
    let mut offset = 0;
    while offset + 4 <= stream.len() {
        let kind = u16::from_le_bytes([stream[offset], stream[offset + 1]]);
        let length = usize::from(u16::from_le_bytes([stream[offset + 2], stream[offset + 3]]));
        if kind == 0x0085 {
            return usize::try_from(u32::from_le_bytes([
                stream[offset + 4],
                stream[offset + 5],
                stream[offset + 6],
                stream[offset + 7],
            ]))
            .unwrap();
        }
        offset += 4 + length;
    }
    panic!("workbook has no BoundSheet8");
}

fn global_eof_offset(stream: &[u8]) -> usize {
    let mut offset = 0;
    loop {
        assert!(offset + 4 <= stream.len());
        let kind = u16::from_le_bytes([stream[offset], stream[offset + 1]]);
        let length = usize::from(u16::from_le_bytes([stream[offset + 2], stream[offset + 3]]));
        if kind == 0x000A {
            return offset;
        }
        offset += 4 + length;
    }
}

fn remove_global_sst(stream: &[u8]) -> Vec<u8> {
    let eof = global_eof_offset(stream);
    let mut output = Vec::with_capacity(stream.len());
    let mut offset = 0;
    while offset < eof {
        assert!(offset + 4 <= stream.len());
        let kind = u16::from_le_bytes([stream[offset], stream[offset + 1]]);
        let length = usize::from(u16::from_le_bytes([stream[offset + 2], stream[offset + 3]]));
        let end = offset + 4 + length;
        if kind == 0x00FC {
            // SST CONTINUE records belong to the removed SST and are
            // contiguous in the globals record sequence.
            offset = end;
            while offset < eof {
                assert!(offset + 4 <= stream.len());
                let next_kind = u16::from_le_bytes([stream[offset], stream[offset + 1]]);
                if next_kind != 0x003C {
                    break;
                }
                let next_length =
                    usize::from(u16::from_le_bytes([stream[offset + 2], stream[offset + 3]]));
                offset += 4 + next_length;
            }
        } else {
            output.extend_from_slice(&stream[offset..end]);
            offset = end;
        }
    }
    output.extend_from_slice(&stream[eof..]);

    // BoundSheet8 stores absolute workbook-stream positions.  Removing the
    // SST shifts every worksheet start by the same negative delta.
    let new_eof = global_eof_offset(&output);
    let delta = i64::try_from(output.len()).unwrap() - i64::try_from(stream.len()).unwrap();
    let mut cursor = 0;
    while cursor < new_eof {
        assert!(cursor + 4 <= output.len());
        let kind = u16::from_le_bytes([output[cursor], output[cursor + 1]]);
        let length = usize::from(u16::from_le_bytes([output[cursor + 2], output[cursor + 3]]));
        if kind == 0x0085 {
            assert!(length >= 8);
            let position = u32::from_le_bytes([
                output[cursor + 4],
                output[cursor + 5],
                output[cursor + 6],
                output[cursor + 7],
            ]);
            let shifted = i64::from(position) + delta;
            assert!(shifted >= 0);
            output[cursor + 4..cursor + 8]
                .copy_from_slice(&u32::try_from(shifted).unwrap().to_le_bytes());
        }
        cursor += 4 + length;
    }
    output
}

fn worksheet_eof_offset(stream: &[u8], start: usize) -> usize {
    let mut offset = start;
    loop {
        assert!(offset + 4 <= stream.len());
        let kind = u16::from_le_bytes([stream[offset], stream[offset + 1]]);
        let length = usize::from(u16::from_le_bytes([stream[offset + 2], stream[offset + 3]]));
        if kind == 0x000A {
            return offset;
        }
        offset += 4 + length;
    }
}

fn patch_later_sheet_offsets(
    output: &mut [u8],
    first_sheet_start: usize,
    insertion_boundary: usize,
    extra_len: usize,
) {
    let delta = u32::try_from(extra_len).unwrap();
    let mut cursor = 0;
    while cursor < first_sheet_start {
        assert!(cursor + 4 <= output.len());
        let kind = u16::from_le_bytes([output[cursor], output[cursor + 1]]);
        let length = usize::from(u16::from_le_bytes([output[cursor + 2], output[cursor + 3]]));
        if kind == 0x0085 && cursor + 8 <= output.len() {
            let position = u32::from_le_bytes([
                output[cursor + 4],
                output[cursor + 5],
                output[cursor + 6],
                output[cursor + 7],
            ]);
            if usize::try_from(position).unwrap() > insertion_boundary {
                let shifted = position.checked_add(delta).unwrap();
                output[cursor + 4..cursor + 8].copy_from_slice(&shifted.to_le_bytes());
            }
        }
        cursor += 4 + length;
    }
}

fn insert_before_worksheet_eof(stream: &[u8], extra: &[u8]) -> Vec<u8> {
    let eof = worksheet_eof_offset(stream, first_sheet_offset(stream));
    let first_sheet_start = first_sheet_offset(stream);
    let mut output = Vec::with_capacity(stream.len() + extra.len());
    output.extend_from_slice(&stream[..eof]);
    output.extend_from_slice(extra);
    output.extend_from_slice(&stream[eof..]);
    patch_later_sheet_offsets(&mut output, first_sheet_start, eof, extra.len());
    output
}

fn first_frame_of_kind(stream: &[u8], wanted: u16) -> (usize, usize) {
    let mut offset = first_sheet_offset(stream);
    loop {
        assert!(offset + 4 <= stream.len());
        let kind = u16::from_le_bytes([stream[offset], stream[offset + 1]]);
        let length = usize::from(u16::from_le_bytes([stream[offset + 2], stream[offset + 3]]));
        let end = offset + 4 + length;
        if kind == wanted {
            return (offset, end);
        }
        assert_ne!(kind, 0x000A, "worksheet has no requested frame");
        offset = end;
    }
}

/// Inserts after the worksheet BOF and before the first selected frame.  The
/// sheet's BoundSheet8 position does not move because the insertion is inside
/// the worksheet substream.
fn insert_before_first_frame_kind(stream: &[u8], wanted: u16, extra: &[u8]) -> Vec<u8> {
    let (offset, _) = first_frame_of_kind(stream, wanted);
    let first_sheet_start = first_sheet_offset(stream);
    let mut output = Vec::with_capacity(stream.len() + extra.len());
    output.extend_from_slice(&stream[..offset]);
    output.extend_from_slice(extra);
    output.extend_from_slice(&stream[offset..]);
    patch_later_sheet_offsets(&mut output, first_sheet_start, offset, extra.len());
    output
}

fn large_numeric_workbook(count: u16) -> Vec<u8> {
    let original = workbook_stream(fixture("Simple.xls"));
    let mut extra = Vec::new();
    for row in 1..=count {
        extra.extend_from_slice(&number_frame(row, 0, f64::from(row)));
    }
    cfb_with_workbook(&insert_before_worksheet_eof(&original, &extra))
}

fn large_malformed_formula_workbook(count: u16) -> Vec<u8> {
    let original = workbook_stream(fixture("Simple.xls"));
    let mut extra = Vec::new();
    // This target-specific UTF-16 refusal is invisible while a different
    // numeric target is selected, so a successful scan can learn that the
    // worksheet index is intrinsically too large before this target is read.
    extra.extend_from_slice(&formula_frame(50_000, 4));
    extra.extend_from_slice(&formula_string_frame(&[1, 0, 1, 0x00, 0xD8]));
    for row in 1..=count {
        extra.extend_from_slice(&number_frame(row, 0, f64::from(row)));
    }
    cfb_with_workbook(&insert_before_worksheet_eof(&original, &extra))
}

fn large_malformed_tail_workbook(count: u16) -> Vec<u8> {
    let original = workbook_stream(fixture("Simple.xls"));
    let mut extra = Vec::new();
    for row in 1..=count {
        extra.extend_from_slice(&number_frame(row, 0, f64::from(row)));
    }
    // The malformed recognized frame is after all valid cells, ensuring that
    // a candidate can never publish a partial index before the refusal.
    extra.extend_from_slice(&frame_bytes(0x0203, &[0, 0, 0, 0, 1, 0]));
    cfb_with_workbook(&insert_before_worksheet_eof(&original, &extra))
}

fn duplicate_out_of_range_workbook() -> Vec<u8> {
    let original = workbook_stream(fixture("Simple.xls"));
    let (row, column) = {
        let (start, end) = first_frame_of_kind(&original, 0x00FD);
        let frame = &original[start..end];
        (
            u16::from_le_bytes([frame[4], frame[5]]),
            u16::from_le_bytes([frame[6], frame[7]]),
        )
    };
    // The existing valid LabelSst follows this inserted occurrence.  The
    // invalid index is a typed CellValue::Error, so the later valid duplicate
    // must overwrite it and must not turn the operation into a refusal.
    let invalid = label_sst_frame(row, column, u32::MAX);
    cfb_with_workbook(&insert_before_first_frame_kind(&original, 0x00FD, &invalid))
}

fn missing_sst_workbook() -> (Vec<u8>, (u16, u16)) {
    let original = workbook_stream(fixture("Simple.xls"));
    let (start, _) = first_frame_of_kind(&original, 0x00FD);
    let row = u16::from_le_bytes([original[start + 4], original[start + 5]]);
    let column = u16::from_le_bytes([original[start + 6], original[start + 7]]);
    (
        cfb_with_workbook(&remove_global_sst(&original)),
        (row, column),
    )
}

fn duplicate_shared_string_workbook() -> Vec<u8> {
    let original = workbook_stream(fixture("Simple.xls"));
    let mut extra = Vec::new();
    // Keep the two equal-coordinate occurrences adjacent and after the
    // ordinary fixture cells.  A different numeric target is used to warm the
    // occurrence index without resolving either shared string.
    extra.extend_from_slice(&label_sst_frame(50_000, 4, 0));
    extra.extend_from_slice(&label_sst_frame(50_000, 4, 0));
    extra.extend_from_slice(&number_frame(50_001, 4, 17.0));
    cfb_with_workbook(&insert_before_worksheet_eof(&original, &extra))
}

fn formula_string_workbook() -> Vec<u8> {
    let original = workbook_stream(fixture("Simple.xls"));
    let mut formula = Vec::new();
    formula.extend_from_slice(&60_005_u16.to_le_bytes());
    formula.extend_from_slice(&6_u16.to_le_bytes());
    formula.extend_from_slice(&0_u16.to_le_bytes());
    formula.extend_from_slice(&[0, 0, 0, 0, 0, 0, 0xFF, 0xFF]);
    formula.extend_from_slice(&0_u16.to_le_bytes());
    formula.extend_from_slice(&0_u32.to_le_bytes());
    formula.extend_from_slice(&0_u16.to_le_bytes());

    let mut string = Vec::new();
    string.extend_from_slice(&9_u16.to_le_bytes());
    string.push(0);
    string.extend_from_slice(b"abc");
    let continuation = |text: &[u8]| {
        let mut payload = vec![0];
        payload.extend_from_slice(text);
        frame_bytes(0x003C, &payload)
    };
    let mut chain = frame_bytes(0x0006, &formula);
    // The BIFF grammar permits these metadata records between the formula and
    // its STRING result.  The following CONTINUE records exercise the second
    // part of the same pending result on both the scan and indexed replay.
    chain.extend_from_slice(&frame_bytes(0x0221, &[]));
    chain.extend_from_slice(&frame_bytes(0x0207, &string));
    chain.extend_from_slice(&continuation(b"def"));
    chain.extend_from_slice(&continuation(b"ghi"));
    cfb_with_workbook(&insert_before_worksheet_eof(&original, &chain))
}

fn formula_frame(row: u16, column: u16) -> Vec<u8> {
    let mut payload = Vec::new();
    payload.extend_from_slice(&row.to_le_bytes());
    payload.extend_from_slice(&column.to_le_bytes());
    payload.extend_from_slice(&0_u16.to_le_bytes());
    payload.extend_from_slice(&[0, 0, 0, 0, 0, 0, 0xFF, 0xFF]);
    payload.extend_from_slice(&0_u16.to_le_bytes());
    payload.extend_from_slice(&0_u32.to_le_bytes());
    payload.extend_from_slice(&0_u16.to_le_bytes());
    frame_bytes(0x0006, &payload)
}

fn formula_string_frame(payload: &[u8]) -> Vec<u8> {
    frame_bytes(0x0207, payload)
}

fn malformed_formula_duplicate_workbook() -> Vec<u8> {
    let original = workbook_stream(fixture("Simple.xls"));
    let mut extra = Vec::new();
    // The first result has a lone UTF-16 high surrogate.  Its BIFF framing is
    // valid, but decoding the selected formula result must refuse it.
    extra.extend_from_slice(&formula_frame(50_000, 4));
    extra.extend_from_slice(&formula_string_frame(&[1, 0, 1, 0x00, 0xD8]));
    // A later equal-coordinate formula is valid and would hide the earlier
    // failure if a cache retained only the latest occurrence.
    extra.extend_from_slice(&formula_frame(50_000, 4));
    extra.extend_from_slice(&formula_string_frame(&[2, 0, 0, b'o', b'k']));
    extra.extend_from_slice(&number_frame(50_001, 4, 19.0));
    cfb_with_workbook(&insert_before_worksheet_eof(&original, &extra))
}

fn formula_without_string_workbook() -> Vec<u8> {
    let original = workbook_stream(fixture("Simple.xls"));
    let mut extra = number_frame(50_001, 4, 23.0);
    // Leave the string-valued FORMULA as the final worksheet record.  The
    // ordinary selected-cell path reports the dedicated missing-result error
    // when this coordinate is requested; an index built for another
    // coordinate must retain enough locator state to report that same error.
    extra.extend_from_slice(&formula_frame(50_000, 4));
    cfb_with_workbook(&insert_before_worksheet_eof(&original, &extra))
}

fn malformed_tail_workbook() -> Vec<u8> {
    let original = workbook_stream(fixture("Simple.xls"));
    // A recognized Number frame with a short payload.  It is after all normal
    // cells, so a failed selected query reaches the same tail on every call.
    let malformed = frame_bytes(0x0203, &[0, 0, 0, 0, 1, 0]);
    cfb_with_workbook(&insert_before_worksheet_eof(&original, &malformed))
}

#[derive(Clone)]
struct CountingSource {
    bytes: Arc<Vec<u8>>,
    ranges: Arc<Mutex<Vec<(u64, usize)>>>,
    cancel_on_read: Arc<Mutex<Option<CancellationSource>>>,
    revision: Arc<AtomicU64>,
    observations: Arc<AtomicU64>,
    bump_before: Arc<AtomicU64>,
    bump_after: Arc<AtomicU64>,
    fail_marker: Arc<Mutex<Option<Vec<u8>>>>,
}

impl CountingSource {
    fn new(bytes: Vec<u8>) -> Self {
        Self {
            bytes: Arc::new(bytes),
            ranges: Arc::new(Mutex::new(Vec::new())),
            cancel_on_read: Arc::new(Mutex::new(None)),
            revision: Arc::new(AtomicU64::new(0)),
            observations: Arc::new(AtomicU64::new(0)),
            bump_before: Arc::new(AtomicU64::new(0)),
            bump_after: Arc::new(AtomicU64::new(0)),
            fail_marker: Arc::new(Mutex::new(None)),
        }
    }

    fn clear_measurement(&self) {
        self.ranges.lock().unwrap().clear();
        self.observations.store(0, Ordering::Relaxed);
    }

    fn bytes_read(&self) -> usize {
        self.ranges
            .lock()
            .unwrap()
            .iter()
            .map(|(_, length)| *length)
            .sum()
    }

    fn read_count(&self) -> usize {
        self.ranges.lock().unwrap().len()
    }

    fn cancel_on_next_read(&self, cancellation: CancellationSource) {
        *self.cancel_on_read.lock().unwrap() = Some(cancellation);
    }

    fn bump_after_observation(&self, ordinal: u64) {
        self.bump_after.store(ordinal, Ordering::Relaxed);
    }

    fn bump_before_observation(&self, ordinal: u64) {
        self.bump_before.store(ordinal, Ordering::Relaxed);
    }

    fn bump(&self) {
        self.revision.fetch_add(1, Ordering::Relaxed);
    }

    fn fail_next_read_containing(&self, marker: &[u8]) {
        *self.fail_marker.lock().unwrap() = Some(marker.to_vec());
    }
}

impl ReadAt for CountingSource {
    fn len(&self) -> io::Result<u64> {
        Ok(self.bytes.len() as u64)
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        let start = usize::try_from(offset)
            .map_err(|_| io::Error::new(io::ErrorKind::InvalidInput, "offset overflow"))?;
        if start >= self.bytes.len() || output.is_empty() {
            return Ok(0);
        }
        let count = output.len().min(self.bytes.len() - start);
        let fail = self
            .fail_marker
            .lock()
            .unwrap()
            .as_ref()
            .is_some_and(|marker| {
                !marker.is_empty()
                    && self.bytes[start..start + count]
                        .windows(marker.len())
                        .any(|window| window == marker)
            });
        if fail {
            let _ = self.fail_marker.lock().unwrap().take();
            return Err(io::Error::other("synthetic shared-string read failure"));
        }
        output[..count].copy_from_slice(&self.bytes[start..start + count]);
        self.ranges.lock().unwrap().push((offset, count));
        if let Some(cancellation) = self.cancel_on_read.lock().unwrap().take() {
            cancellation.cancel();
        }
        Ok(count)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        let ordinal = self.observations.fetch_add(1, Ordering::Relaxed) + 1;
        if self.bump_before.load(Ordering::Relaxed) == ordinal {
            self.revision.fetch_add(1, Ordering::Relaxed);
        }
        let version =
            SourceVersion::new(0x584C_535F_43414348, self.revision.load(Ordering::Relaxed));
        if self.bump_after.load(Ordering::Relaxed) == ordinal {
            self.revision.fetch_add(1, Ordering::Relaxed);
        }
        Ok(version)
    }
}

fn measure_query(
    source: &CountingSource,
    worksheet: &litchi_xls::SourceBackedWorksheet,
    row: u32,
    column: u32,
) -> (Option<CellValue>, usize, usize) {
    source.clear_measurement();
    let value = worksheet.cell_value(row, column).unwrap();
    (value, source.bytes_read(), source.read_count())
}

#[test]
fn disabled_cache_preserves_values_and_repeats_the_complete_scan() {
    let source = Arc::new(CountingSource::new(large_numeric_workbook(2_000)));
    let limits = SourceBackedLimits::default().with_max_query_index_bytes(0);
    let owner = SourceBackedWorkbook::from_read_at_with_limits(source.clone(), limits).unwrap();
    let worksheet = owner.worksheet_by_index(0).unwrap().unwrap();

    let (first, first_bytes, first_reads) = measure_query(&source, &worksheet, 1_999, 0);
    let (second, second_bytes, second_reads) = measure_query(&source, &worksheet, 1_999, 0);
    assert_eq!(first, second);
    assert!(first_bytes > 0 && first_reads > 0);
    assert_eq!(
        second_bytes, first_bytes,
        "disabled cache must not shortcut the scan"
    );
    assert_eq!(
        second_reads, first_reads,
        "disabled cache must not alter the read path"
    );
}

#[test]
fn a_successful_warm_index_shortcuts_later_hits_and_missing_cells() {
    let source = Arc::new(CountingSource::new(large_numeric_workbook(4_000)));
    let owner = SourceBackedWorkbook::from_read_at(source.clone()).unwrap();
    let worksheet = owner.worksheet_by_index(0).unwrap().unwrap();

    let (first, first_bytes, first_reads) = measure_query(&source, &worksheet, 3_999, 0);
    let (second, second_bytes, second_reads) = measure_query(&source, &worksheet, 3_999, 0);
    let (third, third_bytes, third_reads) = measure_query(&source, &worksheet, 3_999, 0);
    assert_eq!(first, second);
    assert_eq!(second, third);
    assert!(first_bytes > 0 && first_reads > 0);
    assert!(second_bytes > 0 && second_reads > 0);
    assert!(
        third_bytes < second_bytes / 2,
        "indexed hit should avoid the complete worksheet scan: first={first_bytes}, second={second_bytes}, third={third_bytes}"
    );
    assert!(
        third_reads < second_reads,
        "indexed hit should issue fewer source reads: first={first_reads}, second={second_reads}, third={third_reads}"
    );

    let (missing, missing_bytes, missing_reads) = measure_query(&source, &worksheet, 50_000, 7);
    assert!(missing.is_none());
    assert_eq!(
        missing_bytes, 0,
        "an indexed miss needs only the freshness fence"
    );
    assert_eq!(
        missing_reads, 0,
        "an indexed miss must not read the worksheet again"
    );
    source.bump();
    assert!(matches!(
        worksheet.cell_value(50_000, 7),
        Err(SourceBackedError::SourceChanged { .. })
    ));
}

#[test]
fn a_tiny_cache_budget_falls_back_without_changing_query_values() {
    let source = Arc::new(CountingSource::new(large_numeric_workbook(2_000)));
    let limits = SourceBackedLimits::default().with_max_query_index_bytes(1);
    let owner = SourceBackedWorkbook::from_read_at_with_limits(source.clone(), limits).unwrap();
    let worksheet = owner.worksheet_by_index(0).unwrap().unwrap();

    let (first, first_bytes, first_reads) = measure_query(&source, &worksheet, 1_999, 0);
    let (second, second_bytes, second_reads) = measure_query(&source, &worksheet, 1_999, 0);
    let (third, third_bytes, third_reads) = measure_query(&source, &worksheet, 1_999, 0);
    assert_eq!(first, second);
    assert_eq!(second, third);
    assert!(first_bytes > 0 && first_reads > 0);
    assert!(second_bytes > 0 && second_reads > 0);
    assert!(third_bytes > 0 && third_reads > 0);
}

#[test]
fn an_intrinsically_oversize_index_stays_on_the_cold_path_after_learning() {
    let source = Arc::new(CountingSource::new(large_numeric_workbook(4_000)));
    // Keep the fixed worksheet-cache metadata admissible while making the
    // complete occurrence table provably too large for the immutable cache
    // ceiling.  The first two queries include the successful scan that learns
    // the lower-bound sentinel; later queries must still return the ordinary
    // scan result rather than a partial or error-derived index hit.
    let limits = SourceBackedLimits::default().with_max_query_index_bytes(1_024);
    let owner = SourceBackedWorkbook::from_read_at_with_limits(source.clone(), limits).unwrap();
    let worksheet = owner.worksheet_by_index(0).unwrap().unwrap();

    for _ in 0..4 {
        let (value, bytes, reads) = measure_query(&source, &worksheet, 3_999, 0);
        assert_eq!(value, Some(CellValue::Float(3_999.0)));
        assert!(bytes > 0 && reads > 0, "oversize index must keep scanning");
    }

    for _ in 0..2 {
        let (value, bytes, reads) = measure_query(&source, &worksheet, 50_000, 7);
        assert_eq!(value, None);
        assert!(
            bytes > 0 && reads > 0,
            "an intrinsically oversize index must not turn a missing cell into a warm hit"
        );
    }
}

#[test]
fn an_oversize_index_sentinel_keeps_the_source_version_fence() {
    let source = Arc::new(CountingSource::new(large_numeric_workbook(4_000)));
    let limits = SourceBackedLimits::default().with_max_query_index_bytes(1_024);
    let owner = SourceBackedWorkbook::from_read_at_with_limits(source.clone(), limits).unwrap();
    let worksheet = owner.worksheet_by_index(0).unwrap().unwrap();

    // The second successful scan learns the intrinsic-size sentinel.
    assert_eq!(
        worksheet.cell_value(3_999, 0).unwrap(),
        Some(CellValue::Float(3_999.0))
    );
    assert_eq!(
        worksheet.cell_value(3_999, 0).unwrap(),
        Some(CellValue::Float(3_999.0))
    );

    source.clear_measurement();
    source.bump();
    let error = worksheet.cell_value(3_999, 0).unwrap_err();
    assert!(matches!(error, SourceBackedError::SourceChanged { .. }));
    assert_eq!(
        source.read_count(),
        0,
        "the sentinel must not bypass the leading freshness fence"
    );
}

#[test]
fn an_oversize_index_keeps_target_formula_refusals_source_backed() {
    let source = Arc::new(CountingSource::new(large_malformed_formula_workbook(4_000)));
    let limits = SourceBackedLimits::default().with_max_query_index_bytes(1_024);
    let owner = SourceBackedWorkbook::from_read_at_with_limits(source.clone(), limits).unwrap();
    let worksheet = owner.worksheet_by_index(0).unwrap().unwrap();

    // Selecting a valid numeric cell leaves the malformed formula unselected,
    // so the successful second scan learns the worksheet-size sentinel.
    assert_eq!(
        worksheet.cell_value(3_999, 0).unwrap(),
        Some(CellValue::Float(3_999.0))
    );
    assert_eq!(
        worksheet.cell_value(3_999, 0).unwrap(),
        Some(CellValue::Float(3_999.0))
    );

    for _ in 0..2 {
        source.clear_measurement();
        let error = worksheet.cell_value(50_000, 4).unwrap_err();
        assert!(
            error.to_string().contains("UTF-16 decoding error"),
            "unexpected target-specific formula refusal: {error:?}"
        );
        assert!(
            source.read_count() > 0,
            "the malformed result must be re-read instead of being cached as an error"
        );
    }
}

#[test]
fn a_failed_oversize_scan_does_not_publish_a_partial_index() {
    let source = Arc::new(CountingSource::new(large_malformed_tail_workbook(4_000)));
    let limits = SourceBackedLimits::default().with_max_query_index_bytes(1_024);
    let owner = SourceBackedWorkbook::from_read_at_with_limits(source.clone(), limits).unwrap();
    let worksheet = owner.worksheet_by_index(0).unwrap().unwrap();

    let mut first_error: Option<String> = None;
    let mut first_bytes = None;
    let mut first_reads = None;
    for _ in 0..3 {
        source.clear_measurement();
        let error = worksheet.cell_value(3_999, 0).unwrap_err();
        if let Some(first) = &first_error {
            assert_eq!(error.to_string(), first.as_str());
            assert_eq!(source.bytes_read(), first_bytes.unwrap());
            assert_eq!(source.read_count(), first_reads.unwrap());
        } else {
            first_error = Some(error.to_string());
            first_bytes = Some(source.bytes_read());
            first_reads = Some(source.read_count());
        }
        assert!(source.bytes_read() > 0 && source.read_count() > 0);
    }
}

#[test]
fn an_indexed_missing_target_still_takes_the_trailing_freshness_fence() {
    let source = Arc::new(CountingSource::new(large_numeric_workbook(4_000)));
    let owner = SourceBackedWorkbook::from_read_at(source.clone()).unwrap();
    let worksheet = owner.worksheet_by_index(0).unwrap().unwrap();

    // Publish the worksheet index with two successful observations, then use
    // a coordinate absent from its occurrence table.  The hit has no source
    // read of its own, so the mutation is injected into the trailing version
    // observation rather than relying on the leading fence.
    assert_eq!(
        worksheet.cell_value(3_999, 0).unwrap(),
        Some(CellValue::Float(3_999.0))
    );
    assert_eq!(
        worksheet.cell_value(3_999, 0).unwrap(),
        Some(CellValue::Float(3_999.0))
    );
    source.clear_measurement();
    source.bump_before_observation(2);
    assert!(matches!(
        worksheet.cell_value(50_000, 7),
        Err(SourceBackedError::SourceChanged { .. })
    ));
    assert_eq!(source.read_count(), 0);
}

#[test]
fn zero_half_one_and_two_mib_budgets_keep_results_identical() {
    // Its 38,950 occurrences fit within 1 MiB with bounded final growth,
    // although the former doubling step required 65,536 slots.
    let bytes = fixture("54016.xls");
    let measure = |budget| {
        let source = Arc::new(CountingSource::new(bytes.clone()));
        let limits = SourceBackedLimits::default().with_max_query_index_bytes(budget);
        let owner = SourceBackedWorkbook::from_read_at_with_limits(source.clone(), limits).unwrap();
        let worksheet = owner.worksheet_by_index(0).unwrap().unwrap();
        let (first, _, _) = measure_query(&source, &worksheet, 0, 0);
        let (second, _, _) = measure_query(&source, &worksheet, 0, 0);
        let (third, third_bytes, third_reads) = measure_query(&source, &worksheet, 0, 0);
        (first, second, third, third_bytes, third_reads)
    };

    let (zero_first, zero_second, zero_third, zero_bytes, zero_reads) = measure(0);
    let (half_first, half_second, half_third, half_bytes, half_reads) = measure(1 << 19);
    let (one_first, one_second, one_third, one_bytes, one_reads) = measure(1 << 20);
    let (two_first, two_second, two_third, two_bytes, two_reads) = measure(2 << 20);
    assert_eq!(zero_first, zero_second);
    assert_eq!(zero_second, zero_third);
    assert_eq!(one_first, one_second);
    assert_eq!(one_second, one_third);
    assert_eq!(two_first, two_second);
    assert_eq!(two_second, two_third);

    assert_eq!(half_first, half_second);
    assert_eq!(half_second, half_third);
    assert_eq!(zero_first, half_first);
    assert_eq!(zero_first, one_first);
    assert_eq!(zero_first, two_first);

    // The intrinsic lower bound exceeds 512 KiB: learning that bound must
    // leave the complete source-backed scan intact. Both larger budgets now
    // admit an identical occurrence index and replay the same source ranges.
    assert!(zero_bytes > 0 && zero_reads > 0);
    assert_eq!((half_bytes, half_reads), (zero_bytes, zero_reads));
    assert_eq!((one_bytes, one_reads), (two_bytes, two_reads));
    assert!(one_bytes < zero_bytes / 2);
    assert!(one_reads < zero_reads);
}

#[test]
fn an_out_of_range_shared_string_duplicate_is_typed_then_overwritten_on_hits() {
    let source = Arc::new(CountingSource::new(duplicate_out_of_range_workbook()));
    let owner = SourceBackedWorkbook::from_read_at(source.clone()).unwrap();
    let worksheet = owner.worksheet_by_index(0).unwrap().unwrap();
    let (first, first_bytes, _) = measure_query(&source, &worksheet, 0, 0);
    let (second, second_bytes, _) = measure_query(&source, &worksheet, 0, 0);
    let (third, third_bytes, _) = measure_query(&source, &worksheet, 0, 0);
    let expected = Some(CellValue::String("replaceMe".to_owned()));
    assert_eq!(first, expected);
    assert_eq!(second, expected);
    assert_eq!(third, expected);
    assert!(
        third_bytes < second_bytes / 2,
        "the warm duplicate hit should avoid the complete scan: first={first_bytes}, second={second_bytes}, third={third_bytes}"
    );
}

#[test]
fn a_label_sst_without_a_global_sst_keeps_the_typed_error_on_warm_hits() {
    let (bytes, (row, column)) = missing_sst_workbook();
    let source = Arc::new(CountingSource::new(bytes));
    let owner = SourceBackedWorkbook::from_read_at(source.clone()).unwrap();
    let worksheet = owner.worksheet_by_index(0).unwrap().unwrap();

    let expected = Some(CellValue::Error("SST not available".to_owned()));
    let (first, first_bytes, _) =
        measure_query(&source, &worksheet, u32::from(row), u32::from(column));
    let (second, second_bytes, _) =
        measure_query(&source, &worksheet, u32::from(row), u32::from(column));
    let (third, third_bytes, _) =
        measure_query(&source, &worksheet, u32::from(row), u32::from(column));
    assert_eq!(first, expected);
    assert_eq!(second, expected);
    assert_eq!(third, expected);
    assert!(first_bytes > 0 && second_bytes > 0);
    assert!(
        third_bytes < second_bytes / 2,
        "the indexed LabelSst replay should preserve the missing-SST typed value: first={first_bytes}, second={second_bytes}, third={third_bytes}"
    );
}

#[test]
fn an_earlier_shared_string_read_error_still_refuses_a_warm_duplicate_hit() {
    let source = Arc::new(CountingSource::new(duplicate_shared_string_workbook()));
    let owner = SourceBackedWorkbook::from_read_at(source.clone()).unwrap();
    let worksheet = owner.worksheet_by_index(0).unwrap().unwrap();

    // Warm and publish the index through a different, numeric coordinate.  A
    // latest-only index would retain only the second LabelSst occurrence and
    // would therefore hide an error while replaying this target.
    assert_eq!(
        worksheet.cell_value(50_001, 4).unwrap(),
        Some(CellValue::Float(17.0))
    );
    assert_eq!(
        worksheet.cell_value(50_001, 4).unwrap(),
        Some(CellValue::Float(17.0))
    );

    // The SST bytes contain this marker.  Injecting a one-shot failure only
    // when that source extent is read makes the first duplicate fail during
    // shared-string resolution; the later equal-coordinate occurrence must
    // not overwrite the operation error.
    source.fail_next_read_containing(b"replaceMe");
    let error = worksheet.cell_value(50_000, 4).unwrap_err();
    assert!(
        error
            .to_string()
            .contains("synthetic shared-string read failure"),
        "unexpected warm duplicate error: {error:?}"
    );

    // The one-shot fault is gone.  Replaying both occurrences now succeeds,
    // proving that the cache retained the occurrence chain rather than a
    // single latest locator or a decoded value.
    assert_eq!(
        worksheet.cell_value(50_000, 4).unwrap(),
        Some(CellValue::String("replaceMe".to_owned()))
    );
}

#[test]
fn an_earlier_malformed_formula_string_still_refuses_a_warm_duplicate_hit() {
    let source = Arc::new(CountingSource::new(malformed_formula_duplicate_workbook()));
    let owner = SourceBackedWorkbook::from_read_at(source).unwrap();
    let worksheet = owner.worksheet_by_index(0).unwrap().unwrap();

    // The numeric coordinate warms the index.  The malformed formula is not
    // the requested target during this scan, so its target-specific STRING
    // decode must not be silently replaced by the later valid duplicate.
    assert_eq!(
        worksheet.cell_value(50_001, 4).unwrap(),
        Some(CellValue::Float(19.0))
    );
    assert_eq!(
        worksheet.cell_value(50_001, 4).unwrap(),
        Some(CellValue::Float(19.0))
    );

    let first = worksheet.cell_value(50_000, 4).unwrap_err();
    let second = worksheet.cell_value(50_000, 4).unwrap_err();
    assert_eq!(first.to_string(), second.to_string());
    assert!(
        first.to_string().contains("UTF-16 decoding error"),
        "unexpected malformed formula error: {first:?}"
    );
}

#[test]
fn an_unselected_formula_without_string_keeps_the_missing_result_error_warm() {
    let bytes = formula_without_string_workbook();

    // The second successful query for another coordinate builds the complete
    // occurrence index while the malformed formula is unselected.  Its
    // locator must remain available for a later selected replay.
    let warm_source = Arc::new(CountingSource::new(bytes.clone()));
    let warm_owner = SourceBackedWorkbook::from_read_at(warm_source).unwrap();
    let warm_sheet = warm_owner.worksheet_by_index(0).unwrap().unwrap();
    assert_eq!(
        warm_sheet.cell_value(50_001, 4).unwrap(),
        Some(CellValue::Float(23.0))
    );
    assert_eq!(
        warm_sheet.cell_value(50_001, 4).unwrap(),
        Some(CellValue::Float(23.0))
    );
    let warm = warm_sheet.cell_value(50_000, 4).unwrap_err();

    // Compare the indexed replay with a genuinely cold selected query so the
    // test pins the public refusal identity rather than only the fact of an
    // error.
    let cold_source = Arc::new(CountingSource::new(bytes));
    let cold_owner = SourceBackedWorkbook::from_read_at(cold_source).unwrap();
    let cold_sheet = cold_owner.worksheet_by_index(0).unwrap().unwrap();
    let cold = cold_sheet.cell_value(50_000, 4).unwrap_err();
    assert_eq!(
        warm.to_string(),
        "string-valued FORMULA lacks STRING result"
    );
    assert_eq!(warm.to_string(), cold.to_string());
}

#[test]
fn a_failed_tail_never_publishes_a_partial_index() {
    let source = Arc::new(CountingSource::new(malformed_tail_workbook()));
    let owner = SourceBackedWorkbook::from_read_at(source.clone()).unwrap();
    let worksheet = owner.worksheet_by_index(0).unwrap().unwrap();

    let first = {
        source.clear_measurement();
        worksheet.cell_value(0, 0).unwrap_err()
    };
    let first_bytes = source.bytes_read();
    let first_reads = source.read_count();
    let second = {
        source.clear_measurement();
        worksheet.cell_value(0, 0).unwrap_err()
    };
    assert_eq!(first.to_string(), second.to_string());
    assert_eq!(source.bytes_read(), first_bytes);
    assert_eq!(source.read_count(), first_reads);
}

#[test]
fn formula_string_continuations_have_the_same_value_on_warm_hits() {
    let source = Arc::new(CountingSource::new(formula_string_workbook()));
    let owner = SourceBackedWorkbook::from_read_at(source.clone()).unwrap();
    let worksheet = owner.worksheet_by_index(0).unwrap().unwrap();
    let (first, _, _) = measure_query(&source, &worksheet, 60_005, 6);
    let (second, _, _) = measure_query(&source, &worksheet, 60_005, 6);
    let (third, third_bytes, third_reads) = measure_query(&source, &worksheet, 60_005, 6);
    let expected = Some(CellValue::String("abcdefghi".to_owned()));
    assert_eq!(first, expected);
    assert_eq!(first, second);
    assert_eq!(second, third);
    assert_eq!(third, expected);
    assert!(third_bytes > 0 && third_reads > 0);
}

#[test]
fn cloned_workbooks_share_the_snapshot_cache_but_reopened_ones_do_not() {
    let source = Arc::new(CountingSource::new(large_numeric_workbook(4_000)));
    let owner = SourceBackedWorkbook::from_read_at(source.clone()).unwrap();
    let clone = owner.clone();
    let reopened = SourceBackedWorkbook::from_read_at(source.clone()).unwrap();
    let original_sheet = owner.worksheet_by_index(0).unwrap().unwrap();
    let clone_sheet = clone.worksheet_by_index(0).unwrap().unwrap();
    let reopened_sheet = reopened.worksheet_by_index(0).unwrap().unwrap();

    let _ = measure_query(&source, &original_sheet, 3_999, 0);
    let _ = measure_query(&source, &original_sheet, 3_999, 0);
    let (_, clone_bytes, clone_reads) = measure_query(&source, &clone_sheet, 3_999, 0);
    let (_, reopened_bytes, reopened_reads) = measure_query(&source, &reopened_sheet, 3_999, 0);
    assert!(clone_bytes < reopened_bytes / 2);
    assert!(clone_reads < reopened_reads);
}

#[test]
fn a_warm_index_keeps_alternating_first_last_missing_and_random_values_identical() {
    let bytes = large_numeric_workbook(4_000);
    let coordinates = [
        (1_u32, 0_u32),
        (4_000, 0),
        (50_000, 7),
        (2_048, 0),
        (4_000, 0),
        (1, 0),
        (1_234, 0),
    ];

    let enabled_source = Arc::new(CountingSource::new(bytes.clone()));
    let enabled_owner = SourceBackedWorkbook::from_read_at(enabled_source.clone()).unwrap();
    let enabled_sheet = enabled_owner.worksheet_by_index(0).unwrap().unwrap();
    // The second observation completes the worksheet scan and publishes the
    // immutable index, including its worksheet BOF chain checkpoint.
    assert_eq!(
        enabled_sheet.cell_value(2_000, 0).unwrap(),
        Some(CellValue::Float(2_000.0))
    );
    assert_eq!(
        enabled_sheet.cell_value(2_000, 0).unwrap(),
        Some(CellValue::Float(2_000.0))
    );
    let mut enabled = Vec::new();
    for &(row, column) in &coordinates {
        enabled.push(measure_query(&enabled_source, &enabled_sheet, row, column));
    }

    let disabled_source = Arc::new(CountingSource::new(bytes));
    let disabled_owner = SourceBackedWorkbook::from_read_at_with_limits(
        disabled_source.clone(),
        SourceBackedLimits::default().with_max_query_index_bytes(0),
    )
    .unwrap();
    let disabled_sheet = disabled_owner.worksheet_by_index(0).unwrap().unwrap();
    let mut disabled = Vec::new();
    for &(row, column) in &coordinates {
        disabled.push(measure_query(
            &disabled_source,
            &disabled_sheet,
            row,
            column,
        ));
    }

    let enabled_values = enabled
        .iter()
        .map(|(value, _, _)| value.clone())
        .collect::<Vec<_>>();
    let disabled_values = disabled
        .iter()
        .map(|(value, _, _)| value.clone())
        .collect::<Vec<_>>();
    assert_eq!(enabled_values, disabled_values);
    assert_eq!(enabled_values[0], Some(CellValue::Float(1.0)));
    assert_eq!(enabled_values[1], Some(CellValue::Float(4_000.0)));
    assert_eq!(enabled_values[2], None);
    assert_eq!(enabled_values[3], Some(CellValue::Float(2_048.0)));

    let enabled_bytes: usize = enabled.iter().map(|(_, bytes, _)| *bytes).sum();
    let disabled_bytes: usize = disabled.iter().map(|(_, bytes, _)| *bytes).sum();
    assert!(
        enabled_bytes < disabled_bytes,
        "warm indexed replays should read less than alternating cold scans: enabled={enabled_bytes}, disabled={disabled_bytes}"
    );
    assert_eq!(
        enabled[2].1, 0,
        "indexed missing targets should read no worksheet bytes"
    );
    assert_eq!(
        enabled[2].2, 0,
        "indexed missing targets should issue no reads"
    );
}

#[test]
fn cloned_handles_can_replay_warm_index_concurrently_with_local_hints() {
    let source = Arc::new(CountingSource::new(large_numeric_workbook(4_000)));
    let owner = SourceBackedWorkbook::from_read_at(source).unwrap();
    let warm_sheet = owner.worksheet_by_index(0).unwrap().unwrap();
    assert_eq!(
        warm_sheet.cell_value(2_000, 0).unwrap(),
        Some(CellValue::Float(2_000.0))
    );
    assert_eq!(
        warm_sheet.cell_value(2_000, 0).unwrap(),
        Some(CellValue::Float(2_000.0))
    );

    let first_sheet = owner.clone().worksheet_by_index(0).unwrap().unwrap();
    let last_sheet = owner.clone().worksheet_by_index(0).unwrap().unwrap();
    let random_sheet = owner.clone().worksheet_by_index(0).unwrap().unwrap();
    let missing_sheet = owner.clone().worksheet_by_index(0).unwrap().unwrap();
    let first = std::thread::spawn(move || first_sheet.cell_value(1, 0));
    let last = std::thread::spawn(move || last_sheet.cell_value(4_000, 0));
    let random = std::thread::spawn(move || random_sheet.cell_value(1_234, 0));
    let missing = std::thread::spawn(move || missing_sheet.cell_value(50_000, 7));

    assert_eq!(first.join().unwrap().unwrap(), Some(CellValue::Float(1.0)));
    assert_eq!(
        last.join().unwrap().unwrap(),
        Some(CellValue::Float(4_000.0))
    );
    assert_eq!(
        random.join().unwrap().unwrap(),
        Some(CellValue::Float(1_234.0))
    );
    assert_eq!(missing.join().unwrap().unwrap(), None);
}

#[test]
fn a_managed_memory_refusal_falls_back_without_changing_the_query_result() {
    let source = Arc::new(CountingSource::new(large_numeric_workbook(4_000)));
    let owner = SourceBackedWorkbook::from_read_at(source.clone()).unwrap();
    let worksheet = owner.worksheet_by_index(0).unwrap().unwrap();
    let budget = Budget::root(
        "xls-query-index-test",
        Limits::new(1, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX),
    );
    let (cancellation, token) = CancellationSource::pair();
    let limits = ExecutionLimits::new(
        NonZeroUsize::new(1).unwrap(),
        NonZeroUsize::new(1).unwrap(),
        NonZeroU64::new(8 * 1024 * 1024).unwrap(),
        1,
    )
    .unwrap();
    let execution = ExecutionContext::new(budget.clone(), token, limits);

    let first = worksheet
        .cell_value_with_execution(3_999, 0, &execution)
        .unwrap();
    let second = worksheet
        .cell_value_with_execution(3_999, 0, &execution)
        .unwrap();
    let third = worksheet
        .cell_value_with_execution(3_999, 0, &execution)
        .unwrap();
    assert_eq!(first, second);
    assert_eq!(second, third);
    assert_eq!(budget.used(Resource::Memory), 0);
    drop(cancellation);
}

#[test]
fn cancellation_during_index_build_drops_the_candidate() {
    let source = Arc::new(CountingSource::new(large_numeric_workbook(4_000)));
    let owner = SourceBackedWorkbook::from_read_at(source.clone()).unwrap();
    let worksheet = owner.worksheet_by_index(0).unwrap().unwrap();

    // The first ordinary query arms the two-observation admission trigger;
    // cancellation on the next read therefore interrupts an actual optional
    // candidate build rather than only the mandatory cold scan.
    assert_eq!(
        worksheet.cell_value(3_999, 0).unwrap(),
        Some(CellValue::Float(3_999.0))
    );
    let (cancellation, token) = CancellationSource::pair();
    let budget = Budget::root(
        "xls-query-index-cancellation",
        Limits::new(u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX),
    );
    let limits = ExecutionLimits::new(
        NonZeroUsize::new(1).unwrap(),
        NonZeroUsize::new(1).unwrap(),
        NonZeroU64::new(8 * 1024 * 1024).unwrap(),
        1,
    )
    .unwrap();
    let execution = ExecutionContext::new(budget, token, limits);
    source.cancel_on_next_read(cancellation);
    let error = worksheet
        .cell_value_with_execution(3_999, 0, &execution)
        .unwrap_err();
    assert!(matches!(
        error,
        SourceBackedError::Execution(ExecutionError::Cancelled)
    ));

    // A cancelled candidate is not usable by a later query.  The fresh query
    // must complete and produce the same value as an ordinary source scan.
    let (fresh_cancellation, fresh_token) = CancellationSource::pair();
    let fresh_budget = Budget::root(
        "xls-query-index-fresh",
        Limits::new(u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX),
    );
    let fresh_execution = ExecutionContext::new(fresh_budget, fresh_token, limits);
    drop(fresh_cancellation);
    let value = worksheet
        .cell_value_with_execution(3_999, 0, &fresh_execution)
        .unwrap();
    assert_eq!(value, Some(CellValue::Float(3_999.0)));
}

#[test]
fn a_source_change_during_index_build_is_refused_without_a_warm_hit() {
    let source = Arc::new(CountingSource::new(large_numeric_workbook(4_000)));
    let owner = SourceBackedWorkbook::from_read_at(source.clone()).unwrap();
    let worksheet = owner.worksheet_by_index(0).unwrap().unwrap();

    // Arm the two-observation admission trigger first.  Bump immediately
    // after the next source-version observation, which lands inside the
    // worksheet read/bracket and prevents candidate publication.
    assert_eq!(
        worksheet.cell_value(3_999, 0).unwrap(),
        Some(CellValue::Float(3_999.0))
    );
    source.clear_measurement();
    source.bump_after_observation(2);
    let error = worksheet.cell_value(3_999, 0).unwrap_err();
    assert!(matches!(error, SourceBackedError::SourceChanged { .. }));
}

#[test]
fn visit_cells_keeps_duplicate_order_after_queries_warm_the_owner() {
    let original = workbook_stream(fixture("Simple.xls"));
    let mut extra = Vec::new();
    extra.extend_from_slice(&number_frame(55_000, 3, 7.0));
    extra.extend_from_slice(&number_frame(55_000, 3, 8.0));
    let source = Arc::new(CountingSource::new(cfb_with_workbook(
        &insert_before_worksheet_eof(&original, &extra),
    )));
    let owner = SourceBackedWorkbook::from_read_at(source).unwrap();
    let worksheet = owner.worksheet_by_index(0).unwrap().unwrap();
    let _ = worksheet.cell_value(55_000, 3).unwrap();
    let _ = worksheet.cell_value(55_000, 3).unwrap();
    let mut values = Vec::new();
    worksheet
        .visit_cells(|cell| {
            if cell.row() == 55_000 && cell.column() == 3 {
                values.push(cell.value().clone());
            }
            Ok(())
        })
        .unwrap();
    assert_eq!(
        values,
        vec![CellValue::Float(7.0), CellValue::Float(8.0)],
        "the occurrence index must not replace the independent visitor walk"
    );
}
