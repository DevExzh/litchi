//! Focused differential coverage for the optional target-local worksheet-chain
//! checkpoint retained by the source-backed XLS query index.

#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "fixture construction and test assertions intentionally panic"
)]

use litchi_cfb::{OleWriter, SharedOleFile};
use litchi_core::sheet::CellValue;
use litchi_core::{OwnedSource, ReadAt, SourceVersion};
use litchi_xls::{SourceBackedLimits, SourceBackedWorkbook};
use std::io::{self, Cursor};
use std::sync::atomic::{AtomicU64, Ordering};
use std::sync::{Arc, Mutex};

const CROSS_SECTOR_MARKER: &[u8] = b"target-checkpoint-tail-marker";

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
    let first_sheet_start = first_sheet_offset(stream);
    let eof = worksheet_eof_offset(stream, first_sheet_start);
    let mut output = Vec::with_capacity(stream.len() + extra.len());
    output.extend_from_slice(&stream[..eof]);
    output.extend_from_slice(extra);
    output.extend_from_slice(&stream[eof..]);
    patch_later_sheet_offsets(&mut output, first_sheet_start, eof, extra.len());
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

/// Places the selected frame across a 512-byte CFB sector boundary and puts a
/// fault marker in a later sector. A warm replay that reads only the selected
/// frame must succeed even when the marker's source read is armed to fail.
fn cross_sector_workbook() -> (Vec<u8>, usize) {
    let original = workbook_stream(fixture("Simple.xls"));
    let eof = worksheet_eof_offset(&original, first_sheet_offset(&original));
    let mut extra = Vec::new();
    let mut filler_row = 50_000_u16;
    while (eof + extra.len()) % 512 < 500 {
        extra.extend_from_slice(&number_frame(filler_row, 1, f64::from(filler_row)));
        filler_row = filler_row.saturating_add(1);
    }
    let target_offset = eof + extra.len();
    assert!((500..512).contains(&(target_offset % 512)));
    extra.extend_from_slice(&number_frame(60_000, 0, 42.0));

    while (eof + extra.len()) / 512 == target_offset / 512 {
        extra.extend_from_slice(&number_frame(filler_row, 1, f64::from(filler_row)));
        filler_row = filler_row.saturating_add(1);
    }
    extra.extend_from_slice(&frame_bytes(0x1234, CROSS_SECTOR_MARKER));
    (
        cfb_with_workbook(&insert_before_worksheet_eof(&original, &extra)),
        target_offset,
    )
}

#[derive(Clone)]
struct FaultSource {
    bytes: Arc<Vec<u8>>,
    fail_marker: Arc<Mutex<Option<Vec<u8>>>>,
    reads: Arc<AtomicU64>,
    revision: Arc<AtomicU64>,
}

impl FaultSource {
    fn new(bytes: Vec<u8>) -> Self {
        Self {
            bytes: Arc::new(bytes),
            fail_marker: Arc::new(Mutex::new(None)),
            reads: Arc::new(AtomicU64::new(0)),
            revision: Arc::new(AtomicU64::new(0)),
        }
    }

    fn fail_next_read_containing(&self, marker: &[u8]) {
        *self.fail_marker.lock().unwrap() = Some(marker.to_vec());
    }

    fn reads(&self) -> u64 {
        self.reads.load(Ordering::Relaxed)
    }
}

impl ReadAt for FaultSource {
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
        let fails = self
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
        if fails {
            let _ = self.fail_marker.lock().unwrap().take();
            return Err(io::Error::other("synthetic tail read failure"));
        }
        output[..count].copy_from_slice(&self.bytes[start..start + count]);
        self.reads.fetch_add(1, Ordering::Relaxed);
        Ok(count)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        Ok(SourceVersion::new(
            0x584C_535F_54415247,
            self.revision.load(Ordering::Relaxed),
        ))
    }
}

#[test]
fn target_checkpoint_keeps_forward_and_backward_replays_identical() {
    let bytes = large_numeric_workbook(4_000);
    let enabled =
        SourceBackedWorkbook::from_read_at(Arc::new(OwnedSource::new(bytes.clone()))).unwrap();
    let enabled_sheet = enabled.worksheet_by_index(0).unwrap().unwrap();

    // The second late query builds the immutable index and retains a
    // target-local checkpoint at the first late occurrence.
    assert_eq!(
        enabled_sheet.cell_value(4_000, 0).unwrap(),
        Some(CellValue::Float(4_000.0))
    );
    assert_eq!(
        enabled_sheet.cell_value(4_000, 0).unwrap(),
        Some(CellValue::Float(4_000.0))
    );

    // The earlier target must select the worksheet-start checkpoint, while a
    // later target can use the retained target-local checkpoint. Both paths
    // must continue to decode the exact stored values.
    assert_eq!(
        enabled_sheet.cell_value(1, 0).unwrap(),
        Some(CellValue::Float(1.0))
    );
    assert_eq!(
        enabled_sheet.cell_value(2_048, 0).unwrap(),
        Some(CellValue::Float(2_048.0))
    );

    let disabled = SourceBackedWorkbook::from_read_at_with_limits(
        Arc::new(OwnedSource::new(bytes)),
        SourceBackedLimits::default().with_max_query_index_bytes(0),
    )
    .unwrap();
    let disabled_sheet = disabled.worksheet_by_index(0).unwrap().unwrap();
    assert_eq!(
        disabled_sheet.cell_value(1, 0).unwrap(),
        Some(CellValue::Float(1.0))
    );
    assert_eq!(
        disabled_sheet.cell_value(2_048, 0).unwrap(),
        Some(CellValue::Float(2_048.0))
    );
}

#[test]
fn target_checkpoint_replay_reads_only_a_cross_sector_target_frame() {
    let (bytes, target_offset) = cross_sector_workbook();
    assert!((500..512).contains(&(target_offset % 512)));
    let source = Arc::new(FaultSource::new(bytes));
    let owner = SourceBackedWorkbook::from_read_at(source.clone()).unwrap();
    let sheet = owner.worksheet_by_index(0).unwrap().unwrap();

    assert_eq!(
        sheet.cell_value(60_000, 0).unwrap(),
        Some(CellValue::Float(42.0))
    );
    assert_eq!(
        sheet.cell_value(60_000, 0).unwrap(),
        Some(CellValue::Float(42.0))
    );
    let reads_before = source.reads();

    // The marker starts in a later sector than the selected frame. A replay
    // that accidentally reads ahead into that sector would surface this
    // source error; the exact-frame replay must not.
    source.fail_next_read_containing(CROSS_SECTOR_MARKER);
    assert_eq!(
        sheet.cell_value(60_000, 0).unwrap(),
        Some(CellValue::Float(42.0))
    );
    assert!(source.reads() > reads_before);
}

#[test]
fn a_missing_build_has_no_target_checkpoint_but_publishes_ordinary_slots() {
    let bytes = large_numeric_workbook(2_000);
    let enabled =
        SourceBackedWorkbook::from_read_at(Arc::new(OwnedSource::new(bytes.clone()))).unwrap();
    let enabled_sheet = enabled.worksheet_by_index(0).unwrap().unwrap();
    assert_eq!(enabled_sheet.cell_value(50_000, 7).unwrap(), None);
    assert_eq!(enabled_sheet.cell_value(50_000, 7).unwrap(), None);
    assert_eq!(
        enabled_sheet.cell_value(1_999, 0).unwrap(),
        Some(CellValue::Float(1_999.0))
    );

    let disabled = SourceBackedWorkbook::from_read_at_with_limits(
        Arc::new(OwnedSource::new(bytes)),
        SourceBackedLimits::default().with_max_query_index_bytes(0),
    )
    .unwrap();
    let disabled_sheet = disabled.worksheet_by_index(0).unwrap().unwrap();
    assert_eq!(
        disabled_sheet.cell_value(1_999, 0).unwrap(),
        Some(CellValue::Float(1_999.0))
    );
}
