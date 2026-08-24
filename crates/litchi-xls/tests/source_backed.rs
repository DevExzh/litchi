use litchi_cfb::{OleWriter, SharedOleFile};
use litchi_core::sheet::{Cell as CellTrait, CellValue};
use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionError, ExecutionLimits, Limits,
    OwnedSource, ReadAt, SourceVersion,
};
use litchi_xls::{SourceBackedError, SourceBackedLimits, SourceBackedWorkbook, Workbook};
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

fn ole_fixture(name: &str) -> Vec<u8> {
    std::fs::read(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/ole/xls")
            .join(name),
    )
    .unwrap()
}

fn libreoffice_fixture(name: &str) -> Vec<u8> {
    std::fs::read(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/libreoffice-core/sc/qa/unit/data/xls")
            .join(name),
    )
    .unwrap()
}

fn workbook_stream(bytes: Vec<u8>, name: &str) -> Vec<u8> {
    SharedOleFile::open(Arc::new(OwnedSource::new(bytes)))
        .unwrap()
        .open_stream(&[name])
        .unwrap()
}

fn cfb_with_streams(streams: &[(&str, &[u8])]) -> Vec<u8> {
    let mut writer = OleWriter::new();
    for (name, bytes) in streams {
        writer.create_stream(&[*name], bytes).unwrap();
    }
    let mut output = Vec::new();
    writer.write_to(&mut Cursor::new(&mut output)).unwrap();
    output
}

fn first_sheet_offset(stream: &[u8]) -> usize {
    let mut cursor = 0;
    while cursor + 4 <= stream.len() {
        let kind = u16::from_le_bytes([stream[cursor], stream[cursor + 1]]);
        let length = usize::from(u16::from_le_bytes([stream[cursor + 2], stream[cursor + 3]]));
        if kind == 0x0085 {
            let payload = cursor + 4;
            return usize::try_from(u32::from_le_bytes([
                stream[payload],
                stream[payload + 1],
                stream[payload + 2],
                stream[payload + 3],
            ]))
            .unwrap();
        }
        cursor += 4 + length;
    }
    panic!("fixture has no BoundSheet8");
}

fn worksheet_eof_offset(stream: &[u8], start: usize) -> usize {
    let mut cursor = start;
    loop {
        assert!(cursor + 4 <= stream.len());
        let kind = u16::from_le_bytes([stream[cursor], stream[cursor + 1]]);
        let length = usize::from(u16::from_le_bytes([stream[cursor + 2], stream[cursor + 3]]));
        if kind == 0x000A {
            return cursor;
        }
        cursor += 4 + length;
    }
}

fn insert_before_worksheet_eof(stream: &[u8], extra: &[u8]) -> Vec<u8> {
    let eof = worksheet_eof_offset(stream, first_sheet_offset(stream));
    let mut output = Vec::with_capacity(stream.len() + extra.len());
    output.extend_from_slice(&stream[..eof]);
    output.extend_from_slice(extra);
    output.extend_from_slice(&stream[eof..]);
    let delta = u32::try_from(extra.len()).unwrap();
    let mut cursor = 0;
    while cursor + 4 <= eof {
        let kind = u16::from_le_bytes([output[cursor], output[cursor + 1]]);
        let length = usize::from(u16::from_le_bytes([output[cursor + 2], output[cursor + 3]]));
        if kind == 0x0085 && cursor + 8 <= output.len() {
            let position = u32::from_le_bytes([
                output[cursor + 4],
                output[cursor + 5],
                output[cursor + 6],
                output[cursor + 7],
            ]);
            if usize::try_from(position).unwrap() > eof {
                let shifted = position.checked_add(delta).unwrap();
                output[cursor + 4..cursor + 8].copy_from_slice(&shifted.to_le_bytes());
            }
        }
        cursor += 4 + length;
    }
    output
}

fn global_eof_offset(stream: &[u8]) -> usize {
    let mut cursor = 0;
    loop {
        assert!(cursor + 4 <= stream.len());
        let kind = u16::from_le_bytes([stream[cursor], stream[cursor + 1]]);
        let length = usize::from(u16::from_le_bytes([stream[cursor + 2], stream[cursor + 3]]));
        if kind == 0x000A {
            return cursor;
        }
        cursor += 4 + length;
    }
}

fn insert_before_global_eof(stream: &[u8], extra: &[u8]) -> Vec<u8> {
    let eof = global_eof_offset(stream);
    let mut output = Vec::with_capacity(stream.len() + extra.len());
    output.extend_from_slice(&stream[..eof]);
    output.extend_from_slice(extra);
    output.extend_from_slice(&stream[eof..]);
    let delta = u32::try_from(extra.len()).unwrap();
    let mut cursor = 0;
    while cursor < eof {
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
            if usize::try_from(position).unwrap() >= eof {
                let shifted = position.checked_add(delta).unwrap();
                output[cursor + 4..cursor + 8].copy_from_slice(&shifted.to_le_bytes());
            }
        }
        cursor += 4 + length;
    }
    output
}

fn late_codepage_bound_sheet_stream(stream: &[u8]) -> Vec<u8> {
    let eof = global_eof_offset(stream);
    let mut modified = stream.to_vec();
    let mut cursor = 0;
    let mut changed = false;
    while cursor < eof {
        assert!(cursor + 4 <= stream.len());
        let kind = u16::from_le_bytes([stream[cursor], stream[cursor + 1]]);
        let length = usize::from(u16::from_le_bytes([stream[cursor + 2], stream[cursor + 3]]));
        let end = cursor + 4 + length;
        if kind == 0x0085 {
            let payload = cursor + 4;
            assert!(length >= 9);
            // CP1251 0xC0 is CYRILLIC CAPITAL LETTER A.  Keep the frame
            // width unchanged while changing only its decoded name.
            let name_length = length - 8;
            assert!(name_length <= usize::from(u8::MAX));
            modified[payload + 6] = u8::try_from(name_length).unwrap();
            modified[payload + 7] = 0;
            modified[payload + 8..payload + 8 + name_length].fill(0xC0);
            changed = true;
            break;
        }
        cursor = end;
    }
    assert!(changed, "fixture has no BoundSheet8");
    insert_before_global_eof(&modified, &frame_bytes(0x0042, &1251_u16.to_le_bytes()))
}

fn split_global_sst_stream(stream: &[u8]) -> Vec<u8> {
    let eof = global_eof_offset(stream);
    let mut output = Vec::with_capacity(stream.len() + 4);
    let mut cursor = 0;
    let mut split = false;
    while cursor < eof {
        assert!(cursor + 4 <= stream.len());
        let kind = u16::from_le_bytes([stream[cursor], stream[cursor + 1]]);
        let length = usize::from(u16::from_le_bytes([stream[cursor + 2], stream[cursor + 3]]));
        let end = cursor + 4 + length;
        if kind == 0x00FC && !split {
            let payload = &stream[cursor + 4..end];
            assert!(payload.len() > 8, "fixture SST is too short to split");
            // Ending the SST segment after its two counts is a valid
            // continuation boundary: the next segment starts at the first
            // shared-string header, while any existing SST CONTINUE records
            // remain in their original order and retain their flags.
            output.extend_from_slice(&frame_bytes(0x00FC, &payload[..8]));
            output.extend_from_slice(&frame_bytes(0x003C, &payload[8..]));
            cursor = end;
            split = true;
        } else {
            output.extend_from_slice(&stream[cursor..end]);
            cursor = end;
        }
    }
    assert!(split, "fixture has no SST");
    output.extend_from_slice(&stream[eof..]);

    let new_eof = global_eof_offset(&output);
    let delta = i64::try_from(output.len()).unwrap() - i64::try_from(stream.len()).unwrap();
    let mut cursor = 0;
    while cursor < new_eof {
        let kind = u16::from_le_bytes([output[cursor], output[cursor + 1]]);
        let length = usize::from(u16::from_le_bytes([output[cursor + 2], output[cursor + 3]]));
        if kind == 0x0085 {
            let position = u32::from_le_bytes([
                output[cursor + 4],
                output[cursor + 5],
                output[cursor + 6],
                output[cursor + 7],
            ]);
            let shifted = if delta >= 0 {
                position.checked_add(u32::try_from(delta).unwrap()).unwrap()
            } else {
                position
                    .checked_sub(u32::try_from(-delta).unwrap())
                    .unwrap()
            };
            output[cursor + 4..cursor + 8].copy_from_slice(&shifted.to_le_bytes());
        }
        cursor += 4 + length;
    }
    output
}

fn number_frame(row: u16, column: u16, xf: u16, value: f64) -> Vec<u8> {
    let mut payload = Vec::with_capacity(14);
    payload.extend_from_slice(&row.to_le_bytes());
    payload.extend_from_slice(&column.to_le_bytes());
    payload.extend_from_slice(&xf.to_le_bytes());
    payload.extend_from_slice(&value.to_le_bytes());
    frame_bytes(0x0203, &payload)
}

fn rk_frame(row: u16, column: u16, xf: u16, value: u32) -> Vec<u8> {
    let mut payload = Vec::with_capacity(10);
    payload.extend_from_slice(&row.to_le_bytes());
    payload.extend_from_slice(&column.to_le_bytes());
    payload.extend_from_slice(&xf.to_le_bytes());
    payload.extend_from_slice(&((value << 2) | 0x02).to_le_bytes());
    frame_bytes(0x027E, &payload)
}

fn mul_rk_frame(row: u16, first_column: u16, values: &[u32]) -> Vec<u8> {
    let mut payload = Vec::with_capacity(6 + values.len() * 6);
    payload.extend_from_slice(&row.to_le_bytes());
    payload.extend_from_slice(&first_column.to_le_bytes());
    for value in values {
        payload.extend_from_slice(&0_u16.to_le_bytes());
        payload.extend_from_slice(&((*value << 2) | 0x02).to_le_bytes());
    }
    payload.extend_from_slice(
        &(first_column + u16::try_from(values.len()).unwrap() - 1).to_le_bytes(),
    );
    frame_bytes(0x00BD, &payload)
}

fn mul_blank_frame(row: u16, first_column: u16, count: usize) -> Vec<u8> {
    let mut payload = Vec::with_capacity(6 + count * 2);
    payload.extend_from_slice(&row.to_le_bytes());
    payload.extend_from_slice(&first_column.to_le_bytes());
    for _ in 0..count {
        payload.extend_from_slice(&0_u16.to_le_bytes());
    }
    payload.extend_from_slice(&(first_column + u16::try_from(count).unwrap() - 1).to_le_bytes());
    frame_bytes(0x00BE, &payload)
}

fn bool_err_frame(row: u16, column: u16, value: u8, is_error: bool) -> Vec<u8> {
    let mut payload = Vec::with_capacity(8);
    payload.extend_from_slice(&row.to_le_bytes());
    payload.extend_from_slice(&column.to_le_bytes());
    payload.extend_from_slice(&0_u16.to_le_bytes());
    payload.push(value);
    payload.push(u8::from(is_error));
    frame_bytes(0x0205, &payload)
}

fn blank_frame(row: u16, column: u16, xf: u16) -> Vec<u8> {
    let mut payload = Vec::with_capacity(6);
    payload.extend_from_slice(&row.to_le_bytes());
    payload.extend_from_slice(&column.to_le_bytes());
    payload.extend_from_slice(&xf.to_le_bytes());
    frame_bytes(0x0201, &payload)
}

fn label_frame(row: u16, column: u16, xf: u16, value: &[u8]) -> Vec<u8> {
    let mut payload = Vec::with_capacity(9 + value.len());
    payload.extend_from_slice(&row.to_le_bytes());
    payload.extend_from_slice(&column.to_le_bytes());
    payload.extend_from_slice(&xf.to_le_bytes());
    payload.extend_from_slice(&u16::try_from(value.len()).unwrap().to_le_bytes());
    payload.push(0);
    payload.extend_from_slice(value);
    frame_bytes(0x0204, &payload)
}

fn first_frame_of_kind(stream: &[u8], wanted: u16) -> Vec<u8> {
    let mut cursor = first_sheet_offset(stream);
    loop {
        assert!(cursor + 4 <= stream.len());
        let kind = u16::from_le_bytes([stream[cursor], stream[cursor + 1]]);
        let length = usize::from(u16::from_le_bytes([stream[cursor + 2], stream[cursor + 3]]));
        let end = cursor + 4 + length;
        if kind == wanted {
            return stream[cursor..end].to_vec();
        }
        assert_ne!(kind, 0x000A, "fixture has no requested BIFF frame");
        cursor = end;
    }
}

fn shifted_label_sst_frame(stream: &[u8], row: u16, column: u16) -> Vec<u8> {
    let mut frame = first_frame_of_kind(stream, 0x00FD);
    frame[4..6].copy_from_slice(&row.to_le_bytes());
    frame[6..8].copy_from_slice(&column.to_le_bytes());
    frame
}

fn first_numeric_formula(stream: &[u8]) -> (u16, u16, Vec<u8>) {
    let mut cursor = first_sheet_offset(stream);
    loop {
        assert!(cursor + 4 <= stream.len());
        let kind = u16::from_le_bytes([stream[cursor], stream[cursor + 1]]);
        let length = usize::from(u16::from_le_bytes([stream[cursor + 2], stream[cursor + 3]]));
        let end = cursor + 4 + length;
        if kind == 0x0006 {
            let payload = &stream[cursor + 4..end];
            if payload.len() >= 14 && payload[12..14] != [0xFF, 0xFF] {
                return (
                    u16::from_le_bytes([payload[0], payload[1]]),
                    u16::from_le_bytes([payload[2], payload[3]]),
                    stream[cursor..end].to_vec(),
                );
            }
        }
        if kind == 0x000A {
            break;
        }
        cursor = end;
    }
    panic!("formula fixture has no numeric cached formula");
}

fn first_label_sst_cell(stream: &[u8]) -> (u16, u16) {
    let frame = first_frame_of_kind(stream, 0x00FD);
    (
        u16::from_le_bytes([frame[4], frame[5]]),
        u16::from_le_bytes([frame[6], frame[7]]),
    )
}

#[derive(Clone)]
struct CountingSource {
    bytes: Arc<Vec<u8>>,
    ranges: Arc<Mutex<Vec<(u64, usize)>>>,
    cancel_on_read: Arc<Mutex<Option<CancellationSource>>>,
    revision: Arc<AtomicU64>,
}

impl CountingSource {
    fn new(bytes: Vec<u8>) -> Self {
        Self {
            bytes: Arc::new(bytes),
            ranges: Arc::new(Mutex::new(Vec::new())),
            cancel_on_read: Arc::new(Mutex::new(None)),
            revision: Arc::new(AtomicU64::new(0)),
        }
    }

    fn bytes_read(&self) -> usize {
        self.ranges
            .lock()
            .unwrap()
            .iter()
            .map(|(_, length)| *length)
            .sum()
    }

    fn ranges(&self) -> Vec<(u64, usize)> {
        self.ranges.lock().unwrap().clone()
    }

    fn clear_ranges(&self) {
        self.ranges.lock().unwrap().clear();
    }

    fn bump(&self) {
        self.revision.fetch_add(1, Ordering::Relaxed);
    }

    fn cancel_on_next_read(&self, source: CancellationSource) {
        *self.cancel_on_read.lock().unwrap() = Some(source);
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
        output[..count].copy_from_slice(&self.bytes[start..start + count]);
        self.ranges.lock().unwrap().push((offset, count));
        if let Some(source) = self.cancel_on_read.lock().unwrap().take() {
            source.cancel();
        }
        Ok(count)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        Ok(SourceVersion::new(
            0x584c_535f_5445_5354,
            self.revision.load(Ordering::Relaxed),
        ))
    }
}

fn ranges_overlap(left: (u64, usize), right: (u64, usize)) -> bool {
    let left_end = left.0.checked_add(left.1 as u64).unwrap();
    let right_end = right.0.checked_add(right.1 as u64).unwrap();
    left.0 < right_end && right.0 < left_end
}

fn overlaps_any(ranges: &[(u64, usize)], probes: &[(u64, usize)]) -> bool {
    ranges.iter().copied().any(|range| {
        probes
            .iter()
            .copied()
            .any(|probe| ranges_overlap(range, probe))
    })
}

fn physical_ranges_for_stream_range(
    source: &CountingSource,
    offset: u64,
    length: usize,
) -> Vec<(u64, usize)> {
    source.clear_ranges();
    let cfb = SharedOleFile::open(Arc::new(source.clone())).unwrap();
    source.clear_ranges();
    let mut output = vec![0; length];
    cfb.read_stream_range(&["Workbook"], offset, &mut output)
        .unwrap();
    let ranges = source.ranges();
    source.clear_ranges();
    ranges
}

fn workbook_bound_sheet_positions(stream: &[u8]) -> Vec<usize> {
    let eof = global_eof_offset(stream);
    let mut positions = Vec::new();
    let mut cursor = 0;
    while cursor < eof {
        assert!(cursor + 4 <= stream.len());
        let kind = u16::from_le_bytes([stream[cursor], stream[cursor + 1]]);
        let length = usize::from(u16::from_le_bytes([stream[cursor + 2], stream[cursor + 3]]));
        if kind == 0x0085 {
            positions.push(
                usize::try_from(u32::from_le_bytes([
                    stream[cursor + 4],
                    stream[cursor + 5],
                    stream[cursor + 6],
                    stream[cursor + 7],
                ]))
                .unwrap(),
            );
        }
        cursor += 4 + length;
    }
    positions
}

fn physical_ranges_for_workbook_sheets(
    source: &CountingSource,
    stream: &[u8],
) -> Vec<Vec<(u64, usize)>> {
    let positions = workbook_bound_sheet_positions(stream);
    let mut boundaries = positions.clone();
    boundaries.sort_unstable();
    let cfb = SharedOleFile::open(Arc::new(source.clone())).unwrap();
    source.clear_ranges();
    let mut all_ranges = Vec::with_capacity(positions.len());
    for start in positions {
        let end = boundaries
            .iter()
            .copied()
            .find(|boundary| *boundary > start)
            .unwrap_or(stream.len());
        let mut ranges = Vec::new();
        let mut offset = start as u64;
        while offset < end as u64 {
            let length = usize::try_from((end as u64 - offset).min(8 * 1024)).unwrap();
            let mut output = vec![0; length];
            source.clear_ranges();
            cfb.read_stream_range(&["Workbook"], offset, &mut output)
                .unwrap();
            ranges.extend(source.ranges());
            offset += length as u64;
        }
        all_ranges.push(ranges);
    }
    source.clear_ranges();
    all_ranges
}

fn execution_pair() -> (CancellationSource, ExecutionContext) {
    let budget = Budget::root(
        "xls-source-backed-test",
        Limits::new(
            8 * 1024 * 1024,
            u64::MAX,
            u64::MAX,
            1_000_000,
            u64::MAX,
            u64::MAX,
        ),
    );
    let (source, token) = CancellationSource::pair();
    let limits = ExecutionLimits::new(
        NonZeroUsize::new(1).unwrap(),
        NonZeroUsize::new(1).unwrap(),
        NonZeroU64::new(8 * 1024 * 1024).unwrap(),
        1,
    )
    .unwrap();
    (source, ExecutionContext::new(budget, token, limits))
}

#[test]
fn opens_without_worksheet_payload_and_matches_eager_selected_cells() {
    let bytes = fixture("Simple.xls");
    let source = Arc::new(CountingSource::new(bytes.clone()));
    let owner = SourceBackedWorkbook::from_read_at(source.clone()).unwrap();
    let worksheet_count = owner.worksheet_count().unwrap();
    assert!(worksheet_count > 0);
    assert_eq!(owner.worksheet_names().unwrap().len(), worksheet_count);
    assert!(source.bytes_read() < bytes.len());

    let eager = Workbook::new(Cursor::new(bytes)).unwrap();
    let mut saw_string = false;
    for worksheet_index in 0..worksheet_count {
        let eager_sheet = eager.xls_worksheet(worksheet_index).unwrap();
        for row in 0..4 {
            for column in 0..4 {
                let expected = eager_sheet.get_cell(row, column).map(CellTrait::value);
                saw_string |= matches!(expected, Some(CellValue::String(_)));
                let actual = owner
                    .cell_value_by_index(worksheet_index, row, column)
                    .unwrap();
                assert_eq!(actual.as_ref(), expected);
            }
        }
    }
    assert!(saw_string);
    assert_eq!(owner.cell_value_by_index(0, 65_535, 65_535).unwrap(), None);
}

#[test]
fn column_beyond_biff8_visible_range_returns_none_without_reading_worksheet() {
    let source = Arc::new(CountingSource::new(fixture("Simple.xls")));
    let owner = SourceBackedWorkbook::from_read_at(source.clone()).unwrap();
    source.clear_ranges();
    assert_eq!(owner.cell_value_by_index(0, 0, 256).unwrap(), None);
    assert_eq!(source.bytes_read(), 0);
}

#[test]
fn materialize_eager_returns_typed_semantic_workbook_with_matching_values() {
    let bytes = fixture("Simple.xls");
    let owner = SourceBackedWorkbook::from_read_at(Arc::new(CountingSource::new(bytes))).unwrap();
    let materialized: Workbook<Cursor<Vec<u8>>> = owner.materialize_eager().unwrap();
    let sheet = materialized.xls_worksheet(0).unwrap();
    for row in 0..4 {
        for column in 0..4 {
            let expected = owner.cell_value_by_index(0, row, column).unwrap();
            let actual = sheet.get_cell(row, column).map(CellTrait::value);
            assert_eq!(actual, expected.as_ref(), "cell ({row}, {column})");
        }
    }
}

#[test]
fn materialize_eager_limit_is_typed_and_independent() {
    let source = Arc::new(CountingSource::new(fixture("Simple.xls")));
    let limits = SourceBackedLimits::default().with_max_materialize_bytes(1);
    let owner = SourceBackedWorkbook::from_read_at_with_limits(source, limits).unwrap();
    assert!(matches!(
        owner.materialize_eager(),
        Err(SourceBackedError::ResourceLimit {
            resource: "materialization bytes",
            ..
        })
    ));
}

#[test]
fn late_codepage_bound_sheet_names_match_eager_workbook() {
    let original = workbook_stream(fixture("Simple.xls"), "Workbook");
    let modified = late_codepage_bound_sheet_stream(&original);
    let bytes = cfb_with_streams(&[("Workbook", &modified)]);
    let owner =
        SourceBackedWorkbook::from_read_at(Arc::new(CountingSource::new(bytes.clone()))).unwrap();
    let eager = Workbook::new(Cursor::new(bytes)).unwrap();
    let expected = eager.sheet(0).unwrap().name().to_owned();
    assert_eq!(owner.worksheet_names().unwrap()[0], expected);
}

#[test]
fn split_global_sst_continue_matches_eager_workbook() {
    let original = workbook_stream(fixture("Simple.xls"), "Workbook");
    let (row, column) = first_label_sst_cell(&original);
    let modified = split_global_sst_stream(&original);
    let bytes = cfb_with_streams(&[("Workbook", &modified)]);
    let owner =
        SourceBackedWorkbook::from_read_at(Arc::new(CountingSource::new(bytes.clone()))).unwrap();
    let eager = Workbook::new(Cursor::new(bytes)).unwrap();
    let expected = eager
        .xls_worksheet(0)
        .unwrap()
        .get_cell(u32::from(row), u32::from(column))
        .map(CellTrait::value);
    assert!(matches!(expected, Some(CellValue::String(_))));
    assert_eq!(
        owner
            .cell_value_by_index(0, u32::from(row), u32::from(column))
            .unwrap()
            .as_ref(),
        expected
    );
}

#[test]
fn formula_fixture_matches_eager_cached_values() {
    let bytes = ole_fixture("FormulaSheetRange.xls");
    let owner =
        SourceBackedWorkbook::from_read_at(Arc::new(CountingSource::new(bytes.clone()))).unwrap();
    let eager = Workbook::new(Cursor::new(bytes)).unwrap();
    for worksheet_index in 0..owner.worksheet_count().unwrap() {
        let eager_sheet = eager.xls_worksheet(worksheet_index).unwrap();
        for row in 0..12 {
            for column in 0..8 {
                let expected = eager_sheet.get_cell(row, column).map(CellTrait::value);
                let actual = owner
                    .cell_value_by_index(worksheet_index, row, column)
                    .unwrap();
                assert_eq!(actual.as_ref(), expected);
            }
        }
    }
}

#[test]
fn date_fixture_matches_eager_date_values() {
    let bytes = libreoffice_fixture("pivottable_dates_grouping.xls");
    let owner =
        SourceBackedWorkbook::from_read_at(Arc::new(CountingSource::new(bytes.clone()))).unwrap();
    let eager = Workbook::new(Cursor::new(bytes)).unwrap();
    let mut dates = Vec::new();
    for worksheet_index in 0..owner.worksheet_count().unwrap() {
        let eager_sheet = eager.xls_worksheet(worksheet_index).unwrap();
        for row in 0..64 {
            for column in 0..16 {
                let expected = eager_sheet.get_cell(row, column).map(CellTrait::value);
                if matches!(expected, Some(CellValue::DateTime(_))) {
                    dates.push((worksheet_index, row, column, expected.cloned()));
                }
            }
        }
    }
    assert!(
        !dates.is_empty(),
        "date fixture did not expose a DateTime cell"
    );
    for (worksheet_index, row, column, expected) in dates {
        let actual = owner
            .cell_value_by_index(worksheet_index, row, column)
            .unwrap();
        assert_eq!(actual, expected);
    }
}

#[test]
fn selected_queries_are_bounded_to_the_selected_owner() {
    let bytes = fixture("TwoSheetsOneHidden.xls");
    let source = Arc::new(CountingSource::new(bytes.clone()));
    let owner = SourceBackedWorkbook::from_read_at(source.clone()).unwrap();
    let open_ranges = source.ranges();
    let workbook = workbook_stream(bytes.clone(), "Workbook");
    let sheet_ranges = physical_ranges_for_workbook_sheets(&source, &workbook);
    let descriptors = owner.worksheet_descriptors().unwrap();
    assert!(descriptors.len() >= 2);
    let first_workbook_index = descriptors[0].workbook_index();
    let second_workbook_index = descriptors[1].workbook_index();
    source.clear_ranges();
    let _ = owner.cell_value_by_index(0, 0, 0).unwrap();
    let first_ranges = source.ranges();
    source.clear_ranges();
    let _ = owner.cell_value_by_index(1, 0, 0).unwrap();
    let second_ranges = source.ranges();
    let first_query: usize = first_ranges.iter().map(|(_, length)| *length).sum();
    let second_query: usize = second_ranges.iter().map(|(_, length)| *length).sum();
    assert!(first_query > 0);
    assert!(second_query > 0);
    assert!(first_query < bytes.len());
    assert!(second_query < bytes.len());
    let all_sheet_ranges = sheet_ranges.iter().flatten().copied().collect::<Vec<_>>();
    assert!(
        open_ranges
            .iter()
            .copied()
            .all(|range| !overlaps_any(&[range], &all_sheet_ranges))
    );
    for (query_ranges, selected_workbook_index) in [
        (&first_ranges, first_workbook_index),
        (&second_ranges, second_workbook_index),
    ] {
        assert!(overlaps_any(
            query_ranges,
            &sheet_ranges[selected_workbook_index]
        ));
        assert!(query_ranges.iter().copied().all(|range| {
            sheet_ranges[selected_workbook_index]
                .iter()
                .copied()
                .any(|selected| ranges_overlap(range, selected))
        }));
        assert!(
            sheet_ranges
                .iter()
                .enumerate()
                .filter(|(index, _)| *index != selected_workbook_index)
                .all(|(_, unselected)| !overlaps_any(query_ranges, unselected))
        );
    }
}

#[test]
fn selected_scan_limits_are_enforced_incrementally() {
    let bytes = fixture("Simple.xls");
    let early_limit = SourceBackedLimits::default().with_max_worksheet_scan_records(32);
    let source = Arc::new(CountingSource::new(bytes.clone()));
    let owner = SourceBackedWorkbook::from_read_at_with_limits(source, early_limit).unwrap();
    assert!(owner.cell_value_by_index(0, 0, 0).unwrap().is_some());

    let bytes_limit = SourceBackedLimits::default().with_max_worksheet_scan_bytes(8);
    let source = Arc::new(CountingSource::new(bytes.clone()));
    let owner = SourceBackedWorkbook::from_read_at_with_limits(source, bytes_limit).unwrap();
    assert!(matches!(
        owner.cell_value_by_index(0, 0, 0),
        Err(SourceBackedError::ResourceLimit {
            resource: "worksheet scan bytes",
            ..
        })
    ));

    let records_limit = SourceBackedLimits::default().with_max_worksheet_scan_records(1);
    let source = Arc::new(CountingSource::new(bytes));
    let owner = SourceBackedWorkbook::from_read_at_with_limits(source, records_limit).unwrap();
    assert!(matches!(
        owner.cell_value_by_index(0, 0, 0),
        Err(SourceBackedError::ResourceLimit {
            resource: "worksheet scan records",
            ..
        })
    ));
}

#[test]
fn duplicate_scalar_keeps_the_latest_value_like_eager_workbook() {
    let original = workbook_stream(fixture("Simple.xls"), "Workbook");
    let modified = insert_before_worksheet_eof(&original, &number_frame(0, 0, 0, 99.0));
    let bytes = cfb_with_streams(&[("Workbook", &modified)]);
    let owner =
        SourceBackedWorkbook::from_read_at(Arc::new(CountingSource::new(bytes.clone()))).unwrap();
    let eager = Workbook::new(Cursor::new(bytes)).unwrap();
    let expected = eager
        .xls_worksheet(0)
        .unwrap()
        .get_cell(0, 0)
        .map(CellTrait::value);
    assert_eq!(
        owner.cell_value_by_index(0, 0, 0).unwrap().as_ref(),
        expected
    );
}

#[test]
fn duplicate_cached_formula_keeps_the_latest_value_like_eager_workbook() {
    let original = workbook_stream(ole_fixture("FormulaSheetRange.xls"), "Workbook");
    let (row, column, mut formula) = first_numeric_formula(&original);
    formula[10..18].copy_from_slice(&987.25_f64.to_le_bytes());
    let modified = insert_before_worksheet_eof(&original, &formula);
    let bytes = cfb_with_streams(&[("Workbook", &modified)]);
    let owner =
        SourceBackedWorkbook::from_read_at(Arc::new(CountingSource::new(bytes.clone()))).unwrap();
    let eager = Workbook::new(Cursor::new(bytes)).unwrap();
    let expected = eager
        .xls_worksheet(0)
        .unwrap()
        .get_cell(u32::from(row), u32::from(column))
        .map(CellTrait::value);
    assert_eq!(
        owner
            .cell_value_by_index(0, u32::from(row), u32::from(column))
            .unwrap()
            .as_ref(),
        expected
    );
}

#[test]
fn synthetic_supported_cell_families_match_eager_values() {
    let original = workbook_stream(fixture("Simple.xls"), "Workbook");
    let mut extra = Vec::new();
    extra.extend_from_slice(&rk_frame(60_000, 1, 0, 42));
    extra.extend_from_slice(&mul_rk_frame(60_001, 2, &[7, 11]));
    extra.extend_from_slice(&mul_blank_frame(60_002, 3, 2));
    extra.extend_from_slice(&bool_err_frame(60_003, 4, 1, false));
    extra.extend_from_slice(&shifted_label_sst_frame(&original, 60_004, 5));
    let modified = insert_before_worksheet_eof(&original, &extra);
    let bytes = cfb_with_streams(&[("Workbook", &modified)]);
    let owner =
        SourceBackedWorkbook::from_read_at(Arc::new(CountingSource::new(bytes.clone()))).unwrap();
    let eager = Workbook::new(Cursor::new(bytes)).unwrap();
    let eager_sheet = eager.xls_worksheet(0).unwrap();
    for (row, column) in [
        (60_000, 1),
        (60_001, 2),
        (60_001, 3),
        (60_002, 3),
        (60_002, 4),
        (60_003, 4),
        (60_004, 5),
    ] {
        let expected = eager_sheet.get_cell(row, column).map(CellTrait::value);
        let actual = owner.cell_value_by_index(0, row, column).unwrap();
        assert_eq!(actual.as_ref(), expected, "cell ({row}, {column})");
    }
}

#[test]
fn scalar_label_blank_and_error_boolerr_match_eager_values() {
    let original = workbook_stream(fixture("Simple.xls"), "Workbook");
    let mut extra = Vec::new();
    extra.extend_from_slice(&label_frame(60_006, 7, 0, b"scalar label"));
    extra.extend_from_slice(&blank_frame(60_007, 8, 0));
    extra.extend_from_slice(&bool_err_frame(60_008, 9, 7, true));
    let modified = insert_before_worksheet_eof(&original, &extra);
    let bytes = cfb_with_streams(&[("Workbook", &modified)]);
    let owner =
        SourceBackedWorkbook::from_read_at(Arc::new(CountingSource::new(bytes.clone()))).unwrap();
    let eager = Workbook::new(Cursor::new(bytes)).unwrap();
    let eager_sheet = eager.xls_worksheet(0).unwrap();
    for (row, column) in [(60_006, 7), (60_007, 8), (60_008, 9)] {
        let expected = eager_sheet.get_cell(row, column).map(CellTrait::value);
        let actual = owner.cell_value_by_index(0, row, column).unwrap();
        assert_eq!(actual.as_ref(), expected, "cell ({row}, {column})");
    }
}

#[test]
fn cached_formula_string_with_multiple_continues_matches_eager_value() {
    let original = workbook_stream(fixture("Simple.xls"), "Workbook");
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
    chain.extend_from_slice(&frame_bytes(0x0207, &string));
    chain.extend_from_slice(&continuation(b"def"));
    chain.extend_from_slice(&continuation(b"ghi"));
    let modified = insert_before_worksheet_eof(&original, &chain);
    let bytes = cfb_with_streams(&[("Workbook", &modified)]);
    let owner =
        SourceBackedWorkbook::from_read_at(Arc::new(CountingSource::new(bytes.clone()))).unwrap();
    let eager = Workbook::new(Cursor::new(bytes)).unwrap();
    let expected = eager
        .xls_worksheet(0)
        .unwrap()
        .get_cell(60_005, 6)
        .map(CellTrait::value);
    assert_eq!(
        owner.cell_value_by_index(0, 60_005, 6).unwrap().as_ref(),
        expected
    );
}

#[test]
fn truncated_tail_is_refused_after_a_matching_cell() {
    let original = workbook_stream(fixture("Simple.xls"), "Workbook");
    let truncated = [0x03_u8, 0x02, 14, 0, 0];
    let modified = insert_before_worksheet_eof(&original, &truncated);
    let bytes = cfb_with_streams(&[("Workbook", &modified)]);
    let owner = SourceBackedWorkbook::from_read_at(Arc::new(CountingSource::new(bytes))).unwrap();
    assert!(matches!(
        owner.cell_by_index(0, 0, 0),
        Err(SourceBackedError::InvalidData(_))
    ));
}

#[test]
fn oversized_unknown_record_is_refused_before_payload_read() {
    let original = workbook_stream(fixture("Simple.xls"), "Workbook");
    let payload = vec![0xA5; litchi_biff::MAX_RECORD_BYTES + 1];
    let oversized = frame_bytes(0x1234, &payload);
    let oversized_offset = worksheet_eof_offset(&original, first_sheet_offset(&original));
    let modified = insert_before_worksheet_eof(&original, &oversized);
    let bytes = cfb_with_streams(&[("Workbook", &modified)]);
    let source = Arc::new(CountingSource::new(bytes));
    let owner = SourceBackedWorkbook::from_read_at(source.clone()).unwrap();
    let header_ranges = physical_ranges_for_stream_range(&source, oversized_offset as u64, 4);
    let payload_ranges =
        physical_ranges_for_stream_range(&source, oversized_offset as u64 + 4, payload.len());
    source.clear_ranges();
    assert!(matches!(
        owner.cell_by_index(0, 0, 0),
        Err(SourceBackedError::ResourceLimit {
            resource: "BIFF record bytes",
            ..
        })
    ));
    let query_ranges = source.ranges();
    assert!(overlaps_any(&query_ranges, &header_ranges));
    assert!(!overlaps_any(&query_ranges, &payload_ranges));
}

#[test]
fn stale_sources_are_rejected_before_selected_reads() {
    let source = Arc::new(CountingSource::new(fixture("Simple.xls")));
    let owner = SourceBackedWorkbook::from_read_at(source.clone()).unwrap();
    source.bump();
    assert!(matches!(
        owner.worksheet_count(),
        Err(SourceBackedError::SourceChanged { .. })
    ));
    assert!(matches!(
        owner.worksheet_names(),
        Err(SourceBackedError::SourceChanged { .. })
    ));
    assert!(matches!(
        owner.date_system(),
        Err(SourceBackedError::SourceChanged { .. })
    ));
    assert!(matches!(
        owner.worksheet_by_index(0),
        Err(SourceBackedError::SourceChanged { .. })
    ));
    assert!(matches!(
        owner.worksheet_by_name("Sheet1"),
        Err(SourceBackedError::SourceChanged { .. })
    ));
    assert!(matches!(
        owner.worksheets(),
        Err(SourceBackedError::SourceChanged { .. })
    ));
    assert!(matches!(
        owner.cell_by_index(0, u32::MAX, 0),
        Err(SourceBackedError::SourceChanged { .. })
    ));
    assert!(matches!(
        owner.cell_value_by_index(0, 0, 0),
        Err(SourceBackedError::SourceChanged { .. })
    ));
    assert!(matches!(
        owner.materialize_eager(),
        Err(SourceBackedError::SourceChanged { .. })
    ));
}

#[test]
fn execution_variants_honor_pre_and_mid_scan_cancellation() {
    let source = Arc::new(CountingSource::new(fixture("Simple.xls")));
    let owner = SourceBackedWorkbook::from_read_at(source.clone()).unwrap();
    let (cancellation, context) = execution_pair();
    cancellation.cancel();
    assert!(matches!(
        owner.cell_value_with_execution(0, 0, 0, &context),
        Err(SourceBackedError::Execution(ExecutionError::Cancelled))
    ));

    let (cancellation, context) = execution_pair();
    cancellation.cancel();
    assert!(matches!(
        owner.materialize_eager_with_execution(&context),
        Err(SourceBackedError::Execution(ExecutionError::Cancelled))
    ));

    let (cancellation, context) = execution_pair();
    source.cancel_on_next_read(cancellation);
    assert!(matches!(
        owner.cell_value_with_execution(0, 0, 0, &context),
        Err(SourceBackedError::Execution(ExecutionError::Cancelled))
    ));

    let (cancellation, context) = execution_pair();
    source.cancel_on_next_read(cancellation);
    assert!(matches!(
        owner.materialize_eager_with_execution(&context),
        Err(SourceBackedError::Execution(ExecutionError::Cancelled))
    ));
}

#[test]
fn explicit_limits_and_filepass_are_typed() {
    let source = Arc::new(CountingSource::new(fixture("Simple.xls")));
    let limits = SourceBackedLimits::default().with_max_global_bytes(1);
    assert!(matches!(
        SourceBackedWorkbook::from_read_at_with_limits(source, limits),
        Err(SourceBackedError::ResourceLimit {
            resource: "global bytes",
            ..
        })
    ));

    let encrypted = Arc::new(CountingSource::new(fixture("xor-encryption-abc.xls")));
    assert!(matches!(
        SourceBackedWorkbook::from_read_at(encrypted),
        Err(SourceBackedError::EncryptedUnsupported)
    ));
}

#[test]
fn filepass_header_scan_never_reads_its_payload_or_worksheets() {
    let original = workbook_stream(fixture("Simple.xls"), "Workbook");
    let filepass_offset = global_eof_offset(&original) as u64;
    let worksheet_positions = workbook_bound_sheet_positions(&original);
    let mut worksheet_boundaries = worksheet_positions.clone();
    worksheet_boundaries.sort_unstable();
    for payload in [vec![0xA5; 8_192], vec![0xA5; 2]] {
        let mut filepass = Vec::with_capacity(6 + payload.len());
        filepass.extend_from_slice(&0x002F_u16.to_le_bytes());
        filepass.extend_from_slice(&8_192_u16.to_le_bytes());
        filepass.extend_from_slice(&payload);
        let modified = insert_before_global_eof(&original, &filepass);
        let bytes = cfb_with_streams(&[("Workbook", &modified)]);
        let source = Arc::new(CountingSource::new(bytes));

        let header_ranges = physical_ranges_for_stream_range(&source, filepass_offset, 4);
        let payload_ranges =
            physical_ranges_for_stream_range(&source, filepass_offset + 4, payload.len());
        let worksheet_ranges = worksheet_positions
            .iter()
            .map(|start| {
                let end = worksheet_boundaries
                    .iter()
                    .copied()
                    .find(|boundary| *boundary > *start)
                    .unwrap_or(original.len());
                physical_ranges_for_stream_range(
                    &source,
                    u64::try_from(*start + filepass.len()).unwrap(),
                    end - start,
                )
            })
            .collect::<Vec<_>>();
        source.clear_ranges();

        assert!(matches!(
            SourceBackedWorkbook::from_read_at(source.clone()),
            Err(SourceBackedError::EncryptedUnsupported)
        ));
        let actual = source.ranges();
        assert!(overlaps_any(&actual, &header_ranges));
        assert!(!overlaps_any(&actual, &payload_ranges));
        assert!(worksheet_ranges.iter().flatten().copied().all(|worksheet| {
            actual
                .iter()
                .copied()
                .all(|read| !ranges_overlap(read, worksheet))
        }));
    }
}

#[test]
fn truncated_global_header_is_rejected_without_overread() {
    let mut stream = frame_bytes(0x0809, &[0; 16]);
    stream.extend_from_slice(&[0x42, 0]);
    let bytes = cfb_with_streams(&[("Workbook", &stream)]);
    let source = Arc::new(CountingSource::new(bytes));
    let retained: Arc<dyn ReadAt> = source.clone();
    let cfb = SharedOleFile::open(Arc::clone(&retained)).unwrap();
    let catalog_ranges = source.ranges();
    drop(cfb);
    source.clear_ranges();
    let first_header_ranges = physical_ranges_for_stream_range(&source, 0, 4);
    source.clear_ranges();

    assert!(matches!(
        SourceBackedWorkbook::from_read_at(source.clone()),
        Err(SourceBackedError::InvalidData(message))
            if message == "truncated BIFF global record header"
    ));
    let mut expected = catalog_ranges;
    expected.extend(first_header_ranges);
    assert_eq!(source.ranges(), expected);
}

#[test]
fn workbook_and_book_selection_only_falls_back_for_missing_stream() {
    let stream = workbook_stream(fixture("Simple.xls"), "Workbook");
    let book_only = cfb_with_streams(&[("Book", &stream)]);
    let owner =
        SourceBackedWorkbook::from_read_at(Arc::new(CountingSource::new(book_only))).unwrap();
    assert!(owner.worksheet_count().unwrap() > 0);

    let malformed = frame_bytes(0x0809, &[0; 16]);
    let both = cfb_with_streams(&[("Workbook", &malformed), ("Book", &stream)]);
    let result = SourceBackedWorkbook::from_read_at(Arc::new(CountingSource::new(both)));
    assert!(matches!(
        result,
        Err(SourceBackedError::Cfb(_))
            | Err(SourceBackedError::Parse(_))
            | Err(SourceBackedError::InvalidData(_))
    ));
}

#[test]
fn worksheet_bof_version_and_substream_type_are_deferred_to_selected_access() {
    for (version, substream_type, expected) in [
        (0x0500_u16, 0x0010_u16, "version"),
        (0x0600_u16, 0x0020_u16, "substream"),
    ] {
        let mut stream = workbook_stream(fixture("Simple.xls"), "Workbook");
        let offset = first_sheet_offset(&stream);
        stream[offset + 4..offset + 6].copy_from_slice(&version.to_le_bytes());
        stream[offset + 6..offset + 8].copy_from_slice(&substream_type.to_le_bytes());
        let bytes = cfb_with_streams(&[("Workbook", &stream)]);
        let owner =
            SourceBackedWorkbook::from_read_at(Arc::new(CountingSource::new(bytes))).unwrap();
        let result = owner.cell_by_index(0, 0, 0);
        if expected == "version" {
            assert!(matches!(result, Err(SourceBackedError::Parse(_))));
        } else {
            assert!(matches!(result, Err(SourceBackedError::InvalidData(_))));
        }
    }
}

#[test]
fn raw_handoff_reads_global_headers_then_one_exact_global_range() {
    let workbook = workbook_stream(fixture("Simple.xls"), "Workbook");
    let global_end = global_eof_offset(&workbook) + 4;
    let mut header_offsets = Vec::new();
    let mut offset = 0_usize;
    loop {
        header_offsets.push(offset);
        let kind = u16::from_le_bytes([workbook[offset], workbook[offset + 1]]);
        let length = usize::from(u16::from_le_bytes([
            workbook[offset + 2],
            workbook[offset + 3],
        ]));
        offset += 4 + length;
        if kind == 0x000A {
            break;
        }
    }

    let source = Arc::new(CountingSource::new(cfb_with_streams(&[(
        "Workbook", &workbook,
    )])));
    let retained: Arc<dyn ReadAt> = source.clone();
    let cfb = Arc::new(SharedOleFile::open(Arc::clone(&retained)).unwrap());
    source.clear_ranges();
    let _owner = litchi_xls::raw::source_backed_workbook_from_shared_ole_file(
        cfb,
        SourceBackedLimits::default(),
    )
    .unwrap();
    let actual = source.ranges();

    let mut expected = Vec::new();
    for header_offset in header_offsets {
        expected.extend(physical_ranges_for_stream_range(
            &source,
            header_offset as u64,
            4,
        ));
    }
    expected.extend(physical_ranges_for_stream_range(&source, 0, global_end));
    assert_eq!(actual, expected);

    let sheet_ranges = physical_ranges_for_workbook_sheets(&source, &workbook);
    assert!(actual.iter().all(|range| {
        sheet_ranges
            .iter()
            .flatten()
            .copied()
            .all(|sheet| !ranges_overlap(*range, sheet))
    }));
}

#[test]
fn shared_cfb_source_arc_preserves_pointer_identity() {
    let source: Arc<dyn ReadAt> = Arc::new(CountingSource::new(fixture("Simple.xls")));
    let cfb = SharedOleFile::open(Arc::clone(&source)).unwrap();
    let retained = cfb.source_arc();
    assert!(Arc::ptr_eq(&source, &retained));
}

#[test]
fn raw_handoff_rejects_a_stale_retained_source_before_global_reads() {
    let source = Arc::new(CountingSource::new(fixture("Simple.xls")));
    let retained: Arc<dyn ReadAt> = source.clone();
    let cfb = Arc::new(SharedOleFile::open(Arc::clone(&retained)).unwrap());
    source.clear_ranges();
    source.bump();
    assert!(matches!(
        litchi_xls::raw::source_backed_workbook_from_shared_ole_file(
            cfb,
            SourceBackedLimits::default(),
        ),
        Err(SourceBackedError::SourceChanged { .. })
    ));
    assert!(source.ranges().is_empty());
}

#[test]
fn raw_handoff_enforces_input_limit_against_existing_cfb_size() {
    let source = Arc::new(CountingSource::new(fixture("Simple.xls")));
    let retained: Arc<dyn ReadAt> = source.clone();
    let cfb = Arc::new(SharedOleFile::open(Arc::clone(&retained)).unwrap());
    let maximum = cfb.file_size().saturating_sub(1);
    assert!(maximum > 0);
    source.clear_ranges();
    assert!(matches!(
        litchi_xls::raw::source_backed_workbook_from_shared_ole_file(
            cfb,
            SourceBackedLimits::default().with_max_input_bytes(maximum),
        ),
        Err(SourceBackedError::ResourceLimit {
            resource: "input bytes",
            ..
        })
    ));
    assert!(source.ranges().is_empty());
}

fn frame_bytes(kind: u16, payload: &[u8]) -> Vec<u8> {
    let mut bytes = Vec::with_capacity(4 + payload.len());
    bytes.extend_from_slice(&kind.to_le_bytes());
    bytes.extend_from_slice(&(u16::try_from(payload.len()).unwrap()).to_le_bytes());
    bytes.extend_from_slice(payload);
    bytes
}
