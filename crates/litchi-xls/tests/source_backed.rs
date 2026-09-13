use litchi_cfb::{OleWriter, SharedOleFile};
use litchi_core::sheet::{Cell as CellTrait, CellValue, WorkbookTrait};
use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionError, ExecutionLimits, Limits,
    OwnedSource, ReadAt, SourceVersion, TextOutputError, TextOutputOptions,
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

/// Upper bound the globals fill window doubles to, from `source.rs`.
const GLOBALS_MAX_WINDOW_BYTES: usize = 64 * 1024;
/// Globals records the scan reads one header and one payload at a time.
const GLOBALS_EXACT_PROLOGUE_RECORDS: usize = 4;

fn global_bof_frame_end(stream: &[u8]) -> usize {
    assert!(stream.len() >= 4);
    assert_eq!(
        u16::from_le_bytes([stream[0], stream[1]]),
        0x0809,
        "fixture globals do not start with BOF"
    );
    4 + usize::from(u16::from_le_bytes([stream[2], stream[3]]))
}

/// End of the globals record at `index`, counting BOF as index zero.
fn global_record_end(stream: &[u8], index: usize) -> usize {
    let mut offset = 0;
    for _ in 0..=index {
        assert!(offset + 4 <= stream.len());
        let length = usize::from(u16::from_le_bytes([stream[offset + 2], stream[offset + 3]]));
        offset += 4 + length;
    }
    offset
}

/// Inserts `extra` immediately after the globals record at `index`, patching
/// `BoundSheet8` stream positions the way [`insert_before_global_eof`] does.
fn insert_after_global_record(stream: &[u8], index: usize, extra: &[u8]) -> Vec<u8> {
    let at = global_record_end(stream, index);
    let eof = global_eof_offset(stream);
    assert!(at <= eof);
    let delta = u32::try_from(extra.len()).unwrap();
    let mut patched = stream.to_vec();
    let mut cursor = 0;
    while cursor < eof {
        assert!(cursor + 4 <= patched.len());
        let kind = u16::from_le_bytes([patched[cursor], patched[cursor + 1]]);
        let length = usize::from(u16::from_le_bytes([
            patched[cursor + 2],
            patched[cursor + 3],
        ]));
        if kind == 0x0085 && cursor + 8 <= patched.len() {
            let position = u32::from_le_bytes([
                patched[cursor + 4],
                patched[cursor + 5],
                patched[cursor + 6],
                patched[cursor + 7],
            ]);
            if usize::try_from(position).unwrap() >= eof {
                let shifted = position.checked_add(delta).unwrap();
                patched[cursor + 4..cursor + 8].copy_from_slice(&shifted.to_le_bytes());
            }
        }
        cursor += 4 + length;
    }
    let mut output = Vec::with_capacity(patched.len() + extra.len());
    output.extend_from_slice(&patched[..at]);
    output.extend_from_slice(extra);
    output.extend_from_slice(&patched[at..]);
    output
}

/// Index of the first globals record of `kind`, counting BOF as index zero.
fn global_record_index_of(stream: &[u8], kind: u16) -> usize {
    let eof = global_eof_offset(stream);
    let mut offset = 0;
    let mut index = 0;
    while offset < eof {
        let found = u16::from_le_bytes([stream[offset], stream[offset + 1]]);
        let length = usize::from(u16::from_le_bytes([stream[offset + 2], stream[offset + 3]]));
        if found == kind {
            return index;
        }
        offset += 4 + length;
        index += 1;
    }
    panic!("fixture has no record of kind {kind:#06x}");
}

/// Offsets of every `BoundSheet8` record in the globals.
fn bound_sheet_record_offsets(stream: &[u8]) -> Vec<usize> {
    let eof = global_eof_offset(stream);
    let mut offsets = Vec::new();
    let mut offset = 0;
    while offset < eof {
        let kind = u16::from_le_bytes([stream[offset], stream[offset + 1]]);
        let length = usize::from(u16::from_le_bytes([stream[offset + 2], stream[offset + 3]]));
        if kind == 0x0085 {
            offsets.push(offset);
        }
        offset += 4 + length;
    }
    offsets
}

/// Builds a fixture whose `BoundSheet8` positions descend and whose
/// `BoundSheet8` run is split wide enough that a window fill lands inside the
/// gap, so only a clamp that every `BoundSheet8` lowers keeps the fills off the
/// sheet bodies.
fn descending_bound_sheets_split_by_a_fill(stream: &[u8]) -> Vec<u8> {
    // Padding before the globals EOF keeps a fill pending after the last
    // BoundSheet8 has been framed; padding inside the run puts a fill boundary
    // between the first BoundSheet8 and the rest.
    let padded = insert_before_global_eof(stream, &frame_bytes(0x1234, &[0xA5; 2_000]));
    let first = global_record_index_of(&padded, 0x0085);
    let mut spread =
        insert_after_global_record(&padded, first, &frame_bytes(0x1234, &[0xA5; 2_000]));
    let offsets = bound_sheet_record_offsets(&spread);
    let mut positions = workbook_bound_sheet_positions(&spread);
    assert!(positions.len() >= 3);
    positions.reverse();
    for (offset, position) in offsets.iter().zip(positions) {
        spread[offset + 4..offset + 8]
            .copy_from_slice(&u32::try_from(position).unwrap().to_le_bytes());
    }
    spread
}

/// One past the last globals byte, so `[0, global_end)` is the globals range.
fn global_end_offset(stream: &[u8]) -> usize {
    global_eof_offset(stream) + 4
}

/// Bytes consumed by the exact prologue and an upper bound on the reads it
/// takes, derived from the fixture's own framing rather than from a recorded
/// schedule. The bound is loose: the prologue's payload fetch also covers the
/// next record's header, so it takes fewer reads than this.
fn globals_prologue(stream: &[u8]) -> (usize, usize) {
    let mut offset = 0;
    let mut reads = 0;
    for _ in 0..GLOBALS_EXACT_PROLOGUE_RECORDS {
        assert!(offset + 4 <= stream.len());
        let kind = u16::from_le_bytes([stream[offset], stream[offset + 1]]);
        let length = usize::from(u16::from_le_bytes([stream[offset + 2], stream[offset + 3]]));
        reads += 1;
        if length > 0 {
            reads += 1;
        }
        offset += 4 + length;
        if kind == 0x000A {
            break;
        }
    }
    (offset, reads)
}

/// Upper bound on window fills needed to cover `bytes`, given a first window of
/// one CFB sector doubling to [`GLOBALS_MAX_WINDOW_BYTES`].
fn globals_fill_bound(bytes: usize) -> usize {
    let mut covered = 0;
    let mut window = 512;
    let mut fills = 1;
    while covered < bytes {
        covered += window;
        window = (window * 2).min(GLOBALS_MAX_WINDOW_BYTES);
        fills += 1;
    }
    fills
}

/// Merges recorded reads into maximal disjoint physical spans.
fn merged_spans(ranges: &[(u64, usize)]) -> Vec<(u64, u64)> {
    let mut spans = ranges
        .iter()
        .map(|(offset, length)| (*offset, offset + *length as u64))
        .collect::<Vec<_>>();
    spans.sort_unstable();
    let mut merged = Vec::<(u64, u64)>::new();
    for (start, end) in spans {
        match merged.last_mut() {
            Some(last) if start <= last.1 => last.1 = last.1.max(end),
            _ => merged.push((start, end)),
        }
    }
    merged
}

/// True when every range in `inner` lies inside one span of `outer`.
fn spans_cover(outer: &[(u64, u64)], inner: &[(u64, usize)]) -> bool {
    inner.iter().all(|(offset, length)| {
        let end = offset + *length as u64;
        outer
            .iter()
            .any(|(start, stop)| *start <= *offset && end <= *stop)
    })
}

/// True when no two recorded reads share a byte.
fn reads_are_disjoint(ranges: &[(u64, usize)]) -> bool {
    let mut spans = ranges
        .iter()
        .map(|(offset, length)| (*offset, offset + *length as u64))
        .collect::<Vec<_>>();
    spans.sort_unstable();
    spans.windows(2).all(|pair| pair[0].1 <= pair[1].0)
}

/// Physical ranges of the Workbook stream at or after `position`.
fn physical_ranges_from(
    source: &CountingSource,
    stream: &[u8],
    position: usize,
) -> Vec<(u64, usize)> {
    physical_ranges_for_stream_range(source, position as u64, stream.len() - position)
}

/// Opens the globals through the raw handoff and returns only the reads the
/// globals scan itself performed, with the CFB catalog reads excluded.
fn globals_reads(
    source: &Arc<CountingSource>,
) -> (Vec<(u64, usize)>, Result<(), SourceBackedError>) {
    globals_reads_with_limits(source, SourceBackedLimits::default())
}

fn globals_reads_with_limits(
    source: &Arc<CountingSource>,
    limits: SourceBackedLimits,
) -> (Vec<(u64, usize)>, Result<(), SourceBackedError>) {
    let retained: Arc<dyn ReadAt> = source.clone();
    let cfb = Arc::new(SharedOleFile::open(Arc::clone(&retained)).unwrap());
    source.clear_ranges();
    let result = litchi_xls::raw::source_backed_workbook_from_shared_ole_file(cfb, limits);
    let ranges = source.ranges();
    (ranges, result.map(|_| ()))
}

/// Asserts the one-pass globals contract: every globals byte read exactly once,
/// nothing read past one window beyond the globals end, and a read count inside
/// the prologue-plus-window bound. Returns the bytes read past the globals end.
fn assert_globals_read_once(
    source: &CountingSource,
    stream: &[u8],
    actual: &[(u64, usize)],
) -> usize {
    let global_end = global_end_offset(stream);
    let globals_ranges = physical_ranges_for_stream_range(source, 0, global_end);
    assert!(
        reads_are_disjoint(actual),
        "two globals reads share a byte: {actual:?}"
    );
    assert!(
        spans_cover(&merged_spans(actual), &globals_ranges),
        "a globals byte was never read"
    );
    let bound = (global_end + GLOBALS_MAX_WINDOW_BYTES).min(stream.len());
    let allowed = merged_spans(&physical_ranges_for_stream_range(source, 0, bound));
    assert!(
        spans_cover(&allowed, actual),
        "a globals read reached past one window beyond the globals end"
    );
    let (prologue_end, prologue_reads) = globals_prologue(stream);
    let fills = globals_fill_bound(bound - prologue_end);
    assert!(
        actual.len() <= prologue_reads + fills + globals_ranges.len(),
        "{} reads exceed the prologue-plus-window bound {}",
        actual.len(),
        prologue_reads + fills + globals_ranges.len()
    );
    let read_bytes: usize = actual.iter().map(|(_, length)| *length).sum();
    assert!(read_bytes >= global_end);
    read_bytes - global_end
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

fn first_sst_payload_span(stream: &[u8]) -> (usize, usize) {
    let mut cursor = 0_usize;
    while cursor + 4 <= stream.len() {
        let kind = u16::from_le_bytes([stream[cursor], stream[cursor + 1]]);
        let length = usize::from(u16::from_le_bytes([stream[cursor + 2], stream[cursor + 3]]));
        if kind == 0x00FC {
            let start = cursor + 4;
            let mut end = start + length;
            cursor = end;
            while cursor + 4 <= stream.len()
                && u16::from_le_bytes([stream[cursor], stream[cursor + 1]]) == 0x003C
            {
                let continuation_length =
                    usize::from(u16::from_le_bytes([stream[cursor + 2], stream[cursor + 3]]));
                end = cursor + 4 + continuation_length;
                cursor = end;
            }
            return (start, end - start);
        }
        cursor += 4 + length;
    }
    panic!("fixture has no SST");
}

fn sst_header(stream: &[u8]) -> (usize, u32, u32) {
    let mut cursor = 0_usize;
    while cursor + 12 <= stream.len() {
        let kind = u16::from_le_bytes([stream[cursor], stream[cursor + 1]]);
        let length = usize::from(u16::from_le_bytes([stream[cursor + 2], stream[cursor + 3]]));
        if kind == 0x00FC {
            let payload = cursor + 4;
            return (
                payload,
                u32::from_le_bytes([
                    stream[payload],
                    stream[payload + 1],
                    stream[payload + 2],
                    stream[payload + 3],
                ]),
                u32::from_le_bytes([
                    stream[payload + 4],
                    stream[payload + 5],
                    stream[payload + 6],
                    stream[payload + 7],
                ]),
            );
        }
        cursor += 4 + length;
    }
    panic!("fixture has no SST");
}

fn label_sst_frame(row: u16, column: u16, string_index: u32) -> Vec<u8> {
    let mut payload = Vec::with_capacity(10);
    payload.extend_from_slice(&row.to_le_bytes());
    payload.extend_from_slice(&column.to_le_bytes());
    payload.extend_from_slice(&0_u16.to_le_bytes());
    payload.extend_from_slice(&string_index.to_le_bytes());
    frame_bytes(0x00FD, &payload)
}

#[derive(Clone)]
struct CountingSource {
    bytes: Arc<Vec<u8>>,
    ranges: Arc<Mutex<Vec<(u64, usize)>>>,
    cancel_on_read: Arc<Mutex<Option<CancellationSource>>>,
    revision: Arc<AtomicU64>,
    versions: Arc<AtomicU64>,
}

impl CountingSource {
    fn new(bytes: Vec<u8>) -> Self {
        Self {
            bytes: Arc::new(bytes),
            ranges: Arc::new(Mutex::new(Vec::new())),
            cancel_on_read: Arc::new(Mutex::new(None)),
            revision: Arc::new(AtomicU64::new(0)),
            versions: Arc::new(AtomicU64::new(0)),
        }
    }

    fn version_calls(&self) -> u64 {
        self.versions.load(Ordering::Relaxed)
    }

    fn clear_version_calls(&self) {
        self.versions.store(0, Ordering::Relaxed);
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
        self.versions.fetch_add(1, Ordering::Relaxed);
        Ok(SourceVersion::new(
            0x584c_535f_5445_5354,
            self.revision.load(Ordering::Relaxed),
        ))
    }
}

#[test]
fn retained_metadata_queries_observe_the_source_once() {
    let source = Arc::new(CountingSource::new(fixture("Simple.xls")));
    let owner = SourceBackedWorkbook::from_read_at(source.clone()).unwrap();
    // Acquiring the handle is itself one metadata query; hold it so the
    // counted region covers exactly one query each.
    let worksheet = owner.worksheet(0).unwrap().unwrap();

    let mut queries: Vec<(&str, Box<dyn Fn()>)> = Vec::new();
    queries.push((
        "worksheet_count",
        Box::new(|| {
            owner.worksheet_count().unwrap();
        }),
    ));
    queries.push((
        "worksheet_names",
        Box::new(|| {
            owner.worksheet_names().unwrap();
        }),
    ));
    queries.push((
        "worksheet_handle",
        Box::new(|| {
            owner.worksheet(0).unwrap().unwrap();
        }),
    ));
    queries.push((
        "worksheet_name",
        Box::new(|| {
            worksheet.name().unwrap();
        }),
    ));
    queries.push((
        "worksheet_visibility",
        Box::new(|| {
            worksheet.visibility().unwrap();
        }),
    ));

    for (query, run) in &queries {
        source.clear_ranges();
        source.clear_version_calls();
        run();
        assert_eq!(
            source.version_calls(),
            1,
            "{query} must fence the retained source exactly once"
        );
        assert_eq!(source.bytes_read(), 0, "{query} must read no source bytes");
    }
}

#[test]
fn retained_metadata_queries_still_refuse_a_changed_source() {
    let source = Arc::new(CountingSource::new(fixture("Simple.xls")));
    let owner = SourceBackedWorkbook::from_read_at(source.clone()).unwrap();
    owner.worksheet_count().unwrap();
    let worksheet = owner.worksheet(0).unwrap().unwrap();
    worksheet.name().unwrap();

    source.bump();

    assert!(
        matches!(
            owner.worksheet_count(),
            Err(SourceBackedError::SourceChanged { .. })
        ),
        "a changed source must still be refused by workbook metadata"
    );
    assert!(
        matches!(
            worksheet.name(),
            Err(SourceBackedError::SourceChanged { .. })
        ),
        "a changed source must still be refused by worksheet metadata"
    );
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
    let source = Arc::new(CountingSource::new(bytes.clone()));
    let owner = SourceBackedWorkbook::from_read_at(source.clone()).unwrap();
    let eager = Workbook::new(Cursor::new(bytes)).unwrap();
    let expected = eager
        .xls_worksheet(0)
        .unwrap()
        .get_cell(u32::from(row), u32::from(column))
        .map(CellTrait::value);
    assert!(matches!(expected, Some(CellValue::String(_))));
    source.clear_ranges();
    assert_eq!(
        owner
            .cell_value_by_index(0, u32::from(row), u32::from(column))
            .unwrap()
            .as_ref(),
        expected
    );
    let query_ranges = source.ranges();
    assert!(!query_ranges.is_empty());
}

#[test]
fn invalid_label_sst_index_is_typed_without_an_sst_read() {
    let original = workbook_stream(fixture("Simple.xls"), "Workbook");
    let (sst_offset, sst_length) = first_sst_payload_span(&original);
    let modified = insert_before_worksheet_eof(&original, &label_sst_frame(60_005, 6, u32::MAX));
    let bytes = cfb_with_streams(&[("Workbook", &modified)]);
    let source = Arc::new(CountingSource::new(bytes));
    let owner = SourceBackedWorkbook::from_read_at(source.clone()).unwrap();
    let sst_ranges = physical_ranges_for_stream_range(&source, sst_offset as u64, sst_length);
    let before_query = source.ranges();

    assert!(matches!(
        owner.cell_value_by_index(0, 60_005, 6).unwrap(),
        Some(CellValue::Error(message))
            if message.starts_with("Invalid SST index: 4294967295 (max: ")
    ));
    let after_query = source.ranges();
    assert!(after_query[before_query.len()..].iter().all(|range| {
        sst_ranges
            .iter()
            .all(|sst_range| !ranges_overlap(*range, *sst_range))
    }));
}

#[test]
fn exact_sst_entry_limit_is_allowed_and_one_over_is_rejected() {
    let original = workbook_stream(fixture("Simple.xls"), "Workbook");
    let (payload, total, unique) = sst_header(&original);
    let maximum = usize::try_from(total.max(unique)).unwrap();
    let exact = cfb_with_streams(&[("Workbook", &original)]);
    assert!(
        SourceBackedWorkbook::from_read_at_with_limits(
            Arc::new(CountingSource::new(exact)),
            SourceBackedLimits::default().with_max_sst_entries(maximum),
        )
        .is_ok()
    );

    let mut over_stream = original;
    let over = total.max(unique).checked_add(1).unwrap();
    over_stream[payload..payload + 4].copy_from_slice(&over.to_le_bytes());
    let source = Arc::new(CountingSource::new(cfb_with_streams(&[(
        "Workbook",
        &over_stream,
    )])));
    assert!(matches!(
        SourceBackedWorkbook::from_read_at_with_limits(
            source,
            SourceBackedLimits::default().with_max_sst_entries(maximum),
        ),
        Err(SourceBackedError::ResourceLimit {
            resource: "SST entries",
            ..
        })
    ));
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
    let (sst_offset, sst_length) = first_sst_payload_span(&workbook);
    let sst_ranges = physical_ranges_for_stream_range(&source, sst_offset as u64, sst_length);
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
    // Open frames the globals and retains their bytes in one pass. A window
    // fill may run past the globals end, so the contract is that open reads no
    // byte at or beyond the smallest BoundSheet8 stream position, and that its
    // total is the CFB catalog plus the globals plus at most one window.
    // The clamp is the minimum over the BoundSheet8 records framed so far; this
    // fixture's lbPlyPos ascend, so it equals the minimum over the whole
    // globals from the first BoundSheet8 on.
    let smallest_sheet = workbook_bound_sheet_positions(&workbook)
        .into_iter()
        .min()
        .unwrap();
    let sheet_tail = physical_ranges_from(&source, &workbook, smallest_sheet);
    assert!(!overlaps_any(&open_ranges, &sheet_tail));
    let catalog_probe = Arc::new(CountingSource::new(bytes.clone()));
    let catalog = SharedOleFile::open(catalog_probe.clone()).unwrap();
    let catalog_bytes = catalog_probe.bytes_read();
    drop(catalog);
    let open_bytes: usize = open_ranges.iter().map(|(_, length)| *length).sum();
    assert!(
        open_bytes <= catalog_bytes + global_end_offset(&workbook) + GLOBALS_MAX_WINDOW_BYTES,
        "open read {open_bytes} bytes"
    );
    for (query_ranges, selected_workbook_index) in [
        (&first_ranges, first_workbook_index),
        (&second_ranges, second_workbook_index),
    ] {
        assert!(overlaps_any(
            query_ranges,
            &sheet_ranges[selected_workbook_index]
        ));
        let allowed_ranges = sheet_ranges[selected_workbook_index]
            .iter()
            .chain(&sst_ranges)
            .copied()
            .collect::<Vec<_>>();
        assert!(query_ranges.iter().copied().all(|range| {
            allowed_ranges
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
fn out_of_order_late_packed_duplicate_keeps_the_latest_value() {
    let original = workbook_stream(fixture("Simple.xls"), "Workbook");
    let mut inserted = number_frame(0, 0, 0, 7.0);
    inserted.extend_from_slice(&number_frame(1, 0, 0, 8.0));
    inserted.extend_from_slice(&mul_rk_frame(0, 0, &[99, 101]));
    let modified = insert_before_worksheet_eof(&original, &inserted);
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
fn malformed_unselected_cell_after_a_match_is_refused() {
    let original = workbook_stream(fixture("Simple.xls"), "Workbook");
    let mut inserted = number_frame(0, 0, 7, 7.0);
    let mut malformed_payload = Vec::new();
    malformed_payload.extend_from_slice(&60_006_u16.to_le_bytes());
    malformed_payload.extend_from_slice(&7_u16.to_le_bytes());
    malformed_payload.extend_from_slice(&0_u16.to_le_bytes());
    malformed_payload.extend_from_slice(&1_u16.to_le_bytes());
    malformed_payload.push(0);
    inserted.extend_from_slice(&frame_bytes(0x0204, &malformed_payload));
    let modified = insert_before_worksheet_eof(&original, &inserted);
    let bytes = cfb_with_streams(&[("Workbook", &modified)]);
    let owner = SourceBackedWorkbook::from_read_at(Arc::new(CountingSource::new(bytes))).unwrap();

    assert!(matches!(
        owner.cell_by_index(0, 0, 0),
        Err(SourceBackedError::Parse(_))
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
fn supported_unknown_payload_is_skipped_without_source_overread() {
    let original = workbook_stream(fixture("Simple.xls"), "Workbook");
    let payload = vec![0xA5; 4_096];
    let unknown_offset = worksheet_eof_offset(&original, first_sheet_offset(&original));
    let modified = insert_before_worksheet_eof(&original, &frame_bytes(0x1234, &payload));
    let bytes = cfb_with_streams(&[("Workbook", &modified)]);
    let source = Arc::new(CountingSource::new(bytes.clone()));
    let owner = SourceBackedWorkbook::from_read_at(source.clone()).unwrap();
    let header_ranges = physical_ranges_for_stream_range(&source, unknown_offset as u64, 4);
    let payload_ranges =
        physical_ranges_for_stream_range(&source, unknown_offset as u64 + 4, payload.len());
    source.clear_ranges();

    let actual = owner.cell_value_by_index(0, 0, 0).unwrap();
    let eager = Workbook::new(Cursor::new(bytes)).unwrap();
    let expected = eager
        .xls_worksheet(0)
        .unwrap()
        .get_cell(0, 0)
        .map(CellTrait::value);
    assert_eq!(actual.as_ref(), expected);

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

/// Builds a `FilePass` frame whose header always claims 8,192 payload bytes.
fn filepass_frame(payload: &[u8]) -> Vec<u8> {
    let mut frame = Vec::with_capacity(4 + payload.len());
    frame.extend_from_slice(&0x002F_u16.to_le_bytes());
    frame.extend_from_slice(&8_192_u16.to_le_bytes());
    frame.extend_from_slice(payload);
    frame
}

#[test]
fn filepass_header_scan_never_reads_its_payload_or_worksheets() {
    let original = workbook_stream(fixture("Simple.xls"), "Workbook");
    // A FilePass framed after the exact prologue may have payload bytes
    // resident in the fill that carried its header; they are never framed,
    // interpreted or published. The bound below is the stream prefix that
    // framing the unmodified fixture's globals reads, so the refusal costs no
    // more of the stream than a plain open of the same fixture and the
    // declared payload is never walked.
    let plain = Arc::new(CountingSource::new(cfb_with_streams(&[(
        "Workbook", &original,
    )])));
    let (plain_ranges, plain_result) = globals_reads(&plain);
    assert!(plain_result.is_ok());
    let plain_bytes: usize = plain_ranges.iter().map(|(_, length)| *length).sum();
    for payload in [vec![0xA5; 8_192], vec![0xA5; 2]] {
        let filepass = filepass_frame(&payload);
        let filepass_offset = global_eof_offset(&original) as u64;
        let modified = insert_before_global_eof(&original, &filepass);
        let bytes = cfb_with_streams(&[("Workbook", &modified)]);
        let source = Arc::new(CountingSource::new(bytes));

        let header_ranges = physical_ranges_for_stream_range(&source, filepass_offset, 4);
        let allowed = merged_spans(&physical_ranges_for_stream_range(&source, 0, plain_bytes));

        let (actual, result) = globals_reads(&source);
        assert!(matches!(
            result,
            Err(SourceBackedError::EncryptedUnsupported)
        ));
        assert!(overlaps_any(&actual, &header_ranges));
        assert!(reads_are_disjoint(&actual));
        assert!(
            spans_cover(&allowed, &actual),
            "the refusal read past the prefix a plain open reads: {actual:?}"
        );
        let read_bytes: usize = actual.iter().map(|(_, length)| *length).sum();
        assert!(read_bytes <= plain_bytes);
    }
}

#[test]
fn spec_position_filepass_is_refused_before_its_payload_is_read() {
    let original = workbook_stream(fixture("Simple.xls"), "Workbook");
    let smallest_sheet = workbook_bound_sheet_positions(&original)
        .into_iter()
        .min()
        .unwrap();
    // [MS-XLS] 2.1.7.20.1 places FilePass immediately after BOF, so index one
    // is the spec position. The exact prologue covers indices one to three and
    // its last fetch also buffers the header of index four, so a FilePass at
    // any of those four positions is refused before a fill is ever issued and
    // no byte of its payload and no worksheet byte is read.
    for after in 0..GLOBALS_EXACT_PROLOGUE_RECORDS {
        for payload in [vec![0xA5; 8_192], vec![0xA5; 2]] {
            let filepass = filepass_frame(&payload);
            let filepass_offset = global_record_end(&original, after) as u64;
            let modified = insert_after_global_record(&original, after, &filepass);
            let bytes = cfb_with_streams(&[("Workbook", &modified)]);
            let source = Arc::new(CountingSource::new(bytes));

            let header_ranges = physical_ranges_for_stream_range(&source, filepass_offset, 4);
            let payload_ranges =
                physical_ranges_for_stream_range(&source, filepass_offset + 4, payload.len());
            let sheet_tail =
                physical_ranges_from(&source, &modified, smallest_sheet + filepass.len());

            let (actual, result) = globals_reads(&source);
            assert!(matches!(
                result,
                Err(SourceBackedError::EncryptedUnsupported)
            ));
            assert!(overlaps_any(&actual, &header_ranges));
            assert!(
                !overlaps_any(&actual, &payload_ranges),
                "FilePass at index {} had payload bytes read",
                after + 1
            );
            assert!(!overlaps_any(&actual, &sheet_tail));
        }
    }
}

#[test]
fn encrypted_fixtures_are_refused_without_reading_a_filepass_payload() {
    // Every FilePass carrier in the corpus places the record at stream offset
    // 20, immediately after BOF, so the globals scan stops with the buffer
    // holding exactly the BOF frame and the FilePass header.
    for bytes in [
        ole_fixture("password.xls"),
        fixture("xor-encryption-abc.xls"),
        fixture("35897-type4.xls"),
    ] {
        let stream = workbook_stream(bytes.clone(), "Workbook");
        let filepass_offset = global_bof_frame_end(&stream);
        assert_eq!(
            u16::from_le_bytes([stream[filepass_offset], stream[filepass_offset + 1]]),
            0x002F
        );
        let source = Arc::new(CountingSource::new(bytes));
        let (actual, result) = globals_reads(&source);
        assert!(matches!(
            result,
            Err(SourceBackedError::EncryptedUnsupported)
        ));
        let framed = physical_ranges_for_stream_range(&source, 0, filepass_offset + 4);
        assert!(
            spans_cover(&merged_spans(&framed), &actual),
            "the scan read past the FilePass header: {actual:?}"
        );
        // The header-only pre-pass this replaced also took two reads to refuse
        // these fixtures, so the refusal costs no more than it did.
        assert_eq!(actual.len(), 2, "{actual:?}");
    }
}

#[test]
fn truncated_global_header_is_rejected_without_reading_past_the_stream() {
    let mut stream = frame_bytes(0x0809, &[0; 16]);
    stream.extend_from_slice(&[0x42, 0]);
    let framed_end = global_record_end(&stream, 0);
    let source = Arc::new(CountingSource::new(cfb_with_streams(&[(
        "Workbook", &stream,
    )])));

    // The exact prologue reads the BOF header, then one fetch covering its
    // payload and the next record's four-byte header, clamped by the stream
    // length. The second record's header is then refused as truncated, so the
    // scan reads at most the framed record plus one header and nothing at all
    // past the stream.
    let whole_stream = merged_spans(&physical_ranges_for_stream_range(&source, 0, stream.len()));
    let (actual, result) = globals_reads(&source);
    assert!(matches!(
        result,
        Err(SourceBackedError::InvalidData(message))
            if message == "truncated BIFF global record header"
    ));
    assert!(
        spans_cover(&whole_stream, &actual),
        "a read passed the Workbook stream: {actual:?}"
    );
    assert!(reads_are_disjoint(&actual));
    let read_bytes: usize = actual.iter().map(|(_, length)| *length).sum();
    assert!(read_bytes <= framed_end + 4, "{read_bytes} bytes read");
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
fn raw_handoff_over_read_is_bounded_by_the_stream_on_a_small_fixture() {
    // This fixture's Workbook stream is 4 KiB, so the binding bound on the
    // over-read is the stream, not the window: no BoundSheet8 has been framed
    // when the last fill is issued, and the fill stops at the stream end or
    // sooner. The window bound biting is covered by the large fixture below.
    let workbook = workbook_stream(fixture("Simple.xls"), "Workbook");
    let global_end = global_end_offset(&workbook);
    assert!(workbook.len() < global_end + GLOBALS_MAX_WINDOW_BYTES);
    let source = Arc::new(CountingSource::new(cfb_with_streams(&[(
        "Workbook", &workbook,
    )])));
    let (actual, result) = globals_reads(&source);
    assert!(result.is_ok());

    let over_read = assert_globals_read_once(&source, &workbook, &actual);
    assert!(over_read < workbook.len() - global_end);

    // The fill reaches into the first sheet body and stops well before the
    // next one, which is the bound a stream-clamped fill can promise.
    let mut positions = workbook_bound_sheet_positions(&workbook);
    positions.sort_unstable();
    assert!(positions.len() >= 2);
    let second_sheet = physical_ranges_from(&source, &workbook, positions[1]);
    assert!(!overlaps_any(&actual, &second_sheet));
}

#[test]
fn source_backed_open_reads_each_global_byte_once() {
    // The globals of this fixture are 621 records and 551,377 bytes inside a
    // 1.3 MB stream, so the one-window bound is well inside the stream and
    // bites. Its first BoundSheet8 is framed long before the fills reach the
    // globals end, so the clamp holds and no byte at or beyond the smallest
    // sheet position is read. Before change 0565 each header was read twice.
    let bytes = ole_fixture("ConditionalFormattingSamples.xls");
    let workbook = workbook_stream(bytes.clone(), "Workbook");
    let global_end = global_end_offset(&workbook);
    assert!(workbook.len() > global_end + GLOBALS_MAX_WINDOW_BYTES);
    let source = Arc::new(CountingSource::new(bytes));
    let (actual, result) = globals_reads(&source);
    assert!(result.is_ok());

    let over_read = assert_globals_read_once(&source, &workbook, &actual);
    assert_eq!(over_read, 0);
    // This fixture's lbPlyPos ascend, so the running minimum the fill clamp
    // keeps equals the minimum over the whole globals from the first
    // BoundSheet8 on. The descending case is covered separately.
    let smallest_sheet = workbook_bound_sheet_positions(&workbook)
        .into_iter()
        .min()
        .unwrap();
    let sheet_tail = physical_ranges_from(&source, &workbook, smallest_sheet);
    assert!(!overlaps_any(&actual, &sheet_tail));
    // Headroom over the count this schedule actually takes.
    assert!(actual.len() <= 64, "{} globals reads", actual.len());
}

#[test]
fn mini_stream_workbook_globals_are_read_once_each() {
    // A Workbook stream below the 4,096-byte CFB cutoff lives in the mini
    // stream, so the fills run through the MiniFAT range reader.
    let bytes = ole_fixture("SimpleWithColours.xls");
    let workbook = workbook_stream(bytes.clone(), "Workbook");
    assert!(workbook.len() < 4_096);
    let source = Arc::new(CountingSource::new(bytes));
    let (actual, result) = globals_reads(&source);
    assert!(result.is_ok());
    let over_read = assert_globals_read_once(&source, &workbook, &actual);
    assert!(over_read <= GLOBALS_MAX_WINDOW_BYTES);
}

#[test]
fn corrupt_bound_sheet_positions_terminate_the_globals_scan() {
    // One position is zero and one points inside the globals, so the fill clamp
    // is contradicted by framing that has already passed it and is dropped for
    // the rest of the scan. With only the stream and `max_global_bytes` left,
    // the scan still stops within one window of the globals end rather than
    // running to the stream end, and reports the offset error the semantic pass
    // reported before change 0565.
    let original = workbook_stream(fixture("Simple.xls"), "Workbook");
    let mut stream = original.clone();
    let offsets = bound_sheet_record_offsets(&stream);
    assert!(offsets.len() >= 3);
    stream[offsets[0] + 4..offsets[0] + 8].copy_from_slice(&0_u32.to_le_bytes());
    stream[offsets[1] + 4..offsets[1] + 8].copy_from_slice(&100_u32.to_le_bytes());
    let global_end = global_end_offset(&stream);
    let source = Arc::new(CountingSource::new(cfb_with_streams(&[(
        "Workbook", &stream,
    )])));
    let (actual, result) = globals_reads(&source);
    assert!(matches!(
        result,
        Err(SourceBackedError::InvalidData(message))
            if message == "BoundSheet8 stream offset is outside the Workbook stream"
    ));
    // The framing itself completed, so the one-pass contract still applies.
    let over_read = assert_globals_read_once(&source, &stream, &actual);
    assert!(
        over_read < stream.len() - global_end,
        "the dropped clamp let the scan run to the stream end"
    );
}

#[test]
fn every_bound_sheet_lowers_the_globals_fill_clamp() {
    // The fixture's BoundSheet8 run is split by padding wide enough that a
    // window fill boundary lands between the first record and the rest, and the
    // declared positions descend, so the first BoundSheet8 alone gives a clamp
    // that is too high. A later fill is still pending when the last BoundSheet8
    // lowers the clamp to the true smallest position; only a clamp that every
    // BoundSheet8 lowers keeps that fill off the sheet bodies.
    let stream = descending_bound_sheets_split_by_a_fill(&workbook_stream(
        fixture("Simple.xls"),
        "Workbook",
    ));
    let offsets = bound_sheet_record_offsets(&stream);
    let positions = workbook_bound_sheet_positions(&stream);
    assert!(positions.windows(2).all(|pair| pair[0] > pair[1]));
    assert!(
        offsets[1] - offsets[0] > 2_000,
        "the BoundSheet8 run was not split"
    );

    let source = Arc::new(CountingSource::new(cfb_with_streams(&[(
        "Workbook", &stream,
    )])));
    let (actual, result) = globals_reads(&source);
    assert!(result.is_ok(), "{result:?}");
    assert_eq!(assert_globals_read_once(&source, &stream, &actual), 0);
    let smallest_sheet = positions.iter().copied().min().unwrap();
    let sheet_tail = physical_ranges_from(&source, &stream, smallest_sheet);
    assert!(!overlaps_any(&actual, &sheet_tail));
}

#[test]
fn mid_globals_byte_limit_reads_at_most_one_header_past_the_limit() {
    // `max_global_bytes` bounds retained globals, not the four header bytes
    // that prove a record crosses it: the header at the limit is read and the
    // record's own end is then reported, which is the order and the `observed`
    // value the header-only pre-pass produced. The limit is a record boundary
    // inside the window phase, so it is reached through a clamped fill rather
    // than through the exact prologue.
    let stream = workbook_stream(fixture("Simple.xls"), "Workbook");
    let mut offset = 0;
    while offset < 1_024 {
        let length = usize::from(u16::from_le_bytes([stream[offset + 2], stream[offset + 3]]));
        offset += 4 + length;
    }
    let limit = offset;
    let length = usize::from(u16::from_le_bytes([stream[limit + 2], stream[limit + 3]]));
    let observed = limit + 4 + length;
    assert!(limit < global_end_offset(&stream));

    let source = Arc::new(CountingSource::new(cfb_with_streams(&[(
        "Workbook", &stream,
    )])));
    let allowed = merged_spans(&physical_ranges_for_stream_range(&source, 0, limit + 4));
    let limits = SourceBackedLimits::default().with_max_global_bytes(limit as u64);
    let (actual, result) = globals_reads_with_limits(&source, limits);
    assert!(
        matches!(
            result,
            Err(SourceBackedError::ResourceLimit {
                resource: "global bytes",
                observed: seen,
                maximum,
            }) if seen == observed as u64 && maximum == limit as u64
        ),
        "{result:?}"
    );
    assert!(
        spans_cover(&allowed, &actual),
        "a read passed four bytes beyond the limit: {actual:?}"
    );
    let read_bytes: usize = actual.iter().map(|(_, length)| *length).sum();
    assert!((limit..=limit + 4).contains(&read_bytes), "{read_bytes}");
}

fn append_eager_text_cell(output: &mut String, value: &CellValue) {
    match value {
        CellValue::Empty => {},
        CellValue::Bool(value) => output.push_str(if *value { "TRUE" } else { "FALSE" }),
        CellValue::Int(value) => output.push_str(&value.to_string()),
        CellValue::Float(value) | CellValue::DateTime(value) => {
            output.push_str(&value.to_string());
        },
        CellValue::String(value) | CellValue::Error(value) => output.push_str(value),
        CellValue::Formula {
            formula,
            cached_value,
            ..
        } => match cached_value.as_deref() {
            Some(value) => append_eager_text_cell(output, value),
            None => {
                output.push('=');
                output.push_str(formula);
            },
        },
    }
}

fn eager_text(workbook: &Workbook<Cursor<Vec<u8>>>) -> String {
    let mut output = String::new();
    for worksheet_index in 0..workbook.worksheet_count() {
        let worksheet = workbook.worksheet_by_index(worksheet_index).unwrap();
        let mut rows = worksheet.rows();
        while let Some(row) = rows.next() {
            let row = row.unwrap();
            for (column, cell) in row.iter().enumerate() {
                if column != 0 {
                    output.push('\t');
                }
                append_eager_text_cell(&mut output, cell);
            }
            output.push('\n');
        }
    }
    output
}

fn assert_source_text_matches_eager(bytes: Vec<u8>) {
    let owner =
        SourceBackedWorkbook::from_read_at(Arc::new(CountingSource::new(bytes.clone()))).unwrap();
    let eager = Workbook::new(Cursor::new(bytes)).unwrap();
    assert_eq!(owner.text().unwrap(), eager_text(&eager));
}

#[test]
fn source_text_sink_matches_legacy_text_and_reports_output_limits() {
    let owner =
        SourceBackedWorkbook::from_read_at(Arc::new(CountingSource::new(fixture("Simple.xls"))))
            .unwrap();
    let legacy = owner.text().unwrap();
    assert!(legacy.ends_with("\n"));

    let mut output = Vec::new();
    let report = owner
        .write_text_to(&mut output, TextOutputOptions::default())
        .unwrap();
    assert_eq!(output, legacy.strip_suffix("\n").unwrap().as_bytes());
    assert_eq!(report.bytes_written(), output.len() as u64);
    assert!(report.objects_written() > 0);

    let mut limited = Vec::new();
    let error = owner
        .write_text_to(
            &mut limited,
            TextOutputOptions::new("\n", "\n\n", 1, u64::MAX),
        )
        .unwrap_err();
    assert!(matches!(error, TextOutputError::Limit { .. }));
    assert_eq!(error.progress().bytes_written(), limited.len() as u64);
}

#[test]
fn source_text_matches_eager_formula_and_out_of_order_packed_duplicate() {
    assert_source_text_matches_eager(ole_fixture("FormulaSheetRange.xls"));

    let original = workbook_stream(fixture("Simple.xls"), "Workbook");
    let mut inserted = number_frame(0, 0, 0, 7.0);
    inserted.extend_from_slice(&number_frame(1, 0, 0, 8.0));
    inserted.extend_from_slice(&mul_rk_frame(0, 0, &[99, 101]));
    let modified = insert_before_worksheet_eof(&original, &inserted);
    assert_source_text_matches_eager(cfb_with_streams(&[("Workbook", &modified)]));
}

#[test]
fn source_text_honors_valid_dimensions_and_retained_text_limits() {
    let original = workbook_stream(fixture("Simple.xls"), "Workbook");
    let mut dimensions = Vec::new();
    dimensions.extend_from_slice(&0_u32.to_le_bytes());
    dimensions.extend_from_slice(&13_u32.to_le_bytes());
    dimensions.extend_from_slice(&0_u16.to_le_bytes());
    dimensions.extend_from_slice(&2_u16.to_le_bytes());
    dimensions.extend_from_slice(&0_u16.to_le_bytes());
    let mut extra = frame_bytes(0x0200, &dimensions);
    extra.extend_from_slice(&number_frame(10, 0, 0, 101.0));
    extra.extend_from_slice(&number_frame(11, 0, 0, 102.0));
    let modified = insert_before_worksheet_eof(&original, &extra);
    let bytes = cfb_with_streams(&[("Workbook", &modified)]);
    assert_source_text_matches_eager(bytes.clone());

    let mut legacy_dimensions = Vec::new();
    legacy_dimensions.extend_from_slice(&0_u16.to_le_bytes());
    legacy_dimensions.extend_from_slice(&13_u16.to_le_bytes());
    legacy_dimensions.extend_from_slice(&0_u16.to_le_bytes());
    legacy_dimensions.extend_from_slice(&2_u16.to_le_bytes());
    legacy_dimensions.extend_from_slice(&0_u16.to_le_bytes());
    let modified = insert_before_worksheet_eof(&original, &frame_bytes(0x0200, &legacy_dimensions));
    assert_source_text_matches_eager(cfb_with_streams(&[("Workbook", &modified)]));

    let owner = SourceBackedWorkbook::from_read_at_with_limits(
        Arc::new(CountingSource::new(bytes.clone())),
        SourceBackedLimits::default().with_max_text_cells(1),
    )
    .unwrap();
    assert!(matches!(
        owner.text(),
        Err(SourceBackedError::ResourceLimit {
            resource: "text cells",
            ..
        })
    ));

    let owner = SourceBackedWorkbook::from_read_at_with_limits(
        Arc::new(CountingSource::new(bytes)),
        SourceBackedLimits::default().with_max_text_bytes(1),
    )
    .unwrap();
    assert!(matches!(
        owner.text(),
        Err(SourceBackedError::ResourceLimit {
            resource: "text bytes",
            ..
        })
    ));
}

#[test]
fn source_text_rejects_malformed_recognized_tail_and_bare_string_without_rows() {
    let original = workbook_stream(fixture("Simple.xls"), "Workbook");
    let malformed =
        insert_before_worksheet_eof(&original, &frame_bytes(0x0204, &[0, 0, 0, 0, 1, 0]));
    let mut output = Vec::new();
    let owner = SourceBackedWorkbook::from_read_at(Arc::new(CountingSource::new(
        cfb_with_streams(&[("Workbook", &malformed)]),
    )))
    .unwrap();
    let error = owner
        .write_text_to(&mut output, TextOutputOptions::default())
        .unwrap_err();
    assert!(matches!(error, TextOutputError::Document { .. }));
    assert!(output.is_empty());
    assert_eq!(error.progress().objects_written(), 0);

    let bare_string = insert_before_worksheet_eof(&original, &frame_bytes(0x0207, &[]));
    let owner = SourceBackedWorkbook::from_read_at(Arc::new(CountingSource::new(
        cfb_with_streams(&[("Workbook", &bare_string)]),
    )))
    .unwrap();
    let mut output = Vec::new();
    let error = owner
        .write_text_to(&mut output, TextOutputOptions::default())
        .unwrap_err();
    assert!(matches!(
        error,
        TextOutputError::Document {
            source: SourceBackedError::InvalidData(_),
            ..
        }
    ));
    assert!(output.is_empty());
    assert_eq!(error.progress().objects_written(), 0);
}

fn frame_bytes(kind: u16, payload: &[u8]) -> Vec<u8> {
    let mut bytes = Vec::with_capacity(4 + payload.len());
    bytes.extend_from_slice(&kind.to_le_bytes());
    bytes.extend_from_slice(&(u16::try_from(payload.len()).unwrap()).to_le_bytes());
    bytes.extend_from_slice(payload);
    bytes
}
