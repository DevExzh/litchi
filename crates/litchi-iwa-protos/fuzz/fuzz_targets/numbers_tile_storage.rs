#![no_main]

use libfuzzer_sys::fuzz_target;
use litchi_iwa_protos::numbers_table_cell_storage_codec::{
    CellSpans, DecodeError, DecodeOptions, StorageVisitor, TileRowInfoSnapshot,
    decode_tile_with_report, decode_tile_with_visitor,
};
use litchi_iwa_protos::tst::{Tile, TileRowInfo};
use prost::Message as _;

// Keep both libFuzzer's input and the codec's aggregate accounting finite. The
// target deliberately skips larger mutations instead of truncating them: the
// bytes passed to both decode paths are always one unchanged source.
const MAX_INPUT_BYTES: usize = 64 * 1024;
const MAX_FIELDS: usize = 8 * 1024;
const MAX_WORK_BYTES: usize = 256 * 1024;
const MAX_REFERENCES: usize = 1024;
const MAX_TEXT_BYTES: usize = 64 * 1024;
const MAX_RECURSION: u32 = 64;
// A row has at least the required fields, so this is a deliberately small
// local staging cap derived from the finite field ceiling. Pointer checks still
// run for every callback after this cap; only compact summaries are retained.
const MAX_RETAINED_ROWS: usize = MAX_FIELDS / 32;

#[derive(Clone)]
struct ObservedSlice {
    pointer: usize,
    length: usize,
}

#[derive(Clone)]
struct ObservedRow {
    ordinal: usize,
    tile_row_index: u32,
    cell_count: u32,
    storage_version: Option<u32>,
    has_wide_offsets: Option<bool>,
    cell_storage_buffer_pre_bnc: ObservedSlice,
    cell_offsets_pre_bnc: ObservedSlice,
    cell_storage_buffer: Option<ObservedSlice>,
    cell_offsets: Option<ObservedSlice>,
}

struct RowCollector<'expected> {
    rows: Vec<ObservedRow>,
    source_start: usize,
    source_end: usize,
    expected_rows: Option<&'expected [TileRowInfo]>,
    total_rows: usize,
    truncated: bool,
}

impl<'expected> RowCollector<'expected> {
    fn new(source: &[u8], expected_rows: Option<&'expected [TileRowInfo]>) -> Self {
        let source_start = source.as_ptr() as usize;
        let source_end = source_start
            .checked_add(source.len())
            .expect("bounded source pointer range");
        Self {
            rows: Vec::with_capacity(MAX_RETAINED_ROWS),
            source_start,
            source_end,
            expected_rows,
            total_rows: 0,
            truncated: false,
        }
    }

    fn assert_borrowed(&self, payload: &[u8]) {
        if payload.is_empty() {
            return;
        }
        let payload_start = payload.as_ptr() as usize;
        let payload_end = payload_start
            .checked_add(payload.len())
            .expect("bounded payload pointer range");
        assert!(
            payload_start >= self.source_start && payload_end <= self.source_end,
            "row payload does not borrow from the source"
        );
    }
}

impl<'expected> StorageVisitor for RowCollector<'expected> {
    fn visit_tile_row(&mut self, row: TileRowInfoSnapshot<'_>) -> Result<(), DecodeError> {
        self.assert_borrowed(row.cell_storage_buffer_pre_bnc());
        self.assert_borrowed(row.cell_offsets_pre_bnc());
        if let Some(payload) = row.cell_storage_buffer() {
            self.assert_borrowed(payload);
        }
        if let Some(payload) = row.cell_offsets() {
            self.assert_borrowed(payload);
        }

        let ordinal = self.total_rows;
        if let Some(expected_rows) = self.expected_rows {
            let expected = expected_rows
                .get(ordinal)
                .unwrap_or_else(|| panic!("streamed row count exceeded Prost row count"));
            assert_row_matches(row, expected);
        }
        self.total_rows = self.total_rows.saturating_add(1);
        if self.rows.len() < MAX_RETAINED_ROWS {
            self.rows.push(ObservedRow {
                ordinal,
                tile_row_index: row.tile_row_index(),
                cell_count: row.cell_count(),
                storage_version: row.storage_version(),
                has_wide_offsets: row.has_wide_offsets(),
                cell_storage_buffer_pre_bnc: observe_slice(row.cell_storage_buffer_pre_bnc()),
                cell_offsets_pre_bnc: observe_slice(row.cell_offsets_pre_bnc()),
                cell_storage_buffer: row.cell_storage_buffer().map(observe_slice),
                cell_offsets: row.cell_offsets().map(observe_slice),
            });
        } else {
            self.truncated = true;
        }
        Ok(())
    }
}

fn observe_slice(bytes: &[u8]) -> ObservedSlice {
    ObservedSlice {
        pointer: bytes.as_ptr() as usize,
        length: bytes.len(),
    }
}

fuzz_target!(|data: &[u8]| {
    let Some(source) = normalize_input(data) else {
        return;
    };
    exercise(&source);
    let wide = source.first().is_some_and(|byte| byte & 1 != 0);
    exercise_spans(
        &source,
        64,
        wide,
        source.first().copied().unwrap_or(0) as usize,
        32,
    );
    let baseline = [0, 0, 255, 255, 8, 0, 255, 255, 16, 0];
    exercise_spans(&baseline, 96, wide, 3, 5);
    let mut mutated = baseline;
    if let Some(byte) = source.first() {
        let index = source.get(1).copied().unwrap_or(0) as usize % mutated.len();
        mutated[index] = *byte;
    }
    exercise_spans(&mutated, 96, wide, 3, 5);
});

fn normalize_input(data: &[u8]) -> Option<Vec<u8>> {
    if let Some(encoded) = data.strip_prefix(b"hex:") {
        return decode_hex(encoded);
    }
    (data.len() <= MAX_INPUT_BYTES).then(|| data.to_vec())
}

fn decode_hex(encoded: &[u8]) -> Option<Vec<u8>> {
    if encoded.len() > MAX_INPUT_BYTES.saturating_mul(2).saturating_add(16) {
        return None;
    }
    let mut output = Vec::with_capacity(encoded.len() / 2);
    let mut high = None;
    for byte in encoded.iter().copied() {
        if byte.is_ascii_whitespace() {
            continue;
        }
        let nibble = hex_nibble(byte)?;
        if let Some(high_nibble) = high.take() {
            output.push((high_nibble << 4) | nibble);
            if output.len() > MAX_INPUT_BYTES {
                return None;
            }
        } else {
            high = Some(nibble);
        }
    }
    high.is_none().then_some(output)
}

fn hex_nibble(byte: u8) -> Option<u8> {
    match byte {
        b'0'..=b'9' => Some(byte - b'0'),
        b'a'..=b'f' => Some(byte - b'a' + 10),
        b'A'..=b'F' => Some(byte - b'A' + 10),
        _ => None,
    }
}

fn options() -> DecodeOptions {
    DecodeOptions::new(
        MAX_INPUT_BYTES,
        MAX_FIELDS,
        MAX_WORK_BYTES,
        MAX_RECURSION,
        MAX_REFERENCES,
        MAX_TEXT_BYTES,
    )
}

fn exercise(source: &[u8]) {
    let before = source.to_vec();
    let options = options();

    // Keep both calls on precisely the same borrowed source. In particular, do
    // not decode one path from a reconstructed or generated Prost value: every
    // visitor callback checks row payload identity against this source.
    let scalar = decode_tile_with_report(source, options);
    // Strict acceptance is authoritative. Only after it succeeds do we invoke
    // Prost; a Prost rejection here is a differential failure and must panic.
    let prost = scalar.as_ref().ok().map(|_| {
        Tile::decode(source)
            .unwrap_or_else(|error| panic!("strict tile acceptance disagreed with Prost: {error}"))
    });
    let expected_rows = prost.as_ref().map(|tile| tile.row_infos.as_slice());
    let mut visitor = RowCollector::new(source, expected_rows);
    let streamed = decode_tile_with_visitor(source, options, &mut visitor);

    assert_eq!(source, before.as_slice(), "decoder modified its source");

    match (scalar, streamed) {
        (Ok((scalar, scalar_report)), Ok((streamed, streamed_report))) => {
            assert_eq!(streamed, scalar);
            assert_eq!(scalar_report, streamed_report);
            let prost = prost
                .as_ref()
                .expect("strict scalar success must have a Prost oracle");
            assert_tile_matches(scalar, prost);
            assert_visitor_rows(&visitor);
        },
        (Err(_), Err(_)) => {
            // No decode result is consumed on this path. The visitor may have
            // observed rows before a later malformed field, but those local
            // summaries are never published as a successful decode.
        },
        _ => panic!("tile decode paths disagreed on success or failure"),
    }
}

fn assert_tile_matches(
    strict: litchi_iwa_protos::numbers_table_cell_storage_codec::TileSnapshot,
    prost: &Tile,
) {
    assert_eq!(strict.max_column(), prost.max_column);
    assert_eq!(strict.max_row(), prost.max_row);
    assert_eq!(strict.num_cells(), prost.num_cells);
    assert_eq!(strict.num_rows(), prost.numrows);
    assert_eq!(strict.storage_version(), prost.storage_version);
    assert_eq!(strict.last_saved_in_bnc(), prost.last_saved_in_bnc);
    assert_eq!(strict.should_use_wide_rows(), prost.should_use_wide_rows);
}

fn assert_row_matches(strict: TileRowInfoSnapshot<'_>, prost: &TileRowInfo) {
    assert_eq!(strict.tile_row_index(), prost.tile_row_index);
    assert_eq!(strict.cell_count(), prost.cell_count);
    assert_eq!(strict.storage_version(), prost.storage_version);
    assert_eq!(strict.has_wide_offsets(), prost.has_wide_offsets);
    assert_eq!(
        strict.cell_storage_buffer_pre_bnc(),
        prost.cell_storage_buffer_pre_bnc.as_slice()
    );
    assert_eq!(
        strict.cell_offsets_pre_bnc(),
        prost.cell_offsets_pre_bnc.as_slice()
    );
    assert_eq!(
        strict.cell_storage_buffer(),
        prost.cell_storage_buffer.as_deref()
    );
    assert_eq!(strict.cell_offsets(), prost.cell_offsets.as_deref());
}

fn assert_visitor_rows(visitor: &RowCollector<'_>) {
    assert!(visitor.rows.len() <= MAX_RETAINED_ROWS);
    if let Some(expected_rows) = visitor.expected_rows {
        assert_eq!(
            visitor.total_rows,
            expected_rows.len(),
            "streamed row count differed from Prost source order"
        );
    }
    if visitor.truncated {
        assert!(visitor.total_rows > MAX_RETAINED_ROWS);
    } else {
        assert_eq!(visitor.total_rows, visitor.rows.len());
    }
    for (index, row) in visitor.rows.iter().enumerate() {
        assert_eq!(row.ordinal, index);
        std::hint::black_box((
            row.tile_row_index,
            row.cell_count,
            row.storage_version,
            row.has_wide_offsets,
        ));
        // The callback did the source-range check before this compact summary
        // was retained. Keep lengths and pointers observable without copying
        // any row payload bytes into the staging vector.
        assert_slice_summary(&row.cell_storage_buffer_pre_bnc);
        assert_slice_summary(&row.cell_offsets_pre_bnc);
        if let Some(payload) = row.cell_storage_buffer.as_ref() {
            assert_slice_summary(payload);
        }
        if let Some(payload) = row.cell_offsets.as_ref() {
            assert_slice_summary(payload);
        }
    }
}

fn assert_slice_summary(summary: &ObservedSlice) {
    assert!(summary.pointer != 0 || summary.length == 0);
}

// Offset validation is exercised independently of protobuf admission, so
// arbitrary mutations reach the borrowed row boundary even when no Tile can
// be decoded. Traversal and retained scratch remain bounded by input size.
fn exercise_spans(
    offsets: &[u8],
    storage_length: usize,
    wide: bool,
    expected: usize,
    columns: usize,
) {
    let Ok((spans, report)) =
        CellSpans::parse(offsets, storage_length, wide, expected, columns, options())
    else {
        return;
    };
    assert_eq!(spans.len(), expected);
    assert_eq!(report.work_bytes(), offsets.len() * 2);
    let storage = [0u8; 96];
    let storage = &storage[..storage_length];
    let mut previous_end = None;
    let mut previous_column = None;
    let mut iterator = spans.iter();
    let mut count = 0;
    while let Some(span) = iterator.next() {
        assert!(span.column() < columns);
        assert!(span.start() < span.end());
        assert!(span.end() <= storage_length);
        if let Some(end) = previous_end {
            assert_eq!(span.start(), end);
        }
        if let Some(column) = previous_column {
            assert!(span.column() > column);
        }
        let bytes = span.bytes(storage).expect("validated storage range");
        assert_eq!(bytes.as_ptr(), storage[span.range()].as_ptr());
        previous_end = Some(span.end());
        previous_column = Some(span.column());
        count += 1;
        assert_eq!(iterator.len(), expected - count);
    }
    assert_eq!(count, expected);
    assert!(iterator.next().is_none());
    if count != 0 {
        assert_eq!(previous_end, Some(storage_length));
    }
    assert!(spans.get(columns).is_none());
}
