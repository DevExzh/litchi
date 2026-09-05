//! Public regression coverage for BIFF8 Formula `RgbExtra` preservation.
//!
//! The native fixture is intentionally used here instead of reproducing its
//! formulas in a test builder.  A small synthetic package supplies the
//! mutation cases so their outcomes are attributable to the ancillary tail
//! rather than unrelated native workbook resources.

use litchi_cfb::{OleFile, OleWriter};
use litchi_core::OwnedSource;
use litchi_xls::cell_values::{FormulaCache, Reference, Selector, Snapshot, Value};
use litchi_xls::{Error, SourceBackedWorkbook, Workbook, Writer};
use std::io::Cursor;
use std::sync::Arc;

const NATIVE_FIXTURE: &str = "../../test-data/ole/xls/FormulaEvalTestData.xls";
const FORMULA_COUNT: usize = 1_416;

#[derive(Clone, Copy)]
struct CensusFormula {
    stream_offset: usize,
    row: u16,
    column: u16,
    tokens: &'static str,
    extra: &'static str,
}

const CENSUS: [CensusFormula; 5] = [
    CensusFormula {
        stream_offset: 43_084,
        row: 46,
        column: 3,
        tokens: "46101a05131300250800080006c00ac02506000b0008c008c00f",
        extra: "01000800080008000800",
    },
    CensusFormula {
        stream_offset: 43_146,
        row: 46,
        column: 4,
        tokens: "26701a05131300250000ffff07400740250700070007c008c00f19100000",
        extra: "01000700070007000700",
    },
    CensusFormula {
        stream_offset: 43_212,
        row: 46,
        column: 5,
        tokens: "46501c0513190024070003c024060004c0151124080004c01524070005c0110f",
        extra: "01000700070004000400",
    },
    CensusFormula {
        stream_offset: 47_984,
        row: 74,
        column: 3,
        tokens: "46701c05130c0024470001c015244d0001c011",
        extra: "010047004d0001000100",
    },
    CensusFormula {
        stream_offset: 48_055,
        row: 74,
        column: 4,
        tokens: "26901c05130c0024060009c024070008c0151119100000",
        extra: "01000600070008000900",
    },
];

fn native_bytes() -> Vec<u8> {
    std::fs::read(std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join(NATIVE_FIXTURE))
        .expect("read native FormulaEvalTestData.xls fixture")
}

fn workbook_stream(bytes: &[u8]) -> Vec<u8> {
    OleFile::open(Cursor::new(bytes))
        .expect("open CFB fixture")
        .open_stream(&["Workbook"])
        .expect("open Workbook stream")
}

fn hex_bytes(value: &str) -> Vec<u8> {
    assert!(value.len().is_multiple_of(2));
    value
        .as_bytes()
        .as_chunks::<2>()
        .0
        .iter()
        .map(|pair| {
            let text = std::str::from_utf8(pair).expect("hex is ASCII");
            u8::from_str_radix(text, 16).expect("census hex byte")
        })
        .collect()
}

fn record_frames(stream: &[u8]) -> Vec<(usize, u16, usize)> {
    let mut frames = Vec::new();
    let mut offset = 0;
    while offset < stream.len() {
        assert!(offset + 4 <= stream.len(), "truncated BIFF record header");
        let kind = u16::from_le_bytes([stream[offset], stream[offset + 1]]);
        let length = usize::from(u16::from_le_bytes([stream[offset + 2], stream[offset + 3]]));
        let end = offset
            .checked_add(4 + length)
            .expect("BIFF record range overflow");
        assert!(end <= stream.len(), "truncated BIFF record payload");
        frames.push((offset, kind, end));
        offset = end;
    }
    frames
}

fn formula_frames(stream: &[u8]) -> Vec<Vec<u8>> {
    record_frames(stream)
        .into_iter()
        .filter(|(_, kind, _)| *kind == 0x0006)
        .map(|(start, _, end)| stream[start..end].to_vec())
        .collect()
}

fn formula_tail_frames(stream: &[u8]) -> Vec<(usize, Vec<u8>)> {
    record_frames(stream)
        .into_iter()
        .filter_map(|(start, kind, end)| {
            if kind != 0x0006 {
                return None;
            }
            let payload = &stream[start + 4..end];
            assert!(payload.len() >= 22, "Formula payload is truncated");
            let cce = usize::from(u16::from_le_bytes([payload[20], payload[21]]));
            let token_end = 22usize.checked_add(cce).expect("Formula cce overflow");
            assert!(token_end <= payload.len(), "Formula cce exceeds payload");
            (token_end < payload.len()).then(|| (start, payload[token_end..].to_vec()))
        })
        .collect()
}

fn formula_tail_bytes(stream: &[u8]) -> Vec<Vec<u8>> {
    formula_tail_frames(stream)
        .into_iter()
        .map(|(_, bytes)| bytes)
        .collect()
}

fn assert_native_census(stream: &[u8]) {
    let frames = formula_frames(stream);
    assert_eq!(frames.len(), FORMULA_COUNT, "native formula census changed");
    assert_eq!(formula_tail_frames(stream).len(), CENSUS.len());

    for expected in CENSUS {
        let start = expected.stream_offset;
        assert_eq!(
            u16::from_le_bytes([stream[start], stream[start + 1]]),
            0x0006
        );
        let payload_len = usize::from(u16::from_le_bytes([stream[start + 2], stream[start + 3]]));
        let payload = &stream[start + 4..start + 4 + payload_len];
        assert_eq!(
            u16::from_le_bytes([payload[0], payload[1]]),
            expected.row,
            "Formula row at stream offset {start}"
        );
        assert_eq!(
            u16::from_le_bytes([payload[2], payload[3]]),
            expected.column,
            "Formula column at stream offset {start}"
        );
        let tokens = hex_bytes(expected.tokens);
        let extra = hex_bytes(expected.extra);
        let cce = usize::from(u16::from_le_bytes([payload[20], payload[21]]));
        assert_eq!(cce, tokens.len(), "cce must end exactly at rgce");
        assert_eq!(&payload[22..22 + cce], tokens.as_slice());
        assert_eq!(&payload[22 + cce..], extra.as_slice());
    }
}

fn assert_eager_formula_metadata(bytes: &[u8]) {
    let workbook = Workbook::new(Cursor::new(bytes.to_vec())).expect("open native XLS");
    let worksheet = workbook
        .xls_worksheet(0)
        .expect("EverythingTests worksheet");
    for expected in CENSUS {
        let cell = worksheet
            .get_cell(u32::from(expected.row), u32::from(expected.column))
            .expect("census Formula cell");
        let tokens = hex_bytes(expected.tokens);
        let extra = hex_bytes(expected.extra);
        assert_eq!(cell.formula_bytes(), Some(tokens.as_slice()));
        assert_eq!(
            cell.formula_metadata()
                .expect("Formula metadata")
                .ancillary_bytes(),
            Some(extra.as_slice())
        );
    }
}

fn assert_single_formula_metadata(bytes: &[u8], row: u16, column: u16) {
    let workbook = Workbook::new(Cursor::new(bytes.to_vec())).expect("open synthetic XLS");
    let worksheet = workbook
        .xls_worksheet(0)
        .expect("synthetic Formula worksheet");
    let cell = worksheet
        .get_cell(u32::from(row), u32::from(column))
        .expect("synthetic Formula cell");
    let tokens = hex_bytes(CENSUS[0].tokens);
    let extra = hex_bytes(CENSUS[0].extra);
    assert_eq!(cell.formula_bytes(), Some(tokens.as_slice()));
    assert_eq!(
        cell.formula_metadata()
            .expect("synthetic Formula metadata")
            .ancillary_bytes(),
        Some(extra.as_slice())
    );
}

#[test]
fn native_formula_census_retains_all_five_rgb_extra_tails() {
    let bytes = native_bytes();
    let stream = workbook_stream(&bytes);
    assert_native_census(&stream);
    assert_eager_formula_metadata(&bytes);
}

#[test]
fn source_backed_formula_access_reaches_the_native_ancillary_parser() {
    let bytes = native_bytes();
    let source = SourceBackedWorkbook::from_read_at(Arc::new(OwnedSource::new(bytes)))
        .expect("open source-backed native XLS");
    let worksheet = source
        .worksheet(0)
        .expect("source-backed worksheet lookup")
        .expect("EverythingTests worksheet");

    // `cell` scans and parses the complete worksheet.  One lookup therefore
    // exercises all five Formula tails without requiring SourceBackedCell to
    // expose eager-only token metadata.
    let cell = worksheet
        .cell(46, 3)
        .expect("source-backed Formula access")
        .expect("source-backed census cell");
    assert_eq!((cell.row(), cell.column()), (46, 3));
}

#[test]
fn native_noop_is_byte_exact_and_source_backed_numeric_edit_keeps_formula_frames() {
    let native = Snapshot::from_bytes(native_bytes()).expect("open native cell-value snapshot");
    let no_op = native.transaction().commit().expect("exact no-op commit");
    assert_eq!(no_op.snapshot().bytes(), native.bytes());
    assert!(std::ptr::eq(
        no_op.snapshot().bytes().as_ptr(),
        native.bytes().as_ptr()
    ));

    let source = Snapshot::from_bytes(ancillary_formula_package())
        .expect("open isolated source-backed cell-value snapshot");
    let (sheet_position, number) = source
        .worksheets()
        .find_map(|worksheet| {
            worksheet
                .numbers()
                .find(|number| {
                    number.value().is_finite()
                        && (number.value() != 0.0 || number.value().is_sign_positive())
                })
                .map(|number| (worksheet.position(), number))
        })
        .expect("synthetic fixture contains an editable Number record");

    let mut source_backed_noop = source.transaction();
    source_backed_noop
        .set_number(
            Selector::Position(sheet_position),
            number.reference(),
            number.value(),
        )
        .expect("stage source-backed numeric no-op");
    let source_backed_noop = source_backed_noop
        .commit_source_backed()
        .expect("publish source-backed numeric no-op");
    assert!(source_backed_noop.is_noop());
    assert_eq!(source_backed_noop.snapshot().bytes(), source.bytes());
    assert!(std::ptr::eq(
        source_backed_noop.snapshot().bytes().as_ptr(),
        source.bytes().as_ptr()
    ));

    let replacement = if number.value().to_bits() == 12_345.25_f64.to_bits() {
        54_321.5
    } else {
        12_345.25
    };

    let mut edit = source.transaction();
    edit.set_number(
        Selector::Position(sheet_position),
        number.reference(),
        replacement,
    )
    .expect("stage unrelated fixed-width Number edit");
    let commit = edit
        .commit_source_backed()
        .expect("source-backed numeric publication");
    let target = commit.snapshot();
    assert_eq!(commit.diagnostics().changed_cells(), 1);
    assert_eq!(commit.diagnostics().touched_streams(), 1);
    assert_eq!(
        source.workbook_stream().len(),
        target.workbook_stream().len()
    );
    assert_eq!(
        formula_frames(source.workbook_stream()),
        formula_frames(target.workbook_stream())
    );
    assert_eq!(
        formula_tail_frames(source.workbook_stream()),
        formula_tail_frames(target.workbook_stream())
    );
    assert_non_workbook_streams_equal(source.bytes(), target.bytes());
    Workbook::new(Cursor::new(target.bytes().to_vec()))
        .expect("source-backed target reopens through eager reader");
    assert_single_formula_metadata(target.bytes(), 0, 0);
    assert_eq!(
        commit
            .patch()
            .inverse()
            .apply(target)
            .expect("source-backed inverse")
            .bytes(),
        source.bytes()
    );
}

#[test]
fn formula_cache_and_style_edits_preserve_the_untouched_rgb_extra() {
    let source = Snapshot::from_bytes(ancillary_formula_package())
        .expect("open isolated ancillary snapshot");
    let reference = Reference::new(0, 0).expect("synthetic reference");

    let old_cache = source
        .worksheet(Selector::Position(0))
        .expect("resolve worksheet")
        .expect("worksheet exists")
        .cell(reference)
        .expect("unique census cell")
        .expect("census cell exists")
        .value()
        .clone();
    let replacement_cache = match old_cache {
        Value::FormulaCache(FormulaCache::Empty) => FormulaCache::Number(1.0),
        _ => FormulaCache::Empty,
    };
    let mut cache_edit = source.transaction();
    cache_edit
        .set_value(
            Selector::Position(0),
            reference,
            Value::FormulaCache(replacement_cache),
        )
        .expect("stage Formula cache replacement");
    let cache_commit = cache_edit.commit().expect("commit Formula cache edit");
    assert_eq!(
        formula_tail_frames(source.workbook_stream()),
        formula_tail_frames(cache_commit.snapshot().workbook_stream())
    );
    assert_single_formula_metadata(cache_commit.snapshot().bytes(), 0, 0);
    assert_eq!(
        cache_commit
            .patch()
            .inverse()
            .apply(cache_commit.snapshot())
            .expect("invert Formula cache edit")
            .bytes(),
        source.bytes()
    );

    let style = source
        .worksheet(Selector::Position(0))
        .expect("resolve worksheet")
        .expect("worksheet exists")
        .cell(reference)
        .expect("unique census cell")
        .expect("census cell exists")
        .style();
    let mut style_edit = source.transaction();
    let replacement_style = style_edit
        .duplicate_style(style)
        .expect("duplicate existing XF resource");
    style_edit
        .set_style(Selector::Position(0), reference, replacement_style)
        .expect("stage Formula style edit");
    let style_commit = style_edit.commit().expect("commit Formula style edit");
    assert_eq!(
        formula_tail_bytes(source.workbook_stream()),
        formula_tail_bytes(style_commit.snapshot().workbook_stream())
    );
    assert_single_formula_metadata(style_commit.snapshot().bytes(), 0, 0);
}

fn assert_non_workbook_streams_equal(before: &[u8], after: &[u8]) {
    let mut source = OleFile::open(Cursor::new(before)).expect("source CFB");
    let mut target = OleFile::open(Cursor::new(after)).expect("target CFB");
    let source_paths = source.list_streams();
    let target_paths = target.list_streams();
    assert_eq!(source_paths, target_paths);
    for path in source_paths {
        if path.len() == 1 && matches!(path[0].as_str(), "Workbook" | "Book") {
            continue;
        }
        let refs = path.iter().map(String::as_str).collect::<Vec<_>>();
        assert_eq!(
            source.open_stream(&refs).expect("source stream"),
            target.open_stream(&refs).expect("target stream"),
            "untouched CFB stream changed at {path:?}"
        );
    }
}

fn ancillary_formula_package() -> Vec<u8> {
    let mut writer = Writer::new();
    let sheet = writer.add_worksheet("Formula").expect("add worksheet");
    writer
        .write_formula(sheet, 0, 0, "A1")
        .expect("write seed formula");
    writer
        .write_number(sheet, 1, 0, 42.0)
        .expect("write seed number");
    let mut base = Cursor::new(Vec::new());
    writer.write_to(&mut base).expect("write seed package");
    let base = base.into_inner();
    let stream = workbook_stream(&base);
    let native_tokens = hex_bytes(CENSUS[0].tokens);
    let native_extra = hex_bytes(CENSUS[0].extra);
    let mut rebuilt = Vec::with_capacity(stream.len() + native_extra.len());
    let mut patched = false;
    for (start, kind, end) in record_frames(&stream) {
        let mut payload = stream[start + 4..end].to_vec();
        if kind == 0x0006 && !patched {
            payload[0..2].copy_from_slice(&0_u16.to_le_bytes());
            payload[2..4].copy_from_slice(&0_u16.to_le_bytes());
            payload[20..22]
                .copy_from_slice(&u16::try_from(native_tokens.len()).unwrap().to_le_bytes());
            payload.truncate(22);
            payload.extend_from_slice(&native_tokens);
            payload.extend_from_slice(&native_extra);
            patched = true;
        }
        rebuilt.extend_from_slice(&kind.to_le_bytes());
        rebuilt.extend_from_slice(&u16::try_from(payload.len()).unwrap().to_le_bytes());
        rebuilt.extend_from_slice(&payload);
    }
    assert!(patched, "seed package has no Formula record");

    let mut container = OleWriter::new();
    container
        .create_stream(&["Workbook"], &rebuilt)
        .expect("create synthetic Workbook stream");
    container
        .create_stream(
            &["UnrelatedData"],
            b"retain this unrelated CFB stream exactly",
        )
        .expect("create untouched opaque stream");
    let mut output = Cursor::new(Vec::new());
    container
        .write_to(&mut output)
        .expect("write synthetic package");
    output.into_inner()
}

#[test]
fn row_and_column_shift_refuse_rgb_extra_atomically() {
    let bytes = ancillary_formula_package();
    let source = Snapshot::from_bytes(bytes).expect("open isolated ancillary package");
    let before = source.bytes().to_vec();

    let mut rows = source.transaction();
    assert!(matches!(
        rows.insert_rows(Selector::Position(0), 0, 1),
        Err(Error::UnsafeEdit(_))
    ));

    let mut columns = source.transaction();
    assert!(matches!(
        columns.insert_columns(Selector::Position(0), 0, 1),
        Err(Error::UnsafeEdit(_))
    ));
    assert_eq!(source.bytes(), before.as_slice());
}
