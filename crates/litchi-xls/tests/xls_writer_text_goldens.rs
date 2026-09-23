#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "golden fixtures favor explicit inputs and panic-driven assertions"
)]

//! Byte-exact goldens for the fresh XLS writer's shared-string path.
//!
//! The fixtures of `fresh_xls_writer_shared_strings_match_the_pre_0753_goldens`
//! were written by the writer at `6d989cad63`, before change 0753 staged the
//! shared-string table by borrowing and hashing each string once, and their
//! SHA-256 recorded here. Each of their worksheets holds at most one distinct
//! string: until change 0757 the SST listed strings in the order a worksheet's
//! cell map iterated them, which for several distinct strings depended on that
//! map's per-process hash seed, so only these shapes were reproducible across
//! processes. Change 0757 replaced the two strings of `one_string_per_sheet`
//! that were longer than an SST entry, which the writer now refuses, with
//! strings exactly at the limit, and pinned that fixture's digest from the
//! writer at `9ff78bbf1c` (0753's head), which writes it identically.
//!
//! Since change 0757 the SST lists each string at its first occurrence in
//! worksheet, row and column order, whatever the cell maps' seeds, so
//! `multi_string_worksheets_are_pinned_across_processes` pins worksheets with
//! many distinct strings too. Those digests were recorded from 0757's writer;
//! no earlier writer produced them reproducibly.

use litchi_core::validation::EvidenceDigest;
use litchi_xls::writer::Writer;
use std::io::Cursor;

fn repeat_to(seed: &str, bytes: usize) -> String {
    let mut text = String::with_capacity(bytes + seed.len());
    while text.len() < bytes {
        text.push_str(seed);
    }
    text
}

fn written(mut writer: Writer) -> Vec<u8> {
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}

fn one_string_per_sheet() -> Vec<u8> {
    let values = [
        String::new(),
        "ascii".to_string(),
        "Latin-1 àéîõüçñ".to_string(),
        "CJK 漢字かなカナ".to_string(),
        "astral 😀𝄞🦀".to_string(),
        "\u{80}\u{7ff}\u{800}\u{ffff}\u{10000}\u{10ffff}".to_string(),
        repeat_to("spans CONTINUE records ", 20_000),
        repeat_to("混合 mixed ✓ 😀 ", 30_000),
        // Exactly 0xFFFF UTF-16 code units, the most an SST entry holds, in
        // each encoding; the last ends with a surrogate pair. One unit more
        // is refused (see `xls_writer_string_limits.rs`).
        "b".repeat(0xFFFF),
        "é".repeat(0xFFFF),
        format!("{}😀", "a".repeat(0xFFFD)),
    ];
    let mut writer = Writer::new();
    for (index, value) in values.iter().enumerate() {
        let sheet = writer.add_worksheet(&format!("Sheet{index:02}")).unwrap();
        writer.write_string(sheet, 0, 0, value).unwrap();
        writer
            .write_number(sheet, 1, 0, f64::from(u32::try_from(index).unwrap()))
            .unwrap();
    }
    written(writer)
}

fn repeated_string() -> Vec<u8> {
    let mut writer = Writer::new();
    let sheet = writer.add_worksheet("Repeated").unwrap();
    for row in 0..200u32 {
        writer.write_string(sheet, row, 0, "same label ✓").unwrap();
        writer.write_number(sheet, row, 1, f64::from(row)).unwrap();
    }
    let other = writer.add_worksheet("Other").unwrap();
    writer.write_string(other, 3, 3, "same label ✓").unwrap();
    written(writer)
}

fn harness_numeric(sheets: usize, rows: usize, columns: usize) -> Vec<u8> {
    let mut writer = Writer::new();
    for sheet in 0..sheets {
        let worksheet = writer.add_worksheet(&format!("Bench{sheet:02}")).unwrap();
        for row in 0..rows {
            for column in 0..columns {
                let value = (sheet * rows * columns + row * columns + column) as f64;
                writer
                    .write_number(
                        worksheet,
                        u32::try_from(row).unwrap(),
                        u16::try_from(column).unwrap(),
                        value,
                    )
                    .unwrap();
            }
        }
    }
    written(writer)
}

fn harness_tiny() -> Vec<u8> {
    harness_numeric(1, 4, 4)
}

fn harness_large() -> Vec<u8> {
    harness_numeric(4, 128, 16)
}

fn harness_payload_heavy() -> Vec<u8> {
    let mut writer = Writer::new();
    for sheet in 0..128usize {
        let worksheet = writer.add_worksheet(&format!("Payload{sheet:03}")).unwrap();
        let mut text = format!(
            "litchi-perf-baseline-xls-v1-{sheet:03}-{:05}-{:03} deterministic payload",
            0, 0
        );
        while text.len() < 32_700 {
            text.push_str("litchi-perf-baseline-payload-heavy-v1 ");
        }
        text.truncate(32_700);
        writer.write_string(worksheet, 0, 0, &text).unwrap();
    }
    written(writer)
}

fn empty_worksheet() -> Vec<u8> {
    let mut writer = Writer::new();
    writer.add_worksheet("Empty").unwrap();
    written(writer)
}

#[test]
fn fresh_xls_writer_shared_strings_match_the_pre_0753_goldens() {
    let fixtures: [(&str, fn() -> Vec<u8>, &str); 6] = [
        (
            "one_string_per_sheet",
            one_string_per_sheet,
            "420df55445de70ab66699a18d1de707a2f35b132319c7592794ce577fec3eab5",
        ),
        (
            "repeated_string",
            repeated_string,
            "b82aba04cd266fbed6e78222edf4e23d323b5b425585ca74c71484e512b855a0",
        ),
        (
            "harness_tiny",
            harness_tiny,
            "cdc133bd87aaa60a91ea5e94df6ff8da0eb6bb0f2432afa4bfdb13cf70c0298b",
        ),
        (
            "harness_large",
            harness_large,
            "228c6585a4d26141aebfaf7b08844a2ee445b269d406006a1fdb0484619120fb",
        ),
        (
            "harness_payload_heavy",
            harness_payload_heavy,
            "f0bec5d6a82d5d739a0993704751fb4da906183a4105b159d10e9b73969491fc",
        ),
        (
            "empty_worksheet",
            empty_worksheet,
            "a9a4887f18865177780657c791df2b6f1d1cb8ca33d4bdf3bbee2435a5aeaefc",
        ),
    ];
    let mut mismatches = Vec::new();
    for (name, build, expected) in fixtures {
        let first = build();
        assert_eq!(first, build(), "{name} is not deterministic");
        let actual = EvidenceDigest::of(&first).to_string();
        if actual != expected {
            mismatches.push(format!("{name}: {actual} ({} bytes)", first.len()));
        }
    }
    assert!(
        mismatches.is_empty(),
        "digest mismatches:\n{}",
        mismatches.join("\n")
    );
}

/// Four worksheets of 900 string cells each, with repeated, empty, non-ASCII,
/// supplementary-plane and CONTINUE-spanning strings, and a numeric worksheet
/// between them; also returns every string cell's expected value.
fn many_strings() -> (Writer, Vec<(usize, u32, u16, String)>) {
    let mut writer = Writer::new();
    let mut expected = Vec::new();
    for sheet in 0..4usize {
        let worksheet = writer.add_worksheet(&format!("Strings{sheet}")).unwrap();
        for row in 0..300u32 {
            for column in 0..3u16 {
                let value = match (row + u32::from(column)) % 7 {
                    0 => "repeated".to_string(),
                    1 => String::new(),
                    2 => format!("distinct {sheet} {row} {column}"),
                    3 => format!("ünïcode {row}"),
                    4 => format!("😀 {}", row % 13),
                    5 => "long ".repeat(200 + usize::try_from(row).unwrap()),
                    _ => format!("{sheet}-{row}"),
                };
                writer.write_string(worksheet, row, column, &value).unwrap();
                expected.push((sheet, row, column, value));
            }
            writer
                .write_number(worksheet, row, 3, f64::from(row))
                .unwrap();
        }
        if sheet == 1 {
            let numbers = writer.add_worksheet("Numbers").unwrap();
            writer.write_number(numbers, 0, 0, 1.5).unwrap();
        }
    }
    (writer, expected)
}

/// Strings written in reverse row and column order, repeated across and within
/// worksheets, so neither insertion order nor a cell map's order is the
/// row-major order the SST follows.
fn scrambled_strings() -> Writer {
    let labels = [
        "zeta",
        "alpha",
        "",
        "Latin-1 àéîõüçñ",
        "CJK 漢字かなカナ",
        "astral 😀𝄞🦀",
        "alpha",
    ];
    let mut writer = Writer::new();
    for sheet in 0..3usize {
        let worksheet = writer.add_worksheet(&format!("Scrambled{sheet}")).unwrap();
        for row in (0..40u32).rev() {
            for column in (0..5u16).rev() {
                let index = (usize::try_from(row).unwrap() * 5 + usize::from(column) + sheet)
                    % labels.len();
                if (row + u32::from(column)) % 4 == 3 {
                    writer
                        .write_number(worksheet, row, column, f64::from(row))
                        .unwrap();
                } else if column == 4 && row % 10 == 0 {
                    writer
                        .write_string(
                            worksheet,
                            row,
                            column,
                            &repeat_to(&format!("{sheet}:{row} spans CONTINUE ✓ "), 9_000),
                        )
                        .unwrap();
                } else {
                    writer
                        .write_string(worksheet, row, column, labels[index])
                        .unwrap();
                }
            }
        }
    }
    writer
}

fn many_strings_workbook() -> Vec<u8> {
    written(many_strings().0)
}

fn scrambled_strings_workbook() -> Vec<u8> {
    written(scrambled_strings())
}

/// Worksheets with many distinct strings, pinned since change 0757 made the
/// SST order independent of the cell maps' hash seeds. Each test process seeds
/// its maps afresh, and each fixture is also built four times in-process, each
/// time with new maps, so a seed-dependent order would fail here.
#[test]
fn multi_string_worksheets_are_pinned_across_processes() {
    let fixtures: [(&str, fn() -> Vec<u8>, &str); 2] = [
        (
            "many_strings",
            many_strings_workbook,
            "c66770be6fb44ebaf6b5b4dab23c67acf8dbc90eb19851611d520442d46384f9",
        ),
        (
            "scrambled_strings",
            scrambled_strings_workbook,
            "ecb5af1fe8c2857a30fd0bdec01c28dbe681d5a0cd9180688319599e188bcd1a",
        ),
    ];
    let mut mismatches = Vec::new();
    for (name, build, expected) in fixtures {
        let first = build();
        for _ in 0..3 {
            assert_eq!(first, build(), "{name} is not deterministic");
        }
        let actual = EvidenceDigest::of(&first).to_string();
        if actual != expected {
            mismatches.push(format!("{name}: {actual} ({} bytes)", first.len()));
        }
    }
    assert!(
        mismatches.is_empty(),
        "digest mismatches:\n{}",
        mismatches.join("\n")
    );
}

/// The SST lists each string once, at its first occurrence in worksheet, row
/// and column order.
#[test]
fn the_sst_lists_strings_in_first_occurrence_order() {
    use litchi_core::sheet::{Cell as _, CellValue};

    let workbook = litchi_xls::Workbook::new(Cursor::new(scrambled_strings_workbook())).unwrap();
    let mut expected: Vec<String> = Vec::new();
    for sheet in 0..3usize {
        let worksheet = workbook.xls_worksheet(sheet).unwrap();
        for row in 0..40u32 {
            for column in 0..5u32 {
                if let Some(cell) = worksheet.get_cell(row, column)
                    && let CellValue::String(value) = cell.value()
                    && !expected.contains(value)
                {
                    expected.push(value.clone());
                }
            }
        }
    }
    let table = workbook.xls_worksheet(0).unwrap().shared_strings().unwrap();
    assert_eq!(table, expected.as_slice());
}

/// Every cell is read back as well as pinned: a wrong `LabelSst` index would
/// name another string.
#[test]
fn many_distinct_and_repeated_strings_read_back_cell_by_cell() {
    use litchi_core::sheet::{Cell as _, CellValue};

    let (writer, expected) = many_strings();
    let workbook = litchi_xls::Workbook::new(Cursor::new(written(writer))).unwrap();
    for (sheet, row, column, value) in expected {
        let sheet_index = if sheet >= 2 { sheet + 1 } else { sheet };
        let worksheet = workbook.xls_worksheet(sheet_index).unwrap();
        let cell = worksheet.get_cell(row, u32::from(column)).unwrap();
        assert_eq!(
            cell.value(),
            &CellValue::String(value),
            "{sheet} {row} {column}"
        );
    }
}

/// The shared-string table is staged per write: writing the same writer twice
/// gives the same bytes, and strings added between writes are staged too.
#[test]
fn a_writer_written_twice_and_extended_restages_its_strings() {
    use litchi_core::sheet::{Cell as _, CellValue};

    let mut writer = Writer::new();
    let sheet = writer.add_worksheet("Reused").unwrap();
    for row in 0..50u32 {
        writer
            .write_string(sheet, row, 0, &format!("value {} ✓", row % 17))
            .unwrap();
    }
    let mut first = Cursor::new(Vec::new());
    writer.write_to(&mut first).unwrap();
    let mut second = Cursor::new(Vec::new());
    writer.write_to(&mut second).unwrap();
    assert_eq!(first.get_ref(), second.get_ref());

    writer.write_string(sheet, 0, 1, "added later 😀").unwrap();
    writer.write_string(sheet, 1, 1, "value 3 ✓").unwrap();
    let mut third = Cursor::new(Vec::new());
    writer.write_to(&mut third).unwrap();
    let workbook = litchi_xls::Workbook::new(Cursor::new(third.into_inner())).unwrap();
    let worksheet = workbook.xls_worksheet(0).unwrap();
    for row in 0..50u32 {
        assert_eq!(
            worksheet.get_cell(row, 0).unwrap().value(),
            &CellValue::String(format!("value {} ✓", row % 17))
        );
    }
    assert_eq!(
        worksheet.get_cell(0, 1).unwrap().value(),
        &CellValue::String("added later 😀".to_string())
    );
    assert_eq!(
        worksheet.get_cell(1, 1).unwrap().value(),
        &CellValue::String("value 3 ✓".to_string())
    );
}
