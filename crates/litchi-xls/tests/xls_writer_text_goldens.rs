#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "golden fixtures favor explicit inputs and panic-driven assertions"
)]

//! Byte-exact goldens for the fresh XLS writer's shared-string path.
//!
//! Every fixture below was written by the writer at `6d989cad63`, before
//! change 0753 staged the shared-string table by borrowing and hashing each
//! string once, and its SHA-256 recorded here. Each worksheet holds at most one
//! distinct string: the SST lists strings in the order a worksheet's cell map
//! iterates them, which for several distinct strings in one worksheet depends
//! on that map's per-process hash seed, so only these shapes are reproducible
//! across processes. The table's order and indices for many distinct strings
//! are pinned against the previous algorithm in-process by the table's own
//! unit tests.

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
        // More than 0xFFFF UTF-16 code units: the SST keeps the first 0xFFFF.
        repeat_to("😀", 140_000),
        repeat_to("b", 70_000),
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
            "552b720e5656a88df6d8d63ebaf50327a0d89c24320ef8a4847994116f376b85",
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

/// Many distinct and repeated strings per worksheet cannot be pinned by digest
/// (see the module comment), so every cell is read back instead: a wrong
/// `LabelSst` index would name another string.
#[test]
fn many_distinct_and_repeated_strings_read_back_cell_by_cell() {
    use litchi_core::sheet::{Cell as _, CellValue};

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
