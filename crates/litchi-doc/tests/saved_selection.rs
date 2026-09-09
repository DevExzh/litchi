//! Public and deferred-reader coverage for the inert DOC `Selsf` cache.

#![allow(
    clippy::expect_used,
    clippy::uninlined_format_args,
    reason = "bounded binary fixtures fail fast with contextual assertions"
)]

use litchi_doc::parts::fib::FileInformationBlock;
use litchi_doc::tracked_revision::Limits;
use litchi_doc::writer::Writer;
use litchi_doc::{Package, SavedSelection, SelectionGeometry, SelectionStyle};
use litchi_ole_common::object::{Editor as PackageEditor, Targets};
use std::io::Cursor;

fn record(flags: u16) -> Vec<u8> {
    let mut data = vec![0; 36];
    data[0..2].copy_from_slice(&flags.to_le_bytes());
    data[2] = 0x81;
    data[4..8].copy_from_slice(&4i32.to_le_bytes());
    data[8..12].copy_from_slice(&8i32.to_le_bytes());
    data[12..16].copy_from_slice(&[0xA5, 0x5A, 0xC3, 0x3C]);
    data[20..24].copy_from_slice(&4i32.to_le_bytes());
    data[24..26].copy_from_slice(&(SelectionStyle::Character as u16).to_le_bytes());
    data[26..28].copy_from_slice(&[0xD7, 0x7D]);
    data[32..34].copy_from_slice(&(-100i16).to_le_bytes());
    data[34..36].copy_from_slice(&100i16.to_le_bytes());
    data
}

fn set_cps(data: &mut [u8], cp_first: i32, cp_lim: i32, cp_anchor: i32) {
    data[4..8].copy_from_slice(&cp_first.to_le_bytes());
    data[8..12].copy_from_slice(&cp_lim.to_le_bytes());
    data[20..24].copy_from_slice(&cp_anchor.to_le_bytes());
}

fn document_with_selsf(source: &[u8], length: u32) -> Vec<u8> {
    let mut writer = Writer::new();
    writer.add_paragraph("Body").expect("fixture paragraph");
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).expect("fixture DOC");
    let mut package =
        PackageEditor::open(output.into_inner(), Targets::default(), Limits::default())
            .expect("package");

    let word_path = ["WordDocument".to_string()];
    let mut word = package.stream(&word_path).expect("WordDocument").to_vec();
    let fib = FileInformationBlock::parse(&word).expect("FIB");
    let table_name = if fib.which_table_stream() {
        "1Table"
    } else {
        "0Table"
    };
    let table_path = [table_name.to_string()];
    let mut table = package.stream(&table_path).expect("table stream").to_vec();
    let offset = u32::try_from(table.len()).expect("table offset");
    table.extend_from_slice(source);
    package
        .put_stream(&table_path, table)
        .expect("Selsf table stream");

    let pair = 154 + 30 * 8;
    word[pair..pair + 4].copy_from_slice(&offset.to_le_bytes());
    word[pair + 4..pair + 8].copy_from_slice(&length.to_le_bytes());
    package.put_stream(&word_path, word).expect("Selsf pointer");
    package.finish().expect("package finish")
}

#[test]
fn parses_fixed_record_and_preserves_ignored_fields() {
    let mut data = record(1 << 13);
    data[16..20].copy_from_slice(&0x0002_0001u32.to_le_bytes());
    let selection = SavedSelection::parse_bytes(&data).expect("valid Selsf");
    assert_eq!(selection.bytes(), data.as_slice());
    assert_eq!(selection.cp_first(), 4);
    assert_eq!(selection.cp_lim(), 8);
    assert_eq!(selection.cp_anchor(), 4);
    assert_eq!(selection.direction_flags(), 0x81);
    assert!(selection.is_forward());
    assert!(selection.prefix_w2007());
    assert_eq!(
        selection.geometry(),
        SelectionGeometry::Block { first: 1, limit: 2 }
    );

    let mut table = record(1 << 11);
    table[16..20].copy_from_slice(&0x0040_0000u32.to_le_bytes());
    table[32..34].copy_from_slice(&(-31_680i16).to_le_bytes());
    table[34..36].copy_from_slice(&31_680i16.to_le_bytes());
    assert_eq!(
        SavedSelection::parse_bytes(&table)
            .expect("whole-row Selsf")
            .geometry(),
        SelectionGeometry::Table {
            first: 0,
            limit: 64
        }
    );
}

#[test]
fn rejects_invalid_domains_and_cross_field_constraints() {
    let mut reserved = record(1 << 14);
    assert!(SavedSelection::parse_bytes(&reserved).is_err());
    reserved[0] = 0;
    reserved[3] = 2;
    assert!(SavedSelection::parse_bytes(&reserved).is_err());

    let mut shape = record(1 << 8);
    shape[3] = 2;
    let shape_selection = SavedSelection::parse_bytes(&shape).expect("shape Selsf");
    assert!(!shape_selection.is_insertion_end());
    shape[3] = 1;
    assert!(SavedSelection::parse_bytes(&shape).is_ok());

    let mut insertion = record(1 << 15);
    insertion[8..12].copy_from_slice(&9i32.to_le_bytes());
    assert!(SavedSelection::parse_bytes(&insertion).is_err());

    for flags in [
        1 << 2 | 1 << 11,
        1 << 4,
        1 << 8 | 1 << 12,
        1 << 10,
        1 << 10 | 1 << 11,
    ] {
        assert!(SavedSelection::parse_bytes(&record(flags)).is_err());
    }

    let mut invalid_style = record(0);
    invalid_style[24..26].copy_from_slice(&6u16.to_le_bytes());
    assert!(SavedSelection::parse_bytes(&invalid_style).is_err());
}

#[test]
fn fib_range_requires_exact_size_and_deferred_errors_are_cached() {
    let source = record(0);
    let offset = 8u32;
    let mut fib_data = vec![0; 154 + 31 * 8];
    fib_data[0..2].copy_from_slice(&0xA5ECu16.to_le_bytes());
    fib_data[2..4].copy_from_slice(&0x00C1u16.to_le_bytes());
    fib_data[152..154].copy_from_slice(&31u16.to_le_bytes());
    let pointer = 154 + 30 * 8;
    fib_data[pointer..pointer + 4].copy_from_slice(&offset.to_le_bytes());
    fib_data[pointer + 4..pointer + 8].copy_from_slice(&36u32.to_le_bytes());
    let fib = FileInformationBlock::parse(&fib_data).expect("FIB");
    let mut table = vec![0xCC; 80];
    table[offset as usize..offset as usize + 36].copy_from_slice(&source);
    assert_eq!(
        SavedSelection::parse(&fib, &table)
            .expect("selected Selsf")
            .expect("record")
            .bytes(),
        source.as_slice()
    );

    let bytes = document_with_selsf(&source[..35], 35);

    let mut package = Package::from_reader(Cursor::new(bytes)).expect("package open");
    let document = package
        .document()
        .expect("document open despite optional error");
    let first = document
        .saved_selection()
        .expect_err("malformed optional Selsf");
    let second = document
        .saved_selection()
        .expect_err("cached malformed optional Selsf");
    assert_eq!(first.to_string(), second.to_string());
}

#[test]
fn document_accessor_returns_selection_and_enforces_main_text_bound() {
    let mut valid = record(0);
    set_cps(&mut valid, 0, 1, 0);
    let bytes = document_with_selsf(&valid, 36);
    let mut package = Package::from_reader(Cursor::new(bytes)).expect("package open");
    let document = package.document().expect("document open");
    let selection = document
        .saved_selection()
        .expect("valid Selsf")
        .expect("selected record");
    assert_eq!(selection.cp_first(), 0);
    assert_eq!(selection.cp_lim(), 1);
    assert!(selection.cp_lim() <= document.fib().get_main_doc_range().1);

    let mut outside = record(0);
    set_cps(&mut outside, i32::MAX, i32::MAX, i32::MAX);
    let bytes = document_with_selsf(&outside, 36);
    let mut package = Package::from_reader(Cursor::new(bytes)).expect("package open");
    let document = package.document().expect("document open");
    let first = document
        .saved_selection()
        .expect_err("cpFirst beyond ccpText");
    let second = document
        .saved_selection()
        .expect_err("cpFirst error is cached");
    assert!(first.to_string().contains("cpFirst"));
    assert_eq!(first.to_string(), second.to_string());

    let mut outside_lim = record(0);
    set_cps(&mut outside_lim, 0, i32::MAX, 0);
    let bytes = document_with_selsf(&outside_lim, 36);
    let mut package = Package::from_reader(Cursor::new(bytes)).expect("package open");
    let document = package.document().expect("document open");
    let first = document
        .saved_selection()
        .expect_err("cpLim beyond ccpText");
    let second = document
        .saved_selection()
        .expect_err("cpLim error is cached");
    assert!(first.to_string().contains("cpLim"));
    assert_eq!(first.to_string(), second.to_string());
}
