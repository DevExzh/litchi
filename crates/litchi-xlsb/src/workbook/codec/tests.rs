#![allow(
    clippy::cast_possible_truncation,
    reason = "test fixture uses bounded literal casts, panic-on-failure extraction, exact floating sentinels, or explicit negative fallback solely to state its assertion"
)]

//! Focused tests for the private workbook codec owners.

use super::super::super::Workbook;
use crate::package::error::{Error, Result};
use crate::package::shared_strings::SharedString;
use crate::raw::{Encoder, Kind, Records, kind};

fn wide_string(value: &str) -> Vec<u8> {
    let mut data = Vec::new();
    data.extend_from_slice(&(value.encode_utf16().count() as u32).to_le_bytes());
    for unit in value.encode_utf16() {
        data.extend_from_slice(&unit.to_le_bytes());
    }
    data
}

fn bundle_sheet_payload(id: u32, relationship_id: &str, name: &str) -> Vec<u8> {
    let mut data = 0_u32.to_le_bytes().to_vec();
    data.extend_from_slice(&id.to_le_bytes());
    data.extend_from_slice(&wide_string(relationship_id));
    data.extend_from_slice(&wide_string(name));
    data
}

fn record_stream(records: &[(Kind, Vec<u8>)]) -> Result<Vec<u8>> {
    let mut data = Vec::new();
    let mut writer = Encoder::new(&mut data, crate::raw::Limits::DEFAULT);
    for (record_kind, payload) in records {
        writer.write_record(*record_kind, payload)?;
    }
    Ok(data)
}

fn push_independent_record(bytes: &mut Vec<u8>, record_kind: Kind, payload: &[u8]) {
    let mut value = u32::from(record_kind.get());
    loop {
        let mut byte = (value & 0x7F) as u8;
        value >>= 7;
        if value != 0 {
            byte |= 0x80;
        }
        bytes.push(byte);
        if value == 0 {
            break;
        }
    }

    let mut value = payload.len() as u32;
    loop {
        let mut byte = (value & 0x7F) as u8;
        value >>= 7;
        if value != 0 {
            byte |= 0x80;
        }
        bytes.push(byte);
        if value == 0 {
            break;
        }
    }
    bytes.extend_from_slice(payload);
}

fn nullable_wide_string(value: Option<&str>) -> Vec<u8> {
    value.map_or_else(
        || u32::MAX.to_le_bytes().to_vec(),
        |value| wide_string(value),
    )
}

fn independent_table_header() -> Vec<u8> {
    let mut payload = Vec::new();
    for value in [0u32, 1, 0, 0] {
        payload.extend_from_slice(&value.to_le_bytes());
    }
    payload.extend_from_slice(&2u32.to_le_bytes()); // LTXML
    payload.extend_from_slice(&7u32.to_le_bytes()); // idList
    payload.extend_from_slice(&1u32.to_le_bytes()); // crwHeader
    payload.extend_from_slice(&0u32.to_le_bytes()); // crwTotals
    payload.extend_from_slice(&0u32.to_le_bytes()); // flags
    for _ in 0..6 {
        payload.extend_from_slice(&u32::MAX.to_le_bytes());
    }
    payload.extend_from_slice(&0u32.to_le_bytes()); // dwConnID
    payload.extend_from_slice(&nullable_wide_string(Some("MappedTable")));
    payload.extend_from_slice(&nullable_wide_string(Some("MappedTable")));
    for _ in 0..4 {
        payload.extend_from_slice(&nullable_wide_string(None));
    }
    payload
}

fn independent_table_column() -> Vec<u8> {
    let mut payload = Vec::new();
    payload.extend_from_slice(&1u32.to_le_bytes()); // idField
    payload.extend_from_slice(&0u32.to_le_bytes()); // ilta
    for _ in 0..3 {
        payload.extend_from_slice(&u32::MAX.to_le_bytes());
    }
    payload.extend_from_slice(&0u32.to_le_bytes()); // idqsif
    payload.extend_from_slice(&nullable_wide_string(None));
    payload.extend_from_slice(&nullable_wide_string(Some("MappedColumn")));
    for _ in 0..4 {
        payload.extend_from_slice(&nullable_wide_string(None));
    }
    payload
}

fn independent_xml_properties() -> Vec<u8> {
    let mut payload = Vec::new();
    payload.extend_from_slice(&1u32.to_le_bytes()); // map ID
    payload.extend_from_slice(&2u32.to_le_bytes()); // can map a single cell
    payload.extend_from_slice(&1u32.to_le_bytes()); // XML data type
    payload.extend_from_slice(&wide_string("/root/item"));
    payload
}

#[test]
fn shared_string_owner_validates_declared_count() {
    let mut item = vec![0];
    item.extend_from_slice(&wide_string("value"));
    let data = record_stream(&[
        (
            kind::BEGIN_SST,
            [1u32.to_le_bytes(), 1u32.to_le_bytes()].concat(),
        ),
        (kind::SST_ITEM, item),
        (kind::END_SST, Vec::new()),
    ])
    .unwrap();
    let mut records = Records::new(&data, crate::raw::Limits::DEFAULT);
    let mut strings: Vec<SharedString> = Vec::new();
    Workbook::read_shared_strings(
        &mut records,
        &mut strings,
        &mut Default::default(),
        crate::raw::Limits::DEFAULT,
    )
    .unwrap();
    assert_eq!(strings.len(), 1);
    assert_eq!(strings[0].text, "value");
}

#[test]
fn pivot_cache_record_owner_preserves_relationship_identity() {
    let mut cache = 12u32.to_le_bytes().to_vec();
    cache.extend_from_slice(&wide_string("rIdCache"));
    let data = record_stream(&[
        (kind::BEGIN_PIVOT_CACHE_IDS, Vec::new()),
        (kind::BEGIN_PIVOT_CACHE_ID, cache),
        (kind::END_PIVOT_CACHE_ID, Vec::new()),
        (kind::END_PIVOT_CACHE_IDS, Vec::new()),
    ])
    .unwrap();
    assert_eq!(
        Workbook::parse_pivot_cache_ids(&data, crate::raw::Limits::DEFAULT).unwrap(),
        vec![(12, "rIdCache".to_string())]
    );
}

/// One row per record sequence around the PivotCache ID collection: `[`/`]`
/// begin and end the collection, `(`/`)` begin and end one cache (each with
/// its own ID), `X` is a record outside the family.
#[test]
fn pivot_cache_id_collection_accepts_and_rejects_the_same_sequences() {
    let cases: &[(&str, &str)] = &[
        ("", "caches 0"),
        ("X", "caches 0"),
        ("[]", "caches 0"),
        ("X[]X", "caches 0"),
        ("[()]", "caches 1"),
        ("[()()]", "caches 2"),
        ("[X(X)X]", "caches 1"),
        ("[()]X", "caches 1"),
        (
            "[",
            "error: Invalid formula: unterminated PivotCache ID collection",
        ),
        (
            "[(",
            "error: Invalid formula: unterminated PivotCache ID collection",
        ),
        (
            "[()",
            "error: Invalid formula: unterminated PivotCache ID collection",
        ),
        (
            "[][]",
            "error: Invalid formula: duplicate BrtBeginPivotCacheIDs collection",
        ),
        (
            "[]X[]",
            "error: Invalid formula: duplicate BrtBeginPivotCacheIDs collection",
        ),
        (
            "[[",
            "error: Invalid formula: duplicate BrtBeginPivotCacheIDs collection",
        ),
        (
            "[([",
            "error: Invalid formula: duplicate BrtBeginPivotCacheIDs collection",
        ),
        (
            "]",
            "error: Invalid formula: unbalanced BrtEndPivotCacheIDs",
        ),
        (
            "[]]",
            "error: Invalid formula: unbalanced BrtEndPivotCacheIDs",
        ),
        (
            "[(]",
            "error: Invalid formula: unbalanced BrtEndPivotCacheIDs",
        ),
        (
            "(",
            "error: Invalid formula: malformed BrtBeginPivotCacheID nesting or payload",
        ),
        (
            "[](",
            "error: Invalid formula: malformed BrtBeginPivotCacheID nesting or payload",
        ),
        (
            "[((",
            "error: Invalid formula: malformed BrtBeginPivotCacheID nesting or payload",
        ),
        (")", "error: Invalid formula: unbalanced BrtEndPivotCacheID"),
        (
            "[)",
            "error: Invalid formula: unbalanced BrtEndPivotCacheID",
        ),
        (
            "[())",
            "error: Invalid formula: unbalanced BrtEndPivotCacheID",
        ),
        (
            "[])",
            "error: Invalid formula: unbalanced BrtEndPivotCacheID",
        ),
    ];
    let mismatches: Vec<String> = cases
        .iter()
        .filter_map(|(sequence, expected)| {
            let mut next_id = 0_u32;
            let records: Vec<(Kind, Vec<u8>)> = sequence
                .chars()
                .map(|token| match token {
                    '[' => (kind::BEGIN_PIVOT_CACHE_IDS, Vec::new()),
                    ']' => (kind::END_PIVOT_CACHE_IDS, Vec::new()),
                    '(' => {
                        next_id += 1;
                        let mut cache = next_id.to_le_bytes().to_vec();
                        cache.extend_from_slice(&wide_string("rId"));
                        (kind::BEGIN_PIVOT_CACHE_ID, cache)
                    },
                    ')' => (kind::END_PIVOT_CACHE_ID, Vec::new()),
                    _ => (Kind::new(4095).unwrap(), vec![1, 2]),
                })
                .collect();
            let data = record_stream(&records).unwrap();
            let actual = match Workbook::parse_pivot_cache_ids(&data, crate::raw::Limits::DEFAULT) {
                Ok(caches) => format!("caches {}", caches.len()),
                Err(error) => format!("error: {error}"),
            };
            (actual != *expected).then(|| format!("(\"{sequence}\", \"{actual}\"),"))
        })
        .collect();
    assert!(mismatches.is_empty(), "\n{}\n", mismatches.join("\n"));
}

#[test]
fn ordinary_table_owner_ignores_opaque_balanced_xml_map_extensions() {
    let mut data = Vec::new();
    push_independent_record(&mut data, kind::BEGIN_LIST, &independent_table_header());
    push_independent_record(&mut data, kind::BEGIN_LIST_COLS, &1u32.to_le_bytes());
    push_independent_record(&mut data, kind::BEGIN_LIST_COL, &independent_table_column());
    push_independent_record(
        &mut data,
        kind::BEGIN_LIST_XML_CPR,
        &independent_xml_properties(),
    );
    push_independent_record(&mut data, Kind::new(0x0FFE).unwrap(), b"opaque");
    push_independent_record(&mut data, kind::FRT_BEGIN, &[0; 4]);
    push_independent_record(&mut data, kind::AC_BEGIN, &[]);
    // A known table record is opaque while protected by extension wrappers;
    // it must not close the host parser's real column.
    push_independent_record(&mut data, kind::END_LIST_COL, &[]);
    push_independent_record(&mut data, kind::AC_END, &[]);
    push_independent_record(&mut data, kind::FRT_END, &[]);
    push_independent_record(&mut data, kind::END_LIST_XML_CPR, &[]);
    push_independent_record(&mut data, kind::END_LIST_COL, &[]);
    push_independent_record(&mut data, kind::END_LIST_COLS, &[]);
    push_independent_record(&mut data, kind::END_LIST, &[]);

    let table = Workbook::parse_table_definition(
        &data,
        3,
        &mut crate::leniency::Tolerance::refusing(),
        crate::raw::Limits::DEFAULT,
    )
    .unwrap();
    assert_eq!(table.table_id(), 7);
    assert_eq!(table.sheet_index(), 3);
    assert_eq!(table.display_name(), "MappedTable");
    assert_eq!(table.columns().len(), 1);
}

#[test]
fn ordinary_table_owner_rejects_mismatched_xml_map_extensions() {
    let mut data = Vec::new();
    push_independent_record(&mut data, kind::BEGIN_LIST, &independent_table_header());
    push_independent_record(&mut data, kind::BEGIN_LIST_COLS, &1u32.to_le_bytes());
    push_independent_record(&mut data, kind::BEGIN_LIST_COL, &independent_table_column());
    push_independent_record(
        &mut data,
        kind::BEGIN_LIST_XML_CPR,
        &independent_xml_properties(),
    );
    push_independent_record(&mut data, kind::FRT_BEGIN, &[0; 4]);
    push_independent_record(&mut data, kind::AC_END, &[]);

    assert!(
        Workbook::parse_table_definition(
            &data,
            0,
            &mut crate::leniency::Tolerance::refusing(),
            crate::raw::Limits::DEFAULT,
        )
        .is_err()
    );
}

#[test]
fn workbook_sheet_name_index_expands_unicode_lowercase_without_normalizing() {
    let duplicate_names = ["\u{0130}", "i\u{0307}"];
    let duplicate_records = duplicate_names
        .iter()
        .enumerate()
        .map(|(index, name)| {
            (
                kind::BUNDLE_SH,
                bundle_sheet_payload(
                    u32::try_from(index + 1).unwrap(),
                    &format!("rId{}", index + 1),
                    name,
                ),
            )
        })
        .collect::<Vec<_>>();
    let duplicate_data = record_stream(&duplicate_records).unwrap();
    let mut duplicate_iter = Records::new(&duplicate_data, crate::raw::Limits::DEFAULT);
    assert!(matches!(
        Workbook::read_workbook(
            &mut duplicate_iter,
            &mut Default::default(),
            crate::raw::Limits::DEFAULT,
        ),
        Err(Error::Unrecognized { typ, .. }) if typ == "BrtBundleSh strName"
    ));

    let canonically_equivalent_names = ["é", "e\u{301}"];
    let distinct_records = canonically_equivalent_names
        .iter()
        .enumerate()
        .map(|(index, name)| {
            (
                kind::BUNDLE_SH,
                bundle_sheet_payload(
                    u32::try_from(index + 1).unwrap(),
                    &format!("rId{}", index + 1),
                    name,
                ),
            )
        })
        .collect::<Vec<_>>();
    let distinct_data = record_stream(&distinct_records).unwrap();
    let mut distinct_iter = Records::new(&distinct_data, crate::raw::Limits::DEFAULT);
    let info = Workbook::read_workbook(
        &mut distinct_iter,
        &mut Default::default(),
        crate::raw::Limits::DEFAULT,
    )
    .unwrap();
    assert_eq!(info.worksheet_names, canonically_equivalent_names);
}

/// Builds a workbook part from one letter per record: `B`/`E` open and close
/// the book-view block (`b`/`e` carry a stray payload byte), `V` is a book
/// view (`v` is one of the wrong length) and `X` is an unrelated record. One
/// sheet follows, so that a view's sheet positions are in range.
fn book_view_stream(letters: &str) -> Vec<u8> {
    let mut records: Vec<(Kind, Vec<u8>)> = letters
        .chars()
        .map(|letter| match letter {
            'B' => (kind::BEGIN_BOOK_VIEWS, Vec::new()),
            'b' => (kind::BEGIN_BOOK_VIEWS, vec![0]),
            'E' => (kind::END_BOOK_VIEWS, Vec::new()),
            'e' => (kind::END_BOOK_VIEWS, vec![0]),
            'V' => (kind::BOOK_VIEW, vec![0; 29]),
            'v' => (kind::BOOK_VIEW, vec![0; 28]),
            'X' => (kind::FILE_VERSION, Vec::new()),
            other => panic!("unknown test record letter {other}"),
        })
        .collect();
    records.push((kind::BUNDLE_SH, bundle_sheet_payload(1, "rId1", "Sheet1")));
    record_stream(&records).unwrap()
}

/// The book-view block is fed every short order of its records: each alone,
/// every pair and triple with duplicates, and an unrelated record before,
/// inside and after.
#[test]
fn book_view_block_accepts_and_rejects_record_orders() {
    let mut cases = vec![String::new()];
    let mut previous = vec![String::new()];
    for _ in 0..3 {
        let mut next = Vec::new();
        for prefix in &previous {
            for letter in ['B', 'V', 'E'] {
                next.push(format!("{prefix}{letter}"));
            }
        }
        cases.extend(next.iter().cloned());
        previous = next;
    }
    cases.extend(
        [
            "BVEB", "BVEV", "BVEE", "BEBE", "BVVE", "X", "BXE", "BXVE", "BVXE", "XBVEX", "BEXV",
            "BEXB", "BX", "b", "e", "v", "bE", "Be", "BvE", "BEb", "BEe", "BEv",
        ]
        .map(str::to_string),
    );
    let actual: Vec<String> = cases
        .iter()
        .map(|case| {
            let data = book_view_stream(case);
            let mut iter = Records::new(&data, crate::raw::Limits::DEFAULT);
            let outcome = match Workbook::read_workbook(
                &mut iter,
                &mut Default::default(),
                crate::raw::Limits::DEFAULT,
            ) {
                Ok(info) => format!("ok active={:?}", info.active_catalog_position),
                Err(error) => error.to_string(),
            };
            format!("{case} => {outcome}")
        })
        .collect();
    assert_eq!(actual, BOOK_VIEW_ORDER_EXPECTED);
}

const BOOK_VIEW_ORDER_EXPECTED: &[&str] = &[
    " => ok active=None",
    "B => Unrecognized BrtBeginBookViews: missing BrtEndBookViews record",
    "V => Unrecognized BrtBookView: record is outside the book-view block",
    "E => Unrecognized BrtEndBookViews: book-view block has no matching begin record",
    "BB => Unrecognized BrtBeginBookViews: duplicate or nested book-view block",
    "BV => Unrecognized BrtBeginBookViews: missing BrtEndBookViews record",
    "BE => ok active=None",
    "VB => Unrecognized BrtBookView: record is outside the book-view block",
    "VV => Unrecognized BrtBookView: record is outside the book-view block",
    "VE => Unrecognized BrtBookView: record is outside the book-view block",
    "EB => Unrecognized BrtEndBookViews: book-view block has no matching begin record",
    "EV => Unrecognized BrtEndBookViews: book-view block has no matching begin record",
    "EE => Unrecognized BrtEndBookViews: book-view block has no matching begin record",
    "BBB => Unrecognized BrtBeginBookViews: duplicate or nested book-view block",
    "BBV => Unrecognized BrtBeginBookViews: duplicate or nested book-view block",
    "BBE => Unrecognized BrtBeginBookViews: duplicate or nested book-view block",
    "BVB => Unrecognized BrtBeginBookViews: duplicate or nested book-view block",
    "BVV => Unrecognized BrtBeginBookViews: missing BrtEndBookViews record",
    "BVE => ok active=Some(0)",
    "BEB => Unrecognized BrtBeginBookViews: duplicate or nested book-view block",
    "BEV => Unrecognized BrtBookView: record is outside the book-view block",
    "BEE => Unrecognized BrtEndBookViews: book-view block has no matching begin record",
    "VBB => Unrecognized BrtBookView: record is outside the book-view block",
    "VBV => Unrecognized BrtBookView: record is outside the book-view block",
    "VBE => Unrecognized BrtBookView: record is outside the book-view block",
    "VVB => Unrecognized BrtBookView: record is outside the book-view block",
    "VVV => Unrecognized BrtBookView: record is outside the book-view block",
    "VVE => Unrecognized BrtBookView: record is outside the book-view block",
    "VEB => Unrecognized BrtBookView: record is outside the book-view block",
    "VEV => Unrecognized BrtBookView: record is outside the book-view block",
    "VEE => Unrecognized BrtBookView: record is outside the book-view block",
    "EBB => Unrecognized BrtEndBookViews: book-view block has no matching begin record",
    "EBV => Unrecognized BrtEndBookViews: book-view block has no matching begin record",
    "EBE => Unrecognized BrtEndBookViews: book-view block has no matching begin record",
    "EVB => Unrecognized BrtEndBookViews: book-view block has no matching begin record",
    "EVV => Unrecognized BrtEndBookViews: book-view block has no matching begin record",
    "EVE => Unrecognized BrtEndBookViews: book-view block has no matching begin record",
    "EEB => Unrecognized BrtEndBookViews: book-view block has no matching begin record",
    "EEV => Unrecognized BrtEndBookViews: book-view block has no matching begin record",
    "EEE => Unrecognized BrtEndBookViews: book-view block has no matching begin record",
    "BVEB => Unrecognized BrtBeginBookViews: duplicate or nested book-view block",
    "BVEV => Unrecognized BrtBookView: record is outside the book-view block",
    "BVEE => Unrecognized BrtEndBookViews: book-view block has no matching begin record",
    "BEBE => Unrecognized BrtBeginBookViews: duplicate or nested book-view block",
    "BVVE => ok active=Some(0)",
    "X => ok active=None",
    "BXE => ok active=None",
    "BXVE => ok active=Some(0)",
    "BVXE => ok active=Some(0)",
    "XBVEX => ok active=Some(0)",
    "BEXV => Unrecognized BrtBookView: record is outside the book-view block",
    "BEXB => Unrecognized BrtBeginBookViews: duplicate or nested book-view block",
    "BX => Unrecognized BrtBeginBookViews: missing BrtEndBookViews record",
    "b => Invalid length: expected 0, found 1",
    "e => Invalid length: expected 0, found 1",
    "v => Unrecognized BrtBookView: record is outside the book-view block",
    "bE => Invalid length: expected 0, found 1",
    "Be => Invalid length: expected 0, found 1",
    "BvE => Invalid length: expected 29, found 28",
    "BEb => Invalid length: expected 0, found 1",
    "BEe => Invalid length: expected 0, found 1",
    "BEv => Unrecognized BrtBookView: record is outside the book-view block",
];
