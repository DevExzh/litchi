#![cfg(feature = "iwork")]

use litchi::iwork::{
    CellView, Document, ErrorKind, Format, Options, Resource, SnapshotLimits, Stage, Value,
};

const FIXTURE_HEX: &str =
    include_str!("../../../test-data/synthetic-iwork/numbers/compatibility-oracles.hex");

fn fixture_bytes() -> Vec<u8> {
    let digits = FIXTURE_HEX
        .bytes()
        .filter(|byte| !byte.is_ascii_whitespace())
        .collect::<Vec<_>>();
    assert_eq!(digits.len(), 1_070, "synthetic fixture hex width changed");
    assert!(digits.len().is_multiple_of(2));
    digits
        .chunks_exact(2)
        .map(|pair| (hex_nibble(pair[0]) << 4) | hex_nibble(pair[1]))
        .collect()
}

fn hex_nibble(byte: u8) -> u8 {
    match byte {
        b'0'..=b'9' => byte - b'0',
        b'a'..=b'f' => byte - b'a' + 10,
        _ => panic!("synthetic fixture contains a non-lowercase-hex byte"),
    }
}

fn exact_three_table_options() -> Options {
    Options::default().with_snapshot(
        SnapshotLimits::new(
            3,
            SnapshotLimits::HARD_MAX_SLIDES,
            SnapshotLimits::HARD_MAX_SECTIONS,
            SnapshotLimits::HARD_MAX_TEXT_BYTES,
        )
        .unwrap_or_else(|error| panic!("exact fixture limits must be valid: {error}")),
    )
}

#[test]
fn root_numbers_preserves_the_global_compatibility_projection() {
    // The focused fixture generator locks this source graph independently:
    // rooted drawable order is canonical model 11 then model 10; global
    // identity order is canonical 10 then 11; legacy model 5 is detached;
    // model 10 carries identical primary 6001 and secondary 6000 payloads.
    let bytes = fixture_bytes();
    let document = Document::from_bytes_with_options(&bytes, exact_three_table_options())
        .unwrap_or_else(|error| panic!("synthetic Numbers fixture must decode: {error}"));
    assert_eq!(document.format(), Format::Numbers);

    let snapshot = document.snapshot();
    assert_eq!(snapshot.table_count(), 3);
    assert_eq!(snapshot.slide_count(), 0);
    assert_eq!(snapshot.section_count(), 0);

    let tables = snapshot.tables().collect::<Vec<_>>();
    assert_eq!(
        tables.iter().map(|table| table.name()).collect::<Vec<_>>(),
        ["First canonical", "Second canonical", "Detached legacy"]
    );
    assert_eq!(
        tables
            .iter()
            .map(|table| table.position())
            .collect::<Vec<_>>(),
        [0, 1, 2]
    );

    let numeric = &tables[0];
    assert_eq!((numeric.row_count(), numeric.column_count()), (1, 2));
    assert_eq!(numeric.cell_count(), 1);
    assert_eq!(numeric.non_empty_cell_count(), 1);
    assert_eq!(numeric.cell(0, 0), Some(CellView::Missing));
    let Some(CellView::Stored(Value::Number(number))) = numeric.cell(0, 1) else {
        panic!("type-nine B1 was not retained as a number");
    };
    assert!(number.is_finite());
    assert_eq!(number.to_bits(), (-1_234.5_f64).to_bits());
    assert_eq!(numeric.cell(1, 0), None);

    for table in &tables[1..] {
        assert_eq!((table.row_count(), table.column_count()), (1, 1));
        assert_eq!(table.cell_count(), 0);
        assert_eq!(table.non_empty_cell_count(), 0);
        assert_eq!(table.cell(0, 0), Some(CellView::Missing));
    }

    // Three results prove the canonical model's secondary legacy payload was
    // deduplicated; its position before lower-ID legacy model 5 proves the
    // canonical-before-legacy candidate groups are preserved.
    assert_eq!(
        snapshot.all_text(),
        ["First canonical", "Second canonical", "Detached legacy"]
    );
}

#[test]
fn root_numbers_table_limit_is_inclusive_and_reports_the_first_excess() {
    let bytes = fixture_bytes();
    let exact = Document::from_bytes_with_options(&bytes, exact_three_table_options())
        .unwrap_or_else(|error| panic!("the exact three-table ceiling must pass: {error}"));
    assert_eq!(exact.snapshot().table_count(), 3);

    let one_under = Options::default().with_snapshot(
        SnapshotLimits::new(
            2,
            SnapshotLimits::HARD_MAX_SLIDES,
            SnapshotLimits::HARD_MAX_SECTIONS,
            SnapshotLimits::HARD_MAX_TEXT_BYTES,
        )
        .unwrap_or_else(|error| panic!("one-under fixture limits must be valid: {error}")),
    );
    let error = Document::from_bytes_with_options(&bytes, one_under)
        .err()
        .unwrap_or_else(|| panic!("the one-under table ceiling must fail"));
    assert_eq!(error.kind(), ErrorKind::LimitExceeded);
    assert_eq!(error.stage(), Stage::Semantic);
    assert_eq!(error.format(), Some(Format::Numbers));
    assert_eq!(error.resource(), Some(Resource::Tables));
    assert_eq!(error.observed(), Some(3));
    assert_eq!(error.maximum(), Some(2));
}

#[test]
fn root_numbers_propagates_a_late_legacy_text_limit_without_authored_content() {
    let bytes = fixture_bytes();
    let canonical_name_bytes = "First canonical".len() + "Second canonical".len();
    assert_eq!(canonical_name_bytes, 31);
    let options = Options::default().with_snapshot(
        SnapshotLimits::new(
            3,
            SnapshotLimits::HARD_MAX_SLIDES,
            SnapshotLimits::HARD_MAX_SECTIONS,
            canonical_name_bytes,
        )
        .unwrap_or_else(|error| panic!("legacy text profile must be valid: {error}")),
    );
    let error = Document::from_bytes_with_options(&bytes, options)
        .err()
        .unwrap_or_else(|| panic!("the detached legacy name must exceed the text ceiling"));
    assert_eq!(error.kind(), ErrorKind::LimitExceeded);
    assert_eq!(error.stage(), Stage::Semantic);
    assert_eq!(error.format(), Some(Format::Numbers));
    assert_eq!(error.resource(), Some(Resource::TextBytes));
    assert_eq!(error.observed(), Some(46));
    assert_eq!(error.maximum(), Some(31));
    assert!(std::error::Error::source(&error).is_none());
    let diagnostic = format!("{error:?} {error}");
    for authored in ["First canonical", "Second canonical", "Detached legacy"] {
        assert!(!diagnostic.contains(authored));
    }
}
