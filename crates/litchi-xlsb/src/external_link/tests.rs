use super::*;
use crate::raw::{Header, Kind as RecordKind, Limits, Records, Writer};

fn workbook_link() -> Link {
    let sheets = SheetRange::sheets(0, 1).unwrap();
    let formula = NameFormula::cell_reference(CellReference::new(sheets, CellLocation::new(2, 3)));
    Link::workbook_with_defined_names(
        "unused-host-source",
        vec!["Data".to_string(), "Rates".to_string()],
        vec![DefinedName::new("Rate").unwrap().with_formula(formula)],
    )
    .unwrap()
}

fn cache() -> ValueMatrix {
    ValueMatrix::new(
        1,
        3,
        vec![
            CachedValue::Number(7.0),
            CachedValue::Boolean(true),
            CachedValue::String("ready".to_string()),
        ],
    )
    .unwrap()
}

fn insert_unknown_after_first_record(source: &[u8]) -> Vec<u8> {
    let limits = Limits::new(MAX_LINK_PART_BYTES, MAX_WIDE_STRING_UNITS);
    let mut records = Records::with_limits(source, limits);
    let first = records.next().unwrap().unwrap();
    let (_, header_len) = Header::parse(&source[first.offset()..], limits).unwrap();
    let split = first.offset() + header_len + first.len();
    let mut opaque = Vec::new();
    Writer::new(&mut opaque)
        .write_record(RecordKind::new(0x3FFE).unwrap(), &[0xF1, 0x00, 0x7E])
        .unwrap();

    let mut result = Vec::with_capacity(source.len() + opaque.len());
    result.extend_from_slice(&source[..split]);
    result.extend_from_slice(&opaque);
    result.extend_from_slice(&source[split..]);
    result
}

#[test]
fn workbook_edits_are_source_checked_and_reversible() {
    let link = workbook_link();
    let source = write_external_link_stream(&link, Some("rIdOld")).unwrap();
    let snapshot = Snapshot::read(&source).unwrap();

    let mut edit = snapshot.edit();
    edit.set_relationship_id("rIdNew").unwrap();
    edit.set_sheet_names(vec![
        "Data".to_string(),
        "Rates".to_string(),
        "Audit".to_string(),
    ])
    .unwrap();
    edit.upsert_defined_name(DefinedName::new("Total").unwrap())
        .unwrap();
    let commit = edit.commit().unwrap();

    assert!(!commit.patch().is_empty());
    assert_eq!(commit.snapshot().relationship_id(), Some("rIdNew"));
    assert_eq!(commit.snapshot().link().defined_names().len(), 2);

    let changed = commit.patch().apply(&source).unwrap();
    assert_eq!(changed, commit.patch().after());
    let restored = commit.patch().inverse().apply(&changed).unwrap();
    assert_eq!(restored, source);
    assert!(commit.patch().apply(b"stale").is_err());
}

#[test]
fn unknown_records_survive_typed_edits_byte_for_byte() {
    let source = write_external_link_stream(&workbook_link(), Some("rIdPath")).unwrap();
    let source = insert_unknown_after_first_record(&source);
    let snapshot = Snapshot::read(&source).unwrap();
    assert_eq!(snapshot.unknown_records().len(), 1);
    assert_eq!(snapshot.unknown_records()[0].kind(), 0x3FFE);
    assert_eq!(snapshot.unknown_records()[0].payload(), [0xF1, 0x00, 0x7E]);

    let mut edit = snapshot.edit();
    edit.set_source("rIdChanged").unwrap();
    let commit = edit.commit().unwrap();
    assert!(
        commit
            .snapshot()
            .source_bytes()
            .windows(snapshot.unknown_records()[0].bytes().len())
            .any(|window| window == snapshot.unknown_records()[0].bytes())
    );
}

#[test]
fn dde_and_ole_metadata_and_caches_edit_without_activation() {
    let dde = Link::dde_with_items(
        "Excel",
        "System",
        vec![DdeItem::new("Status").unwrap().with_cached_values(cache())],
    )
    .unwrap();
    let source = write_external_link_stream(&dde, None).unwrap();
    let mut edit = Snapshot::read(&source).unwrap().edit();
    edit.set_dde_flags("status", true, true, false).unwrap();
    edit.set_dde_cache("Status", None).unwrap();
    edit.set_dde_topic("Sheet1").unwrap();
    let commit = edit.commit().unwrap();
    let changed = Snapshot::read(commit.patch().apply(&source).unwrap().as_slice()).unwrap();
    assert_eq!(changed.link().dde_topic(), Some("Sheet1"));
    assert!(changed.link().dde_items()[0].wants_advise());
    assert!(changed.link().dde_items()[0].cached_values().is_none());

    let ole = Link::ole_with_items(
        "rIdOle",
        "Acme.Server",
        vec![OleItem::new("Report").unwrap().with_cached_values(cache())],
    )
    .unwrap();
    let source = write_external_link_stream(&ole, Some("rIdOle")).unwrap();
    let mut edit = Snapshot::read(&source).unwrap().edit();
    edit.set_ole_program_id("Acme.NewServer").unwrap();
    edit.set_ole_flags("Report", true, true, true).unwrap();
    edit.set_ole_cache("Report", Some(cache())).unwrap();
    let commit = edit.commit().unwrap();
    let changed_bytes = commit.patch().apply(&source).unwrap();
    let changed = Snapshot::read(&changed_bytes).unwrap();
    assert_eq!(changed.link().ole_program_id(), Some("Acme.NewServer"));
    assert!(changed.link().ole_items()[0].displays_as_icon());
    assert!(changed.link().ole_items()[0].cached_values().is_some());
}

#[test]
fn failed_typed_operations_are_atomic_and_no_op_is_exact() {
    let source = write_external_link_stream(&workbook_link(), Some("rIdPath")).unwrap();
    let snapshot = Snapshot::read(&source).unwrap();

    let mut edit = snapshot.edit();
    let before = edit.link().clone();
    assert!(edit.set_dde_topic("System").is_err());
    assert_eq!(edit.link(), &before);

    let commit = snapshot.edit().commit().unwrap();
    assert!(commit.patch().is_empty());
    assert_eq!(commit.patch().before(), source);
    assert_eq!(commit.patch().after(), source);
    assert_eq!(commit.patch().apply(&source).unwrap(), source);
}

#[test]
fn ownership_and_cache_bounds_are_rejected() {
    let dde = Link::dde("Excel", "System", vec!["Status".to_string()]).unwrap();
    assert!(write_external_link_stream(&dde, Some("rIdDde")).is_err());
    assert!(ValueMatrix::new(MAX_XLSB_EXTERNAL_CACHE_ROWS + 1, 1, vec![]).is_err());
    assert!(ValueMatrix::new(1, MAX_XLSB_EXTERNAL_CACHE_COLUMNS + 1, vec![]).is_err());
    assert!(
        ValueMatrix::new(
            MAX_XLSB_EXTERNAL_CACHE_ROWS,
            MAX_XLSB_EXTERNAL_CACHE_COLUMNS,
            vec![],
        )
        .is_err()
    );
}
