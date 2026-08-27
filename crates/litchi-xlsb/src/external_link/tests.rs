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
    insert_unknown_records_after_first_record(
        source,
        &[(RecordKind::new(0x3FFE).unwrap(), &[0xF1, 0x00, 0x7E])],
    )
}

fn insert_unknown_records_after_first_record(
    source: &[u8],
    records_to_insert: &[(RecordKind, &[u8])],
) -> Vec<u8> {
    insert_unknown_records_after_known_record(source, 1, records_to_insert)
}

fn insert_unknown_records_after_known_record(
    source: &[u8],
    after_known: usize,
    records_to_insert: &[(RecordKind, &[u8])],
) -> Vec<u8> {
    let limits = Limits::new(MAX_LINK_PART_BYTES, MAX_WIDE_STRING_UNITS);
    let mut known = 0;
    let mut split = source.len();
    for record in Records::with_limits(source, limits) {
        let record = record.unwrap();
        if !codec::is_modeled_record(record.kind()) {
            continue;
        }
        if known == after_known {
            split = record.offset();
            break;
        }
        known += 1;
    }
    let mut opaque = Vec::new();
    let mut writer = Writer::new(&mut opaque);
    for &(kind, payload) in records_to_insert {
        writer.write_record(kind, payload).unwrap();
    }

    let mut result = Vec::with_capacity(source.len() + opaque.len());
    result.extend_from_slice(&source[..split]);
    result.extend_from_slice(&opaque);
    result.extend_from_slice(&source[split..]);
    result
}

fn insert_repeated_unknown_after_first_record(
    source: &[u8],
    count: usize,
    payload: &[u8],
) -> Vec<u8> {
    let limits = Limits::new(MAX_LINK_PART_BYTES, MAX_WIDE_STRING_UNITS);
    let mut records = Records::with_limits(source, limits);
    let first = records.next().unwrap().unwrap();
    let (_, header_len) = Header::parse(&source[first.offset()..], limits).unwrap();
    let split = first.offset() + header_len + first.len();
    let mut opaque = Vec::new();
    let mut writer = Writer::new(&mut opaque);
    let kind = RecordKind::new(0x3FFE).unwrap();
    for _ in 0..count {
        writer.write_record(kind, payload).unwrap();
    }

    let mut result = Vec::with_capacity(source.len() + opaque.len());
    result.extend_from_slice(&source[..split]);
    result.extend_from_slice(&opaque);
    result.extend_from_slice(&source[split..]);
    result
}

fn replace_record_payload(
    source: &[u8],
    target_kind: RecordKind,
    occurrence: usize,
    payload: &[u8],
) -> Vec<u8> {
    let limits = Limits::new(MAX_LINK_PART_BYTES, MAX_WIDE_STRING_UNITS);
    let mut records = Records::with_limits(source, limits);
    let mut seen = 0;
    let mut range = None;
    while let Some(record) = records.next() {
        let record = record.unwrap();
        if record.kind() != target_kind {
            continue;
        }
        if seen == occurrence {
            let (_, header_len) = Header::parse(&source[record.offset()..], limits).unwrap();
            range = Some((record.offset(), record.offset() + header_len + record.len()));
            break;
        }
        seen += 1;
    }
    let (start, end) = range.unwrap();
    let mut replacement = Vec::new();
    Writer::new(&mut replacement)
        .write_record(target_kind, payload)
        .unwrap();

    let mut result = Vec::with_capacity(source.len() - (end - start) + replacement.len());
    result.extend_from_slice(&source[..start]);
    result.extend_from_slice(&replacement);
    result.extend_from_slice(&source[end..]);
    result
}

fn first_sup_name_bits(source: &[u8]) -> [u8; 7] {
    let limits = Limits::new(MAX_LINK_PART_BYTES, MAX_WIDE_STRING_UNITS);
    for record in Records::with_limits(source, limits) {
        let record = record.unwrap();
        if record.kind() == crate::raw::kind::SUP_NAME_BITS {
            return record.payload().try_into().unwrap();
        }
    }
    panic!("source has no BrtSupNameBits record");
}

fn record_kinds(source: &[u8]) -> Vec<u16> {
    let limits = Limits::new(MAX_LINK_PART_BYTES, MAX_WIDE_STRING_UNITS);
    Records::with_limits(source, limits)
        .map(|record| u16::from(record.unwrap().kind()))
        .collect()
}

fn assert_opaque_cache_run_is_between_tabs_and_name(
    source: &[u8],
    opaque_specs: &[(RecordKind, &[u8])],
) {
    let kinds = record_kinds(source);
    let tabs = kinds
        .iter()
        .position(|kind| *kind == u16::from(crate::raw::kind::SUP_TABS))
        .unwrap();
    let cache_start = tabs + 1;
    let cache_end = cache_start + opaque_specs.len();
    let expected_kinds: Vec<_> = opaque_specs
        .iter()
        .map(|(kind, _)| u16::from(*kind))
        .collect();
    assert_eq!(&kinds[cache_start..cache_end], expected_kinds.as_slice());
    assert_eq!(
        kinds[cache_end],
        u16::from(crate::raw::kind::SUP_NAME_START)
    );
    assert_eq!(
        kinds[cache_end + 1],
        u16::from(crate::raw::kind::SUP_NAME_FORMULA)
    );
    assert_eq!(
        kinds[cache_end + 2],
        u16::from(crate::raw::kind::SUP_NAME_BITS)
    );
    assert_eq!(
        kinds[cache_end + 3],
        u16::from(crate::raw::kind::SUP_NAME_END)
    );
    assert_eq!(
        kinds[cache_end + 4],
        u16::from(crate::raw::kind::END_SUP_BOOK)
    );
}

fn assert_external_formula_length_error(source: &[u8], found: usize) {
    match parse_external_link(source).unwrap_err() {
        Error::InvalidLength {
            expected: 13,
            found: actual,
        } => assert_eq!(actual, found),
        other @ (Error::Wire(_)
        | Error::InvalidFormula(_)
        | Error::InvalidLength { .. }
        | Error::Allocation { .. }
        | Error::LimitExceeded { .. }) => {
            panic!("unexpected external formula error: {other:?}")
        },
    }
}

fn assert_opaque_sequence_in_order(source: &[u8], records: &[UnknownRecord]) {
    let mut search_start = 0;
    for record in records {
        let offset = source[search_start..]
            .windows(record.bytes().len())
            .position(|window| window == record.bytes())
            .map(|offset| search_start + offset)
            .unwrap();
        search_start = offset + record.bytes().len();
    }
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

#[test]
fn workbook_sup_name_bits_preserve_ignored_bytes_and_change_only_modeled_fields() {
    let source = write_external_link_stream(&workbook_link(), Some("rIdPath")).unwrap();
    let source = replace_record_payload(
        &source,
        crate::raw::kind::SUP_NAME_BITS,
        0,
        &[0x00, 0xA5, 0x00, 0x00, 0x00, 0x00, 0xFE],
    );
    assert_eq!(
        first_sup_name_bits(&source),
        [0x00, 0xA5, 0x00, 0x00, 0x00, 0x00, 0xFE]
    );

    let snapshot = Snapshot::read(&source).unwrap();
    let no_op = snapshot.edit().commit().unwrap();
    assert_eq!(no_op.patch().before(), source);
    assert_eq!(no_op.patch().after(), source);

    let mut unrelated = snapshot.edit();
    unrelated.set_source("rIdChanged").unwrap();
    let unrelated = unrelated.commit().unwrap();
    let unrelated_bytes = unrelated.patch().after();
    assert_eq!(
        first_sup_name_bits(unrelated_bytes),
        [0x00, 0xA5, 0x00, 0x00, 0x00, 0x00, 0xFE]
    );

    let mut modeled = Snapshot::read(unrelated_bytes).unwrap().edit();
    let mut name = modeled.link().defined_names()[0].clone();
    name = name.with_built_in(true).with_sheet_scope(1);
    modeled.set_defined_names(vec![name]).unwrap();
    let modeled = modeled.commit().unwrap();
    assert_eq!(
        first_sup_name_bits(modeled.patch().after()),
        [0x01, 0xA5, 0x02, 0x00, 0x00, 0x00, 0xFE]
    );
}

#[test]
fn dde_sup_name_bits_preserve_ignored_bytes_and_change_only_modeled_fields() {
    let dde = Link::dde_with_items(
        "Excel",
        "System",
        vec![DdeItem::new("StdDocumentName").unwrap()],
    )
    .unwrap();
    let source = write_external_link_stream(&dde, None).unwrap();
    let source = replace_record_payload(
        &source,
        crate::raw::kind::SUP_NAME_BITS,
        0,
        &[0x00, 0xA5, 0x00, 0x00, 0x00, 0x00, 0xA1],
    );
    let snapshot = Snapshot::read(&source).unwrap();
    let no_op = snapshot.edit().commit().unwrap();
    assert_eq!(no_op.patch().before(), source);
    assert_eq!(no_op.patch().after(), source);

    let mut unrelated = snapshot.edit();
    unrelated.set_dde_topic("Sheet2").unwrap();
    let unrelated = unrelated.commit().unwrap();
    assert_eq!(
        first_sup_name_bits(unrelated.patch().after()),
        [0x00, 0xA5, 0x00, 0x00, 0x00, 0x00, 0xA1]
    );

    let mut modeled = Snapshot::read(unrelated.patch().after()).unwrap().edit();
    modeled
        .set_dde_flags("StdDocumentName", true, true, true)
        .unwrap();
    let modeled = modeled.commit().unwrap();
    assert_eq!(
        first_sup_name_bits(modeled.patch().after()),
        [0x0E, 0xA5, 0x00, 0x00, 0x00, 0x00, 0xA1]
    );
}

#[test]
fn ole_sup_name_bits_preserve_ignored_bytes_and_change_only_modeled_fields() {
    let ole = Link::ole_with_items(
        "rIdOle",
        "Acme.Server",
        vec![OleItem::new("Report").unwrap()],
    )
    .unwrap();
    let source = write_external_link_stream(&ole, Some("rIdOle")).unwrap();
    let source = replace_record_payload(
        &source,
        crate::raw::kind::SUP_NAME_BITS,
        0,
        &[0x10, 0xA5, 0x00, 0x00, 0x00, 0x00, 0xA1],
    );
    let snapshot = Snapshot::read(&source).unwrap();
    let no_op = snapshot.edit().commit().unwrap();
    assert_eq!(no_op.patch().before(), source);
    assert_eq!(no_op.patch().after(), source);

    let mut unrelated = snapshot.edit();
    unrelated.set_ole_program_id("Acme.NewServer").unwrap();
    let unrelated = unrelated.commit().unwrap();
    assert_eq!(
        first_sup_name_bits(unrelated.patch().after()),
        [0x10, 0xA5, 0x00, 0x00, 0x00, 0x00, 0xA1]
    );

    let mut modeled = Snapshot::read(unrelated.patch().after()).unwrap().edit();
    modeled.set_ole_flags("Report", true, true, true).unwrap();
    let modeled = modeled.commit().unwrap();
    assert_eq!(
        first_sup_name_bits(modeled.patch().after()),
        [0x36, 0xA5, 0x00, 0x00, 0x00, 0x00, 0xA1]
    );
}

#[test]
fn oversized_external_name_formulas_are_rejected_before_copy() {
    let source = write_external_link_stream(&workbook_link(), Some("rIdPath")).unwrap();
    let mut oversized = 14_u32.to_le_bytes().to_vec();
    oversized.extend_from_slice(&[0x00; 14]);
    let oversized =
        replace_record_payload(&source, crate::raw::kind::SUP_NAME_FORMULA, 0, &oversized);
    assert_external_formula_length_error(&oversized, 14);

    let huge = replace_record_payload(
        &source,
        crate::raw::kind::SUP_NAME_FORMULA,
        0,
        &u32::MAX.to_le_bytes(),
    );
    assert_external_formula_length_error(&huge, u32::MAX as usize);
}

#[test]
fn opaque_external_table_records_and_future_frames_survive_in_order() {
    let source = write_external_link_stream(&workbook_link(), Some("rIdPath")).unwrap();
    let opaque_specs = [
        (crate::raw::kind::EXTERN_TABLE_START, &[0x00; 5][..]),
        (crate::raw::kind::EXTERN_ROW_HDR, &[0x00; 4][..]),
        (crate::raw::kind::EXTERN_CELL_BLANK, &[0x00; 4][..]),
        (
            crate::raw::kind::EXTERN_CELL_REAL,
            &[
                0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
            ][..],
        ),
        (
            crate::raw::kind::EXTERN_CELL_BOOL,
            &[0x02, 0x00, 0x00, 0x00, 0x01][..],
        ),
        (
            crate::raw::kind::EXTERN_CELL_ERROR,
            &[0x03, 0x00, 0x00, 0x00, 0x07][..],
        ),
        (
            crate::raw::kind::EXTERN_CELL_STRING,
            &[0x04, 0x00, 0x00, 0x00, 1, 0, 0, 0, b'x', 0][..],
        ),
        (crate::raw::kind::EXTERN_TABLE_END, &[][..]),
        (RecordKind::new(0x3FFD).unwrap(), &[0xD0, 0xD1][..]),
    ];
    let source = insert_unknown_records_after_known_record(&source, 2, &opaque_specs);
    let snapshot = Snapshot::read(&source).unwrap();
    assert_eq!(snapshot.unknown_records().len(), opaque_specs.len());
    for (record, (kind, payload)) in snapshot.unknown_records().iter().zip(opaque_specs) {
        assert_eq!(record.kind(), u16::from(kind));
        assert_eq!(record.payload(), payload);
    }
    assert_opaque_cache_run_is_between_tabs_and_name(&source, &opaque_specs);

    let no_op = snapshot.edit().commit().unwrap();
    assert_eq!(no_op.patch().before(), source);
    assert_eq!(no_op.patch().after(), source);

    let mut edit = snapshot.edit();
    edit.set_source("rIdOpaqueChanged").unwrap();
    let commit = edit.commit().unwrap();
    assert_opaque_sequence_in_order(commit.patch().after(), snapshot.unknown_records());
    assert_opaque_cache_run_is_between_tabs_and_name(commit.patch().after(), &opaque_specs);
}

#[test]
fn opaque_external_record_count_and_byte_limits_are_typed() {
    let source = write_external_link_stream(&workbook_link(), Some("rIdPath")).unwrap();
    let too_many =
        insert_repeated_unknown_after_first_record(&source, MAX_UNKNOWN_RECORDS + 1, &[0xB0]);
    match Snapshot::read(&too_many).unwrap_err() {
        Error::InvalidLength { expected, found } => {
            assert_eq!(expected, MAX_UNKNOWN_RECORDS);
            assert_eq!(found, MAX_UNKNOWN_RECORDS + 1);
        },
        other @ (Error::Wire(_)
        | Error::InvalidFormula(_)
        | Error::Allocation { .. }
        | Error::LimitExceeded { .. }) => {
            panic!("unexpected unknown-record count error: {other:?}")
        },
    }

    let mut one_record = Vec::new();
    Writer::new(&mut one_record)
        .write_record(RecordKind::new(0x3FFE).unwrap(), &[0xC0; 128])
        .unwrap();
    let count = MAX_UNKNOWN_BYTES / one_record.len() + 1;
    assert!(count <= MAX_UNKNOWN_RECORDS);
    let too_many_bytes = insert_repeated_unknown_after_first_record(&source, count, &[0xC0; 128]);
    match Snapshot::read(&too_many_bytes).unwrap_err() {
        Error::InvalidLength { expected, found } => {
            assert_eq!(expected, MAX_UNKNOWN_BYTES);
            assert_eq!(found, count * one_record.len());
        },
        other @ (Error::Wire(_)
        | Error::InvalidFormula(_)
        | Error::Allocation { .. }
        | Error::LimitExceeded { .. }) => {
            panic!("unexpected unknown-record byte error: {other:?}")
        },
    }
}

fn assert_limit_exceeded<T>(
    result: Result<T>,
    resource: ExternalLinkResource,
    actual: usize,
    maximum: usize,
) {
    let Err(error) = result else {
        panic!("expected {resource} limit failure")
    };
    let Error::LimitExceeded {
        resource: found_resource,
        actual: found_actual,
        maximum: found_maximum,
    } = error
    else {
        panic!("expected {resource} limit failure, found {error:?}")
    };
    assert_eq!(found_resource, resource);
    assert_eq!(found_actual, actual);
    assert_eq!(found_maximum, maximum);
}

#[test]
fn external_link_limits_builder_and_default_are_publicly_parity_checked() {
    let default_from_builder = ExternalLinkLimits::builder().build().unwrap();
    assert_eq!(default_from_builder, ExternalLinkLimits::DEFAULT);
    assert_eq!(ExternalLinkLimits::default(), default_from_builder);

    let primary = ExternalLinkLimits::builder()
        .max_part_bytes(64)
        .max_total_part_bytes(128)
        .max_opaque_bytes(8)
        .max_total_opaque_bytes(16)
        .max_utf16_units(2)
        .max_total_utf16_units(4)
        .max_records(3)
        .max_cache_records(4)
        .max_opaque_records(5)
        .max_links(6)
        .max_items(7)
        .max_matrices(8)
        .max_cells(9)
        .max_decoded_semantic_bytes(10)
        .max_retained_objects(11)
        .build()
        .unwrap();
    let aliases = ExternalLinkLimits::builder()
        .max_part_bytes(64)
        .max_total_part_bytes(128)
        .max_opaque_bytes(8)
        .max_total_opaque_bytes(16)
        .max_utf16_units(2)
        .max_total_utf16_units(4)
        .max_total_records(3)
        .max_total_cache_records(4)
        .max_total_opaque_records(5)
        .max_total_links(6)
        .max_total_items(7)
        .max_total_matrices(8)
        .max_total_cells(9)
        .max_total_decoded_semantic_bytes(10)
        .max_total_retained_objects(11)
        .build()
        .unwrap();
    assert_eq!(primary, aliases);
    assert_eq!(primary.max_total_records(), primary.max_records());
    assert_eq!(
        primary.max_total_cache_records(),
        primary.max_cache_records()
    );
    assert_eq!(
        primary.max_total_opaque_records(),
        primary.max_opaque_records()
    );
    assert_eq!(primary.max_total_links(), primary.max_links());
    assert_eq!(primary.max_total_items(), primary.max_items());
    assert_eq!(primary.max_total_matrices(), primary.max_matrices());
    assert_eq!(primary.max_total_cells(), primary.max_cells());
    assert_eq!(
        primary.max_total_decoded_semantic_bytes(),
        primary.max_decoded_semantic_bytes()
    );
    assert_eq!(
        primary.max_total_retained_objects(),
        primary.max_retained_objects()
    );
}

#[test]
fn external_link_part_limit_accepts_exact_bytes_and_reports_one_below() {
    let source = write_external_link_stream(&workbook_link(), Some("rIdPath")).unwrap();
    let exact = ExternalLinkLimits::builder()
        .max_part_bytes(source.len())
        .max_total_part_bytes(source.len())
        .build()
        .unwrap();
    assert!(parse_external_link_with_limits(&source, exact).is_ok());

    let below = ExternalLinkLimits::builder()
        .max_part_bytes(source.len() - 1)
        .max_total_part_bytes(source.len() - 1)
        .build()
        .unwrap();
    assert_limit_exceeded(
        parse_external_link_with_limits(&source, below),
        ExternalLinkResource::PartBytes,
        source.len(),
        source.len() - 1,
    );
}

#[test]
fn external_link_record_limit_accepts_exact_records_and_reports_one_below() {
    let source = write_external_link_stream(&workbook_link(), Some("rIdPath")).unwrap();
    let record_count = record_kinds(&source).len();
    let exact = ExternalLinkLimits::builder()
        .max_records(record_count)
        .build()
        .unwrap();
    assert!(parse_external_link_with_limits(&source, exact).is_ok());

    let below = ExternalLinkLimits::builder()
        .max_records(record_count - 1)
        .build()
        .unwrap();
    assert_limit_exceeded(
        parse_external_link_with_limits(&source, below),
        ExternalLinkResource::Records,
        record_count,
        record_count - 1,
    );
}

#[test]
fn external_link_utf16_limits_count_non_bmp_units_per_string_and_in_total() {
    let link = Link::dde("X", "😀", vec!["A".to_string()]).unwrap();
    let source = write_external_link_stream(&link, None).unwrap();
    let non_bmp_units = "😀".encode_utf16().count();
    let total_units =
        "X".encode_utf16().count() + "😀".encode_utf16().count() + "A".encode_utf16().count();
    assert_eq!(non_bmp_units, 2);

    let exact = ExternalLinkLimits::builder()
        .max_utf16_units(non_bmp_units)
        .max_total_utf16_units(total_units)
        .build()
        .unwrap();
    assert!(parse_external_link_with_limits(&source, exact).is_ok());

    let one_below = ExternalLinkLimits::builder()
        .max_utf16_units(non_bmp_units - 1)
        .max_total_utf16_units(total_units)
        .build()
        .unwrap();
    assert_limit_exceeded(
        parse_external_link_with_limits(&source, one_below),
        ExternalLinkResource::Utf16Units,
        non_bmp_units,
        non_bmp_units - 1,
    );

    let total_below = ExternalLinkLimits::builder()
        .max_utf16_units(non_bmp_units)
        .max_total_utf16_units(total_units - 1)
        .build()
        .unwrap();
    assert_limit_exceeded(
        parse_external_link_with_limits(&source, total_below),
        ExternalLinkResource::TotalUtf16Units,
        total_units,
        total_units - 1,
    );
}

#[test]
fn external_link_decoded_semantic_bytes_accept_exact_and_report_one_below() {
    let link = Link::dde("X", "Y", vec!["A".to_string()]).unwrap();
    let source = write_external_link_stream(&link, None).unwrap();
    let semantic_bytes = "X".len() + "Y".len() + "A".len();
    let exact = ExternalLinkLimits::builder()
        .max_decoded_semantic_bytes(semantic_bytes)
        .build()
        .unwrap();
    assert!(parse_external_link_with_limits(&source, exact).is_ok());

    let one_below = ExternalLinkLimits::builder()
        .max_decoded_semantic_bytes(semantic_bytes - 1)
        .build()
        .unwrap();
    assert_limit_exceeded(
        parse_external_link_with_limits(&source, one_below),
        ExternalLinkResource::DecodedSemanticBytes,
        semantic_bytes,
        semantic_bytes - 1,
    );
}

#[test]
fn external_link_item_limits_aggregate_workbook_names_and_dde_ole_items() {
    let workbook = Link::workbook_with_defined_names(
        "unused-host-source",
        vec!["Data".to_string()],
        vec![
            DefinedName::new("A").unwrap(),
            DefinedName::new("B").unwrap(),
        ],
    )
    .unwrap();
    let dde = Link::dde_with_items(
        "X",
        "Y",
        vec![DdeItem::new("A").unwrap(), DdeItem::new("B").unwrap()],
    )
    .unwrap();
    let ole = Link::ole_with_items(
        "rIdOle",
        "Acme.Server",
        vec![OleItem::new("A").unwrap(), OleItem::new("B").unwrap()],
    )
    .unwrap();

    for (link, relationship_id) in [
        (&workbook, Some("rIdPath")),
        (&dde, None),
        (&ole, Some("rIdOle")),
    ] {
        let source = write_external_link_stream(link, relationship_id).unwrap();
        let exact = ExternalLinkLimits::builder().max_items(2).build().unwrap();
        assert!(parse_external_link_with_limits(&source, exact).is_ok());

        let one_below = ExternalLinkLimits::builder().max_items(1).build().unwrap();
        assert_limit_exceeded(
            parse_external_link_with_limits(&source, one_below),
            ExternalLinkResource::Items,
            2,
            1,
        );
    }
}

#[test]
fn external_link_cache_limits_aggregate_matrices_and_cells_across_items() {
    let matrix = ValueMatrix::new(1, 1, vec![CachedValue::Empty]).unwrap();
    let link = Link::dde_with_items(
        "X",
        "Y",
        vec![
            DdeItem::new("A")
                .unwrap()
                .with_cached_values(matrix.clone()),
            DdeItem::new("B").unwrap().with_cached_values(matrix),
        ],
    )
    .unwrap();
    let source = write_external_link_stream(&link, None).unwrap();

    let exact = ExternalLinkLimits::builder()
        .max_matrices(2)
        .max_cells(2)
        .build()
        .unwrap();
    assert!(parse_external_link_with_limits(&source, exact).is_ok());

    let matrix_below = ExternalLinkLimits::builder()
        .max_matrices(1)
        .max_cells(2)
        .build()
        .unwrap();
    assert_limit_exceeded(
        parse_external_link_with_limits(&source, matrix_below),
        ExternalLinkResource::Matrices,
        2,
        1,
    );

    let cells_below = ExternalLinkLimits::builder()
        .max_matrices(2)
        .max_cells(1)
        .build()
        .unwrap();
    assert_limit_exceeded(
        parse_external_link_with_limits(&source, cells_below),
        ExternalLinkResource::Cells,
        2,
        1,
    );
}

#[test]
fn external_link_snapshot_limits_bound_opaque_records_and_bytes() {
    let source = insert_unknown_after_first_record(
        &write_external_link_stream(&workbook_link(), Some("rIdPath")).unwrap(),
    );
    let default_snapshot = Snapshot::read(&source).unwrap();
    let opaque_bytes = default_snapshot.unknown_records()[0].bytes().len();

    let exact = ExternalLinkLimits::builder()
        .max_opaque_bytes(opaque_bytes)
        .max_total_opaque_bytes(opaque_bytes)
        .max_opaque_records(1)
        .build()
        .unwrap();
    let snapshot = Snapshot::read_with_limits(&source, exact).unwrap();
    assert_eq!(snapshot.limits(), exact);
    assert_eq!(snapshot.unknown_records().len(), 1);

    let record_below = ExternalLinkLimits::builder()
        .max_opaque_records(0)
        .build()
        .unwrap();
    assert_limit_exceeded(
        Snapshot::read_with_limits(&source, record_below),
        ExternalLinkResource::OpaqueRecords,
        1,
        0,
    );

    let bytes_below = ExternalLinkLimits::builder()
        .max_opaque_bytes(opaque_bytes - 1)
        .max_total_opaque_bytes(opaque_bytes - 1)
        .build()
        .unwrap();
    assert_limit_exceeded(
        Snapshot::read_with_limits(&source, bytes_below),
        ExternalLinkResource::OpaqueBytes,
        opaque_bytes,
        opaque_bytes - 1,
    );
}

#[test]
fn external_link_snapshot_limits_are_reused_and_enforced_by_changed_commit() {
    let source = write_external_link_stream(&workbook_link(), Some("rIdPath")).unwrap();
    let maximum = source.len();
    let limits = ExternalLinkLimits::builder()
        .max_part_bytes(maximum)
        .max_total_part_bytes(maximum)
        .build()
        .unwrap();
    let snapshot = Snapshot::read_with_limits(&source, limits).unwrap();
    assert_eq!(snapshot.limits(), limits);

    let mut edit = snapshot.edit();
    let long_source = format!("rId{}", "x".repeat(256));
    edit.set_source(long_source).unwrap();
    let staged = edit.link();
    let expected_after = write_external_link_stream(staged, Some(staged.source())).unwrap();
    assert!(expected_after.len() > maximum);
    assert_limit_exceeded(
        edit.commit(),
        ExternalLinkResource::PartBytes,
        expected_after.len(),
        maximum,
    );
}

#[test]
fn external_link_apply_with_limits_is_fallible_and_accepts_exact_patch_bytes() {
    let source = write_external_link_stream(&workbook_link(), Some("rIdPath")).unwrap();
    let mut edit = Snapshot::read(&source).unwrap().edit();
    edit.set_source("rIdChangedToLonger").unwrap();
    let commit = edit.commit().unwrap();
    let patch = commit.patch();
    let maximum = patch.after().len();
    let exact = ExternalLinkLimits::builder()
        .max_part_bytes(maximum)
        .max_total_part_bytes(maximum)
        .build()
        .unwrap();
    assert_eq!(
        apply_with_limits(patch.before(), patch, exact).unwrap(),
        patch.after()
    );
    assert!(apply_with_limits(b"stale", patch, exact).is_err());

    let one_below = ExternalLinkLimits::builder()
        .max_part_bytes(maximum - 1)
        .max_total_part_bytes(maximum - 1)
        .build()
        .unwrap();
    assert_limit_exceeded(
        apply_with_limits(patch.before(), patch, one_below),
        ExternalLinkResource::PartBytes,
        maximum,
        maximum - 1,
    );
}

#[test]
fn external_link_malformed_wire_keeps_wire_and_structural_errors_typed() {
    let source = write_external_link_stream(&workbook_link(), Some("rIdPath")).unwrap();
    let truncated = replace_record_payload(
        &source,
        crate::raw::kind::SUP_NAME_START,
        0,
        &[2, 0, 0, 0, b'A', 0],
    );
    let wire_error =
        parse_external_link_with_limits(&truncated, ExternalLinkLimits::DEFAULT).unwrap_err();
    assert!(matches!(wire_error, Error::Wire(_)));

    let mut structurally_invalid = source;
    Writer::new(&mut structurally_invalid)
        .write_record(crate::raw::kind::SUP_NAME_END, &[])
        .unwrap();
    let structural_error =
        parse_external_link_with_limits(&structurally_invalid, ExternalLinkLimits::DEFAULT)
            .unwrap_err();
    assert!(matches!(structural_error, Error::InvalidFormula(_)));
}

#[test]
fn external_link_writer_part_limits_accept_exact_output_and_reject_one_below() {
    let link = workbook_link();
    let canonical = write_external_link_stream(&link, Some("rIdPath")).unwrap();
    let exact = ExternalLinkLimits::builder()
        .max_part_bytes(canonical.len())
        .max_total_part_bytes(canonical.len())
        .build()
        .unwrap();
    let bounded = write_external_link_stream_with_limits(&link, Some("rIdPath"), exact).unwrap();
    assert_eq!(bounded, canonical);

    let one_below = ExternalLinkLimits::builder()
        .max_part_bytes(canonical.len() - 1)
        .max_total_part_bytes(canonical.len() - 1)
        .build()
        .unwrap();
    assert_limit_exceeded(
        write_external_link_stream_with_limits(&link, Some("rIdPath"), one_below),
        ExternalLinkResource::PartBytes,
        canonical.len(),
        canonical.len() - 1,
    );
}

#[test]
fn external_link_writer_rejects_authored_semantic_limits_with_typed_resources() {
    let workbook = workbook_link();
    let items = ExternalLinkLimits::builder().max_items(0).build().unwrap();
    assert_limit_exceeded(
        write_external_link_stream_with_limits(&workbook, Some("rIdPath"), items),
        ExternalLinkResource::Items,
        1,
        0,
    );

    let matrix = ValueMatrix::new(1, 1, vec![CachedValue::Empty]).unwrap();
    let cached = Link::dde_with_items(
        "X",
        "Y",
        vec![DdeItem::new("A").unwrap().with_cached_values(matrix)],
    )
    .unwrap();
    let cells = ExternalLinkLimits::builder().max_cells(0).build().unwrap();
    assert_limit_exceeded(
        write_external_link_stream_with_limits(&cached, None, cells),
        ExternalLinkResource::Cells,
        1,
        0,
    );

    let short = Link::dde("X", "Y", vec!["A".to_string()]).unwrap();
    let utf16 = ExternalLinkLimits::builder()
        .max_utf16_units(0)
        .build()
        .unwrap();
    assert_limit_exceeded(
        write_external_link_stream_with_limits(&short, None, utf16),
        ExternalLinkResource::Utf16Units,
        1,
        0,
    );

    let semantic_bytes = ExternalLinkLimits::builder()
        .max_decoded_semantic_bytes(0)
        .build()
        .unwrap();
    assert_limit_exceeded(
        write_external_link_stream_with_limits(&short, None, semantic_bytes),
        ExternalLinkResource::DecodedSemanticBytes,
        1,
        0,
    );
}

#[test]
fn external_link_strict_snapshot_commit_rejects_larger_model_and_remains_retryable() {
    let link = Link::dde_with_items("X", "Y", vec![DdeItem::new("A").unwrap()]).unwrap();
    let source = write_external_link_stream(&link, None).unwrap();
    let limits = ExternalLinkLimits::builder().max_items(1).build().unwrap();
    let snapshot = Snapshot::read_with_limits(&source, limits).unwrap();

    let mut rejected = snapshot.edit();
    rejected
        .upsert_dde_item(DdeItem::new("B").unwrap())
        .unwrap();
    assert_limit_exceeded(rejected.commit(), ExternalLinkResource::Items, 2, 1);
    assert_eq!(snapshot.source_bytes(), source.as_slice());
    assert_eq!(snapshot.link(), &link);
    assert_eq!(snapshot.limits(), limits);

    let mut retry = snapshot.edit();
    retry.set_dde_topic("Z").unwrap();
    let committed = retry.commit().unwrap();
    assert_eq!(committed.snapshot().limits(), limits);
    assert_eq!(committed.snapshot().link().dde_topic(), Some("Z"));
}

#[test]
fn external_link_patch_apply_with_limits_rejects_semantic_after_image() {
    let link = Link::dde_with_items("X", "Y", vec![DdeItem::new("A").unwrap()]).unwrap();
    let source = write_external_link_stream(&link, None).unwrap();
    let mut edit = Snapshot::read(&source).unwrap().edit();
    edit.upsert_dde_item(DdeItem::new("B").unwrap()).unwrap();
    let commit = edit.commit().unwrap();
    let patch = commit.patch();
    let limits = ExternalLinkLimits::builder()
        .max_part_bytes(patch.after().len())
        .max_total_part_bytes(patch.after().len())
        .max_items(1)
        .build()
        .unwrap();
    assert_limit_exceeded(
        apply_with_limits(patch.before(), patch, limits),
        ExternalLinkResource::Items,
        2,
        1,
    );
}

#[test]
fn dde_relationship_id_refusal_is_atomic_and_valid_edits_remain_retryable() {
    let link = Link::dde_with_items("X", "Y", vec![DdeItem::new("A").unwrap()]).unwrap();
    let source = write_external_link_stream(&link, None).unwrap();
    let snapshot = Snapshot::read(&source).unwrap();
    let mut edit = snapshot.edit();
    let before = edit.link().clone();

    assert!(matches!(
        edit.set_relationship_id("rIdDde"),
        Err(Error::InvalidFormula(_))
    ));
    assert_eq!(edit.link(), &before);
    assert_eq!(snapshot.source_bytes(), source.as_slice());

    edit.set_source("Changed").unwrap();
    let committed = edit.commit().unwrap();
    assert_eq!(committed.snapshot().link().source(), "Changed");
    assert_eq!(committed.snapshot().link().dde_topic(), Some("Y"));
}

#[test]
fn external_link_writer_cache_matrix_record_and_object_limits_have_exact_edges() {
    let matrix = ValueMatrix::new(1, 1, vec![CachedValue::Empty]).unwrap();
    let link = Link::dde_with_items(
        "X",
        "Y",
        vec![DdeItem::new("A").unwrap().with_cached_values(matrix)],
    )
    .unwrap();
    let canonical = write_external_link_stream(&link, None).unwrap();
    let kinds = record_kinds(&canonical);
    let record_count = kinds.len();

    let exact = ExternalLinkLimits::builder()
        .max_matrices(1)
        .max_cells(1)
        .max_records(record_count)
        .build()
        .unwrap();
    assert_eq!(
        write_external_link_stream_with_limits(&link, None, exact).unwrap(),
        canonical
    );

    let matrix_below = ExternalLinkLimits::builder()
        .max_matrices(0)
        .build()
        .unwrap();
    assert_limit_exceeded(
        write_external_link_stream_with_limits(&link, None, matrix_below),
        ExternalLinkResource::Matrices,
        1,
        0,
    );

    let cells_below = ExternalLinkLimits::builder().max_cells(0).build().unwrap();
    assert_limit_exceeded(
        write_external_link_stream_with_limits(&link, None, cells_below),
        ExternalLinkResource::Cells,
        1,
        0,
    );

    let records_below = ExternalLinkLimits::builder()
        .max_records(record_count - 1)
        .build()
        .unwrap();
    assert_limit_exceeded(
        write_external_link_stream_with_limits(&link, None, records_below),
        ExternalLinkResource::Records,
        record_count,
        record_count - 1,
    );
}

#[test]
fn external_link_writer_extern_cache_record_limit_uses_opaque_table_records() {
    let link = workbook_link();
    let base = write_external_link_stream(&link, Some("rIdPath")).unwrap();
    let opaque_specs = [
        (crate::raw::kind::EXTERN_TABLE_START, &[0x00; 5][..]),
        (crate::raw::kind::EXTERN_ROW_HDR, &[0x00; 4][..]),
        (crate::raw::kind::EXTERN_CELL_BLANK, &[0x00; 4][..]),
        (
            crate::raw::kind::EXTERN_CELL_REAL,
            &[
                0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
            ][..],
        ),
        (
            crate::raw::kind::EXTERN_CELL_BOOL,
            &[0x02, 0x00, 0x00, 0x00, 0x01][..],
        ),
        (
            crate::raw::kind::EXTERN_CELL_ERROR,
            &[0x03, 0x00, 0x00, 0x00, 0x07][..],
        ),
        (
            crate::raw::kind::EXTERN_CELL_STRING,
            &[0x04, 0x00, 0x00, 0x00, 1, 0, 0, 0, b'x', 0][..],
        ),
        (crate::raw::kind::EXTERN_TABLE_END, &[][..]),
        (RecordKind::new(0x3FFD).unwrap(), &[0xD0, 0xD1][..]),
    ];
    let source = insert_unknown_records_after_known_record(&base, 2, &opaque_specs);
    let snapshot = Snapshot::read(&source).unwrap();
    assert_eq!(snapshot.unknown_records().len(), opaque_specs.len());
    let cache_records = snapshot
        .unknown_records()
        .iter()
        .take_while(|record| record.kind() != 0x3FFD)
        .count();
    assert_eq!(cache_records, 8);

    let exact = ExternalLinkLimits::builder()
        .max_cache_records(cache_records)
        .build()
        .unwrap();
    assert!(
        package::write_external_link_stream_with_unknown_and_limits(
            &link,
            Some("rIdPath"),
            snapshot.unknown_records(),
            exact,
        )
        .is_ok()
    );

    let one_below = ExternalLinkLimits::builder()
        .max_cache_records(cache_records - 1)
        .build()
        .unwrap();
    assert_limit_exceeded(
        package::write_external_link_stream_with_unknown_and_limits(
            &link,
            Some("rIdPath"),
            snapshot.unknown_records(),
            one_below,
        ),
        ExternalLinkResource::CacheRecords,
        cache_records,
        cache_records - 1,
    );
}

#[test]
fn external_link_writer_retained_object_limit_has_exact_edge() {
    let link = Link::workbook_with_defined_names(
        "unused-host-source",
        vec!["Data".to_string()],
        Vec::new(),
    )
    .unwrap();
    let exact = ExternalLinkLimits::builder()
        .max_retained_objects(1)
        .build()
        .unwrap();
    assert!(write_external_link_stream_with_limits(&link, Some("rIdPath"), exact).is_ok());

    let one_below = ExternalLinkLimits::builder()
        .max_retained_objects(0)
        .build()
        .unwrap();
    assert_limit_exceeded(
        write_external_link_stream_with_limits(&link, Some("rIdPath"), one_below),
        ExternalLinkResource::RetainedObjects,
        1,
        0,
    );
}
