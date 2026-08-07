use super::*;
use crate::property_set::{CodePage, DOCUMENT_SUMMARY_INFORMATION_FMTID, Section, Value};

const LINK_BASE_FIXTURE: [u8; 4] = [0x62, 0x00, 0x00, 0x00];
const HYPERLINKS_FIXTURE: [u8; 64] = [
    0x3c, 0x00, 0x00, 0x00, 0x06, 0x00, 0x00, 0x00, // cbData, cElements
    0x03, 0x00, 0x00, 0x00, 0x61, 0x00, 0x00, 0x00, // hash: "A" ^ ""
    0x03, 0x00, 0x00, 0x00, 0x07, 0x00, 0x00, 0x00, // app
    0x03, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, // OfficeArt
    0x03, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, // info
    0x1f, 0x00, 0x00, 0x00, 0x02, 0x00, 0x00, 0x00, // target tag, cch
    0x41, 0x00, 0x00, 0x00, // target "A" and terminator
    0x1f, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, // location tag, cch
    0x00, 0x00, 0x00, 0x00, // location terminator and required alignment
];

fn section() -> Section {
    let mut section = Section::new(crate::property_set::USER_DEFINED_PROPERTIES_FMTID);
    section.set_page(CodePage::Utf16Le);
    section
}

fn hyperlink() -> Hyperlink {
    Hyperlink::new(7, 0, 0, "https://example.test/a", "bookmark").unwrap()
}

#[test]
fn compact_fixture_round_trips_and_preserves_the_stored_hash() {
    let mut source = section();
    let mut edit = Edit::new(&mut source).unwrap();
    edit.set_link_base(LinkBase::new("https://example.test/base/").unwrap())
        .unwrap();
    edit.set_hyperlinks(Hyperlinks::new(vec![hyperlink()]))
        .unwrap();

    let typed = Properties::new(&source).unwrap();
    assert_eq!(
        typed.link_base().unwrap().unwrap().value(),
        "https://example.test/base/"
    );
    let links = typed.hyperlinks().unwrap().unwrap();
    let link = &links.links()[0];
    assert!(link.hash_matches());
    assert_eq!(link.app(), 7);
    assert_eq!(link.target(), "https://example.test/a");
    assert_eq!(link.location(), "bookmark");

    let raw = source.find_named(HYPERLINKS).unwrap().1;
    assert!(matches!(raw, Value::Blob(data) if data.len() >= 8));
}

#[test]
fn independent_compact_fixtures_have_exact_tags_hash_and_alignment() {
    let mut source = section();
    source
        .add_named(2, LINK_BASE.into(), Value::Blob(LINK_BASE_FIXTURE.to_vec()))
        .unwrap();
    source
        .add_named(
            3,
            HYPERLINKS.into(),
            Value::Blob(HYPERLINKS_FIXTURE.to_vec()),
        )
        .unwrap();
    let typed = Properties::new(&source).unwrap();
    assert_eq!(typed.link_base().unwrap().unwrap().value(), "b");
    let links = typed.hyperlinks().unwrap().unwrap();
    assert_eq!(links.links()[0].stored_hash(), 0x61);
    assert!(links.links()[0].hash_matches());
    assert_eq!(links.links()[0].app(), 7);
    assert_eq!(links.links()[0].target(), "A");
    assert_eq!(links.links()[0].location(), "");

    let mut edit = Edit::new(&mut source).unwrap();
    edit.set_link_base(LinkBase::new("b").unwrap()).unwrap();
    edit.set_hyperlinks(links).unwrap();
    assert_eq!(
        edit.section().find_named(LINK_BASE).unwrap().1,
        &Value::Blob(LINK_BASE_FIXTURE.to_vec())
    );
    assert_eq!(
        edit.section().find_named(HYPERLINKS).unwrap().1,
        &Value::Blob(HYPERLINKS_FIXTURE.to_vec())
    );

    let mut padded = HYPERLINKS_FIXTURE;
    padded[42] = 0x34;
    padded[43] = 0x12;
    let mut padded_source = section();
    padded_source
        .add_named(2, HYPERLINKS.into(), Value::Blob(padded.to_vec()))
        .unwrap();
    let padded_links = Properties::new(&padded_source)
        .unwrap()
        .hyperlinks()
        .unwrap()
        .unwrap();
    assert_eq!(padded_links.links()[0].target(), "A");
    assert_eq!(padded_links.links()[0].location(), "");
    Edit::new(&mut padded_source)
        .unwrap()
        .set_hyperlinks(padded_links)
        .unwrap();
    assert_eq!(
        padded_source.find_named(HYPERLINKS).unwrap().1,
        &Value::Blob(HYPERLINKS_FIXTURE.to_vec())
    );
}

#[test]
fn property_stream_outer_hyperlink_padding_is_ignored_before_typed_overlay() {
    let mut source = section();
    source
        .add_named(
            2,
            HYPERLINKS.into(),
            Value::Blob(HYPERLINKS_FIXTURE.to_vec()),
        )
        .unwrap();
    let mut bytes = crate::property_set::Stream::new(source).to_bytes().unwrap();
    let marker = [
        0x41, 0x00, 0x00, 0x00, 0x40, 0x00, 0x00, 0x00, 0x3c, 0x00, 0x00, 0x00,
    ];
    let offset = bytes
        .windows(marker.len())
        .position(|window| window == marker)
        .expect("fixed VtHyperlinks outer value fixture");
    bytes[offset + 2] = 0xa5;
    bytes[offset + 3] = 0x5a;
    let parsed = crate::property_set::Stream::parse(&bytes).unwrap();
    let section = parsed
        .section(crate::property_set::USER_DEFINED_PROPERTIES_FMTID)
        .unwrap();
    let links = Properties::new(section)
        .unwrap()
        .hyperlinks()
        .unwrap()
        .unwrap();
    assert_eq!(links.links()[0].target(), "A");
    assert_eq!(links.links()[0].stored_hash(), 0x61);
}

#[test]
fn reserved_names_are_case_insensitive_and_never_numeric_piddsi() {
    let mut source = section();
    source
        .add_named(42, "_pid_linkbase".into(), Value::Blob(vec![0, 0]))
        .unwrap();
    source.add(0x14, Value::Blob(vec![0, 0])).unwrap();
    let typed = Properties::new(&source).unwrap();
    assert_eq!(typed.link_base().unwrap().unwrap().value(), "");

    source.update(42, Value::Blob(Vec::new())).unwrap();
    assert_eq!(
        Properties::new(&source)
            .unwrap()
            .link_base()
            .unwrap()
            .unwrap()
            .value(),
        ""
    );

    let mut edit = Edit::new(&mut source).unwrap();
    edit.set_link_base(LinkBase::new("base").unwrap()).unwrap();
    assert!(source.find_named("_PID_LINKBASE").is_some());
    assert!(matches!(source.property(0x14), Some(Value::Blob(_))));
    assert_eq!(source.property_name(42), Some("_pid_linkbase"));
}

#[test]
fn malformed_values_are_lazy_and_bounded() {
    let mut source = section();
    source
        .add_named(2, HYPERLINKS.into(), Value::Blob(vec![4, 0, 0, 0]))
        .unwrap();
    let typed = Properties::new(&source).unwrap();
    assert!(typed.hyperlinks().is_err());

    let mut source = section();
    source
        .add_named(2, LINK_BASE.into(), Value::Blob(vec![0; 8]))
        .unwrap();
    let limits = Limits::builder().max_blob_bytes(4).build().unwrap();
    assert!(
        Properties::with_limits(&source, limits)
            .unwrap()
            .link_base()
            .is_err()
    );
}

#[test]
fn hyperlink_grammar_rejects_wrong_elements_but_accepts_retained_padding() {
    let mut source = section();
    let mut malformed = vec![4, 0, 0, 0, 1, 0, 0, 0];
    source
        .add_named(2, HYPERLINKS.into(), Value::Blob(malformed.clone()))
        .unwrap();
    assert!(Properties::new(&source).unwrap().hyperlinks().is_err());

    malformed[4] = 6;
    let mut source = section();
    source
        .add_named(2, HYPERLINKS.into(), Value::Blob(malformed))
        .unwrap();
    assert!(Properties::new(&source).unwrap().hyperlinks().is_err());

    let mut source = section();
    let mut edit = Edit::new(&mut source).unwrap();
    edit.set_hyperlinks(Hyperlinks::new(vec![hyperlink()]))
        .unwrap();
    let Value::Blob(raw) = source.find_named(HYPERLINKS).unwrap().1 else {
        unreachable!()
    };
    let mut raw = raw.clone();
    raw[10] = 1;
    source.update(2, Value::Blob(raw)).unwrap();
    let links = Properties::new(&source)
        .unwrap()
        .hyperlinks()
        .unwrap()
        .unwrap();
    assert_eq!(links.links()[0], hyperlink());

    let Value::Blob(raw) = source.find_named(HYPERLINKS).unwrap().1 else {
        unreachable!()
    };
    let mut raw = raw.clone();
    raw[42] = 0x7f;
    source.update(2, Value::Blob(raw)).unwrap();
    let links = Properties::new(&source)
        .unwrap()
        .hyperlinks()
        .unwrap()
        .unwrap();
    assert_eq!(links.links()[0], hyperlink());
}

#[test]
fn fixtures_enforce_exact_bounds_and_malformed_inner_fields() {
    let mut source = section();
    source
        .add_named(2, LINK_BASE.into(), Value::Blob(LINK_BASE_FIXTURE.to_vec()))
        .unwrap();
    assert!(
        Properties::with_limits(
            &source,
            Limits::builder()
                .max_blob_bytes(LINK_BASE_FIXTURE.len())
                .build()
                .unwrap()
        )
        .unwrap()
        .link_base()
        .is_ok()
    );
    assert!(
        Properties::with_limits(
            &source,
            Limits::builder()
                .max_blob_bytes(LINK_BASE_FIXTURE.len() - 1)
                .build()
                .unwrap()
        )
        .unwrap()
        .link_base()
        .is_err()
    );

    let mut source = section();
    source
        .add_named(
            2,
            HYPERLINKS.into(),
            Value::Blob(HYPERLINKS_FIXTURE.to_vec()),
        )
        .unwrap();
    assert!(
        Properties::with_limits(
            &source,
            Limits::builder()
                .max_blob_bytes(HYPERLINKS_FIXTURE.len())
                .build()
                .unwrap()
        )
        .unwrap()
        .hyperlinks()
        .is_ok()
    );
    assert!(
        Properties::with_limits(
            &source,
            Limits::builder()
                .max_blob_bytes(HYPERLINKS_FIXTURE.len() - 1)
                .build()
                .unwrap()
        )
        .unwrap()
        .hyperlinks()
        .is_err()
    );

    for (offset, value) in [(0, 0x3b), (4, 5), (40, 0x1e), (48, 0)] {
        let mut malformed = HYPERLINKS_FIXTURE;
        malformed[offset] = value;
        let mut source = section();
        source
            .add_named(2, HYPERLINKS.into(), Value::Blob(malformed.to_vec()))
            .unwrap();
        assert!(Properties::new(&source).unwrap().hyperlinks().is_err());
    }
    let mut malformed = HYPERLINKS_FIXTURE;
    malformed[48] = 0;
    malformed[49] = 0xd8;
    let mut source = section();
    source
        .add_named(2, HYPERLINKS.into(), Value::Blob(malformed.to_vec()))
        .unwrap();
    assert!(Properties::new(&source).unwrap().hyperlinks().is_err());
}

#[test]
fn limits_cover_link_count_string_and_aggregate_units() {
    let values = Hyperlinks::new(vec![
        Hyperlink::new(0, 0, 0, "a", "b").unwrap(),
        Hyperlink::new(0, 0, 0, "c", "d").unwrap(),
    ]);
    let mut source = section();
    assert!(
        Edit::with_limits(&mut source, Limits::builder().max_links(1).build().unwrap())
            .unwrap()
            .set_hyperlinks(values.clone())
            .is_err()
    );
    assert!(source.find_named(HYPERLINKS).is_none());

    let mut source = section();
    assert!(
        Edit::with_limits(
            &mut source,
            Limits::builder().max_string_units(1).build().unwrap()
        )
        .unwrap()
        .set_hyperlinks(values.clone())
        .is_err()
    );

    let mut source = section();
    assert!(
        Edit::with_limits(
            &mut source,
            Limits::builder().max_total_utf16_units(3).build().unwrap()
        )
        .unwrap()
        .set_hyperlinks(values)
        .is_err()
    );
}

#[test]
fn decode_limits_accept_exact_boundaries_and_reject_one_over() {
    let mut source = section();
    source
        .add_named(
            2,
            HYPERLINKS.into(),
            Value::Blob(HYPERLINKS_FIXTURE.to_vec()),
        )
        .unwrap();
    assert!(
        Properties::with_limits(
            &source,
            Limits::builder()
                .max_links(1)
                .max_string_units(2)
                .max_total_utf16_units(3)
                .build()
                .unwrap(),
        )
        .unwrap()
        .hyperlinks()
        .is_ok()
    );
    assert!(
        Properties::with_limits(
            &source,
            Limits::builder().max_string_units(1).build().unwrap(),
        )
        .unwrap()
        .hyperlinks()
        .is_err()
    );
    assert!(
        Properties::with_limits(
            &source,
            Limits::builder().max_total_utf16_units(2).build().unwrap(),
        )
        .unwrap()
        .hyperlinks()
        .is_err()
    );

    let mut two_links = section();
    Edit::new(&mut two_links)
        .unwrap()
        .set_hyperlinks(Hyperlinks::new(vec![
            Hyperlink::new(0, 0, 0, "a", "").unwrap(),
            Hyperlink::new(0, 0, 0, "b", "").unwrap(),
        ]))
        .unwrap();
    assert!(
        Properties::with_limits(&two_links, Limits::builder().max_links(1).build().unwrap(),)
            .unwrap()
            .hyperlinks()
            .is_err()
    );
}

#[test]
fn edits_preserve_existing_identity_order_unknown_values_and_remove_cleanly() {
    let mut source = section();
    source
        .add_named(
            18,
            "_pid_hlinks".into(),
            Value::Blob(vec![4, 0, 0, 0, 0, 0, 0, 0]),
        )
        .unwrap();
    source
        .add_named(
            19,
            "Other".into(),
            Value::Unknown {
                variant_type: 0x7777,
                data: vec![1, 2],
            },
        )
        .unwrap();
    let original_order: Vec<_> = source.property_ids().collect();
    let other = source.property(19).cloned();

    let mut edit = Edit::new(&mut source).unwrap();
    edit.set_hyperlinks(Hyperlinks::new(vec![hyperlink()]))
        .unwrap();
    assert_eq!(edit.section().property_name(18), Some("_pid_hlinks"));
    assert_eq!(edit.section().property(19), other.as_ref());
    assert_eq!(
        edit.section().property_ids().collect::<Vec<_>>(),
        original_order
    );
    assert!(edit.remove_hyperlinks());
    assert!(source.find_named(HYPERLINKS).is_none());
    assert_eq!(source.property(19), other.as_ref());
}

#[test]
fn insertion_establishes_required_scaffolding_and_edits_rollback() {
    let mut blank = Section::new(crate::property_set::USER_DEFINED_PROPERTIES_FMTID);
    Edit::new(&mut blank)
        .unwrap()
        .set_link_base(LinkBase::new("base").unwrap())
        .unwrap();
    assert_eq!(blank.page(), Some(CodePage::WINDOWS_1252));
    assert!(blank.find_named(LINK_BASE).is_some());

    let mut unicode = Section::new(crate::property_set::USER_DEFINED_PROPERTIES_FMTID);
    unicode.set_page(CodePage::Utf16Le);
    Edit::new(&mut unicode)
        .unwrap()
        .set_link_base(LinkBase::new("base").unwrap())
        .unwrap();
    assert_eq!(unicode.page(), Some(CodePage::Utf16Le));

    let mut existing = Section::new(crate::property_set::USER_DEFINED_PROPERTIES_FMTID);
    existing
        .add_named(2, LINK_BASE.into(), Value::Blob(LINK_BASE_FIXTURE.to_vec()))
        .unwrap();
    Edit::new(&mut existing)
        .unwrap()
        .set_link_base(LinkBase::new("replacement").unwrap())
        .unwrap();
    assert_eq!(existing.page(), Some(CodePage::WINDOWS_1252));
    assert_eq!(existing.property_name(2), Some(LINK_BASE));

    let mut high = section();
    high.add_named(
        0x01_000000,
        LINK_BASE.into(),
        Value::Blob(LINK_BASE_FIXTURE.to_vec()),
    )
    .unwrap();
    assert!(Edit::new(&mut high).is_err());

    let mut at_max = section();
    at_max
        .add_named(
            0x00ff_ffff,
            LINK_BASE.into(),
            Value::Blob(LINK_BASE_FIXTURE.to_vec()),
        )
        .unwrap();
    Edit::new(&mut at_max)
        .unwrap()
        .set_link_base(LinkBase::new("accepted").unwrap())
        .unwrap();
    assert_eq!(at_max.property_name(0x00ff_ffff), Some(LINK_BASE));

    let mut linked = section();
    linked
        .add_named(30, "LinkedValue".into(), Value::Lpstr("base".into()))
        .unwrap();
    linked
        .add(0x01_00001e, Value::Lpstr("companion".into()))
        .unwrap();
    let linked_order: Vec<_> = linked.property_ids().collect();
    let linked_base = linked.property(30).cloned();
    let linked_companion = linked.property(0x01_00001e).cloned();
    let mut linked_edit = Edit::new(&mut linked).unwrap();
    linked_edit
        .set_link_base(LinkBase::new("reserved").unwrap())
        .unwrap();
    assert_eq!(linked_edit.section().property(30), linked_base.as_ref());
    assert_eq!(
        linked_edit.section().property(0x01_00001e),
        linked_companion.as_ref()
    );
    assert_eq!(
        linked_edit.section().property_ids().collect::<Vec<_>>(),
        [linked_order.as_slice(), &[2]].concat()
    );
    assert!(linked_edit.remove_link_base());
    assert_eq!(
        linked_edit.section().property_ids().collect::<Vec<_>>(),
        linked_order
    );
    assert_eq!(linked_edit.section().property(30), linked_base.as_ref());
    assert_eq!(
        linked_edit.section().property(0x01_00001e),
        linked_companion.as_ref()
    );

    let mut populated = section();
    let mut edit = Edit::new(&mut populated).unwrap();
    edit.set_link_base(LinkBase::new("base").unwrap()).unwrap();
    edit.set_hyperlinks(Hyperlinks::new(vec![hyperlink()]))
        .unwrap();
    let original = edit.section().clone();
    drop(edit);
    let mut limited = Edit::with_limits(
        &mut populated,
        Limits::builder().max_blob_bytes(3).build().unwrap(),
    )
    .unwrap();
    assert!(
        limited
            .set_link_base(LinkBase::new("replacement").unwrap())
            .is_err()
    );
    assert_eq!(limited.section(), &original);
    assert!(
        limited
            .set_hyperlinks(Hyperlinks::new(vec![hyperlink()]))
            .is_err()
    );
    assert_eq!(limited.section(), &original);
}

#[test]
fn hash_is_case_insensitive_and_constructor_uses_the_should_algorithm() {
    let one = Hyperlink::new(0, 0, 0, "AbC", "Target").unwrap();
    let two = Hyperlink::new(0, 0, 0, "aBc", "target").unwrap();
    assert_eq!(one.stored_hash(), two.stored_hash());
    assert!(one.hash_matches());

    let parsed = Hyperlink::from_wire(
        one.stored_hash().wrapping_add(1),
        0,
        0,
        0,
        "AbC".into(),
        "Target".into(),
    );
    assert!(!parsed.hash_matches());

    let truncated = "a".repeat(255);
    let longer = format!("{truncated}z");
    assert_eq!(
        Hyperlink::new(0, 0, 0, truncated, "")
            .unwrap()
            .stored_hash(),
        Hyperlink::new(0, 0, 0, longer, "").unwrap().stored_hash()
    );

    let boundary = format!("{}\u{1f600}", "a".repeat(254));
    assert_eq!(
        Hyperlink::new(0, 0, 0, boundary, "").unwrap().stored_hash(),
        0x0061_d85c
    );
}

#[test]
fn wrong_section_is_rejected_without_interpreting_piddsi() {
    let mut wrong = Section::new(DOCUMENT_SUMMARY_INFORMATION_FMTID);
    wrong.set_page(CodePage::Utf16Le);
    assert!(Properties::new(&wrong).is_err());
    assert!(Edit::new(&mut wrong).is_err());
}
