//! Independent admission checks for public metadata plans.
#![allow(clippy::unwrap_used, clippy::expect_used, reason = "test assertions")]

use litchi_opc::{
    ContentTypeEdit, OpcError, OpcPackage, PackURI, ReadLimits, ReadResource, RelationshipEdit,
    Relationships, TargetMode,
};
use quick_xml::{Reader, events::Event};

fn package(relationships: Option<&[u8]>) -> OpcPackage {
    let mut writer = soapberry_zip::office::StreamingArchiveWriter::new();
    writer.write_stored("[Content_Types].xml", br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="bin" ContentType="application/octet-stream"/></Types>"#).unwrap();
    if let Some(xml) = relationships {
        writer
            .write_stored("custom/_rels/item.bin.rels", xml)
            .unwrap();
    }
    writer.write_stored("custom/item.bin", b"opaque").unwrap();
    OpcPackage::from_vec(writer.finish_to_bytes().unwrap()).unwrap()
}

#[test]
fn content_type_plan_reports_exact_events_for_default_to_override_replacement() {
    let package = package(None);
    let before = package.source_content_types().unwrap();
    let part_name = PackURI::new("/custom/item.bin").unwrap();
    let additions = [ContentTypeEdit {
        part_name: &part_name,
        content_type: "application/octet-stream",
    }];
    let plan = before
        .plan_edit_with_defaults(&additions, &[], &["bin"], ReadLimits::default())
        .unwrap();
    assert_eq!(plan.mapping_count(), 2);
    assert!(!plan.is_noop());
    let expected_events = plan.event_count();
    let expected_len = plan.final_len();
    let limits = ReadLimits::builder()
        .max_xml_events(expected_events)
        .unwrap()
        .build()
        .unwrap();
    let after = plan.materialize(limits).unwrap();
    assert_eq!(after.bytes().len(), expected_len);
    assert_eq!(events(after.bytes()), expected_events);
    let text = std::str::from_utf8(after.bytes()).unwrap();
    assert!(text.contains("PartName=\"/custom/item.bin\""));
    assert!(!text.contains("Extension=\"bin\""));
    assert!(
        before
            .plan_edit_with_defaults(
                &additions,
                &[],
                &["bin"],
                ReadLimits::builder()
                    .max_xml_events(expected_events - 1)
                    .unwrap()
                    .build()
                    .unwrap()
            )
            .is_err()
    );
}

fn events(xml: &[u8]) -> usize {
    let mut reader = Reader::from_reader(xml);
    let mut count = 0;
    loop {
        count += 1;
        if matches!(reader.read_event().unwrap(), Event::Eof) {
            return count;
        }
    }
}

#[test]
fn relationship_plan_admits_exact_events_before_materialization_for_all_root_forms() {
    let empty = br#"<r:Relationships xmlns:r='http://schemas.openxmlformats.org/package/2006/relationships'/>"#;
    let paired = br#"<r:Relationships xmlns:r='http://schemas.openxmlformats.org/package/2006/relationships'><!-- preserved --></r:Relationships>"#;
    for source in [None, Some(empty.as_slice()), Some(paired.as_slice())] {
        let package = package(source);
        let before = package
            .source_relationships(&PackURI::new("/custom/item.bin").unwrap())
            .unwrap();
        let additions = [RelationshipEdit {
            id: "rId1",
            reltype: "urn:test",
            target: "https://example.test/a?x=1&y=2",
            mode: TargetMode::External,
        }];
        let plan = before
            .plan_edit(&additions, &[], ReadLimits::default())
            .unwrap();
        let planned_events = plan.event_count();
        let planned_bytes = plan.final_len();
        assert_eq!(plan.relationship_count(), 1);
        assert!(plan.member_present());
        let limits = ReadLimits::builder()
            .max_xml_events(planned_events)
            .unwrap()
            .build()
            .unwrap();
        let after = before
            .plan_edit(&additions, &[], limits)
            .unwrap()
            .materialize(limits)
            .unwrap();
        assert_eq!(events(after.bytes()), planned_events);
        assert_eq!(after.bytes().len(), planned_bytes);
        assert_eq!(before.member_present(), source.is_some());
        if let Some(raw) = source {
            assert_eq!(before.bytes(), raw);
        }
        if source == Some(paired.as_slice()) {
            assert!(
                std::str::from_utf8(after.bytes())
                    .unwrap()
                    .contains("<!-- preserved -->")
            );
        }
        let under = ReadLimits::builder()
            .max_xml_events(planned_events - 1)
            .unwrap()
            .build()
            .unwrap();
        assert!(matches!(
            before.plan_edit(&additions, &[], under),
            Err(OpcError::ReadLimit {
                resource: ReadResource::XmlEvents,
                ..
            })
        ));
    }
}

#[test]
fn mixed_relationship_plan_counts_removed_paired_tokens_and_preserves_unrelated_bytes() {
    let source = br#"<?xml version='1.0'?><r:Relationships xmlns:r='http://schemas.openxmlformats.org/package/2006/relationships'><!-- keep --><r:Relationship Id='rId1' Type='urn:test' Target='https://example.test/old' TargetMode='External'></r:Relationship><?keep data?><r:Relationship Id='rId2' Type='urn:other' Target='https://example.test/untouched' TargetMode='External'/></r:Relationships>"#;
    let package = package(Some(source));
    let before = package
        .source_relationships(&PackURI::new("/custom/item.bin").unwrap())
        .unwrap();
    let additions = [RelationshipEdit {
        id: "rId1",
        reltype: "urn:test",
        target: "https://example.test/new",
        mode: TargetMode::External,
    }];
    let plan = before
        .plan_edit(&additions, &["rId1"], ReadLimits::default())
        .unwrap();
    let final_len = plan.final_len();
    let final_events = plan.event_count();
    assert_eq!(plan.relationship_count(), 2);
    let after = plan.materialize(ReadLimits::default()).unwrap();
    assert_eq!(after.bytes().len(), final_len);
    assert_eq!(events(after.bytes()), final_events);
    assert_eq!(final_events + 1, events(source));
    let text = std::str::from_utf8(after.bytes()).unwrap();
    assert!(text.contains("<!-- keep -->"));
    assert!(text.contains("<?keep data?>"));
    assert!(text.contains("<r:Relationship Id='rId2' Type='urn:other' Target='https://example.test/untouched' TargetMode='External'/>"));
    assert!(!text.contains("https://example.test/old"));
    assert!(text.contains("https://example.test/new"));
    assert_eq!(before.bytes(), source);
}

#[test]
fn relationship_plan_enforces_archive_entry_cap_for_noop_and_changed_output() {
    let xml = br#"<r:Relationships xmlns:r='http://schemas.openxmlformats.org/package/2006/relationships'/>"#;
    let package = package(Some(xml));
    let before = package
        .source_relationships(&PackURI::new("/custom/item.bin").unwrap())
        .unwrap();
    let addition = [RelationshipEdit {
        id: "rId1",
        reltype: "urn:test",
        target: "https://example.test/new",
        mode: TargetMode::External,
    }];
    for additions in [&[][..], addition.as_slice()] {
        let size = before
            .plan_edit(additions, &[], ReadLimits::default())
            .unwrap()
            .final_len();
        let exact = ReadLimits::builder()
            .max_archive_entry_bytes(size as u64)
            .unwrap()
            .build()
            .unwrap();
        let under = ReadLimits::builder()
            .max_archive_entry_bytes((size - 1) as u64)
            .unwrap()
            .build()
            .unwrap();
        let after = before
            .plan_edit(additions, &[], exact)
            .unwrap()
            .materialize(exact)
            .unwrap();
        assert_eq!(after.bytes().len(), size);
        assert!(
            matches!(before.plan_edit(additions, &[], under), Err(OpcError::ReadLimit { resource: ReadResource::ArchiveEntryBytes, actual, maximum }) if actual == size as u64 && maximum == (size - 1) as u64)
        );
        assert!(
            matches!(before.plan_edit(additions, &[], ReadLimits::default()).unwrap().materialize(under), Err(OpcError::ReadLimit { resource: ReadResource::ArchiveEntryBytes, actual, maximum }) if actual == size as u64 && maximum == (size - 1) as u64)
        );
        assert_eq!(before.bytes(), xml);
    }
}

#[test]
fn content_type_override_conflicts_are_refused_during_planning() {
    let package = package(None);
    let initial = package.source_content_types().unwrap();
    let part = PackURI::new("/custom/item.bin").unwrap();
    let additions = [ContentTypeEdit {
        part_name: &part,
        content_type: "application/octet-stream",
    }];
    let before = initial
        .plan_edit(&additions, &[], ReadLimits::default())
        .unwrap()
        .materialize(ReadLimits::default())
        .unwrap();
    assert!(matches!(
        before.plan_edit(&additions, &[], ReadLimits::default()),
        Err(OpcError::InvalidContentTypesManifest(_))
    ));
    let alias = PackURI::new("/CUSTOM/ITEM.BIN").unwrap();
    let replacement = [ContentTypeEdit {
        part_name: &alias,
        content_type: "application/x-test",
    }];
    let after = before
        .plan_edit(&replacement, &[part], ReadLimits::default())
        .unwrap()
        .materialize(ReadLimits::default())
        .unwrap();
    assert!(
        std::str::from_utf8(after.bytes())
            .unwrap()
            .contains("application/x-test")
    );
}

#[test]
fn content_type_default_selectors_use_only_xml_token_whitespace() {
    let package = package(None);
    let before = package.source_content_types().unwrap();
    let limits = ReadLimits::default();
    let plain = before
        .plan_edit_with_defaults(&[], &[], &["bin"], limits)
        .unwrap()
        .materialize(limits)
        .unwrap();
    let spaced = before
        .plan_edit_with_defaults(&[], &[], &[" \tBIN\r\n "], limits)
        .unwrap()
        .materialize(limits)
        .unwrap();
    assert_eq!(plain.bytes(), spaced.bytes());
    for selectors in [vec!["bin", " BIN\t"], vec!["b in"], vec!["\u{a0}bin\u{a0}"]] {
        assert!(
            before
                .plan_edit_with_defaults(&[], &[], &selectors, limits)
                .is_err(),
            "unexpected selector admission: {selectors:?}"
        );
    }
}

#[test]
fn new_owner_canonical_plan_reports_actual_output_before_materialization() {
    for populated in [false, true] {
        let mut relationships = Relationships::new("/new/item.xml".to_owned());
        if populated {
            relationships
                .try_add_relationship(
                    "urn:test".to_owned(),
                    "https://example.test/é?a=1&b=2".to_owned(),
                    "rId1".to_owned(),
                    TargetMode::External,
                )
                .unwrap();
        }
        let plan = relationships.plan_canonical(ReadLimits::default()).unwrap();
        let size = plan.final_len();
        let count = plan.event_count();
        assert_eq!(plan.relationship_count(), usize::from(populated));
        let exact = ReadLimits::builder()
            .max_archive_entry_bytes(size as u64)
            .unwrap()
            .max_xml_events(count)
            .unwrap()
            .build()
            .unwrap();
        let xml = plan.materialize(exact).unwrap();
        assert_eq!(xml.len(), size);
        assert_eq!(events(&xml), count);
        let under = ReadLimits::builder()
            .max_archive_entry_bytes((size - 1) as u64)
            .unwrap()
            .build()
            .unwrap();
        assert!(matches!(
            relationships.plan_canonical(under),
            Err(OpcError::ReadLimit {
                resource: ReadResource::ArchiveEntryBytes,
                ..
            })
        ));
        let under = ReadLimits::builder()
            .max_xml_events(count - 1)
            .unwrap()
            .build()
            .unwrap();
        assert!(matches!(
            relationships.plan_canonical(under),
            Err(OpcError::ReadLimit {
                resource: ReadResource::XmlEvents,
                ..
            })
        ));
    }
}
