use super::super::support::*;

#[test]
fn smart_tags_round_trip_through_both_output_paths() {
    let mut writer = Writer::new();
    writer.add_paragraph("abcdefghijklmnopqrst").unwrap();
    writer.add_smart_tag(
        SmartTagEntry::new(0, 10, "urn:example:geo", "place")
            .with_origin(crate::SmartTagOrigin::ExternalRecognizer)
            .with_native_export(true)
            .with_property("city", "東京"),
    );
    writer.add_smart_tag(
        SmartTagEntry::new(5, 15, "urn:example:geo", "place")
            .with_sub_entity(true)
            .with_property("city", "Paris"),
    );
    writer.add_smart_tag(SmartTagEntry::new(5, 5, "urn:example:point", "cursor"));
    writer.add_smart_tag_recognizer_range(SmartTagRecognizerRange {
        start: 0,
        end: 5,
        state: crate::SmartTagRecognizerState::Dirty,
    });
    writer.add_smart_tag_recognizer_range(SmartTagRecognizerRange {
        start: 5,
        end: 20,
        state: crate::SmartTagRecognizerState::Clean,
    });

    let mut cursor = Cursor::new(Vec::new());
    writer.write_to(&mut cursor).unwrap();
    let mut package = crate::Package::from_reader(Cursor::new(cursor.into_inner())).unwrap();
    let document = package.document().unwrap();
    for index in [114usize, 115, 117, 118, 132] {
        assert!(document.fib().get_table_pointer(index).unwrap().1 > 0);
    }
    let smart_tags = document.smart_tags().unwrap().clone();
    assert_eq!(smart_tags.tags.len(), 3);
    assert_eq!(smart_tags.store.as_ref().unwrap().types.len(), 2);
    assert_eq!(
        smart_tags.tags[0].info.origin,
        crate::SmartTagOrigin::ExternalRecognizer
    );
    assert!(smart_tags.tags[0].is_native);
    assert_eq!(
        smart_tags
            .store
            .as_ref()
            .unwrap()
            .resolve_property(smart_tags.tags[0].property_bag.properties[0]),
        Some(("city", "東京"))
    );
    assert_eq!(
        (smart_tags.tags[1].start_depth, smart_tags.tags[1].end_depth),
        (3, 0)
    );
    assert_eq!(smart_tags.recognizer_ranges.len(), 2);

    let path = std::env::temp_dir().join(format!(
        "litchi-doc-smart-tags-{}-{}.doc",
        std::process::id(),
        std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .unwrap()
            .as_nanos()
    ));
    writer.save(&path).unwrap();
    let mut package = crate::Package::open(&path).unwrap();
    assert_eq!(package.document().unwrap().smart_tags(), Some(&smart_tags));
    std::fs::remove_file(path).unwrap();
}
