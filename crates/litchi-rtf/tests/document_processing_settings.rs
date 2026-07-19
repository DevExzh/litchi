use litchi_rtf::{
    AbstractNumberingCleanupStatus, DocumentEventMask, DocumentProcessingSettings, RtfDocument,
    RtfWriter,
};

fn write(document: &RtfDocument<'_>) -> Vec<u8> {
    let mut output = Vec::new();
    RtfWriter::new(&mut output)
        .write_document(document)
        .unwrap();
    output
}

#[test]
fn parses_explicit_values_and_preserves_zero_separately_from_omission() {
    let explicit =
        RtfDocument::parse(r#"{\rtf1\fracwidth\ilfomacatclnup0\grfdocevents0 Body}"#).unwrap();
    assert!(
        explicit
            .processing_settings()
            .fractional_character_widths_for_printing
    );
    assert_eq!(
        explicit.processing_settings().abstract_numbering_cleanup,
        Some(AbstractNumberingCleanupStatus::Reviewed)
    );
    assert_eq!(
        explicit.processing_settings().event_mask,
        DocumentEventMask::from_bits(0)
    );
    let serialized = String::from_utf8(write(&explicit)).unwrap();
    assert!(serialized.contains("\\ilfomacatclnup0"));
    assert!(serialized.contains("\\grfdocevents0"));

    let omitted = RtfDocument::parse(r#"{\rtf1 Body}"#).unwrap();
    assert!(omitted.processing_settings().is_empty());
    assert_eq!(
        omitted
            .processing_settings()
            .effective_abstract_numbering_cleanup(),
        AbstractNumberingCleanupStatus::Reviewed
    );
    let serialized = String::from_utf8(write(&omitted)).unwrap();
    assert!(!serialized.contains("\\ilfomacatclnup"));
    assert!(!serialized.contains("\\grfdocevents"));
}

#[test]
fn every_document_event_bit_including_reserved_bits_round_trips() {
    let flags = [
        DocumentEventMask::NEW,
        DocumentEventMask::OPEN,
        DocumentEventMask::CLOSE,
        DocumentEventMask::SYNC,
        DocumentEventMask::XML_AFTER_INSERT,
        DocumentEventMask::XML_BEFORE_DELETE,
        DocumentEventMask::RESERVED_INTERNAL_6,
        DocumentEventMask::RESERVED_INTERNAL_7,
        DocumentEventMask::CONTENT_CONTROL_AFTER_ADD,
        DocumentEventMask::CONTENT_CONTROL_BEFORE_DELETE,
        DocumentEventMask::CONTENT_CONTROL_ON_EXIT,
        DocumentEventMask::CONTENT_CONTROL_ON_ENTER,
        DocumentEventMask::CONTENT_CONTROL_BEFORE_STORE_UPDATE,
        DocumentEventMask::CONTENT_CONTROL_BEFORE_CONTENT_UPDATE,
        DocumentEventMask::BUILDING_BLOCK_INSERT,
    ];

    for (bit, flag) in flags.into_iter().enumerate() {
        assert_eq!(flag.bits(), 1 << bit);
        assert!(DocumentEventMask::ALL.contains(flag));
        let source = format!(r#"{{\rtf1\grfdocevents{} Body}}"#, flag.bits());
        let document = RtfDocument::parse(&source).unwrap();
        assert_eq!(document.processing_settings().event_mask, Some(flag));
        let output = write(&document);
        assert_eq!(
            RtfDocument::parse_bytes(&output)
                .unwrap()
                .processing_settings()
                .event_mask,
            Some(flag)
        );
    }
    assert!(DocumentEventMask::from_bits(0x8000).is_none());
}

#[test]
fn typed_api_round_trips_all_bits_in_stable_order_without_side_effects() {
    let mut document = RtfDocument::parse(r#"{\rtf1 Body}"#).unwrap();
    document.set_processing_settings(DocumentProcessingSettings {
        fractional_character_widths_for_printing: true,
        abstract_numbering_cleanup: Some(AbstractNumberingCleanupStatus::Incomplete),
        event_mask: Some(DocumentEventMask::ALL),
    });
    let output = write(&document);
    let serialized = String::from_utf8(output.clone()).unwrap();
    assert!(
        serialized.find("\\fracwidth").unwrap() < serialized.find("\\ilfomacatclnup1").unwrap()
    );
    assert!(
        serialized.find("\\ilfomacatclnup1").unwrap()
            < serialized.find("\\grfdocevents32767").unwrap()
    );
    let reparsed = RtfDocument::parse_bytes(&output).unwrap();
    assert_eq!(
        reparsed.processing_settings(),
        document.processing_settings()
    );
    assert_eq!(reparsed.text(), "Body");

    document.clear_processing_settings();
    assert!(document.processing_settings().is_empty());
    assert_eq!(document.text(), "Body");
}

#[test]
fn coexists_with_adjacent_output_and_rendering_properties() {
    let document = RtfDocument::parse(concat!(
        r#"{\rtf1\grfdocevents16387\jexpand\fracwidth\psover"#,
        r#"\ilfomacatclnup1\horzdoc Body}"#,
    ))
    .unwrap();
    assert_eq!(
        document.processing_settings().event_mask,
        DocumentEventMask::from_bits(16387)
    );
    let output = write(&document);
    let reparsed = RtfDocument::parse_bytes(&output).unwrap();
    assert_eq!(
        reparsed.processing_settings(),
        document.processing_settings()
    );
    assert_eq!(reparsed.rendering_settings(), document.rendering_settings());
    assert_eq!(reparsed.output_settings(), document.output_settings());
    assert_eq!(reparsed.text(), "Body");
}

#[test]
fn rejects_missing_invalid_overflow_duplicate_starred_grouped_and_late_forms() {
    for source in [
        r#"{\rtf1\fracwidth0 Body}"#,
        r#"{\rtf1\ilfomacatclnup Body}"#,
        r#"{\rtf1\ilfomacatclnup-1 Body}"#,
        r#"{\rtf1\ilfomacatclnup2 Body}"#,
        r#"{\rtf1\ilfomacatclnup6 Body}"#,
        r#"{\rtf1\ilfomacatclnup99999999999 Body}"#,
        r#"{\rtf1\grfdocevents Body}"#,
        r#"{\rtf1\grfdocevents-1 Body}"#,
        r#"{\rtf1\grfdocevents32768 Body}"#,
        r#"{\rtf1\grfdocevents2147483647 Body}"#,
        r#"{\rtf1\grfdocevents99999999999 Body}"#,
        r#"{\rtf1\fracwidth\fracwidth Body}"#,
        r#"{\rtf1\ilfomacatclnup0\ilfomacatclnup1 Body}"#,
        r#"{\rtf1\grfdocevents0\grfdocevents1 Body}"#,
        r#"{\rtf1{\*\fracwidth}Body}"#,
        r#"{\rtf1{\*\ilfomacatclnup0}Body}"#,
        r#"{\rtf1{\*\grfdocevents0}Body}"#,
        r#"{\rtf1{\fracwidth}Body}"#,
        r#"{\rtf1{\ilfomacatclnup0}Body}"#,
        r#"{\rtf1{\grfdocevents0}Body}"#,
        r#"{\rtf1 Body\fracwidth}"#,
        r#"{\rtf1 Body\ilfomacatclnup0}"#,
        r#"{\rtf1 Body\grfdocevents0}"#,
    ] {
        assert!(
            RtfDocument::parse(source).is_err(),
            "accepted malformed {source}"
        );
    }
}
