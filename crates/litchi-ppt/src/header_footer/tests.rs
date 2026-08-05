use super::*;
use crate::PptRecordType;
use crate::records::PptRecord;

const POI_SLIDE_BYTES: &[u8] = &[
    0x3F, 0x00, 0xD9, 0x0F, 0x2E, 0, 0, 0, 0, 0, 0xDA, 0x0F, 4, 0, 0, 0, 0, 0, 0x23, 0, 0x20, 0,
    0xBA, 0x0F, 0x1A, 0, 0, 0, 0x4D, 0, 0x79, 0, 0x20, 0, 0x46, 0, 0x6F, 0, 0x6F, 0, 0x74, 0, 0x65,
    0, 0x72, 0, 0x20, 0, 0x2D, 0, 0x20, 0, 0x31, 0,
];
const POI_NOTES_BYTES: &[u8] = &[
    0x4F, 0, 0xD9, 0x0F, 0x48, 0, 0, 0, 0, 0, 0xDA, 0x0F, 4, 0, 0, 0, 0, 0, 0x3D, 0, 0x10, 0, 0xBA,
    0x0F, 0x16, 0, 0, 0, 0x4E, 0, 0x6F, 0, 0x74, 0, 0x65, 0, 0x20, 0, 0x48, 0, 0x65, 0, 0x61, 0,
    0x64, 0, 0x65, 0, 0x72, 0, 0x20, 0, 0xBA, 0x0F, 0x16, 0, 0, 0, 0x4E, 0, 0x6F, 0, 0x74, 0, 0x65,
    0, 0x20, 0, 0x46, 0, 0x6F, 0, 0x6F, 0, 0x74, 0, 0x65, 0, 0x72, 0,
];

fn parsed(bytes: &[u8], scope: PowerPointHeaderFooterScope) -> PowerPointHeaderFooter {
    let (record, consumed) = PptRecord::parse(bytes, 0).expect("record");
    assert_eq!(consumed, bytes.len());
    PowerPointHeaderFooter::parse_record(&record, scope).expect("header/footer")
}

#[test]
fn poi_record_arrays_are_byte_identical() {
    let slide = parsed(
        POI_SLIDE_BYTES,
        PowerPointHeaderFooterScope::PresentationSlides,
    );
    assert_eq!(slide.footer.as_deref(), Some("My Footer - 1"));
    assert_eq!(slide.to_record_bytes().unwrap(), POI_SLIDE_BYTES);

    let notes = parsed(
        POI_NOTES_BYTES,
        PowerPointHeaderFooterScope::NotesAndHandouts,
    );
    assert_eq!(notes.header.as_deref(), Some("Note Header"));
    assert_eq!(notes.footer.as_deref(), Some("Note Footer"));
    assert_eq!(notes.to_record_bytes().unwrap(), POI_NOTES_BYTES);
}

#[test]
fn all_flags_format_13_empty_and_local_roundtrip() {
    let value = PowerPointHeaderFooter {
        scope: PowerPointHeaderFooterScope::Local {
            parent: PowerPointHeaderFooterParent::Slide,
            parent_ordinal: PowerPointHeaderFooterParentOrdinal(7),
        },
        options: PowerPointHeaderFooterOptions {
            datetime_format: PowerPointDateTimeFormatId::new(13).unwrap(),
            show_date: true,
            use_current_datetime: true,
            use_user_date: true,
            show_slide_number: true,
            show_header: true,
            show_footer: true,
        },
        user_date: Some(String::new()),
        header: None,
        footer: Some(String::new()),
        placeholder_display: None,
    };
    let bytes = value.to_record_bytes().unwrap();
    let reparsed = parsed(&bytes, value.scope);
    assert_eq!(reparsed, value);
}

#[test]
fn malformed_record_matrix_is_rejected() {
    let mut cases = Vec::new();
    for (offset, value) in [
        (0, 0x3Eu8),
        (1, 0x01),
        (2, 0xD8),
        (4, 0xFF),
        (8, 0x01),
        (10, 0xD9),
        (12, 0x05),
        (16, 0x0E),
        (17, 0x80),
        (18, 0x40),
        (24, 0x19),
        (28, 0x01),
    ] {
        let mut bytes = POI_SLIDE_BYTES.to_vec();
        bytes[offset] = value;
        cases.push(bytes);
    }
    let mut invalid_utf16 = POI_SLIDE_BYTES.to_vec();
    invalid_utf16[28] = 0x00;
    invalid_utf16[29] = 0xD8;
    cases.push(invalid_utf16);

    for bytes in cases {
        let rejected = PptRecord::parse(&bytes, 0)
            .and_then(|(record, _)| {
                PowerPointHeaderFooter::parse_record(
                    &record,
                    PowerPointHeaderFooterScope::PresentationSlides,
                )
                .map(|_| (record, 0))
            })
            .is_err();
        assert!(rejected, "malformed bytes were accepted");
    }
}

#[test]
fn illegal_header_controls_and_oversize_user_date_are_rejected() {
    let invalid_header = PowerPointHeaderFooter {
        scope: PowerPointHeaderFooterScope::PresentationSlides,
        options: PowerPointHeaderFooterOptions::default(),
        user_date: None,
        header: Some("not permitted".to_string()),
        footer: None,
        placeholder_display: None,
    };
    assert!(invalid_header.to_record_bytes().is_err());

    let control = PowerPointHeaderFooter {
        scope: PowerPointHeaderFooterScope::NotesAndHandouts,
        options: PowerPointHeaderFooterOptions::default(),
        user_date: None,
        header: None,
        footer: Some("bad\nfooter".to_string()),
        placeholder_display: None,
    };
    assert!(control.to_record_bytes().is_err());

    let user_date = PowerPointHeaderFooter {
        scope: PowerPointHeaderFooterScope::NotesAndHandouts,
        options: PowerPointHeaderFooterOptions::default(),
        user_date: Some("x".repeat(256)),
        header: None,
        footer: None,
        placeholder_display: None,
    };
    assert!(user_date.to_record_bytes().is_err());
}

#[test]
fn placement_duplicate_and_order_violations_are_rejected() {
    let (container, _) = PptRecord::parse(POI_SLIDE_BYTES, 0).unwrap();
    let atom = container.children[0].clone();
    let footer = container.children[1].clone();

    let mut out_of_order = container.clone();
    out_of_order.data.clear();
    out_of_order.children = vec![footer.clone(), atom.clone()];
    out_of_order.data_length = 0;
    assert!(
        PowerPointHeaderFooter::parse_record(
            &out_of_order,
            PowerPointHeaderFooterScope::PresentationSlides,
        )
        .is_err()
    );

    let document = PptRecord {
        record_type: PptRecordType::Document,
        record_type_raw: 1000,
        version: 0xF,
        instance: 0,
        data_length: 0,
        data: Vec::new(),
        children: vec![container.clone(), container.clone()],
    };
    let records = vec![&document, &document.children[0], &document.children[1]];
    assert!(PowerPointHeaderFooters::parse_record_tree(&records).is_err());

    let wrong_parent = PptRecord {
        record_type: PptRecordType::Notes,
        record_type_raw: 1008,
        version: 0xF,
        instance: 0,
        data_length: 0,
        data: Vec::new(),
        children: vec![container],
    };
    let empty_document = PptRecord {
        record_type: PptRecordType::Document,
        record_type_raw: 1000,
        version: 0xF,
        instance: 0,
        data_length: 0,
        data: Vec::new(),
        children: Vec::new(),
    };
    let records = vec![&empty_document, &wrong_parent, &wrong_parent.children[0]];
    assert!(PowerPointHeaderFooters::parse_record_tree(&records).is_err());
}
