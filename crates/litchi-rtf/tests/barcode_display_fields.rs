#![allow(
    clippy::expect_used,
    clippy::shadow_reuse,
    clippy::shadow_same,
    clippy::shadow_unrelated,
    clippy::unwrap_used,
    reason = "test assertions panic on failure by design and rebind fixture names across steps"
)]

use litchi_rtf::{BarcodeDisplayFieldKind, Field, FieldOwner, FieldStatus, RtfDocument, RtfWriter};

const FIXTURE: &str = include_str!(concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/../../test-data/rtf/barcode_display_fields.rtf"
));

fn write(document: &RtfDocument<'_>) -> Vec<u8> {
    let mut output = Vec::new();
    RtfWriter::new(&mut output)
        .write_document(document)
        .unwrap();
    output
}

#[test]
fn inspects_barcode_display_metadata_without_validating_or_generating_barcodes() {
    let document = RtfDocument::parse(FIXTURE).unwrap();

    assert_eq!(document.text(), "BeforeAfter");
    assert_eq!(document.barcode_field_count(), 0);
    assert_eq!(document.barcode_display_field_count(), 2);

    let fields = document.barcode_display_fields();
    let display = &fields[0];
    assert_eq!(display.kind(), BarcodeDisplayFieldKind::DisplayBarcode);
    assert_eq!(display.data_argument(), "https://example.invalid/qr");
    assert_eq!(display.barcode_type(), "QR");
    assert_eq!(display.switches().len(), 5);
    assert_eq!(display.switches()[0].name, "q");
    assert_eq!(display.switches()[0].value.as_deref(), Some("3"));
    assert_eq!(display.switches()[1].name, "h");
    assert_eq!(display.switches()[1].value.as_deref(), Some("720"));
    assert_eq!(display.switches()[2].name, "s");
    assert_eq!(display.switches()[2].value.as_deref(), Some("125"));
    assert_eq!(display.switches()[3].name, "t");
    assert!(display.switches()[3].value.is_none());
    assert_eq!(display.switches()[4].name, "*");
    assert_eq!(display.switches()[4].value.as_deref(), Some("MERGEFORMAT"));
    assert_eq!(display.cached_result(), Some("cached display"));
    assert_eq!(
        display.status(),
        FieldStatus {
            dirty: true,
            locked: true,
            ..FieldStatus::default()
        }
    );
    assert_eq!(display.owner(), FieldOwner::Body);
    assert_eq!(display.position(), "Before".len());

    let merge = &fields[1];
    assert_eq!(merge.kind(), BarcodeDisplayFieldKind::MergeBarcode);
    assert_eq!(merge.data_argument(), "CustomerCode");
    assert_eq!(merge.barcode_type(), "CODE128");
    assert_eq!(merge.switches().len(), 3);
    assert_eq!(merge.switches()[0].name, "t");
    assert!(merge.switches()[0].value.is_none());
    assert_eq!(merge.switches()[1].name, "x");
    assert!(merge.switches()[1].value.is_none());
    assert_eq!(merge.switches()[2].name, "z");
    assert_eq!(merge.switches()[2].value.as_deref(), Some("opaque"));
    assert_eq!(merge.cached_result(), Some("cached merge"));
    assert_eq!(
        merge.status(),
        FieldStatus {
            edited: true,
            ..FieldStatus::default()
        }
    );

    assert!(
        Field::parse_instruction("DISPLAYBARCODE")
            .barcode_display_field()
            .is_none()
    );
    assert!(
        Field::parse_instruction("MERGEBARCODE CustomerCode")
            .barcode_display_field()
            .is_none()
    );
}

#[test]
fn barcode_display_metadata_round_trips_deterministically() {
    let document = RtfDocument::parse(FIXTURE).unwrap();
    let reparsed = RtfDocument::parse_bytes(&write(&document)).unwrap();

    assert_eq!(reparsed.text(), document.text());
    assert_eq!(
        reparsed.barcode_display_field_count(),
        document.barcode_display_field_count()
    );

    let fields = reparsed.barcode_display_fields();
    assert_eq!(fields[0].kind(), BarcodeDisplayFieldKind::DisplayBarcode);
    assert_eq!(fields[0].data_argument(), "https://example.invalid/qr");
    assert_eq!(fields[0].barcode_type(), "QR");
    assert_eq!(fields[0].switches().len(), 5);
    assert_eq!(fields[1].kind(), BarcodeDisplayFieldKind::MergeBarcode);
    assert_eq!(fields[1].data_argument(), "CustomerCode");
    assert_eq!(fields[1].barcode_type(), "CODE128");
    assert_eq!(fields[1].switches().len(), 3);
}
