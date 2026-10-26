//! Body Date & Time smart fields (`litchi_pages::date_time`, record 0371).

use litchi_pages::date_time::{
    DisplayText, Format, Instant, LocaleIdentifier, Settings, TextRange,
};
use litchi_pages::{Package, TextPosition};

const BASIC: &[u8] = include_bytes!("../../../test-data/iwork/pages/basic.pages");

fn bytes(package: &Package) -> Vec<u8> {
    let mut output = Vec::new();
    package.write_to(&mut output).unwrap();
    output
}

fn settings(seconds: f64) -> Settings {
    Settings::fixed(
        Format::new("yyyy-MM-dd").unwrap(),
        LocaleIdentifier::new("en_US").unwrap(),
        Instant::from_reference_date_seconds(seconds).unwrap(),
    )
}

#[test]
fn insert_update_remove_round_trip_and_invert() {
    let package = Package::from_bytes(BASIC).unwrap();
    let text = package.text().unwrap();
    let mut edit = package.edit_body_date_time_fields().unwrap();
    edit.insert(
        TextPosition::from_utf16_index(0).unwrap(),
        DisplayText::new("2026-10-26").unwrap(),
        settings(1.0),
    )
    .unwrap();
    let inserted = edit.commit().unwrap();
    let fields = inserted.package().body_date_time_fields().unwrap();
    assert_eq!(fields.len(), 1);
    assert_eq!(
        fields[0].range(),
        TextRange::from_utf16_indexes(0, 10).unwrap()
    );
    assert!(inserted.package().text().unwrap().starts_with("2026-10-26"));
    let restored = inserted
        .package()
        .apply_body_date_time_fields(&inserted.patch().inverse())
        .unwrap();
    assert_eq!(bytes(restored.package()), BASIC);

    let mut edit = inserted.package().edit_body_date_time_fields().unwrap();
    edit.update(
        0,
        TextRange::from_utf16_indexes(0, 4).unwrap(),
        settings(2.0),
    )
    .unwrap();
    let updated = edit.commit().unwrap();
    assert_eq!(
        updated.package().body_date_time_fields().unwrap()[0].settings(),
        &settings(2.0)
    );

    let mut edit = updated.package().edit_body_date_time_fields().unwrap();
    edit.remove(0).unwrap();
    let removed = edit.commit().unwrap();
    assert!(
        removed
            .package()
            .body_date_time_fields()
            .unwrap()
            .is_empty()
    );
    assert_eq!(
        removed.package().text().unwrap(),
        format!("2026-10-26{text}")
    );
}
