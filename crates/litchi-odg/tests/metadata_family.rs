//! The ADR 0034 §7 metadata family over the shared ODF engine (record 0288).

use litchi_odg::Package;
use litchi_odg::metadata::Field;

fn meta_xml(package: &Package) -> String {
    let owned = litchi_odf_common::OwnedPackage::from_bytes(package.as_bytes().to_vec()).unwrap();
    String::from_utf8(owned.get_file("meta.xml").unwrap()).unwrap()
}

fn round_trip(package: &Package) {
    let mut edit = package.edit_metadata();
    edit.set(Field::Title, Some("Title & more"))
        .unwrap()
        .set(Field::Keywords, Some("alpha, beta"))
        .unwrap()
        .set(Field::Author, Some("Ann"))
        .unwrap();
    let commit = edit.commit().unwrap();
    assert!(commit.diagnostics().changed());
    assert_eq!(commit.diagnostics().field_count(), 3);
    let next = commit.package().clone();
    let meta = meta_xml(&next);
    assert!(meta.contains(">Title &amp; more</"), "{meta}");
    assert!(meta.contains(">alpha</") && meta.contains(">beta</"));
    let reopened = Package::from_bytes(next.as_bytes()).unwrap();
    assert_eq!(reopened.as_bytes(), next.as_bytes());

    // A second edit replaces the spliced element in place.
    let mut edit = next.edit_metadata();
    edit.set(Field::Title, Some("Second")).unwrap();
    let second = edit.commit().unwrap().into_package();
    let meta = meta_xml(&second);
    assert!(meta.contains(">Second</") && !meta.contains("Title &amp; more"));

    // Unchanged values are an exact no-op.
    let mut edit = next.edit_metadata();
    edit.set(Field::Title, Some("Title & more")).unwrap();
    let commit = edit.commit().unwrap();
    assert!(!commit.diagnostics().changed());
    assert!(commit.patch().is_noop());
    assert_eq!(commit.package().as_bytes(), next.as_bytes());

    // The inverse patch restores the source bytes; a stale source is refused.
    let mut edit = package.edit_metadata();
    edit.set(Field::Subject, Some("S")).unwrap();
    let (after, patch) = edit.commit().unwrap().into_parts();
    let restored = after
        .apply_metadata(&patch.inverse())
        .unwrap()
        .into_package();
    assert_eq!(restored.as_bytes(), package.as_bytes());
    assert_eq!(
        package.apply_metadata(&patch).unwrap().package().as_bytes(),
        after.as_bytes()
    );
    assert!(after.apply_metadata(&patch).is_err());
}

#[test]
fn fresh_package_metadata_round_trip() {
    round_trip(&Package::new().unwrap());
}

#[test]
fn removing_a_field_drops_its_element() {
    let package = Package::new().unwrap();
    let mut edit = package.edit_metadata();
    edit.set(Field::Description, Some("D")).unwrap();
    let next = edit.commit().unwrap().into_package();
    let mut edit = next.edit_metadata();
    edit.set(Field::Description, None).unwrap();
    let cleared = edit.commit().unwrap().into_package();
    assert!(!meta_xml(&cleared).contains(">D</"));
}

#[test]
fn values_xml_cannot_carry_are_refused() {
    let package = Package::new().unwrap();
    let mut edit = package.edit_metadata();
    assert!(edit.set(Field::Title, Some("bad\u{1}")).is_err());
}

/// A producer package: its `meta.xml` (with LibreOffice's prolog) is
/// spliced, every other byte kept.
#[test]
fn producer_package_metadata_round_trip() {
    let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data")
        .join("libreoffice-core/desktop/qa/data/3page.odg");
    let package = Package::from_bytes(&std::fs::read(path).unwrap()).unwrap();
    round_trip(&package);
}
