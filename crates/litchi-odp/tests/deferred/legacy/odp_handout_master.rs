//! Tests for `style:handout-master` against real presentation packages.

use litchi_odp::Package;
use std::path::PathBuf;

fn fixture(name: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/odf/odp")
        .join(name)
}

#[test]
fn reads_handout_master_from_real_presentations() {
    for name in [
        "cellspan.odp",
        "tdf102223.odp",
        "tdf105502.odp",
        "tdf169979.odp",
        "text-in-image.odp",
    ] {
        let package = Package::open(fixture(name)).unwrap();
        let master = package
            .handout_master()
            .unwrap_or_else(|error| panic!("{name}: {error}"))
            .unwrap_or_else(|| panic!("{name} has no handout master"));
        assert!(!master.page_layout_name.is_empty(), "{name}");
        assert!(
            master.xml.starts_with("<style:handout-master"),
            "{name}: {}",
            &master.xml[..master.xml.len().min(60)]
        );
    }
}

#[test]
fn handout_master_is_absent_in_text_documents() {
    // An ODT without a handout master reports None rather than erroring.
    let package = Package::open(
        PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/odf/odt/note-tracked-changes.odt"),
    );
    if let Ok(package) = package {
        assert!(package.handout_master().unwrap().is_none());
    }
}

#[test]
fn bom_prefixed_styles_xml_keeps_fragments_exact() {
    // tdf169979.odp stores styles.xml with a UTF-8 BOM; quick-xml reports
    // positions relative to the stripped text, and fragments must slice
    // against the same view.
    let package = Package::open(fixture("tdf169979.odp")).unwrap();
    for page in package.master_pages().unwrap() {
        assert!(
            page.xml.starts_with("<style:master-page"),
            "BOM shifted master-page fragment: {:?}",
            &page.xml[..page.xml.len().min(20)]
        );
    }
    let master = package.handout_master().unwrap().unwrap();
    assert!(master.xml.starts_with("<style:handout-master"));
}
