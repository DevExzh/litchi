//! LENIENCY_AUDIT F9.2: a manifest without a `/` root entry, or whose root
//! entry declares an empty media type, falls back to the `mimetype` member
//! the way LibreOffice does. Every package owner opens it by default and
//! records the fallback; a conforming manifest records nothing, and a
//! genuinely conflicting non-empty root stays a hard error.

#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "hand-built fixtures keep tolerance assertions focused"
)]

#[path = "support/odf_rezip.rs"]
mod odf_rezip;

use std::sync::Arc;

use litchi_core::{FileFormat, OwnedSource};
use litchi_odf_common::{
    Defect, OwnedPackage, PackageDefect, SourceBackedPackage, ToleranceReport, detect,
};

const ODT: &str = "application/vnd.oasis.opendocument.text";
const MANIFEST: &str = "META-INF/manifest.xml";

fn recorded(report: &ToleranceReport) -> Vec<(Defect, u16)> {
    report
        .defects()
        .iter()
        .map(|tolerated| (tolerated.defect, tolerated.context))
        .collect()
}

fn conforming_fixture() -> Vec<u8> {
    std::fs::read(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/odf/corpus/writer-table.odt"),
    )
    .unwrap()
}

fn root_fallback(context: u16) -> Vec<(Defect, u16)> {
    vec![(
        Defect::Package(PackageDefect::ManifestRootMediaTypeMissing),
        context,
    )]
}

/// Open `bytes` through every package owner and assert each accepts it,
/// adopts the `mimetype` member as the root media type, and records exactly
/// `defects`.
fn assert_opens_everywhere_recording(bytes: &[u8], defects: &[(Defect, u16)], label: &str) {
    // Detection never consulted the manifest root; it still classifies.
    assert_eq!(detect::bytes(bytes), Some(FileFormat::Odt), "{label}");

    // Owned open: the borrowed view adopts the fallback.
    let owned = OwnedPackage::from_bytes(bytes.to_vec()).unwrap();
    assert_eq!(
        recorded(owned.tolerance_report()),
        defects,
        "{label}: owned"
    );
    assert_eq!(owned.package().unwrap().manifest().mimetype, ODT, "{label}");
    assert_eq!(
        recorded(owned.clone_without_password().tolerance_report()),
        defects,
        "{label}: clone"
    );

    // Positional source-backed open, and materialization across owners.
    let positional =
        SourceBackedPackage::from_read_at(Arc::new(OwnedSource::new(bytes.to_vec()))).unwrap();
    assert_eq!(positional.mimetype().unwrap(), ODT, "{label}");
    assert_eq!(positional.manifest().unwrap().mimetype, ODT, "{label}");
    assert_eq!(
        recorded(positional.tolerance_report()),
        defects,
        "{label}: source-backed"
    );
    let materialized = positional.materialize().unwrap();
    assert_eq!(
        recorded(materialized.tolerance_report()),
        defects,
        "{label}: materialized"
    );

    // The prepared detection handoff retains the report.
    let prepared = detect::prepared_or_original(bytes.to_vec())
        .unwrap_or_else(|_| panic!("{label}: package must prepare"));
    assert_eq!(
        recorded(prepared.tolerance_report()),
        defects,
        "{label}: prepared"
    );

    // The family contract accepts it and carries the record.
    let family =
        litchi_odf_common::core::Package::from_bytes(bytes.to_vec(), ODT, "<office:text", "ODT")
            .unwrap();
    assert_eq!(
        recorded(family.tolerance_report()),
        defects,
        "{label}: family"
    );
}

#[test]
fn rootless_manifest_opens_and_records_the_fallback() {
    let bytes = odf_rezip::edit_member(
        &conforming_fixture(),
        MANIFEST,
        odf_rezip::without_manifest_root,
    );
    assert_opens_everywhere_recording(&bytes, &root_fallback(0), "rootless");
}

#[test]
fn empty_root_media_type_opens_and_records_the_fallback() {
    let bytes = odf_rezip::edit_member(
        &conforming_fixture(),
        MANIFEST,
        odf_rezip::with_empty_manifest_root_media_type,
    );
    assert_opens_everywhere_recording(&bytes, &root_fallback(1), "empty media type");
}

#[test]
fn the_same_package_with_its_root_entry_repairs_nothing() {
    let bytes = odf_rezip::rezip(&conforming_fixture(), odf_rezip::MimetypeLayout::CONFORMING);
    assert_opens_everywhere_recording(&bytes, &[], "conforming");
    assert_opens_everywhere_recording(&conforming_fixture(), &[], "original");
}

#[test]
fn a_conflicting_non_empty_root_stays_a_hard_error() {
    let bytes = odf_rezip::edit_member(&conforming_fixture(), MANIFEST, |manifest| {
        odf_rezip::with_manifest_root_media_type(
            manifest,
            "application/vnd.oasis.opendocument.spreadsheet",
        )
    });
    let error = SourceBackedPackage::from_read_at(Arc::new(OwnedSource::new(bytes.clone())))
        .unwrap_err()
        .to_string();
    assert!(error.contains("does not match mimetype"), "{error}");

    // The owned owner never compared the two values and records nothing:
    // a conflict is not a fallback.
    let owned = OwnedPackage::from_bytes(bytes).unwrap();
    assert!(owned.tolerance_report().is_clean());
}

#[test]
fn a_missing_manifest_stays_a_hard_error() {
    let conforming = conforming_fixture();
    let source = OwnedPackage::from_bytes(conforming).unwrap();
    let content = source.get_file("content.xml").unwrap();
    let bytes = odf_rezip::package(
        ODT.as_bytes(),
        &[("content.xml", content.as_slice())],
        odf_rezip::MimetypeLayout::CONFORMING,
    );
    assert!(SourceBackedPackage::from_read_at(Arc::new(OwnedSource::new(bytes.clone()))).is_err());
    assert!(
        litchi_odf_common::core::Package::from_bytes(bytes, ODT, "<office:text", "ODT").is_err()
    );
}

#[test]
fn a_manifest_without_any_entry_is_not_a_rootless_manifest() {
    // A document that lists nothing is a missing or foreign manifest, not
    // one that merely omits its root; the fallback does not engage.
    let bytes = odf_rezip::edit_member(&conforming_fixture(), MANIFEST, |_| {
        r#"<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"/>"#
            .to_string()
    });
    assert!(SourceBackedPackage::from_read_at(Arc::new(OwnedSource::new(bytes.clone()))).is_err());
    let owned = OwnedPackage::from_bytes(bytes).unwrap();
    assert!(owned.tolerance_report().is_clean());
}
