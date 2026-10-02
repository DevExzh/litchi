//! LENIENCY_AUDIT F9.1: a re-zipped package whose `mimetype` member is
//! deflated, carries a local extra field, or is not the first entry opens by
//! default through every package owner and detector, and the deviation is
//! recorded in the opening [`ToleranceReport`]. A conforming package records
//! nothing, and structural or content defects stay hard errors.

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
use odf_rezip::{MimetypeLayout, minimal_package, minimal_package_with_mimetype_content};

const ODT: &str = "application/vnd.oasis.opendocument.text";

const DEFLATED: MimetypeLayout = MimetypeLayout {
    deflated: true,
    local_extra: false,
    last: false,
};
const EXTRA_FIELD: MimetypeLayout = MimetypeLayout {
    deflated: false,
    local_extra: true,
    last: false,
};
const NOT_FIRST: MimetypeLayout = MimetypeLayout {
    deflated: false,
    local_extra: false,
    last: true,
};
const ALL_THREE: MimetypeLayout = MimetypeLayout {
    deflated: true,
    local_extra: true,
    last: true,
};

fn recorded(report: &ToleranceReport) -> Vec<(Defect, u16)> {
    report
        .defects()
        .iter()
        .map(|tolerated| (tolerated.defect, tolerated.context))
        .collect()
}

fn expected(layout: MimetypeLayout) -> Vec<(Defect, u16)> {
    let mut defects = Vec::new();
    if layout.deflated {
        defects.push((Defect::Package(PackageDefect::MimetypeDeflated), 8));
    }
    if layout.local_extra {
        defects.push((Defect::Package(PackageDefect::MimetypeExtraField), 0));
    }
    if layout.last {
        defects.push((Defect::Package(PackageDefect::MimetypeNotFirst), 0));
    }
    defects
}

/// Open `bytes` through every package owner and detector and assert each one
/// accepts the package and records exactly `defects`.
fn assert_opens_everywhere_recording(bytes: &[u8], defects: &[(Defect, u16)], label: &str) {
    // Detection classifies the package from every input shape.
    assert_eq!(
        detect::packaged_mime(bytes),
        Some(FileFormat::Odt),
        "{label}"
    );
    assert_eq!(detect::bytes(bytes), Some(FileFormat::Odt), "{label}");
    let source = OwnedSource::new(bytes.to_vec());
    assert_eq!(
        detect::packaged_mime_read_at(&source).unwrap(),
        Some(FileFormat::Odt),
        "{label}"
    );
    let mut cursor = std::io::Cursor::new(bytes.to_vec());
    assert_eq!(
        detect::reader(&mut cursor),
        Some(FileFormat::Odt),
        "{label}"
    );

    // Owned open.
    let owned = OwnedPackage::from_bytes(bytes.to_vec()).unwrap();
    assert_eq!(owned.mimetype().unwrap(), ODT, "{label}");
    assert_eq!(recorded(owned.tolerance_report()), defects, "{label}");
    assert_eq!(
        recorded(owned.clone_without_password().tolerance_report()),
        defects,
        "{label}"
    );

    // Positional source-backed open, and materialization across owners.
    let positional =
        SourceBackedPackage::from_read_at(Arc::new(OwnedSource::new(bytes.to_vec()))).unwrap();
    assert_eq!(positional.mimetype().unwrap(), ODT, "{label}");
    assert_eq!(recorded(positional.tolerance_report()), defects, "{label}");
    let materialized = positional.materialize().unwrap();
    assert_eq!(
        recorded(materialized.tolerance_report()),
        defects,
        "{label}"
    );

    // The prepared detection handoff retains the report.
    let prepared = detect::prepared_or_original(bytes.to_vec())
        .unwrap_or_else(|_| panic!("{label}: tolerated package must prepare"));
    assert_eq!(prepared.format(), FileFormat::Odt, "{label}");
    assert_eq!(recorded(prepared.tolerance_report()), defects, "{label}");
}

#[test]
fn deflated_mimetype_opens_and_records_the_compression_method() {
    let bytes = minimal_package(ODT, "office:text", DEFLATED);
    assert_opens_everywhere_recording(&bytes, &expected(DEFLATED), "deflated");
}

#[test]
fn mimetype_local_extra_field_opens_and_is_recorded() {
    let bytes = minimal_package(ODT, "office:text", EXTRA_FIELD);
    assert_opens_everywhere_recording(&bytes, &expected(EXTRA_FIELD), "extra field");
}

#[test]
fn relocated_mimetype_opens_and_is_recorded() {
    let bytes = minimal_package(ODT, "office:text", NOT_FIRST);
    assert_opens_everywhere_recording(&bytes, &expected(NOT_FIRST), "not first");
}

#[test]
fn combined_deviations_are_each_recorded_in_roadmap_order() {
    let bytes = minimal_package(ODT, "office:text", ALL_THREE);
    let defects = expected(ALL_THREE);
    assert_eq!(defects.len(), 3);
    assert_opens_everywhere_recording(&bytes, &defects, "all three");
}

#[test]
fn the_same_package_without_the_deviation_repairs_nothing() {
    let bytes = minimal_package(ODT, "office:text", MimetypeLayout::CONFORMING);
    assert_opens_everywhere_recording(&bytes, &[], "conforming");
}

#[test]
fn rezipped_corpus_fixture_records_its_deviations() {
    let conforming = std::fs::read(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/odf/corpus/writer-table.odt"),
    )
    .unwrap();
    let clean = OwnedPackage::from_bytes(conforming.clone()).unwrap();
    assert!(clean.tolerance_report().is_clean());

    let rezipped = odf_rezip::rezip(&conforming, ALL_THREE);
    assert_opens_everywhere_recording(&rezipped, &expected(ALL_THREE), "rezipped corpus");
}

#[test]
fn garbage_mimetype_content_still_fails_under_every_layout() {
    for layout in [DEFLATED, EXTRA_FIELD, NOT_FIRST, ALL_THREE] {
        let bytes = minimal_package_with_mimetype_content(
            ODT,
            b"application/x-not-an-odf-type",
            "office:text",
            layout,
        );
        assert_eq!(detect::packaged_mime(&bytes), None, "{layout:?}");
        assert_eq!(detect::bytes(&bytes), None, "{layout:?}");
        assert_eq!(
            detect::packaged_mime_read_at(&OwnedSource::new(bytes.clone())).unwrap(),
            None,
            "{layout:?}"
        );
        assert!(
            detect::prepared_or_original(bytes.clone()).is_err(),
            "{layout:?}"
        );
        // The family contract still names the expected MIME type exactly.
        assert!(
            litchi_odf_common::core::Package::from_bytes(bytes, ODT, "<office:text", "ODT")
                .is_err(),
            "{layout:?}"
        );
    }
}

#[test]
fn structural_defects_stay_hard_errors() {
    // Non-ZIP input.
    assert_eq!(detect::packaged_mime(b"not a zip archive at all"), None);
    assert!(OwnedPackage::from_bytes(b"not a zip archive at all".to_vec()).is_err());

    // A truncated archive whose central directory is gone.
    let bytes = minimal_package(ODT, "office:text", NOT_FIRST);
    let truncated = bytes[..bytes.len() - 30].to_vec();
    assert_eq!(detect::packaged_mime(&truncated), None);
    assert!(
        SourceBackedPackage::from_read_at(Arc::new(OwnedSource::new(truncated.clone()))).is_err()
    );
    assert!(OwnedPackage::from_bytes(truncated).is_err());

    // A corrupted deflated `mimetype` payload fails its CRC check.
    let mut corrupt = minimal_package(ODT, "office:text", DEFLATED);
    let payload = 30 + "mimetype".len() + 5;
    corrupt[payload] ^= 0x20;
    assert_eq!(detect::packaged_mime(&corrupt), None);
    assert!(detect::prepared_or_original(corrupt).is_err());

    // A missing manifest is still rejected by the positional owner.
    let no_manifest = odf_rezip::package(
        ODT.as_bytes(),
        &[(
            "content.xml",
            br#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><office:body><office:text/></office:body></office:document-content>"#,
        )],
        NOT_FIRST,
    );
    assert!(SourceBackedPackage::from_read_at(Arc::new(OwnedSource::new(no_manifest))).is_err());

    // A prefixed archive is not a relocated member.
    let mut prefixed = b"prefix bytes".to_vec();
    prefixed.extend_from_slice(&minimal_package(
        ODT,
        "office:text",
        MimetypeLayout::CONFORMING,
    ));
    assert!(
        SourceBackedPackage::from_read_at(Arc::new(OwnedSource::new(prefixed.clone()))).is_err()
    );
    assert_eq!(detect::packaged_mime(&prefixed), None);
}

#[test]
fn an_ooxml_first_entry_skips_the_tolerant_lookup() {
    // An archive led by the OPC catalog is an OOXML candidate; the tolerant
    // `mimetype` lookup does not engage for it.
    let bytes = odf_rezip::package(
        ODT.as_bytes(),
        &[("[Content_Types].xml", b"<Types/>")],
        NOT_FIRST,
    );
    assert_eq!(detect::packaged_mime(&bytes), None);
    assert_eq!(
        detect::packaged_mime_read_at(&OwnedSource::new(bytes)).unwrap(),
        None
    );
}
