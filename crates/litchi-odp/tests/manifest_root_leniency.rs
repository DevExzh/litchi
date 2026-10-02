//! LENIENCY_AUDIT F9.2 at the ODP facade: a manifest without a `/` root
//! entry, or whose root entry has an empty media type, falls back to the
//! `mimetype` member through every ODP open path, and the fallback is
//! recorded on the document handle's tolerance report. A conflicting
//! non-empty root stays a hard error.

#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "hand-built fixtures keep tolerance assertions focused"
)]

#[path = "../../litchi-odf-common/tests/support/odf_rezip.rs"]
mod odf_rezip;

use std::sync::Arc;

use litchi_core::OwnedSource;
use litchi_odf_common::{Defect, PackageDefect, ToleranceReport};
use litchi_odp::{Presentation, SourceBackedPresentation};

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
            .join("../../test-data/odf/corpus/impress-basic.odp"),
    )
    .unwrap()
}

fn assert_every_open_path_records(bytes: &[u8], defects: &[(Defect, u16)], label: &str) {
    let owned = Presentation::from_bytes(bytes.to_vec()).unwrap();
    assert_eq!(
        recorded(owned.tolerance_report()),
        defects,
        "{label}: from_bytes"
    );

    let prepared = litchi_odf_common::detect::prepared(bytes.to_vec())
        .unwrap_or_else(|| panic!("{label}: package must prepare"));
    let adopted = Presentation::from_prepared_package(prepared).unwrap();
    assert_eq!(
        recorded(adopted.tolerance_report()),
        defects,
        "{label}: prepared"
    );

    let source =
        SourceBackedPresentation::from_read_at(Arc::new(OwnedSource::new(bytes.to_vec()))).unwrap();
    assert_eq!(
        recorded(source.tolerance_report()),
        defects,
        "{label}: source-backed"
    );
}

#[test]
fn rootless_manifest_opens_by_default_and_records_the_fallback() {
    let conforming = conforming_fixture();
    assert_every_open_path_records(&conforming, &[], "conforming");

    let rootless = odf_rezip::edit_member(&conforming, MANIFEST, odf_rezip::without_manifest_root);
    assert_every_open_path_records(
        &rootless,
        &[(
            Defect::Package(PackageDefect::ManifestRootMediaTypeMissing),
            0,
        )],
        "rootless",
    );

    let empty = odf_rezip::edit_member(
        &conforming,
        MANIFEST,
        odf_rezip::with_empty_manifest_root_media_type,
    );
    assert_every_open_path_records(
        &empty,
        &[(
            Defect::Package(PackageDefect::ManifestRootMediaTypeMissing),
            1,
        )],
        "empty media type",
    );
}

#[test]
fn conflicting_manifest_root_stays_a_hard_error() {
    let conflicting = odf_rezip::edit_member(&conforming_fixture(), MANIFEST, |manifest| {
        odf_rezip::with_manifest_root_media_type(
            manifest,
            "application/vnd.oasis.opendocument.text",
        )
    });
    assert!(
        SourceBackedPresentation::from_read_at(Arc::new(OwnedSource::new(conflicting))).is_err()
    );
}
