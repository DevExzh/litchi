//! LENIENCY_AUDIT F9.1 at the ODS facade: a re-zipped package whose
//! `mimetype` member is deflated, carries a local extra field, or is not the
//! first entry opens by default through every ODS open path, and the
//! deviation is recorded on the document handle's tolerance report.

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
use litchi_ods::{SourceBackedSpreadsheet, Spreadsheet};
use odf_rezip::MimetypeLayout;

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
            .join("../../test-data/odf/corpus/calc-two-sheets.ods"),
    )
    .unwrap()
}

fn assert_every_open_path_records(bytes: &[u8], defects: &[(Defect, u16)], label: &str) {
    let owned = Spreadsheet::from_bytes(bytes.to_vec()).unwrap();
    assert_eq!(
        recorded(owned.tolerance_report()),
        defects,
        "{label}: from_bytes"
    );
    assert_eq!(owned.sheet_names().len(), 2, "{label}: sheets");

    let prepared = litchi_odf_common::detect::prepared(bytes.to_vec())
        .unwrap_or_else(|| panic!("{label}: package must prepare"));
    let adopted = Spreadsheet::from_prepared_package(prepared).unwrap();
    assert_eq!(
        recorded(adopted.tolerance_report()),
        defects,
        "{label}: prepared"
    );

    let source =
        SourceBackedSpreadsheet::from_read_at(Arc::new(OwnedSource::new(bytes.to_vec()))).unwrap();
    assert_eq!(
        recorded(source.tolerance_report()),
        defects,
        "{label}: source-backed"
    );
}

#[test]
fn rezipped_ods_opens_by_default_and_records_each_deviation() {
    let conforming = conforming_fixture();
    assert_every_open_path_records(&conforming, &[], "conforming");

    let cases = [
        (
            MimetypeLayout {
                deflated: true,
                local_extra: false,
                last: false,
            },
            vec![(Defect::Package(PackageDefect::MimetypeDeflated), 8)],
        ),
        (
            MimetypeLayout {
                deflated: false,
                local_extra: true,
                last: false,
            },
            vec![(Defect::Package(PackageDefect::MimetypeExtraField), 0)],
        ),
        (
            MimetypeLayout {
                deflated: false,
                local_extra: false,
                last: true,
            },
            vec![(Defect::Package(PackageDefect::MimetypeNotFirst), 0)],
        ),
        (
            MimetypeLayout {
                deflated: true,
                local_extra: true,
                last: true,
            },
            vec![
                (Defect::Package(PackageDefect::MimetypeDeflated), 8),
                (Defect::Package(PackageDefect::MimetypeExtraField), 0),
                (Defect::Package(PackageDefect::MimetypeNotFirst), 0),
            ],
        ),
    ];
    for (layout, defects) in cases {
        let rezipped = odf_rezip::rezip(&conforming, layout);
        assert_every_open_path_records(&rezipped, &defects, &format!("{layout:?}"));
    }

    // The same re-zip without any deviation repairs nothing.
    let clean = odf_rezip::rezip(&conforming, MimetypeLayout::CONFORMING);
    assert_every_open_path_records(&clean, &[], "re-zipped conforming");
}
