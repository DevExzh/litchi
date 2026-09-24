//! Real host-origin VBA data vendored from Apache POI (Apache-2.0).
//!
//! POI ships each source beside its macro-enabled Office host fixture. These
//! tests exercise the MS-OVBA codec directly, with finite input and output
//! budgets, so the runtime crate does not depend on an OOXML ZIP reader.

#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "fixture setup and assertions intentionally panic on failure"
)]

use litchi_cfb::OleFile;
use litchi_vba::{Limits, codec, project::Project};
use std::io::Cursor;

const FIXTURE_LIMITS: Limits = Limits {
    max_cfb_bytes: 32 * 1024,
    max_compressed_stream_bytes: 16 * 1024,
    max_decompressed_stream_bytes: 32 * 1024,
    max_modules: 32,
    max_references: 32,
    max_licenses: 32,
    max_string_bytes: 16 * 1024,
    max_total_source_bytes: 32 * 1024,
};

const WORD: &[u8] = include_bytes!("../../../test-data/poi/test-data/document/SimpleMacro.vba");
const EXCEL: &[u8] = include_bytes!("../../../test-data/poi/test-data/spreadsheet/SimpleMacro.vba");
const POWERPOINT: &[u8] =
    include_bytes!("../../../test-data/poi/test-data/slideshow/SimpleMacro.vba");
const EXCEL_CFB: &[u8] =
    include_bytes!("../../../test-data/poi/test-data/spreadsheet/SimpleMacro.xls");
const EXCEL_FORM_CFB: &[u8] =
    include_bytes!("../../../test-data/poi/test-data/spreadsheet/59858.xls");

#[test]
fn real_word_vba_source_is_bounded_and_lossless() {
    assert_fixture(WORD, "This is a macro word processing document");
}

#[test]
fn real_excel_vba_source_is_bounded_and_lossless() {
    assert_fixture(EXCEL, "This is a macro workbook");
}

#[test]
fn real_powerpoint_vba_source_is_bounded_and_lossless() {
    assert_fixture(POWERPOINT, "This is a macro slideshow");
}

#[test]
fn real_excel_cfb_project_exposes_references_and_project_wm() {
    let mut ole = OleFile::open(Cursor::new(EXCEL_CFB)).expect("Excel CFB fixture");
    let project = Project::open(
        &mut ole,
        &["_VBA_PROJECT_CUR"],
        &Limits {
            max_cfb_bytes: EXCEL_CFB.len(),
            ..FIXTURE_LIMITS
        },
    )
    .expect("MS-OVBA project");
    assert!(!project.modules().is_empty());
    assert!(!project.references().is_empty());
    assert!(
        project
            .references()
            .iter()
            .all(|reference| reference.raw().is_some())
    );
    assert_eq!(
        project
            .project_wm(&FIXTURE_LIMITS)
            .expect("PROJECTwm parse")
            .map_or(0, |maps| maps.len()),
        project.modules().len()
    );
    assert_eq!(
        project
            .project_text(&FIXTURE_LIMITS)
            .expect("typed PROJECT stream")
            .project_name(),
        Some(project.name())
    );
}

#[test]
fn real_excel_form_project_exposes_designer_frame_metadata() {
    let mut ole = OleFile::open(Cursor::new(EXCEL_FORM_CFB)).expect("Excel form CFB fixture");
    let limits = Limits {
        max_cfb_bytes: EXCEL_FORM_CFB.len(),
        max_compressed_stream_bytes: 64 * 1024,
        max_decompressed_stream_bytes: 128 * 1024,
        max_total_source_bytes: 256 * 1024,
        ..FIXTURE_LIMITS
    };
    let project =
        Project::open(&mut ole, &["_VBA_PROJECT_CUR"], &limits).expect("MS-OVBA form project");
    // The native dir stream contains stdole and Office registered libraries,
    // followed by an MSForms REFERENCEORIGINAL with its nested control record.
    // REFERENCEPROJECT is covered by codec vectors, not this producer fixture.
    assert_eq!(project.references().len(), 3);
    assert_eq!(project.references()[0].name(), Some("stdole"));
    assert_eq!(project.references()[1].name(), Some("Office"));
    assert_eq!(project.references()[2].name(), Some("MSForms"));
    for reference in &project.references()[..2] {
        assert!(matches!(
            reference.kind(),
            litchi_vba::dir::ReferenceKind::Registered { .. }
        ));
    }
    assert!(matches!(
        project.references()[2].kind(),
        litchi_vba::dir::ReferenceKind::Original { .. }
    ));
    assert_eq!(project.designer_frames().len(), 1);
    let frame = &project.designer_frames()[0];
    assert_eq!(frame.module_name(), "frmPageCreation");
    assert_eq!(frame.class_id(), "{C62A69F0-16DC-11CE-9E98-00AA00574A4F}");
    assert!(
        frame
            .properties()
            .iter()
            .any(|property| { property.name() == "StartUpPosition" && property.value() == "1" })
    );
}

fn assert_fixture(source: &[u8], expected_source: &str) {
    assert!(source.len() <= FIXTURE_LIMITS.max_decompressed_stream_bytes);
    assert!(
        std::str::from_utf8(source)
            .unwrap()
            .contains(expected_source)
    );

    let compressed = codec::encode(source, &FIXTURE_LIMITS).unwrap();
    assert!(compressed.len() <= FIXTURE_LIMITS.max_compressed_stream_bytes);
    assert_eq!(codec::decode(&compressed, &FIXTURE_LIMITS).unwrap(), source);
    assert_eq!(codec::encode(source, &FIXTURE_LIMITS).unwrap(), compressed);
}
