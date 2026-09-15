#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]
//! A value-only patch must be accepted or refused by what the workbook is, not
//! by which hash seed a particular load happened to draw.
//!
//! `capture_auxiliary` and `capture_auxiliary_source` build the styles-and-theme
//! `PartState` array by walking `workbook.rels().iter()`, which is a `HashMap`
//! iteration, and `SourceState::same_owner` compares two such arrays with slice
//! `==`. On the ordinary two-relationship shape — one styles relationship and
//! one theme relationship, which is what Excel emits for nearly every workbook —
//! the array is `[styles, theme]` in one load and `[theme, styles]` in another,
//! so a patch captured against one load is refused against another load of the
//! same bytes as `Error::PatchConflict`. Every other captured relationship array
//! in that file is sorted (`capture_relationships` ends in
//! `sort_unstable_by(|left, right| left.id.cmp(&right.id))`); this one was the
//! outlier. Change 0628 found it and left it for this change.
//!
//! The exact no-op used below is the sharpest form of the defect: nothing about
//! the document changed, the patch stages no edit, and the refusal is a coin
//! flip. ADR 0006 requires serialization to be deterministic unless a `Clock`,
//! actor identity, or cryptographic RNG is explicitly supplied, and none of the
//! three reaches this path; ADR 0003's source-checked patches must fail on a
//! real overlap, not on a hash seed.
//!
//! Each `Relationships::new()` builds a fresh `HashMap`, whose `RandomState` is
//! seeded from a per-thread counter that advances on every construction, so
//! repeating a load in one process varies the visit order exactly as separate
//! processes do. With two candidate orders a single trial agrees by chance about
//! half the time, so the loop repeats enough times that an accidental agreement
//! across every trial is not worth considering.

use std::sync::Arc;

use litchi_core::OwnedSource;
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{BlobPart, OpcPackage, PackURI, PackageWriter, TargetMode};
use litchi_xlsx::Selector;
use litchi_xlsx::cell_values::SourceBackedEditor;

/// Repetitions per determinism assertion; 2^-128 accidental agreement.
const REPEATS: usize = 128;

const SML: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const MAIN: &str = "/xl/workbook.xml";
const SHEET: &str = "/xl/worksheets/sheet1.xml";
const STYLES: &str = "/xl/styles.xml";
const THEME: &str = "/xl/theme/theme1.xml";

/// The ordinary shape: one workbook owning a worksheet, a styles part and a
/// theme part. The two auxiliary relationships are what `capture_auxiliary`
/// walks.
fn workbook_with_styles_and_theme() -> Vec<u8> {
    let workbook_xml = format!(
        r#"<workbook xmlns="{SML}" xmlns:r="{REL}"><bookViews><workbookView/></bookViews><sheets><sheet name="Sheet1" sheetId="1" r:id="rIdSheet"/></sheets></workbook>"#
    );
    let sheet_xml = format!(
        r#"<worksheet xmlns="{SML}"><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>"#
    );
    let styles_xml =
        format!(r#"<styleSheet xmlns="{SML}"><cellXfs count="1"><xf/></cellXfs></styleSheet>"#);

    let mut package = OpcPackage::new();
    for (name, content_type, blob) in [
        (MAIN, ct::SML_SHEET_MAIN, workbook_xml.into_bytes()),
        (SHEET, ct::SML_WORKSHEET, sheet_xml.into_bytes()),
        (STYLES, ct::SML_STYLES, styles_xml.into_bytes()),
        (THEME, ct::OFC_THEME, b"<theme/>".to_vec()),
    ] {
        package
            .try_add_part(Box::new(BlobPart::new(
                PackURI::new(name).unwrap(),
                content_type.to_owned(),
                blob,
            )))
            .unwrap();
    }
    let workbook = package
        .get_part_mut(&PackURI::new(MAIN).unwrap())
        .unwrap()
        .rels_mut();
    for (reltype, target, r_id) in [
        (rt::WORKSHEET, "worksheets/sheet1.xml", "rIdSheet"),
        (rt::STYLES, "styles.xml", "rIdStyles"),
        (rt::THEME, "theme/theme1.xml", "rIdTheme"),
    ] {
        workbook
            .try_add_relationship(
                reltype.to_owned(),
                target.to_owned(),
                r_id.to_owned(),
                TargetMode::Internal,
            )
            .unwrap();
    }
    package.relate_to("xl/workbook.xml", rt::OFFICE_DOCUMENT);
    PackageWriter::to_bytes(&package).unwrap()
}

/// A workbook whose only auxiliary relationship is the styles part: one
/// captured `PartState`, so no order exists to get wrong. This is the negative
/// control — it passed before the fix and must keep passing after it.
fn workbook_with_styles_only() -> Vec<u8> {
    let mut package = OpcPackage::from_bytes(&workbook_with_styles_and_theme()).unwrap();
    package
        .get_part_mut(&PackURI::new(MAIN).unwrap())
        .unwrap()
        .rels_mut()
        .remove("rIdTheme");
    package.remove_part(&PackURI::new(THEME).unwrap());
    PackageWriter::to_bytes(&package).unwrap()
}

/// Capture a multi-sheet transaction against one load and publish its patch
/// against an independent load of the same bytes, the way a caller that keeps a
/// patch across a reopen does.
fn noop_patch_survives_a_reload(bytes: &[u8]) -> litchi_xlsx::Result<()> {
    let editor =
        SourceBackedEditor::from_read_at(Arc::new(OwnedSource::new(bytes.to_vec()))).unwrap();
    let commit = editor
        .edit_sheets([Selector::from("Sheet1")])
        .unwrap()
        .commit()
        .unwrap();
    assert!(!commit.changed(), "the transaction stages no edit");
    assert!(commit.patch().is_empty(), "the patch is an exact no-op");
    let mut replay = OpcPackage::from_bytes(bytes).unwrap();
    commit.patch().apply(&mut replay)
}

#[test]
fn exact_noop_patch_is_accepted_against_every_reload_of_the_same_workbook() {
    let bytes = workbook_with_styles_and_theme();
    for attempt in 0..REPEATS {
        noop_patch_survives_a_reload(&bytes).unwrap_or_else(|error| {
            panic!(
                "attempt {attempt}: an exact no-op patch was refused against a reload of the same \
                 bytes: {error}"
            )
        });
    }
}

#[test]
fn single_auxiliary_relationship_keeps_accepting_its_noop_patch() {
    let bytes = workbook_with_styles_only();
    for attempt in 0..REPEATS {
        noop_patch_survives_a_reload(&bytes).unwrap_or_else(|error| {
            panic!("attempt {attempt}: single-auxiliary no-op patch was refused: {error}")
        });
    }
}

#[test]
fn a_genuinely_different_workbook_is_still_refused() {
    let bytes = workbook_with_styles_and_theme();
    let editor =
        SourceBackedEditor::from_read_at(Arc::new(OwnedSource::new(bytes.clone()))).unwrap();
    let commit = editor
        .edit_sheets([Selector::from("Sheet1")])
        .unwrap()
        .commit()
        .unwrap();

    let mut altered = OpcPackage::from_bytes(&bytes).unwrap();
    altered
        .get_part_mut(&PackURI::new(THEME).unwrap())
        .unwrap()
        .set_blob(b"<theme><changed/></theme>".to_vec());
    assert!(
        matches!(
            commit.patch().apply(&mut altered),
            Err(litchi_xlsx::Error::PatchConflict { .. })
        ),
        "a changed theme part must still be a conflict"
    );
}
