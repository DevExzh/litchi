//! Observe real commit-local collector calls without instrumenting release builds.
use std::cell::Cell;

use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{BlobPart, OpcPackage, PackURI, PackageWriter, SourceBackedPackage, TargetMode};
use litchi_sheet::Cell as Address;

use crate::cell_values::SourceBackedEditor;

std::thread_local! {
    static COUNTS: Cell<(usize, usize)> = const { Cell::new((0, 0)) };
}

pub(super) fn record_attempt() {
    COUNTS.with(|counts| {
        let (attempts, accepted) = counts.get();
        counts.set((attempts + 1, accepted));
    });
}

pub(super) fn record_acceptance() {
    COUNTS.with(|counts| {
        let (attempts, accepted) = counts.get();
        counts.set((attempts, accepted + 1));
    });
}

fn editor() -> SourceBackedEditor {
    const SML: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    let mut package = OpcPackage::new();
    for (path, content_type, xml) in [
        (
            "/xl/workbook.xml",
            ct::SML_SHEET_MAIN,
            format!(
                r#"<workbook xmlns="{SML}" xmlns:r="{REL}"><sheets><sheet name="Sheet1" sheetId="1" r:id="rIdSheet"/></sheets></workbook>"#
            ),
        ),
        (
            "/xl/worksheets/sheet1.xml",
            ct::SML_WORKSHEET,
            format!(
                r#"<worksheet xmlns="{SML}"><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>"#
            ),
        ),
    ] {
        package
            .try_add_part(Box::new(BlobPart::new(
                PackURI::new(path).expect("fixture part URI"),
                content_type.to_owned(),
                xml.into_bytes(),
            )))
            .expect("fixture part");
    }
    package
        .get_part_mut(&PackURI::new("/xl/workbook.xml").expect("workbook URI"))
        .expect("workbook part")
        .rels_mut()
        .try_add_relationship(
            rt::WORKSHEET.to_owned(),
            "worksheets/sheet1.xml".to_owned(),
            "rIdSheet".to_owned(),
            TargetMode::Internal,
        )
        .expect("worksheet relationship");
    package.relate_to("xl/workbook.xml", rt::OFFICE_DOCUMENT);
    let bytes = PackageWriter::to_bytes(&package).expect("fixture package bytes");
    let package = SourceBackedPackage::from_vec(bytes).expect("source-backed package");
    SourceBackedEditor::from_source_backed_package(package).expect("source-backed editor")
}

#[test]
fn empty_and_effective_noop_commits_never_call_collector() {
    let editor = editor();
    let address = Address::from_a1("A1").expect("cell address");
    for (value, expected_changed, expected_counts) in [
        (None, false, (0, 0)),
        (Some(1u32), false, (0, 0)),
        (Some(2u32), true, (1, 1)),
    ] {
        COUNTS.with(|counts| counts.set((0, 0)));
        let mut edit = editor
            .edit_sheets(["Sheet1".into()])
            .expect("plan sheet edit");
        assert_eq!(COUNTS.with(Cell::get), (0, 0), "planning must not collect");
        if let Some(value) = value {
            edit.set("Sheet1", address, value)
                .expect("stage scalar value");
        }
        assert_eq!(COUNTS.with(Cell::get), (0, 0), "staging must not collect");
        let commit = edit.commit().expect("commit sheet edit");
        assert_eq!(commit.changed(), expected_changed);
        assert_eq!(COUNTS.with(Cell::get), expected_counts, "value {value:?}");
    }
}
