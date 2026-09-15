//! Change 0631 probe B: repeat open, edit and save `N` times on one fixture per
//! crate so a callgrind isolation pair at `N` and `N + M` can be differenced.
//!
//! The three repaired sites are on snapshot and patch paths, not on open or
//! publication, and the only added work is sorting a two- or three-element
//! array. This probe exists to show that with counts rather than to assert it.
//!
//! No fixture in this repository's corpus is admissible to the XLSX value-only
//! closure — every real workbook carries core-properties and extended-properties
//! package relationships, which `validate_package_relationships` refuses — so
//! the XLSX leg passes the literal fixture name `synthetic` and builds the
//! two-auxiliary-relationship workbook in process, once, before the loop.
//!
//! Usage: open_edit_save_iters <xlsx|pptx|docx> <fixture> <iterations>

use std::hint::black_box;
use std::path::Path;
use std::sync::Arc;

use litchi_core::{OwnedSource, Selector as CoreSelector};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{BlobPart, PackURI, TargetMode};

const SML: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

/// The smallest workbook the value-only closure admits that still owns both a
/// styles and a theme relationship: the shape `capture_auxiliary` captures and
/// `SourceState::same_owner` compares.
fn synthetic_workbook() -> Vec<u8> {
    let workbook_xml = format!(
        r#"<workbook xmlns="{SML}" xmlns:r="{REL}"><bookViews><workbookView/></bookViews><sheets><sheet name="Sheet1" sheetId="1" r:id="rIdSheet"/></sheets></workbook>"#
    );
    let sheet_xml = format!(
        r#"<worksheet xmlns="{SML}"><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>"#
    );
    let styles_xml =
        format!(r#"<styleSheet xmlns="{SML}"><cellXfs count="1"><xf/></cellXfs></styleSheet>"#);
    let mut package = litchi_opc::OpcPackage::new();
    for (name, content_type, blob) in [
        ("/xl/workbook.xml", ct::SML_SHEET_MAIN, workbook_xml.into_bytes()),
        ("/xl/worksheets/sheet1.xml", ct::SML_WORKSHEET, sheet_xml.into_bytes()),
        ("/xl/styles.xml", ct::SML_STYLES, styles_xml.into_bytes()),
        ("/xl/theme/theme1.xml", ct::OFC_THEME, b"<theme/>".to_vec()),
    ] {
        package
            .try_add_part(Box::new(BlobPart::new(
                PackURI::new(name).expect("part name"),
                content_type.to_owned(),
                blob,
            )))
            .expect("add part");
    }
    let workbook = package
        .get_part_mut(&PackURI::new("/xl/workbook.xml").expect("part name"))
        .expect("workbook")
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
            .expect("relationship");
    }
    package.relate_to("xl/workbook.xml", rt::OFFICE_DOCUMENT);
    litchi_opc::PackageWriter::to_bytes(&package).expect("publish")
}

/// Open a workbook twice — once source-backed for a value-only transaction and
/// once owning for a value-only snapshot — and serialize the result.
///
/// The two opens exercise both repaired captures: `capture_auxiliary_source`
/// for the source-backed leg and `capture_auxiliary` for the owning one. The
/// patch is deliberately not published here, because publishing it is the very
/// verdict the before leg gets wrong, and a round that refuses on one leg and
/// succeeds on the other would not be the same work.
fn xlsx_round(bytes: &[u8]) -> usize {
    use litchi_xlsx::cell_values::{Snapshot, SourceBackedEditor};
    let editor = SourceBackedEditor::from_read_at(Arc::new(OwnedSource::new(bytes.to_vec())))
        .expect("source-backed editor");
    let selector: litchi_xlsx::Selector<'_> = CoreSelector::Position(0.into());
    let commit = editor
        .edit_sheets([selector.clone()])
        .expect("value-only transaction")
        .commit()
        .expect("value-only commit");
    black_box(commit.patch().is_empty());
    let package = litchi_opc::OpcPackage::from_bytes(bytes).expect("owning open");
    let snapshot = Snapshot::load(&package, selector).expect("value-only snapshot");
    black_box(snapshot.source_xml().len());
    litchi_opc::PackageWriter::to_bytes(&package)
        .expect("publish")
        .len()
}

/// Open a presentation, capture and publish an exact no-op `ActiveX` control
/// transaction, and serialize the result.
///
/// The transaction is a no-op rather than a rename because the only corpus
/// fixture carrying controls has formatting whitespace in its slide XML, and a
/// rewritten part must be compact to be published; a no-op leaves the part
/// pristine. `load` calls the repaired `relationship_states` three times per
/// snapshot and `apply_patch` loads again, so the changed function is on the
/// measured path either way.
fn pptx_round(path: &Path) -> usize {
    use litchi_pptx::presentation::embedded::controls::{Limits, slide};
    let package = litchi_pptx::Package::open(path).expect("presentation open");
    let opc = package.opc().expect("opc");
    let slides = package
        .presentation()
        .expect("presentation")
        .slides()
        .expect("slides");
    let sheet = slides.first().expect("one slide");
    let part = sheet.part().part();
    let snapshot = slide::load(opc, 0, part, 0, &mut Limits::default()).expect("control snapshot");
    let commit = snapshot.edit().commit().expect("control commit");
    let mut target = opc.clone();
    slide::apply_patch(&mut target, commit.patch()).expect("control patch");
    litchi_opc::PackageWriter::to_bytes(&target)
        .expect("publish")
        .len()
}

/// Open a document, capture and publish an exact no-op content-control package
/// patch against an independent load, and serialize the result.
fn docx_round(path: &Path) -> usize {
    let source = litchi_docx::Package::open(path).expect("document open");
    let commit = source
        .content_control_snapshot()
        .expect("content-control snapshot")
        .edit()
        .expect("package transaction")
        .commit()
        .expect("package commit");
    let mut target = litchi_docx::Package::open(path).expect("target open");
    target
        .apply_content_controls(&commit)
        .expect("content-control publication");
    litchi_opc::PackageWriter::to_bytes(target.opc_package())
        .expect("publish")
        .len()
}

fn main() {
    let mut args = std::env::args().skip(1);
    let format = args.next().expect("format");
    let path = args.next().expect("fixture path");
    let iterations: usize = args.next().expect("iterations").parse().expect("iterations");
    let path = Path::new(&path);
    let bytes = if path.as_os_str() == "synthetic" {
        synthetic_workbook()
    } else {
        std::fs::read(path).expect("fixture bytes")
    };

    let mut total = 0usize;
    for _ in 0..iterations {
        total = total.wrapping_add(black_box(match format.as_str() {
            "xlsx" => xlsx_round(&bytes),
            "pptx" => pptx_round(path),
            "docx" => docx_round(path),
            other => panic!("unknown format {other}"),
        }));
    }
    println!("{format} {} iterations={iterations} total={total}", path.display());
}
