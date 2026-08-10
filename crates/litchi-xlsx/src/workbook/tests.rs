//! Behavioral tests for the immutable workbook facade and OPC graph boundary.

use std::io::{Cursor, Write};
use std::mem::size_of;
use std::path::Path;
use std::sync::Arc;

use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{BlobPart, OpcPackage, PackURI, Part, TargetMode};
use litchi_sheet::Rect;

use super::*;
use crate::Cell;
use crate::cell::{Extents, Value};
use crate::error::Error;
use crate::formula::Cache;
use crate::{LocalStyle, Package, ReadLimits, Style, Styles};

const CHARTSHEET_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet";
const CHARTSHEET_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml";

#[derive(Default)]
struct WriteOnly(Vec<u8>);

impl Write for WriteOnly {
    fn write(&mut self, bytes: &[u8]) -> std::io::Result<usize> {
        self.0.extend_from_slice(bytes);
        Ok(bytes.len())
    }

    fn flush(&mut self) -> std::io::Result<()> {
        Ok(())
    }
}

fn with_eocd_comment(mut archive: Vec<u8>, comment: &[u8]) -> Vec<u8> {
    let comment_len = u16::try_from(comment.len()).expect("ZIP comment fits in EOCD");
    let eocd = archive.len().checked_sub(22).expect("archive has an EOCD");
    assert_eq!(&archive[eocd..eocd + 4], b"PK\x05\x06");
    archive[eocd + 20..eocd + 22].copy_from_slice(&comment_len.to_le_bytes());
    archive.extend_from_slice(comment);
    archive
}

#[test]
fn owned_workbook_noop_output_preserves_exact_archive() {
    let source = with_eocd_comment(
        Workbook::new()
            .expect("valid workbook")
            .to_bytes()
            .expect("serialize workbook"),
        b"xlsx exact source",
    );
    let workbook = Workbook::from_bytes(source.clone()).expect("open owned workbook");

    assert_eq!(
        workbook.to_bytes().expect("serialize no-op workbook"),
        source
    );
}

#[test]
fn new_workbook_is_deterministic_and_selector_first() {
    let first = Workbook::new().expect("valid baseline");
    let second = Workbook::new().expect("valid baseline");

    assert_eq!(first.to_bytes().ok(), second.to_bytes().ok());

    let mut streamed = WriteOnly::default();
    first.write_to(&mut streamed).expect("stream workbook");
    assert_eq!(
        streamed.0,
        first.to_bytes().expect("buffered serialization")
    );
    let streamed = Workbook::from_slice(&streamed.0).expect("reopen streamed workbook");
    assert_eq!(streamed.len(), 1);
    assert_eq!(
        streamed
            .sheet("Sheet1")
            .expect("lookup")
            .expect("default sheet")
            .name(),
        "Sheet1"
    );
    assert_eq!(first.len(), 1);
    assert_eq!(first.flavor(), Flavor::Workbook);
    assert_eq!(first.date_system(), DateSystem::Excel1900);

    let by_name = first.sheet("sheet1").expect("lookup").expect("present");
    let by_position = first.sheet(0usize).expect("lookup").expect("present");
    let checked_name = crate::sheet::Name::new("SHEET1").expect("checked name");
    assert!(
        first
            .sheet(&checked_name)
            .expect("checked lookup")
            .is_some()
    );
    assert!(first.sheet(checked_name).expect("moved lookup").is_some());
    assert_eq!(by_name.name(), "Sheet1");
    assert_eq!(by_position.position(), 0);
    assert!(by_name.same_workbook(&by_position));
    assert!(matches!(by_name.kind(), WorksheetKind::Worksheet));
    assert!(matches!(by_name.visibility(), Visibility::Visible));
    assert!(first.sheet(1usize).expect("lookup").is_none());
    let extents = by_name.extents().expect("empty extents");
    assert_eq!(extents.declared().map(Rect::a1).as_deref(), Some("A1"));
    assert_eq!(extents.stored(), None);
    assert_eq!(extents.content(), None);
    assert_eq!(extents.styled(), None);

    let reopened = Workbook::from_bytes(first.to_bytes().expect("serialize"))
        .expect("reopen generated workbook");
    assert_eq!(
        reopened.active_sheet().map(|sheet| sheet.name().to_owned()),
        Some("Sheet1".into())
    );
}

#[test]
fn package_and_workbook_ingress_honor_exact_read_limits() {
    let bytes = Workbook::new()
        .expect("minimal workbook")
        .to_bytes()
        .expect("serialize minimal workbook");
    let input_bytes = u64::try_from(bytes.len()).expect("input length fits u64");
    let exact = ReadLimits::builder()
        .max_input_bytes(input_bytes)
        .expect("exact input limit")
        .build()
        .expect("valid exact limit");
    let over = ReadLimits::builder()
        .max_input_bytes(input_bytes - 1)
        .expect("smaller input limit")
        .build()
        .expect("valid smaller limit");
    let file = tempfile::NamedTempFile::new().expect("temporary workbook path");
    std::fs::write(file.path(), &bytes).expect("write workbook");

    assert!(Package::open(file.path()).is_ok());
    assert!(Package::from_bytes(bytes.clone()).is_ok());
    assert!(Package::from_slice(&bytes).is_ok());
    assert!(Package::from_reader(Cursor::new(bytes.clone())).is_ok());
    assert!(Workbook::open(file.path()).is_ok());
    assert!(Workbook::from_bytes(bytes.clone()).is_ok());
    assert!(Workbook::from_slice(&bytes).is_ok());
    assert!(Workbook::from_reader(Cursor::new(bytes.clone())).is_ok());

    assert!(Package::open_with_limits(file.path(), exact).is_ok());
    assert!(Package::from_bytes_with_limits(bytes.clone(), exact).is_ok());
    assert!(Package::from_slice_with_limits(&bytes, exact).is_ok());
    assert!(Package::from_reader_with_limits(Cursor::new(bytes.clone()), exact).is_ok());
    assert!(Workbook::open_with_limits(file.path(), exact).is_ok());
    assert!(Workbook::from_bytes_with_limits(bytes.clone(), exact).is_ok());
    assert!(Workbook::from_slice_with_limits(&bytes, exact).is_ok());
    assert!(Workbook::from_reader_with_limits(Cursor::new(bytes.clone()), exact).is_ok());

    assert!(Package::open_with_limits(file.path(), over).is_err());
    assert!(Package::from_bytes_with_limits(bytes.clone(), over).is_err());
    assert!(Package::from_slice_with_limits(&bytes, over).is_err());
    assert!(Package::from_reader_with_limits(Cursor::new(bytes.clone()), over).is_err());
    assert!(Workbook::open_with_limits(file.path(), over).is_err());
    assert!(Workbook::from_bytes_with_limits(bytes.clone(), over).is_err());
    assert!(Workbook::from_slice_with_limits(&bytes, over).is_err());
    assert!(Workbook::from_reader_with_limits(Cursor::new(bytes), over).is_err());
}

#[test]
fn clones_share_the_snapshot_and_handles_pin_it() {
    fn assert_send_sync<T: Send + Sync>() {}
    assert_send_sync::<Workbook>();
    assert_send_sync::<Worksheet>();
    assert_send_sync::<Style>();
    assert_send_sync::<Styles>();
    assert_send_sync::<crate::StyleKey>();
    assert_send_sync::<LocalStyle>();
    assert_send_sync::<Extents>();

    let workbook = Workbook::new().expect("valid baseline");
    let clone = workbook.clone();
    let sheet = workbook.active_sheet().expect("active sheet");
    let style = workbook
        .styles()
        .expect("styles")
        .base()
        .expect("base style");
    drop(workbook);

    assert_eq!(sheet.name(), "Sheet1");
    assert_eq!(style.fan_out().expect("fan-out"), 0);
    assert_eq!(
        clone.active_sheet().map(|sheet| sheet.name().to_owned()),
        Some("Sheet1".into())
    );
    assert!(size_of::<Workbook>() <= 2 * size_of::<usize>());
    assert!(size_of::<Style>() <= 2 * size_of::<usize>());
    assert!(size_of::<Styles>() <= 2 * size_of::<usize>());
}

#[test]
fn web_facades_are_lazy_borrowed_and_snapshot_scoped() {
    use litchi_ooxml_common::web::{
        AddIn, Conformance, Pane, Panes, Reference, Store as CatalogStore,
    };

    let baseline = Workbook::new().expect("valid baseline");
    assert!(baseline.task_panes().expect("load absent panes").is_none());
    assert!(baseline.task_panes().expect("reuse absent panes").is_none());

    let mut package = baseline.inner.package.clone();
    let reference = Reference::new("wa1", "1.0.0.0", CatalogStore::Omex).expect("valid reference");
    let mut panes = Panes::new();
    panes
        .push(Pane::new(
            AddIn::new("add-in-1", reference).expect("valid add-in"),
        ))
        .expect("valid pane");
    litchi_ooxml_common::web::put(&mut package, panes, Conformance::Transitional)
        .expect("install task panes");

    let with_panes = Workbook::from_package(package).expect("reopen task-pane workbook");
    std::thread::scope(|scope| {
        let reads = (0..8)
            .map(|_| {
                let workbook = with_panes.clone();
                scope.spawn(move || {
                    workbook
                        .task_panes()
                        .expect("load task panes concurrently")
                        .expect("task panes present")
                        .len()
                })
            })
            .collect::<Vec<_>>();
        for read in reads {
            assert_eq!(read.join().expect("task-pane reader"), 1);
        }
    });
    let first = with_panes
        .task_panes()
        .expect("load task panes")
        .expect("task panes present");
    let second = with_panes
        .task_panes()
        .expect("reuse task panes")
        .expect("task panes present");
    assert!(std::ptr::eq(first, second));
    assert_eq!(first.len(), 1);
    assert!(first.get("add-in-1").is_some());

    let mut package = baseline.inner.package.clone();
    let sheet_uri = PackURI::new("/xl/worksheets/sheet1.xml").expect("valid URI");
    package
        .get_part_mut(&sheet_uri)
        .expect("worksheet part")
        .set_blob(
            include_bytes!("../../../../test-data/ooxml/web_extensions/worksheet_bindings.xml")
                .to_vec(),
        );
    let with_bindings = Workbook::from_package(package).expect("reopen binding workbook");
    let sheet = with_bindings
        .sheet("Sheet1")
        .expect("lookup")
        .expect("worksheet present");
    std::thread::scope(|scope| {
        let reads = (0..8)
            .map(|_| {
                let sheet = sheet.clone();
                scope.spawn(move || {
                    sheet
                        .web_bindings()
                        .expect("load bindings concurrently")
                        .len()
                })
            })
            .collect::<Vec<_>>();
        for read in reads {
            assert_eq!(read.join().expect("web-binding reader"), 2);
        }
    });
    let first = sheet.web_bindings().expect("load bindings");
    let second = sheet.web_bindings().expect("reuse bindings");
    assert!(std::ptr::eq(first, second));
    assert_eq!(first.len(), 2);
    let first_binding = first.iter().next().expect("first binding");
    assert_eq!(first_binding.app_ref(), "sales-table");
}

#[test]
fn flavor_is_content_derived() {
    let mut package = OpcPackage::new();
    let workbook_uri = PackURI::new("/custom/main.xml").expect("valid URI");
    let worksheet_uri = PackURI::new("/custom/sheet.xml").expect("valid URI");
    let mut workbook = BlobPart::new(
        workbook_uri,
        ct::SML_TEMPLATE_MAIN.into(),
        br#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="8" state="veryHidden" r:id="rId1"/></sheets></workbook>"#.to_vec(),
    );
    workbook.relate_to("sheet.xml", rt::WORKSHEET);
    package.add_part(Box::new(workbook));
    package.add_part(Box::new(BlobPart::new(
        worksheet_uri,
        ct::SML_WORKSHEET.into(),
        br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>"#.to_vec(),
    )));
    package.relate_to("custom/main.xml", rt::OFFICE_DOCUMENT);

    let workbook = Workbook::from_package(package).expect("valid template");
    let sheet = workbook.sheet("Data").expect("lookup").expect("present");
    assert_eq!(workbook.flavor(), Flavor::Template);
    assert!(workbook.flavor().is_template());
    assert!(matches!(sheet.visibility(), Visibility::VeryHidden));
}

#[test]
fn duplicate_names_and_dangling_relationships_are_typed_errors() {
    let duplicate_xml = br#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/><sheet name="data" sheetId="2" r:id="rId2"/></sheets></workbook>"#;
    let mut package = package_with_workbook(duplicate_xml);
    let workbook_uri = PackURI::new("/xl/workbook.xml").expect("valid URI");
    let workbook = package.get_part_mut(&workbook_uri).expect("workbook part");
    workbook.relate_to("worksheets/sheet1.xml", rt::WORKSHEET);
    workbook.relate_to("worksheets/sheet2.xml", rt::WORKSHEET);
    for index in 1..=2 {
        package.add_part(Box::new(BlobPart::new(
            PackURI::new(format!("/xl/worksheets/sheet{index}.xml")).expect("valid URI"),
            ct::SML_WORKSHEET.into(),
            br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>"#.to_vec(),
        )));
    }
    assert!(matches!(
        Workbook::from_package(package),
        Err(Error::SheetNameConflict {
            first: 0,
            second: 1,
            ..
        })
    ));

    let dangling_xml = br#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Missing" sheetId="1" r:id="absent"/></sheets></workbook>"#;
    assert!(matches!(
        Workbook::from_package(package_with_workbook(dangling_xml)),
        Err(Error::Invalid(_))
    ));

    let aliased_xml = br#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="One" sheetId="1" r:id="rId1"/><sheet name="Two" sheetId="2" r:id="rId2"/></sheets></workbook>"#;
    let mut package = package_with_workbook(aliased_xml);
    let workbook_uri = PackURI::new("/xl/workbook.xml").expect("valid URI");
    let workbook = package.get_part_mut(&workbook_uri).expect("workbook part");
    for id in ["rId1", "rId2"] {
        workbook
            .rels_mut()
            .try_add_relationship(
                rt::WORKSHEET.to_owned(),
                "worksheets/sheet1.xml".to_owned(),
                id.to_owned(),
                TargetMode::Internal,
            )
            .expect("sheet relationship");
    }
    package.add_part(Box::new(BlobPart::new(
        PackURI::new("/xl/worksheets/sheet1.xml").expect("valid URI"),
        ct::SML_WORKSHEET.into(),
        br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>"#.to_vec(),
    )));
    assert!(matches!(
        Workbook::from_package(package),
        Err(Error::Invalid(message)) if message.contains("referenced by both 'One' and 'Two'")
    ));
}

#[test]
fn styles_graph_table_and_cell_references_are_checked() {
    let baseline = Workbook::new().expect("baseline");

    let mut duplicate = baseline.inner.package.clone();
    duplicate
        .get_part_mut(&baseline.inner.workbook_uri)
        .expect("workbook part")
        .rels_mut()
        .try_add_relationship(
            rt::STYLES.into(),
            "styles.xml".into(),
            "rId3".into(),
            TargetMode::Internal,
        )
        .expect("second styles relationship");
    assert!(matches!(
        Workbook::from_package(duplicate),
        Err(Error::Invalid(message)) if message.contains("multiple styles relationships")
    ));

    let mut external = baseline.inner.package.clone();
    let rels = external
        .get_part_mut(&baseline.inner.workbook_uri)
        .expect("workbook part")
        .rels_mut();
    rels.remove("rId2").expect("styles relationship");
    rels.try_add_relationship(
        rt::STYLES.into(),
        "https://example.invalid/styles.xml".into(),
        "rId2".into(),
        TargetMode::External,
    )
    .expect("external styles relationship");
    assert!(matches!(
        Workbook::from_package(external),
        Err(Error::Invalid(message)) if message.contains("cannot be external")
    ));

    let styles_uri = PackURI::new("/xl/styles.xml").expect("styles URI");
    let mut wrong_type = baseline.inner.package.clone();
    wrong_type
        .get_part_mut(&styles_uri)
        .expect("styles part")
        .set_content_type("application/xml".into())
        .expect("replace content type");
    assert!(matches!(
        Workbook::from_package(wrong_type),
        Err(Error::Invalid(message)) if message.contains("styles part has content type")
    ));

    let mut malformed = baseline.inner.package.clone();
    malformed
        .get_part_mut(&styles_uri)
        .expect("styles part")
        .set_blob(
            br#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cellXfs count="2"><xf/></cellXfs></styleSheet>"#.to_vec(),
        );
    let malformed = Workbook::from_package(malformed).expect("graph remains lazy");
    assert!(matches!(malformed.styles(), Err(Error::Invalid(_))));

    let mut dangling = baseline.inner.package.clone();
    dangling
        .get_part_mut(&PackURI::new("/xl/worksheets/sheet1.xml").expect("sheet URI"))
        .expect("sheet part")
        .set_blob(
            br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" s="1"/></row></sheetData></worksheet>"#.to_vec(),
        );
    let dangling = Workbook::from_package(dangling).expect("lazy worksheet");
    assert!(matches!(
        dangling
            .sheet(0usize)
            .expect("lookup")
            .expect("sheet")
            .cell("A1"),
        Err(Error::Invalid(message)) if message.contains("A1 references shared style 1")
    ));

    let mut dangling_column = baseline.inner.package.clone();
    dangling_column
        .get_part_mut(&PackURI::new("/xl/worksheets/sheet1.xml").expect("sheet URI"))
        .expect("sheet part")
        .set_blob(
            br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cols><col min="2" max="2" style="1"/></cols><sheetData/></worksheet>"#.to_vec(),
        );
    let dangling_column = Workbook::from_package(dangling_column).expect("lazy column style");
    assert!(matches!(
        dangling_column
            .sheet(0usize)
            .expect("lookup")
            .expect("sheet")
            .column(1),
        Err(Error::Invalid(message)) if message.contains("column 1 references shared style 1")
    ));

    let mut dangling_row = baseline.inner.package.clone();
    dangling_row
        .get_part_mut(&PackURI::new("/xl/worksheets/sheet1.xml").expect("sheet URI"))
        .expect("sheet part")
        .set_blob(
            br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="2" s="1" customFormat="1"/></sheetData></worksheet>"#.to_vec(),
        );
    let dangling_row = Workbook::from_package(dangling_row).expect("lazy row style");
    assert!(matches!(
        dangling_row
            .sheet(0usize)
            .expect("lookup")
            .expect("sheet")
            .row(1),
        Err(Error::Invalid(message)) if message.contains("row 1 references shared style 1")
    ));
}

#[test]
fn concurrent_snapshot_reads_need_no_public_locking() {
    let workbook = Workbook::new().expect("valid baseline");
    std::thread::scope(|scope| {
        for _ in 0..8 {
            let workbook = workbook.clone();
            scope.spawn(move || {
                for _ in 0..1_000 {
                    let sheet = workbook.sheet("Sheet1").expect("lookup").expect("present");
                    assert_eq!(sheet.position(), 0);
                }
            });
        }
    });
}

#[test]
fn cell_facade_is_sparse_exact_and_non_mutating() {
    let workbook = Workbook::from_package(package_with_cells()).expect("valid workbook");
    let bytes_before = workbook.to_bytes().expect("serialize before lazy read");
    let sheet = workbook.sheet("data").expect("lookup").expect("present");

    assert!(matches!(
        sheet.cell("A1").expect("cell lookup").stored(),
        Some(Cell::Value(Value::Text(text))) if text.as_str() == "Office & Litchi"
    ));
    assert!(sheet.cell((0, 1)).expect("missing lookup").is_missing());
    assert!(matches!(
        sheet.cell((1, 2)).expect("number lookup").stored(),
        Some(Cell::Value(Value::Number(number))) if number.as_str() == "-0.000"
    ));
    let Some(Cell::Formula(formula)) = sheet.cell((2, 1)).expect("formula lookup").stored() else {
        panic!("expected formula cell")
    };
    assert_eq!(formula.text(), "C2*2");
    assert!(matches!(
        formula.cached().map(Cache::value),
        Some(Value::Number(number)) if number.as_str() == "0"
    ));
    assert!(matches!(
        sheet.cell((4, 3)).expect("empty lookup").stored(),
        Some(Cell::Empty)
    ));
    assert!(matches!(
        sheet.cell((litchi_sheet::ROWS, 0)),
        Err(Error::Coordinate(_))
    ));
    assert!(sheet.row(0).expect("stored row 1").stored());
    assert!(!sheet.row(3).expect("implicit row 4").stored());
    assert!(!sheet.row(4).expect("stored row 5").hidden());
    assert!(matches!(
        sheet.row(litchi_sheet::ROWS),
        Err(Error::Coordinate(_))
    ));
    assert_eq!(
        sheet
            .rows()
            .expect("stored rows")
            .map(|row| row.index().get())
            .collect::<Vec<_>>(),
        [0, 1, 2, 4]
    );

    let addresses = sheet
        .cells("B1:D4")
        .expect("sparse traversal")
        .map(|(address, _)| (address.row().get(), address.column().get()))
        .collect::<Vec<_>>();
    assert_eq!(addresses, [(1, 2), (2, 1)]);
    assert!(matches!(sheet.cells("B2:A1"), Err(Error::Range(_))));
    let extents = sheet.extents().expect("cell extents");
    assert_eq!(extents.declared(), None);
    assert_eq!(extents.stored().map(Rect::a1).as_deref(), Some("A1:D5"));
    assert_eq!(extents.content().map(Rect::a1).as_deref(), Some("A1:C3"));
    assert_eq!(extents.styled().map(Rect::a1).as_deref(), Some("D5"));
    assert_eq!(extents.used().map(Rect::a1).as_deref(), Some("A1:D5"));
    assert_eq!(
        sheet.stored_extent().expect("extent").map(Rect::end),
        Some((5, 4))
    );
    assert_eq!(
        workbook.to_bytes().expect("serialize after lazy read"),
        bytes_before
    );
}

#[test]
fn concurrent_first_cell_read_publishes_one_safe_snapshot() {
    let workbook = Workbook::from_package(package_with_cells()).expect("valid workbook");
    let barrier = Arc::new(std::sync::Barrier::new(8));
    std::thread::scope(|scope| {
        for _ in 0..8 {
            let workbook = workbook.clone();
            let barrier = Arc::clone(&barrier);
            scope.spawn(move || {
                barrier.wait();
                let sheet = workbook.sheet("Data").expect("lookup").expect("present");
                assert!(matches!(
                    sheet.cell((0, 0)).expect("cell lookup").stored(),
                    Some(Cell::Value(Value::Text(text))) if text.as_str() == "Office & Litchi"
                ));
            });
        }
    });
}

#[test]
fn worksheet_operations_reject_other_sheet_kinds_without_parsing_them() {
    let mut package = OpcPackage::new();
    let workbook_uri = PackURI::new("/xl/workbook.xml").expect("valid URI");
    let chart_uri = PackURI::new("/xl/chartsheets/sheet1.xml").expect("valid URI");
    let mut workbook = BlobPart::new(
        workbook_uri,
        ct::SML_SHEET_MAIN.into(),
        br#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Chart" sheetId="1" r:id="rId1"/></sheets></workbook>"#.to_vec(),
    );
    workbook
        .rels_mut()
        .try_add_relationship(
            CHARTSHEET_REL.into(),
            "chartsheets/sheet1.xml".into(),
            "rId1".into(),
            TargetMode::Internal,
        )
        .expect("chart relationship");
    package.add_part(Box::new(workbook));
    package.add_part(Box::new(BlobPart::new(
        chart_uri,
        CHARTSHEET_CONTENT_TYPE.into(),
        b"not parsed by a worksheet operation".to_vec(),
    )));
    package.relate_to("xl/workbook.xml", rt::OFFICE_DOCUMENT);

    let workbook = Workbook::from_package(package).expect("valid chart graph");
    let chart = workbook.sheet("Chart").expect("lookup").expect("present");
    assert!(matches!(
        chart.cell((0, 0)),
        Err(Error::NotWorksheet { .. })
    ));
}

#[test]
fn poi_and_libreoffice_shared_formula_oracles_match() {
    let root = Path::new(env!("CARGO_MANIFEST_DIR")).join("../..");
    let cases = [
        (
            root.join("test-data/poi/test-data/spreadsheet/shared_formulas.xlsx"),
            (40, 0),
            "B41",
        ),
        (
            root.join("test-data/libreoffice-core/sc/qa/unit/data/xlsx/shared-formula/basic.xlsx"),
            (18, 1),
            "A19*10",
        ),
    ];
    for (path, address, expected) in cases {
        if !path.exists() {
            continue;
        }
        let workbook = Workbook::open(path).expect("corpus workbook");
        let sheet = workbook.sheet(0usize).expect("lookup").expect("present");
        let Some(Cell::Formula(formula)) = sheet.cell(address).expect("formula lookup").stored()
        else {
            panic!("expected formula at {address:?}")
        };
        assert_eq!(formula.text(), expected);
    }
}

fn package_with_workbook(xml: &[u8]) -> OpcPackage {
    let mut package = OpcPackage::new();
    package.add_part(Box::new(BlobPart::new(
        PackURI::new("/xl/workbook.xml").expect("valid URI"),
        ct::SML_SHEET_MAIN.into(),
        xml.to_vec(),
    )));
    package.relate_to("xl/workbook.xml", rt::OFFICE_DOCUMENT);
    package
}

fn package_with_cells() -> OpcPackage {
    let mut package = OpcPackage::new();
    let workbook_uri = PackURI::new("/xl/workbook.xml").expect("valid URI");
    let worksheet_uri = PackURI::new("/xl/worksheets/sheet1.xml").expect("valid URI");
    let strings_uri = PackURI::new("/xl/sharedStrings.xml").expect("valid URI");
    let styles_uri = PackURI::new("/xl/styles.xml").expect("valid URI");
    let mut workbook = BlobPart::new(
        workbook_uri,
        ct::SML_SHEET_MAIN.into(),
        br#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>"#.to_vec(),
    );
    workbook
        .rels_mut()
        .try_add_relationship(
            rt::WORKSHEET.into(),
            "worksheets/sheet1.xml".into(),
            "rId1".into(),
            TargetMode::Internal,
        )
        .expect("worksheet relationship");
    workbook
        .rels_mut()
        .try_add_relationship(
            rt::SHARED_STRINGS.into(),
            "sharedStrings.xml".into(),
            "rId2".into(),
            TargetMode::Internal,
        )
        .expect("shared-string relationship");
    workbook
        .rels_mut()
        .try_add_relationship(
            rt::STYLES.into(),
            "styles.xml".into(),
            "rId3".into(),
            TargetMode::Internal,
        )
        .expect("styles relationship");
    package.add_part(Box::new(workbook));
    package.add_part(Box::new(BlobPart::new(
        worksheet_uri,
        ct::SML_WORKSHEET.into(),
        br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c></row><row r="2"><c r="C2"><v>-0.000</v></c></row><row r="3"><c r="B3"><f>C2*2</f><v>0</v></c></row><row r="5"><c r="D5" s="2"/></row></sheetData></worksheet>"#.to_vec(),
    )));
    package.add_part(Box::new(BlobPart::new(
        strings_uri,
        ct::SML_SHARED_STRINGS.into(),
        br#"<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1"><si><r><t>Office &amp; </t></r><r><t>Litchi</t></r></si></sst>"#.to_vec(),
    )));
    package.add_part(Box::new(BlobPart::new(
        styles_uri,
        ct::SML_STYLES.into(),
        br#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cellXfs count="3"><xf/><xf numFmtId="1"/><xf numFmtId="2"/></cellXfs></styleSheet>"#.to_vec(),
    )));
    package.relate_to("xl/workbook.xml", rt::OFFICE_DOCUMENT);
    package
}
