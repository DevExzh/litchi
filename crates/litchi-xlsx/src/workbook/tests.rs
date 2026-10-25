//! Behavioral tests for the immutable workbook facade and OPC graph boundary.

use std::io::{Cursor, Write};
use std::mem::size_of;
use std::path::Path;
use std::sync::Arc;

use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{BlobPart, OpcPackage, PackURI, Part, TargetMode};
use litchi_sheet::Rect;
use soapberry_zip::office::StreamingArchiveWriter;

use super::*;
use crate::Cell;
use crate::cell::{Extents, Value};
use crate::error::Error;
use crate::formula::Cache;
use crate::{LocalStyle, Package, SourceLimits, Style, Styles};

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
        Package::new()
            .expect("valid workbook")
            .to_bytes()
            .expect("serialize workbook"),
        b"xlsx exact source",
    );
    let workbook = Package::from_bytes(&source).expect("open owned workbook");

    assert_eq!(
        workbook.to_bytes().expect("serialize no-op workbook"),
        source
    );
}

#[test]
fn new_workbook_is_deterministic_and_selector_first() {
    let first = Package::new().expect("valid baseline");
    let second = Package::new().expect("valid baseline");

    assert_eq!(first.to_bytes().ok(), second.to_bytes().ok());

    let mut streamed = WriteOnly::default();
    first.write_to(&mut streamed).expect("stream workbook");
    assert_eq!(
        streamed.0,
        first.to_bytes().expect("buffered serialization")
    );
    let streamed = Package::from_bytes(&streamed.0).expect("reopen streamed workbook");
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

    let reopened = Package::from_bytes(&first.to_bytes().expect("serialize"))
        .expect("reopen generated workbook");
    assert_eq!(
        reopened.active_sheet().map(|sheet| sheet.name().to_owned()),
        Some("Sheet1".into())
    );
}

#[test]
fn same_content_ignores_zip_framing_when_native_graph_matches() {
    let bytes = Package::new()
        .expect("valid workbook")
        .to_bytes()
        .expect("serialize workbook");
    let left = Package::from_bytes(&bytes).expect("left workbook");
    let right = Package::from_bytes(&with_eocd_comment(bytes, b"different ZIP comment"))
        .expect("right workbook");

    assert_ne!(
        left.inner.package.exact_source_sha256(),
        right.inner.package.exact_source_sha256(),
        "the test must exercise the native graph comparison"
    );
    assert!(left.same_content(&right).expect("compare native graph"));
}

#[test]
fn same_content_uses_shared_source_proof_for_opaque_members() {
    let source =
        Package::from_bytes(&workbook_with_opaque_member()).expect("workbook with opaque member");
    assert_eq!(source.inner.package.non_part_members().len(), 1);

    // Replacing a part with the same bytes revokes exact-source publication,
    // while the cloned package still shares the immutable source archive and
    // therefore retains a proof for the opaque member's bytes.
    let worksheet = PackURI::new("/xl/worksheets/sheet1.xml").expect("worksheet URI");
    let mut package = source.inner.package.clone();
    let bytes = package
        .get_part(&worksheet)
        .expect("worksheet part")
        .blob()
        .to_vec();
    package
        .get_part_mut(&worksheet)
        .expect("mutable worksheet part")
        .set_blob(bytes);
    let target = Package::from_opc(package).expect("target workbook");

    assert!(!target.inner.package.is_unmodified_owned_source());
    assert!(source.same_content(&target).expect("compare opaque graph"));
}

#[test]
fn same_content_refuses_unshared_opaque_members() {
    let left = Package::from_bytes(&workbook_with_opaque_member()).expect("left workbook");
    let right = Package::from_bytes(&workbook_with_opaque_payload(b"different opaque bytes"))
        .expect("right workbook");

    assert!(matches!(
        left.same_content(&right),
        Err(Error::Unsupported {
            feature: "content comparison of opaque OPC members"
        })
    ));
}

#[test]
fn same_content_rejects_unmodeled_publication_manifests() {
    let manifest_left =
        Package::from_bytes(&workbook_with_unmodeled_metadata("one", "same", "same"))
            .expect("manifest-left workbook");
    let manifest_right =
        Package::from_bytes(&workbook_with_unmodeled_metadata("two", "same", "same"))
            .expect("manifest-right workbook");
    assert!(
        !manifest_left
            .same_content(&manifest_right)
            .expect("compare content-type manifests")
    );

    let relationship_left =
        Package::from_bytes(&workbook_with_unmodeled_metadata("same", "one", "same"))
            .expect("relationship-left workbook");
    let relationship_right =
        Package::from_bytes(&workbook_with_unmodeled_metadata("same", "two", "same"))
            .expect("relationship-right workbook");
    assert!(
        !relationship_left
            .same_content(&relationship_right)
            .expect("compare relationship manifests")
    );

    let part_relationship_left =
        Package::from_bytes(&workbook_with_unmodeled_metadata("same", "same", "one"))
            .expect("part-relationship-left workbook");
    let part_relationship_right =
        Package::from_bytes(&workbook_with_unmodeled_metadata("same", "same", "two"))
            .expect("part-relationship-right workbook");
    assert!(
        !part_relationship_left
            .same_content(&part_relationship_right)
            .expect("compare part relationship manifests")
    );
}

#[test]
fn patch_exposes_endpoints_and_conservative_retention() {
    let source = Package::new().expect("source workbook");
    let mut edit = source.edit().expect("start edit");
    edit.sheet("Sheet1")
        .expect("sheet lookup")
        .expect("Sheet1")
        .set("A1", "changed")
        .expect("set cell");
    let commit = edit.commit().expect("commit edit");
    let patch = commit.patch().clone();

    assert!(
        patch
            .source()
            .expect("patch source")
            .same_content(&source)
            .unwrap()
    );
    assert!(!patch.is_noop());
    assert!(patch.retained_bytes() > 0);
    assert!(Patch::default().source().is_err());
    assert_eq!(Patch::default().retained_bytes(), 0);
}

fn workbook_with_opaque_member() -> Vec<u8> {
    workbook_with_opaque_payload(b"opaque member bytes")
}

fn workbook_with_opaque_payload(payload: &[u8]) -> Vec<u8> {
    const MAIN: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    const PACKAGE_REL: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
    let mut writer =
        StreamingArchiveWriter::new(soapberry_zip::office::StreamingArchiveLimits::default());
    writer
        .write_stored(
            "[Content_Types].xml",
            format!(
                r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="{workbook}"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="{worksheet}"/></Types>"#,
                workbook = ct::SML_SHEET_MAIN,
                worksheet = ct::SML_WORKSHEET,
            )
            .as_bytes(),
        )
        .expect("content types");
    writer
        .write_stored(
            "_rels/.rels",
            format!(
                r#"<Relationships xmlns="{PACKAGE_REL}"><Relationship Id="rId1" Type="{REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>"#
            )
            .as_bytes(),
        )
        .expect("package relationships");
    writer
        .write_stored(
            "xl/workbook.xml",
            format!(
                r#"<workbook xmlns="{MAIN}" xmlns:r="{REL}"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>"#
            )
            .as_bytes(),
        )
        .expect("workbook");
    writer
        .write_stored(
            "xl/_rels/workbook.xml.rels",
            format!(
                r#"<Relationships xmlns="{PACKAGE_REL}"><Relationship Id="rId1" Type="{REL}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#
            )
            .as_bytes(),
        )
        .expect("workbook relationships");
    writer
        .write_stored(
            "xl/worksheets/sheet1.xml",
            format!(r#"<worksheet xmlns="{MAIN}"><sheetData/></worksheet>"#).as_bytes(),
        )
        .expect("worksheet");
    writer
        .write_stored("opaque.bin", payload)
        .expect("opaque member");
    writer.finish_to_bytes().expect("finish workbook")
}

fn workbook_with_unmodeled_metadata(
    content_marker: &str,
    relationship_marker: &str,
    part_relationship_marker: &str,
) -> Vec<u8> {
    const MAIN: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    const PACKAGE_REL: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
    let mut writer =
        StreamingArchiveWriter::new(soapberry_zip::office::StreamingArchiveLimits::default());
    writer
        .write_stored(
            "[Content_Types].xml",
            format!(
                r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types" xmlns:x="urn:litchi:test"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml" x:marker="{content_marker}"/><Override PartName="/xl/workbook.xml" ContentType="{workbook}"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="{worksheet}"/></Types>"#,
                workbook = ct::SML_SHEET_MAIN,
                worksheet = ct::SML_WORKSHEET,
            )
            .as_bytes(),
        )
        .expect("content types");
    writer
        .write_stored(
            "_rels/.rels",
            format!(
                r#"<Relationships xmlns="{PACKAGE_REL}" xmlns:x="urn:litchi:test"><Relationship Id="rId1" Type="{REL}/officeDocument" Target="xl/workbook.xml" x:marker="{relationship_marker}"/></Relationships>"#
            )
            .as_bytes(),
        )
        .expect("package relationships");
    writer
        .write_stored(
            "xl/workbook.xml",
            format!(
                r#"<workbook xmlns="{MAIN}" xmlns:r="{REL}"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>"#
            )
            .as_bytes(),
        )
        .expect("workbook");
    writer
        .write_stored(
            "xl/_rels/workbook.xml.rels",
            format!(
                r#"<Relationships xmlns="{PACKAGE_REL}" xmlns:x="urn:litchi:test"><Relationship Id="rId1" Type="{REL}/worksheet" Target="worksheets/sheet1.xml" x:marker="{part_relationship_marker}"/></Relationships>"#
            )
            .as_bytes(),
        )
        .expect("workbook relationships");
    writer
        .write_stored(
            "xl/worksheets/sheet1.xml",
            format!(r#"<worksheet xmlns="{MAIN}"><sheetData/></worksheet>"#).as_bytes(),
        )
        .expect("worksheet");
    writer.finish_to_bytes().expect("finish workbook")
}

#[test]
fn package_and_workbook_ingress_honor_exact_read_limits() {
    let bytes = Package::new()
        .expect("minimal workbook")
        .to_bytes()
        .expect("serialize minimal workbook");
    let input_bytes = u64::try_from(bytes.len()).expect("input length fits u64");
    let exact = SourceLimits::builder()
        .max_input_bytes(input_bytes)
        .expect("exact input limit")
        .build()
        .expect("valid exact limit");
    let over = SourceLimits::builder()
        .max_input_bytes(input_bytes - 1)
        .expect("smaller input limit")
        .build()
        .expect("valid smaller limit");
    let file = tempfile::NamedTempFile::new().expect("temporary workbook path");
    std::fs::write(file.path(), &bytes).expect("write workbook");

    assert!(Package::open(file.path()).is_ok());
    assert!(Package::from_bytes(&bytes).is_ok());
    assert!(Package::from_bytes(&bytes).is_ok());
    assert!(Package::from_reader(Cursor::new(bytes.clone())).is_ok());
    assert!(Package::open(file.path()).is_ok());
    assert!(Package::from_bytes(&bytes).is_ok());
    assert!(Package::from_bytes(&bytes).is_ok());
    assert!(Package::from_reader(Cursor::new(bytes.clone())).is_ok());

    assert!(
        Package::open_with_options(file.path(), crate::OpenOptions::new().source_limits(exact))
            .is_ok()
    );
    assert!(
        Package::from_bytes_with_options(&bytes, crate::OpenOptions::new().source_limits(exact))
            .is_ok()
    );
    assert!(
        Package::from_bytes_with_options(&bytes, crate::OpenOptions::new().source_limits(exact))
            .is_ok()
    );
    assert!(
        Package::from_reader_with_options(
            Cursor::new(bytes.clone()),
            crate::OpenOptions::new().source_limits(exact)
        )
        .is_ok()
    );
    assert!(
        Package::open_with_options(file.path(), crate::OpenOptions::new().source_limits(exact))
            .is_ok()
    );
    assert!(
        Package::from_bytes_with_options(&bytes, crate::OpenOptions::new().source_limits(exact))
            .is_ok()
    );
    assert!(
        Package::from_bytes_with_options(&bytes, crate::OpenOptions::new().source_limits(exact))
            .is_ok()
    );
    assert!(
        Package::from_reader_with_options(
            Cursor::new(bytes.clone()),
            crate::OpenOptions::new().source_limits(exact)
        )
        .is_ok()
    );

    assert!(
        Package::open_with_options(file.path(), crate::OpenOptions::new().source_limits(over))
            .is_err()
    );
    assert!(
        Package::from_bytes_with_options(&bytes, crate::OpenOptions::new().source_limits(over))
            .is_err()
    );
    assert!(
        Package::from_bytes_with_options(&bytes, crate::OpenOptions::new().source_limits(over))
            .is_err()
    );
    assert!(
        Package::from_reader_with_options(
            Cursor::new(bytes.clone()),
            crate::OpenOptions::new().source_limits(over)
        )
        .is_err()
    );
    assert!(
        Package::open_with_options(file.path(), crate::OpenOptions::new().source_limits(over))
            .is_err()
    );
    assert!(
        Package::from_bytes_with_options(&bytes, crate::OpenOptions::new().source_limits(over))
            .is_err()
    );
    assert!(
        Package::from_bytes_with_options(&bytes, crate::OpenOptions::new().source_limits(over))
            .is_err()
    );
    assert!(
        Package::from_reader_with_options(
            Cursor::new(bytes),
            crate::OpenOptions::new().source_limits(over)
        )
        .is_err()
    );
}

#[test]
fn derived_owned_edit_keeps_the_physical_part_limit() {
    let baseline = Package::new().expect("minimal workbook");
    let baseline_bytes = baseline.to_bytes().expect("serialize minimal workbook");
    let maximum_part = baseline
        .opc()
        .try_iter_parts()
        .map(|part| part.expect("baseline part").blob().len())
        .max()
        .expect("baseline parts");
    let limits = SourceLimits::builder()
        .max_part_bytes(maximum_part as u64)
        .expect("part limit")
        .build()
        .expect("source limits");
    let package = Package::from_bytes_with_options(
        &baseline_bytes,
        crate::OpenOptions::new().source_limits(limits),
    )
    .expect("limited workbook");
    assert_eq!(package.opc().read_limits(), limits);
    let before = package.to_bytes().expect("serialize source");

    let mut edit = package.edit().expect("start edit");
    edit.sheet(0usize)
        .expect("lookup")
        .expect("worksheet")
        .set("A1", "x".repeat(maximum_part))
        .expect("stage oversized cell");
    let committed = edit.commit().expect("derived edit remains valid");
    assert_eq!(committed.package().opc().read_limits(), limits);
    assert!(
        Package::from_bytes_with_options(
            &committed
                .package()
                .to_bytes()
                .expect("serialize derived edit"),
            crate::OpenOptions::new().source_limits(limits),
        )
        .is_err()
    );
    assert_eq!(package.to_bytes().expect("source remains readable"), before);
    assert_eq!(package.opc().read_limits(), limits);
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

    let workbook = Package::new().expect("valid baseline");
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

    let baseline = Package::new().expect("valid baseline");
    assert!(
        baseline
            .workbook_ref()
            .task_panes()
            .expect("load absent panes")
            .is_none()
    );
    assert!(
        baseline
            .workbook_ref()
            .task_panes()
            .expect("reuse absent panes")
            .is_none()
    );

    let mut package = baseline.inner.package.clone();
    let reference = Reference::new("wa1", "1.0.0.0", CatalogStore::Omex).expect("valid reference");
    let mut panes = Panes::new();
    panes
        .push(Pane::new(
            AddIn::new("add-in-1", reference).expect("valid add-in"),
        ))
        .expect("valid pane");
    let mut edit = litchi_ooxml_common::web::Edit::new(&package).expect("task-pane edit");
    edit.put(panes, Conformance::Transitional)
        .expect("stage task panes");
    package = edit.commit().expect("install task panes").into_package();

    let with_panes = Package::from_opc(package).expect("reopen task-pane workbook");
    std::thread::scope(|scope| {
        let reads = (0..8)
            .map(|_| {
                let workbook = with_panes.clone();
                scope.spawn(move || {
                    workbook
                        .workbook_ref()
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
        .workbook_ref()
        .task_panes()
        .expect("load task panes")
        .expect("task panes present");
    let second = with_panes
        .workbook_ref()
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
    let with_bindings = Package::from_opc(package).expect("reopen binding workbook");
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

    let workbook = Package::from_opc(package).expect("valid template");
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
        Package::from_opc(package),
        Err(Error::SheetNameConflict {
            first: 0,
            second: 1,
            ..
        })
    ));

    let dangling_xml = br#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Missing" sheetId="1" r:id="absent"/></sheets></workbook>"#;
    assert!(matches!(
        Package::from_opc(package_with_workbook(dangling_xml)),
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
        Package::from_opc(package),
        Err(Error::Invalid(message)) if message.contains("referenced by both 'One' and 'Two'")
    ));
}

#[test]
fn styles_graph_table_and_cell_references_are_checked() {
    let baseline = Package::new().expect("baseline");

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
        Package::from_opc(duplicate),
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
        Package::from_opc(external),
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
        Package::from_opc(wrong_type),
        Err(Error::Invalid(message)) if message.contains("styles part has content type")
    ));

    let mut malformed = baseline.inner.package.clone();
    malformed
        .get_part_mut(&styles_uri)
        .expect("styles part")
        .set_blob(
            br#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cellXfs count="2"><xf/></cellXfs></styleSheet>"#.to_vec(),
        );
    let malformed = Package::from_opc(malformed).expect("graph remains lazy");
    assert!(matches!(malformed.styles(), Err(Error::Invalid(_))));

    let mut dangling = baseline.inner.package.clone();
    dangling
        .get_part_mut(&PackURI::new("/xl/worksheets/sheet1.xml").expect("sheet URI"))
        .expect("sheet part")
        .set_blob(
            br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" s="1"/></row></sheetData></worksheet>"#.to_vec(),
        );
    // Audit finding 4.2: a stale style index no longer kills the sheet. It
    // reads with the default format (cellXfs 0) and the repair is recorded.
    let dangling = Package::from_opc(dangling).expect("lazy worksheet");
    let sheet = dangling.sheet(0usize).expect("lookup").expect("sheet");
    sheet.cell("A1").expect("stale cell style is repaired");
    assert_eq!(sheet.store().expect("store").entries()[0].style, Some(0));
    assert_style_repair(
        &dangling,
        &sheet,
        crate::ToleranceLocation::Cell(litchi_sheet::Cell::from_a1("A1").expect("A1")),
    );

    let mut dangling_column = baseline.inner.package.clone();
    dangling_column
        .get_part_mut(&PackURI::new("/xl/worksheets/sheet1.xml").expect("sheet URI"))
        .expect("sheet part")
        .set_blob(
            br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cols><col min="2" max="2" style="1"/></cols><sheetData/></worksheet>"#.to_vec(),
        );
    let dangling_column = Package::from_opc(dangling_column).expect("lazy column style");
    let sheet = dangling_column
        .sheet(0usize)
        .expect("lookup")
        .expect("sheet");
    sheet.column(1).expect("stale column style is repaired");
    assert_eq!(
        sheet.store().expect("store").column_entries()[0]
            .properties
            .style,
        Some(0)
    );
    let column = litchi_sheet::Column::new(1).expect("column B");
    assert_style_repair(
        &dangling_column,
        &sheet,
        crate::ToleranceLocation::Columns {
            first: column,
            last: column,
        },
    );

    let mut dangling_row = baseline.inner.package.clone();
    dangling_row
        .get_part_mut(&PackURI::new("/xl/worksheets/sheet1.xml").expect("sheet URI"))
        .expect("sheet part")
        .set_blob(
            br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="2" s="1" customFormat="1"/></sheetData></worksheet>"#.to_vec(),
        );
    let dangling_row = Package::from_opc(dangling_row).expect("lazy row style");
    let sheet = dangling_row.sheet(0usize).expect("lookup").expect("sheet");
    sheet.row(1).expect("stale row style is repaired");
    assert_eq!(
        sheet.store().expect("store").row_entries()[0]
            .properties
            .style,
        Some(0)
    );
    assert_style_repair(
        &dangling_row,
        &sheet,
        crate::ToleranceLocation::Row(litchi_sheet::Row::new(1).expect("row 2")),
    );
}

fn assert_style_repair(workbook: &Package, sheet: &Worksheet, location: crate::ToleranceLocation) {
    let defect = crate::Defect::Cell(crate::CellDefect::StyleIndexOutOfRange);
    let report = sheet.tolerance_report().expect("sheet report");
    assert_eq!(report.total(), 1);
    assert_eq!(report.defects()[0].defect, defect);
    assert_eq!(report.defects()[0].sheet, None);
    assert_eq!(report.defects()[0].location, location);
    let aggregate = workbook.tolerance_report();
    assert_eq!(aggregate.total(), 1);
    assert_eq!(aggregate.defects()[0].sheet, Some(0));
    assert_eq!(aggregate.defects()[0].location, location);
}

#[test]
fn concurrent_snapshot_reads_need_no_public_locking() {
    let workbook = Package::new().expect("valid baseline");
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
    let workbook = Package::from_opc(package_with_cells()).expect("valid workbook");
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
    let workbook = Package::from_opc(package_with_cells()).expect("valid workbook");
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

    let workbook = Package::from_opc(package).expect("valid chart graph");
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
        let workbook = Package::open(path).expect("corpus workbook");
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

#[test]
fn validated_worksheet_store_adoption_requires_exact_bytes_and_lineage() {
    use super::model::ValidatedWorksheetStore;

    let source = Package::new().expect("source workbook");
    let uri = source.inner.sheets[0].part_uri.clone();

    let same_lineage =
        Workbook::from_package_with_styles(source.inner.package.clone(), Some(&source))
            .expect("same-lineage target");
    let equal_but_distinct = Arc::new(
        same_lineage
            .inner
            .package
            .get_part(&uri)
            .expect("worksheet part")
            .blob()
            .to_vec(),
    );
    let store = crate::raw::worksheet::parse(&equal_but_distinct, || source.inner.shared_strings())
        .expect("worksheet store");
    same_lineage
        .adopt_validated_worksheet_stores(
            &source,
            vec![ValidatedWorksheetStore {
                uri: uri.clone(),
                content: equal_but_distinct,
                store,
            }],
        )
        .expect("skip byte-mismatched cache");
    assert!(same_lineage.inner.sheets[0].cells.get().is_none());

    let distinct_lineage =
        Package::from_opc(source.inner.package.clone()).expect("distinct-lineage target");
    let exact_content = distinct_lineage
        .inner
        .package
        .get_part(&uri)
        .expect("worksheet part")
        .blob_arc();
    let store = crate::raw::worksheet::parse(&exact_content, || source.inner.shared_strings())
        .expect("worksheet store");
    distinct_lineage
        .adopt_validated_worksheet_stores(
            &source,
            vec![ValidatedWorksheetStore {
                uri,
                content: exact_content,
                store,
            }],
        )
        .expect("skip lineage-mismatched cache");
    assert!(distinct_lineage.inner.sheets[0].cells.get().is_none());
}

/// Audit finding 4.4 through the owned facade: a dangling shared-string
/// reference opens by default, reads empty, and is recorded; the conformant
/// control reports clean; malformed worksheet XML is still fatal.
#[test]
fn dangling_shared_string_reference_opens_with_a_recorded_repair() {
    let baseline = Package::new().expect("valid baseline");
    let sheet_uri = PackURI::new("/xl/worksheets/sheet1.xml").expect("sheet URI");
    let with_sheet = |xml: &[u8]| {
        let mut package = baseline.inner.package.clone();
        package
            .get_part_mut(&sheet_uri)
            .expect("sheet part")
            .set_blob(xml.to_vec());
        Package::from_opc(package).expect("lazy worksheet")
    };

    let dangling = with_sheet(
        br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="s"><v>3</v></c><c r="B1"><v>2</v></c></row></sheetData></worksheet>"#,
    );
    assert!(dangling.tolerance_report().is_clean(), "nothing parsed yet");
    let sheet = dangling.sheet(0usize).expect("lookup").expect("sheet");
    assert!(matches!(
        sheet.cell("A1").expect("repaired cell"),
        crate::cell::View::Stored(Cell::Value(Value::Text(text))) if text.as_str().is_empty()
    ));
    let report = dangling.tolerance_report();
    assert_eq!(report.total(), 1);
    assert_eq!(
        report.defects()[0].defect,
        crate::Defect::Cell(crate::CellDefect::MissingSharedStringTable)
    );
    assert_eq!(report.defects()[0].sheet, Some(0));
    assert_eq!(
        report.defects()[0].location,
        crate::ToleranceLocation::Cell(litchi_sheet::Cell::from_a1("A1").expect("A1"))
    );

    let control = with_sheet(
        br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="B1"><v>2</v></c></row></sheetData></worksheet>"#,
    );
    let sheet = control.sheet(0usize).expect("lookup").expect("sheet");
    sheet.cell("B1").expect("conformant cell");
    assert!(sheet.tolerance_report().expect("report").is_clean());
    assert!(control.tolerance_report().is_clean());

    let malformed = with_sheet(
        br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="s"><v>3</v></c></row></worksheet>"#,
    );
    let sheet = malformed.sheet(0usize).expect("lookup").expect("sheet");
    assert!(sheet.cell("A1").is_err());
    assert!(sheet.tolerance_report().is_err());
}
