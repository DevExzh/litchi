#![allow(
    clippy::unwrap_used,
    reason = "focused integration tests use panic-on-failure assertions"
)]

use std::io;
use std::io::Write;
use std::num::{NonZeroU64, NonZeroUsize};
use std::sync::Arc;
use std::sync::atomic::{AtomicU64, Ordering};

use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits, ReadAt, Resource,
    SourceVersion,
};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{
    BlobPart, OpcError, OpcPackage, PackURI, PackageWriter, SourceCacheLimits, TargetMode,
};
use litchi_xlsx::page_setup::{SourceBackedEditor, SourceEdit};
use litchi_xlsx::{
    Error, Order, Orientation, Package, Paper, ReadLimits, Scale, Setup,
    parse_worksheet_page_setup_relationship_id,
};

const SML: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const MAIN: &str = "/xl/workbook.xml";
const SHEET: &str = "/xl/worksheets/sheet1.xml";
const SECOND: &str = "/xl/worksheets/sheet2.xml";
const DRAWING: &str = "/xl/drawings/drawing1.xml";
const MEDIA: &str = "/xl/media/image1.png";
const DRAWING_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing";
const PRINTER_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings";
const PRINTER_CT: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings";

struct VersionedSource {
    bytes: Vec<u8>,
    revision: AtomicU64,
}

impl VersionedSource {
    fn new(bytes: Vec<u8>) -> Self {
        Self {
            bytes,
            revision: AtomicU64::new(0),
        }
    }

    fn changed(&self) {
        self.revision.fetch_add(1, Ordering::SeqCst);
    }
}

impl ReadAt for VersionedSource {
    fn len(&self) -> io::Result<u64> {
        Ok(self.bytes.len() as u64)
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        let offset = usize::try_from(offset)
            .map_err(|_| io::Error::new(io::ErrorKind::InvalidInput, "offset too large"))?;
        if offset >= self.bytes.len() {
            return Ok(0);
        }
        let end = offset.saturating_add(output.len()).min(self.bytes.len());
        output[..end - offset].copy_from_slice(&self.bytes[offset..end]);
        Ok(end - offset)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        Ok(SourceVersion::new(83, self.revision.load(Ordering::SeqCst)))
    }
}

struct FailingSink {
    accepted: usize,
    limit: usize,
}

impl Write for FailingSink {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        if self.accepted >= self.limit {
            return Err(io::Error::other("injected sink failure"));
        }
        let written = bytes.len().min(self.limit - self.accepted);
        self.accepted += written;
        Ok(written)
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

fn fixture(sheet_xml: String, workbook_suffix: &str, signed: bool) -> Vec<u8> {
    let workbook = format!(
        r#"<workbook xmlns="{SML}" xmlns:r="{REL}"><sheets><sheet name="Sheet1" sheetId="1" r:id="rIdSheet"/><sheet name="Unused" sheetId="2" r:id="rIdUnused"/></sheets><calcPr calcId="7"/>{workbook_suffix}</workbook>"#
    );
    let mut package = OpcPackage::new();
    for (name, content_type, bytes) in [
        (MAIN, ct::SML_SHEET_MAIN, workbook.into_bytes()),
        (SHEET, ct::SML_WORKSHEET, sheet_xml.into_bytes()),
        (
            SECOND,
            ct::SML_WORKSHEET,
            format!(
                r#"<worksheet xmlns="{SML}"><sheetData/><extLst><ext uri="urn:unselected">{}</ext></extLst></worksheet>"#,
                "x".repeat(128 * 1024)
            )
            .into_bytes(),
        ),
        (
            DRAWING,
            "application/vnd.openxmlformats-officedocument.drawing+xml",
            b"<drawing xmlns=\"urn:drawing\"/>".to_vec(),
        ),
        (
            MEDIA,
            "image/png",
            (0..128 * 1024).map(|value| (value % 251) as u8).collect(),
        ),
    ] {
        package
            .try_add_part(Box::new(BlobPart::new(
                PackURI::new(name).unwrap(),
                content_type.to_string(),
                bytes,
            )))
            .unwrap();
    }
    let workbook_part = package.get_part_mut(&PackURI::new(MAIN).unwrap()).unwrap();
    for (relationship_type, target, id) in [
        (rt::WORKSHEET, "worksheets/sheet1.xml", "rIdSheet"),
        (rt::WORKSHEET, "worksheets/sheet2.xml", "rIdUnused"),
    ] {
        workbook_part
            .rels_mut()
            .try_add_relationship(
                relationship_type.to_string(),
                target.to_string(),
                id.to_string(),
                TargetMode::Internal,
            )
            .unwrap();
    }
    package
        .get_part_mut(&PackURI::new(SHEET).unwrap())
        .unwrap()
        .rels_mut()
        .try_add_relationship(
            DRAWING_REL.to_string(),
            "../drawings/drawing1.xml".to_string(),
            "rIdDrawing".to_string(),
            TargetMode::Internal,
        )
        .unwrap();
    package
        .get_part_mut(&PackURI::new(DRAWING).unwrap())
        .unwrap()
        .rels_mut()
        .try_add_relationship(
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image".to_string(),
            "../media/image1.png".to_string(),
            "rIdImage".to_string(),
            TargetMode::Internal,
        )
        .unwrap();
    package.relate_to("xl/workbook.xml", rt::OFFICE_DOCUMENT);
    if signed {
        package
            .try_add_part(Box::new(BlobPart::new(
                PackURI::new("/_xmlsignatures/origin.sigs").unwrap(),
                ct::OPC_DIGITAL_SIGNATURE_ORIGIN.to_string(),
                b"<origin/>".to_vec(),
            )))
            .unwrap();
        package.relate_to("_xmlsignatures/origin.sigs", rt::DIGITAL_SIGNATURE_ORIGIN);
    }
    PackageWriter::to_bytes(&package).unwrap()
}

fn ordinary_fixture(workbook_suffix: &str, signed: bool) -> Vec<u8> {
    fixture(
        format!(r#"<worksheet xmlns="{SML}"><sheetData/><drawing r:id="rIdDrawing"/></worksheet>"#),
        workbook_suffix,
        signed,
    )
}

fn target_setup() -> Setup {
    let mut setup = Setup::new(Paper::A4);
    setup.scale = Some(Scale::new(85).unwrap());
    setup.order = Some(Order::OverThenDown);
    setup.orientation = Some(Orientation::Landscape);
    setup.use_printer_defaults = Some(false);
    setup
}

fn changed_commit(editor: &SourceBackedEditor) -> litchi_xlsx::page_setup::Commit {
    let mut edit = editor.edit("Sheet1").unwrap();
    assert!(edit.set(target_setup()));
    let commit = edit.commit().unwrap();
    assert!(commit.changed());
    assert_eq!(commit.diagnostics().touched_worksheets(), 1);
    commit
}

fn part_len(bytes: &[u8], member: &str) -> u64 {
    OpcPackage::from_bytes(bytes)
        .unwrap()
        .get_part(&PackURI::new(member).unwrap())
        .unwrap()
        .blob()
        .len() as u64
}

fn managed_context(memory: u64) -> (Budget, CancellationSource, ExecutionContext) {
    let budget = Budget::root(
        "xlsx-page-setup-managed-test",
        Limits::new(memory, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX),
    );
    let (cancellation_source, cancellation) = CancellationSource::pair();
    let execution_limits = ExecutionLimits::new(
        NonZeroUsize::new(1).unwrap(),
        NonZeroUsize::new(1).unwrap(),
        NonZeroU64::new(memory.max(1)).unwrap(),
        0,
    )
    .unwrap();
    let context = ExecutionContext::new(budget.clone(), cancellation, execution_limits);
    (budget, cancellation_source, context)
}

#[test]
fn managed_payload_retention_publication_and_owning_escape_are_explicit() {
    let bytes = ordinary_fixture("", false);
    let exact = part_len(&bytes, MAIN) + part_len(&bytes, SHEET);
    let (budget, _cancellation_source, context) = managed_context(exact);
    let editor =
        SourceBackedEditor::from_read_at_with_limits_and_cache_limits_and_execution_context(
            Arc::new(VersionedSource::new(bytes.clone())),
            ReadLimits::default(),
            SourceCacheLimits::new(usize::try_from(exact).unwrap(), 8).unwrap(),
            context,
        )
        .unwrap();
    assert!(editor.cache_diagnostics().budget_managed);
    assert_eq!(budget.used(Resource::Memory), 0);
    let commit = changed_commit(&editor);
    assert_eq!(budget.used(Resource::Memory), exact);
    let mut replay = OpcPackage::from_bytes(&bytes).unwrap();
    commit.patch().apply(&mut replay).unwrap();
    assert!(matches!(
        commit.patch().inverse().apply(&mut replay),
        Err(Error::Package(OpcError::ManagedPartDataArcEscape))
    ));
    let mut output = Vec::new();
    editor
        .publish_commit_to_stream(&mut output, &commit)
        .unwrap();
    assert_eq!(
        litchi_xlsx::page_setup::Snapshot::load(
            &OpcPackage::from_bytes(&output).unwrap(),
            "Sheet1",
        )
        .unwrap()
        .page_setup(),
        Some(&target_setup())
    );
    drop(commit);
    assert_eq!(budget.used(Resource::Memory), 0);
}

fn relationship_signatures(relationships: &litchi_opc::Relationships) -> Vec<String> {
    let mut signatures = relationships
        .iter()
        .map(|relationship| {
            format!(
                "{}\u{1f}{}\u{1f}{}\u{1f}{:?}",
                relationship.r_id(),
                relationship.reltype(),
                relationship.target_ref(),
                relationship.target_mode()
            )
        })
        .collect::<Vec<_>>();
    signatures.sort_unstable();
    signatures
}

fn eager_output(source: &[u8]) -> Vec<u8> {
    let package = Package::from_bytes(source.to_vec()).unwrap();
    let workbook = package.into_workbook().unwrap();
    let mut edit = workbook.edit().unwrap();
    assert!(
        edit.put_page_setup("Sheet1", target_setup())
            .unwrap()
            .is_some()
    );
    let commit = edit.commit().unwrap();
    let mut output = Vec::new();
    commit.workbook().write_to(&mut output).unwrap();
    output
}

#[test]
fn changed_edit_matches_eager_reopens_inverts_and_preserves_every_relationship() {
    let source_bytes = ordinary_fixture(r#"<extLst><ext uri="urn:keep"/></extLst>"#, false);
    let editor =
        SourceBackedEditor::from_read_at(Arc::new(VersionedSource::new(source_bytes.clone())))
            .unwrap();
    assert_eq!(editor.cache_diagnostics().successful_loads, 0);
    let commit = changed_commit(&editor);
    assert_eq!(editor.cache_diagnostics().successful_loads, 2);

    let mut replay = OpcPackage::from_bytes(&source_bytes).unwrap();
    commit.patch().apply(&mut replay).unwrap();
    assert_eq!(
        litchi_xlsx::page_setup::Snapshot::load(&replay, "Sheet1")
            .unwrap()
            .page_setup(),
        Some(&target_setup())
    );
    commit.patch().inverse().apply(&mut replay).unwrap();
    assert_eq!(
        replay
            .get_part(&PackURI::new(SHEET).unwrap())
            .unwrap()
            .blob(),
        OpcPackage::from_bytes(&source_bytes)
            .unwrap()
            .get_part(&PackURI::new(SHEET).unwrap())
            .unwrap()
            .blob()
    );

    let mut output = Vec::new();
    let published = editor
        .publish_commit_to_stream(&mut output, &commit)
        .unwrap();
    assert_eq!(published.page_setup(), Some(&target_setup()));
    assert_eq!(output, eager_output(&source_bytes));

    let source = OpcPackage::from_bytes(&source_bytes).unwrap();
    let candidate = OpcPackage::from_bytes(&output).unwrap();
    assert_eq!(source.part_count(), candidate.part_count());
    assert_eq!(
        relationship_signatures(source.rels()),
        relationship_signatures(candidate.rels())
    );
    for part in source.iter_parts() {
        let output_part = candidate.get_part(part.partname()).unwrap();
        assert_eq!(part.content_type(), output_part.content_type());
        assert_eq!(
            relationship_signatures(part.rels()),
            relationship_signatures(output_part.rels())
        );
        if part.partname().as_str() == SHEET {
            assert_ne!(part.blob(), output_part.blob());
        } else {
            assert_eq!(part.blob(), output_part.blob());
        }
    }
    let reopened = Package::from_bytes(output).unwrap();
    let sheet = reopened
        .workbook()
        .unwrap()
        .sheet("Sheet1")
        .unwrap()
        .unwrap();
    assert_eq!(sheet.page_setup().unwrap(), Some(target_setup()));
    assert!(
        parse_worksheet_page_setup_relationship_id(
            candidate
                .get_part(&PackURI::new(SHEET).unwrap())
                .unwrap()
                .blob()
        )
        .unwrap()
        .is_none()
    );
}

#[test]
fn add_replace_remove_noop_and_signed_refusal_are_exact() {
    let bytes = ordinary_fixture("", false);
    let editor =
        SourceBackedEditor::from_read_at(Arc::new(VersionedSource::new(bytes.clone()))).unwrap();
    let mut edit = editor.edit("Sheet1").unwrap();
    assert!(edit.set(target_setup()));
    assert!(!edit.set(target_setup()));
    assert!(edit.remove());
    assert!(!edit.remove());
    let noop = edit.commit().unwrap();
    assert!(!noop.changed());
    let mut output = Vec::new();
    editor.publish_commit_to_stream(&mut output, &noop).unwrap();
    assert_eq!(output, bytes);

    let existing = fixture(
        format!(
            r#"<worksheet xmlns="{SML}" xmlns:r="{REL}"><sheetData/><pageSetup paperSize="1"/><drawing r:id="rIdDrawing"/></worksheet>"#
        ),
        "",
        false,
    );
    let editor =
        SourceBackedEditor::from_read_at(Arc::new(VersionedSource::new(existing))).unwrap();
    let mut edit = editor.edit("Sheet1").unwrap();
    assert!(edit.set(target_setup()));
    assert!(edit.commit().unwrap().changed());

    let signed = ordinary_fixture("", true);
    let editor =
        SourceBackedEditor::from_read_at(Arc::new(VersionedSource::new(signed.clone()))).unwrap();
    let noop = editor.edit("Sheet1").unwrap().commit().unwrap();
    output.clear();
    editor.publish_commit_to_stream(&mut output, &noop).unwrap();
    assert_eq!(output, signed);

    let editor = SourceBackedEditor::from_read_at(Arc::new(VersionedSource::new(signed))).unwrap();
    let commit = changed_commit(&editor);
    output.clear();
    assert!(matches!(
        editor.publish_commit_to_stream(&mut output, &commit),
        Err(Error::Package(OpcError::SignedSourceRequiresExplicitPolicy))
    ));
    assert!(output.is_empty());
}

#[test]
fn printer_references_and_relationships_are_refused_without_materializing_resources() {
    let dangling = fixture(
        format!(
            r#"<worksheet xmlns="{SML}" xmlns:r="{REL}"><sheetData/><pageSetup r:id="rIdPrinter"/><drawing r:id="rIdDrawing"/></worksheet>"#
        ),
        "",
        false,
    );
    let editor =
        SourceBackedEditor::from_read_at(Arc::new(VersionedSource::new(dangling))).unwrap();
    assert!(matches!(
        editor.edit("Sheet1"),
        Err(Error::Unsupported { .. })
    ));

    let mut package = OpcPackage::from_bytes(&ordinary_fixture("", false)).unwrap();
    package
        .try_add_part(Box::new(BlobPart::new(
            PackURI::new("/xl/printerSettings/printerSettings1.bin").unwrap(),
            PRINTER_CT.to_string(),
            vec![0; 256],
        )))
        .unwrap();
    package
        .get_part_mut(&PackURI::new(SHEET).unwrap())
        .unwrap()
        .rels_mut()
        .try_add_relationship(
            PRINTER_REL.to_string(),
            "../printerSettings/printerSettings1.bin".to_string(),
            "rIdPrinter".to_string(),
            TargetMode::Internal,
        )
        .unwrap();
    let bytes = PackageWriter::to_bytes(&package).unwrap();
    let editor = SourceBackedEditor::from_read_at(Arc::new(VersionedSource::new(bytes))).unwrap();
    assert!(matches!(
        editor.edit("Sheet1"),
        Err(Error::Unsupported { .. })
    ));
    assert_eq!(editor.cache_diagnostics().successful_loads, 2);
}

#[test]
fn foreign_version_retarget_and_outbound_relationship_changes_are_rejected() {
    let bytes = ordinary_fixture("", false);
    let editor =
        SourceBackedEditor::from_read_at(Arc::new(VersionedSource::new(bytes.clone()))).unwrap();
    let commit = changed_commit(&editor);
    let foreign = ordinary_fixture(r#"<extLst><ext uri="urn:foreign"/></extLst>"#, false);
    let foreign_editor =
        SourceBackedEditor::from_read_at(Arc::new(VersionedSource::new(foreign))).unwrap();
    let mut output = Vec::new();
    assert!(matches!(
        foreign_editor.publish_commit_to_stream(&mut output, &commit),
        Err(Error::PatchConflict { .. })
    ));
    assert!(output.is_empty());

    let source = Arc::new(VersionedSource::new(bytes.clone()));
    let editor = SourceBackedEditor::from_read_at(source.clone()).unwrap();
    let commit = changed_commit(&editor);
    source.changed();
    assert!(matches!(
        editor.publish_commit_to_stream(&mut output, &commit),
        Err(Error::Package(OpcError::SourceChanged { .. }))
    ));
    assert!(output.is_empty());

    let mut retargeted = OpcPackage::from_bytes(&bytes).unwrap();
    retargeted
        .get_part_mut(&PackURI::new(MAIN).unwrap())
        .unwrap()
        .rels_mut()
        .retarget("rIdSheet", "worksheets/sheet2.xml".to_string())
        .unwrap();
    assert!(matches!(
        commit.patch().apply(&mut retargeted),
        Err(Error::PatchConflict { .. })
    ));

    let mut changed_relationship = OpcPackage::from_bytes(&bytes).unwrap();
    changed_relationship
        .get_part_mut(&PackURI::new(SHEET).unwrap())
        .unwrap()
        .rels_mut()
        .retarget("rIdDrawing", "../drawings/other.xml".to_string())
        .unwrap();
    assert!(matches!(
        commit.patch().apply(&mut changed_relationship),
        Err(Error::PatchConflict { .. })
    ));
}

#[test]
fn mce_limits_and_partial_sink_are_checked() {
    let mce_sheet = format!(
        concat!(
            r#"<worksheet xmlns="{SML}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:future">"#,
            r#"<mc:AlternateContent><mc:Choice Requires="x"><x:future/></mc:Choice><mc:Fallback>"#,
            r#"<pageSetup paperSize="1"/>"#,
            r#"</mc:Fallback></mc:AlternateContent></worksheet>"#,
        ),
        SML = SML,
    );
    let mce = fixture(mce_sheet, "", false);
    let editor = SourceBackedEditor::from_read_at(Arc::new(VersionedSource::new(mce))).unwrap();
    let mut edit = editor.edit("Sheet1").unwrap();
    assert!(edit.set(target_setup()));
    assert!(edit.commit().is_err());

    let limits = ReadLimits::builder()
        .max_part_bytes(1)
        .unwrap()
        .build()
        .unwrap();
    assert!(matches!(
        SourceBackedEditor::from_read_at_with_limits(
            Arc::new(VersionedSource::new(ordinary_fixture("", false))),
            limits,
        ),
        Err(Error::Package(OpcError::ReadLimit { .. }))
    ));

    let editor = SourceBackedEditor::from_read_at(Arc::new(VersionedSource::new(
        ordinary_fixture("", false),
    )))
    .unwrap();
    let commit = changed_commit(&editor);
    let mut sink = FailingSink {
        accepted: 0,
        limit: 128,
    };
    assert!(matches!(
        editor.publish_commit_to_stream(&mut sink, &commit),
        Err(Error::Package(OpcError::IncompleteOutput { .. }))
    ));
    assert_eq!(sink.accepted, 128);
}

#[test]
fn source_edit_is_send_and_chartsheets_are_refused() {
    fn assert_send<T: Send>() {}
    assert_send::<SourceEdit>();

    const CHARTSHEET_REL: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet";
    const CHARTSHEET_CT: &str =
        "application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml";
    let mut package = OpcPackage::from_bytes(&ordinary_fixture("", false)).unwrap();
    let sheet = package.get_part_mut(&PackURI::new(SHEET).unwrap()).unwrap();
    sheet.set_content_type(CHARTSHEET_CT.to_owned()).unwrap();
    sheet.set_blob(format!(r#"<chartsheet xmlns="{SML}"/>"#).into_bytes());
    let workbook = package.get_part_mut(&PackURI::new(MAIN).unwrap()).unwrap();
    workbook.rels_mut().remove("rIdSheet").unwrap();
    workbook
        .rels_mut()
        .try_add_relationship(
            CHARTSHEET_REL.to_owned(),
            "worksheets/sheet1.xml".to_owned(),
            "rIdSheet".to_owned(),
            TargetMode::Internal,
        )
        .unwrap();
    let bytes = PackageWriter::to_bytes(&package).unwrap();
    let editor = SourceBackedEditor::from_read_at(Arc::new(VersionedSource::new(bytes))).unwrap();
    assert!(matches!(
        editor.edit("Sheet1"),
        Err(Error::NotWorksheet { .. })
    ));
}
