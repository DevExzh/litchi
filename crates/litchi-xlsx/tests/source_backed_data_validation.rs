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
    Budget, CancellationSource, ExecutionContext, ExecutionError, ExecutionLimits, Limits, ReadAt,
    Resource, SourceVersion,
};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{
    BlobPart, OpcError, OpcPackage, PackURI, PackageWriter, ReadLimits, SourceCacheLimits,
    TargetMode,
};
use litchi_xlsx::data_validation::{
    Collection, Formula, ListSource, Source, SourceBackedEditor, SourceEdit, Sqref, Validation,
    ValidationOperator, ValidationType, replace_data_validation_collections,
};
use litchi_xlsx::{Error, Package};

const SML: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const MAIN: &str = "/xl/workbook.xml";
const SHEET: &str = "/xl/worksheets/sheet1.xml";
const SECOND: &str = "/xl/worksheets/sheet2.xml";
const UNUSED: &str = "/xl/media/unused.bin";

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
        Ok(SourceVersion::new(
            101,
            self.revision.load(Ordering::SeqCst),
        ))
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

fn fixture(sheet_xml: String, signed: bool) -> Vec<u8> {
    let workbook = format!(
        r#"<workbook xmlns="{SML}" xmlns:r="{REL}"><sheets><sheet name="Sheet1" sheetId="1" r:id="rIdSheet"/><sheet name="Unused" sheetId="2" r:id="rIdUnused"/></sheets><calcPr calcId="7"/></workbook>"#
    );
    let mut package = OpcPackage::new();
    for (name, content_type, bytes) in [
        (MAIN, ct::SML_SHEET_MAIN, workbook.into_bytes()),
        (SHEET, ct::SML_WORKSHEET, sheet_xml.into_bytes()),
        (
            SECOND,
            ct::SML_WORKSHEET,
            format!(
                r#"<worksheet xmlns="{SML}"><sheetData/><!--{}--></worksheet>"#,
                "x".repeat(64 * 1024)
            )
            .into_bytes(),
        ),
        (
            UNUSED,
            "application/octet-stream",
            (0..64 * 1024).map(|value| (value % 251) as u8).collect(),
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
    workbook_part
        .rels_mut()
        .try_add_relationship(
            rt::WORKSHEET.to_string(),
            "worksheets/sheet1.xml".to_string(),
            "rIdSheet".to_string(),
            TargetMode::Internal,
        )
        .unwrap();
    workbook_part
        .rels_mut()
        .try_add_relationship(
            rt::WORKSHEET.to_string(),
            "worksheets/sheet2.xml".to_string(),
            "rIdUnused".to_string(),
            TargetMode::Internal,
        )
        .unwrap();
    package
        .get_part_mut(&PackURI::new(SHEET).unwrap())
        .unwrap()
        .rels_mut()
        .try_add_relationship(
            rt::DRAWING.to_string(),
            "../media/unused.bin".to_string(),
            "rIdDrawing".to_string(),
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

fn ordinary_fixture(signed: bool) -> Vec<u8> {
    fixture(
        format!(r#"<worksheet xmlns="{SML}"><sheetData/></worksheet>"#),
        signed,
    )
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
        "xlsx-data-validation-managed-test",
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

fn target_collections() -> Vec<Collection> {
    let mut rule = Validation::new(
        Source::Core,
        ValidationType::Whole,
        Sqref::parse("A1:A8").unwrap(),
    );
    rule.set_operator(ValidationOperator::Between);
    rule.set_allow_blank(true);
    rule.set_formula1(Some(ListSource::Formula(Formula::new("1").unwrap())))
        .unwrap();
    rule.set_formula2(Some(Formula::new("10").unwrap()))
        .unwrap();
    vec![Collection::new(Source::Core, vec![rule]).unwrap()]
}

fn changed_commit(editor: &SourceBackedEditor) -> litchi_xlsx::data_validation::Commit {
    let mut edit = editor.edit("Sheet1").unwrap();
    assert!(edit.set_collections(target_collections()).unwrap());
    let commit = edit.commit().unwrap();
    assert!(commit.changed());
    assert_eq!(commit.diagnostics().touched_worksheets(), 1);
    commit
}

fn relationship_signatures(relationships: &litchi_opc::Relationships) -> Vec<String> {
    let mut values = relationships
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
    values.sort_unstable();
    values
}

#[test]
fn changed_edit_reopens_inverts_and_preserves_unselected_parts() {
    let source_bytes = ordinary_fixture(false);
    let editor =
        SourceBackedEditor::from_read_at(Arc::new(VersionedSource::new(source_bytes.clone())))
            .unwrap();
    let commit = changed_commit(&editor);
    assert_eq!(editor.cache_diagnostics().successful_loads, 2);

    let mut replay = OpcPackage::from_bytes(&source_bytes).unwrap();
    commit.patch().apply(&mut replay).unwrap();
    assert_eq!(
        litchi_xlsx::data_validation::Snapshot::load(&replay, "Sheet1")
            .unwrap()
            .collections(),
        target_collections()
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
    let editor =
        SourceBackedEditor::from_read_at(Arc::new(VersionedSource::new(source_bytes.clone())))
            .unwrap();
    let commit = changed_commit(&editor);
    editor
        .publish_commit_to_stream(&mut output, &commit)
        .unwrap();
    let source = OpcPackage::from_bytes(&source_bytes).unwrap();
    let output_package = OpcPackage::from_bytes(&output).unwrap();
    for part in source.iter_parts() {
        let rewritten = output_package.get_part(part.partname()).unwrap();
        assert_eq!(rewritten.content_type(), part.content_type());
        assert_eq!(
            relationship_signatures(rewritten.rels()),
            relationship_signatures(part.rels())
        );
        if part.partname() != &PackURI::new(SHEET).unwrap() {
            assert_eq!(rewritten.blob(), part.blob());
        }
    }
    let package = Package::from_opc(output_package).unwrap();
    assert_eq!(
        package
            .workbook()
            .unwrap()
            .sheet("Sheet1")
            .unwrap()
            .unwrap()
            .data_validations()
            .unwrap(),
        target_collections()
    );
}

#[test]
fn managed_snapshot_changed_publication_and_budget_release_are_exact() {
    let source_bytes = ordinary_fixture(false);
    let exact = part_len(&source_bytes, MAIN) + part_len(&source_bytes, SHEET);
    let (budget, _cancellation_source, context) = managed_context(exact);
    let editor =
        SourceBackedEditor::from_read_at_with_limits_and_cache_limits_and_execution_context(
            Arc::new(VersionedSource::new(source_bytes.clone())),
            ReadLimits::default(),
            SourceCacheLimits::new(usize::try_from(exact).unwrap(), 8).unwrap(),
            context,
        )
        .unwrap();
    let snapshot = editor.snapshot("Sheet1").unwrap();
    assert!(snapshot.collections().is_empty());
    assert_eq!(budget.used(Resource::Memory), exact);
    drop(snapshot);

    let commit = changed_commit(&editor);
    let mut replay = OpcPackage::from_bytes(&source_bytes).unwrap();
    commit.patch().apply(&mut replay).unwrap();
    assert!(matches!(
        commit.patch().inverse().apply(&mut replay),
        Err(Error::Package(OpcError::ManagedPartDataArcEscape))
    ));
    let mut output = Vec::new();
    editor
        .publish_commit_to_stream(&mut output, &commit)
        .unwrap();
    let published = OpcPackage::from_bytes(&output).unwrap();
    assert_eq!(
        litchi_xlsx::data_validation::Snapshot::load(&published, "Sheet1")
            .unwrap()
            .collections(),
        target_collections()
    );
    assert_eq!(
        published
            .get_part(&PackURI::new(UNUSED).unwrap())
            .unwrap()
            .blob(),
        OpcPackage::from_bytes(&source_bytes)
            .unwrap()
            .get_part(&PackURI::new(UNUSED).unwrap())
            .unwrap()
            .blob()
    );
    drop(commit);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn managed_one_under_budget_refuses_before_selected_payload_retention() {
    let source_bytes = ordinary_fixture(false);
    let exact = part_len(&source_bytes, MAIN) + part_len(&source_bytes, SHEET);
    let (budget, _cancellation_source, context) = managed_context(exact - 1);
    let editor = SourceBackedEditor::from_read_at_with_execution_context(
        Arc::new(VersionedSource::new(source_bytes)),
        ReadLimits::default(),
        context,
    )
    .unwrap();
    assert!(matches!(
        editor.snapshot("Sheet1"),
        Err(Error::Package(OpcError::Execution(
            ExecutionError::ResourceLimit(_)
        )))
    ));
    assert_eq!(editor.cache_diagnostics().successful_loads, 1);
    drop(editor);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn no_op_signed_source_is_exact_but_changed_signed_source_is_refused() {
    let source_bytes = ordinary_fixture(true);
    let editor =
        SourceBackedEditor::from_read_at(Arc::new(VersionedSource::new(source_bytes.clone())))
            .unwrap();
    let commit = editor.edit("Sheet1").unwrap().commit().unwrap();
    assert!(!commit.changed());
    let mut output = Vec::new();
    editor
        .publish_commit_to_stream(&mut output, &commit)
        .unwrap();
    assert_eq!(output, source_bytes);

    let editor =
        SourceBackedEditor::from_read_at(Arc::new(VersionedSource::new(source_bytes))).unwrap();
    let commit = changed_commit(&editor);
    assert!(
        editor
            .publish_commit_to_stream(Vec::new(), &commit)
            .is_err()
    );
}

#[test]
fn add_replace_clear_noop_and_atomic_update_are_exact() {
    let source_bytes = ordinary_fixture(false);
    let mut seeded = OpcPackage::from_bytes(&source_bytes).unwrap();
    let sheet = seeded.get_part_mut(&PackURI::new(SHEET).unwrap()).unwrap();
    sheet.set_blob(
        replace_data_validation_collections(sheet.blob(), &target_collections()).unwrap(),
    );
    let seeded = PackageWriter::to_bytes(&seeded).unwrap();
    let editor = SourceBackedEditor::from_read_at(Arc::new(VersionedSource::new(seeded))).unwrap();
    let mut edit = editor.edit("Sheet1").unwrap();
    assert!(!edit.set_collections(target_collections()).unwrap());
    assert!(edit.clear());
    assert!(!edit.clear());
    let commit = edit.commit().unwrap();
    assert!(commit.changed());
    let mut output = Vec::new();
    let published = editor
        .publish_commit_to_stream(&mut output, &commit)
        .unwrap();
    assert!(published.collections().is_empty());

    let editor =
        SourceBackedEditor::from_read_at(Arc::new(VersionedSource::new(source_bytes))).unwrap();
    let mut edit = editor.edit("Sheet1").unwrap();
    assert!(matches!(
        edit.update(|collections| {
            *collections = target_collections();
            Err(Error::UnsupportedSelector)
        }),
        Err(Error::UnsupportedSelector)
    ));
    assert!(edit.collections().is_empty());
    assert!(
        edit.update(|collections| {
            *collections = target_collections();
            Ok(())
        })
        .unwrap()
    );
    assert!(!edit.set_collections(target_collections()).unwrap());
    assert!(edit.clear());
    assert!(!edit.commit().unwrap().changed());
}

#[test]
fn patch_conflicts_on_owner_or_outbound_relationship_change_and_source_version_change() {
    let source_bytes = ordinary_fixture(false);
    let editor =
        SourceBackedEditor::from_read_at(Arc::new(VersionedSource::new(source_bytes.clone())))
            .unwrap();
    let commit = changed_commit(&editor);
    let mut package = OpcPackage::from_bytes(&source_bytes).unwrap();
    package
        .get_part_mut(&PackURI::new(SHEET).unwrap())
        .unwrap()
        .rels_mut()
        .try_add_relationship(
            rt::HYPERLINK.to_string(),
            "https://example.invalid".to_string(),
            "rIdChanged".to_string(),
            TargetMode::External,
        )
        .unwrap();
    assert!(matches!(
        commit.patch().apply(&mut package),
        Err(Error::PatchConflict { .. })
    ));

    let source = Arc::new(VersionedSource::new(source_bytes));
    let editor = SourceBackedEditor::from_read_at(source.clone()).unwrap();
    let commit = changed_commit(&editor);
    source.changed();
    assert!(
        editor
            .publish_commit_to_stream(Vec::new(), &commit)
            .is_err()
    );

    let foreign = ordinary_fixture(false);
    let mut foreign_package = OpcPackage::from_bytes(&foreign).unwrap();
    foreign_package
        .get_part_mut(&PackURI::new(SHEET).unwrap())
        .unwrap()
        .set_blob(
            format!(r#"<worksheet xmlns="{SML}"><sheetData/><extLst/></worksheet>"#).into_bytes(),
        );
    let foreign = PackageWriter::to_bytes(&foreign_package).unwrap();
    let foreign_editor =
        SourceBackedEditor::from_read_at(Arc::new(VersionedSource::new(foreign))).unwrap();
    assert!(matches!(
        foreign_editor.publish_commit_to_stream(Vec::new(), &commit),
        Err(Error::PatchConflict { .. })
    ));

    let mut retargeted = OpcPackage::from_bytes(&ordinary_fixture(false)).unwrap();
    let workbook = retargeted
        .get_part_mut(&PackURI::new(MAIN).unwrap())
        .unwrap();
    workbook.rels_mut().remove("rIdSheet").unwrap();
    workbook
        .rels_mut()
        .try_add_relationship(
            rt::WORKSHEET.to_string(),
            "worksheets/sheet2.xml".to_string(),
            "rIdSheet".to_string(),
            TargetMode::Internal,
        )
        .unwrap();
    assert!(matches!(
        commit.patch().apply(&mut retargeted),
        Err(Error::PatchConflict { .. })
    ));
}

#[test]
fn mce_selected_collections_and_partial_sinks_are_refused() {
    let mce = format!(
        r#"<worksheet xmlns="{SML}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" mc:Ignorable="x14"><mc:AlternateContent><mc:Choice Requires="x14"><dataValidations count="1"><dataValidation type="whole" sqref="A1"><formula1>1</formula1><formula2>2</formula2></dataValidation></dataValidations></mc:Choice><mc:Fallback/></mc:AlternateContent><sheetData/></worksheet>"#
    );
    let editor =
        SourceBackedEditor::from_read_at(Arc::new(VersionedSource::new(fixture(mce, false))))
            .unwrap();
    let mut edit = editor.edit("Sheet1").unwrap();
    edit.set_collections(target_collections()).unwrap();
    let error = edit.commit().unwrap_err();
    assert!(error.to_string().contains("selected through MCE"));

    let limits = ReadLimits::builder()
        .max_part_bytes(1)
        .unwrap()
        .build()
        .unwrap();
    assert!(matches!(
        SourceBackedEditor::from_read_at_with_limits(
            Arc::new(VersionedSource::new(ordinary_fixture(false))),
            limits,
        ),
        Err(Error::Package(OpcError::ReadLimit { .. }))
    ));

    let source_bytes = ordinary_fixture(false);
    let editor =
        SourceBackedEditor::from_read_at(Arc::new(VersionedSource::new(source_bytes))).unwrap();
    let commit = changed_commit(&editor);
    let error = editor
        .publish_commit_to_stream(
            FailingSink {
                accepted: 0,
                limit: 32,
            },
            &commit,
        )
        .unwrap_err();
    assert!(error.to_string().contains("sink") || error.to_string().contains("injected"));
}

#[test]
fn chartsheet_selection_is_refused_and_editor_is_send() {
    fn assert_send<T: Send>() {}
    assert_send::<SourceBackedEditor>();
    assert_send::<SourceEdit>();

    const CHARTSHEET_REL: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet";
    const CHARTSHEET_CT: &str =
        "application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml";

    let mut package = OpcPackage::from_bytes(&ordinary_fixture(false)).unwrap();
    let sheet_uri = PackURI::new(SHEET).unwrap();
    let sheet = package.get_part_mut(&sheet_uri).unwrap();
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
