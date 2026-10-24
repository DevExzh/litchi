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
use litchi_xlsx::SourceBackedWorkbook;
use litchi_xlsx::sheet_protection::{
    Metadata, ProtectedRange, ProtectedRangeCollection, ProtectedRangeSource, Protection,
    ProtectionPasswordVerifier, ProtectionRangeSqref,
};
use litchi_xlsx::{Error, Package, SourceLimits};

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
        Ok(SourceVersion::new(73, self.revision.load(Ordering::SeqCst)))
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
    package
        .try_add_part(Box::new(BlobPart::new(
            PackURI::new(MAIN).unwrap(),
            ct::SML_SHEET_MAIN.to_string(),
            workbook.into_bytes(),
        )))
        .unwrap();
    package
        .try_add_part(Box::new(BlobPart::new(
            PackURI::new(SHEET).unwrap(),
            ct::SML_WORKSHEET.to_string(),
            sheet_xml.into_bytes(),
        )))
        .unwrap();
    package
        .try_add_part(Box::new(BlobPart::new(
            PackURI::new(SECOND).unwrap(),
            ct::SML_WORKSHEET.to_string(),
            format!(
                r#"<worksheet xmlns="{SML}"><sheetData/><extLst><ext uri="urn:unselected">{}</ext></extLst></worksheet>"#,
                "x".repeat(128 * 1024)
            )
            .into_bytes(),
        )))
        .unwrap();
    package
        .try_add_part(Box::new(BlobPart::new(
            PackURI::new(UNUSED).unwrap(),
            "application/octet-stream".to_string(),
            (0..128 * 1024).map(|value| (value % 251) as u8).collect(),
        )))
        .unwrap();
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
        format!(r#"<worksheet xmlns="{SML}"><sheetData/></worksheet>"#),
        workbook_suffix,
        signed,
    )
}

fn part_len(bytes: &[u8], member: &str) -> u64 {
    OpcPackage::from_bytes(bytes, SourceLimits::default())
        .unwrap()
        .get_part(&PackURI::new(member).unwrap())
        .unwrap()
        .blob()
        .len() as u64
}

fn managed_context(memory: u64) -> (Budget, CancellationSource, ExecutionContext) {
    let budget = Budget::root(
        "xlsx-sheet-protection-managed-test",
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

fn target_metadata() -> Metadata {
    let mut protection = Protection::new();
    protection
        .set_verifier(Some(ProtectionPasswordVerifier::Legacy(0x83af)))
        .unwrap();
    protection.set_sheet_locked(true);
    protection.set_objects_locked(true);
    protection.set_auto_filter_locked(false);

    let mut range = ProtectedRange::new(
        ProtectedRangeSource::Core,
        "Input cells",
        ProtectionRangeSqref::parse("A1:B4 D6").unwrap(),
    )
    .unwrap();
    range
        .set_verifier(Some(ProtectionPasswordVerifier::Legacy(0x1a2b)))
        .unwrap();
    let collection =
        ProtectedRangeCollection::new(ProtectedRangeSource::Core, vec![range]).unwrap();
    let mut value = Metadata::new();
    value.set_sheet_protection(Some(protection)).unwrap();
    value
        .set_protected_range_collections(vec![collection])
        .unwrap();
    value
}

fn changed_edit(editor: &SourceBackedWorkbook) -> litchi_xlsx::sheet_protection::source::Edit {
    let mut edit = editor.edit_sheet_protection("Sheet1").unwrap();
    assert!(edit.set(target_metadata()).unwrap());
    edit
}

fn changed_commit(editor: &SourceBackedWorkbook) -> litchi_xlsx::sheet_protection::source::Commit {
    let commit = changed_edit(editor).commit().unwrap();
    assert!(commit.diagnostics().changed());
    commit
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

#[test]
fn changed_edit_reopens_inverts_and_changes_only_the_selected_worksheet() {
    let source_bytes = ordinary_fixture(r#"<extLst><ext uri="urn:keep"/></extLst>"#, false);
    let editor =
        SourceBackedWorkbook::from_read_at(Arc::new(VersionedSource::new(source_bytes.clone())))
            .unwrap();
    let opened = editor.cache_diagnostics().successful_loads;
    let commit = changed_commit(&editor);
    assert_eq!(editor.cache_diagnostics().successful_loads, opened + 1);

    let mut replay = OpcPackage::from_bytes(&source_bytes, SourceLimits::default()).unwrap();
    replay_patch(&mut replay, commit.patch()).unwrap();
    assert_eq!(
        litchi_xlsx::sheet_protection::Snapshot::load(&replay, "Sheet1")
            .unwrap()
            .metadata(),
        &target_metadata()
    );
    replay_patch(&mut replay, &commit.patch().inverse()).unwrap();
    assert_eq!(
        replay
            .get_part(&PackURI::new(SHEET).unwrap())
            .unwrap()
            .blob(),
        OpcPackage::from_bytes(&source_bytes, SourceLimits::default())
            .unwrap()
            .get_part(&PackURI::new(SHEET).unwrap())
            .unwrap()
            .blob()
    );

    let mut output = Vec::new();
    commit.package().write_to(&mut output).unwrap();
    let published = commit.patch().after().clone();
    assert_eq!(published.metadata(), &target_metadata());

    let source = OpcPackage::from_bytes(&source_bytes, SourceLimits::default()).unwrap();
    let candidate = OpcPackage::from_bytes(&output, SourceLimits::default()).unwrap();
    assert_eq!(source.part_count(), candidate.part_count());
    assert_eq!(
        relationship_signatures(source.rels()),
        relationship_signatures(candidate.rels())
    );
    for part in source
        .try_iter_parts()
        .map(|part| part.expect("part payload decodes"))
    {
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
    let reopened = Package::from_bytes(&output).unwrap();
    assert_eq!(
        reopened
            .workbook()
            .sheet("Sheet1")
            .unwrap()
            .unwrap()
            .protection()
            .unwrap(),
        target_metadata()
    );
}

#[test]
fn managed_snapshot_changed_publication_and_budget_release_are_exact() {
    let source_bytes = ordinary_fixture(r#"<extLst><ext uri="urn:keep"/></extLst>"#, false);
    let exact = part_len(&source_bytes, MAIN) + part_len(&source_bytes, SHEET);
    let (budget, _cancellation_source, context) = managed_context(exact);
    let editor = SourceBackedWorkbook::from_read_at_with_options(
        Arc::new(VersionedSource::new(source_bytes.clone())),
        litchi_xlsx::OpenOptions::new()
            .source_limits(SourceLimits::default())
            .cache_limits(SourceCacheLimits::new(usize::try_from(exact).unwrap(), 8).unwrap())
            .execution_context(context),
    )
    .unwrap();
    let snapshot = editor
        .edit_sheet_protection("Sheet1")
        .and_then(|edit| edit.before().cloned())
        .unwrap();
    assert_eq!(snapshot.metadata(), &Metadata::new());
    assert_eq!(budget.used(Resource::Memory), exact);
    drop(snapshot);

    let commit = changed_commit(&editor);
    let mut replay = OpcPackage::from_bytes(&source_bytes, SourceLimits::default()).unwrap();
    replay_patch(&mut replay, commit.patch()).unwrap();
    replay_patch(&mut replay, &commit.patch().inverse()).unwrap();
    let mut output = Vec::new();
    commit.package().write_to(&mut output).unwrap();
    let published = OpcPackage::from_bytes(&output, SourceLimits::default()).unwrap();
    assert_eq!(
        litchi_xlsx::sheet_protection::Snapshot::load(&published, "Sheet1")
            .unwrap()
            .metadata(),
        &target_metadata()
    );
    assert_eq!(
        published
            .get_part(&PackURI::new(UNUSED).unwrap())
            .unwrap()
            .blob(),
        OpcPackage::from_bytes(&source_bytes, SourceLimits::default())
            .unwrap()
            .get_part(&PackURI::new(UNUSED).unwrap())
            .unwrap()
            .blob()
    );
    drop(commit);
    assert!(budget.used(Resource::Memory) <= exact);
    drop(editor);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn add_replace_remove_noop_and_signed_refusal_are_exact() {
    let existing = target_metadata();
    let authored = litchi_xlsx::sheet_protection::write_core(
        &existing,
        litchi_xlsx::sheet_protection::Conformance::Transitional,
    )
    .unwrap();
    let existing_xml = format!(r#"<worksheet xmlns="{SML}"><sheetData/>{authored}</worksheet>"#);
    let bytes = fixture(existing_xml, "", false);
    let editor = SourceBackedWorkbook::from_read_at(Arc::new(VersionedSource::new(bytes))).unwrap();
    let mut edit = editor.edit_sheet_protection("Sheet1").unwrap();
    assert!(!edit.set(existing).unwrap());
    assert!(edit.clear().unwrap());
    let remove = edit.commit().unwrap();
    assert!(remove.diagnostics().changed());
    let mut output = Vec::new();
    remove.package().write_to(&mut output).unwrap();
    let published = remove.patch().after().clone();
    assert_eq!(published.metadata(), &Metadata::new());
    assert!(
        !OpcPackage::from_bytes(&output, SourceLimits::default())
            .unwrap()
            .get_part(&PackURI::new(SHEET).unwrap())
            .unwrap()
            .blob()
            .windows(b"sheetProtection".len())
            .any(|window| window == b"sheetProtection")
    );

    let bytes = ordinary_fixture("", false);
    let editor = SourceBackedWorkbook::from_read_at(Arc::new(VersionedSource::new(bytes))).unwrap();
    let mut edit = editor.edit_sheet_protection("Sheet1").unwrap();
    let unchanged = edit.metadata().unwrap().clone();
    assert!(matches!(
        edit.update(|metadata| {
            *metadata = target_metadata();
            Err(Error::UnsupportedSelector)
        }),
        Err(Error::UnsupportedSelector)
    ));
    assert_eq!(edit.metadata().unwrap(), &unchanged);
    assert!(
        edit.update(|metadata| {
            *metadata = target_metadata();
            Ok(())
        })
        .unwrap()
    );
    assert!(!edit.set(target_metadata()).unwrap());
    assert!(edit.clear().unwrap());
    assert!(!edit.clear().unwrap());
    assert!(!edit.commit().unwrap().diagnostics().changed());

    let signed = ordinary_fixture("", true);
    let editor =
        SourceBackedWorkbook::from_read_at(Arc::new(VersionedSource::new(signed.clone()))).unwrap();
    let noop = editor
        .edit_sheet_protection("Sheet1")
        .unwrap()
        .commit()
        .unwrap();
    output.clear();
    noop.package().write_to(&mut output).unwrap();
    assert_eq!(output, signed);

    let editor =
        SourceBackedWorkbook::from_read_at(Arc::new(VersionedSource::new(signed))).unwrap();
    assert!(matches!(
        changed_edit(&editor).commit(),
        Err(Error::Package(OpcError::SignedSourceRequiresExplicitPolicy))
    ));
}

#[test]
fn foreign_changed_and_retargeted_sources_are_rejected() {
    let bytes = ordinary_fixture("", false);
    let editor =
        SourceBackedWorkbook::from_read_at(Arc::new(VersionedSource::new(bytes.clone()))).unwrap();
    let commit = changed_commit(&editor);
    let foreign = ordinary_fixture(r#"<extLst><ext uri="urn:foreign"/></extLst>"#, false);
    let foreign_editor =
        SourceBackedWorkbook::from_read_at(Arc::new(VersionedSource::new(foreign))).unwrap();
    let mut output = Vec::new();
    assert!(matches!(
        foreign_editor
            .apply_sheet_protection(commit.patch())
            .map(|_| ()),
        Err(Error::PatchConflict { .. })
    ));
    assert!(output.is_empty());

    let source = Arc::new(VersionedSource::new(bytes.clone()));
    let editor = SourceBackedWorkbook::from_read_at(source.clone()).unwrap();
    let commit = changed_commit(&editor);
    source.changed();
    assert!(matches!(
        commit.package().write_to(&mut output),
        Err(Error::Package(OpcError::SourceChanged { .. }))
    ));
    assert!(output.is_empty());

    let mut retargeted = OpcPackage::from_bytes(&bytes, SourceLimits::default()).unwrap();
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
    // An owner refuses the retargeted graph at open.
    assert!(replay_patch(&mut retargeted, commit.patch()).is_err());

    let mut relationship_changed = OpcPackage::from_bytes(&bytes, SourceLimits::default()).unwrap();
    relationship_changed
        .get_part_mut(&PackURI::new(SHEET).unwrap())
        .unwrap()
        .rels_mut()
        .try_add_relationship(
            rt::HYPERLINK.to_owned(),
            "https://example.invalid/changed".to_owned(),
            "rIdChanged".to_owned(),
            TargetMode::External,
        )
        .unwrap();
    assert!(matches!(
        replay_patch(&mut relationship_changed, commit.patch()),
        Err(Error::PatchConflict { .. })
    ));
}

#[test]
fn mce_limits_and_partial_sink_are_checked() {
    let mce_sheet = format!(
        concat!(
            r#"<worksheet xmlns="{SML}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:future">"#,
            r#"<mc:AlternateContent><mc:Choice Requires="x"><x:future/></mc:Choice><mc:Fallback>"#,
            r#"<sheetProtection sheet="1"/></mc:Fallback></mc:AlternateContent></worksheet>"#,
        ),
        SML = SML,
    );
    let mce = fixture(mce_sheet, "", false);
    let editor = SourceBackedWorkbook::from_read_at(Arc::new(VersionedSource::new(mce))).unwrap();
    let mut edit = editor.edit_sheet_protection("Sheet1").unwrap();
    assert!(edit.set(target_metadata()).unwrap());
    assert!(edit.commit().is_err());

    let limits = SourceLimits::builder()
        .max_part_bytes(1)
        .unwrap()
        .build()
        .unwrap();
    assert!(matches!(
        SourceBackedWorkbook::from_read_at_with_options(
            Arc::new(VersionedSource::new(ordinary_fixture("", false))),
            litchi_xlsx::OpenOptions::new().source_limits(limits)
        ),
        Err(Error::Package(OpcError::Limit { .. }))
    ));

    let editor = SourceBackedWorkbook::from_read_at(Arc::new(VersionedSource::new(
        ordinary_fixture("", false),
    )))
    .unwrap();
    let commit = changed_commit(&editor);
    let mut sink = FailingSink {
        accepted: 0,
        limit: 128,
    };
    assert!(matches!(
        commit.package().write_to(&mut sink),
        Err(Error::Package(OpcError::IncompleteOutput { .. }))
    ));
    assert_eq!(sink.accepted, 128);
}

#[test]
fn chartsheet_selection_is_refused() {
    fn assert_send<T: Send>() {}
    assert_send::<SourceBackedWorkbook>();
    assert_send::<litchi_xlsx::sheet_protection::source::Edit>();

    const CHARTSHEET_REL: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet";
    const CHARTSHEET_CT: &str =
        "application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml";

    let mut package =
        OpcPackage::from_bytes(&ordinary_fixture("", false), SourceLimits::default()).unwrap();
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
    let editor = SourceBackedWorkbook::from_read_at(Arc::new(VersionedSource::new(bytes))).unwrap();
    assert!(matches!(
        editor.edit_sheet_protection("Sheet1").unwrap().commit(),
        Err(Error::NotWorksheet { .. })
    ));
}

/// Replay a source-backed patch onto `package` through a source-backed
/// workbook over its bytes (record 0360: source-backed patches replay onto
/// source-backed workbooks).
fn replay_patch(
    package: &mut OpcPackage,
    patch: &litchi_xlsx::sheet_protection::source::Patch,
) -> litchi_xlsx::Result<()> {
    let bytes = PackageWriter::to_bytes(package)?;
    let workbook =
        SourceBackedWorkbook::from_read_at(Arc::new(litchi_core::OwnedSource::new(bytes)))?;
    let next = workbook
        .apply_sheet_protection(patch)?
        .into_package()
        .to_bytes()?;
    *package = OpcPackage::from_bytes(&next, SourceLimits::default())?;
    Ok(())
}
