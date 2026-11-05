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
use litchi_xlsx::page_breaks::{Break, Collection};
use litchi_xlsx::{Error, OpenOptions, Package, SourceBackedWorkbook, SourceLimits};

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
        Ok(SourceVersion::new(53, self.revision.load(Ordering::SeqCst)))
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

fn changed_edit(editor: &SourceBackedWorkbook) -> litchi_xlsx::page_breaks::source::Edit {
    let mut edit = editor.edit_page_breaks("Sheet1").unwrap();
    assert!(
        edit.set_horizontal(
            Collection::horizontal([Break::new(100, 0, 16_383).unwrap().with_manual(true)])
                .unwrap()
        )
        .unwrap()
    );
    assert!(
        edit.set_vertical(Collection::vertical([Break::new(4, 2, 90).unwrap()]).unwrap())
            .unwrap()
    );
    edit
}

fn changed_commit(editor: &SourceBackedWorkbook) -> litchi_xlsx::page_breaks::source::Commit {
    let commit = changed_edit(editor).commit().unwrap();
    assert!(commit.diagnostics().changed());
    assert!(!commit.patch().is_noop());
    commit
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
        "xlsx-page-breaks-managed-test",
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
    let editor = SourceBackedWorkbook::from_read_at_with_options(
        Arc::new(VersionedSource::new(bytes.clone())),
        OpenOptions::new()
            .source_limits(SourceLimits::default())
            .cache_limits(SourceCacheLimits::new(usize::try_from(exact).unwrap(), 8).unwrap())
            .execution_context(context),
    )
    .unwrap();
    assert!(editor.cache_diagnostics().budget_managed);
    // A managed source-backed workbook pins its workbook catalog.
    assert_eq!(budget.used(Resource::Memory), part_len(&bytes, MAIN));
    let commit = changed_commit(&editor);
    assert_eq!(budget.used(Resource::Memory), exact);
    // A managed patch replays onto an owning package, whose inverse would
    // detach the managed allocation.
    let replay = Package::from_bytes(&bytes)
        .unwrap()
        .apply_page_breaks(&commit.patch().clone().into())
        .unwrap()
        .into_package();
    assert!(matches!(
        replay.apply_page_breaks(&commit.patch().inverse().into()),
        Err(Error::Package(OpcError::ManagedPartDataArcEscape))
    ));
    let mut output = Vec::new();
    commit.package().write_to(&mut output).unwrap();
    assert_eq!(
        litchi_xlsx::page_breaks::Snapshot::load(
            &OpcPackage::from_bytes(&output, SourceLimits::default()).unwrap(),
            "Sheet1",
        )
        .unwrap()
        .page_breaks(),
        commit.patch().after().page_breaks()
    );
    drop(commit);
    drop(replay);
    // The exact-capacity cache may keep the selected worksheet warm, never
    // more than the budget.
    assert!(budget.used(Resource::Memory) <= exact);
    drop(editor);
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

#[test]
fn changed_edit_reopens_and_changes_only_the_selected_worksheet() {
    let source_bytes = ordinary_fixture(r#"<extLst><ext uri="urn:keep"/></extLst>"#, false);
    let editor =
        SourceBackedWorkbook::from_read_at(Arc::new(VersionedSource::new(source_bytes.clone())))
            .unwrap();
    let opened = editor.cache_diagnostics().successful_loads;
    let commit = changed_commit(&editor);
    // Only the selected worksheet is loaded by the edit.
    assert_eq!(editor.cache_diagnostics().successful_loads, opened + 1);

    let replay = Package::from_bytes(&source_bytes)
        .unwrap()
        .apply_page_breaks(&commit.patch().clone().into())
        .unwrap()
        .into_package();
    assert_eq!(
        replay.page_breaks("Sheet1").unwrap().page_breaks(),
        commit.patch().after().page_breaks()
    );
    let replay = replay
        .apply_page_breaks(&commit.patch().inverse().into())
        .unwrap()
        .into_package()
        .into_plain_opc();
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
    let published = commit
        .package()
        .edit_page_breaks("Sheet1")
        .unwrap()
        .before()
        .clone();
    assert_eq!(
        published.page_breaks(),
        commit.patch().after().page_breaks()
    );
    // A second write of the committed workbook is byte-identical.
    assert_eq!(commit.package().to_bytes().unwrap(), output);

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
    let page_breaks = reopened.page_breaks("Sheet1").unwrap();
    assert_eq!(
        page_breaks.page_breaks(),
        commit.patch().after().page_breaks()
    );
}

#[test]
fn noop_is_byte_exact_and_signed_change_is_refused() {
    let signed = ordinary_fixture("", true);
    let editor =
        SourceBackedWorkbook::from_read_at(Arc::new(VersionedSource::new(signed.clone()))).unwrap();
    let noop = editor.edit_page_breaks("Sheet1").unwrap().commit().unwrap();
    assert!(!noop.diagnostics().changed());
    assert!(noop.patch().is_noop());
    let mut output = Vec::new();
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
    // The patch replays onto a second open of the same bytes...
    let same =
        SourceBackedWorkbook::from_read_at(Arc::new(VersionedSource::new(bytes.clone()))).unwrap();
    let replayed = same.apply_page_breaks(commit.patch()).unwrap();
    assert_eq!(
        replayed.package().to_bytes().unwrap(),
        commit.package().to_bytes().unwrap()
    );
    // ...but not onto a foreign workbook.
    assert!(matches!(
        foreign_editor.apply_page_breaks(commit.patch()),
        Err(Error::PatchConflict { .. })
    ));

    let source = Arc::new(VersionedSource::new(bytes.clone()));
    let editor = SourceBackedWorkbook::from_read_at(source.clone()).unwrap();
    let commit = changed_commit(&editor);
    source.changed();
    let mut output = Vec::new();
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
    // An owner refuses a retargeted graph that references one worksheet
    // twice, so no owner over it holds the patch's before state.
    assert!(Package::from_opc(retargeted).is_err());
}

#[test]
fn mce_limits_and_partial_sink_are_checked() {
    let mce_sheet = format!(
        concat!(
            r#"<worksheet xmlns="{SML}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:future">"#,
            r#"<mc:AlternateContent><mc:Choice Requires="x"><x:future/></mc:Choice><mc:Fallback>"#,
            r#"<rowBreaks count="1"><brk id="6"/></rowBreaks>"#,
            r#"</mc:Fallback></mc:AlternateContent></worksheet>"#,
        ),
        SML = SML,
    );
    let mce = fixture(mce_sheet, "", false);
    let editor = SourceBackedWorkbook::from_read_at(Arc::new(VersionedSource::new(mce))).unwrap();
    let mut edit = editor.edit_page_breaks("Sheet1").unwrap();
    edit.clear();
    assert!(edit.commit().is_err());

    let limits = SourceLimits::builder()
        .max_part_bytes(1)
        .unwrap()
        .build()
        .unwrap();
    assert!(matches!(
        SourceBackedWorkbook::from_read_at_with_options(
            Arc::new(VersionedSource::new(ordinary_fixture("", false))),
            OpenOptions::new().source_limits(limits),
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
