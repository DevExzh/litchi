#![allow(clippy::unwrap_used, reason = "focused integration assertions")]

use std::io;
use std::num::{NonZeroU64, NonZeroUsize};
use std::sync::Arc;
use std::sync::atomic::{AtomicU64, Ordering};

use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionError, ExecutionLimits, Limits, ReadAt,
    Resource, SourceVersion,
};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{
    BlobPart, OpcError, OpcPackage, PackURI, PackageWriter, SourceBackedPackage, SourceCacheLimits,
    TargetMode,
};
use litchi_xlsx::row_visibility::{
    MAX_BATCH_EDITS, RowVisibilityEdit, Snapshot, SourceBackedEditor,
};
use litchi_xlsx::{Error, RowIndex};

const SML: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const MAIN: &str = "/xl/workbook.xml";
const SHEET: &str = "/xl/worksheets/sheet1.xml";
const UNUSED: &str = "/xl/media/unused.bin";
// Changed publication retains the selected payload and reserves bounded OPC
// topology metadata. Keep this allowance explicit and separate from payload
// capacity, matching the managed XLSX publication contract.
const MANAGED_PUBLICATION_PLANNING_HEADROOM: u64 = 64 * 1024;
static NEXT_SOURCE_ID: AtomicU64 = AtomicU64::new(20_000);

struct VersionedSource {
    bytes: Vec<u8>,
    id: u64,
    revision: AtomicU64,
    rejected_read_offset: AtomicU64,
    allowed_matching_reads: AtomicU64,
    matching_read_count: AtomicU64,
}

impl VersionedSource {
    fn new(bytes: Vec<u8>) -> Self {
        Self {
            bytes,
            id: NEXT_SOURCE_ID.fetch_add(1, Ordering::Relaxed),
            revision: AtomicU64::new(0),
            rejected_read_offset: AtomicU64::new(u64::MAX),
            allowed_matching_reads: AtomicU64::new(0),
            matching_read_count: AtomicU64::new(0),
        }
    }

    fn changed(&self) {
        self.revision.fetch_add(1, Ordering::SeqCst);
    }

    fn reject_small_read_at(&self, offset: u64) {
        self.allowed_matching_reads.store(0, Ordering::SeqCst);
        self.matching_read_count.store(0, Ordering::SeqCst);
        self.rejected_read_offset.store(offset, Ordering::SeqCst);
    }

    fn allow_one_small_read_at(&self, offset: u64) {
        self.allowed_matching_reads.store(1, Ordering::SeqCst);
        self.matching_read_count.store(0, Ordering::SeqCst);
        self.rejected_read_offset.store(offset, Ordering::SeqCst);
    }

    fn matching_small_read_count(&self) -> u64 {
        self.matching_read_count.load(Ordering::SeqCst)
    }
}

impl ReadAt for VersionedSource {
    fn len(&self) -> io::Result<u64> {
        Ok(self.bytes.len() as u64)
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        if offset == self.rejected_read_offset.load(Ordering::SeqCst) && output.len() < 64 * 1024 {
            let matching = self.matching_read_count.fetch_add(1, Ordering::SeqCst);
            if matching >= self.allowed_matching_reads.load(Ordering::SeqCst) {
                return Err(io::Error::other("selected worksheet payload read rejected"));
            }
        }
        let offset = usize::try_from(offset).map_err(|_| io::Error::other("offset"))?;
        if offset >= self.bytes.len() {
            return Ok(0);
        }
        let count = output.len().min(self.bytes.len() - offset);
        output[..count].copy_from_slice(&self.bytes[offset..offset + count]);
        Ok(count)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        Ok(SourceVersion::new(
            self.id,
            self.revision.load(Ordering::SeqCst),
        ))
    }
}

fn row(index: u32) -> RowIndex {
    RowIndex::new(index).unwrap()
}

fn fixture(sheet_xml: String, main_type: &str, signed: bool) -> Vec<u8> {
    let workbook = format!(
        r#"<workbook xmlns="{SML}" xmlns:r="{REL}"><bookViews><workbookView/></bookViews><sheets><sheet name="Sheet1" sheetId="1" r:id="rIdSheet"/></sheets></workbook>"#
    );
    let mut package = OpcPackage::new();
    package
        .try_add_part(Box::new(BlobPart::new(
            PackURI::new(MAIN).unwrap(),
            main_type.to_owned(),
            workbook.into_bytes(),
        )))
        .unwrap();
    package
        .try_add_part(Box::new(BlobPart::new(
            PackURI::new(SHEET).unwrap(),
            ct::SML_WORKSHEET.to_owned(),
            sheet_xml.into_bytes(),
        )))
        .unwrap();
    package
        .try_add_part(Box::new(BlobPart::new(
            PackURI::new(UNUSED).unwrap(),
            "application/octet-stream".to_owned(),
            (0..128 * 1024).map(|value| (value % 251) as u8).collect(),
        )))
        .unwrap();
    package
        .get_part_mut(&PackURI::new(MAIN).unwrap())
        .unwrap()
        .rels_mut()
        .try_add_relationship(
            rt::WORKSHEET.to_owned(),
            "worksheets/sheet1.xml".to_owned(),
            "rIdSheet".to_owned(),
            TargetMode::Internal,
        )
        .unwrap();
    package.relate_to("xl/workbook.xml", rt::OFFICE_DOCUMENT);
    if signed {
        package
            .try_add_part(Box::new(BlobPart::new(
                PackURI::new("/_xmlsignatures/origin.sigs").unwrap(),
                ct::OPC_DIGITAL_SIGNATURE_ORIGIN.to_owned(),
                b"<origin/>".to_vec(),
            )))
            .unwrap();
        package.relate_to("_xmlsignatures/origin.sigs", rt::DIGITAL_SIGNATURE_ORIGIN);
    }
    PackageWriter::to_bytes(&package).unwrap()
}

fn ordinary(sheet_xml: String) -> Vec<u8> {
    fixture(sheet_xml, ct::SML_SHEET_MAIN, false)
}

fn exact_rows() -> (String, String) {
    let before = format!(
        r#"<worksheet xmlns="{SML}"><dimension ref="A1:C3"/><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetData><row r='1' spans='1:1' ht='20' customHeight='1' hidden='true'><c r="A1"><v>1</v></c></row><row r="2"><c r='B2' t="b"><v>1</v></c></row><row r="3" hidden="0"><c r="C3" t='inlineStr'><is><t xml:space="preserve"> keep </t></is></c></row></sheetData></worksheet>"#
    );
    let after = format!(
        r#"<worksheet xmlns="{SML}"><dimension ref="A1:C3"/><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetData><row r='1' spans='1:1' ht='20' customHeight='1' hidden="1"><c r="A1"><v>1</v></c></row><row r="2" hidden="1"><c r='B2' t="b"><v>1</v></c></row><row r="3"><c r="C3" t='inlineStr'><is><t xml:space="preserve"> keep </t></is></c></row></sheetData></worksheet>"#
    );
    (before, after)
}

fn oversized_row_worksheet_xml(padding: usize) -> String {
    let mut sheet = String::with_capacity(padding + 512);
    sheet.push_str(&format!(r#"<worksheet xmlns="{SML}">"#));
    let mut remaining = padding;
    while remaining != 0 {
        let chunk = remaining.min(1024 * 1024);
        sheet.push_str("<!--");
        sheet.extend(std::iter::repeat_n('x', chunk));
        sheet.push_str("-->");
        remaining -= chunk;
    }
    sheet.push_str(r#"<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>"#);
    sheet
}

fn zip_member_data_offset(bytes: &[u8], member: &[u8]) -> u64 {
    for offset in 0..bytes.len().saturating_sub(30) {
        if bytes.get(offset..offset + 4) != Some(b"PK\x03\x04") {
            continue;
        }
        let name_len = usize::from(u16::from_le_bytes([bytes[offset + 26], bytes[offset + 27]]));
        let extra_len = usize::from(u16::from_le_bytes([bytes[offset + 28], bytes[offset + 29]]));
        let name_start = offset + 30;
        let name_end = name_start.checked_add(name_len).unwrap();
        if bytes.get(name_start..name_end) == Some(member) {
            return u64::try_from(name_end + extra_len).unwrap();
        }
    }
    panic!("ZIP member local header not found")
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
        "xlsx-row-visibility-managed-test",
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

fn worksheet_xml(package: &OpcPackage) -> &[u8] {
    package
        .get_part(&PackURI::new(SHEET).unwrap())
        .unwrap()
        .blob()
}

#[test]
fn batch_changes_only_hidden_owners_and_exact_inverse_restores_source() {
    let (before_xml, after_xml) = exact_rows();
    let source_bytes = ordinary(before_xml.clone());
    let editor =
        SourceBackedEditor::from_read_at(Arc::new(VersionedSource::new(source_bytes.clone())))
            .unwrap();
    let mut edit = editor.edit("Sheet1").unwrap();
    edit.apply_batch([
        RowVisibilityEdit::hide(row(0)),
        RowVisibilityEdit::hide(row(1)),
        RowVisibilityEdit::unhide(row(2)),
    ])
    .unwrap();
    let commit = edit.commit().unwrap();
    assert!(commit.changed());
    assert_eq!(commit.diagnostics().changed_rows(), 3);
    assert_eq!(commit.diagnostics().touched_worksheets(), 1);
    assert_eq!(commit.snapshot().source_xml(), after_xml.as_bytes());
    assert_eq!(commit.snapshot().is_hidden(row(0)), Some(true));
    assert_eq!(commit.snapshot().is_hidden(row(1)), Some(true));
    assert_eq!(commit.snapshot().is_hidden(row(2)), Some(false));

    let mut replay = OpcPackage::from_bytes(&source_bytes).unwrap();
    let unused = replay
        .get_part(&PackURI::new(UNUSED).unwrap())
        .unwrap()
        .blob()
        .to_vec();
    commit.patch().apply(&mut replay).unwrap();
    assert_eq!(worksheet_xml(&replay), after_xml.as_bytes());
    assert_eq!(
        replay
            .get_part(&PackURI::new(UNUSED).unwrap())
            .unwrap()
            .blob(),
        unused,
    );
    commit.patch().inverse().apply(&mut replay).unwrap();
    assert_eq!(worksheet_xml(&replay), before_xml.as_bytes());

    let mut published = Vec::new();
    editor
        .publish_commit_to_stream(&mut published, &commit)
        .unwrap();
    let published = OpcPackage::from_bytes(&published).unwrap();
    assert_eq!(worksheet_xml(&published), after_xml.as_bytes());
    assert_eq!(
        published
            .get_part(&PackURI::new(UNUSED).unwrap())
            .unwrap()
            .blob(),
        unused,
    );
}

#[test]
fn canonical_requests_are_exact_noops() {
    let xml = format!(
        r#"<worksheet xmlns="{SML}"><sheetData><row r="1" hidden="1"/><row r="2"/></sheetData></worksheet>"#
    );
    let bytes = ordinary(xml.clone());
    let editor = SourceBackedEditor::from_read_at(Arc::new(VersionedSource::new(bytes))).unwrap();
    let mut edit = editor.edit(0usize).unwrap();
    edit.hide(row(0)).unwrap();
    edit.unhide(row(1)).unwrap();
    let commit = edit.commit().unwrap();
    assert!(!commit.changed());
    assert!(commit.patch().is_empty());
    assert_eq!(commit.snapshot().source_xml(), xml.as_bytes());
}

#[test]
fn changed_publication_reuses_matched_provenance_without_selected_reload() {
    let bytes = ordinary(oversized_row_worksheet_xml(8 * 1024 * 1024 + 64 * 1024));
    let worksheet_data_offset = zip_member_data_offset(&bytes, b"xl/worksheets/sheet1.xml");
    let unused = OpcPackage::from_bytes(&bytes)
        .unwrap()
        .get_part(&PackURI::new(UNUSED).unwrap())
        .unwrap()
        .blob()
        .to_vec();
    let source = Arc::new(VersionedSource::new(bytes));
    let editor = SourceBackedEditor::from_read_at(source.clone()).unwrap();
    let mut edit = editor.edit("Sheet1").unwrap();
    edit.hide(row(0)).unwrap();
    let commit = edit.commit().unwrap();
    assert!(commit.changed());

    // Publication must retain the one selected-member read needed by the OPC
    // overlay path, but a second semantic snapshot reload would exceed this
    // allowance and fail before output completes.
    source.allow_one_small_read_at(worksheet_data_offset);

    let mut output = Vec::new();
    let published_snapshot = editor
        .publish_commit_to_stream(&mut output, &commit)
        .unwrap();
    assert_eq!(source.matching_small_read_count(), 1);
    assert_eq!(published_snapshot.is_hidden(row(0)), Some(true));

    let published = OpcPackage::from_bytes(&output).unwrap();
    let reopened = Snapshot::load(&published, "Sheet1").unwrap();
    assert_eq!(reopened.is_hidden(row(0)), Some(true));
    assert_eq!(
        published
            .get_part(&PackURI::new(UNUSED).unwrap())
            .unwrap()
            .blob(),
        unused,
    );
}

#[test]
fn xml_boolean_forms_and_self_closing_layout_state_are_owned_narrowly() {
    let before = format!(
        r#"<worksheet xmlns="{SML}"><sheetData><row r="1" hidden="true" outlineLevel="2" collapsed="1"/><row r="2" hidden="false" ht="22" customHeight="1"/><row r="3"/></sheetData></worksheet>"#
    );
    let bytes = ordinary(before.clone());
    let editor = SourceBackedEditor::from_read_at(Arc::new(VersionedSource::new(bytes))).unwrap();
    let source = editor.snapshot("Sheet1").unwrap();
    assert_eq!(source.is_hidden(row(0)), Some(true));
    assert_eq!(source.is_hidden(row(1)), Some(false));
    assert_eq!(source.is_hidden(row(2)), Some(false));

    let mut edit = editor.edit("Sheet1").unwrap();
    edit.hide(row(0)).unwrap();
    edit.unhide(row(1)).unwrap();
    edit.unhide(row(2)).unwrap();
    let commit = edit.commit().unwrap();
    assert_eq!(commit.diagnostics().changed_rows(), 2);
    assert_eq!(
        commit.snapshot().source_xml(),
        format!(
            r#"<worksheet xmlns="{SML}"><sheetData><row r="1" outlineLevel="2" collapsed="1" hidden="1"/><row r="2" ht="22" customHeight="1"/><row r="3"/></sheetData></worksheet>"#
        )
        .as_bytes()
    );
}

#[test]
fn duplicate_missing_and_oversized_batches_are_atomic() {
    let rows = (1..=MAX_BATCH_EDITS + 1)
        .map(|number| format!(r#"<row r="{number}"/>"#))
        .collect::<String>();
    let xml = format!(r#"<worksheet xmlns="{SML}"><sheetData>{rows}</sheetData></worksheet>"#);
    let bytes = ordinary(xml);
    let editor = SourceBackedEditor::from_read_at(Arc::new(VersionedSource::new(bytes))).unwrap();
    let mut edit = editor.edit("Sheet1").unwrap();
    assert!(edit.hide(row(400)).is_err());
    assert_eq!(edit.len(), 0);
    assert!(
        edit.apply_batch([
            RowVisibilityEdit::hide(row(0)),
            RowVisibilityEdit::unhide(row(0))
        ])
        .is_err()
    );
    assert_eq!(edit.len(), 0);
    let too_many = (0..=MAX_BATCH_EDITS)
        .map(|index| RowVisibilityEdit::hide(row(u32::try_from(index).unwrap())));
    assert!(edit.apply_batch(too_many).is_err());
    assert_eq!(edit.len(), 0);
}

#[test]
fn stale_package_and_source_revision_are_refused_without_output() {
    let xml =
        format!(r#"<worksheet xmlns="{SML}"><sheetData><row r="1"/></sheetData></worksheet>"#);
    let bytes = ordinary(xml);
    let source = Arc::new(VersionedSource::new(bytes.clone()));
    let editor = SourceBackedEditor::from_read_at(source.clone()).unwrap();
    let mut edit = editor.edit("Sheet1").unwrap();
    edit.hide(row(0)).unwrap();
    let commit = edit.commit().unwrap();

    let mut stale = OpcPackage::from_bytes(&bytes).unwrap();
    stale
        .get_part_mut(&PackURI::new(SHEET).unwrap())
        .unwrap()
        .set_blob(b"<different/>".to_vec());
    let before = worksheet_xml(&stale).to_vec();
    assert!(matches!(
        commit.patch().apply(&mut stale),
        Err(Error::PatchConflict { .. })
    ));
    assert_eq!(worksheet_xml(&stale), before);

    source.changed();
    let mut output = Vec::new();
    assert!(matches!(
        editor.publish_commit_to_stream(&mut output, &commit),
        Err(Error::Package(OpcError::SourceChanged { .. }))
    ));
    assert!(output.is_empty());
}

#[test]
fn unsafe_xml_macros_and_signatures_fail_closed() {
    for xml in [
        format!(
            r#"<worksheet xmlns="{SML}"><sheetProtection sheet="1"/><sheetData><row r="1"/></sheetData></worksheet>"#
        ),
        format!(
            r#"<worksheet xmlns="{SML}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x"><sheetData><row r="1"/></sheetData></worksheet>"#
        ),
        format!(
            r#"<worksheet xmlns="{SML}"><sheetData><row r="1"><c r="A1"><f>A2</f><v>1</v></c></row></sheetData></worksheet>"#
        ),
        format!(
            r#"<worksheet xmlns="{SML}"><sheetData><row r="1"/><row r="1"/></sheetData></worksheet>"#
        ),
        format!(
            r#"<worksheet xmlns="{SML}"><sheetData><row r="1" hidden="yes"/></sheetData></worksheet>"#
        ),
    ] {
        let editor =
            SourceBackedEditor::from_read_at(Arc::new(VersionedSource::new(ordinary(xml))))
                .unwrap();
        assert!(editor.edit("Sheet1").is_err());
    }

    let plain =
        format!(r#"<worksheet xmlns="{SML}"><sheetData><row r="1"/></sheetData></worksheet>"#);
    let macro_bytes = fixture(plain.clone(), ct::SML_SHEET_MACRO_MAIN, false);
    let macro_editor =
        SourceBackedEditor::from_read_at(Arc::new(VersionedSource::new(macro_bytes))).unwrap();
    assert!(macro_editor.edit("Sheet1").is_err());

    let signed_bytes = fixture(plain, ct::SML_SHEET_MAIN, true);
    let signed_editor =
        SourceBackedEditor::from_read_at(Arc::new(VersionedSource::new(signed_bytes.clone())))
            .unwrap();
    let mut edit = signed_editor.edit("Sheet1").unwrap();
    edit.hide(row(0)).unwrap();
    let commit = edit.commit().unwrap();
    let mut signed = OpcPackage::from_bytes(&signed_bytes).unwrap();
    assert!(matches!(
        commit.patch().apply(&mut signed),
        Err(Error::Signed)
    ));
}

#[test]
fn public_snapshot_distinguishes_missing_and_visible_row_owners() {
    let xml = format!(
        r#"<worksheet xmlns="{SML}"><sheetData><row r="2" hidden="false"/></sheetData></worksheet>"#
    );
    let package = OpcPackage::from_bytes(&ordinary(xml)).unwrap();
    let snapshot = Snapshot::load(&package, "Sheet1").unwrap();
    assert_eq!(snapshot.is_hidden(row(0)), None);
    assert_eq!(snapshot.is_hidden(row(1)), Some(false));
    assert!(snapshot.contains_row(row(1)));
}

#[test]
fn managed_editor_retains_payload_closure_and_releases_budget() {
    let (before_xml, _after_xml) = exact_rows();
    let bytes = ordinary(before_xml);
    let exact = part_len(&bytes, MAIN) + part_len(&bytes, SHEET);
    let (budget, _cancellation_source, context) = managed_context(exact);
    let editor =
        SourceBackedEditor::from_read_at_with_limits_and_cache_limits_and_execution_context(
            Arc::new(VersionedSource::new(bytes)),
            litchi_xlsx::ReadLimits::default(),
            SourceCacheLimits::new(usize::try_from(exact).unwrap(), 8).unwrap(),
            context,
        )
        .unwrap();

    assert!(editor.cache_diagnostics().budget_managed);
    assert_eq!(budget.used(Resource::Memory), 0);
    let snapshot = editor.snapshot("Sheet1").unwrap();
    assert_eq!(snapshot.is_hidden(row(0)), Some(true));
    assert_eq!(snapshot.is_hidden(row(1)), Some(false));
    assert_eq!(budget.used(Resource::Memory), exact);
    assert_eq!(editor.cache_diagnostics().retained_entries, 2);

    drop(snapshot);
    assert_eq!(budget.used(Resource::Memory), exact);
    drop(editor);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn managed_one_under_budget_refuses_before_selected_worksheet_io() {
    let (before_xml, _after_xml) = exact_rows();
    let bytes = ordinary(before_xml);
    let exact = part_len(&bytes, MAIN) + part_len(&bytes, SHEET);
    let worksheet_data_offset = zip_member_data_offset(&bytes, b"xl/worksheets/sheet1.xml");
    let source = Arc::new(VersionedSource::new(bytes));
    source.reject_small_read_at(worksheet_data_offset);
    let (budget, _cancellation_source, context) = managed_context(exact - 1);
    let editor = SourceBackedEditor::from_read_at_with_execution_context(
        source.clone(),
        litchi_xlsx::ReadLimits::default(),
        context,
    )
    .unwrap();

    assert_eq!(budget.used(Resource::Memory), 0);
    let result = editor.snapshot("Sheet1");
    assert!(matches!(
        result,
        Err(Error::Package(OpcError::Execution(
            ExecutionError::ResourceLimit(_)
        )))
    ));
    assert_eq!(source.matching_small_read_count(), 0);
    assert_eq!(editor.cache_diagnostics().successful_loads, 1);

    drop(editor);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn managed_cancellation_is_checked_before_snapshot_and_stream_output() {
    let (before_xml, _after_xml) = exact_rows();
    let bytes = ordinary(before_xml);
    let exact = part_len(&bytes, MAIN) + part_len(&bytes, SHEET);
    let (budget, cancellation_source, context) = managed_context(exact);
    let editor = SourceBackedEditor::from_read_at_with_execution_context(
        Arc::new(VersionedSource::new(bytes.clone())),
        litchi_xlsx::ReadLimits::default(),
        context,
    )
    .unwrap();
    cancellation_source.cancel();
    assert!(matches!(
        editor.snapshot("Sheet1"),
        Err(Error::Package(OpcError::Cancelled))
    ));
    drop(editor);
    assert_eq!(budget.used(Resource::Memory), 0);

    let (budget, cancellation_source, context) = managed_context(exact);
    let editor = SourceBackedEditor::from_read_at_with_execution_context(
        Arc::new(VersionedSource::new(bytes.clone())),
        litchi_xlsx::ReadLimits::default(),
        context,
    )
    .unwrap();
    let mut edit = editor.edit("Sheet1").unwrap();
    edit.hide(row(1)).unwrap();
    let commit = edit.commit().unwrap();
    cancellation_source.cancel();
    let mut output = Vec::new();
    assert!(matches!(
        editor.publish_commit_to_stream(&mut output, &commit),
        Err(Error::Package(OpcError::Cancelled))
    ));
    assert!(output.is_empty());
    drop(commit);
    assert_eq!(budget.used(Resource::Memory), 0);

    let (budget, cancellation_source, context) = managed_context(exact);
    let editor = SourceBackedEditor::from_read_at_with_execution_context(
        Arc::new(VersionedSource::new(bytes.clone())),
        litchi_xlsx::ReadLimits::default(),
        context,
    )
    .unwrap();
    let mut edit = editor.edit("Sheet1").unwrap();
    cancellation_source.cancel();
    assert!(matches!(
        edit.hide(row(0)),
        Err(Error::Package(OpcError::Cancelled))
    ));
    assert_eq!(edit.len(), 0);
    drop(edit);
    drop(editor);
    assert_eq!(budget.used(Resource::Memory), 0);

    let (budget, cancellation_source, context) = managed_context(exact);
    let editor = SourceBackedEditor::from_read_at_with_execution_context(
        Arc::new(VersionedSource::new(bytes)),
        litchi_xlsx::ReadLimits::default(),
        context,
    )
    .unwrap();
    let edit = editor.edit("Sheet1").unwrap();
    cancellation_source.cancel();
    assert!(matches!(
        edit.commit(),
        Err(Error::Package(OpcError::Cancelled))
    ));
    drop(editor);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn managed_exact_noop_publication_is_byte_exact_and_releases_budget() {
    let xml = format!(
        r#"<worksheet xmlns="{SML}"><sheetData><row r="1" hidden="1"/><row r="2"/></sheetData></worksheet>"#
    );
    let bytes = ordinary(xml);
    let exact = part_len(&bytes, MAIN) + part_len(&bytes, SHEET);
    let (budget, _cancellation_source, context) = managed_context(exact);
    let editor = SourceBackedEditor::from_read_at_with_execution_context(
        Arc::new(VersionedSource::new(bytes.clone())),
        litchi_xlsx::ReadLimits::default(),
        context,
    )
    .unwrap();
    let mut edit = editor.edit("Sheet1").unwrap();
    edit.hide(row(0)).unwrap();
    edit.unhide(row(1)).unwrap();
    let commit = edit.commit().unwrap();
    assert!(!commit.changed());
    assert!(commit.patch().is_empty());

    let mut output = Vec::new();
    editor
        .publish_commit_to_stream(&mut output, &commit)
        .unwrap();
    assert_eq!(output, bytes);
    drop(commit);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn managed_changed_publication_preserves_unknown_members_and_releases_budget() {
    let (before_xml, after_xml) = exact_rows();
    let before_len = before_xml.len();
    let bytes = ordinary(before_xml.clone());
    let exact = part_len(&bytes, MAIN) + part_len(&bytes, SHEET);
    let unused = OpcPackage::from_bytes(&bytes)
        .unwrap()
        .get_part(&PackURI::new(UNUSED).unwrap())
        .unwrap()
        .blob()
        .to_vec();
    let (budget, _cancellation_source, context) =
        managed_context(exact + MANAGED_PUBLICATION_PLANNING_HEADROOM);
    let editor = SourceBackedEditor::from_read_at_with_execution_context(
        Arc::new(VersionedSource::new(bytes.clone())),
        litchi_xlsx::ReadLimits::default(),
        context,
    )
    .unwrap();
    let mut edit = editor.edit("Sheet1").unwrap();
    edit.apply_batch([
        RowVisibilityEdit::hide(row(0)),
        RowVisibilityEdit::hide(row(1)),
        RowVisibilityEdit::unhide(row(2)),
    ])
    .unwrap();
    let commit = edit.commit().unwrap();
    assert!(commit.changed());
    assert_eq!(commit.snapshot().source_xml(), after_xml.as_bytes());
    assert_eq!(budget.used(Resource::Memory), exact);

    let mut replay = OpcPackage::from_bytes(&bytes).unwrap();
    commit.patch().apply(&mut replay).unwrap();
    assert_eq!(worksheet_xml(&replay), after_xml.as_bytes());
    assert!(matches!(
        commit.patch().inverse().apply(&mut replay),
        Err(Error::Package(OpcError::ManagedPartDataArcEscape))
    ));
    commit
        .patch()
        .inverse()
        .apply_materialized(&mut replay, before_len)
        .unwrap();
    assert_eq!(worksheet_xml(&replay), before_xml.as_bytes());
    assert_eq!(
        replay
            .get_part(&PackURI::new(UNUSED).unwrap())
            .unwrap()
            .blob(),
        unused,
    );

    let mut output = Vec::new();
    editor
        .publish_commit_to_stream(&mut output, &commit)
        .unwrap();
    let published = OpcPackage::from_bytes(&output).unwrap();
    assert_eq!(worksheet_xml(&published), after_xml.as_bytes());
    assert_eq!(
        published
            .get_part(&PackURI::new(UNUSED).unwrap())
            .unwrap()
            .blob(),
        unused,
    );
    drop(commit);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn managed_changed_publication_refuses_without_planning_headroom() {
    let (before_xml, _after_xml) = exact_rows();
    let bytes = ordinary(before_xml);
    let exact = part_len(&bytes, MAIN) + part_len(&bytes, SHEET);
    let (budget, _cancellation_source, context) = managed_context(exact);
    let editor = SourceBackedEditor::from_read_at_with_execution_context(
        Arc::new(VersionedSource::new(bytes)),
        litchi_xlsx::ReadLimits::default(),
        context,
    )
    .unwrap();
    let mut edit = editor.edit("Sheet1").unwrap();
    edit.hide(row(1)).unwrap();
    let commit = edit.commit().unwrap();
    assert!(commit.changed());
    assert_eq!(budget.used(Resource::Memory), exact);

    let mut output = Vec::new();
    let result = editor.publish_commit_to_stream(&mut output, &commit);
    assert!(matches!(
        result,
        Err(Error::Package(OpcError::Execution(
            ExecutionError::ResourceLimit(_)
        )))
    ));
    assert!(output.is_empty());
    drop(commit);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn managed_editor_adopts_an_indexed_package_and_releases_budget() {
    let (before_xml, _after_xml) = exact_rows();
    let bytes = ordinary(before_xml);
    let exact = part_len(&bytes, MAIN) + part_len(&bytes, SHEET);
    let (budget, _cancellation_source, context) = managed_context(exact);
    let package = SourceBackedPackage::from_read_at_with_execution_context(
        Arc::new(VersionedSource::new(bytes)),
        litchi_xlsx::ReadLimits::default(),
        context,
    )
    .unwrap();
    let editor = SourceBackedEditor::from_source_backed_package(package).unwrap();
    assert!(editor.cache_diagnostics().budget_managed);
    let snapshot = editor.snapshot("Sheet1").unwrap();
    assert_eq!(snapshot.is_hidden(row(0)), Some(true));
    assert_eq!(budget.used(Resource::Memory), exact);
    drop(snapshot);
    drop(editor);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn managed_signature_noop_and_changed_protection_contracts_remain_fail_closed() {
    let plain =
        format!(r#"<worksheet xmlns="{SML}"><sheetData><row r="1"/></sheetData></worksheet>"#);
    let signed = fixture(plain.clone(), ct::SML_SHEET_MAIN, true);
    let signed_exact = part_len(&signed, MAIN) + part_len(&signed, SHEET);
    let (budget, _cancellation_source, context) = managed_context(signed_exact);
    let editor = SourceBackedEditor::from_read_at_with_execution_context(
        Arc::new(VersionedSource::new(signed.clone())),
        litchi_xlsx::ReadLimits::default(),
        context,
    )
    .unwrap();
    let noop = editor.edit("Sheet1").unwrap().commit().unwrap();
    let mut output = Vec::new();
    editor.publish_commit_to_stream(&mut output, &noop).unwrap();
    assert_eq!(output, signed);
    drop(noop);
    assert_eq!(budget.used(Resource::Memory), 0);

    let (budget, _cancellation_source, context) = managed_context(signed_exact);
    let editor = SourceBackedEditor::from_read_at_with_execution_context(
        Arc::new(VersionedSource::new(signed)),
        litchi_xlsx::ReadLimits::default(),
        context,
    )
    .unwrap();
    let mut edit = editor.edit("Sheet1").unwrap();
    edit.hide(row(0)).unwrap();
    let commit = edit.commit().unwrap();
    let mut output = Vec::new();
    assert!(matches!(
        editor.publish_commit_to_stream(&mut output, &commit),
        Err(Error::Package(OpcError::SignedSourceRequiresExplicitPolicy))
    ));
    assert!(output.is_empty());
    drop(commit);
    assert_eq!(budget.used(Resource::Memory), 0);

    let protected = ordinary(format!(
        r#"<worksheet xmlns="{SML}"><sheetProtection sheet="1"/><sheetData><row r="1"/></sheetData></worksheet>"#
    ));
    let protected_exact = part_len(&protected, MAIN) + part_len(&protected, SHEET);
    let (budget, _cancellation_source, context) = managed_context(protected_exact);
    let editor = SourceBackedEditor::from_read_at_with_execution_context(
        Arc::new(VersionedSource::new(protected)),
        litchi_xlsx::ReadLimits::default(),
        context,
    )
    .unwrap();
    assert!(editor.edit("Sheet1").is_err());
    drop(editor);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn managed_source_revision_is_refused_before_output() {
    let (before_xml, _after_xml) = exact_rows();
    let bytes = ordinary(before_xml);
    let exact = part_len(&bytes, MAIN) + part_len(&bytes, SHEET);
    let source = Arc::new(VersionedSource::new(bytes));
    let (budget, _cancellation_source, context) = managed_context(exact);
    let editor = SourceBackedEditor::from_read_at_with_execution_context(
        source.clone(),
        litchi_xlsx::ReadLimits::default(),
        context,
    )
    .unwrap();
    let mut edit = editor.edit("Sheet1").unwrap();
    edit.hide(row(0)).unwrap();
    let commit = edit.commit().unwrap();
    source.changed();

    let mut output = Vec::new();
    assert!(matches!(
        editor.publish_commit_to_stream(&mut output, &commit),
        Err(Error::Package(OpcError::SourceChanged { .. }))
    ));
    assert!(output.is_empty());
    drop(commit);
    assert_eq!(budget.used(Resource::Memory), 0);
}
