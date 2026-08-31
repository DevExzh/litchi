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
    BlobPart, OpcError, OpcPackage, PackURI, PackageWriter, Part, Relationships, SourceCacheLimits,
    TargetMode,
};
use litchi_xlsx::cell_values::{
    CellValueEdit, MAX_BATCH_EDITS, SheetCellValueEdit, SourceBackedEditor,
};
use litchi_xlsx::{Address, Cell, Error, ErrorValue, Formula, Number, Value};
use soapberry_zip::office::ArchiveReader;

const SML: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const MAIN: &str = "/xl/workbook.xml";
const SHEET: &str = "/xl/worksheets/sheet1.xml";
const UNUSED: &str = "/xl/media/unused.bin";
const CALC_CHAIN: &str = "/xl/calcChain.xml";
const CALC_CHAIN_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml";
const CALC_CHAIN_REL_ID: &str = "rIdCalculationChain";
static NEXT_SOURCE_ID: AtomicU64 = AtomicU64::new(1_000);

struct VersionedSource {
    bytes: Vec<u8>,
    id: u64,
    revision: AtomicU64,
    read_calls: AtomicU64,
    read_bytes: AtomicU64,
    selected_worksheet_offset: AtomicU64,
    unselected_worksheet_offset: AtomicU64,
    selected_worksheet_data_reads: AtomicU64,
    unselected_worksheet_data_reads: AtomicU64,
    rejected_read_offset: AtomicU64,
    rejected_read_count: AtomicU64,
    flip_after_read_offset: AtomicU64,
    pending_version_flip: AtomicU64,
}

#[derive(Clone)]
struct CancelOnBlobArcPart {
    inner: BlobPart,
    cancellation_source: CancellationSource,
    blob_arc_calls: Arc<AtomicU64>,
    cancel_on_call: u64,
}

impl CancelOnBlobArcPart {
    fn new(
        partname: PackURI,
        content_type: String,
        blob: Vec<u8>,
        cancellation_source: CancellationSource,
        cancel_on_call: u64,
    ) -> Self {
        Self {
            inner: BlobPart::new(partname, content_type, blob),
            cancellation_source,
            blob_arc_calls: Arc::new(AtomicU64::new(0)),
            cancel_on_call,
        }
    }
}

impl Part for CancelOnBlobArcPart {
    fn blob(&self) -> &[u8] {
        self.inner.blob()
    }

    fn blob_arc(&self) -> Arc<Vec<u8>> {
        let call = self.blob_arc_calls.fetch_add(1, Ordering::SeqCst) + 1;
        if call == self.cancel_on_call {
            self.cancellation_source.cancel();
        }
        self.inner.blob_arc()
    }

    fn content_type(&self) -> &str {
        self.inner.content_type()
    }

    fn partname(&self) -> &PackURI {
        self.inner.partname()
    }

    fn rels(&self) -> &Relationships {
        self.inner.rels()
    }

    fn rels_mut(&mut self) -> &mut Relationships {
        self.inner.rels_mut()
    }

    fn set_blob(&mut self, blob: Vec<u8>) {
        self.inner.set_blob(blob);
    }
}

impl VersionedSource {
    fn new(bytes: Vec<u8>) -> Self {
        Self {
            bytes,
            id: NEXT_SOURCE_ID.fetch_add(1, Ordering::Relaxed),
            revision: AtomicU64::new(0),
            read_calls: AtomicU64::new(0),
            read_bytes: AtomicU64::new(0),
            selected_worksheet_offset: AtomicU64::new(u64::MAX),
            unselected_worksheet_offset: AtomicU64::new(u64::MAX),
            selected_worksheet_data_reads: AtomicU64::new(0),
            unselected_worksheet_data_reads: AtomicU64::new(0),
            rejected_read_offset: AtomicU64::new(u64::MAX),
            rejected_read_count: AtomicU64::new(0),
            flip_after_read_offset: AtomicU64::new(u64::MAX),
            pending_version_flip: AtomicU64::new(0),
        }
    }
    fn changed(&self) {
        self.revision.fetch_add(1, Ordering::SeqCst);
    }
    fn reject_read_at(&self, offset: u64) {
        self.rejected_read_count.store(0, Ordering::SeqCst);
        self.rejected_read_offset.store(offset, Ordering::SeqCst);
    }
    fn rejected_read_count(&self) -> u64 {
        self.rejected_read_count.load(Ordering::SeqCst)
    }
    fn flip_after_read_at(&self, offset: u64) {
        self.flip_after_read_offset.store(offset, Ordering::SeqCst);
        self.pending_version_flip.store(0, Ordering::SeqCst);
    }
    fn set_worksheet_diagnostic_offsets(&self, selected: u64, unselected: u64) {
        self.selected_worksheet_offset
            .store(selected, Ordering::SeqCst);
        self.unselected_worksheet_offset
            .store(unselected, Ordering::SeqCst);
    }
    fn reset_read_diagnostics(&self) {
        self.read_calls.store(0, Ordering::SeqCst);
        self.read_bytes.store(0, Ordering::SeqCst);
        self.selected_worksheet_data_reads
            .store(0, Ordering::SeqCst);
        self.unselected_worksheet_data_reads
            .store(0, Ordering::SeqCst);
    }
    fn read_diagnostics(&self) -> (u64, u64, u64, u64) {
        (
            self.read_calls.load(Ordering::SeqCst),
            self.read_bytes.load(Ordering::SeqCst),
            self.selected_worksheet_data_reads.load(Ordering::SeqCst),
            self.unselected_worksheet_data_reads.load(Ordering::SeqCst),
        )
    }
}

impl ReadAt for VersionedSource {
    fn len(&self) -> io::Result<u64> {
        Ok(self.bytes.len() as u64)
    }
    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        self.read_calls.fetch_add(1, Ordering::SeqCst);
        if offset == self.selected_worksheet_offset.load(Ordering::SeqCst) {
            self.selected_worksheet_data_reads
                .fetch_add(1, Ordering::SeqCst);
        }
        if offset == self.unselected_worksheet_offset.load(Ordering::SeqCst) {
            self.unselected_worksheet_data_reads
                .fetch_add(1, Ordering::SeqCst);
        }
        let flip_offset = self.flip_after_read_offset.load(Ordering::SeqCst);
        if offset == flip_offset
            && self
                .flip_after_read_offset
                .compare_exchange(flip_offset, u64::MAX, Ordering::SeqCst, Ordering::SeqCst)
                .is_ok()
        {
            // Delay the flip until the selected Part's post-load checks have
            // observed the old revision; the facade's final post-parse check
            // must observe the new one.
            self.pending_version_flip.store(3, Ordering::SeqCst);
        }
        if offset == self.rejected_read_offset.load(Ordering::SeqCst) {
            self.rejected_read_count.fetch_add(1, Ordering::SeqCst);
            return Err(io::Error::other("selected worksheet payload read rejected"));
        }
        let offset = usize::try_from(offset).map_err(|_| io::Error::other("offset"))?;
        if offset >= self.bytes.len() {
            return Ok(0);
        }
        let count = output.len().min(self.bytes.len() - offset);
        output[..count].copy_from_slice(&self.bytes[offset..offset + count]);
        self.read_bytes.fetch_add(
            u64::try_from(count).map_err(|_| io::Error::other("read size"))?,
            Ordering::SeqCst,
        );
        Ok(count)
    }
    fn version(&self) -> io::Result<SourceVersion> {
        if self.pending_version_flip.load(Ordering::SeqCst) != 0
            && self.pending_version_flip.fetch_sub(1, Ordering::SeqCst) == 1
        {
            self.revision.fetch_add(1, Ordering::SeqCst);
        }
        Ok(SourceVersion::new(
            self.id,
            self.revision.load(Ordering::SeqCst),
        ))
    }
}

fn address(value: &str) -> Address {
    Address::from_a1(value).unwrap()
}

fn fixture_package(sheet_xml: String, signed: bool) -> OpcPackage {
    let workbook = format!(
        r#"<workbook xmlns="{SML}" xmlns:r="{REL}"><bookViews><workbookView/></bookViews><sheets><sheet name="Sheet1" sheetId="1" r:id="rIdSheet"/></sheets></workbook>"#
    );
    let mut package = OpcPackage::new();
    package
        .try_add_part(Box::new(BlobPart::new(
            PackURI::new(MAIN).unwrap(),
            ct::SML_SHEET_MAIN.to_owned(),
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
            (0..256 * 1024).map(|value| (value % 251) as u8).collect(),
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
    package
}

fn fixture(sheet_xml: String, signed: bool) -> Vec<u8> {
    PackageWriter::to_bytes(&fixture_package(sheet_xml, signed)).unwrap()
}

fn replace_sheet_with_cancel_on_blob_arc(
    package: &mut OpcPackage,
    cancellation_source: CancellationSource,
    cancel_on_call: u64,
) {
    let uri = PackURI::new(SHEET).unwrap();
    let (content_type, blob) = {
        let part = package.get_part(&uri).unwrap();
        (part.content_type().to_owned(), part.blob().to_vec())
    };
    assert!(package.remove_part(&uri));
    package.add_part(Box::new(CancelOnBlobArcPart::new(
        uri,
        content_type,
        blob,
        cancellation_source,
        cancel_on_call,
    )));
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
        "xlsx-cell-values-managed-test",
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

fn with_style_and_theme(bytes: &[u8]) -> Vec<u8> {
    let mut package = OpcPackage::from_bytes(bytes).unwrap();
    package
        .try_add_part(Box::new(BlobPart::new(
            PackURI::new("/xl/theme/theme1.xml").unwrap(),
            ct::OFC_THEME.to_owned(),
            b"<theme/>".to_vec(),
        )))
        .unwrap();
    package
        .get_part_mut(&PackURI::new(MAIN).unwrap())
        .unwrap()
        .rels_mut()
        .try_add_relationship(
            rt::THEME.to_owned(),
            "theme/theme1.xml".to_owned(),
            "rIdTheme".to_owned(),
            TargetMode::Internal,
        )
        .unwrap();
    PackageWriter::to_bytes(&package).unwrap()
}

fn oversized_worksheet_xml(padding: usize, value: &str) -> Vec<u8> {
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
    sheet.push_str(&format!(
        r#"<sheetData><row r="1"><c r="A1"><v>{value}</v></c></row></sheetData></worksheet>"#
    ));
    sheet.into_bytes()
}

fn oversized_multi_sheet_source() -> Vec<u8> {
    let mut package = OpcPackage::from_bytes(&two_sheets()).unwrap();
    package
        .get_part_mut(&PackURI::new("/xl/worksheets/sheet1.xml").unwrap())
        .unwrap()
        // Keep Sheet1 above the default 8 MiB source payload-cache limit so
        // an independent semantic reload must perform physical ZIP I/O.
        .set_blob(oversized_worksheet_xml(8 * 1024 * 1024 + 64 * 1024, "1"));
    PackageWriter::to_bytes(&package).unwrap()
}

fn three_cells() -> Vec<u8> {
    fixture(
        format!(
            r#"<worksheet xmlns="{SML}"><dimension ref="A1:C3"/><sheetViews><sheetView workbookViewId="0"><selection activeCell="A1" sqref="A1"/></sheetView></sheetViews><sheetData><row r="1"><c r="A1"><v>1</v></c></row><row r="2"><c r="B2" t="b"><v>1</v></c></row><row r="3"><c r="C3" t="inlineStr"><is><t>old</t></is></c></row></sheetData></worksheet>"#
        ),
        false,
    )
}

fn two_sheets_with_signature(signed: bool) -> Vec<u8> {
    let mut package = fixture_package(
        format!(
            r#"<worksheet xmlns="{SML}"><dimension ref="A1:C3"/><sheetViews><sheetView workbookViewId="0"><selection activeCell="A1" sqref="A1"/></sheetView></sheetViews><sheetData><row r="1"><c r="A1"><v>1</v></c></row><row r="2"><c r="B2" t="b"><v>1</v></c></row><row r="3"><c r="C3" t="inlineStr"><is><t>old</t></is></c></row></sheetData></worksheet>"#
        ),
        signed,
    );
    package
        .try_add_part(Box::new(BlobPart::new(
            PackURI::new("/xl/worksheets/sheet2.xml").unwrap(),
            ct::SML_WORKSHEET.to_owned(),
            format!(
                r#"<worksheet xmlns="{SML}"><dimension ref="A1:C3"/><sheetData><row r="1"><c r="A1"><v>20</v></c></row><row r="2"><c r="B2" t="b"><v>1</v></c></row><row r="3"><c r="C3" t="inlineStr"><is><t>second</t></is></c></row></sheetData></worksheet>"#
            )
            .into_bytes(),
        )))
        .unwrap();
    package
        .get_part_mut(&PackURI::new(MAIN).unwrap())
        .unwrap()
        .set_blob(
            format!(
                r#"<workbook xmlns="{SML}" xmlns:r="{REL}"><bookViews><workbookView/></bookViews><sheets><sheet name="Sheet1" sheetId="1" r:id="rIdSheet"/><sheet name="Sheet2" sheetId="2" r:id="rIdSheet2"/></sheets></workbook>"#
            )
            .into_bytes(),
        );
    package
        .get_part_mut(&PackURI::new(MAIN).unwrap())
        .unwrap()
        .rels_mut()
        .try_add_relationship(
            rt::WORKSHEET.to_owned(),
            "worksheets/sheet2.xml".to_owned(),
            "rIdSheet2".to_owned(),
            TargetMode::Internal,
        )
        .unwrap();
    PackageWriter::to_bytes(&package).unwrap()
}

fn two_sheets() -> Vec<u8> {
    two_sheets_with_signature(false)
}

fn signed_two_sheets() -> Vec<u8> {
    two_sheets_with_signature(true)
}

fn with_two_cell_styles(bytes: &[u8]) -> (Vec<u8>, PackURI, Vec<u8>) {
    let mut package = OpcPackage::from_bytes(bytes).unwrap();
    let style_uri = PackURI::new("/xl/styles.xml").unwrap();
    let style_xml = format!(
        r#"<styleSheet xmlns="{SML}"><cellXfs count="2"><xf/><xf numFmtId="1"/></cellXfs></styleSheet>"#
    )
    .into_bytes();
    package
        .try_add_part(Box::new(BlobPart::new(
            style_uri.clone(),
            ct::SML_STYLES.to_owned(),
            style_xml.clone(),
        )))
        .unwrap();
    package
        .get_part_mut(&PackURI::new(MAIN).unwrap())
        .unwrap()
        .rels_mut()
        .try_add_relationship(
            rt::STYLES.to_owned(),
            "styles.xml".to_owned(),
            "rIdStyles".to_owned(),
            TargetMode::Internal,
        )
        .unwrap();
    (
        PackageWriter::to_bytes(&package).unwrap(),
        style_uri,
        style_xml,
    )
}

fn changed_commit(editor: &SourceBackedEditor) -> litchi_xlsx::cell_values::Commit {
    let mut edit = editor.edit("Sheet1").unwrap();
    edit.apply_batch([
        CellValueEdit::set(address("A1"), Number::new("10.50").unwrap()),
        CellValueEdit::set(address("B2"), false),
        CellValueEdit::set(address("C3"), Value::Error(ErrorValue::NotAvailable)),
    ])
    .unwrap();
    let commit = edit.commit().unwrap();
    assert!(commit.changed());
    assert_eq!(commit.diagnostics().changed_cells(), 3);
    assert_eq!(commit.diagnostics().touched_worksheets(), 1);
    commit
}

fn zip_member(bytes: &[u8], member: &str) -> Vec<u8> {
    ArchiveReader::new(bytes).unwrap().read(member).unwrap()
}

fn calculation_chain_target(package: &OpcPackage) -> Option<PackURI> {
    package
        .get_part(&PackURI::new(MAIN).unwrap())
        .ok()?
        .rels()
        .iter()
        .find_map(|relationship| {
            if matches!(
                relationship.reltype(),
                rt::CALC_CHAIN | rt::STRICT_CALC_CHAIN
            ) {
                relationship.target_partname().ok()
            } else {
                None
            }
        })
}

fn has_calculation_chain(package: &OpcPackage) -> bool {
    calculation_chain_target(package).is_some_and(|target| package.get_part(&target).is_ok())
}

fn scalar_formula_source() -> Vec<u8> {
    let mut package = fixture_package(
        format!(
            r#"<worksheet xmlns="{SML}"><dimension ref="A1:B1"/><sheetData><row r="1"><c r="A1"><f>A2+1</f><v>2</v></c><c r="B1"><v>99</v></c></row></sheetData></worksheet>"#
        ),
        false,
    );
    package
        .get_part_mut(&PackURI::new(MAIN).unwrap())
        .unwrap()
        .set_blob(
            format!(
                r#"<workbook xmlns="{SML}" xmlns:r="{REL}"><bookViews><workbookView/></bookViews><sheets><sheet name="Sheet1" sheetId="1" r:id="rIdSheet"/></sheets><calcPr calcId="123" calcMode="manual" fullCalcOnLoad="0" calcCompleted="1" forceFullCalc="0"/></workbook>"#
            )
            .into_bytes(),
        );
    package
        .try_add_part(Box::new(BlobPart::new(
            PackURI::new(CALC_CHAIN).unwrap(),
            CALC_CHAIN_CONTENT_TYPE.to_owned(),
            format!(r#"<calcChain xmlns="{SML}"><c r="A1" i="1"/></calcChain>"#).into_bytes(),
        )))
        .unwrap();
    package
        .get_part_mut(&PackURI::new(MAIN).unwrap())
        .unwrap()
        .rels_mut()
        .try_add_relationship(
            rt::CALC_CHAIN.to_owned(),
            "calcChain.xml".to_owned(),
            CALC_CHAIN_REL_ID.to_owned(),
            TargetMode::Internal,
        )
        .unwrap();
    PackageWriter::to_bytes(&package).unwrap()
}

fn date_source() -> Vec<u8> {
    fixture(
        format!(
            r#"<worksheet xmlns="{SML}"><sheetData><row r="1"><c r="A1" t="d"><v>2025-01-01</v></c></row></sheetData></worksheet>"#
        ),
        false,
    )
}

#[test]
fn first_middle_last_batch_reopens_preserves_unselected_parts_and_inverts() {
    let source_bytes = three_cells();
    let editor =
        SourceBackedEditor::from_read_at(Arc::new(VersionedSource::new(source_bytes.clone())))
            .unwrap();
    assert_eq!(editor.cache_diagnostics().successful_loads, 0);
    let commit = changed_commit(&editor);
    assert_eq!(editor.cache_diagnostics().successful_loads, 2);

    let mut replay = OpcPackage::from_bytes(&source_bytes).unwrap();
    commit.patch().apply(&mut replay).unwrap();
    assert_eq!(
        litchi_xlsx::cell_values::Snapshot::load(&replay, "Sheet1")
            .unwrap()
            .value(address("B2")),
        Some(&Value::Bool(false))
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
            .blob(),
    );

    let mut output = Vec::new();
    editor
        .publish_commit_to_stream(&mut output, &commit)
        .unwrap();
    let source = OpcPackage::from_bytes(&source_bytes).unwrap();
    let published = OpcPackage::from_bytes(&output).unwrap();
    litchi_xlsx::Package::from_bytes(output).unwrap();
    assert_eq!(
        source
            .get_part(&PackURI::new(UNUSED).unwrap())
            .unwrap()
            .blob(),
        published
            .get_part(&PackURI::new(UNUSED).unwrap())
            .unwrap()
            .blob(),
    );
    assert_eq!(
        litchi_xlsx::cell_values::Snapshot::load(&published, "Sheet1")
            .unwrap()
            .value(address("C3")),
        Some(&Value::Error(ErrorValue::NotAvailable))
    );
}

#[test]
fn scalar_staging_and_batch_staging_are_equivalent_and_clear_retains_record() {
    let bytes = three_cells();
    let scalar =
        SourceBackedEditor::from_read_at(Arc::new(VersionedSource::new(bytes.clone()))).unwrap();
    let mut edit = scalar.edit("Sheet1").unwrap();
    edit.set(address("A1"), 5u32).unwrap();
    edit.clear(address("C3")).unwrap();
    let scalar_commit = edit.commit().unwrap();

    let batch = SourceBackedEditor::from_read_at(Arc::new(VersionedSource::new(bytes))).unwrap();
    let mut edit = batch.edit("Sheet1").unwrap();
    edit.apply_batch([
        CellValueEdit::set(address("A1"), 5u32),
        CellValueEdit::clear(address("C3")),
    ])
    .unwrap();
    let batch_commit = edit.commit().unwrap();
    assert_eq!(
        scalar_commit.snapshot().source_xml(),
        batch_commit.snapshot().source_xml()
    );
    assert_eq!(batch_commit.snapshot().value(address("C3")), None);
    assert!(batch_commit.snapshot().contains_cell(address("C3")));
    assert!(
        String::from_utf8_lossy(batch_commit.snapshot().source_xml()).contains(r#"<c r="C3"></c>"#)
    );
}

#[test]
fn remove_deletes_exact_cell_owners_but_retains_rows_dimension_and_style_table() {
    let base = fixture(
        format!(
            r#"<worksheet xmlns="{SML}"><dimension ref="A1:C3"/><sheetData><row r="1" spans="1:3"><c r="A1" s="1"><v>1</v></c><c r="B1"><v>2</v></c></row><row r="2" spans="2:2"><c r="B2" t="b"><v>1</v></c></row><row r="3" spans="3:3"><c r="C3" t="inlineStr"><is><t>tail</t></is></c></row></sheetData></worksheet>"#
        ),
        false,
    );
    let (bytes, style_uri, style_xml) = with_two_cell_styles(&base);
    let original = OpcPackage::from_bytes(&bytes)
        .unwrap()
        .get_part(&PackURI::new(SHEET).unwrap())
        .unwrap()
        .blob()
        .to_vec();

    let clear_editor =
        SourceBackedEditor::from_read_at(Arc::new(VersionedSource::new(bytes.clone()))).unwrap();
    let mut clear = clear_editor.edit("Sheet1").unwrap();
    clear.clear(address("A1")).unwrap();
    let clear = clear.commit().unwrap();
    assert!(clear.snapshot().contains_cell(address("A1")));
    assert!(
        String::from_utf8_lossy(clear.snapshot().source_xml()).contains(r#"<c s="1" r="A1"></c>"#)
    );

    let editor =
        SourceBackedEditor::from_read_at(Arc::new(VersionedSource::new(bytes.clone()))).unwrap();
    let mut edit = editor.edit("Sheet1").unwrap();
    edit.apply_batch([
        CellValueEdit::remove(address("A1")),
        CellValueEdit::remove(address("B2")),
        CellValueEdit::remove(address("C3")),
    ])
    .unwrap();
    let commit = edit.commit().unwrap();
    assert_eq!(commit.diagnostics().changed_cells(), 3);
    for cell in ["A1", "B2", "C3"] {
        assert!(!commit.snapshot().contains_cell(address(cell)));
    }
    let xml = String::from_utf8_lossy(commit.snapshot().source_xml());
    assert!(xml.contains(r#"<dimension ref="A1:C3"/>"#));
    assert!(xml.contains(r#"<row r="1"><c r="B1"><v>2</v></c></row>"#));
    assert!(xml.contains(r#"<row r="2"></row>"#));
    assert!(xml.contains(r#"<row r="3"></row>"#));
    assert!(!xml.contains("spans="));
    assert!(!xml.contains(r#"r="A1""#));
    assert!(!xml.contains(r#"r="B2""#));
    assert!(!xml.contains(r#"r="C3""#));

    let mut replay = OpcPackage::from_bytes(&bytes).unwrap();
    commit.patch().apply(&mut replay).unwrap();
    assert_eq!(replay.get_part(&style_uri).unwrap().blob(), style_xml);
    commit.patch().inverse().apply(&mut replay).unwrap();
    assert_eq!(
        replay
            .get_part(&PackURI::new(SHEET).unwrap())
            .unwrap()
            .blob(),
        original,
    );

    let mut output = Vec::new();
    editor
        .publish_commit_to_stream(&mut output, &commit)
        .unwrap();
    assert_eq!(
        OpcPackage::from_bytes(&output)
            .unwrap()
            .get_part(&style_uri)
            .unwrap()
            .blob(),
        style_xml,
    );
    assert_eq!(
        OpcPackage::from_bytes(&output)
            .unwrap()
            .get_part(&PackURI::new(UNUSED).unwrap())
            .unwrap()
            .blob(),
        OpcPackage::from_bytes(&bytes)
            .unwrap()
            .get_part(&PackURI::new(UNUSED).unwrap())
            .unwrap()
            .blob(),
    );
}

#[test]
fn remove_missing_and_duplicate_selectors_fail_atomically() {
    let bytes = three_cells();
    let editor = SourceBackedEditor::from_read_at(Arc::new(VersionedSource::new(bytes))).unwrap();
    let mut edit = editor.edit("Sheet1").unwrap();
    assert!(edit.remove(address("Z99")).is_err());
    assert!(edit.is_empty());
    assert!(
        edit.apply_batch([
            CellValueEdit::remove(address("B2")),
            CellValueEdit::set(address("B2"), false),
        ])
        .is_err()
    );
    assert!(edit.is_empty());
    edit.remove(address("B2")).unwrap();
    assert!(edit.clear(address("B2")).is_err());
    assert_eq!(edit.len(), 1);
}

#[test]
fn exact_noop_duplicate_and_late_failure_are_atomic() {
    let bytes = three_cells();
    let editor =
        SourceBackedEditor::from_read_at(Arc::new(VersionedSource::new(bytes.clone()))).unwrap();
    let mut edit = editor.edit("Sheet1").unwrap();
    edit.set(address("A1"), Number::new("1").unwrap()).unwrap();
    let before = edit.len();
    assert!(
        edit.apply_batch([
            CellValueEdit::set(address("B2"), false),
            CellValueEdit::set(address("Z99"), 9u32),
        ])
        .is_err()
    );
    assert_eq!(edit.len(), before);
    assert!(edit.set(address("A1"), 2u32).is_err());
    let commit = edit.commit().unwrap();
    assert!(!commit.changed());
    let mut output = Vec::new();
    editor
        .publish_commit_to_stream(&mut output, &commit)
        .unwrap();
    assert_eq!(output, bytes);
}

#[test]
fn managed_scalar_exact_noop_publishes_without_detaching_source() {
    let bytes = three_cells();
    let exact = part_len(&bytes, MAIN) + part_len(&bytes, SHEET);
    let (budget, _cancellation_source, context) = managed_context(exact);
    let editor = SourceBackedEditor::from_read_at_with_execution_context(
        Arc::new(VersionedSource::new(bytes.clone())),
        litchi_xlsx::ReadLimits::default(),
        context,
    )
    .unwrap();
    let commit = editor.edit("Sheet1").unwrap().commit().unwrap();
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
fn exact_batch_limit_accepts_n_and_rejects_n_plus_one_atomically() {
    let mut rows = String::new();
    for row in 1..=MAX_BATCH_EDITS + 1 {
        rows.push_str(&format!(
            r#"<row r="{row}"><c r="A{row}"><v>{row}</v></c></row>"#
        ));
    }
    let bytes = fixture(
        format!(
            r#"<worksheet xmlns="{SML}"><dimension ref="A1:A257"/><sheetData>{rows}</sheetData></worksheet>"#
        ),
        false,
    );
    let editor =
        SourceBackedEditor::from_read_at(Arc::new(VersionedSource::new(bytes.clone()))).unwrap();
    let mut edit = editor.edit("Sheet1").unwrap();
    edit.apply_batch(
        (1..=MAX_BATCH_EDITS).map(|row| CellValueEdit::set(address(&format!("A{row}")), 0u32)),
    )
    .unwrap();
    assert_eq!(edit.len(), MAX_BATCH_EDITS);
    assert!(edit.set(address("A257"), 0u32).is_err());
    assert_eq!(edit.len(), MAX_BATCH_EDITS);

    let editor = SourceBackedEditor::from_read_at(Arc::new(VersionedSource::new(bytes))).unwrap();
    let mut edit = editor.edit("Sheet1").unwrap();
    edit.apply_batch(
        (1..=MAX_BATCH_EDITS).map(|row| CellValueEdit::remove(address(&format!("A{row}")))),
    )
    .unwrap();
    assert_eq!(edit.len(), MAX_BATCH_EDITS);
    assert!(edit.remove(address("A257")).is_err());
    assert_eq!(edit.len(), MAX_BATCH_EDITS);
}

#[test]
fn bounded_multi_sheet_transaction_publishes_one_overlay_and_inverts() {
    let source_bytes = two_sheets();
    let editor =
        SourceBackedEditor::from_read_at(Arc::new(VersionedSource::new(source_bytes.clone())))
            .unwrap();
    assert!(editor.snapshot("Sheet1").is_err());

    let commit = editor
        .edit_many([
            SheetCellValueEdit::set("Sheet1", address("A1"), 10u32),
            SheetCellValueEdit::clear("Sheet2", address("B2")),
            SheetCellValueEdit::remove("Sheet2", address("C3")),
        ])
        .unwrap()
        .commit()
        .unwrap();
    assert!(commit.changed());
    assert_eq!(commit.diagnostics().changed_cells(), 3);
    assert_eq!(commit.diagnostics().touched_worksheets(), 2);
    assert_eq!(commit.snapshot().len(), 2);
    assert_eq!(
        commit.snapshot().value(0, address("A1")),
        Some(&Value::Number(Number::new("10").unwrap()))
    );
    assert_eq!(commit.snapshot().value(1, address("B2")), None);
    assert!(!commit.snapshot().contains_cell(1, address("C3")));

    let mut replay = OpcPackage::from_bytes(&source_bytes).unwrap();
    commit.patch().apply(&mut replay).unwrap();
    assert_eq!(
        litchi_xlsx::cell_values::Snapshot::load_multi(&replay, "Sheet2")
            .unwrap()
            .value(address("B2")),
        None
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
            .blob(),
    );

    let mut output = Vec::new();
    editor
        .publish_multi_commit_to_stream(&mut output, &commit)
        .unwrap();
    let published = OpcPackage::from_bytes(&output).unwrap();
    assert_eq!(
        published
            .get_part(&PackURI::new(UNUSED).unwrap())
            .unwrap()
            .blob(),
        OpcPackage::from_bytes(&source_bytes)
            .unwrap()
            .get_part(&PackURI::new(UNUSED).unwrap())
            .unwrap()
            .blob(),
    );
    assert_eq!(
        litchi_xlsx::cell_values::Snapshot::load_multi(&published, "Sheet1")
            .unwrap()
            .value(address("A1")),
        Some(&Value::Number(Number::new("10").unwrap()))
    );
}

#[test]
fn multi_sheet_duplicate_and_late_failure_are_atomic() {
    let bytes = two_sheets();
    let editor = SourceBackedEditor::from_read_at(Arc::new(VersionedSource::new(bytes))).unwrap();
    let mut edit = editor
        .edit_sheets(["Sheet1".into(), "Sheet2".into()])
        .unwrap();
    assert!(
        edit.apply_batch([
            SheetCellValueEdit::set("Sheet1", address("A1"), 4u32),
            SheetCellValueEdit::set("Sheet1", address("A1"), 5u32),
        ])
        .is_err()
    );
    assert!(edit.is_empty());
    assert!(
        edit.apply_batch([
            SheetCellValueEdit::set("Sheet1", address("A1"), 4u32),
            SheetCellValueEdit::set("Sheet2", address("Z99"), 5u32),
        ])
        .is_err()
    );
    assert!(edit.is_empty());
}

#[test]
fn multi_sheet_commit_rejects_foreign_source_lineage() {
    let bytes = two_sheets();
    let first =
        SourceBackedEditor::from_read_at(Arc::new(VersionedSource::new(bytes.clone()))).unwrap();
    let commit = first
        .edit_many([SheetCellValueEdit::set("Sheet1", address("A1"), 4u32)])
        .unwrap()
        .commit()
        .unwrap();
    let second = SourceBackedEditor::from_read_at(Arc::new(VersionedSource::new(bytes))).unwrap();
    let mut output = Vec::new();
    assert!(matches!(
        second.publish_multi_commit_to_stream(&mut output, &commit),
        Err(Error::PatchConflict { .. })
    ));
    assert!(output.is_empty());
}

#[test]
fn multi_sheet_publication_rejects_source_revision_before_output() {
    let bytes = two_sheets();
    let source = Arc::new(VersionedSource::new(bytes));
    let editor = SourceBackedEditor::from_read_at(source.clone()).unwrap();
    let commit = editor
        .edit_many([SheetCellValueEdit::set("Sheet1", address("A1"), 4u32)])
        .unwrap()
        .commit()
        .unwrap();
    source.changed();
    let mut output = Vec::new();
    assert!(matches!(
        editor.publish_multi_commit_to_stream(&mut output, &commit),
        Err(Error::Package(OpcError::SourceChanged { .. }))
    ));
    assert!(output.is_empty());
}

#[test]
fn multi_sheet_exact_noop_publishes_source_byte_for_byte() {
    let bytes = two_sheets();
    let editor =
        SourceBackedEditor::from_read_at(Arc::new(VersionedSource::new(bytes.clone()))).unwrap();
    let commit = editor
        .edit_sheets(["Sheet1".into(), "Sheet2".into()])
        .unwrap()
        .commit()
        .unwrap();
    assert!(!commit.changed());
    assert!(commit.patch().is_empty());
    let mut output = Vec::new();
    editor
        .publish_multi_commit_to_stream(&mut output, &commit)
        .unwrap();
    assert_eq!(output, bytes);
}

#[test]
fn managed_multi_sheet_exact_noop_publishes_without_detaching_sources() {
    let bytes = two_sheets();
    let exact = part_len(&bytes, MAIN)
        + part_len(&bytes, SHEET)
        + part_len(&bytes, "/xl/worksheets/sheet2.xml");
    let (budget, _cancellation_source, context) = managed_context(exact);
    let editor = SourceBackedEditor::from_read_at_with_execution_context(
        Arc::new(VersionedSource::new(bytes.clone())),
        litchi_xlsx::ReadLimits::default(),
        context,
    )
    .unwrap();
    let commit = editor
        .edit_sheets(["Sheet1".into(), "Sheet2".into()])
        .unwrap()
        .commit()
        .unwrap();
    assert!(commit.patch().is_empty());

    let mut output = Vec::new();
    editor
        .publish_multi_commit_to_stream(&mut output, &commit)
        .unwrap();
    assert_eq!(output, bytes);
    drop(commit);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn multi_sheet_patch_rejects_unselected_graph_tampering() {
    let bytes = two_sheets();
    let editor =
        SourceBackedEditor::from_read_at(Arc::new(VersionedSource::new(bytes.clone()))).unwrap();
    let commit = editor
        .edit_many([SheetCellValueEdit::set("Sheet1", address("A1"), 4u32)])
        .unwrap()
        .commit()
        .unwrap();

    let sheet2 = PackURI::new("/xl/worksheets/sheet2.xml").unwrap();
    let mut wrong_type = OpcPackage::from_bytes(&bytes).unwrap();
    wrong_type
        .get_part_mut(&sheet2)
        .unwrap()
        .set_content_type("application/octet-stream".to_owned())
        .unwrap();
    assert!(matches!(
        commit.patch().apply(&mut wrong_type),
        Err(Error::PatchConflict { .. })
    ));

    let mut wrong_relationship = OpcPackage::from_bytes(&bytes).unwrap();
    wrong_relationship
        .get_part_mut(&sheet2)
        .unwrap()
        .rels_mut()
        .try_add_relationship(
            rt::IMAGE.to_owned(),
            "../media/unused.bin".to_owned(),
            "rIdInjected".to_owned(),
            TargetMode::Internal,
        )
        .unwrap();
    assert!(matches!(
        commit.patch().apply(&mut wrong_relationship),
        Err(Error::PatchConflict { .. })
    ));

    let mut wrong_owner = OpcPackage::from_bytes(&bytes).unwrap();
    wrong_owner
        .rels_mut()
        .try_add_relationship(
            rt::OFFICE_DOCUMENT.to_owned(),
            "xl/worksheets/sheet2.xml".to_owned(),
            "rIdSecondOwner".to_owned(),
            TargetMode::Internal,
        )
        .unwrap();
    assert!(matches!(
        commit.patch().apply(&mut wrong_owner),
        Err(Error::PatchConflict { .. })
    ));
}

#[test]
fn multi_sheet_patch_composes_with_unselected_payload_edits_and_inverse() {
    let bytes = two_sheets();
    let editor =
        SourceBackedEditor::from_read_at(Arc::new(VersionedSource::new(bytes.clone()))).unwrap();
    let commit = editor
        .edit_many([SheetCellValueEdit::set("Sheet1", address("A1"), 42u32)])
        .unwrap()
        .commit()
        .unwrap();

    let sheet2 = PackURI::new("/xl/worksheets/sheet2.xml").unwrap();
    let independent_sheet2 = format!(
        r#"<worksheet xmlns="{SML}"><dimension ref="A1:C3"/><sheetData><row r="1"><c r="A1"><v>20</v></c></row><row r="2"><c r="B2" t="b"><v>0</v></c></row><row r="3"><c r="C3" t="inlineStr"><is><t>independent</t></is></c></row></sheetData></worksheet>"#
    )
    .into_bytes();
    let mut composed = OpcPackage::from_bytes(&bytes).unwrap();
    composed
        .get_part_mut(&sheet2)
        .unwrap()
        .set_blob(independent_sheet2.clone());
    commit.patch().apply(&mut composed).unwrap();
    assert_eq!(
        composed.get_part(&sheet2).unwrap().blob(),
        independent_sheet2.as_slice()
    );
    assert_eq!(
        litchi_xlsx::cell_values::Snapshot::load_multi(&composed, "Sheet1")
            .unwrap()
            .value(address("A1")),
        Some(&Value::Number(Number::new("42").unwrap()))
    );

    commit.patch().inverse().apply(&mut composed).unwrap();
    assert_eq!(
        composed.get_part(&sheet2).unwrap().blob(),
        independent_sheet2.as_slice()
    );
    assert_eq!(
        litchi_xlsx::cell_values::Snapshot::load_multi(&composed, "Sheet1")
            .unwrap()
            .value(address("A1")),
        Some(&Value::Number(Number::new("1").unwrap()))
    );
}

#[test]
fn multi_sheet_selector_bounds_and_semantic_duplicates_are_atomic() {
    let bytes = two_sheets();
    let editor = SourceBackedEditor::from_read_at(Arc::new(VersionedSource::new(bytes))).unwrap();
    let mut selectors = Vec::new();
    selectors.extend(std::iter::repeat_n(0usize.into(), 65));
    assert!(editor.edit_sheets(selectors).is_err());
    assert!(
        editor
            .edit_sheets(vec!["Sheet1".into(), 0usize.into()])
            .is_err()
    );
}

#[test]
fn signed_multi_sheet_noop_is_exact_and_changed_publication_refuses() {
    let bytes = signed_two_sheets();
    let editor =
        SourceBackedEditor::from_read_at(Arc::new(VersionedSource::new(bytes.clone()))).unwrap();
    let noop = editor
        .edit_sheets(["Sheet1".into(), "Sheet2".into()])
        .unwrap()
        .commit()
        .unwrap();
    let mut output = Vec::new();
    editor
        .publish_multi_commit_to_stream(&mut output, &noop)
        .unwrap();
    assert_eq!(output, bytes);

    let editor = SourceBackedEditor::from_read_at(Arc::new(VersionedSource::new(bytes))).unwrap();
    let changed = editor
        .edit_many([SheetCellValueEdit::set("Sheet1", address("A1"), 42u32)])
        .unwrap()
        .commit()
        .unwrap();
    let mut output = Vec::new();
    assert!(matches!(
        editor.publish_multi_commit_to_stream(&mut output, &changed),
        Err(Error::Package(OpcError::SignedSourceRequiresExplicitPolicy))
    ));
    assert!(output.is_empty());
}

#[test]
fn mce_shared_strings_relationships_and_signed_changes_are_refused() {
    for xml in [
        format!(
            r#"<worksheet xmlns="{SML}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><sheetData/><mc:AlternateContent/></worksheet>"#
        ),
        format!(r#"<worksheet xmlns="{SML}" future="unknown"><sheetData/></worksheet>"#),
        format!(r#"<worksheet xmlns="{SML}"><sheetData>payload</sheetData></worksheet>"#),
        format!(
            r#"<worksheet xmlns="{SML}" xmlns:s="http://purl.oclc.org/ooxml/spreadsheetml/main"><sheetData><s:row r="1"><s:c r="A1"><s:v>1</s:v></s:c></s:row></sheetData></worksheet>"#
        ),
    ] {
        let editor =
            SourceBackedEditor::from_read_at(Arc::new(VersionedSource::new(fixture(xml, false))))
                .unwrap();
        assert!(editor.edit("Sheet1").is_err());
    }

    let mut shared = OpcPackage::from_bytes(&three_cells()).unwrap();
    shared
        .try_add_part(Box::new(BlobPart::new(
            PackURI::new("/xl/sharedStrings.xml").unwrap(),
            ct::SML_SHARED_STRINGS.to_owned(),
            format!(r#"<sst xmlns="{SML}"><si><t>unused</t></si></sst>"#).into_bytes(),
        )))
        .unwrap();
    shared
        .get_part_mut(&PackURI::new(MAIN).unwrap())
        .unwrap()
        .rels_mut()
        .try_add_relationship(
            rt::SHARED_STRINGS.to_owned(),
            "sharedStrings.xml".to_owned(),
            "rIdShared".to_owned(),
            TargetMode::Internal,
        )
        .unwrap();
    let editor = SourceBackedEditor::from_read_at(Arc::new(VersionedSource::new(
        PackageWriter::to_bytes(&shared).unwrap(),
    )))
    .unwrap();
    assert!(editor.edit("Sheet1").is_err());

    let signed = fixture(
        format!(
            r#"<worksheet xmlns="{SML}"><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>"#
        ),
        true,
    );
    let editor = SourceBackedEditor::from_read_at(Arc::new(VersionedSource::new(signed))).unwrap();
    let mut edit = editor.edit("Sheet1").unwrap();
    edit.remove(address("A1")).unwrap();
    let commit = edit.commit().unwrap();
    let mut output = Vec::new();
    assert!(matches!(
        editor.publish_commit_to_stream(&mut output, &commit),
        Err(Error::Package(OpcError::SignedSourceRequiresExplicitPolicy))
    ));
    assert!(output.is_empty());
}

#[test]
fn changed_positional_source_and_stale_in_memory_patch_are_refused() {
    let bytes = three_cells();
    let source = Arc::new(VersionedSource::new(bytes.clone()));
    let editor = SourceBackedEditor::from_read_at(source.clone()).unwrap();
    let commit = changed_commit(&editor);
    source.changed();
    let mut output = Vec::new();
    assert!(matches!(
        editor.publish_commit_to_stream(&mut output, &commit),
        Err(Error::Package(OpcError::SourceChanged { .. }))
    ));
    assert!(output.is_empty());

    let mut stale = OpcPackage::from_bytes(&bytes).unwrap();
    stale
        .get_part_mut(&PackURI::new(SHEET).unwrap())
        .unwrap()
        .set_blob(b"<broken/>".to_vec());
    assert!(matches!(
        commit.patch().apply(&mut stale),
        Err(Error::PatchConflict { .. })
    ));
}

#[test]
fn source_version_flip_after_selected_read_rejects_single_and_multi_snapshots() {
    let bytes = three_cells();
    let worksheet_data_offset = zip_member_data_offset(&bytes, b"xl/worksheets/sheet1.xml");
    let source = Arc::new(VersionedSource::new(bytes.clone()));
    let editor = SourceBackedEditor::from_read_at(source.clone()).unwrap();
    source.flip_after_read_at(worksheet_data_offset);
    assert!(matches!(
        editor.snapshot("Sheet1"),
        Err(Error::Package(OpcError::SourceChanged { .. }))
    ));

    let source = Arc::new(VersionedSource::new(bytes));
    let editor = SourceBackedEditor::from_read_at(source.clone()).unwrap();
    source.flip_after_read_at(worksheet_data_offset);
    assert!(matches!(
        editor.edit_sheets(["Sheet1".into()]),
        Err(Error::Package(OpcError::SourceChanged { .. }))
    ));
}

#[test]
fn exact_noop_publication_does_not_reparse_an_oversized_selected_worksheet() {
    let bytes = oversized_multi_sheet_source();
    let worksheet_data_offset = zip_member_data_offset(&bytes, b"xl/worksheets/sheet1.xml");
    let source = Arc::new(VersionedSource::new(bytes.clone()));
    let editor = SourceBackedEditor::from_read_at(source.clone()).unwrap();
    let commit = editor
        .edit_sheets(["Sheet1".into(), "Sheet2".into()])
        .unwrap()
        .commit()
        .unwrap();
    assert!(commit.patch().is_empty());

    source.reject_read_at(worksheet_data_offset);
    let uncached_reload = SourceBackedEditor::from_read_at(source.clone()).unwrap();
    assert!(uncached_reload.edit_sheets(["Sheet1".into()]).is_err());
    let rejected_reads = source.rejected_read_count();
    assert!(rejected_reads > 0);

    let mut output = Vec::new();
    editor
        .publish_multi_commit_to_stream(&mut output, &commit)
        .unwrap();
    assert_eq!(source.rejected_read_count(), rejected_reads);
    assert_eq!(output, bytes);
}

#[test]
fn changed_publication_reuses_selected_semantics_without_extra_source_reads() {
    let bytes = oversized_multi_sheet_source();
    let selected_offset = zip_member_data_offset(&bytes, b"xl/worksheets/sheet1.xml");
    let unselected_offset = zip_member_data_offset(&bytes, b"xl/worksheets/sheet2.xml");
    let source = Arc::new(VersionedSource::new(bytes));
    source.set_worksheet_diagnostic_offsets(selected_offset, unselected_offset);
    let editor = SourceBackedEditor::from_read_at(source.clone()).unwrap();
    assert_eq!(editor.cache_diagnostics().successful_loads, 0);

    let edit = editor
        .edit_many([SheetCellValueEdit::set("Sheet1", address("A1"), 2u32)])
        .unwrap();
    let commit = edit.commit().unwrap();
    assert!(commit.changed());
    let loaded = editor.cache_diagnostics();
    assert_eq!(loaded.successful_loads, 2);
    assert_eq!(loaded.cold_loads, 2);

    source.reset_read_diagnostics();
    let mut output = Vec::new();
    editor
        .publish_multi_commit_to_stream(&mut output, &commit)
        .unwrap();
    let (raw_read_calls, raw_read_bytes, selected_data_reads, _unselected_data_reads) =
        source.read_diagnostics();
    assert!(raw_read_calls > 0);
    assert!(raw_read_bytes > 0);
    assert_eq!(selected_data_reads, 1);
    assert!(!output.is_empty());
}

#[test]
fn commit_is_bound_to_the_exact_positional_source_identity() {
    let bytes = three_cells();
    let source =
        SourceBackedEditor::from_read_at(Arc::new(VersionedSource::new(bytes.clone()))).unwrap();
    let commit = changed_commit(&source);
    let foreign = SourceBackedEditor::from_read_at(Arc::new(VersionedSource::new(bytes))).unwrap();
    let mut output = Vec::new();
    assert!(matches!(
        foreign.publish_commit_to_stream(&mut output, &commit),
        Err(Error::PatchConflict { .. })
    ));
    assert!(output.is_empty());
}

#[test]
fn shared_style_owner_is_preserved_and_bound_into_patch_closure() {
    let bytes = fixture(
        format!(
            r#"<worksheet xmlns="{SML}"><sheetData><row r="1"><c r="A1" s="1"><v>1</v></c></row></sheetData></worksheet>"#
        ),
        false,
    );
    let mut package = OpcPackage::from_bytes(&bytes).unwrap();
    let style_uri = PackURI::new("/xl/styles.xml").unwrap();
    let style_xml = format!(
        r#"<styleSheet xmlns="{SML}"><cellXfs count="2"><xf/><xf/></cellXfs></styleSheet>"#
    )
    .into_bytes();
    package
        .try_add_part(Box::new(BlobPart::new(
            style_uri.clone(),
            ct::SML_STYLES.to_owned(),
            style_xml.clone(),
        )))
        .unwrap();
    package
        .get_part_mut(&PackURI::new(MAIN).unwrap())
        .unwrap()
        .rels_mut()
        .try_add_relationship(
            rt::STYLES.to_owned(),
            "styles.xml".to_owned(),
            "rIdStyles".to_owned(),
            TargetMode::Internal,
        )
        .unwrap();
    let bytes = PackageWriter::to_bytes(&package).unwrap();
    let editor =
        SourceBackedEditor::from_read_at(Arc::new(VersionedSource::new(bytes.clone()))).unwrap();
    let mut edit = editor.edit("Sheet1").unwrap();
    edit.set(address("A1"), 2u32).unwrap();
    let commit = edit.commit().unwrap();
    let mut output = Vec::new();
    editor
        .publish_commit_to_stream(&mut output, &commit)
        .unwrap();
    assert_eq!(
        OpcPackage::from_bytes(&output)
            .unwrap()
            .get_part(&style_uri)
            .unwrap()
            .blob(),
        style_xml
    );

    let mut hostile = OpcPackage::from_bytes(&bytes).unwrap();
    hostile
        .get_part_mut(&style_uri)
        .unwrap()
        .set_blob(
            format!(
                r#"<styleSheet xmlns="{SML}"><cellXfs count="2"><xf/><xf numFmtId="1"/></cellXfs></styleSheet>"#
            )
            .into_bytes(),
        );
    assert!(matches!(
        commit.patch().apply(&mut hostile),
        Err(Error::PatchConflict { .. })
    ));
}

#[test]
fn noop_patch_rejects_added_worksheet_relationship() {
    let bytes = three_cells();
    let editor =
        SourceBackedEditor::from_read_at(Arc::new(VersionedSource::new(bytes.clone()))).unwrap();
    let noop = editor.edit("Sheet1").unwrap().commit().unwrap();
    let mut hostile = OpcPackage::from_bytes(&bytes).unwrap();
    hostile
        .get_part_mut(&PackURI::new(SHEET).unwrap())
        .unwrap()
        .rels_mut()
        .try_add_relationship(
            rt::IMAGE.to_owned(),
            "../media/unused.bin".to_owned(),
            "rIdInjected".to_owned(),
            TargetMode::Internal,
        )
        .unwrap();
    assert!(matches!(
        noop.patch().apply(&mut hostile),
        Err(Error::PatchConflict { .. })
    ));
}

#[test]
fn managed_editor_retains_complete_payload_closure_and_releases_budget() {
    let base = fixture(
        format!(
            r#"<worksheet xmlns="{SML}"><sheetData><row r="1"><c r="A1" s="1"><v>1</v></c></row></sheetData></worksheet>"#
        ),
        false,
    );
    let (with_styles, _style_uri, _style_xml) = with_two_cell_styles(&base);
    let bytes = with_style_and_theme(&with_styles);
    let exact = part_len(&bytes, MAIN)
        + part_len(&bytes, SHEET)
        + part_len(&bytes, "/xl/styles.xml")
        + part_len(&bytes, "/xl/theme/theme1.xml");
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
    let commit = editor.edit("Sheet1").unwrap().commit().unwrap();
    assert_eq!(budget.used(Resource::Memory), exact);
    let diagnostics = commit.snapshot();
    assert_eq!(
        diagnostics.value(address("A1")),
        Some(&Value::Number(Number::new("1").unwrap()))
    );
    assert_eq!(
        commit.snapshot().source_xml(),
        commit.patch().before().source_xml()
    );
    assert_eq!(
        commit.patch().before().source_xml().len() as u64,
        part_len(&with_styles, SHEET)
    );
    drop(commit);
    drop(editor);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn managed_multi_sheet_batch_retains_each_selected_payload_and_streams_once() {
    let bytes = two_sheets();
    let exact = part_len(&bytes, MAIN)
        + part_len(&bytes, SHEET)
        + part_len(&bytes, "/xl/worksheets/sheet2.xml");
    let publication_limit = exact + u64::try_from(bytes.len()).unwrap();
    let (budget, _cancellation_source, context) = managed_context(publication_limit);
    let editor =
        SourceBackedEditor::from_read_at_with_limits_and_cache_limits_and_execution_context(
            Arc::new(VersionedSource::new(bytes.clone())),
            litchi_xlsx::ReadLimits::default(),
            SourceCacheLimits::new(usize::try_from(exact).unwrap(), 8).unwrap(),
            context,
        )
        .unwrap();
    let commit = editor
        .edit_many([
            SheetCellValueEdit::set("Sheet1", address("A1"), 40u32),
            SheetCellValueEdit::set("Sheet2", address("A1"), 50u32),
        ])
        .unwrap()
        .commit()
        .unwrap();
    assert_eq!(budget.used(Resource::Memory), exact);
    let mut output = Vec::new();
    editor
        .publish_multi_commit_to_stream(&mut output, &commit)
        .unwrap();
    let published = OpcPackage::from_bytes(&output).unwrap();
    assert_eq!(
        litchi_xlsx::cell_values::Snapshot::load_multi(&published, "Sheet1")
            .unwrap()
            .value(address("A1")),
        Some(&Value::Number(Number::new("40").unwrap()))
    );
    assert_eq!(
        litchi_xlsx::cell_values::Snapshot::load_multi(&published, "Sheet2")
            .unwrap()
            .value(address("A1")),
        Some(&Value::Number(Number::new("50").unwrap()))
    );
    drop(commit);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn managed_one_under_budget_refuses_before_selected_worksheet_io() {
    let bytes = three_cells();
    let workbook_bytes = part_len(&bytes, MAIN);
    let worksheet_bytes = part_len(&bytes, SHEET);
    let exact = workbook_bytes + worksheet_bytes;
    let worksheet_data_offset = zip_member_data_offset(&bytes, b"xl/worksheets/sheet1.xml");
    let source = Arc::new(VersionedSource::new(bytes));
    source.reject_read_at(worksheet_data_offset);
    let (budget, _cancellation_source, context) = managed_context(exact - 1);
    let editor = SourceBackedEditor::from_read_at_with_execution_context(
        source.clone(),
        litchi_xlsx::ReadLimits::default(),
        context,
    )
    .unwrap();
    assert_eq!(budget.used(Resource::Memory), 0);
    let result = editor.edit("Sheet1");
    assert!(matches!(
        result,
        Err(Error::Package(OpcError::Execution(
            ExecutionError::ResourceLimit(_)
        )))
    ));
    assert_eq!(source.rejected_read_count(), 0);
    assert_eq!(editor.cache_diagnostics().successful_loads, 1);
    drop(editor);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn managed_stream_publication_works_and_inverse_materialization_is_explicit() {
    let bytes = three_cells();
    let workbook_bytes = part_len(&bytes, MAIN);
    let worksheet_bytes = part_len(&bytes, SHEET);
    let exact = workbook_bytes + worksheet_bytes;
    let publication_limit = exact + u64::try_from(bytes.len()).unwrap();
    let (budget, _cancellation_source, context) = managed_context(publication_limit);
    let editor = SourceBackedEditor::from_read_at_with_execution_context(
        Arc::new(VersionedSource::new(bytes.clone())),
        litchi_xlsx::ReadLimits::default(),
        context,
    )
    .unwrap();
    let commit = changed_commit(&editor);
    assert_eq!(budget.used(Resource::Memory), exact);

    let mut replay = OpcPackage::from_bytes(&bytes).unwrap();
    commit.patch().apply(&mut replay).unwrap();
    assert!(matches!(
        commit.patch().inverse().apply(&mut replay),
        Err(Error::Package(OpcError::ManagedPartDataArcEscape))
    ));
    commit
        .patch()
        .inverse()
        .apply_materialized(&mut replay, exact as usize)
        .unwrap();
    assert_eq!(
        replay
            .get_part(&PackURI::new(SHEET).unwrap())
            .unwrap()
            .blob(),
        OpcPackage::from_bytes(&bytes)
            .unwrap()
            .get_part(&PackURI::new(SHEET).unwrap())
            .unwrap()
            .blob(),
    );

    let mut output = Vec::new();
    editor
        .publish_commit_to_stream(&mut output, &commit)
        .unwrap();
    assert_eq!(
        litchi_xlsx::cell_values::Snapshot::load(
            &OpcPackage::from_bytes(&output).unwrap(),
            "Sheet1",
        )
        .unwrap()
        .value(address("A1")),
        Some(&Value::Number(Number::new("10.50").unwrap())),
    );
    drop(commit);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn managed_cancellation_is_checked_after_snapshot_and_before_stream_output() {
    let bytes = three_cells();
    let exact = part_len(&bytes, MAIN) + part_len(&bytes, SHEET);
    let source = Arc::new(VersionedSource::new(bytes.clone()));
    let (budget, cancellation_source, context) = managed_context(exact);
    let editor = SourceBackedEditor::from_read_at_with_execution_context(
        source,
        litchi_xlsx::ReadLimits::default(),
        context,
    )
    .unwrap();
    let commit = changed_commit(&editor);
    cancellation_source.cancel();
    let mut output = Vec::new();
    assert!(matches!(
        editor.publish_commit_to_stream(&mut output, &commit),
        Err(Error::Package(OpcError::Cancelled))
    ));
    assert!(output.is_empty());
    drop(commit);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn managed_cancellation_during_batch_iterators_keeps_staging_atomic() {
    let bytes = three_cells();
    let exact = part_len(&bytes, MAIN) + part_len(&bytes, SHEET);
    let (budget, cancellation_source, context) = managed_context(exact);
    let editor = SourceBackedEditor::from_read_at_with_execution_context(
        Arc::new(VersionedSource::new(bytes.clone())),
        litchi_xlsx::ReadLimits::default(),
        context,
    )
    .unwrap();
    let mut edit = editor.edit("Sheet1").unwrap();
    let cancellation_for_iterator = cancellation_source.clone();
    let mut index = 0;
    let result = edit.apply_batch(std::iter::from_fn(move || {
        let value = match index {
            0 => Some(CellValueEdit::set(address("A1"), 2u32)),
            1 => {
                cancellation_for_iterator.cancel();
                Some(CellValueEdit::set(address("B2"), false))
            },
            _ => None,
        };
        index += 1;
        value
    }));
    assert!(matches!(result, Err(Error::Package(OpcError::Cancelled))));
    assert!(edit.is_empty());
    drop(edit);
    drop(editor);
    assert_eq!(budget.used(Resource::Memory), 0);

    let bytes = two_sheets();
    let exact = part_len(&bytes, MAIN)
        + part_len(&bytes, SHEET)
        + part_len(&bytes, "/xl/worksheets/sheet2.xml");
    let (budget, cancellation_source, context) = managed_context(exact);
    let editor = SourceBackedEditor::from_read_at_with_execution_context(
        Arc::new(VersionedSource::new(bytes)),
        litchi_xlsx::ReadLimits::default(),
        context,
    )
    .unwrap();
    let mut edit = editor
        .edit_sheets(["Sheet1".into(), "Sheet2".into()])
        .unwrap();
    let cancellation_for_iterator = cancellation_source.clone();
    let mut index = 0;
    let result = edit.apply_batch(std::iter::from_fn(move || {
        let value = match index {
            0 => Some(SheetCellValueEdit::set("Sheet1", address("A1"), 2u32)),
            1 => {
                cancellation_for_iterator.cancel();
                Some(SheetCellValueEdit::set("Sheet2", address("A1"), 3u32))
            },
            _ => None,
        };
        index += 1;
        value
    }));
    assert!(matches!(result, Err(Error::Package(OpcError::Cancelled))));
    assert!(edit.is_empty());
    drop(edit);
    drop(editor);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn managed_patch_cancellation_during_finalization_does_not_publish_candidate() {
    let bytes = three_cells();
    let exact = part_len(&bytes, MAIN) + part_len(&bytes, SHEET);
    let (budget, cancellation_source, context) = managed_context(exact);
    let editor = SourceBackedEditor::from_read_at_with_execution_context(
        Arc::new(VersionedSource::new(bytes.clone())),
        litchi_xlsx::ReadLimits::default(),
        context,
    )
    .unwrap();
    let commit = changed_commit(&editor);
    let mut replay = OpcPackage::from_bytes(&bytes).unwrap();
    replace_sheet_with_cancel_on_blob_arc(&mut replay, cancellation_source, 1);
    assert!(matches!(
        commit.patch().apply(&mut replay),
        Err(Error::Package(OpcError::Cancelled))
    ));
    assert_eq!(
        replay
            .get_part(&PackURI::new(SHEET).unwrap())
            .unwrap()
            .blob(),
        OpcPackage::from_bytes(&bytes)
            .unwrap()
            .get_part(&PackURI::new(SHEET).unwrap())
            .unwrap()
            .blob(),
    );
    drop(commit);
    drop(editor);
    assert_eq!(budget.used(Resource::Memory), 0);

    let bytes = two_sheets();
    let exact = part_len(&bytes, MAIN)
        + part_len(&bytes, SHEET)
        + part_len(&bytes, "/xl/worksheets/sheet2.xml");
    let (budget, cancellation_source, context) = managed_context(exact);
    let editor = SourceBackedEditor::from_read_at_with_execution_context(
        Arc::new(VersionedSource::new(bytes.clone())),
        litchi_xlsx::ReadLimits::default(),
        context,
    )
    .unwrap();
    let commit = editor
        .edit_many([SheetCellValueEdit::set("Sheet1", address("A1"), 10u32)])
        .unwrap()
        .commit()
        .unwrap();
    let mut replay = OpcPackage::from_bytes(&bytes).unwrap();
    // MultiPatch's initial owning readback consumes the first blob_arc call;
    // the second call is the candidate readback after staged replacement.
    replace_sheet_with_cancel_on_blob_arc(&mut replay, cancellation_source, 2);
    assert!(matches!(
        commit.patch().apply(&mut replay),
        Err(Error::Package(OpcError::Cancelled))
    ));
    assert_eq!(
        replay
            .get_part(&PackURI::new(SHEET).unwrap())
            .unwrap()
            .blob(),
        OpcPackage::from_bytes(&bytes)
            .unwrap()
            .get_part(&PackURI::new(SHEET).unwrap())
            .unwrap()
            .blob(),
    );
    drop(commit);
    drop(editor);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn managed_empty_multi_edit_iterator_reports_cancellation_before_invalid_batch() {
    let bytes = three_cells();
    let exact = part_len(&bytes, MAIN) + part_len(&bytes, SHEET);
    let (budget, cancellation_source, context) = managed_context(exact);
    let editor = SourceBackedEditor::from_read_at_with_execution_context(
        Arc::new(VersionedSource::new(bytes)),
        litchi_xlsx::ReadLimits::default(),
        context,
    )
    .unwrap();
    let result = editor.edit_many(std::iter::from_fn(move || {
        cancellation_source.cancel();
        None::<SheetCellValueEdit<'static>>
    }));
    assert!(matches!(result, Err(Error::Package(OpcError::Cancelled))));
    drop(editor);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn source_backed_date_edit_retains_date_semantics() {
    let bytes = date_source();
    let editor = SourceBackedEditor::from_read_at(Arc::new(VersionedSource::new(bytes))).unwrap();
    let mut edit = editor.edit("Sheet1").unwrap();
    edit.set(address("A1"), Value::date("2026-08-14").unwrap())
        .unwrap();
    let commit = edit.commit().unwrap();
    assert!(commit.changed());
    assert!(matches!(
        commit.snapshot().cell(address("A1")),
        Some(Cell::Value(Value::Date(date))) if date.as_str() == "2026-08-14"
    ));

    let mut output = Vec::new();
    editor
        .publish_commit_to_stream(&mut output, &commit)
        .unwrap();
    let xml = String::from_utf8(zip_member(&output, "xl/worksheets/sheet1.xml")).unwrap();
    assert!(xml.contains(r#"r="A1" t="d""#));
    assert!(xml.contains("2026-08-14"));
}

#[test]
fn scalar_formula_replacement_drops_cache_invalidates_calculation_and_preserves_members() {
    let bytes = scalar_formula_source();
    let source_package = OpcPackage::from_bytes(&bytes).unwrap();
    assert!(has_calculation_chain(&source_package));
    let unused = source_package
        .get_part(&PackURI::new(UNUSED).unwrap())
        .unwrap()
        .blob()
        .to_vec();

    let editor = SourceBackedEditor::from_read_at(Arc::new(VersionedSource::new(bytes))).unwrap();
    let mut edit = editor.edit("Sheet1").unwrap();
    edit.set_formula(address("A1"), Formula::new("A2+2").unwrap())
        .unwrap();
    let commit = edit.commit().unwrap();
    assert!(commit.changed());
    assert!(matches!(
        commit.snapshot().cell(address("A1")),
        Some(Cell::Formula(formula))
            if formula.text() == "A2+2" && formula.cached().is_none()
    ));

    let mut output = Vec::new();
    editor
        .publish_commit_to_stream(&mut output, &commit)
        .unwrap();
    let published = OpcPackage::from_bytes(&output).unwrap();
    assert!(!has_calculation_chain(&published));
    assert!(
        published
            .get_part(&PackURI::new(CALC_CHAIN).unwrap())
            .is_err()
    );
    assert_eq!(
        published
            .get_part(&PackURI::new(UNUSED).unwrap())
            .unwrap()
            .blob(),
        unused.as_slice(),
    );

    let workbook_xml = String::from_utf8(zip_member(&output, "xl/workbook.xml")).unwrap();
    assert!(workbook_xml.contains(r#"calcId="0""#));
    assert!(workbook_xml.contains(r#"fullCalcOnLoad="true""#));
    assert!(workbook_xml.contains(r#"calcCompleted="false""#));
    assert!(workbook_xml.contains(r#"forceFullCalc="true""#));
    let sheet_xml = String::from_utf8(zip_member(&output, "xl/worksheets/sheet1.xml")).unwrap();
    assert!(sheet_xml.contains("<f>A2+2</f>"));
    assert!(!sheet_xml.contains("<v>2</v>"));
    let content_types = String::from_utf8(zip_member(&output, "[Content_Types].xml")).unwrap();
    assert!(!content_types.contains(CALC_CHAIN));
    assert!(!content_types.contains(CALC_CHAIN_CONTENT_TYPE));
}

#[test]
fn grouped_formula_edits_are_refused_without_staging() {
    let cases = [
        (
            "array",
            address("A1"),
            format!(
                r#"<worksheet xmlns="{SML}"><sheetData><row r="1"><c r="A1"><f t="array" ref="A1:B1">SUM(A1:A2)</f><v>2</v></c><c r="B1"><v>2</v></c></row></sheetData></worksheet>"#
            ),
        ),
        (
            "data table",
            address("A1"),
            format!(
                r#"<worksheet xmlns="{SML}"><sheetData><row r="1"><c r="A1"><f t="dataTable" ref="A1:B1">A1</f><v>2</v></c><c r="B1"><v>2</v></c></row></sheetData></worksheet>"#
            ),
        ),
        (
            "shared",
            address("B1"),
            format!(
                r#"<worksheet xmlns="{SML}"><sheetData><row r="1"><c r="A1"><f t="shared" si="0" ref="A1:B1">A1+1</f><v>2</v></c><c r="B1"><f t="shared" si="0"/><v>3</v></c></row></sheetData></worksheet>"#
            ),
        ),
    ];

    for (kind, target, xml) in cases {
        let editor =
            SourceBackedEditor::from_read_at(Arc::new(VersionedSource::new(fixture(xml, false))))
                .unwrap();
        let mut edit = editor.edit("Sheet1").unwrap();
        assert!(
            edit.set_formula(target, Formula::new("A1+9").unwrap())
                .is_err(),
            "{kind} formula must remain group-scoped"
        );
        assert!(edit.is_empty(), "{kind} refusal must be atomic");
    }
}

#[test]
fn formula_noop_is_byte_exact_and_retains_calculation_chain() {
    let bytes = scalar_formula_source();
    let editor =
        SourceBackedEditor::from_read_at(Arc::new(VersionedSource::new(bytes.clone()))).unwrap();
    let commit = editor.edit("Sheet1").unwrap().commit().unwrap();
    assert!(!commit.changed());
    assert!(commit.patch().is_empty());

    let mut output = Vec::new();
    editor
        .publish_commit_to_stream(&mut output, &commit)
        .unwrap();
    assert_eq!(output, bytes);
    assert!(has_calculation_chain(
        &OpcPackage::from_bytes(&output).unwrap()
    ));
}

#[test]
fn formula_patch_inverse_restores_calculation_topology_and_formula_cache() {
    let bytes = scalar_formula_source();
    let original = OpcPackage::from_bytes(&bytes).unwrap();
    let original_chain = original
        .get_part(&PackURI::new(CALC_CHAIN).unwrap())
        .unwrap()
        .blob()
        .to_vec();
    let original_content_types = zip_member(&bytes, "[Content_Types].xml");
    let editor = SourceBackedEditor::from_read_at(Arc::new(VersionedSource::new(bytes))).unwrap();
    let mut edit = editor.edit("Sheet1").unwrap();
    edit.set_formula(address("A1"), Formula::new("A2+2").unwrap())
        .unwrap();
    let commit = edit.commit().unwrap();

    let mut replay = original.clone();
    commit.patch().apply(&mut replay).unwrap();
    assert!(!has_calculation_chain(&replay));
    assert!(replay.get_part(&PackURI::new(CALC_CHAIN).unwrap()).is_err());
    let removed_content_types = PackageWriter::to_bytes(&replay).unwrap();
    assert!(
        !String::from_utf8(zip_member(&removed_content_types, "[Content_Types].xml"))
            .unwrap()
            .contains(CALC_CHAIN)
    );

    commit.patch().inverse().apply(&mut replay).unwrap();
    assert!(has_calculation_chain(&replay));
    assert_eq!(
        replay
            .get_part(&PackURI::new(CALC_CHAIN).unwrap())
            .unwrap()
            .blob(),
        original_chain.as_slice(),
    );
    let restored = PackageWriter::to_bytes(&replay).unwrap();
    assert_eq!(
        zip_member(&restored, "[Content_Types].xml"),
        original_content_types,
    );
    let snapshot = litchi_xlsx::cell_values::Snapshot::load_multi(&replay, "Sheet1").unwrap();
    assert!(matches!(
        snapshot.cell(address("A1")),
        Some(Cell::Formula(formula))
            if formula.text() == "A2+1" && formula.cached().is_some()
    ));
}

#[test]
fn formula_patch_refuses_another_calc_chain_inbound_edge_atomically() {
    let mut source = OpcPackage::from_bytes(&scalar_formula_source()).unwrap();
    source
        .get_part_mut(&PackURI::new(UNUSED).unwrap())
        .unwrap()
        .rels_mut()
        .try_add_relationship(
            "urn:litchi:test:calc-chain-alias".to_owned(),
            "../CALCCHAIN.XML".to_owned(),
            "rIdCalcChainAlias".to_owned(),
            TargetMode::Internal,
        )
        .unwrap();
    let bytes = PackageWriter::to_bytes(&source).unwrap();
    let editor =
        SourceBackedEditor::from_read_at(Arc::new(VersionedSource::new(bytes.clone()))).unwrap();
    let mut edit = editor.edit("Sheet1").unwrap();
    edit.set_formula(address("A1"), Formula::new("A2+2").unwrap())
        .unwrap();
    let commit = edit.commit().unwrap();

    let mut replay = OpcPackage::from_bytes(&bytes).unwrap();
    assert!(commit.patch().apply(&mut replay).is_err());
    assert!(has_calculation_chain(&replay));
    assert!(
        replay
            .get_part(&PackURI::new(UNUSED).unwrap())
            .unwrap()
            .rels()
            .get("rIdCalcChainAlias")
            .is_some()
    );
}

#[test]
fn formula_publication_rejects_a_stale_source_before_output() {
    let bytes = scalar_formula_source();
    let source = Arc::new(VersionedSource::new(bytes));
    let editor = SourceBackedEditor::from_read_at(source.clone()).unwrap();
    let mut edit = editor.edit("Sheet1").unwrap();
    edit.set_formula(address("A1"), Formula::new("A2+2").unwrap())
        .unwrap();
    let commit = edit.commit().unwrap();
    source.changed();

    let mut output = Vec::new();
    assert!(matches!(
        editor.publish_commit_to_stream(&mut output, &commit),
        Err(Error::Package(OpcError::SourceChanged { .. }))
    ));
    assert!(output.is_empty());
}
