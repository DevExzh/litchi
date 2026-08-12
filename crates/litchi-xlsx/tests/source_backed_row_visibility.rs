#![allow(clippy::unwrap_used, reason = "focused integration assertions")]

use std::io;
use std::sync::Arc;
use std::sync::atomic::{AtomicU64, Ordering};

use litchi_core::{ReadAt, SourceVersion};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{BlobPart, OpcError, OpcPackage, PackURI, PackageWriter, TargetMode};
use litchi_xlsx::row_visibility::{
    MAX_BATCH_EDITS, RowVisibilityEdit, Snapshot, SourceBackedEditor,
};
use litchi_xlsx::{Error, RowIndex};

const SML: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const MAIN: &str = "/xl/workbook.xml";
const SHEET: &str = "/xl/worksheets/sheet1.xml";
const UNUSED: &str = "/xl/media/unused.bin";
static NEXT_SOURCE_ID: AtomicU64 = AtomicU64::new(20_000);

struct VersionedSource {
    bytes: Vec<u8>,
    id: u64,
    revision: AtomicU64,
}

impl VersionedSource {
    fn new(bytes: Vec<u8>) -> Self {
        Self {
            bytes,
            id: NEXT_SOURCE_ID.fetch_add(1, Ordering::Relaxed),
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
