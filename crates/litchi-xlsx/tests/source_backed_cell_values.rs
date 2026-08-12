#![allow(clippy::unwrap_used, reason = "focused integration assertions")]

use std::io;
use std::sync::Arc;
use std::sync::atomic::{AtomicU64, Ordering};

use litchi_core::{ReadAt, SourceVersion};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{BlobPart, OpcError, OpcPackage, PackURI, PackageWriter, TargetMode};
use litchi_xlsx::cell_values::{CellValueEdit, MAX_BATCH_EDITS, SourceBackedEditor};
use litchi_xlsx::{Address, Error, ErrorValue, Number, Value};

const SML: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const MAIN: &str = "/xl/workbook.xml";
const SHEET: &str = "/xl/worksheets/sheet1.xml";
const UNUSED: &str = "/xl/media/unused.bin";
static NEXT_SOURCE_ID: AtomicU64 = AtomicU64::new(1_000);

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

fn address(value: &str) -> Address {
    Address::from_a1(value).unwrap()
}

fn fixture(sheet_xml: String, signed: bool) -> Vec<u8> {
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
fn formulas_mce_shared_strings_relationships_and_signed_changes_are_refused() {
    for xml in [
        format!(
            r#"<worksheet xmlns="{SML}"><sheetData><row r="1"><c r="A1"><f>1+1</f><v>2</v></c></row></sheetData></worksheet>"#
        ),
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
