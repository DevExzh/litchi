#![allow(
    clippy::unwrap_used,
    reason = "focused integration tests use panic-on-failure assertions"
)]

use std::io::{self, Write};
use std::num::{NonZeroU64, NonZeroUsize};
use std::sync::Arc;
use std::sync::atomic::{AtomicU64, Ordering};

use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits, ReadAt, SourceVersion,
};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{BlobPart, OpcError, OpcPackage, PackURI, PackageWriter, TargetMode};
use litchi_xlsx::{Error, MergeEditBlock, Rect, SourceBackedMergeEditor};

const SML: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const MAIN: &str = "/xl/workbook.xml";
const SHEET: &str = "/xl/worksheets/sheet1.xml";
const SECOND: &str = "/xl/worksheets/sheet2.xml";
const UNUSED: &str = "/xl/media/untouched.bin";

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
            9_021,
            self.revision.load(Ordering::SeqCst),
        ))
    }
}

fn shared_strings_xml(value: &str) -> Vec<u8> {
    format!(r#"<sst xmlns="{SML}" count="1" uniqueCount="1"><si><t>{value}</t></si></sst>"#)
        .into_bytes()
}

fn shared_strings_fixture() -> Vec<u8> {
    let sheet = format!(
        r#"<worksheet xmlns="{SML}"><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c></row></sheetData></worksheet>"#
    );
    let mut package = OpcPackage::from_bytes(&fixture(&sheet)).unwrap();
    package
        .try_add_part(Box::new(BlobPart::new(
            PackURI::new("/xl/sharedStrings.xml").unwrap(),
            ct::SML_SHARED_STRINGS.to_owned(),
            shared_strings_xml("anchor"),
        )))
        .unwrap();
    package
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
    PackageWriter::to_bytes(&package).unwrap()
}

fn changed_shared_strings_bytes(bytes: &[u8]) -> Vec<u8> {
    let mut package = OpcPackage::from_bytes(bytes).unwrap();
    package
        .get_part_mut(&PackURI::new("/xl/sharedStrings.xml").unwrap())
        .unwrap()
        .set_blob(shared_strings_xml("changed"));
    PackageWriter::to_bytes(&package).unwrap()
}

fn changed_shared_strings_relationship(bytes: &[u8]) -> Vec<u8> {
    let mut package = OpcPackage::from_bytes(bytes).unwrap();
    package
        .get_part_mut(&PackURI::new(MAIN).unwrap())
        .unwrap()
        .rels_mut()
        .remove("rIdShared")
        .unwrap();
    PackageWriter::to_bytes(&package).unwrap()
}

struct FailingWriter {
    remaining: usize,
}

impl Write for FailingWriter {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        if self.remaining == 0 {
            return Err(io::Error::new(io::ErrorKind::BrokenPipe, "test sink"));
        }
        let accepted = bytes.len().min(self.remaining);
        self.remaining -= accepted;
        Ok(accepted)
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

struct MutatingWriter {
    source: Arc<VersionedSource>,
    mutated: bool,
}

impl Write for MutatingWriter {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        if !self.mutated {
            self.source.changed();
            self.mutated = true;
        }
        Ok(bytes.len())
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

fn fixture(sheet_xml: &str) -> Vec<u8> {
    fixture_with_options(sheet_xml, false)
}

fn signed_fixture(sheet_xml: &str) -> Vec<u8> {
    fixture_with_options(sheet_xml, true)
}

fn fixture_with_options(sheet_xml: &str, signed: bool) -> Vec<u8> {
    let workbook = format!(
        r#"<workbook xmlns="{SML}" xmlns:r="{REL}"><sheets><sheet name="Sheet1" sheetId="1" r:id="rIdSheet"/><sheet name="Other" sheetId="2" r:id="rIdOther"/></sheets></workbook>"#
    );
    let mut package = OpcPackage::new();
    for (name, content_type, bytes) in [
        (MAIN, ct::SML_SHEET_MAIN, workbook.into_bytes()),
        (SHEET, ct::SML_WORKSHEET, sheet_xml.as_bytes().to_vec()),
        (
            SECOND,
            ct::SML_WORKSHEET,
            format!(
                r#"<worksheet xmlns="{SML}"><sheetData/><extLst><ext uri="urn:other">untouched</ext></extLst></worksheet>"#
            )
            .into_bytes(),
        ),
        (
            UNUSED,
            "application/octet-stream",
            (0..4096).map(|value| (value % 251) as u8).collect(),
        ),
    ] {
        package
            .try_add_part(Box::new(BlobPart::new(
                PackURI::new(name).unwrap(),
                content_type.to_owned(),
                bytes,
            )))
            .unwrap();
    }
    let workbook_part = package.get_part_mut(&PackURI::new(MAIN).unwrap()).unwrap();
    workbook_part
        .rels_mut()
        .try_add_relationship(
            rt::WORKSHEET.to_owned(),
            "worksheets/sheet1.xml".to_owned(),
            "rIdSheet".to_owned(),
            TargetMode::Internal,
        )
        .unwrap();
    workbook_part
        .rels_mut()
        .try_add_relationship(
            rt::WORKSHEET.to_owned(),
            "worksheets/sheet2.xml".to_owned(),
            "rIdOther".to_owned(),
            TargetMode::Internal,
        )
        .unwrap();
    package
        .get_part_mut(&PackURI::new(SHEET).unwrap())
        .unwrap()
        .rels_mut()
        .try_add_relationship(
            rt::DRAWING.to_owned(),
            "../media/untouched.bin".to_owned(),
            "rIdDrawing".to_owned(),
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

fn ordinary_sheet() -> String {
    format!(
        r#"<worksheet xmlns="{SML}"><dimension ref="A1:C3"/><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>anchor</t></is></c></row></sheetData><mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells></worksheet>"#
    )
}

fn part_bytes(bytes: &[u8], name: &str) -> Vec<u8> {
    OpcPackage::from_bytes(bytes)
        .unwrap()
        .get_part(&PackURI::new(name).unwrap())
        .unwrap()
        .blob()
        .to_vec()
}

fn package_part_bytes(package: &OpcPackage, name: &str) -> Vec<u8> {
    package
        .get_part(&PackURI::new(name).unwrap())
        .unwrap()
        .blob()
        .to_vec()
}

fn part_relationship_count(bytes: &[u8], name: &str) -> usize {
    OpcPackage::from_bytes(bytes)
        .unwrap()
        .get_part(&PackURI::new(name).unwrap())
        .unwrap()
        .rels()
        .len()
}

fn rect(reference: &str) -> Rect {
    Rect::from_a1(reference).unwrap()
}

#[test]
fn source_merge_commit_unmerge_and_overlay_preserve_unselected_members() {
    let source_bytes = fixture(&ordinary_sheet());
    let untouched = part_bytes(&source_bytes, UNUSED);
    let selected_relationships = part_relationship_count(&source_bytes, SHEET);
    let source = Arc::new(VersionedSource::new(source_bytes.clone()));
    let editor = SourceBackedMergeEditor::from_read_at(source).unwrap();

    let snapshot = editor.snapshot("Sheet1").unwrap().unwrap();
    assert_eq!(snapshot.merge_count(), 1);
    assert!(snapshot.contains_merge(rect("A1:B1")));

    let mut edit = editor.edit("Sheet1").unwrap().unwrap();
    edit.merge("A2:B2").unwrap();
    let commit = edit.commit().unwrap();
    assert!(commit.changed());
    assert_eq!(commit.diagnostics().touched_worksheets(), 1);
    assert!(commit.snapshot().contains_merge(rect("A2:B2")));

    let mut output = Vec::new();
    editor
        .publish_commit_to_stream(&mut output, &commit)
        .unwrap();
    let selected = String::from_utf8(part_bytes(&output, SHEET)).unwrap();
    assert!(selected.contains(r#"ref="A1:B1""#));
    assert!(selected.contains(r#"ref="A2:B2""#));
    assert_eq!(part_bytes(&output, UNUSED), untouched);
    assert_eq!(
        part_relationship_count(&output, SHEET),
        selected_relationships
    );
    assert!(
        String::from_utf8(part_bytes(&output, SECOND))
            .unwrap()
            .contains("untouched")
    );

    let inverse = commit.patch().inverse();
    assert!(inverse.before().contains_merge(rect("A2:B2")));
    assert!(inverse.after().contains_merge(rect("A1:B1")));
}

#[test]
fn source_merge_noop_publication_preserves_exact_bytes_including_signed_sources() {
    for bytes in [
        fixture(&ordinary_sheet()),
        signed_fixture(&ordinary_sheet()),
    ] {
        let editor =
            SourceBackedMergeEditor::from_read_at(Arc::new(VersionedSource::new(bytes.clone())))
                .unwrap();
        let commit = editor.edit("Sheet1").unwrap().unwrap().commit().unwrap();
        assert!(!commit.changed());
        let mut output = Vec::new();
        editor
            .publish_commit_to_stream(&mut output, &commit)
            .unwrap();
        assert_eq!(output, bytes);
    }
}

#[test]
fn source_merge_patch_applies_forward_and_inverse_exactly() {
    let source_bytes = fixture(&ordinary_sheet());
    let original_sheet = part_bytes(&source_bytes, SHEET);
    let editor =
        SourceBackedMergeEditor::from_read_at(Arc::new(VersionedSource::new(source_bytes.clone())))
            .unwrap();
    let mut edit = editor.edit("Sheet1").unwrap().unwrap();
    edit.merge("A2:B2").unwrap();
    let commit = edit.commit().unwrap();
    let mut package = OpcPackage::from_bytes(&source_bytes).unwrap();

    commit.patch().apply(&mut package).unwrap();
    assert!(
        String::from_utf8(package_part_bytes(&package, SHEET))
            .unwrap()
            .contains(r#"ref="A2:B2""#)
    );
    commit.patch().inverse().apply(&mut package).unwrap();
    assert_eq!(package_part_bytes(&package, SHEET), original_sheet);
}

#[test]
fn source_merge_patch_rejects_shared_strings_payload_and_relationship_drift() {
    let source_bytes = shared_strings_fixture();
    for changed_bytes in [
        changed_shared_strings_bytes(&source_bytes),
        changed_shared_strings_relationship(&source_bytes),
    ] {
        let editor = SourceBackedMergeEditor::from_read_at(Arc::new(VersionedSource::new(
            source_bytes.clone(),
        )))
        .unwrap();
        let mut edit = editor.edit("Sheet1").unwrap().unwrap();
        edit.merge("A2:B2").unwrap();
        let commit = edit.commit().unwrap();
        let mut package = OpcPackage::from_bytes(&changed_bytes).unwrap();
        assert!(matches!(
            commit.patch().apply(&mut package),
            Err(Error::PatchConflict { .. })
        ));
    }
}

#[test]
fn source_merge_managed_inverse_requires_bounded_materialization() {
    let source_bytes = fixture(&ordinary_sheet());
    let worksheet_bytes = part_bytes(&source_bytes, SHEET).len();
    let budget = Budget::root(
        "xlsx-source-merge-materialization-test",
        Limits::new(u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX),
    );
    let (cancellation_source, cancellation) = CancellationSource::pair();
    let execution_limits = ExecutionLimits::new(
        NonZeroUsize::new(32).unwrap(),
        NonZeroUsize::new(32).unwrap(),
        NonZeroU64::new(u64::MAX).unwrap(),
        0,
    )
    .unwrap();
    let context = ExecutionContext::new(budget, cancellation, execution_limits);
    let editor = SourceBackedMergeEditor::from_read_at_with_execution_context(
        Arc::new(VersionedSource::new(source_bytes.clone())),
        litchi_xlsx::ReadLimits::default(),
        context,
    )
    .unwrap();
    let mut edit = editor.edit("Sheet1").unwrap().unwrap();
    edit.merge("A2:B2").unwrap();
    let commit = edit.commit().unwrap();
    let mut package = OpcPackage::from_bytes(&source_bytes).unwrap();
    commit.patch().apply(&mut package).unwrap();
    assert!(matches!(
        commit.patch().inverse().apply(&mut package),
        Err(Error::Package(OpcError::ManagedPartDataArcEscape))
    ));
    commit
        .patch()
        .inverse()
        .apply_materialized(&mut package, worksheet_bytes)
        .unwrap();
    assert_eq!(
        package_part_bytes(&package, SHEET),
        part_bytes(&source_bytes, SHEET)
    );
    cancellation_source.cancel();
}

#[test]
fn source_merge_changed_publication_refuses_signed_source_before_output() {
    let signed_bytes = signed_fixture(&ordinary_sheet());
    let signed_editor =
        SourceBackedMergeEditor::from_read_at(Arc::new(VersionedSource::new(signed_bytes.clone())))
            .unwrap();
    let mut edit = signed_editor.edit("Sheet1").unwrap().unwrap();
    edit.merge("A2:B2").unwrap();
    let commit = edit.commit().unwrap();
    let mut output = Vec::new();
    assert!(matches!(
        signed_editor.publish_commit_to_stream(&mut output, &commit),
        Err(Error::Package(OpcError::SignedSourceRequiresExplicitPolicy))
    ));
    assert!(output.is_empty());
}

#[test]
fn source_merge_unmerge_is_a_semantic_noop_for_missing_ranges() {
    let bytes = fixture(&ordinary_sheet());
    let editor =
        SourceBackedMergeEditor::from_read_at(Arc::new(VersionedSource::new(bytes))).unwrap();
    let mut edit = editor.edit("Sheet1").unwrap().unwrap();
    edit.unmerge("Z99").unwrap();
    let commit = edit.commit().unwrap();
    assert!(!commit.changed());
    assert!(commit.patch().is_empty());
}

#[test]
fn source_merge_selector_miss_is_an_option() {
    let editor = SourceBackedMergeEditor::from_read_at(Arc::new(VersionedSource::new(fixture(
        &ordinary_sheet(),
    ))))
    .unwrap();
    assert!(editor.snapshot("Missing").unwrap().is_none());
    assert!(editor.edit("Missing").unwrap().is_none());
}

#[test]
fn source_merge_staging_is_canonical_and_set_based() {
    let editor = SourceBackedMergeEditor::from_read_at(Arc::new(VersionedSource::new(fixture(
        &ordinary_sheet(),
    ))))
    .unwrap();
    let mut edit = editor.edit("Sheet1").unwrap().unwrap();
    edit.unmerge("A1").unwrap().merge("A1:B1").unwrap();
    assert!(!edit.is_changed());
    assert_eq!(edit.merges().collect::<Vec<_>>(), vec![rect("A1:B1")]);
    assert!(!edit.commit().unwrap().changed());

    let editor = SourceBackedMergeEditor::from_read_at(Arc::new(VersionedSource::new(fixture(
        &ordinary_sheet(),
    ))))
    .unwrap();
    let mut edit = editor.edit("Sheet1").unwrap().unwrap();
    edit.unmerge("A1").unwrap().merge("A1:C1").unwrap();
    let commit = edit.commit().unwrap();
    assert!(commit.changed());
    assert_eq!(
        commit.snapshot().merges().collect::<Vec<_>>(),
        vec![rect("A1:C1")]
    );

    let empty = format!(r#"<worksheet xmlns="{SML}"><sheetData/></worksheet>"#);
    let editor =
        SourceBackedMergeEditor::from_read_at(Arc::new(VersionedSource::new(fixture(&empty))))
            .unwrap();
    let mut edit = editor.edit("Sheet1").unwrap().unwrap();
    edit.merge("C1:D1").unwrap().merge("A1:B1").unwrap();
    assert_eq!(
        edit.merges().collect::<Vec<_>>(),
        vec![rect("A1:B1"), rect("C1:D1")]
    );
    let commit = edit.commit().unwrap();
    assert_eq!(
        commit.snapshot().merges().collect::<Vec<_>>(),
        vec![rect("A1:B1"), rect("C1:D1")]
    );
}

#[test]
fn source_merge_structural_refusal_precedes_follower_content() {
    let protected = format!(
        r#"<worksheet xmlns="{SML}"><sheetProtection sheet="1"/><sheetData><row r="1"><c r="B1"><v>7</v></c></row></sheetData></worksheet>"#
    );
    let editor =
        SourceBackedMergeEditor::from_read_at(Arc::new(VersionedSource::new(fixture(&protected))))
            .unwrap();
    let mut edit = editor.edit("Sheet1").unwrap().unwrap();
    edit.merge("A1:B1").unwrap();
    assert!(matches!(
        edit.commit(),
        Err(Error::MergeEditBlocked {
            reason: MergeEditBlock::ProtectedSheet,
            ..
        })
    ));

    let markup_compatibility = format!(
        r#"<worksheet xmlns="{SML}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:litchi:unknown" mc:Ignorable="x"><sheetData><row r="1"><c r="B1"><v>7</v></c></row></sheetData><mc:AlternateContent><mc:Choice Requires="x"><mergeCells count="1"><mergeCell ref="C1:D1"/></mergeCells></mc:Choice><mc:Fallback/></mc:AlternateContent></worksheet>"#
    );
    let editor = SourceBackedMergeEditor::from_read_at(Arc::new(VersionedSource::new(fixture(
        &markup_compatibility,
    ))))
    .unwrap();
    let mut edit = editor.edit("Sheet1").unwrap().unwrap();
    edit.merge("A1:B1").unwrap();
    assert!(matches!(
        edit.commit(),
        Err(Error::MergeEditBlocked {
            reason: MergeEditBlock::MarkupCompatibility,
            ..
        })
    ));
}

#[test]
fn source_merge_unmerge_removes_existing_range_and_publishes_one_overlay() {
    let source_bytes = fixture(&ordinary_sheet());
    let editor =
        SourceBackedMergeEditor::from_read_at(Arc::new(VersionedSource::new(source_bytes)))
            .unwrap();
    let mut edit = editor.edit("Sheet1").unwrap().unwrap();
    edit.unmerge("A1").unwrap();
    let commit = edit.commit().unwrap();
    assert!(commit.changed());
    let mut output = Vec::new();
    editor
        .publish_commit_to_stream(&mut output, &commit)
        .unwrap();
    let selected = String::from_utf8(part_bytes(&output, SHEET)).unwrap();
    assert!(!selected.contains(r#"ref="A1:B1""#));
}

#[test]
fn source_merge_refuses_followers_overlap_protection_and_unknown_payload() {
    let follower = format!(
        r#"<worksheet xmlns="{SML}"><sheetData><row r="1"><c r="B1"><v>7</v></c></row></sheetData></worksheet>"#
    );
    let editor =
        SourceBackedMergeEditor::from_read_at(Arc::new(VersionedSource::new(fixture(&follower))))
            .unwrap();
    let mut edit = editor.edit("Sheet1").unwrap().unwrap();
    edit.merge("A1:B1").unwrap();
    assert!(matches!(edit.commit(), Err(Error::MergeEditBlocked { .. })));

    let overlap = format!(
        r#"<worksheet xmlns="{SML}"><sheetData/><mergeCells count="1"><mergeCell ref="A1:B2"/></mergeCells></worksheet>"#
    );
    let editor =
        SourceBackedMergeEditor::from_read_at(Arc::new(VersionedSource::new(fixture(&overlap))))
            .unwrap();
    let mut edit = editor.edit("Sheet1").unwrap().unwrap();
    edit.merge("B2:C3").unwrap();
    assert!(matches!(edit.commit(), Err(Error::MergeEditBlocked { .. })));

    let protected =
        format!(r#"<worksheet xmlns="{SML}"><sheetProtection sheet="1"/><sheetData/></worksheet>"#);
    let editor =
        SourceBackedMergeEditor::from_read_at(Arc::new(VersionedSource::new(fixture(&protected))))
            .unwrap();
    let mut edit = editor.edit("Sheet1").unwrap().unwrap();
    edit.merge("A1:B1").unwrap();
    assert!(matches!(edit.commit(), Err(Error::MergeEditBlocked { .. })));

    let unknown = format!(
        r#"<worksheet xmlns="{SML}"><sheetData/><mergeCells><mergeCell ref="A1:B1"/><unknown/></mergeCells></worksheet>"#
    );
    assert!(
        SourceBackedMergeEditor::from_read_at(Arc::new(VersionedSource::new(fixture(&unknown))))
            .unwrap()
            .snapshot("Sheet1")
            .is_err()
    );
}

#[test]
fn source_merge_accepts_empty_followers_and_preserves_a_rich_anchor() {
    let rich = format!(
        r#"<worksheet xmlns="{SML}"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><r><rPr><b/><color rgb="FFFF0000"/></rPr><t>rich anchor</t></r></is></c><c r="B1"/></row></sheetData></worksheet>"#
    );
    let source_bytes = fixture(&rich);
    let editor =
        SourceBackedMergeEditor::from_read_at(Arc::new(VersionedSource::new(source_bytes)))
            .unwrap();
    let mut edit = editor.edit("Sheet1").unwrap().unwrap();
    edit.merge("A1:B1").unwrap();
    let commit = edit.commit().unwrap();
    assert!(commit.snapshot().contains_merge(rect("A1:B1")));

    let mut output = Vec::new();
    editor
        .publish_commit_to_stream(&mut output, &commit)
        .unwrap();
    let selected = String::from_utf8(part_bytes(&output, SHEET)).unwrap();
    assert!(selected.contains(r#"<rPr><b/><color rgb="FFFF0000"/></rPr>"#));
    assert!(selected.contains("rich anchor"));
}

#[test]
fn source_merge_refuses_array_data_table_and_shared_formula_groups() {
    for formula_type in ["array", "dataTable", "shared"] {
        let sheet = format!(
            r#"<worksheet xmlns="{SML}"><sheetData><row r="1"><c r="A1"><f t="{formula_type}" ref="A1:B1" si="0">SUM(A1)</f><v>1</v></c><c r="B1"/></row></sheetData></worksheet>"#
        );
        let editor =
            SourceBackedMergeEditor::from_read_at(Arc::new(VersionedSource::new(fixture(&sheet))))
                .unwrap();
        let mut edit = editor.edit("Sheet1").unwrap().unwrap();
        edit.merge("A1:B1").unwrap();
        assert!(matches!(
            edit.commit(),
            Err(Error::MergeEditBlocked {
                reason: MergeEditBlock::GroupFormula,
                ..
            })
        ));
    }
}

#[test]
fn source_merge_rejects_stale_and_foreign_lineage_before_output() {
    let bytes = fixture(&ordinary_sheet());
    let source = Arc::new(VersionedSource::new(bytes.clone()));
    let stale_editor = SourceBackedMergeEditor::from_read_at(source.clone()).unwrap();
    let mut edit = stale_editor.edit("Sheet1").unwrap().unwrap();
    edit.merge("A2:B2").unwrap();
    let commit = edit.commit().unwrap();
    source.changed();
    assert!(matches!(
        stale_editor.publish_commit_to_stream(Vec::new(), &commit),
        Err(Error::Package(OpcError::SourceChanged { .. }))
    ));

    let foreign_editor =
        SourceBackedMergeEditor::from_read_at(Arc::new(VersionedSource::new(bytes))).unwrap();
    assert!(matches!(
        foreign_editor.publish_commit_to_stream(Vec::new(), &commit),
        Err(Error::PatchConflict { .. })
    ));
}

#[test]
fn source_merge_limit_refusal_is_typed() {
    let bytes = fixture(&ordinary_sheet());
    let limits = litchi_xlsx::ReadLimits::builder()
        .max_input_bytes(1)
        .unwrap()
        .build()
        .unwrap();
    assert!(matches!(
        SourceBackedMergeEditor::from_read_at_with_limits(
            Arc::new(VersionedSource::new(bytes)),
            limits,
        ),
        Err(Error::Package(_))
    ));
}

#[test]
fn source_merge_honors_managed_cancellation_before_selected_read() {
    let bytes = fixture(&ordinary_sheet());
    let budget = Budget::root(
        "xlsx-source-merge-cancellation-test",
        Limits::new(u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX),
    );
    let (cancellation_source, cancellation) = CancellationSource::pair();
    let execution_limits = ExecutionLimits::new(
        NonZeroUsize::new(1).unwrap(),
        NonZeroUsize::new(1).unwrap(),
        NonZeroU64::new(u64::MAX).unwrap(),
        0,
    )
    .unwrap();
    let context = ExecutionContext::new(budget, cancellation, execution_limits);
    let editor = SourceBackedMergeEditor::from_read_at_with_execution_context(
        Arc::new(VersionedSource::new(bytes)),
        litchi_xlsx::ReadLimits::default(),
        context,
    )
    .unwrap();
    cancellation_source.cancel();
    assert!(matches!(editor.snapshot("Sheet1"), Err(Error::Package(_))));
}

#[test]
fn source_merge_publication_reports_bounded_sink_and_source_failures() {
    let source = Arc::new(VersionedSource::new(fixture(&ordinary_sheet())));
    let editor = SourceBackedMergeEditor::from_read_at(source.clone()).unwrap();
    let mut edit = editor.edit("Sheet1").unwrap().unwrap();
    edit.merge("A2:B2").unwrap();
    let commit = edit.commit().unwrap();
    let result = editor.publish_commit_to_stream(FailingWriter { remaining: 1 }, &commit);
    assert!(matches!(
        result,
        Err(Error::Package(OpcError::IncompleteOutput { written, .. })) if written > 0
    ));

    let source = Arc::new(VersionedSource::new(fixture(&ordinary_sheet())));
    let editor = SourceBackedMergeEditor::from_read_at(source.clone()).unwrap();
    let mut edit = editor.edit("Sheet1").unwrap().unwrap();
    edit.merge("A2:B2").unwrap();
    let commit = edit.commit().unwrap();
    let result = editor.publish_commit_to_stream(
        MutatingWriter {
            source,
            mutated: false,
        },
        &commit,
    );
    assert!(matches!(
        result,
        Err(Error::Package(OpcError::IncompleteOutput { source, .. }))
            if matches!(source.as_ref(), OpcError::SourceChanged { .. })
    ));
}
