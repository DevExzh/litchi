#![allow(
    clippy::unwrap_used,
    reason = "focused integration tests use panic-on-failure assertions"
)]

use std::io::{self, Write};
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
use litchi_xlsx::auto_filter::{
    Calendar, Color, Column, Condition, Definition, Item, Payload, Range, SourceBackedEditor,
    State, Values,
};
use litchi_xlsx::sort::{SortBy, SortMethod};
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
            202,
            self.revision.load(Ordering::SeqCst),
        ))
    }
}

struct FailingSink(usize);

impl Write for FailingSink {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        if self.0 == 0 {
            return Err(io::Error::other("injected sink failure"));
        }
        let written = bytes.len().min(self.0);
        self.0 -= written;
        Ok(written)
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

fn fixture(sheet_xml: &str, signed: bool) -> Vec<u8> {
    let workbook = format!(
        r#"<workbook xmlns="{SML}" xmlns:r="{REL}"><sheets><sheet name="Sheet1" sheetId="1" r:id="rIdSheet"/><sheet name="Unused" sheetId="2" r:id="rIdUnused"/></sheets><calcPr calcId="7"/></workbook>"#
    );
    let mut package = OpcPackage::new();
    for (name, content_type, bytes) in [
        (MAIN, ct::SML_SHEET_MAIN, workbook.into_bytes()),
        (SHEET, ct::SML_WORKSHEET, sheet_xml.as_bytes().to_vec()),
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
            "rIdUnused".to_owned(),
            TargetMode::Internal,
        )
        .unwrap();
    package
        .get_part_mut(&PackURI::new(SHEET).unwrap())
        .unwrap()
        .rels_mut()
        .try_add_relationship(
            rt::DRAWING.to_owned(),
            "../media/unused.bin".to_owned(),
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

fn ordinary_fixture(signed: bool) -> Vec<u8> {
    fixture(
        &format!(r#"<worksheet xmlns="{SML}"><sheetData/></worksheet>"#),
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
        "xlsx-auto-filter-managed-test",
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

fn definition(updated: bool) -> Definition {
    let mut value = Definition::new(Some(
        Range::new(if updated { "A1:C12" } else { "A1:C8" }).unwrap(),
    ));
    let mut column = Column::new(0).unwrap();
    column.set_payload(Some(Payload::Values(
        Values::new(
            false,
            Calendar::None,
            vec![Item::Value(
                if updated { "violet" } else { "blue" }.to_owned(),
            )],
        )
        .unwrap(),
    )));
    value.columns.push(column);
    value
        .set_sort_state(Some(
            State::new(
                Range::new(if updated { "A2:C12" } else { "A2:C8" }).unwrap(),
                false,
                updated,
                Some(SortMethod::None),
                vec![Condition::new(
                    Range::new(if updated { "B2:B12" } else { "B2:B8" }).unwrap(),
                    updated,
                    SortBy::Value,
                )],
            )
            .unwrap(),
        ))
        .unwrap();
    value
}

fn changed_commit(editor: &SourceBackedEditor) -> litchi_xlsx::auto_filter::Commit {
    let mut edit = editor.edit("Sheet1").unwrap();
    assert!(edit.set(definition(true)).unwrap());
    edit.commit().unwrap()
}

fn color_definition() -> Definition {
    let mut value = Definition::new(Some(Range::new("A1:C12").unwrap()));
    let mut column = Column::new(0).unwrap();
    column.set_payload(Some(Payload::Color(Color::new(0, true))));
    value.columns.push(column);
    value
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
    assert!(commit.changed());
    assert_eq!(commit.diagnostics().touched_worksheets(), 1);
    assert_eq!(editor.cache_diagnostics().successful_loads, 2);

    let mut replay = OpcPackage::from_bytes(&source_bytes).unwrap();
    commit.patch().apply(&mut replay).unwrap();
    assert_eq!(
        litchi_xlsx::auto_filter::Snapshot::load(&replay, "Sheet1")
            .unwrap()
            .auto_filter(),
        Some(&definition(true))
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

    let editor =
        SourceBackedEditor::from_read_at(Arc::new(VersionedSource::new(source_bytes.clone())))
            .unwrap();
    let commit = changed_commit(&editor);
    let mut output = Vec::new();
    editor
        .publish_commit_to_stream(&mut output, &commit)
        .unwrap();
    let source = OpcPackage::from_bytes(&source_bytes).unwrap();
    let candidate = OpcPackage::from_bytes(&output).unwrap();
    for part in source.iter_parts() {
        let rewritten = candidate.get_part(part.partname()).unwrap();
        assert_eq!(rewritten.content_type(), part.content_type());
        assert_eq!(
            relationship_signatures(rewritten.rels()),
            relationship_signatures(part.rels())
        );
        if part.partname() != &PackURI::new(SHEET).unwrap() {
            assert_eq!(rewritten.blob(), part.blob());
        }
    }
    assert_eq!(
        Package::from_opc(candidate)
            .unwrap()
            .workbook()
            .unwrap()
            .sheet("Sheet1")
            .unwrap()
            .unwrap()
            .auto_filter()
            .unwrap(),
        Some(definition(true))
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
            litchi_xlsx::ReadLimits::default(),
            SourceCacheLimits::new(usize::try_from(exact).unwrap(), 8).unwrap(),
            context,
        )
        .unwrap();
    let snapshot = editor.snapshot("Sheet1").unwrap();
    assert!(snapshot.auto_filter().is_none());
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
        litchi_xlsx::auto_filter::Snapshot::load(&published, "Sheet1")
            .unwrap()
            .auto_filter(),
        Some(&definition(true))
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
fn managed_cancellation_is_checked_before_stream_output() {
    let source_bytes = ordinary_fixture(false);
    let exact = part_len(&source_bytes, MAIN) + part_len(&source_bytes, SHEET);
    let (budget, cancellation_source, context) = managed_context(exact);
    let editor = SourceBackedEditor::from_read_at_with_execution_context(
        Arc::new(VersionedSource::new(source_bytes)),
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
fn add_replace_clear_noop_protection_and_mce_are_checked() {
    let source = ordinary_fixture(false);
    let editor =
        SourceBackedEditor::from_read_at(Arc::new(VersionedSource::new(source.clone()))).unwrap();
    let mut edit = editor.edit("Sheet1").unwrap();
    assert!(edit.set(definition(false)).unwrap());
    assert!(!edit.set(definition(false)).unwrap());
    assert!(edit.clear());
    assert!(!edit.clear());
    assert!(!edit.commit().unwrap().changed());

    let protected = fixture(
        &format!(
            r#"<worksheet xmlns="{SML}"><sheetData/><sheetProtection sheet="1" sort="1" autoFilter="1"/></worksheet>"#
        ),
        false,
    );
    let editor =
        SourceBackedEditor::from_read_at(Arc::new(VersionedSource::new(protected))).unwrap();
    let mut edit = editor.edit("Sheet1").unwrap();
    edit.set(definition(true)).unwrap();
    assert!(edit.commit().is_err());

    let mce = fixture(
        &format!(
            r#"<worksheet xmlns="{SML}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><sheetData/><mc:AlternateContent><mc:Choice Requires="x14"><autoFilter ref="A1:C8"/></mc:Choice><mc:Fallback/></mc:AlternateContent></worksheet>"#
        ),
        false,
    );
    let editor = SourceBackedEditor::from_read_at(Arc::new(VersionedSource::new(mce))).unwrap();
    let mut edit = editor.edit("Sheet1").unwrap();
    edit.set(definition(true)).unwrap();
    assert!(edit.commit().is_err());
}

#[test]
fn signed_noop_is_exact_and_conflicts_and_partial_output_fail_closed() {
    let signed = ordinary_fixture(true);
    let editor =
        SourceBackedEditor::from_read_at(Arc::new(VersionedSource::new(signed.clone()))).unwrap();
    let commit = editor.edit("Sheet1").unwrap().commit().unwrap();
    let mut output = Vec::new();
    editor
        .publish_commit_to_stream(&mut output, &commit)
        .unwrap();
    assert_eq!(output, signed);

    let editor =
        SourceBackedEditor::from_read_at(Arc::new(VersionedSource::new(ordinary_fixture(true))))
            .unwrap();
    let commit = changed_commit(&editor);
    assert!(
        editor
            .publish_commit_to_stream(Vec::new(), &commit)
            .is_err()
    );

    let source = Arc::new(VersionedSource::new(ordinary_fixture(false)));
    let editor = SourceBackedEditor::from_read_at(source.clone()).unwrap();
    let commit = changed_commit(&editor);
    source.changed();
    assert!(
        editor
            .publish_commit_to_stream(Vec::new(), &commit)
            .is_err()
    );

    let editor =
        SourceBackedEditor::from_read_at(Arc::new(VersionedSource::new(ordinary_fixture(false))))
            .unwrap();
    let commit = changed_commit(&editor);
    assert!(
        editor
            .publish_commit_to_stream(FailingSink(32), &commit)
            .is_err()
    );

    let mut foreign = OpcPackage::from_bytes(&ordinary_fixture(false)).unwrap();
    foreign
        .get_part_mut(&PackURI::new(SHEET).unwrap())
        .unwrap()
        .set_blob(
            format!(r#"<worksheet xmlns="{SML}"><sheetData/><extLst/></worksheet>"#).into_bytes(),
        );
    assert!(matches!(
        commit.patch().apply(&mut foreign),
        Err(Error::PatchConflict { .. })
    ));

    let mut foreign_styles = OpcPackage::from_bytes(&ordinary_fixture(false)).unwrap();
    foreign_styles
        .try_add_part(Box::new(BlobPart::new(
            PackURI::new("/xl/styles.xml").unwrap(),
            ct::SML_STYLES.to_owned(),
            format!(r#"<styleSheet xmlns="{SML}"/>"#).into_bytes(),
        )))
        .unwrap();
    foreign_styles
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
    assert!(matches!(
        commit.patch().apply(&mut foreign_styles),
        Err(Error::PatchConflict { .. })
    ));
}

#[test]
fn differential_format_references_bind_the_styles_part() {
    let mut package = OpcPackage::from_bytes(&ordinary_fixture(false)).unwrap();
    let styles_uri = PackURI::new("/xl/styles.xml").unwrap();
    package
        .try_add_part(Box::new(BlobPart::new(
            styles_uri.clone(),
            ct::SML_STYLES.to_owned(),
            format!(
                r#"<styleSheet xmlns="{SML}"><dxfs count="1"><dxf><fill/></dxf></dxfs></styleSheet>"#
            )
            .into_bytes(),
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
    let source = PackageWriter::to_bytes(&package).unwrap();
    let editor = SourceBackedEditor::from_read_at(Arc::new(VersionedSource::new(source))).unwrap();
    let mut edit = editor.edit("Sheet1").unwrap();
    assert!(edit.set(color_definition()).unwrap());
    assert!(edit.commit().is_ok());
    assert_eq!(editor.cache_diagnostics().successful_loads, 3);

    let editor =
        SourceBackedEditor::from_read_at(Arc::new(VersionedSource::new(ordinary_fixture(false))))
            .unwrap();
    let mut edit = editor.edit("Sheet1").unwrap();
    assert!(edit.set(color_definition()).unwrap());
    assert!(edit.commit().is_err());
}
