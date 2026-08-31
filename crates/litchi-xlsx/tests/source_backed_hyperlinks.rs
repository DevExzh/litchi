#![allow(
    clippy::unwrap_used,
    reason = "focused integration tests use panic-on-failure assertions"
)]

use std::collections::BTreeMap;
use std::io::{self, Write};
use std::num::{NonZeroU64, NonZeroUsize};
use std::sync::Arc;
use std::sync::atomic::{AtomicU64, AtomicUsize, Ordering};

use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionError, ExecutionLimits, Limits, ReadAt,
    Resource, SourceVersion,
};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{OpcError, OpcPackage, PackURI};
use litchi_xlsx::{
    Error, Hyperlink, HyperlinkReference, Package, ReadLimits, SourceBackedHyperlinkCommit,
    SourceBackedHyperlinkEdit, SourceBackedHyperlinkEditor, SourceBackedHyperlinkPatch,
    SourceBackedHyperlinkSnapshot,
};
use soapberry_zip::{PreservationIndex, ZipArchive};

const SML: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PACKAGE_REL: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const MAIN: &str = "/xl/workbook.xml";
const SHEET: &str = "/xl/worksheets/sheet1.xml";
const SECOND: &str = "/xl/worksheets/sheet2.xml";
const SECOND_MARKER: &str = "source-backed-hyperlink-unselected-sheet-marker";
const EXTERNAL_TARGET: &str = "https://127.0.0.1:9/original?q=1#fragment";
const NEW_EXTERNAL_TARGET: &str = "https://127.0.0.1:9/new?q=2#fragment";

struct CountingSource {
    bytes: Vec<u8>,
    second_marker_offset: usize,
    second_body_marker_reads: AtomicUsize,
    revision: AtomicU64,
}

impl CountingSource {
    fn new(bytes: Vec<u8>) -> Self {
        let second_marker_offset = bytes
            .windows(SECOND_MARKER.len())
            .position(|window| window == SECOND_MARKER.as_bytes())
            .expect("second worksheet marker is stored in archive");
        Self {
            bytes,
            second_marker_offset,
            second_body_marker_reads: AtomicUsize::new(0),
            revision: AtomicU64::new(0),
        }
    }

    fn changed(&self) {
        self.revision.fetch_add(1, Ordering::SeqCst);
    }
}

impl ReadAt for CountingSource {
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
        if offset < self.second_marker_offset + SECOND_MARKER.len()
            && self.second_marker_offset < end
        {
            self.second_body_marker_reads.fetch_add(1, Ordering::SeqCst);
        }
        output[..end - offset].copy_from_slice(&self.bytes[offset..end]);
        Ok(end - offset)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        Ok(SourceVersion::new(
            342,
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

#[derive(Debug, PartialEq, Eq)]
struct RawMember {
    local: Vec<u8>,
    central_without_offset: Vec<u8>,
}

fn raw_members(bytes: &[u8]) -> BTreeMap<String, RawMember> {
    let archive = ZipArchive::from_slice(bytes).unwrap().into_zip_archive();
    let mut buffer = vec![0_u8; soapberry_zip::RECOMMENDED_BUFFER_SIZE];
    let index = PreservationIndex::new(&archive, &mut buffer).unwrap();
    let mut records = archive.entries(&mut buffer);
    index
        .entries()
        .iter()
        .map(|preserved| {
            let record = records.next_entry().unwrap().unwrap();
            let name = record
                .file_path()
                .try_normalize()
                .unwrap()
                .as_ref()
                .to_owned();
            let local = bytes
                [preserved.local_span().start as usize..preserved.local_span().end as usize]
                .to_vec();
            let central_range = preserved.central_record();
            let mut central_without_offset =
                bytes[central_range.start as usize..central_range.end as usize].to_vec();
            central_without_offset[42..46].fill(0);
            (
                name,
                RawMember {
                    local,
                    central_without_offset,
                },
            )
        })
        .collect()
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

fn internal(
    reference: &str,
    location: &str,
    display: Option<&str>,
    tooltip: Option<&str>,
) -> Hyperlink {
    Hyperlink::new(
        HyperlinkReference::new(reference).unwrap(),
        Some(location.to_owned()),
        None,
        display.map(String::from),
        tooltip.map(String::from),
    )
    .unwrap()
}

fn external(
    reference: &str,
    target: &str,
    display: Option<&str>,
    tooltip: Option<&str>,
) -> Hyperlink {
    Hyperlink::new(
        HyperlinkReference::new(reference).unwrap(),
        None,
        Some(target.to_owned()),
        display.map(String::from),
        tooltip.map(String::from),
    )
    .unwrap()
}

fn worksheet_xml() -> String {
    format!(
        r#"<worksheet xmlns="{SML}" xmlns:r="{REL}"><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData><hyperlinks><hyperlink ref="A1" location="Sheet1!B2" display="local"/><hyperlink ref="C3" r:id="rIdHyperlink" display="remote" tooltip="old"/></hyperlinks><extLst><ext uri="urn:opaque"><vendor:opaque xmlns:vendor="urn:vendor">keep</vendor:opaque></ext></extLst></worksheet>"#
    )
}

fn fixture_with_sheet(sheet_xml: String, signed: bool) -> Vec<u8> {
    let workbook_xml = format!(
        r#"<workbook xmlns="{SML}" xmlns:r="{REL}"><sheets><sheet name="Sheet1" sheetId="1" r:id="rIdSheet"/><sheet name="Second" sheetId="2" r:id="rIdSecond"/></sheets></workbook>"#
    );
    let second_xml = format!(
        r#"<worksheet xmlns="{SML}"><sheetData/><!--{SECOND_MARKER}{}--></worksheet>"#,
        "x".repeat(16 * 1024)
    );
    let mut content_types = format!(
        r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="bin" ContentType="application/octet-stream"/><Override PartName="{MAIN}" ContentType="{}"/><Override PartName="{SHEET}" ContentType="{}"/><Override PartName="{SECOND}" ContentType="{}"/>"#,
        ct::SML_SHEET_MAIN,
        ct::SML_WORKSHEET,
        ct::SML_WORKSHEET,
    );
    if signed {
        content_types.push_str(&format!(
            r#"<Override PartName="/_xmlsignatures/origin.sigs" ContentType="{}"/>"#,
            ct::OPC_DIGITAL_SIGNATURE_ORIGIN
        ));
    }
    content_types.push_str("</Types>");

    let root_relationships = if signed {
        format!(
            r#"<Relationships xmlns="{PACKAGE_REL}"><Relationship Id="rIdOffice" Type="{office}" Target="xl/workbook.xml"/><Relationship Id="rIdSignature" Type="{signature}" Target="_xmlsignatures/origin.sigs"/></Relationships>"#,
            office = rt::OFFICE_DOCUMENT,
            signature = rt::DIGITAL_SIGNATURE_ORIGIN,
        )
    } else {
        format!(
            r#"<Relationships xmlns="{PACKAGE_REL}"><Relationship Id="rIdOffice" Type="{office}" Target="xl/workbook.xml"/></Relationships>"#,
            office = rt::OFFICE_DOCUMENT,
        )
    };
    let workbook_relationships = format!(
        r#"<Relationships xmlns="{PACKAGE_REL}"><Relationship Id="rIdSheet" Type="{worksheet}" Target="worksheets/sheet1.xml"/><Relationship Id="rIdSecond" Type="{worksheet}" Target="worksheets/sheet2.xml"/></Relationships>"#,
        worksheet = rt::WORKSHEET,
    );
    let sheet_relationships = format!(
        r#"<Relationships xmlns="{PACKAGE_REL}"><Relationship Id="rIdHyperlink" Type="{hyperlink}" Target="{EXTERNAL_TARGET}" TargetMode="External"/><Relationship Id="rIdDrawing" Type="{drawing}" Target="../media/opaque.bin"/></Relationships>"#,
        hyperlink = rt::HYPERLINK,
        drawing = rt::DRAWING,
    );

    let mut writer = soapberry_zip::office::StreamingArchiveWriter::new();
    writer
        .write_stored("[Content_Types].xml", content_types.as_bytes())
        .unwrap();
    writer
        .write_stored("_rels/.rels", root_relationships.as_bytes())
        .unwrap();
    writer
        .write_stored("xl/workbook.xml", workbook_xml.as_bytes())
        .unwrap();
    writer
        .write_stored(
            "xl/_rels/workbook.xml.rels",
            workbook_relationships.as_bytes(),
        )
        .unwrap();
    writer
        .write_stored("xl/worksheets/sheet1.xml", sheet_xml.as_bytes())
        .unwrap();
    writer
        .write_stored("xl/worksheets/sheet2.xml", second_xml.as_bytes())
        .unwrap();
    writer
        .write_stored(
            "xl/worksheets/_rels/sheet1.xml.rels",
            sheet_relationships.as_bytes(),
        )
        .unwrap();
    writer
        .write_stored(
            "xl/media/opaque.bin",
            &(0..8 * 1024)
                .map(|value| (value % 251) as u8)
                .collect::<Vec<_>>(),
        )
        .unwrap();
    if signed {
        writer
            .write_stored("_xmlsignatures/origin.sigs", b"<origin/>")
            .unwrap();
    }
    writer.finish_to_bytes().unwrap()
}

fn fixture(signed: bool) -> Vec<u8> {
    fixture_with_sheet(worksheet_xml(), signed)
}

fn eager_links(bytes: &[u8]) -> Vec<Hyperlink> {
    let workbook = Package::from_bytes(bytes.to_vec())
        .unwrap()
        .into_workbook()
        .unwrap();
    workbook
        .sheet("Sheet1")
        .unwrap()
        .unwrap()
        .hyperlinks()
        .unwrap()
}

fn eager_links_from_opc(package: OpcPackage) -> Vec<Hyperlink> {
    let workbook = Package::from_opc(package).unwrap().into_workbook().unwrap();
    workbook
        .sheet("Sheet1")
        .unwrap()
        .unwrap()
        .hyperlinks()
        .unwrap()
}

fn source_links(snapshot: &SourceBackedHyperlinkSnapshot) -> Vec<Hyperlink> {
    snapshot.hyperlinks().to_vec()
}

fn stage_change(edit: &mut SourceBackedHyperlinkEdit) {
    edit.put_hyperlink(internal("B2", "Sheet1!D4", Some("added"), None))
        .unwrap();
    edit.replace_hyperlink(internal("A1", "Sheet1!C3", Some("replaced"), Some("tip")))
        .unwrap();
    edit.put_hyperlink(external(
        "C3",
        EXTERNAL_TARGET,
        Some("updated remote"),
        Some("updated tip"),
    ))
    .unwrap();
    assert!(
        edit.remove_hyperlink(HyperlinkReference::new("B2").unwrap())
            .unwrap()
            .is_some()
    );
}

fn changed_commit(editor: &SourceBackedHyperlinkEditor) -> SourceBackedHyperlinkCommit {
    let mut edit = editor.edit("Sheet1").unwrap();
    stage_change(&mut edit);
    let commit = edit.commit().unwrap();
    assert!(commit.changed());
    commit
}

fn eager_changed_links(bytes: &[u8]) -> Vec<Hyperlink> {
    let workbook = Package::from_bytes(bytes.to_vec())
        .unwrap()
        .into_workbook()
        .unwrap();
    let mut edit = workbook.edit().unwrap();
    {
        let mut sheet = edit.sheet("Sheet1").unwrap().unwrap();
        sheet
            .put_hyperlink(internal("B2", "Sheet1!D4", Some("added"), None))
            .unwrap();
        sheet
            .replace_hyperlink(internal("A1", "Sheet1!C3", Some("replaced"), Some("tip")))
            .unwrap();
        sheet
            .replace_hyperlink(external(
                "C3",
                EXTERNAL_TARGET,
                Some("updated remote"),
                Some("updated tip"),
            ))
            .unwrap();
        assert!(
            sheet
                .remove_hyperlink(HyperlinkReference::new("B2").unwrap())
                .unwrap()
                .is_some()
        );
    }
    let committed = edit.commit().unwrap().into_workbook();
    committed
        .sheet("Sheet1")
        .unwrap()
        .unwrap()
        .hyperlinks()
        .unwrap()
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
        "xlsx-source-hyperlink-managed-test",
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
fn internal_edits_and_external_metadata_match_eager_semantics_and_stay_local() {
    let source_bytes = fixture(false);
    let eager_before = eager_links(&source_bytes);
    let source = Arc::new(CountingSource::new(source_bytes.clone()));
    let editor = SourceBackedHyperlinkEditor::from_read_at(source.clone()).unwrap();
    let before = editor.snapshot("Sheet1").unwrap();
    assert_eq!(source_links(&before), eager_before);
    assert_eq!(
        source.second_body_marker_reads.load(Ordering::SeqCst),
        0,
        "selected hyperlink snapshot must not materialize the unselected worksheet"
    );

    let commit = changed_commit(&editor);
    assert_eq!(
        source_links(commit.snapshot()),
        eager_changed_links(&source_bytes)
    );
    assert_eq!(
        source.second_body_marker_reads.load(Ordering::SeqCst),
        0,
        "source-backed hyperlink commit must remain selected-sheet local"
    );

    let mut output = Vec::new();
    editor
        .publish_commit_to_stream(&mut output, &commit)
        .unwrap();
    assert_eq!(eager_links(&output), source_links(commit.snapshot()));
}

#[test]
fn exact_noop_including_signed_source_bytes_and_changed_signed_source_refusal() {
    for signed in [false, true] {
        let source_bytes = fixture(signed);
        let editor = SourceBackedHyperlinkEditor::from_read_at(Arc::new(CountingSource::new(
            source_bytes.clone(),
        )))
        .unwrap();
        let noop = editor.edit("Sheet1").unwrap().commit().unwrap();
        assert!(!noop.changed());
        assert!(noop.patch().is_empty());
        let mut output = Vec::new();
        editor.publish_commit_to_stream(&mut output, &noop).unwrap();
        assert_eq!(output, source_bytes);
    }

    let source_bytes = fixture(true);
    let editor =
        SourceBackedHyperlinkEditor::from_read_at(Arc::new(CountingSource::new(source_bytes)))
            .unwrap();
    let mut edit = editor.edit("Sheet1").unwrap();
    edit.put_hyperlink(internal("D4", "Sheet1!D4", None, None))
        .unwrap();
    let commit = edit.commit().unwrap();
    let mut output = Vec::new();
    assert!(matches!(
        editor.publish_commit_to_stream(&mut output, &commit),
        Err(Error::Package(OpcError::SignedSourceRequiresExplicitPolicy))
    ));
    assert!(output.is_empty());
}

#[test]
fn patch_forward_inverse_restore_exact_worksheet_and_relationships_and_reject_stale_foreign() {
    let source_bytes = fixture(false);
    let editor = SourceBackedHyperlinkEditor::from_read_at(Arc::new(CountingSource::new(
        source_bytes.clone(),
    )))
    .unwrap();
    let commit = changed_commit(&editor);
    let patch: &SourceBackedHyperlinkPatch = commit.patch();

    let mut replay = OpcPackage::from_bytes(&source_bytes).unwrap();
    let sheet_uri = PackURI::new(SHEET).unwrap();
    let before_sheet = replay.get_part(&sheet_uri).unwrap().blob().to_vec();
    let before_relationships = relationship_signatures(replay.get_part(&sheet_uri).unwrap().rels());
    patch.apply(&mut replay).unwrap();
    assert_eq!(
        eager_links_from_opc(replay.clone()),
        source_links(commit.snapshot())
    );
    patch.inverse().apply(&mut replay).unwrap();
    assert_eq!(replay.get_part(&sheet_uri).unwrap().blob(), before_sheet);
    assert_eq!(
        relationship_signatures(replay.get_part(&sheet_uri).unwrap().rels()),
        before_relationships
    );

    let source = Arc::new(CountingSource::new(source_bytes.clone()));
    let editor = SourceBackedHyperlinkEditor::from_read_at(source.clone()).unwrap();
    let commit = changed_commit(&editor);
    source.changed();
    let mut output = Vec::new();
    assert!(matches!(
        editor.publish_commit_to_stream(&mut output, &commit),
        Err(Error::Package(OpcError::SourceChanged { .. }))
    ));
    assert!(output.is_empty());

    let editor =
        SourceBackedHyperlinkEditor::from_read_at(Arc::new(CountingSource::new(source_bytes)))
            .unwrap();
    let commit = changed_commit(&editor);
    let foreign = fixture_with_sheet(
        format!(
            r#"<worksheet xmlns="{SML}"><sheetData/><hyperlinks/><extLst><ext uri="urn:foreign"/></extLst></worksheet>"#
        ),
        false,
    );
    let foreign_editor =
        SourceBackedHyperlinkEditor::from_read_at(Arc::new(CountingSource::new(foreign))).unwrap();
    let mut output = Vec::new();
    assert!(matches!(
        foreign_editor.publish_commit_to_stream(&mut output, &commit),
        Err(Error::PatchConflict { .. })
    ));
    assert!(output.is_empty());
}

#[test]
fn changed_publication_raw_preserves_relationships_and_unselected_members() {
    let source_bytes = fixture(false);
    let source = Arc::new(CountingSource::new(source_bytes.clone()));
    let editor = SourceBackedHyperlinkEditor::from_read_at(source.clone()).unwrap();
    let commit = changed_commit(&editor);
    assert_eq!(source.second_body_marker_reads.load(Ordering::SeqCst), 0);

    let mut output = Vec::new();
    editor
        .publish_commit_to_stream(&mut output, &commit)
        .unwrap();
    let source_package = OpcPackage::from_bytes(&source_bytes).unwrap();
    let output_package = OpcPackage::from_bytes(&output).unwrap();
    assert_eq!(source_package.part_count(), output_package.part_count());
    for source_part in source_package.iter_parts() {
        let output_part = output_package.get_part(source_part.partname()).unwrap();
        assert_eq!(output_part.content_type(), source_part.content_type());
        assert_eq!(
            relationship_signatures(output_part.rels()),
            relationship_signatures(source_part.rels())
        );
        if source_part.partname().as_str() != SHEET {
            assert_eq!(output_part.blob(), source_part.blob());
        }
    }

    let source_raw = raw_members(&source_bytes);
    let output_raw = raw_members(&output);
    assert_eq!(
        source_raw.keys().collect::<Vec<_>>(),
        output_raw.keys().collect::<Vec<_>>()
    );
    for (name, source_member) in source_raw {
        if name == "xl/worksheets/sheet1.xml" {
            assert_ne!(output_raw[&name].local, source_member.local);
        } else {
            assert_eq!(
                output_raw.get(&name),
                Some(&source_member),
                "raw member {name}"
            );
        }
    }
}

#[test]
fn relationship_changes_and_opaque_owners_refuse_before_output() {
    let editor =
        SourceBackedHyperlinkEditor::from_read_at(Arc::new(CountingSource::new(fixture(false))))
            .unwrap();
    let mut edit = editor.edit("Sheet1").unwrap();
    let baseline = edit.values().unwrap();
    edit.remove_hyperlink(HyperlinkReference::new("C3").unwrap())
        .unwrap();
    assert!(edit.set(baseline).unwrap());
    assert!(!edit.commit().unwrap().changed());

    let editor =
        SourceBackedHyperlinkEditor::from_read_at(Arc::new(CountingSource::new(fixture(false))))
            .unwrap();
    let mut edit = editor.edit("Sheet1").unwrap();
    edit.replace_at(
        HyperlinkReference::new("C3").unwrap(),
        external("D4", EXTERNAL_TARGET, Some("moved"), None),
    )
    .unwrap();
    let staged = edit.values().unwrap();
    assert!(!edit.set(staged).unwrap());
    assert!(edit.commit().unwrap().changed());

    let editor =
        SourceBackedHyperlinkEditor::from_read_at(Arc::new(CountingSource::new(fixture(false))))
            .unwrap();
    let mut edit = editor.edit("Sheet1").unwrap();
    edit.put_hyperlink(external("D4", NEW_EXTERNAL_TARGET, None, None))
        .unwrap();
    assert!(matches!(edit.commit(), Err(Error::Unsupported { .. })));

    let editor =
        SourceBackedHyperlinkEditor::from_read_at(Arc::new(CountingSource::new(fixture(false))))
            .unwrap();
    let mut edit = editor.edit("Sheet1").unwrap();
    edit.replace_hyperlink(external(
        "C3",
        NEW_EXTERNAL_TARGET,
        Some("retargeted"),
        None,
    ))
    .unwrap();
    assert!(matches!(edit.commit(), Err(Error::Unsupported { .. })));

    let editor =
        SourceBackedHyperlinkEditor::from_read_at(Arc::new(CountingSource::new(fixture(false))))
            .unwrap();
    let mut edit = editor.edit("Sheet1").unwrap();
    assert!(
        edit.remove_hyperlink(HyperlinkReference::new("C3").unwrap())
            .unwrap()
            .is_some()
    );
    assert!(matches!(edit.commit(), Err(Error::Unsupported { .. })));

    let opaque = fixture_with_sheet(
        format!(
            r#"<worksheet xmlns="{SML}"><sheetData/><hyperlinks><hyperlink ref="A1" location="Sheet1!B2"><vendor/></hyperlink></hyperlinks></worksheet>"#
        ),
        false,
    );
    let editor =
        SourceBackedHyperlinkEditor::from_read_at(Arc::new(CountingSource::new(opaque))).unwrap();
    assert!(matches!(editor.edit("Sheet1"), Err(Error::Invalid(_))));

    let mce = fixture_with_sheet(
        format!(
            r#"<worksheet xmlns="{SML}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:u="urn:future"><sheetData/><mc:AlternateContent><mc:Choice Requires="u"><u:hyperlinks/></mc:Choice><mc:Fallback><hyperlinks/></mc:Fallback></mc:AlternateContent></worksheet>"#
        ),
        false,
    );
    let editor =
        SourceBackedHyperlinkEditor::from_read_at(Arc::new(CountingSource::new(mce))).unwrap();
    assert!(editor.edit("Sheet1").is_err());
}

#[test]
fn managed_limits_cancellation_and_partial_sink_are_typed() {
    let source_bytes = fixture(false);
    let exact = part_len(&source_bytes, MAIN) + part_len(&source_bytes, SHEET);
    let (budget, _cancellation_source, context) = managed_context(exact - 1);
    let editor = SourceBackedHyperlinkEditor::from_read_at_with_execution_context(
        Arc::new(CountingSource::new(source_bytes.clone())),
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
    drop(editor);
    assert_eq!(budget.used(Resource::Memory), 0);

    let limits = ReadLimits::builder()
        .max_part_bytes(1)
        .unwrap()
        .build()
        .unwrap();
    assert!(matches!(
        SourceBackedHyperlinkEditor::from_read_at_with_limits(
            Arc::new(CountingSource::new(source_bytes.clone())),
            limits,
        ),
        Err(Error::Package(OpcError::ReadLimit { .. }))
    ));

    let (budget, cancellation_source, context) = managed_context(exact);
    let editor = SourceBackedHyperlinkEditor::from_read_at_with_execution_context(
        Arc::new(CountingSource::new(source_bytes.clone())),
        ReadLimits::default(),
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

    let editor =
        SourceBackedHyperlinkEditor::from_read_at(Arc::new(CountingSource::new(source_bytes)))
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
