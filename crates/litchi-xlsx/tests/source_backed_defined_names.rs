use std::collections::BTreeMap;
use std::io;
use std::io::Write;
use std::sync::Arc;
use std::sync::atomic::{AtomicU64, Ordering};

use litchi_core::{ReadAt, SourceVersion};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{BlobPart, OpcError, OpcPackage, PackURI, PackageWriter, TargetMode};
use litchi_xlsx::defined_names::{Commit, Snapshot, SourceBackedEditor};
use litchi_xlsx::raw::DefinedName;
use litchi_xlsx::{Error, ReadLimits};
use soapberry_zip::{PreservationIndex, ZipArchive};

const SML: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const MAIN: &str = "/xl/workbook.xml";
const SHEET: &str = "/xl/worksheets/sheet1.xml";
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
        Ok(SourceVersion::new(31, self.revision.load(Ordering::SeqCst)))
    }
}

struct FailingSink {
    accepted: usize,
    limit: usize,
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

fn fixture(workbook_prefix: &str, workbook_suffix: &str, signed: bool) -> Vec<u8> {
    let workbook = format!(
        r#"<workbook xmlns="{SML}" xmlns:r="{REL}">{workbook_prefix}<sheets><sheet name="Sheet1" sheetId="1" r:id="rIdSheet"/></sheets><calcPr calcId="7"/>{workbook_suffix}</workbook>"#
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
            format!(r#"<worksheet xmlns="{SML}"><sheetData/></worksheet>"#).into_bytes(),
        )))
        .unwrap();
    package
        .try_add_part(Box::new(BlobPart::new(
            PackURI::new(UNUSED).unwrap(),
            "application/octet-stream".to_string(),
            (0..128 * 1024).map(|value| (value % 251) as u8).collect(),
        )))
        .unwrap();
    package
        .get_part_mut(&PackURI::new(MAIN).unwrap())
        .unwrap()
        .rels_mut()
        .try_add_relationship(
            rt::WORKSHEET.to_string(),
            "worksheets/sheet1.xml".to_string(),
            "rIdSheet".to_string(),
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

fn target_names() -> Vec<DefinedName> {
    vec![
        DefinedName {
            name: "GlobalRange".to_owned(),
            reference: "'Sheet1'!$A$1:$B$2".to_owned(),
            comment: Some("kept & escaped".to_owned()),
            ..DefinedName::default()
        },
        DefinedName {
            name: "LocalCell".to_owned(),
            reference: "'Sheet1'!$C$3".to_owned(),
            local_sheet_id: Some(0),
            hidden: true,
            ..DefinedName::default()
        },
    ]
}

fn changed_commit(editor: &SourceBackedEditor) -> Commit {
    let mut edit = editor.edit();
    assert!(edit.replace(target_names()).unwrap());
    let commit = edit.commit().unwrap();
    assert!(commit.changed());
    assert_eq!(commit.diagnostics().touched_workbooks(), 1);
    assert_eq!(commit.snapshot().defined_names(), target_names());
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
fn changed_edit_reopens_and_changes_only_workbook_xml() {
    let source_bytes = fixture("", r#"<extLst><ext uri="urn:keep"/></extLst>"#, false);
    let editor =
        SourceBackedEditor::from_read_at(Arc::new(VersionedSource::new(source_bytes.clone())))
            .unwrap();
    assert_eq!(editor.cache_diagnostics().successful_loads, 1);
    let commit = changed_commit(&editor);
    assert_eq!(editor.cache_diagnostics().successful_loads, 1);

    let mut replay = OpcPackage::from_bytes(&source_bytes).unwrap();
    commit.patch().apply(&mut replay).unwrap();
    assert_eq!(
        Snapshot::load(&replay).unwrap().defined_names(),
        target_names()
    );
    commit.patch().inverse().apply(&mut replay).unwrap();
    assert!(Snapshot::load(&replay).unwrap().defined_names().is_empty());

    let mut output = Vec::new();
    let published = editor
        .publish_commit_to_stream(&mut output, &commit)
        .unwrap();
    assert_eq!(published.defined_names(), target_names());

    let source = OpcPackage::from_bytes(&source_bytes).unwrap();
    let candidate = OpcPackage::from_bytes(&output).unwrap();
    assert_eq!(source.part_count(), candidate.part_count());
    assert_eq!(
        relationship_signatures(source.rels()),
        relationship_signatures(candidate.rels())
    );
    for part in source.iter_parts() {
        let output_part = candidate.get_part(part.partname()).unwrap();
        assert_eq!(part.content_type(), output_part.content_type());
        assert_eq!(
            relationship_signatures(part.rels()),
            relationship_signatures(output_part.rels())
        );
        if part.partname().as_str() == MAIN {
            assert_ne!(part.blob(), output_part.blob());
            let output_xml = std::str::from_utf8(output_part.blob()).unwrap();
            assert!(output_xml.contains("GlobalRange"));
            assert!(output_xml.contains(r#"<extLst><ext uri="urn:keep"/></extLst>"#));
        } else {
            assert_eq!(part.blob(), output_part.blob());
        }
    }
    assert_eq!(
        Snapshot::load(&candidate).unwrap().defined_names(),
        target_names()
    );
    let source_raw = raw_members(&source_bytes);
    let candidate_raw = raw_members(&output);
    assert_eq!(
        source_raw.keys().collect::<Vec<_>>(),
        candidate_raw.keys().collect::<Vec<_>>()
    );
    for (name, source_member) in source_raw {
        if name != "xl/workbook.xml" {
            assert_eq!(
                candidate_raw.get(&name),
                Some(&source_member),
                "raw member {name}"
            );
        }
    }
}

#[test]
fn replacement_clear_and_noop_are_exact() {
    let bytes = fixture("", "", false);
    let editor =
        SourceBackedEditor::from_read_at(Arc::new(VersionedSource::new(bytes.clone()))).unwrap();
    let noop = editor.edit().commit().unwrap();
    assert!(!noop.changed());
    assert!(noop.patch().is_empty());
    let mut output = Vec::new();
    editor.publish_commit_to_stream(&mut output, &noop).unwrap();
    assert_eq!(output, bytes);

    let editor = SourceBackedEditor::from_read_at(Arc::new(VersionedSource::new(bytes))).unwrap();
    let commit = changed_commit(&editor);
    let mut output = Vec::new();
    editor
        .publish_commit_to_stream(&mut output, &commit)
        .unwrap();
    let editor = SourceBackedEditor::from_read_at(Arc::new(VersionedSource::new(output))).unwrap();
    let mut edit = editor.edit();
    assert!(edit.clear().unwrap());
    let cleared = edit.commit().unwrap();
    assert!(cleared.snapshot().defined_names().is_empty());
}

#[test]
fn signed_change_foreign_commit_and_changed_source_refuse_before_output() {
    let signed = fixture("", "", true);
    let editor =
        SourceBackedEditor::from_read_at(Arc::new(VersionedSource::new(signed.clone()))).unwrap();
    let noop = editor.edit().commit().unwrap();
    let mut output = Vec::new();
    editor.publish_commit_to_stream(&mut output, &noop).unwrap();
    assert_eq!(output, signed);

    let editor = SourceBackedEditor::from_read_at(Arc::new(VersionedSource::new(signed))).unwrap();
    let commit = changed_commit(&editor);
    output.clear();
    assert!(matches!(
        editor.publish_commit_to_stream(&mut output, &commit),
        Err(Error::Package(OpcError::SignedSourceRequiresExplicitPolicy))
    ));
    assert!(output.is_empty());

    let bytes = fixture("", "", false);
    let editor =
        SourceBackedEditor::from_read_at(Arc::new(VersionedSource::new(bytes.clone()))).unwrap();
    let commit = changed_commit(&editor);
    let foreign = fixture("", r#"<extLst><ext uri="urn:foreign"/></extLst>"#, false);
    let foreign_editor =
        SourceBackedEditor::from_read_at(Arc::new(VersionedSource::new(foreign))).unwrap();
    assert!(matches!(
        foreign_editor.publish_commit_to_stream(&mut output, &commit),
        Err(Error::PatchConflict { .. })
    ));
    assert!(output.is_empty());

    let source = Arc::new(VersionedSource::new(bytes));
    let editor = SourceBackedEditor::from_read_at(source.clone()).unwrap();
    let commit = changed_commit(&editor);
    source.changed();
    assert!(matches!(
        editor.publish_commit_to_stream(&mut output, &commit),
        Err(Error::Package(OpcError::SourceChanged { .. }))
    ));
    assert!(output.is_empty());
}

#[test]
fn protection_mce_limits_invalid_scopes_and_partial_sink_are_checked() {
    let protected = fixture(r#"<workbookProtection lockStructure="1"/>"#, "", false);
    let editor =
        SourceBackedEditor::from_read_at(Arc::new(VersionedSource::new(protected))).unwrap();
    assert!(editor.edit().replace(target_names()).is_err());

    let mce = fixture(
        "",
        r#"<definedNames><definedName name="Direct">1</definedName><future/></definedNames>"#,
        false,
    );
    let editor = SourceBackedEditor::from_read_at(Arc::new(VersionedSource::new(mce))).unwrap();
    assert!(editor.edit().replace(target_names()).is_err());

    let bytes = fixture("", "", false);
    let editor =
        SourceBackedEditor::from_read_at(Arc::new(VersionedSource::new(bytes.clone()))).unwrap();
    let mut invalid = target_names();
    invalid[1].local_sheet_id = Some(1);
    assert!(editor.edit().replace(invalid).is_err());

    let limits = ReadLimits::builder()
        .max_part_bytes(1)
        .unwrap()
        .build()
        .unwrap();
    assert!(matches!(
        SourceBackedEditor::from_read_at_with_limits(
            Arc::new(VersionedSource::new(bytes.clone())),
            limits,
        ),
        Err(Error::Package(OpcError::ReadLimit { .. }))
    ));

    let editor = SourceBackedEditor::from_read_at(Arc::new(VersionedSource::new(bytes))).unwrap();
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
