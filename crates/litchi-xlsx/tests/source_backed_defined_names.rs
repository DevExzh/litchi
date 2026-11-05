use std::collections::BTreeMap;
use std::io;
use std::io::Write;
use std::num::{NonZeroU64, NonZeroUsize};
use std::sync::Arc;
use std::sync::atomic::{AtomicU64, Ordering};

use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionError, ExecutionLimits,
    Limits as BudgetLimits, ReadAt, Resource, SourceVersion,
};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{
    BlobPart, OpcError, OpcPackage, PackURI, PackageWriter, SourceCacheLimits, TargetMode,
};
use litchi_xlsx::SourceBackedWorkbook;
use litchi_xlsx::defined_names::Snapshot;
use litchi_xlsx::defined_names::source::Commit;
use litchi_xlsx::raw::DefinedName;
use litchi_xlsx::{Error, SourceLimits};
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
    let index = PreservationIndex::new(
        &archive,
        &mut buffer,
        soapberry_zip::office::ArchiveLimits::default(),
    )
    .unwrap();
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

fn chained_names() -> Vec<DefinedName> {
    let mut names = target_names();
    names.push(DefinedName {
        name: "ChainedRange".to_owned(),
        reference: "'Sheet1'!$D$4".to_owned(),
        ..DefinedName::default()
    });
    names
}

fn changed_edit(editor: &SourceBackedWorkbook) -> litchi_xlsx::defined_names::source::Edit {
    let mut edit = editor.edit_defined_names().unwrap();
    assert!(edit.replace(target_names()).unwrap());
    edit
}

fn changed_commit(editor: &SourceBackedWorkbook) -> Commit {
    let commit = changed_edit(editor).commit().unwrap();
    assert!(commit.diagnostics().changed());
    assert_eq!(commit.patch().after().defined_names(), target_names());
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
        "xlsx-defined-names-managed-test",
        BudgetLimits::new(memory, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX),
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
fn changed_edit_reopens_and_changes_only_workbook_xml() {
    let source_bytes = fixture("", r#"<extLst><ext uri="urn:keep"/></extLst>"#, false);
    let editor =
        SourceBackedWorkbook::from_read_at(Arc::new(VersionedSource::new(source_bytes.clone())))
            .unwrap();
    let opened = editor.cache_diagnostics().successful_loads;
    assert_eq!(editor.cache_diagnostics().successful_loads, opened);
    let commit = changed_commit(&editor);
    assert_eq!(editor.cache_diagnostics().successful_loads, opened);

    let mut replay = OpcPackage::from_bytes(&source_bytes, SourceLimits::default()).unwrap();
    replay_patch(&mut replay, commit.patch()).unwrap();
    assert_eq!(
        Snapshot::load(&replay).unwrap().defined_names(),
        target_names()
    );
    replay_patch(&mut replay, &commit.patch().inverse()).unwrap();
    assert!(Snapshot::load(&replay).unwrap().defined_names().is_empty());

    let mut output = Vec::new();
    commit.package().write_to(&mut output).unwrap();
    assert_eq!(commit.package().defined_names(), target_names());
    assert_eq!(commit.package().active_sheet_position(), Some(0));
    let published = commit.patch().after().clone();
    assert_eq!(published.defined_names(), target_names());

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
        SourceBackedWorkbook::from_read_at(Arc::new(VersionedSource::new(bytes.clone()))).unwrap();
    let noop = editor.edit_defined_names().unwrap().commit().unwrap();
    assert!(!noop.diagnostics().changed());
    assert!(noop.patch().is_noop());
    let mut output = Vec::new();
    noop.package().write_to(&mut output).unwrap();
    assert_eq!(output, bytes);

    let editor = SourceBackedWorkbook::from_read_at(Arc::new(VersionedSource::new(bytes))).unwrap();
    let commit = changed_commit(&editor);
    let mut output = Vec::new();
    commit.package().write_to(&mut output).unwrap();
    let editor =
        SourceBackedWorkbook::from_read_at(Arc::new(VersionedSource::new(output))).unwrap();
    let mut edit = editor.edit_defined_names().unwrap();
    assert!(edit.clear().unwrap());
    let cleared = edit.commit().unwrap();
    assert!(cleared.patch().after().defined_names().is_empty());
}

#[test]
fn chained_edit_and_replay_resolve_the_deferred_defined_name_catalog() {
    let bytes = fixture("", "", false);
    let editor = SourceBackedWorkbook::from_read_at(Arc::new(VersionedSource::new(bytes))).unwrap();
    let first = changed_commit(&editor);

    // The first commit is still deferred. Its next edit must prove against
    // the committed target catalog, not the original owner's empty catalog.
    let mut chained = first.package().edit_defined_names().unwrap();
    assert_eq!(chained.defined_names().unwrap(), target_names().as_slice());
    let expected = chained_names();
    assert!(chained.replace(expected.clone()).unwrap());
    let second = chained.commit().unwrap();
    assert_eq!(second.patch().after().defined_names(), expected.as_slice());

    // Replay likewise creates a deferred output. Starting a second edit from
    // that replayed owner must observe the same target catalog and publish it.
    let replayed = editor
        .apply_defined_names(first.patch())
        .unwrap()
        .into_package();
    let mut replay_edit = replayed.edit_defined_names().unwrap();
    assert_eq!(
        replay_edit.defined_names().unwrap(),
        target_names().as_slice()
    );
    assert!(replay_edit.replace(expected.clone()).unwrap());
    let replayed_second = replay_edit.commit().unwrap();
    assert_eq!(
        replayed_second.patch().after().defined_names(),
        expected.as_slice()
    );
}

#[test]
fn signed_change_foreign_commit_and_changed_source_refuse_before_output() {
    let signed = fixture("", "", true);
    let editor =
        SourceBackedWorkbook::from_read_at(Arc::new(VersionedSource::new(signed.clone()))).unwrap();
    let noop = editor.edit_defined_names().unwrap().commit().unwrap();
    let mut output = Vec::new();
    noop.package().write_to(&mut output).unwrap();
    assert_eq!(output, signed);

    let editor =
        SourceBackedWorkbook::from_read_at(Arc::new(VersionedSource::new(signed))).unwrap();
    assert!(matches!(
        changed_edit(&editor).commit(),
        Err(Error::Package(OpcError::SignedSourceRequiresExplicitPolicy))
    ));

    let bytes = fixture("", "", false);
    let editor =
        SourceBackedWorkbook::from_read_at(Arc::new(VersionedSource::new(bytes.clone()))).unwrap();
    let commit = changed_commit(&editor);
    let foreign = fixture("", r#"<extLst><ext uri="urn:foreign"/></extLst>"#, false);
    let foreign_editor =
        SourceBackedWorkbook::from_read_at(Arc::new(VersionedSource::new(foreign))).unwrap();
    assert!(matches!(
        foreign_editor
            .apply_defined_names(commit.patch())
            .map(|_| ()),
        Err(Error::PatchConflict { .. })
    ));

    let source = Arc::new(VersionedSource::new(bytes));
    let editor = SourceBackedWorkbook::from_read_at(source.clone()).unwrap();
    let commit = changed_commit(&editor);
    source.changed();
    output.clear();
    assert!(matches!(
        commit.package().write_to(&mut output),
        Err(Error::Package(OpcError::SourceChanged { .. }))
    ));
    assert!(output.is_empty());
}

#[test]
fn protection_mce_limits_invalid_scopes_and_partial_sink_are_checked() {
    let protected = fixture(r#"<workbookProtection lockStructure="1"/>"#, "", false);
    let editor =
        SourceBackedWorkbook::from_read_at(Arc::new(VersionedSource::new(protected))).unwrap();
    assert!(
        editor
            .edit_defined_names()
            .unwrap()
            .replace(target_names())
            .is_err()
    );

    let mce = fixture(
        "",
        r#"<definedNames><definedName name="Direct">1</definedName><future/></definedNames>"#,
        false,
    );
    let editor = SourceBackedWorkbook::from_read_at(Arc::new(VersionedSource::new(mce))).unwrap();
    assert!(
        editor
            .edit_defined_names()
            .unwrap()
            .replace(target_names())
            .is_err()
    );

    let bytes = fixture("", "", false);
    let editor =
        SourceBackedWorkbook::from_read_at(Arc::new(VersionedSource::new(bytes.clone()))).unwrap();
    let mut invalid = target_names();
    invalid[1].local_sheet_id = Some(1);
    assert!(
        editor
            .edit_defined_names()
            .unwrap()
            .replace(invalid)
            .is_err()
    );

    let limits = SourceLimits::builder()
        .max_part_bytes(1)
        .unwrap()
        .build()
        .unwrap();
    assert!(matches!(
        SourceBackedWorkbook::from_read_at_with_options(
            Arc::new(VersionedSource::new(bytes.clone())),
            litchi_xlsx::OpenOptions::new().source_limits(limits)
        ),
        Err(Error::Package(OpcError::Limit { .. }))
    ));

    let editor = SourceBackedWorkbook::from_read_at(Arc::new(VersionedSource::new(bytes))).unwrap();
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
fn replay_rechecks_target_catalog_under_recipient_source_limits() {
    // Source-backed extraction reads in bounded chunks. Keep the admitted
    // source above that scratch size so an exact `max_part_bytes` policy can
    // admit the before workbook while the larger authored target exceeds it.
    let padding = format!("<!--{}-->", "x".repeat(192 * 1024));
    let source_bytes = fixture("", &padding, false);
    let producer =
        SourceBackedWorkbook::from_read_at(Arc::new(VersionedSource::new(source_bytes.clone())))
            .unwrap();
    let commit = changed_commit(&producer);
    let before_workbook_bytes = part_len(&source_bytes, MAIN);
    let recipient_limits = SourceLimits::builder()
        .max_part_bytes(before_workbook_bytes)
        .unwrap()
        .build()
        .unwrap();
    let recipient = SourceBackedWorkbook::from_read_at_with_options(
        Arc::new(VersionedSource::new(source_bytes)),
        litchi_xlsx::OpenOptions::new().source_limits(recipient_limits),
    )
    .unwrap();
    let mut output = Vec::new();
    commit.package().write_to(&mut output).unwrap();
    assert!(part_len(&output, MAIN) > before_workbook_bytes);

    assert!(matches!(
        recipient.apply_defined_names(commit.patch()),
        Err(Error::Package(OpcError::Limit { .. }))
    ));
}

#[test]
fn managed_workbook_payload_is_budgeted_streamed_and_released() {
    let bytes = fixture("", "", false);
    let exact = part_len(&bytes, MAIN);
    let (budget, _cancellation_source, context) = managed_context(exact);
    let editor = SourceBackedWorkbook::from_read_at_with_options(
        Arc::new(VersionedSource::new(bytes.clone())),
        litchi_xlsx::OpenOptions::new()
            .source_limits(SourceLimits::default())
            .cache_limits(SourceCacheLimits::new(usize::try_from(exact).unwrap(), 4).unwrap())
            .execution_context(context),
    )
    .unwrap();
    assert!(editor.cache_diagnostics().budget_managed);
    assert_eq!(budget.used(Resource::Memory), exact);

    let commit = changed_commit(&editor);
    let mut replay = OpcPackage::from_bytes(&bytes, SourceLimits::default()).unwrap();
    replay_patch(&mut replay, commit.patch()).unwrap();
    replay_patch(&mut replay, &commit.patch().inverse()).unwrap();

    let mut output = Vec::new();
    commit.package().write_to(&mut output).unwrap();
    let published = commit.patch().after().clone();
    assert_eq!(published.defined_names(), target_names());
    assert_eq!(
        Snapshot::load(&OpcPackage::from_bytes(&output, SourceLimits::default()).unwrap())
            .unwrap()
            .defined_names(),
        target_names()
    );
    drop(commit);
    assert!(budget.used(Resource::Memory) <= exact);
    drop(editor);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn managed_one_under_budget_fails_before_payload_retention_and_cancellation_stops_output() {
    let bytes = fixture("", "", false);
    let exact = part_len(&bytes, MAIN);
    let (budget, _cancellation_source, context) = managed_context(exact - 1);
    assert!(matches!(
        SourceBackedWorkbook::from_read_at_with_options(
            Arc::new(VersionedSource::new(bytes.clone())),
            litchi_xlsx::OpenOptions::new()
                .source_limits(SourceLimits::default())
                .execution_context(context)
        ),
        Err(Error::Package(OpcError::Execution(
            ExecutionError::ResourceLimit(_)
        )))
    ));
    assert_eq!(budget.used(Resource::Memory), 0);

    let (budget, cancellation_source, context) = managed_context(exact);
    let editor = SourceBackedWorkbook::from_read_at_with_options(
        Arc::new(VersionedSource::new(bytes.clone())),
        litchi_xlsx::OpenOptions::new()
            .source_limits(SourceLimits::default())
            .execution_context(context),
    )
    .unwrap();
    let edit = editor.edit_defined_names().unwrap();
    cancellation_source.cancel();
    assert!(matches!(
        edit.commit(),
        Err(Error::Package(OpcError::Cancelled))
    ));
    drop(editor);
    assert_eq!(budget.used(Resource::Memory), 0);

    let (budget, cancellation_source, context) = managed_context(exact);
    let editor = SourceBackedWorkbook::from_read_at_with_options(
        Arc::new(VersionedSource::new(bytes)),
        litchi_xlsx::OpenOptions::new()
            .source_limits(SourceLimits::default())
            .execution_context(context),
    )
    .unwrap();
    let commit = changed_commit(&editor);
    cancellation_source.cancel();
    let mut output = Vec::new();
    assert!(matches!(
        commit.package().write_to(&mut output),
        Err(Error::Package(OpcError::Cancelled))
    ));
    assert!(output.is_empty());
    drop(commit);
    assert!(budget.used(Resource::Memory) <= exact);
    drop(editor);
    assert_eq!(budget.used(Resource::Memory), 0);
}

/// Replay a source-backed patch onto `package` through a source-backed
/// workbook over its bytes (record 0360: source-backed patches replay onto
/// source-backed workbooks).
fn replay_patch(
    package: &mut OpcPackage,
    patch: &litchi_xlsx::defined_names::source::Patch,
) -> litchi_xlsx::Result<()> {
    let bytes = PackageWriter::to_bytes(package)?;
    let workbook =
        SourceBackedWorkbook::from_read_at(Arc::new(litchi_core::OwnedSource::new(bytes)))?;
    let next = workbook
        .apply_defined_names(patch)?
        .into_package()
        .to_bytes()?;
    *package = OpcPackage::from_bytes(&next, SourceLimits::default())?;
    Ok(())
}
