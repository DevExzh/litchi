#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "focused topology assertions intentionally fail on fixture errors"
)]

//! Integration coverage for the public source-backed prepared-topology seam.
//!
//! The tests inspect the logical candidate before publication, then reopen the
//! published physical package through the same source-backed API.  Unchanged
//! Parts are deliberately left out of the plan so their payload remains lazy.

use std::io::{self, Write};
use std::ops::Range;
use std::sync::{
    Arc,
    atomic::{AtomicBool, AtomicU64, Ordering},
};

use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionError, ExecutionLimits, Limits, ReadAt,
    Resource, SourceVersion,
};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{
    AuthoredXmlFragment, BlobPart, OpcError, OpcPackage, PackURI, PackageWriter,
    SourceBackedPackage, SourceTopologyPlan,
};
use quick_xml::{events::Event, reader::NsReader};
use soapberry_zip::office::StreamingArchiveWriter;

const DOCUMENT: &str = "/word/document.xml";
const UNTOUCHED: &str = "/custom/untouched.xml";
const ADDED: &str = "/custom/added.xml";
const LARGE_UNTOUCHED: &str = "/custom/large.bin";
const CUSTOM_REL: &str = "http://example.invalid/relationships/custom";
const DOCUMENT_REL_MEMBER: &str = "word/_rels/document.xml.rels";
const PACKAGE_REL_MEMBER: &str = "_rels/.rels";
const RELATIONSHIPS_NS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const CONTENT_TYPES_NS: &str = "http://schemas.openxmlformats.org/package/2006/content-types";
const OFFICE_DOCUMENT_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
const HYPERLINK_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";

fn pack(uri: &str) -> PackURI {
    PackURI::new(uri).unwrap()
}

#[derive(Debug)]
struct VersionedSource {
    bytes: Arc<Vec<u8>>,
    revision: AtomicU64,
    reads: AtomicU64,
    denied_start: AtomicU64,
    denied_end: AtomicU64,
    checks_before_change: AtomicU64,
    fail_and_change_on_read: AtomicBool,
}

impl VersionedSource {
    fn new(bytes: Vec<u8>) -> Self {
        Self {
            bytes: Arc::new(bytes),
            revision: AtomicU64::new(0),
            reads: AtomicU64::new(0),
            denied_start: AtomicU64::new(0),
            denied_end: AtomicU64::new(0),
            checks_before_change: AtomicU64::new(0),
            fail_and_change_on_read: AtomicBool::new(false),
        }
    }

    fn bump(&self) {
        self.revision.fetch_add(1, Ordering::SeqCst);
    }

    fn read_count(&self) -> u64 {
        self.reads.load(Ordering::SeqCst)
    }

    fn deny_reads_in(&self, range: Range<u64>) {
        self.denied_start.store(range.start, Ordering::SeqCst);
        self.denied_end.store(range.end, Ordering::SeqCst);
    }

    fn allow_all_reads(&self) {
        self.denied_end.store(0, Ordering::SeqCst);
    }
}

impl ReadAt for VersionedSource {
    fn len(&self) -> io::Result<u64> {
        u64::try_from(self.bytes.len()).map_err(|_| io::Error::other("fixture length exceeds u64"))
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        self.reads.fetch_add(1, Ordering::SeqCst);
        if self.fail_and_change_on_read.swap(false, Ordering::SeqCst) {
            self.bump();
            return Err(io::Error::other(
                "source changed during failed payload read",
            ));
        }
        let read_end = offset.saturating_add(output.len() as u64);
        let denied_start = self.denied_start.load(Ordering::SeqCst);
        let denied_end = self.denied_end.load(Ordering::SeqCst);
        if denied_end != 0 && offset < denied_end && denied_start < read_end {
            return Err(io::Error::other("test source denied member range"));
        }
        let offset = usize::try_from(offset)
            .map_err(|_| io::Error::new(io::ErrorKind::InvalidInput, "offset exceeds usize"))?;
        if offset >= self.bytes.len() {
            return Ok(0);
        }
        let count = output.len().min(self.bytes.len() - offset);
        output[..count].copy_from_slice(&self.bytes[offset..offset + count]);
        Ok(count)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        if self
            .checks_before_change
            .fetch_update(Ordering::SeqCst, Ordering::SeqCst, |n| n.checked_sub(1))
            == Ok(1)
        {
            self.bump();
        }
        Ok(SourceVersion::new(
            0x5052_4550,
            self.revision.load(Ordering::SeqCst),
        ))
    }
}

fn source_bytes() -> Vec<u8> {
    let mut package = OpcPackage::new();
    package
        .try_add_part(Box::new(BlobPart::new(
            pack(DOCUMENT),
            ct::WML_DOCUMENT.to_owned(),
            b"<before/>".to_vec(),
        )))
        .unwrap();
    package
        .try_add_part(Box::new(BlobPart::new(
            pack(UNTOUCHED),
            "application/xml".to_owned(),
            b"<untouched source='yes'/>".to_vec(),
        )))
        .unwrap();
    package.relate_to("word/document.xml", rt::OFFICE_DOCUMENT);
    package
        .get_part_mut(&pack(DOCUMENT))
        .unwrap()
        .relate_to("../custom/untouched.xml", CUSTOM_REL);
    PackageWriter::to_bytes(&package).unwrap()
}

fn stored_member_data_range(bytes: &[u8], wanted: &str) -> Range<u64> {
    let mut offset = 0usize;
    while offset.checked_add(30).is_some_and(|end| end <= bytes.len())
        && bytes[offset..offset + 4] == [0x50, 0x4b, 0x03, 0x04]
    {
        let name_len = u16::from_le_bytes([bytes[offset + 26], bytes[offset + 27]]) as usize;
        let extra_len = u16::from_le_bytes([bytes[offset + 28], bytes[offset + 29]]) as usize;
        let data_start = offset
            .checked_add(30)
            .and_then(|value| value.checked_add(name_len))
            .and_then(|value| value.checked_add(extra_len))
            .expect("test ZIP local-header range must fit usize");
        let data_len = u32::from_le_bytes([
            bytes[offset + 18],
            bytes[offset + 19],
            bytes[offset + 20],
            bytes[offset + 21],
        ]) as usize;
        let data_end = data_start
            .checked_add(data_len)
            .expect("test ZIP member range must fit usize");
        let name = std::str::from_utf8(&bytes[offset + 30..data_start])
            .expect("test ZIP member names are UTF-8");
        if name == wanted {
            return (data_start as u64)..(data_end as u64);
        }
        offset = data_end;
    }
    panic!("test ZIP member {wanted:?} is missing");
}

fn source_bytes_with_large_untouched_part() -> (Vec<u8>, Vec<u8>) {
    let large_payload = vec![b'L'; 128 * 1024];
    let mut package = OpcPackage::new();
    package
        .try_add_part(Box::new(BlobPart::new(
            pack(DOCUMENT),
            ct::WML_DOCUMENT.to_owned(),
            b"<before/>".to_vec(),
        )))
        .unwrap();
    package
        .try_add_part(Box::new(BlobPart::new(
            pack(UNTOUCHED),
            "application/xml".to_owned(),
            b"<untouched source='yes'/>".to_vec(),
        )))
        .unwrap();
    package
        .try_add_part(Box::new(BlobPart::new(
            pack(LARGE_UNTOUCHED),
            "application/octet-stream".to_owned(),
            large_payload.clone(),
        )))
        .unwrap();
    package.relate_to("word/document.xml", rt::OFFICE_DOCUMENT);
    package
        .get_part_mut(&pack(DOCUMENT))
        .unwrap()
        .relate_to("../custom/untouched.xml", CUSTOM_REL);
    (PackageWriter::to_bytes(&package).unwrap(), large_payload)
}

fn open(source: Arc<VersionedSource>) -> SourceBackedPackage {
    SourceBackedPackage::from_read_at(source).unwrap()
}

fn open_with_limits(
    source: Arc<VersionedSource>,
    limits: litchi_opc::ReadLimits,
) -> litchi_opc::Result<SourceBackedPackage> {
    SourceBackedPackage::from_read_at_with_limits(source, limits)
}

fn mutation_plan() -> SourceTopologyPlan {
    let mut plan = SourceTopologyPlan::new();
    plan.try_replace_part(pack(DOCUMENT), b"<after/>".to_vec())
        .unwrap();
    plan.try_add_part(pack(ADDED), "application/example+xml", b"<added/>".to_vec())
        .unwrap();
    plan.try_add_internal_relationship(pack("/"), "rIdAdded", CUSTOM_REL, pack(ADDED))
        .unwrap();
    plan.try_add_internal_relationship(pack(DOCUMENT), "rIdAdded", CUSTOM_REL, pack(ADDED))
        .unwrap();
    plan
}

fn zip_member(bytes: &[u8], name: &str) -> Vec<u8> {
    soapberry_zip::office::ArchiveReader::new(bytes)
        .unwrap()
        .read(name)
        .unwrap()
}

fn whitespace_rich_relationship_source() -> Vec<u8> {
    let content_types = format!(
        r#"<Types xmlns="{CONTENT_TYPES_NS}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/></Types>"#
    );
    let package_relationships = format!(
        r#"<Relationships xmlns="{RELATIONSHIPS_NS}"><Relationship Id="rId1" Type="{OFFICE_DOCUMENT_REL}" Target="word/document.xml"/></Relationships>"#
    );
    let document_relationships = br#"<?xml version="1.0" encoding="UTF-8"?>
<!--preserve-before-->
<r:Relationships producer="opaque" xmlns="urn:litchi:unused" xmlns:r="http://schemas.openxmlformats.org/package/2006/relationships">
  <?preserve instruction?>
  <r:Relationship TargetMode='External'
      Target='https://before.invalid/?a=1&amp;b=2#fragment'
      Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'
      Id='rExisting'></r:Relationship>
  <!--preserve-between-->
  <r:Relationship Id='rInternal' Target="./target.xml" Type="urn:litchi:test"><?remaining?><!----></r:Relationship>
  <?preserve-after?>
</r:Relationships>
<!--preserve-tail-->"#;

    let mut writer = StreamingArchiveWriter::new();
    writer
        .write_stored("[Content_Types].xml", content_types.as_bytes())
        .unwrap();
    writer
        .write_stored(PACKAGE_REL_MEMBER, package_relationships.as_bytes())
        .unwrap();
    writer
        .write_stored(DOCUMENT_REL_MEMBER, document_relationships)
        .unwrap();
    writer
        .write_deflated_sized("word/document.xml", b"<document/>")
        .unwrap();
    writer.finish_to_bytes().unwrap()
}

fn relationship_event_count(xml: &[u8]) -> usize {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(true);
    reader.config_mut().check_end_names = true;
    let mut events = 0usize;
    loop {
        events += 1;
        if matches!(reader.read_resolved_event().unwrap().1, Event::Eof) {
            return events;
        }
    }
}

fn relationship_limits(
    source: &[u8],
    expected: &[u8],
    member_limit: usize,
    total_limit: usize,
    event_limit: usize,
) -> litchi_opc::ReadLimits {
    let package_relationships = zip_member(source, PACKAGE_REL_MEMBER);
    let expected_relationships = zip_member(expected, DOCUMENT_REL_MEMBER);
    assert_eq!(expected_relationships.len(), member_limit);
    assert_eq!(
        package_relationships.len() + expected_relationships.len(),
        total_limit
    );
    assert_eq!(
        relationship_event_count(&package_relationships)
            + relationship_event_count(&expected_relationships),
        event_limit
    );
    litchi_opc::ReadLimits::builder()
        .max_relationship_xml_bytes(member_limit)
        .unwrap()
        .max_total_relationship_xml_bytes(total_limit)
        .unwrap()
        .max_total_relationship_xml_events(event_limit)
        .unwrap()
        .build()
        .unwrap()
}

fn append_external_plan() -> SourceTopologyPlan {
    let mut plan = SourceTopologyPlan::new();
    plan.try_add_external_relationship(
        pack(DOCUMENT),
        "rAdded",
        HYPERLINK_REL,
        "https://added.invalid/?q=two#fragment",
    )
    .unwrap();
    plan
}

fn remove_canonical_plan() -> SourceTopologyPlan {
    let mut plan = SourceTopologyPlan::new();
    plan.try_remove_relationship(pack(DOCUMENT), "rId1")
        .unwrap();
    plan
}

fn remove_noncanonical_plan() -> SourceTopologyPlan {
    let mut plan = SourceTopologyPlan::new();
    plan.try_remove_external_relationship(pack(DOCUMENT), "rExisting")
        .unwrap();
    plan
}

fn publish_relationship_plan(
    source: &[u8],
    limits: litchi_opc::ReadLimits,
    plan: SourceTopologyPlan,
) -> litchi_opc::Result<Vec<u8>> {
    let package = open_with_limits(Arc::new(VersionedSource::new(source.to_vec())), limits)?;
    let prepared = package.prepare_topology(plan)?;
    prepared.with_candidate(|candidate| {
        let _ = candidate.relationships(&pack(DOCUMENT))?;
        Ok(())
    })?;
    let mut output = Vec::new();
    prepared.publish_to_stream(&mut output)?;
    Ok(output)
}

fn assert_relationship_limit_boundaries(source: &[u8], plan: fn() -> SourceTopologyPlan) {
    let expected =
        publish_relationship_plan(source, litchi_opc::ReadLimits::default(), plan()).unwrap();
    let expected_member = zip_member(&expected, DOCUMENT_REL_MEMBER);
    let package_relationships = zip_member(source, PACKAGE_REL_MEMBER);
    let exact_member = expected_member.len();
    let exact_total = package_relationships.len() + exact_member;
    let exact_events = relationship_event_count(&package_relationships)
        + relationship_event_count(&expected_member);

    let exact_limits =
        relationship_limits(source, &expected, exact_member, exact_total, exact_events);
    let exact_output = publish_relationship_plan(source, exact_limits, plan()).unwrap();
    assert_eq!(
        zip_member(&exact_output, DOCUMENT_REL_MEMBER),
        expected_member
    );

    let cases = [
        (
            "per-member bytes",
            litchi_opc::ReadLimits::builder()
                .max_relationship_xml_bytes(exact_member - 1)
                .unwrap()
                .max_total_relationship_xml_bytes(exact_total)
                .unwrap()
                .max_total_relationship_xml_events(exact_events)
                .unwrap()
                .build()
                .unwrap(),
            litchi_opc::ReadResource::RelationshipXmlBytes,
        ),
        (
            "aggregate relationship bytes",
            litchi_opc::ReadLimits::builder()
                .max_relationship_xml_bytes(exact_member)
                .unwrap()
                .max_total_relationship_xml_bytes(exact_total - 1)
                .unwrap()
                .max_total_relationship_xml_events(exact_events)
                .unwrap()
                .build()
                .unwrap(),
            litchi_opc::ReadResource::TotalRelationshipXmlBytes,
        ),
        (
            "aggregate relationship events",
            litchi_opc::ReadLimits::builder()
                .max_relationship_xml_bytes(exact_member)
                .unwrap()
                .max_total_relationship_xml_bytes(exact_total)
                .unwrap()
                .max_total_relationship_xml_events(exact_events - 1)
                .unwrap()
                .build()
                .unwrap(),
            litchi_opc::ReadResource::TotalRelationshipXmlEvents,
        ),
    ];
    for (label, limits, resource) in cases {
        let package =
            open_with_limits(Arc::new(VersionedSource::new(source.to_vec())), limits).unwrap();
        let error = match package.prepare_topology(plan()) {
            Err(error) => error,
            Ok(_) => panic!("{label} must refuse one under the exact final boundary"),
        };
        assert!(
            matches!(error, OpcError::ReadLimit { resource: actual, .. } if actual == resource),
            "{label} returned unexpected error: {error:?}"
        );
    }
}

fn assert_relationship_removal_subtracts_source_usage(
    source: &[u8],
    plan: fn() -> SourceTopologyPlan,
) {
    let package_relationships = zip_member(source, PACKAGE_REL_MEMBER);
    let document_relationships = zip_member(source, DOCUMENT_REL_MEMBER);
    let source_member_limit = package_relationships
        .len()
        .max(document_relationships.len());
    let source_total_limit = package_relationships.len() + document_relationships.len();
    let source_event_limit = relationship_event_count(&package_relationships)
        + relationship_event_count(&document_relationships);
    let exact_limits = litchi_opc::ReadLimits::builder()
        .max_relationship_xml_bytes(source_member_limit)
        .unwrap()
        .max_total_relationship_xml_bytes(source_total_limit)
        .unwrap()
        .max_total_relationship_xml_events(source_event_limit)
        .unwrap()
        .build()
        .unwrap();

    // The source itself consumes the entire configured relationship budget.
    // A removal must subtract the removed member/relationship from the
    // effective candidate instead of charging the unchanged source total a
    // second time.
    let output = publish_relationship_plan(source, exact_limits, plan()).unwrap();
    assert!(zip_member(&output, DOCUMENT_REL_MEMBER).len() < document_relationships.len());

    // One-under caps are rejected at source admission, before a topology
    // candidate exists. This keeps the exact-cap success above distinct from
    // an ingress refusal caused by an actually-too-small source budget.
    let under_member = litchi_opc::ReadLimits::builder()
        .max_relationship_xml_bytes(source_member_limit - 1)
        .unwrap()
        .build()
        .unwrap();
    let member_error = match open_with_limits(
        Arc::new(VersionedSource::new(source.to_vec())),
        under_member,
    ) {
        Err(error) => error,
        Ok(_) => panic!("one-under source relationship byte cap must refuse ingress"),
    };
    assert!(matches!(
        member_error,
        OpcError::ReadLimit {
            resource: litchi_opc::ReadResource::RelationshipXmlBytes,
            ..
        }
    ));

    let under_total = litchi_opc::ReadLimits::builder()
        .max_total_relationship_xml_bytes(source_total_limit - 1)
        .unwrap()
        .build()
        .unwrap();
    let total_error =
        match open_with_limits(Arc::new(VersionedSource::new(source.to_vec())), under_total) {
            Err(error) => error,
            Ok(_) => panic!("one-under source aggregate relationship cap must refuse ingress"),
        };
    assert!(matches!(
        total_error,
        OpcError::ReadLimit {
            resource: litchi_opc::ReadResource::TotalRelationshipXmlBytes,
            ..
        }
    ));
}

fn managed_context() -> (Budget, CancellationSource, ExecutionContext) {
    managed_context_with_resources(16 * 1024 * 1024, u64::MAX, u64::MAX, u64::MAX, u64::MAX)
}

fn managed_context_with_resources(
    memory: u64,
    input_bytes: u64,
    output_bytes: u64,
    objects: u64,
    work: u64,
) -> (Budget, CancellationSource, ExecutionContext) {
    let budget = Budget::root(
        "prepared-topology-managed-test",
        Limits::new(memory, input_bytes, output_bytes, objects, u64::MAX, work),
    );
    let (cancellation_source, cancellation) = CancellationSource::pair();
    let execution_limits = ExecutionLimits::new(
        std::num::NonZeroUsize::new(1).unwrap(),
        std::num::NonZeroUsize::new(1).unwrap(),
        std::num::NonZeroU64::new(memory).unwrap(),
        0,
    )
    .unwrap();
    let context = ExecutionContext::new(budget.clone(), cancellation, execution_limits);
    (budget, cancellation_source, context)
}

struct PartialWriter {
    bytes: Vec<u8>,
    remaining: usize,
}

impl Write for PartialWriter {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        if self.remaining == 0 {
            return Err(io::Error::new(io::ErrorKind::BrokenPipe, "test sink"));
        }
        let accepted = self.remaining.min(bytes.len());
        self.bytes.extend_from_slice(&bytes[..accepted]);
        self.remaining -= accepted;
        Ok(accepted)
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

#[test]
fn prepare_effective_graph_publish_and_reopen_are_consistent() {
    let source = source_bytes();
    let source_handle = Arc::new(VersionedSource::new(source.clone()));
    let reads_before_prepare = source_handle.read_count();
    let prepared = open(Arc::clone(&source_handle))
        .prepare_topology(mutation_plan())
        .unwrap();
    let reads_before_candidate = source_handle.read_count();
    assert!(reads_before_candidate >= reads_before_prepare);

    let observed = prepared
        .with_candidate(|candidate| {
            let document = candidate.part(&pack(DOCUMENT))?;
            assert_eq!(document.data()?.as_bytes(), b"<after/>");

            let added = candidate.part(&pack(ADDED))?;
            assert_eq!(added.content_type().as_str(), "application/example+xml");
            assert_eq!(added.data()?.as_bytes(), b"<added/>");

            let package_relationship = candidate
                .package_relationships()
                .get("rIdAdded")
                .ok_or_else(|| OpcError::RelationshipNotFound("rIdAdded".to_owned()))?;
            assert_eq!(package_relationship.reltype(), CUSTOM_REL);
            assert_eq!(package_relationship.target_ref(), "custom/added.xml");

            let document_relationship =
                candidate
                    .relationships(&pack(DOCUMENT))?
                    .get("rIdAdded")
                    .ok_or_else(|| OpcError::RelationshipNotFound("rIdAdded".to_owned()))?;
            assert_eq!(document_relationship.reltype(), CUSTOM_REL);
            assert_eq!(document_relationship.target_ref(), "../custom/added.xml");
            assert_eq!(
                candidate.content_type(&pack(ADDED))?.as_str(),
                "application/example+xml"
            );
            assert!(candidate.has_physical_member("custom/added.xml")?);
            assert!(!candidate.has_physical_member("custom/missing.xml")?);

            let untouched = candidate.part(&pack(UNTOUCHED))?;
            Ok((untouched.partname().clone(), source_handle.read_count()))
        })
        .unwrap();
    assert_eq!(observed.0, pack(UNTOUCHED));
    assert_eq!(
        observed.1,
        source_handle.read_count(),
        "catalog inspection does not lazily read an unchanged Part"
    );

    // Prepare again so the untouched payload can be observed crossing the
    // lazy effective-Part boundary without borrowing a consumed topology.
    let prepared = open(Arc::clone(&source_handle))
        .prepare_topology(mutation_plan())
        .unwrap();
    let before_untouched_data = source_handle.read_count();
    let untouched_bytes = prepared
        .with_candidate(|candidate| {
            Ok(candidate
                .part(&pack(UNTOUCHED))?
                .data()?
                .as_bytes()
                .to_vec())
        })
        .unwrap();
    assert_eq!(untouched_bytes, b"<untouched source='yes'/>");
    assert!(source_handle.read_count() > before_untouched_data);

    let prepared = open(Arc::clone(&source_handle))
        .prepare_topology(mutation_plan())
        .unwrap();
    let mut output = Vec::new();
    prepared.publish_to_stream(&mut output).unwrap();
    assert_ne!(output, source);
    assert_eq!(
        zip_member(&output, "custom/untouched.xml"),
        zip_member(&source, "custom/untouched.xml")
    );

    let reopened_source = Arc::new(VersionedSource::new(output.clone()));
    let reopened = open(reopened_source);
    assert_eq!(
        reopened
            .part(&pack(DOCUMENT))
            .unwrap()
            .data()
            .unwrap()
            .as_bytes(),
        b"<after/>"
    );
    assert_eq!(
        reopened
            .part(&pack(ADDED))
            .unwrap()
            .data()
            .unwrap()
            .as_bytes(),
        b"<added/>"
    );
    assert_eq!(
        reopened
            .part(&pack(UNTOUCHED))
            .unwrap()
            .data()
            .unwrap()
            .as_bytes(),
        b"<untouched source='yes'/>"
    );
    assert_eq!(
        reopened.rels().get("rIdAdded").unwrap().target_ref(),
        "custom/added.xml"
    );
    assert_eq!(
        reopened
            .part(&pack(DOCUMENT))
            .unwrap()
            .rels()
            .get("rIdAdded")
            .unwrap()
            .reltype(),
        CUSTOM_REL
    );
}

#[test]
fn removed_part_is_absent_from_effective_content_type_lookup() {
    let source = source_bytes();
    let mut plan = SourceTopologyPlan::new();
    plan.try_remove_relationship(pack(DOCUMENT), "rId1")
        .unwrap();
    plan.try_remove_part(pack(UNTOUCHED)).unwrap();

    let prepared = open(Arc::new(VersionedSource::new(source)))
        .prepare_topology(plan)
        .unwrap();
    prepared
        .with_candidate(|candidate| {
            assert!(matches!(
                candidate.part(&pack(UNTOUCHED)),
                Err(OpcError::PartNotFound(_))
            ));
            assert!(matches!(
                candidate.content_type(&pack(UNTOUCHED)),
                Err(OpcError::PartNotFound(_))
            ));
            assert!(!candidate.has_physical_member("custom/untouched.xml")?);
            Ok(())
        })
        .unwrap();
}

#[test]
fn effective_relationship_graph_cap_rejects_one_new_internal_target_before_candidate() {
    let limits = litchi_opc::ReadLimits::builder()
        .max_relationship_graph_nodes(2)
        .unwrap()
        .build()
        .unwrap();
    // The fixture's package->document and document->untouched edges consume
    // exactly two graph nodes.  The added Part and internal edge would be a
    // third effective target, so preparation must refuse before a candidate or
    // publication sink can be reached.
    let package = open_with_limits(Arc::new(VersionedSource::new(source_bytes())), limits).unwrap();
    let mut plan = SourceTopologyPlan::new();
    plan.try_add_part(pack(ADDED), "application/example+xml", b"<added/>".to_vec())
        .unwrap();
    plan.try_add_internal_relationship(pack(DOCUMENT), "rIdAdded", CUSTOM_REL, pack(ADDED))
        .unwrap();

    let error = match package.prepare_topology(plan) {
        Err(error) => error,
        Ok(_) => panic!("effective graph cap must reject the new target during preparation"),
    };
    assert!(matches!(
        error,
        OpcError::ReadLimit {
            resource: litchi_opc::ReadResource::RelationshipGraphNodes,
            ..
        }
    ));
}

#[test]
fn shrinking_effective_candidate_is_accepted_at_exact_source_caps() {
    let source = source_bytes();
    let limits = litchi_opc::ReadLimits::builder()
        .max_parts(2)
        .unwrap()
        .max_total_part_bytes((b"<before/>".len() + b"<untouched source='yes'/>".len()) as u64)
        .unwrap()
        .max_total_relationships(2)
        .unwrap()
        .max_relationship_graph_nodes(2)
        .unwrap()
        .build()
        .unwrap();
    let package = open_with_limits(Arc::new(VersionedSource::new(source.clone())), limits).unwrap();
    let mut plan = SourceTopologyPlan::new();
    plan.try_replace_part(pack(DOCUMENT), b"<after/>".to_vec())
        .unwrap();
    plan.try_remove_relationship(pack(DOCUMENT), "rId1")
        .unwrap();
    plan.try_remove_part(pack(UNTOUCHED)).unwrap();

    let prepared = package.prepare_topology(plan).unwrap();
    prepared
        .with_candidate(|candidate| {
            assert_eq!(candidate.parts().len(), 1);
            assert_eq!(
                candidate.part(&pack(DOCUMENT))?.data()?.as_bytes(),
                b"<after/>"
            );
            assert!(matches!(
                candidate.part(&pack(UNTOUCHED)),
                Err(OpcError::PartNotFound(_))
            ));
            assert!(candidate.relationships(&pack(DOCUMENT))?.is_empty());
            Ok(())
        })
        .unwrap();

    let mut output = Vec::new();
    prepared.publish_to_stream(&mut output).unwrap();
    let reopened = open(Arc::new(VersionedSource::new(output)));
    assert_eq!(
        reopened
            .part(&pack(DOCUMENT))
            .unwrap()
            .data()
            .unwrap()
            .as_bytes(),
        b"<after/>"
    );
    assert!(matches!(
        reopened.part(&pack(UNTOUCHED)),
        Err(OpcError::PartNotFound(_))
    ));
}

#[test]
fn content_types_byte_cap_refuses_new_mapping_before_candidate() {
    let source = source_bytes();
    let content_types_bytes = zip_member(&source, "[Content_Types].xml");
    let limits = litchi_opc::ReadLimits::builder()
        .max_content_types_bytes(content_types_bytes.len())
        .unwrap()
        .build()
        .unwrap();
    let (_budget, _cancellation_source, context) = managed_context();
    let package = SourceBackedPackage::from_read_at_with_execution_context(
        Arc::new(VersionedSource::new(source)),
        limits,
        context,
    )
    .unwrap();
    let mut plan = SourceTopologyPlan::new();
    plan.try_add_part(pack(ADDED), "application/example+xml", b"<added/>".to_vec())
        .unwrap();

    let error = match package.prepare_topology(plan) {
        Err(error) => error,
        Ok(_) => panic!("content-types cap must reject the new mapping during preparation"),
    };
    assert!(matches!(
        error,
        OpcError::ReadLimit {
            resource: litchi_opc::ReadResource::ContentTypesBytes,
            ..
        }
    ));
}

#[test]
fn canonical_relationship_append_and_removal_accept_exact_final_boundaries() {
    let source = source_bytes();
    assert_relationship_limit_boundaries(&source, append_external_plan);
    assert_relationship_removal_subtracts_source_usage(&source, remove_canonical_plan);
}

#[test]
fn whitespace_rich_noncanonical_relationship_append_and_removal_accept_exact_boundaries() {
    let source = whitespace_rich_relationship_source();
    assert_relationship_limit_boundaries(&source, append_external_plan);
    assert_relationship_removal_subtracts_source_usage(&source, remove_noncanonical_plan);
}

fn archive_uncompressed_total(bytes: &[u8]) -> u64 {
    let archive = soapberry_zip::ZipArchive::from_slice(bytes)
        .unwrap()
        .into_zip_archive();
    let mut scratch = vec![0; soapberry_zip::RECOMMENDED_BUFFER_SIZE];
    let index = soapberry_zip::PreservationIndex::new_with_limits(
        &archive,
        &mut scratch,
        soapberry_zip::office::ArchiveLimits::UNBOUNDED,
    )
    .unwrap();
    index
        .entries()
        .iter()
        .map(|entry| entry.uncompressed_size())
        .sum()
}

#[test]
fn combined_content_types_and_relationship_growth_obeys_exact_archive_total_cap() {
    let source = source_bytes();
    let expected = publish_relationship_plan(&source, litchi_opc::ReadLimits::default(), {
        let mut plan = SourceTopologyPlan::new();
        plan.try_add_part(pack(ADDED), "application/example+xml", b"<added/>".to_vec())
            .unwrap();
        plan.try_add_external_relationship(
            pack(DOCUMENT),
            "rAdded",
            HYPERLINK_REL,
            "https://added.invalid/?q=two#fragment",
        )
        .unwrap();
        plan
    })
    .unwrap();
    let exact_total = archive_uncompressed_total(&expected);
    let source_total = archive_uncompressed_total(&source);
    assert!(exact_total > source_total);

    let make_plan = || {
        let mut plan = SourceTopologyPlan::new();
        plan.try_add_part(pack(ADDED), "application/example+xml", b"<added/>".to_vec())
            .unwrap();
        plan.try_add_external_relationship(
            pack(DOCUMENT),
            "rAdded",
            HYPERLINK_REL,
            "https://added.invalid/?q=two#fragment",
        )
        .unwrap();
        plan
    };
    let exact_limits = litchi_opc::ReadLimits::builder()
        .max_archive_total_bytes(exact_total)
        .unwrap()
        .build()
        .unwrap();
    let exact_output = publish_relationship_plan(&source, exact_limits, make_plan()).unwrap();
    assert_eq!(archive_uncompressed_total(&exact_output), exact_total);

    let under_limits = litchi_opc::ReadLimits::builder()
        .max_archive_total_bytes(exact_total - 1)
        .unwrap()
        .build()
        .unwrap();
    let package =
        open_with_limits(Arc::new(VersionedSource::new(source.clone())), under_limits).unwrap();
    let error = match package.prepare_topology(make_plan()) {
        Err(error) => error,
        Ok(_) => panic!("one byte under the final archive total must refuse preparation"),
    };
    assert!(matches!(
        error,
        OpcError::ReadLimit {
            resource: litchi_opc::ReadResource::ArchiveTotalBytes,
            ..
        }
    ));
}

#[test]
fn callback_failure_writes_no_bytes_and_prepared_value_remains_publishable() {
    let source = source_bytes();
    let prepared = open(Arc::new(VersionedSource::new(source.clone())))
        .prepare_topology(mutation_plan())
        .unwrap();
    let callback = prepared.with_candidate::<(), _>(|_| {
        Err(OpcError::SourceBackedOverlayUnavailable {
            reason: "semantic callback refused candidate".to_owned(),
        })
    });
    assert!(matches!(
        callback,
        Err(OpcError::SourceBackedOverlayUnavailable { reason })
            if reason.contains("callback")
    ));

    let mut output = Vec::new();
    prepared.publish_to_stream(&mut output).unwrap();
    assert!(!output.is_empty());
    assert_eq!(zip_member(&output, "word/document.xml"), b"<after/>");
}

#[test]
fn borrowed_prepared_candidate_keeps_source_usable_and_matches_consuming_publish() {
    let source = source_bytes();
    let source_handle = Arc::new(VersionedSource::new(source.clone()));
    let package = open(Arc::clone(&source_handle));
    package
        .with_prepared_topology(mutation_plan(), |candidate| {
            assert_eq!(
                candidate.part(&pack(DOCUMENT))?.data()?.as_bytes(),
                b"<after/>"
            );
            assert_eq!(
                candidate.part(&pack(ADDED))?.data()?.as_bytes(),
                b"<added/>"
            );
            assert_eq!(
                candidate
                    .relationships(&pack(DOCUMENT))?
                    .get("rIdAdded")
                    .unwrap()
                    .target_ref(),
                "../custom/added.xml"
            );
            assert_eq!(
                candidate.content_type(&pack(ADDED))?.as_str(),
                "application/example+xml"
            );
            Ok(())
        })
        .unwrap();

    // The borrowed helper does not consume the source-backed package. The
    // original source remains readable and the same package can later build a
    // consuming PreparedTopology whose publication matches an independent
    // source-backed run.
    assert_eq!(
        package
            .part(&pack(DOCUMENT))
            .unwrap()
            .data()
            .unwrap()
            .as_bytes(),
        b"<before/>"
    );
    let mut actual = Vec::new();
    package
        .prepare_topology(mutation_plan())
        .unwrap()
        .publish_to_stream(&mut actual)
        .unwrap();
    let mut expected = Vec::new();
    open(Arc::new(VersionedSource::new(source)))
        .prepare_topology(mutation_plan())
        .unwrap()
        .publish_to_stream(&mut expected)
        .unwrap();
    assert_eq!(actual, expected);
}

#[test]
fn replacement_only_candidate_uses_source_content_type_without_manifest_reparse() {
    let source = source_bytes();
    let content_types_range = stored_member_data_range(&source, "[Content_Types].xml");
    let source_handle = Arc::new(VersionedSource::new(source));
    let package = open(Arc::clone(&source_handle));
    let mut plan = SourceTopologyPlan::new();
    plan.try_replace_part(pack(DOCUMENT), b"<after/>".to_vec())
        .unwrap();

    source_handle.deny_reads_in(content_types_range);
    let prepared = package.prepare_topology(plan).unwrap();
    prepared
        .with_candidate(|candidate| {
            assert_eq!(
                candidate.content_type(&pack(DOCUMENT))?.as_str(),
                ct::WML_DOCUMENT
            );
            assert_eq!(
                candidate.part(&pack(UNTOUCHED))?.content_type().as_str(),
                "application/xml"
            );
            assert_eq!(
                candidate.part(&pack(DOCUMENT))?.data()?.as_bytes(),
                b"<after/>"
            );
            Ok(())
        })
        .unwrap();

    source_handle.allow_all_reads();
    let mut output = Vec::new();
    prepared.publish_to_stream(&mut output).unwrap();
    assert_eq!(zip_member(&output, "word/document.xml"), b"<after/>");
    assert!(source_handle.read_count() > 0);
}

#[test]
fn borrowed_prepared_callback_failure_leaves_source_reusable() {
    let source = source_bytes();
    let package = open(Arc::new(VersionedSource::new(source)));
    let error: litchi_opc::Result<()> = package.with_prepared_topology(mutation_plan(), |_| {
        Err(OpcError::SourceBackedOverlayUnavailable {
            reason: "borrowed callback refused candidate".to_owned(),
        })
    });
    assert!(matches!(
        error,
        Err(OpcError::SourceBackedOverlayUnavailable { reason })
            if reason.contains("borrowed callback")
    ));
    assert_eq!(
        package
            .part(&pack(DOCUMENT))
            .unwrap()
            .data()
            .unwrap()
            .as_bytes(),
        b"<before/>"
    );
    let mut output = Vec::new();
    package
        .prepare_topology(mutation_plan())
        .unwrap()
        .publish_to_stream(&mut output)
        .unwrap();
    assert_eq!(zip_member(&output, "word/document.xml"), b"<after/>");
}

#[test]
fn borrowed_prepared_callback_source_change_is_refused_after_callback() {
    let source_handle = Arc::new(VersionedSource::new(source_bytes()));
    let package = open(Arc::clone(&source_handle));
    let entered = Arc::new(AtomicBool::new(false));
    let callback_entered = Arc::clone(&entered);
    let error = package.with_prepared_topology(mutation_plan(), |_| {
        callback_entered.store(true, Ordering::SeqCst);
        source_handle.bump();
        Ok(())
    });
    assert!(entered.load(Ordering::SeqCst));
    assert!(matches!(error, Err(OpcError::SourceChanged { .. })));
}

#[test]
fn borrowed_noop_does_not_invoke_callback_and_keeps_source_usable() {
    let package = open(Arc::new(VersionedSource::new(source_bytes())));
    let invoked = Arc::new(AtomicBool::new(false));
    let callback_invoked = Arc::clone(&invoked);
    let error = package.with_prepared_topology(SourceTopologyPlan::new(), |_| {
        callback_invoked.store(true, Ordering::SeqCst);
        Ok(())
    });
    assert!(matches!(
        error,
        Err(OpcError::SourceBackedOverlayUnavailable { .. })
    ));
    assert!(!invoked.load(Ordering::SeqCst));
    assert_eq!(
        package
            .part(&pack(DOCUMENT))
            .unwrap()
            .data()
            .unwrap()
            .as_bytes(),
        b"<before/>"
    );
}

#[test]
fn borrowed_candidate_refuses_transferred_source_change_after_callback() {
    let foreign_handle = Arc::new(VersionedSource::new(source_bytes()));
    let foreign = open(Arc::clone(&foreign_handle));
    let (_, token) = foreign
        .part(&pack(UNTOUCHED))
        .unwrap()
        .data_and_authorize_precompressed()
        .unwrap();
    let package = open(Arc::new(VersionedSource::new(source_bytes())));
    let mut plan = SourceTopologyPlan::new();
    plan.try_add_precompressed_part(pack("/custom/transferred.xml"), "application/xml", token)
        .unwrap();
    let mut entered = false;
    let result = package.with_prepared_topology(plan, |candidate| {
        assert_eq!(
            candidate
                .part(&pack("/custom/transferred.xml"))?
                .data()?
                .as_bytes(),
            b"<untouched source='yes'/>"
        );
        entered = true;
        foreign_handle.bump();
        Ok(())
    });
    assert!(entered);
    assert!(matches!(result, Err(OpcError::SourceChanged { .. })));
    assert_eq!(
        package
            .part(&pack(DOCUMENT))
            .unwrap()
            .data()
            .unwrap()
            .as_bytes(),
        b"<before/>"
    );
}

#[test]
fn managed_borrowed_candidate_keeps_prepared_payload_reserved_during_callback() {
    let (budget, _cancellation_source, context) = managed_context();
    let package = SourceBackedPackage::from_read_at_with_execution_context(
        Arc::new(VersionedSource::new(source_bytes())),
        litchi_opc::ReadLimits::default(),
        context,
    )
    .unwrap();
    let before = budget.used(Resource::Memory);
    package
        .with_prepared_topology(mutation_plan(), |candidate| {
            assert_eq!(
                candidate.part(&pack(ADDED))?.data()?.as_bytes(),
                b"<added/>"
            );
            assert!(budget.used(Resource::Memory) >= before);
            Ok(())
        })
        .unwrap();
}

#[test]
fn large_unchanged_part_stays_cold_and_publishes_under_smaller_managed_memory_cap() {
    let (source, large_payload) = source_bytes_with_large_untouched_part();
    // The managed cache reserves a Part's declared uncompressed size before
    // reading it. Keep this below the 128 KiB large member so a successful
    // candidate inspection proves that the unchanged member was never decoded
    // into a managed PartData allocation. Fixture creation and the
    // caller-owned publication Vec are intentionally outside this budget.
    const MANAGED_MEMORY_LIMIT: u64 = 64 * 1024;
    let source_handle = Arc::new(VersionedSource::new(source.clone()));
    let (budget, _cancellation_source, context) = managed_context_with_resources(
        MANAGED_MEMORY_LIMIT,
        u64::MAX,
        u64::MAX,
        u64::MAX,
        u64::MAX,
    );
    let source_for_package: Arc<dyn ReadAt> = source_handle.clone();
    let package = SourceBackedPackage::from_read_at_with_execution_context(
        source_for_package,
        litchi_opc::ReadLimits::default(),
        context,
    )
    .unwrap();
    let mut plan = SourceTopologyPlan::new();
    plan.try_replace_part(pack(DOCUMENT), b"<after/>".to_vec())
        .unwrap();
    let diagnostics_before_prepare = package.cache_diagnostics();
    package
        .with_prepared_topology(plan, |candidate| {
            let diagnostics_before_metadata = package.cache_diagnostics();
            let large = candidate.part(&pack(LARGE_UNTOUCHED))?;
            assert_eq!(large.content_type().as_str(), "application/octet-stream");
            assert!(candidate.has_physical_member("custom/large.bin")?);
            let diagnostics_after_metadata = package.cache_diagnostics();
            assert_eq!(
                diagnostics_after_metadata.cold_loads, diagnostics_before_metadata.cold_loads,
                "candidate metadata inspection must not decode the large unchanged Part"
            );
            assert_eq!(
                diagnostics_after_metadata.successful_loads,
                diagnostics_before_metadata.successful_loads
            );
            Ok(())
        })
        .unwrap();
    let diagnostics_after_prepare = package.cache_diagnostics();
    assert!(diagnostics_after_prepare.cold_loads >= diagnostics_before_prepare.cold_loads);
    assert!(
        diagnostics_after_prepare.successful_loads >= diagnostics_before_prepare.successful_loads
    );
    assert!(diagnostics_after_prepare.budget_cache_reserved_bytes < large_payload.len() as u64);
    assert_eq!(
        diagnostics_after_prepare.budget_memory_limit,
        Some(MANAGED_MEMORY_LIMIT)
    );
    assert!(budget.used(Resource::Memory) < large_payload.len() as u64);

    // Prove that the same managed session would refuse to decode this member
    // itself. The successful candidate path below therefore demonstrates
    // cold metadata access, rather than merely relying on a smaller numeric
    // budget assertion.
    let (_large_read_budget, _large_read_cancellation, large_read_context) =
        managed_context_with_resources(
            MANAGED_MEMORY_LIMIT,
            u64::MAX,
            u64::MAX,
            u64::MAX,
            u64::MAX,
        );
    let large_read_source: Arc<dyn ReadAt> = source_handle.clone();
    let large_read_package = SourceBackedPackage::from_read_at_with_execution_context(
        large_read_source,
        litchi_opc::ReadLimits::default(),
        large_read_context,
    )
    .unwrap();
    let large_read_error = large_read_package
        .part(&pack(LARGE_UNTOUCHED))
        .unwrap()
        .data()
        .unwrap_err();
    assert!(matches!(
        large_read_error,
        OpcError::Execution(ExecutionError::ResourceLimit(limit))
            if limit.resource == Resource::Memory
    ));

    let (publication_budget, _cancellation_source, publication_context) =
        managed_context_with_resources(
            MANAGED_MEMORY_LIMIT,
            u64::MAX,
            u64::MAX,
            u64::MAX,
            u64::MAX,
        );
    let publication_source: Arc<dyn ReadAt> = source_handle.clone();
    let publication_package = SourceBackedPackage::from_read_at_with_execution_context(
        publication_source,
        litchi_opc::ReadLimits::default(),
        publication_context,
    )
    .unwrap();
    let mut publication_plan = SourceTopologyPlan::new();
    publication_plan
        .try_replace_part(pack(DOCUMENT), b"<after/>".to_vec())
        .unwrap();
    let prepared = publication_package
        .prepare_topology(publication_plan)
        .unwrap();
    let mut output = Vec::new();
    prepared.publish_to_stream(&mut output).unwrap();
    assert!(publication_budget.used(Resource::Memory) < large_payload.len() as u64);
    assert_eq!(zip_member(&output, "custom/large.bin"), large_payload);
}

#[test]
fn managed_low_memory_refuses_preparation_before_candidate() {
    let (_budget, _cancellation_source, context) =
        managed_context_with_resources(64 * 1024, u64::MAX, u64::MAX, u64::MAX, u64::MAX);
    let package = SourceBackedPackage::from_read_at_with_execution_context(
        Arc::new(VersionedSource::new(whitespace_rich_relationship_source())),
        litchi_opc::ReadLimits::default(),
        context,
    )
    .unwrap();
    let invoked = Arc::new(AtomicBool::new(false));
    let callback_invoked = Arc::clone(&invoked);
    let result: litchi_opc::Result<()> =
        package.with_prepared_topology(append_external_plan(), |_| {
            callback_invoked.store(true, Ordering::SeqCst);
            Ok(())
        });
    assert!(matches!(
        result,
        Err(OpcError::Execution(ExecutionError::ResourceLimit(limit)))
            if limit.resource == Resource::Memory
    ));
    assert!(!invoked.load(Ordering::SeqCst));
}

#[test]
fn managed_low_input_limit_refuses_source_open() {
    let (_budget, _cancellation_source, context) =
        managed_context_with_resources(16 * 1024 * 1024, 1, u64::MAX, u64::MAX, u64::MAX);
    let result = SourceBackedPackage::from_read_at_with_execution_context(
        Arc::new(VersionedSource::new(source_bytes())),
        litchi_opc::ReadLimits::default(),
        context,
    );
    assert!(matches!(
        result,
        Err(OpcError::Execution(ExecutionError::ResourceLimit(limit)))
            if limit.resource == Resource::InputBytes
    ));
}

#[test]
fn managed_low_work_limit_refuses_candidate_before_callback() {
    let (_budget, _cancellation_source, context) =
        managed_context_with_resources(16 * 1024 * 1024, u64::MAX, u64::MAX, u64::MAX, 1);
    let package = SourceBackedPackage::from_read_at_with_execution_context(
        Arc::new(VersionedSource::new(source_bytes())),
        litchi_opc::ReadLimits::default(),
        context,
    )
    .unwrap();
    let invoked = Arc::new(AtomicBool::new(false));
    let callback_invoked = Arc::clone(&invoked);
    let result: litchi_opc::Result<()> = package.with_prepared_topology(mutation_plan(), |_| {
        callback_invoked.store(true, Ordering::SeqCst);
        Ok(())
    });
    assert!(matches!(
        result,
        Err(OpcError::Execution(ExecutionError::ResourceLimit(limit)))
            if limit.resource == Resource::Work
    ));
    assert!(!invoked.load(Ordering::SeqCst));
}

#[test]
fn managed_low_output_limit_reports_sink_accounting() {
    let output_limit = 512;
    let (budget, _cancellation_source, context) = managed_context_with_resources(
        16 * 1024 * 1024,
        u64::MAX,
        output_limit,
        u64::MAX,
        u64::MAX,
    );
    let package = SourceBackedPackage::from_read_at_with_execution_context(
        Arc::new(VersionedSource::new(source_bytes())),
        litchi_opc::ReadLimits::default(),
        context,
    )
    .unwrap();
    let prepared = package.prepare_topology(mutation_plan()).unwrap();
    let mut output = Vec::new();
    let result = prepared.publish_to_stream(&mut output);
    match result {
        Err(OpcError::IncompleteOutput { written, .. }) => {
            assert_eq!(written, output.len() as u64);
            assert!(written > 0);
        },
        Err(OpcError::Execution(ExecutionError::ResourceLimit(limit)))
            if limit.resource == Resource::OutputBytes =>
        {
            assert!(output.is_empty());
        },
        other => panic!("unexpected low-output result: {other:?}"),
    }
    assert_eq!(budget.used(Resource::OutputBytes), output.len() as u64);
    assert!(output.len() as u64 <= output_limit);
}

#[test]
fn source_change_after_prepare_refuses_candidate_and_publication_before_sink_bytes() {
    let source = source_bytes();
    let source_handle = Arc::new(VersionedSource::new(source));
    let prepared = open(Arc::clone(&source_handle))
        .prepare_topology(mutation_plan())
        .unwrap();
    source_handle.bump();
    assert!(matches!(
        prepared.with_candidate(|_| Ok(())),
        Err(OpcError::SourceChanged { .. })
    ));
    let mut output = Vec::new();
    assert!(matches!(
        prepared.publish_to_stream(&mut output),
        Err(OpcError::SourceChanged { .. })
    ));
    assert!(output.is_empty());

    let callback_source = Arc::new(VersionedSource::new(source_bytes()));
    let callback_prepared = open(Arc::clone(&callback_source))
        .prepare_topology(mutation_plan())
        .unwrap();
    let callback_result = callback_prepared.with_candidate(|_| {
        callback_source.bump();
        Ok(())
    });
    assert!(matches!(
        callback_result,
        Err(OpcError::SourceChanged { .. })
    ));
    let mut callback_output = Vec::new();
    assert!(matches!(
        callback_prepared.publish_to_stream(&mut callback_output),
        Err(OpcError::SourceChanged { .. })
    ));
    assert!(callback_output.is_empty());
}

#[test]
fn exact_noop_preparation_is_source_byte_exact_and_has_no_candidate() {
    let source = source_bytes();
    let prepared = open(Arc::new(VersionedSource::new(source.clone())))
        .prepare_topology(SourceTopologyPlan::new())
        .unwrap();
    assert!(prepared.with_candidate(|_| Ok(())).is_err());
    let mut output = Vec::new();
    prepared.publish_to_stream(&mut output).unwrap();
    assert_eq!(output, source);
}

#[test]
fn stale_source_refuses_even_empty_preparation() {
    let source = Arc::new(VersionedSource::new(source_bytes()));
    let package = open(Arc::clone(&source));
    source.bump();
    let result: litchi_opc::Result<()> = package
        .with_prepared_topology(SourceTopologyPlan::new(), |_| {
            panic!("empty stale preparation cannot enter a callback")
        });
    assert!(matches!(result, Err(OpcError::SourceChanged { .. })));
    assert!(matches!(
        package.prepare_topology(SourceTopologyPlan::new()),
        Err(OpcError::SourceChanged { .. })
    ));
}

#[test]
fn trailing_archive_bytes_refuse_changed_candidate_but_preserve_exact_noop() {
    let mut source = source_bytes();
    source.extend_from_slice(b"opaque trailing source bytes");
    let package = open(Arc::new(VersionedSource::new(source.clone())));
    let mut invoked = false;
    let result = package.with_prepared_topology(mutation_plan(), |_| {
        invoked = true;
        Ok(())
    });
    assert!(result.is_err());
    assert!(
        !invoked,
        "publication boundary refusal must precede semantic readback"
    );
    assert!(package.prepare_topology(mutation_plan()).is_err());

    let package = open(Arc::new(VersionedSource::new(source.clone())));
    let mut output = Vec::new();
    package
        .prepare_topology(SourceTopologyPlan::new())
        .unwrap()
        .publish_to_stream(&mut output)
        .unwrap();
    assert_eq!(output, source);
}

#[test]
fn prepared_publication_reports_sequential_sink_failure_accounting() {
    let prepared = open(Arc::new(VersionedSource::new(source_bytes())))
        .prepare_topology(mutation_plan())
        .unwrap();
    let mut sink = PartialWriter {
        bytes: Vec::new(),
        remaining: 37,
    };
    let error = prepared.publish_to_stream(&mut sink).unwrap_err();
    match error {
        OpcError::IncompleteOutput { written, .. } => {
            assert_eq!(written, sink.bytes.len() as u64);
            assert!(written > 0);
        },
        other => panic!("expected incomplete output, got {other:?}"),
    }
}

#[test]
fn managed_prepared_candidate_reads_and_publishes_with_reservations_attached() {
    let source = source_bytes();
    let (_budget, cancellation_source, context) = managed_context();
    let source_handle = Arc::new(VersionedSource::new(source));
    let package = SourceBackedPackage::from_read_at_with_execution_context(
        source_handle,
        litchi_opc::ReadLimits::default(),
        context,
    )
    .unwrap();
    assert!(package.cache_diagnostics().budget_managed);
    let prepared = package.prepare_topology(mutation_plan()).unwrap();
    let untouched = prepared
        .with_candidate(|candidate| {
            Ok(candidate
                .part(&pack(UNTOUCHED))?
                .data()?
                .as_bytes()
                .to_vec())
        })
        .unwrap();
    assert_eq!(untouched, b"<untouched source='yes'/>");
    let mut output = Vec::new();
    prepared.publish_to_stream(&mut output).unwrap();
    assert_eq!(zip_member(&output, "custom/added.xml"), b"<added/>");
    drop(cancellation_source);
}

#[test]
fn prepared_source_xml_token_keeps_managed_reservation_until_publication_drop() {
    let (budget, _cancellation_source, context) = managed_context();
    let source_handle = Arc::new(VersionedSource::new(source_bytes()));
    let package = SourceBackedPackage::from_read_at_with_execution_context(
        source_handle,
        litchi_opc::ReadLimits::default(),
        context,
    )
    .unwrap();
    let before_capture = budget.used(Resource::Memory);
    let source_xml = package.part(&pack(DOCUMENT)).unwrap().source_xml().unwrap();
    let after_capture = budget.used(Resource::Memory);
    assert!(after_capture > before_capture);
    let proof = source_xml
        .checked_range(0..b"<before/>".len(), b"<before/>")
        .unwrap();
    let mut publication = source_xml.into_publication().unwrap();
    publication
        .replace(
            proof,
            AuthoredXmlFragment::markup(b"<changed/>".to_vec()).unwrap(),
        )
        .unwrap();
    let staged = publication.finish().unwrap();
    let after_staging = budget.used(Resource::Memory);
    assert!(after_staging >= after_capture);

    let mut plan = SourceTopologyPlan::new();
    plan.try_replace_source_xml_part(pack(DOCUMENT), staged)
        .unwrap();
    let prepared = package.prepare_topology(plan).unwrap();
    let after_prepare = budget.used(Resource::Memory);
    assert!(
        after_prepare >= after_staging,
        "prepared topology must retain source XML token reservations"
    );
    assert_eq!(
        prepared
            .with_candidate(|candidate| {
                Ok(candidate.part(&pack(DOCUMENT))?.data()?.as_bytes().to_vec())
            })
            .unwrap(),
        b"<changed/>"
    );

    let mut output = Vec::new();
    prepared.publish_to_stream(&mut output).unwrap();
    assert_eq!(zip_member(&output, "word/document.xml"), b"<changed/>");
    drop(prepared);
    assert!(
        budget.used(Resource::Memory) < after_prepare,
        "source XML token reservation should release after prepared publication drops it"
    );
}

#[test]
fn prepared_precompressed_token_keeps_managed_reservation_until_publication_drop() {
    let (budget, _cancellation_source, context) = managed_context();
    let source_handle = Arc::new(VersionedSource::new(source_bytes()));
    let package = SourceBackedPackage::from_read_at_with_execution_context(
        source_handle,
        litchi_opc::ReadLimits::default(),
        context,
    )
    .unwrap();
    let before_capture = budget.used(Resource::Memory);
    let (decoded, token) = package
        .part(&pack(UNTOUCHED))
        .unwrap()
        .data_and_authorize_precompressed()
        .unwrap();
    let expected = decoded.as_bytes().to_vec();
    let after_capture = budget.used(Resource::Memory);
    assert!(after_capture > before_capture);
    drop(decoded);
    let after_decoded_drop = budget.used(Resource::Memory);
    assert!(after_decoded_drop >= before_capture);

    let mut plan = SourceTopologyPlan::new();
    plan.try_add_precompressed_part(pack("/custom/transferred.xml"), "application/xml", token)
        .unwrap();
    let prepared = package.prepare_topology(plan).unwrap();
    let after_prepare = budget.used(Resource::Memory);
    assert!(
        after_prepare >= after_decoded_drop,
        "prepared topology must retain the precompressed token reservation"
    );
    let candidate_bytes = prepared
        .with_candidate(|candidate| {
            Ok(candidate
                .part(&pack("/custom/transferred.xml"))?
                .data()?
                .as_bytes()
                .to_vec())
        })
        .unwrap();
    assert_eq!(candidate_bytes, expected);

    let mut output = Vec::new();
    prepared.publish_to_stream(&mut output).unwrap();
    assert_eq!(zip_member(&output, "custom/transferred.xml"), expected);
    drop(prepared);
    assert!(
        budget.used(Resource::Memory) < after_prepare,
        "precompressed token reservation should release after prepared publication drops it"
    );
}

#[test]
fn retained_preparation_retries_and_publishes_shared_candidate_repeatedly() {
    fn assert_shared<T: Clone + Send + Sync>() {}
    assert_shared::<litchi_opc::PreparedTopology>();
    let source = Arc::new(VersionedSource::new(source_bytes()));
    let prepared = open(Arc::clone(&source))
        .prepare_topology(mutation_plan())
        .unwrap();
    let reads = source.read_count();
    let shared = prepared.clone();
    assert_eq!(source.read_count(), reads);
    let mut failing = PartialWriter {
        bytes: Vec::new(),
        remaining: 37,
    };
    assert!(matches!(
        prepared.publish_to_stream(&mut failing),
        Err(OpcError::IncompleteOutput { .. })
    ));
    let mut first = Vec::new();
    prepared.publish_to_stream(&mut first).unwrap();
    let mut second = Vec::new();
    shared.publish_to_stream(&mut second).unwrap();
    assert_eq!(first, second);
    assert_eq!(zip_member(&second, "custom/added.xml"), b"<added/>");
    drop(prepared);
    shared
        .with_candidate(|candidate| {
            assert_eq!(
                candidate.part(&pack(DOCUMENT))?.data()?.as_bytes(),
                b"<after/>"
            );
            assert_eq!(
                candidate.part(&pack(ADDED))?.data()?.as_bytes(),
                b"<added/>"
            );
            Ok(())
        })
        .unwrap();
    source.bump();
    let mut rejected = Vec::new();
    assert!(matches!(
        shared.publish_to_stream(&mut rejected),
        Err(OpcError::SourceChanged { .. })
    ));
    assert!(rejected.is_empty());
}

#[test]
fn prepared_clones_publish_concurrently_from_one_retained_topology() {
    let prepared = open(Arc::new(VersionedSource::new(source_bytes())))
        .prepare_topology(mutation_plan())
        .unwrap();
    let outputs = std::thread::scope(|scope| {
        let left = prepared.clone();
        let right = prepared.clone();
        let publish = |prepared: litchi_opc::PreparedTopology| {
            let mut output = Vec::new();
            prepared.publish_to_stream(&mut output).unwrap();
            output
        };
        let left = scope.spawn(move || publish(left));
        let right = scope.spawn(move || publish(right));
        (left.join().unwrap(), right.join().unwrap())
    });
    assert_eq!(outputs.0, outputs.1);
    prepared
        .with_candidate(|candidate| {
            assert_eq!(
                candidate.part(&pack(DOCUMENT))?.data()?.as_bytes(),
                b"<after/>"
            );
            Ok(())
        })
        .unwrap();
}

#[test]
fn retained_prepared_reservations_release_after_the_last_share() {
    let (budget, _cancellation, context) = managed_context();
    let package = SourceBackedPackage::from_read_at_with_execution_context(
        Arc::new(VersionedSource::new(source_bytes())),
        litchi_opc::ReadLimits::default(),
        context,
    )
    .unwrap();
    let prepared = package.prepare_topology(mutation_plan()).unwrap();
    let before = budget.used(Resource::Memory);
    let objects = budget.used(Resource::Objects);
    assert!(before > 0);
    let shared = prepared.clone();
    assert_eq!(budget.used(Resource::Memory), before);
    assert_eq!(budget.used(Resource::Objects), objects);
    let mut first = Vec::new();
    prepared.publish_to_stream(&mut first).unwrap();
    drop(prepared);
    assert!(budget.used(Resource::Memory) > 0);
    let mut second = Vec::new();
    shared.publish_to_stream(&mut second).unwrap();
    assert_eq!(first, second);
    drop(shared);
    assert_eq!(budget.used(Resource::Memory), 0);
    assert_eq!(budget.used(Resource::Objects), 0);
}

#[test]
fn retained_preparation_rechecks_a_transfer_source_on_every_publication() {
    let foreign_handle = Arc::new(VersionedSource::new(source_bytes()));
    let foreign = open(Arc::clone(&foreign_handle));
    let (_, token) = foreign
        .part(&pack(UNTOUCHED))
        .unwrap()
        .data_and_authorize_precompressed()
        .unwrap();
    let package = open(Arc::new(VersionedSource::new(source_bytes())));
    let mut plan = SourceTopologyPlan::new();
    plan.try_add_precompressed_part(pack("/custom/transferred.xml"), "application/xml", token)
        .unwrap();
    let prepared = package.prepare_topology(plan).unwrap();
    let next = prepared
        .edit_xml(&pack(DOCUMENT), |xml, edit| {
            edit.replace(
                xml.checked_range(0..xml.bytes().len(), xml.bytes())?,
                AuthoredXmlFragment::markup(b"<chained/>".to_vec())?,
            )
        })
        .unwrap();
    drop(prepared);
    let prepared = next;
    drop(foreign);
    let mut first = Vec::new();
    prepared.publish_to_stream(&mut first).unwrap();
    assert_eq!(
        zip_member(&first, "custom/transferred.xml"),
        b"<untouched source='yes'/>"
    );
    let transferred = prepared.part(&pack("/custom/transferred.xml")).unwrap();
    let captured = transferred.data().unwrap();
    foreign_handle.bump();
    assert!(matches!(
        transferred.data(),
        Err(OpcError::SourceChanged { .. })
    ));
    assert!(matches!(
        prepared.main_document_part(),
        Err(OpcError::SourceChanged { .. })
    ));
    assert_eq!(captured.as_bytes(), b"<untouched source='yes'/>");
    let mut second = Vec::new();
    assert!(matches!(
        prepared.publish_to_stream(&mut second),
        Err(OpcError::SourceChanged { .. })
    ));
    assert!(second.is_empty());
    assert!(matches!(
        prepared.with_candidate(|_| Ok(())),
        Err(OpcError::SourceChanged { .. })
    ));
}

#[test]
fn owned_effective_views_are_lazy_and_outlive_an_unchanged_candidate() {
    fn send_sync<T: Send + Sync>() {}
    send_sync::<litchi_opc::PreparedPartView>();
    send_sync::<litchi_opc::PreparedPartData>();
    let (generated, _) = source_bytes_with_large_untouched_part();
    let reader = litchi_opc::phys_pkg::PhysPkgReader::new(&generated).unwrap();
    let mut writer = StreamingArchiveWriter::new();
    for name in reader.member_names().unwrap() {
        writer
            .write_stored(&name, &reader.read_member(&name).unwrap())
            .unwrap();
    }
    let bytes = writer.finish_to_bytes().unwrap();
    let denied = stored_member_data_range(&bytes, "custom/large.bin");
    let source = Arc::new(VersionedSource::new(bytes));
    let candidate = open(Arc::clone(&source))
        .prepare_topology(SourceTopologyPlan::new())
        .unwrap();
    source.deny_reads_in(denied);
    let before = source.read_count();
    let main = candidate.main_document_part().unwrap();
    let parts = candidate.iter_parts().unwrap();
    assert_eq!(main.partname(), &pack(DOCUMENT));
    assert_eq!(main.content_type(), ct::WML_DOCUMENT);
    assert_eq!(
        main.declared_uncompressed_size().unwrap(),
        b"<before/>".len() as u64
    );
    assert!(candidate.is_noop());
    assert_eq!(candidate.read_limits(), litchi_opc::ReadLimits::default());
    assert!(candidate.execution_context().is_none());
    assert_eq!(
        source.read_count(),
        before,
        "capturing metadata must not decode payloads"
    );
    drop(candidate);
    let names: Vec<_> = parts.map(|part| part.partname().clone()).collect();
    assert!(names.contains(&pack(DOCUMENT)));
    assert!(names.contains(&pack(UNTOUCHED)));
    let data = main.data().unwrap();
    assert_eq!(data.as_bytes(), b"<before/>");
    let cloned = data.clone();
    assert_eq!(data.as_bytes().as_ptr(), cloned.as_bytes().as_ptr());
    drop(main);
    drop(data);
    assert_eq!(cloned.as_bytes(), b"<before/>");
}

#[test]
fn owned_effective_views_read_changes_additions_removals_and_relationships() {
    let package = open(Arc::new(VersionedSource::new(source_bytes())));
    let mut plan = mutation_plan();
    plan.try_remove_relationship(pack(DOCUMENT), "rId1")
        .unwrap();
    plan.try_remove_part(pack(UNTOUCHED)).unwrap();
    let candidate = package.prepare_topology(plan).unwrap();
    assert!(matches!(
        candidate.part(&pack(UNTOUCHED)),
        Err(OpcError::PartNotFound(_))
    ));
    let main = candidate.main_document_part().unwrap();
    assert!(main.rels().get("rId1").is_none());
    assert!(main.rels().get("rIdAdded").is_some());
    assert!(candidate.rels().get("rIdAdded").is_some());
    assert_eq!(main.data().unwrap().as_bytes(), b"<after/>");
    assert!(
        main.source_xml().is_err(),
        "authored bytes are readable but have no physical XML proof"
    );
    let added = candidate.part(&pack(ADDED)).unwrap();
    assert_eq!(added.content_type(), "application/example+xml");
    assert_eq!(
        added.declared_uncompressed_size().unwrap(),
        b"<added/>".len() as u64
    );
    assert_eq!(
        main.declared_uncompressed_size().unwrap(),
        b"<after/>".len() as u64
    );
    let data = added.data().unwrap();
    assert!(added.source_xml().is_err());
    drop(candidate);
    drop(main);
    drop(added);
    assert_eq!(data.as_bytes(), b"<added/>");
    assert!(!format!("{data:?}").contains("<added/>"));
}

#[test]
fn owned_main_view_resolves_the_effective_root_relationship() {
    let package = open(Arc::new(VersionedSource::new(source_bytes())));
    let mut plan = mutation_plan();
    plan.try_replace_internal_relationship(pack("/"), "rId1", rt::OFFICE_DOCUMENT, pack(ADDED))
        .unwrap();
    let candidate = package.prepare_topology(plan).unwrap();
    assert_eq!(
        candidate.main_document_part().unwrap().partname(),
        &pack(ADDED)
    );
}

#[test]
fn owned_effective_views_keep_prior_revisions_and_capture_current_xml_provenance() {
    let root = open(Arc::new(VersionedSource::new(source_bytes())))
        .prepare_topology(SourceTopologyPlan::new())
        .unwrap();
    let first = root
        .edit_xml(&pack(DOCUMENT), |xml, edit| {
            edit.replace(
                xml.checked_range(0..xml.bytes().len(), xml.bytes())?,
                AuthoredXmlFragment::markup(b"<first/>".to_vec())?,
            )
        })
        .unwrap();
    let prior = first.main_document_part().unwrap();
    let second = first
        .edit_xml(&pack(DOCUMENT), |xml, edit| {
            edit.replace(
                xml.checked_range(0..xml.bytes().len(), xml.bytes())?,
                AuthoredXmlFragment::markup(b"<second/>".to_vec())?,
            )
        })
        .unwrap();
    let current = second.main_document_part().unwrap();
    assert!(!second.is_noop());
    assert_eq!(
        second.physical_source_lineage(),
        root.physical_source_lineage()
    );
    assert_eq!(
        second.physical_source_version().unwrap(),
        root.physical_source_version().unwrap()
    );
    assert_eq!(
        second.physical_source_artifact().fingerprint().unwrap(),
        root.physical_source_artifact().fingerprint().unwrap()
    );
    assert_eq!(prior.data().unwrap().as_bytes(), b"<first/>");
    assert_eq!(current.data().unwrap().as_bytes(), b"<second/>");
    assert_eq!(current.source_xml().unwrap().bytes(), b"<second/>");
    let old = prior.source_xml().unwrap();
    let stale = old.checked_range(0..0, b"").unwrap();
    assert!(
        second
            .edit_xml(&pack(DOCUMENT), |_, edit| edit
                .replace(stale, AuthoredXmlFragment::empty()))
            .is_err()
    );
    drop(root);
    drop(first);
    drop(second);
    let result = std::thread::spawn(move || current.data().unwrap())
        .join()
        .unwrap();
    assert_eq!(result.as_bytes(), b"<second/>");
}

#[test]
fn owned_payloads_pin_reservations_and_new_reads_recheck_source_and_cancellation() {
    let (budget, _cancel, context) = managed_context();
    let source = Arc::new(VersionedSource::new(source_bytes()));
    let package = SourceBackedPackage::from_read_at_with_execution_context(
        source.clone(),
        litchi_opc::ReadLimits::default(),
        context,
    )
    .unwrap();
    let candidate = package.prepare_topology(mutation_plan()).unwrap();
    let main = candidate.main_document_part().unwrap();
    let untouched = candidate.part(&pack(UNTOUCHED)).unwrap();
    let data = main.data().unwrap();
    let original = untouched.data().unwrap();
    let used = budget.used(Resource::Memory);
    let clone = main.clone();
    let data_clone = data.clone();
    assert_eq!(budget.used(Resource::Memory), used);
    drop(candidate);
    drop(main);
    drop(untouched);
    source.bump();
    assert!(matches!(clone.data(), Err(OpcError::SourceChanged { .. })));
    assert!(matches!(
        clone.source_xml(),
        Err(OpcError::SourceChanged { .. })
    ));
    assert_eq!(data_clone.as_bytes(), b"<after/>");
    drop(clone);
    drop(data);
    drop(data_clone);
    assert!(budget.used(Resource::Memory) > 0);
    drop(original);
    assert_eq!(budget.used(Resource::Memory), 0);
    assert_eq!(budget.used(Resource::Objects), 0);

    let (budget, cancel, context) = managed_context();
    let package = SourceBackedPackage::from_read_at_with_execution_context(
        Arc::new(VersionedSource::new(source_bytes())),
        litchi_opc::ReadLimits::default(),
        context,
    )
    .unwrap();
    let candidate = package.prepare_topology(SourceTopologyPlan::new()).unwrap();
    let part = candidate.main_document_part().unwrap();
    cancel.cancel();
    assert!(part.data().is_err());
    assert!(part.source_xml().is_err());
    assert!(candidate.iter_parts().is_err());
    drop(part);
    drop(candidate);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn owned_missing_lookup_rechecks_source_before_reporting_absence() {
    let source = Arc::new(VersionedSource::new(source_bytes()));
    let candidate = open(Arc::clone(&source))
        .prepare_topology(SourceTopologyPlan::new())
        .unwrap();
    source.checks_before_change.store(2, Ordering::SeqCst);
    assert!(matches!(
        candidate.part(&pack("/missing.xml")),
        Err(OpcError::SourceChanged { .. })
    ));
    assert!(candidate.check_source().is_err());
}

#[test]
fn materialization_uses_effective_graph_and_outlives_candidate() {
    let package = open(Arc::new(VersionedSource::new(source_bytes())));
    let mut plan = mutation_plan();
    plan.try_remove_relationship(pack(DOCUMENT), "rId1")
        .unwrap();
    plan.try_remove_part(pack(UNTOUCHED)).unwrap();
    plan.try_replace_internal_relationship(pack("/"), "rId1", rt::OFFICE_DOCUMENT, pack(ADDED))
        .unwrap();
    let candidate = package.prepare_topology(plan).unwrap();
    let retained = candidate.part(&pack(DOCUMENT)).unwrap().data().unwrap();
    let owned = candidate.to_opc_package().unwrap();
    assert!(owned.get_part(&pack(UNTOUCHED)).is_err());
    let main = owned.get_part(&pack(DOCUMENT)).unwrap();
    assert_eq!(main.blob(), b"<after/>");
    assert_eq!(main.blob().as_ptr(), retained.as_bytes().as_ptr());
    assert!(main.rels().get("rId1").is_none());
    assert!(main.rels().get("rIdAdded").is_some());
    let added = owned.get_part(&pack(ADDED)).unwrap();
    assert_eq!(added.content_type(), "application/example+xml");
    assert_eq!(added.blob(), b"<added/>");
    assert!(owned.rels().get("rIdAdded").is_some());
    assert_eq!(
        owned.rels().get("rId1").unwrap().target_ref(),
        "custom/added.xml"
    );
    drop(candidate);
    drop(retained);
    let bytes = PackageWriter::to_bytes(&owned).unwrap();
    let reopened = SourceBackedPackage::from_vec(bytes).unwrap();
    assert_eq!(
        reopened.main_document_part().unwrap().partname(),
        &pack(ADDED)
    );
    assert_eq!(
        reopened
            .part(&pack(DOCUMENT))
            .unwrap()
            .data()
            .unwrap()
            .as_bytes(),
        b"<after/>"
    );
}

#[test]
fn materialization_noop_and_changed_preserve_unchanged_parts() {
    for changed in [false, true] {
        let source = Arc::new(VersionedSource::new(source_bytes()));
        let candidate = open(Arc::clone(&source))
            .prepare_topology(if changed {
                mutation_plan()
            } else {
                SourceTopologyPlan::new()
            })
            .unwrap();
        let owned = candidate.to_opc_package().unwrap();
        assert_eq!(
            owned.get_part(&pack(UNTOUCHED)).unwrap().blob(),
            b"<untouched source='yes'/>"
        );
        assert_eq!(
            owned.get_part(&pack(DOCUMENT)).unwrap().blob(),
            if changed {
                b"<after/>".as_slice()
            } else {
                b"<before/>".as_slice()
            }
        );
        source.bump();
        assert!(matches!(
            candidate.to_opc_package(),
            Err(OpcError::SourceChanged { .. })
        ));
        assert_eq!(
            owned.get_part(&pack(UNTOUCHED)).unwrap().blob(),
            b"<untouched source='yes'/>"
        );
    }
}

#[test]
fn materialization_refuses_managed_candidates_before_cold_reads() {
    for changed in [false, true] {
        let source = Arc::new(VersionedSource::new(source_bytes()));
        let (_budget, cancel, context) = managed_context();
        let package = SourceBackedPackage::from_read_at_with_execution_context(
            source.clone(),
            litchi_opc::ReadLimits::default(),
            context,
        )
        .unwrap();
        let candidate = package
            .prepare_topology(if changed {
                mutation_plan()
            } else {
                SourceTopologyPlan::new()
            })
            .unwrap();
        let before = source.read_count();
        assert!(matches!(
            candidate.to_opc_package(),
            Err(OpcError::ManagedPackageMaterialization)
        ));
        assert_eq!(before, source.read_count());
        cancel.cancel();
        assert!(matches!(
            candidate.to_opc_package(),
            Err(OpcError::Cancelled)
        ));
    }
}

#[test]
fn materialization_rechecks_transfer_sources() {
    let donor_source = Arc::new(VersionedSource::new(source_bytes()));
    let donor = open(donor_source.clone());
    let (_, token) = donor
        .part(&pack(UNTOUCHED))
        .unwrap()
        .data_and_authorize_precompressed()
        .unwrap();
    let mut plan = SourceTopologyPlan::new();
    plan.try_add_precompressed_part(pack(ADDED), "application/xml", token)
        .unwrap();
    let candidate = open(Arc::new(VersionedSource::new(source_bytes())))
        .prepare_topology(plan)
        .unwrap();
    let owned = candidate.to_opc_package().unwrap();
    assert_eq!(
        owned.get_part(&pack(ADDED)).unwrap().blob(),
        b"<untouched source='yes'/>"
    );
    donor_source.bump();
    assert!(matches!(
        candidate.to_opc_package(),
        Err(OpcError::SourceChanged { .. })
    ));
}

#[test]
fn materialization_source_change_wins_over_failed_cold_read() {
    for changed in [false, true] {
        let source = Arc::new(VersionedSource::new(source_bytes()));
        let candidate = open(source.clone())
            .prepare_topology(if changed {
                mutation_plan()
            } else {
                SourceTopologyPlan::new()
            })
            .unwrap();
        source.fail_and_change_on_read.store(true, Ordering::SeqCst);
        assert!(matches!(
            candidate.to_opc_package(),
            Err(OpcError::SourceChanged { .. })
        ));
    }
}

#[test]
fn materialization_refuses_managed_donor_on_unmanaged_base() {
    for xml_transfer in [false, true] {
        let donor_source = Arc::new(VersionedSource::new(source_bytes()));
        let (_budget, cancel, context) = managed_context();
        let donor = SourceBackedPackage::from_read_at_with_execution_context(
            donor_source.clone(),
            litchi_opc::ReadLimits::default(),
            context,
        )
        .unwrap();
        let mut plan = SourceTopologyPlan::new();
        if xml_transfer {
            let xml = donor.part(&pack(UNTOUCHED)).unwrap().source_xml().unwrap();
            plan.try_add_source_xml_part(pack(ADDED), xml).unwrap();
        } else {
            let (_, token) = donor
                .part(&pack(UNTOUCHED))
                .unwrap()
                .data_and_authorize_precompressed()
                .unwrap();
            plan.try_add_precompressed_part(pack(ADDED), "application/xml", token)
                .unwrap();
        }
        let base_source = Arc::new(VersionedSource::new(source_bytes()));
        let candidate = open(base_source.clone()).prepare_topology(plan).unwrap();
        let before = (base_source.read_count(), donor_source.read_count());
        assert!(matches!(
            candidate.to_opc_package(),
            Err(OpcError::ManagedPackageMaterialization)
        ));
        assert_eq!(
            before,
            (base_source.read_count(), donor_source.read_count())
        );
        cancel.cancel();
        assert!(matches!(
            candidate.to_opc_package(),
            Err(OpcError::Cancelled)
        ));
    }
}
