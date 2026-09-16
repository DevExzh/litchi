#![allow(
    clippy::expect_used,
    clippy::indexing_slicing,
    clippy::unwrap_used,
    reason = "focused source-publication assertions intentionally panic on fixture errors"
)]

//! Public API coverage for source-preserving XML publication.
//!
//! These tests keep the source bytes deliberately formatted and exercise the
//! boundary between exact source bytes and compact authored fragments.

use std::num::{NonZeroU64, NonZeroUsize};
use std::sync::Arc;

use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits, OwnedSource, Resource,
};
use litchi_opc::{
    AuthoredXmlFragment, OpcError, PackURI, ReadLimits, ReadResource, SourceBackedPackage,
    SourceTopologyPlan, authored_xml_requires_source_proof,
};
use soapberry_zip::office::StreamingArchiveWriter;

const CONTENT_TYPES_NS: &str = "http://schemas.openxmlformats.org/package/2006/content-types";
const RELATIONSHIPS_NS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const OFFICE_DOCUMENT_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
const DOCUMENT_URI: &str = "/word/document.xml";

const FORMATTED_DOCUMENT: &[u8] = br#"<?xml version="1.0" encoding="UTF-8"?>
<document>
  <before attr="one"/>
  <!--keep-->
  <nested>
    <leaf/>
  </nested>
</document>
"#;

fn document_uri() -> PackURI {
    PackURI::new(DOCUMENT_URI).expect("fixture URI must be valid")
}

fn archive_bytes(document: &[u8]) -> Vec<u8> {
    let content_types = format!(
        r#"<Types xmlns="{CONTENT_TYPES_NS}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/></Types>"#
    );
    let root_relationships = format!(
        r#"<Relationships xmlns="{RELATIONSHIPS_NS}"><Relationship Id="rId1" Type="{OFFICE_DOCUMENT_REL}" Target="word/document.xml"/></Relationships>"#
    );

    let mut writer = StreamingArchiveWriter::new();
    writer
        .write_stored("[Content_Types].xml", content_types.as_bytes())
        .expect("content types fixture must be writable");
    writer
        .write_stored("_rels/.rels", root_relationships.as_bytes())
        .expect("root relationships fixture must be writable");
    writer
        .write_stored("word/document.xml", document)
        .expect("document fixture must be writable");
    writer
        .finish_to_bytes()
        .expect("fixture archive must finish")
}

fn open(document: &[u8]) -> SourceBackedPackage {
    SourceBackedPackage::from_vec(archive_bytes(document)).expect("fixture package must open")
}

fn open_with_limits(document: &[u8], limits: ReadLimits) -> SourceBackedPackage {
    SourceBackedPackage::from_vec_with_limits(archive_bytes(document), limits)
        .expect("fixture package must open with limits")
}

fn insertion_range(bytes: &[u8]) -> usize {
    bytes
        .windows(b"</document>".len())
        .position(|window| window == b"</document>")
        .expect("formatted fixture must have a document close tag")
}

fn before_range(bytes: &[u8]) -> std::ops::Range<usize> {
    let start = bytes
        .windows(b"<before attr=\"one\"/>".len())
        .position(|window| window == b"<before attr=\"one\"/>")
        .expect("formatted fixture must have a before element");
    start..start + b"<before attr=\"one\"/>".len()
}

fn managed_context(memory: u64) -> (Budget, CancellationSource, ExecutionContext) {
    let budget = Budget::root(
        "source-xml-publication-test",
        Limits::new(memory, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX),
    );
    let (cancellation_source, cancellation) = CancellationSource::pair();
    let execution_limits = ExecutionLimits::new(
        NonZeroUsize::new(1).expect("one worker limit is nonzero"),
        NonZeroUsize::new(1).expect("one operation limit is nonzero"),
        NonZeroU64::new(memory).expect("memory limit is nonzero"),
        0,
    )
    .expect("execution limits must be valid");
    let context = ExecutionContext::new(budget.clone(), cancellation, execution_limits);
    (budget, cancellation_source, context)
}

#[test]
fn source_xml_publication_preserves_formatted_bytes_and_emits_checked_fragment() {
    let package = open(FORMATTED_DOCUMENT);
    let part = package
        .part(&document_uri())
        .expect("document part must exist");
    let source = part.source_xml().expect("formatted XML is source-valid");
    let original = source.bytes().to_vec();
    let insertion = insertion_range(&original);
    let proof = source
        .checked_range(insertion..insertion, &[])
        .expect("close-tag insertion point must be source-authorized");

    let mut publication = source
        .into_publication()
        .expect("source XML must start an edit transaction");
    publication
        .replace(
            proof,
            AuthoredXmlFragment::markup(b"<inserted/>".to_vec())
                .expect("compact authored fragment must pass its audit"),
        )
        .expect("checked insertion must be accepted");
    let edited = publication.finish().expect("source splice must remain XML");

    let expected = [
        &original[..insertion],
        b"<inserted/>".as_slice(),
        &original[insertion..],
    ]
    .concat();
    assert_eq!(edited.bytes(), expected.as_slice());
    assert_eq!(
        &edited.bytes()[..insertion],
        &original[..insertion],
        "source prefix must remain byte exact"
    );
    assert_eq!(
        &edited.bytes()[insertion + b"<inserted/>".len()..],
        &original[insertion..],
        "source suffix must remain byte exact"
    );

    let mut plan = SourceTopologyPlan::new();
    plan.try_replace_source_xml_part(document_uri(), edited)
        .expect("source XML token must be accepted by the topology plan");
    let mut output = Vec::new();
    package
        .write_topology_to_stream(&mut output, plan)
        .expect("source XML replacement must publish");
    let published = SourceBackedPackage::from_vec(output).expect("published package must reopen");
    let published_bytes = published
        .part(&document_uri())
        .expect("published document must exist")
        .data()
        .expect("published document must be readable");
    assert_eq!(published_bytes.as_bytes(), expected.as_slice());
}

#[test]
fn the_authored_classifier_still_separates_compact_from_formatted_xml() {
    // Change 0657 stopped making compactness a publication refusal on the
    // source-backed replacement route: a replacement there is a splice of the
    // Part's own source, so it carries the producer's formatting and the
    // contract is now a property of this library's serializers, asserted by
    // their own tests. The classifier that tells the two apart is unchanged,
    // and so is its refusal of malformed input.
    let uri = document_uri();
    assert!(
        !authored_xml_requires_source_proof(
            &uri,
            "application/xml",
            b"<document><child/></document>"
        )
        .expect("compact authored XML must be classified")
    );
    assert!(
        authored_xml_requires_source_proof(&uri, "application/xml", FORMATTED_DOCUMENT)
            .expect("formatted XML should be a source-proof candidate")
    );
    assert!(authored_xml_requires_source_proof(&uri, "application/xml", b"<document>").is_err());

    // The same formatted replacement now publishes, and every other member
    // stays byte-exact.
    let package = open(b"<document/>");
    let mut plan = SourceTopologyPlan::new();
    plan.try_replace_part(uri.clone(), FORMATTED_DOCUMENT.to_vec())
        .expect("topology plan may stage a replacement");
    let mut output = Vec::new();
    package
        .write_topology_to_stream(&mut output, plan)
        .expect("a formatted replacement publishes since change 0657");
    assert!(!output.is_empty());

    // Malformed input is still refused, and still emits no archive.
    let package = open(b"<document/>");
    let mut plan = SourceTopologyPlan::new();
    plan.try_replace_part(uri, b"<document>".to_vec())
        .expect("topology plan may stage a replacement");
    let mut output = Vec::new();
    let error = package
        .write_topology_to_stream(&mut output, plan)
        .expect_err("an unterminated replacement must be refused");
    assert!(matches!(error, OpcError::XmlPublication { .. }));
    assert!(output.is_empty(), "refused XML must emit no archive");
}

#[test]
fn identical_source_splice_retains_original_authority_and_exact_output() {
    let package = open(FORMATTED_DOCUMENT);
    let source = package.part(&document_uri()).unwrap().source_xml().unwrap();
    let range = before_range(source.bytes());
    let bytes = source.bytes()[range.clone()].to_vec();
    let proof = source.checked_range(range, &bytes).unwrap();
    let mut publication = source.into_publication().unwrap();
    publication
        .replace(proof, AuthoredXmlFragment::markup(bytes).unwrap())
        .unwrap();
    let finished = publication.finish().unwrap();
    assert_eq!(finished.bytes(), FORMATTED_DOCUMENT);
    assert!(finished.clone().into_publication().is_ok());
    let mut plan = SourceTopologyPlan::new();
    plan.try_replace_source_xml_part(document_uri(), finished)
        .unwrap();
    let mut output = Vec::new();
    package.write_topology_to_stream(&mut output, plan).unwrap();
    assert_eq!(output, archive_bytes(FORMATTED_DOCUMENT));
}

#[test]
fn source_provenance_rejects_foreign_lineage_and_stale_ranges_without_output() {
    let first = open(FORMATTED_DOCUMENT);
    let second = open(FORMATTED_DOCUMENT);
    let uri = document_uri();
    let first_source = first
        .part(&uri)
        .expect("first document must exist")
        .source_xml()
        .expect("first source XML must be valid");
    let second_source = second
        .part(&uri)
        .expect("second document must exist")
        .source_xml()
        .expect("second source XML must be valid");
    let first_bytes = first_source.bytes().to_vec();
    let range = before_range(&first_bytes);

    assert!(
        first_source
            .checked_range(range.clone(), b"<before attr=\"two\"/>")
            .is_err()
    );
    let foreign_proof = first_source
        .checked_range(range, &first_bytes[before_range(&first_bytes)])
        .expect("first source range must be valid");
    let mut publication = second_source
        .into_publication()
        .expect("second source must start an edit transaction");
    let error = publication
        .replace(
            foreign_proof,
            AuthoredXmlFragment::markup(b"<foreign/>".to_vec())
                .expect("compact foreign fragment must pass its audit"),
        )
        .expect_err("a source proof cannot cross package lineage");
    assert!(matches!(
        error,
        OpcError::SourceBackedOverlayUnavailable { .. }
    ));

    let foreign_source = first
        .part(&uri)
        .expect("first document must exist")
        .source_xml()
        .expect("first source XML must be valid");
    let mut plan = SourceTopologyPlan::new();
    plan.try_replace_source_xml_part(uri, foreign_source)
        .expect("lineage is checked at publication, after plan staging");
    let mut output = Vec::new();
    let error = second
        .write_topology_to_stream(&mut output, plan)
        .expect_err("a foreign source token must be refused by the destination");
    assert!(matches!(
        error,
        OpcError::SourceBackedOverlayUnavailable { .. }
    ));
    assert!(output.is_empty(), "lineage refusal must emit no archive");
}

#[test]
fn source_splice_rejects_reedit_and_overlapping_ranges_without_second_output() {
    let package = open(FORMATTED_DOCUMENT);
    let uri = document_uri();
    let source = package
        .part(&uri)
        .expect("document must exist")
        .source_xml()
        .expect("document source must be valid");
    let original = source.bytes().to_vec();
    let first_range = before_range(&original);
    let overlapping_range = first_range.start + 1..first_range.end;
    let first_proof = source
        .checked_range(first_range.clone(), &original[first_range.clone()])
        .expect("first replacement range must be valid");
    let overlapping_proof = source
        .checked_range(
            overlapping_range.clone(),
            &original[overlapping_range.clone()],
        )
        .expect("overlapping range must be source-authorized before staging");

    let mut publication = source
        .into_publication()
        .expect("source must start an edit transaction");
    publication
        .replace(
            first_proof,
            AuthoredXmlFragment::markup(b"<first/>".to_vec())
                .expect("first fragment must be compact"),
        )
        .expect("first replacement must be accepted");
    let error = publication
        .replace(
            overlapping_proof,
            AuthoredXmlFragment::markup(b"<second/>".to_vec())
                .expect("second fragment must be compact"),
        )
        .expect_err("overlapping source ranges must be refused");
    assert!(matches!(
        error,
        OpcError::SourceBackedOverlayUnavailable { .. }
    ));

    let edited = publication
        .finish()
        .expect("the accepted edit must remain publishable");
    assert!(
        edited
            .bytes()
            .windows(b"<first/>".len())
            .any(|window| window == b"<first/>")
    );
    assert!(
        !edited
            .bytes()
            .windows(b"<second/>".len())
            .any(|window| window == b"<second/>")
    );

    let mut plan = SourceTopologyPlan::new();
    plan.try_replace_source_xml_part(uri, edited.clone())
        .expect("the accepted edit must be stageable");
    let mut output = Vec::new();
    package
        .write_topology_to_stream(&mut output, plan)
        .expect("the accepted edit must publish");
    let published = SourceBackedPackage::from_vec(output).expect("published package must reopen");
    let published_bytes = published
        .part(&document_uri())
        .expect("published document must exist")
        .data()
        .expect("published document must be readable");
    assert!(
        published_bytes
            .as_bytes()
            .windows(b"<first/>".len())
            .any(|window| window == b"<first/>")
    );
    assert!(
        !published_bytes
            .as_bytes()
            .windows(b"<second/>".len())
            .any(|window| window == b"<second/>")
    );
    assert!(
        edited.into_publication().is_err(),
        "a derived source cannot be reedited"
    );
}

#[test]
fn destination_depth_and_event_limits_refuse_source_xml_additions() {
    let source = open(FORMATTED_DOCUMENT);
    let uri = document_uri();
    let source_xml = source
        .part(&uri)
        .expect("source document must exist")
        .source_xml()
        .expect("source document must be valid");

    let depth_limits = ReadLimits::builder()
        .max_xml_depth(1)
        .expect("depth limit must be valid")
        .build()
        .expect("depth profile must be valid");
    let depth_destination = open_with_limits(b"<destination/>", depth_limits);
    let mut depth_plan = SourceTopologyPlan::new();
    depth_plan
        .try_add_source_xml_part(
            PackURI::new("/custom/copied.xml").unwrap(),
            source_xml.clone(),
        )
        .expect("source addition must stage before destination validation");
    let mut depth_output = Vec::new();
    assert!(matches!(
        depth_destination.write_topology_to_stream(&mut depth_output, depth_plan),
        Err(OpcError::ReadLimit {
            resource: ReadResource::XmlDepth,
            ..
        })
    ));
    assert!(depth_output.is_empty());

    let event_limits = ReadLimits::builder()
        .max_xml_events(8)
        .expect("event limit must be valid")
        .build()
        .expect("event profile must be valid");
    let event_destination = open_with_limits(b"<destination/>", event_limits);
    let mut event_plan = SourceTopologyPlan::new();
    event_plan
        .try_add_source_xml_part(PackURI::new("/custom/copied.xml").unwrap(), source_xml)
        .expect("source addition must stage before destination validation");
    let mut event_output = Vec::new();
    assert!(matches!(
        event_destination.write_topology_to_stream(&mut event_output, event_plan),
        Err(OpcError::ReadLimit {
            resource: ReadResource::XmlEvents,
            ..
        })
    ));
    assert!(event_output.is_empty());
}

#[test]
fn managed_source_xml_honors_cancellation_and_releases_memory_gauge() {
    let memory = 8 * 1024 * 1024;
    let (budget, cancellation_source, context) = managed_context(memory);
    let package = SourceBackedPackage::from_read_at_with_execution_context(
        Arc::new(OwnedSource::new(archive_bytes(FORMATTED_DOCUMENT))),
        ReadLimits::default(),
        context,
    )
    .expect("managed fixture package must open");
    cancellation_source.cancel();
    let cancelled = package
        .part(&document_uri())
        .expect("part lookup does not read a payload")
        .source_xml()
        .expect_err("source XML capture must honor cancellation");
    assert!(matches!(cancelled, OpcError::Cancelled));
    drop(package);
    assert_eq!(budget.used(Resource::Memory), 0);
    assert_eq!(budget.used(Resource::Objects), 0);

    let (budget, _cancellation_source, context) = managed_context(memory);
    {
        let package = SourceBackedPackage::from_read_at_with_execution_context(
            Arc::new(OwnedSource::new(archive_bytes(FORMATTED_DOCUMENT))),
            ReadLimits::default(),
            context,
        )
        .expect("managed fixture package must open");
        let source = package
            .part(&document_uri())
            .expect("document must exist")
            .source_xml()
            .expect("source XML must be capturable");
        let before = budget.used(Resource::Memory);
        let clones = vec![source.clone(); 16];
        assert_eq!(budget.used(Resource::Memory), before);
        let independent = package.part(&document_uri()).unwrap().source_xml().unwrap();
        assert!(budget.used(Resource::Memory) > before);
        drop(independent);
        drop(clones);
        assert_eq!(budget.used(Resource::Memory), before);
        let insertion = insertion_range(source.bytes());
        let proof = source
            .checked_range(insertion..insertion, &[])
            .expect("insertion proof must be valid");
        let mut publication = source.into_publication().expect("publication must start");
        publication
            .replace(
                proof,
                AuthoredXmlFragment::markup(b"<managed/>".to_vec()).unwrap(),
            )
            .expect("managed splice must be accepted");
        let edited = publication.finish().expect("managed splice must finish");
        assert!(budget.used(Resource::Memory) > before);
        drop(edited);
    }
    assert_eq!(budget.used(Resource::Memory), 0);
    assert_eq!(budget.used(Resource::Objects), 0);
}

#[test]
fn managed_source_xml_refuses_retained_fragment_capacity_over_budget() {
    let memory = 1024 * 1024;
    let (budget, _cancellation_source, context) = managed_context(memory);
    {
        let package = SourceBackedPackage::from_read_at_with_execution_context(
            Arc::new(OwnedSource::new(archive_bytes(FORMATTED_DOCUMENT))),
            ReadLimits::default(),
            context,
        )
        .unwrap();
        let source = package.part(&document_uri()).unwrap().source_xml().unwrap();
        let insertion = insertion_range(source.bytes());
        let proof = source.checked_range(insertion..insertion, &[]).unwrap();
        let mut publication = source.into_publication().unwrap();
        let mut oversized = Vec::with_capacity(2 * 1024 * 1024);
        oversized.extend_from_slice(b"<added/>");
        let fragment = AuthoredXmlFragment::markup(oversized).unwrap();
        let error = publication.replace(proof, fragment).unwrap_err();
        assert!(matches!(
            error,
            OpcError::Execution(litchi_core::ExecutionError::ResourceLimit(limit))
                if limit.resource == Resource::Memory
        ));
        let unchanged = publication.finish().unwrap();
        assert_eq!(unchanged.bytes(), FORMATTED_DOCUMENT);
    }
    assert_eq!(budget.used(Resource::Memory), 0);
    assert_eq!(budget.used(Resource::Objects), 0);
}

/// A fixture with a second Part, so that a replacement of the first leaves an
/// untouched Part whose published bytes can be compared with the source's.
fn archive_bytes_with_sibling(document: &[u8], styles: &[u8]) -> Vec<u8> {
    let content_types = format!(
        r#"<Types xmlns="{CONTENT_TYPES_NS}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/></Types>"#
    );
    let root_relationships = format!(
        r#"<Relationships xmlns="{RELATIONSHIPS_NS}"><Relationship Id="rId1" Type="{OFFICE_DOCUMENT_REL}" Target="word/document.xml"/></Relationships>"#
    );
    let mut writer = StreamingArchiveWriter::new();
    writer
        .write_stored("[Content_Types].xml", content_types.as_bytes())
        .expect("content types fixture must be writable");
    writer
        .write_stored("_rels/.rels", root_relationships.as_bytes())
        .expect("root relationships fixture must be writable");
    writer
        .write_stored("word/document.xml", document)
        .expect("document fixture must be writable");
    writer
        .write_stored("word/styles.xml", styles)
        .expect("styles fixture must be writable");
    writer
        .finish_to_bytes()
        .expect("fixture archive must finish")
}

/// Change 0654, under change 0652 decision 2.
///
/// The original bytes of a replaced Part are audited for structure and finite
/// budgets, never for this repository's compact output contract. The
/// replacement keeps the authored contract, which
/// `the_authored_classifier_still_separates_compact_from_formatted_xml` above pins.
#[test]
fn noncompact_original_bytes_publish_and_leave_every_other_member_byte_exact() {
    for (label, original) in [
        (
            "declaration line ending and indentation",
            FORMATTED_DOCUMENT,
        ),
        (
            "attribute separation and whitespace before close",
            b"<document  a=\"1\"\n\tb = \"2\" ><child /></document >".as_slice(),
        ),
        (
            "plain-space run",
            b"<document> <child/></document>".as_slice(),
        ),
    ] {
        const SIBLING_STYLES: &[u8] = b"<styles>\n  <style id=\"a\" />\n</styles>";
        let source = archive_bytes_with_sibling(original, SIBLING_STYLES);
        let package = SourceBackedPackage::from_vec(source.clone())
            .unwrap_or_else(|error| panic!("{label}: fixture must open: {error:?}"));
        let mut output = Vec::new();
        package
            .write_part_overlay_to_stream(
                &mut output,
                &document_uri(),
                b"<document><replaced/></document>".to_vec(),
            )
            .unwrap_or_else(|error| {
                panic!("{label}: non-compact original must publish: {error:?}")
            });

        let before = SourceBackedPackage::from_vec(source).expect("source must reopen");
        let after = SourceBackedPackage::from_vec(output).expect("published package must reopen");
        let mut compared = 0_usize;
        for part in before.iter_parts() {
            let name = part.partname().clone();
            let published = after
                .part(&name)
                .unwrap_or_else(|error| panic!("{label}: {name} must survive: {error:?}"))
                .data()
                .expect("published payload");
            if name == document_uri() {
                assert_eq!(
                    published.as_bytes(),
                    b"<document><replaced/></document>",
                    "{label}: the replaced Part must carry the replacement"
                );
                continue;
            }
            assert_eq!(
                published.as_bytes(),
                part.data().expect("source payload").as_bytes(),
                "{label}: untouched Part {name} must be byte-identical"
            );
            compared += 1;
        }
        assert!(compared > 0, "{label}: at least one untouched Part");
    }
}

/// Every refusal that is not a compactness verdict still fires on the original
/// bytes, with the same `OpcError::XmlPublication` identity, before any byte
/// reaches the sink.
#[test]
fn original_bytes_keep_every_structural_doctype_encoding_and_limit_refusal() {
    for (label, original) in [
        ("unclosed document element", b"<document>".as_slice()),
        ("two document elements", b"<document/><other/>".as_slice()),
        ("unquoted attribute value", b"<document a=1/>".as_slice()),
        ("DOCTYPE", b"<!DOCTYPE document><document/>".as_slice()),
        ("invalid UTF-8", b"<document>\xff</document>".as_slice()),
    ] {
        let package = SourceBackedPackage::from_vec(archive_bytes(original))
            .unwrap_or_else(|error| panic!("{label}: fixture must open: {error:?}"));
        let mut output = Vec::new();
        let error = package
            .write_part_overlay_to_stream(
                &mut output,
                &document_uri(),
                b"<document><replaced/></document>".to_vec(),
            )
            .expect_err("a refused original must stay refused");
        assert!(
            matches!(error, OpcError::XmlPublication { ref part, .. } if part == "/word/document.xml"),
            "{label}: refusal must keep its identity, got {error:?}"
        );
        assert!(output.is_empty(), "{label}: a refusal emits no archive");
    }

    // The audit's own finite budgets are unchanged and are exercised at the
    // auditor in `crates/xml-minifier/tests/audit.rs`; `ReadLimits` does not
    // feed them, so there is nothing package-level to narrow here.
}

/// An exact byte no-op is still exact: it never reaches either audit, so a
/// Part whose source payload is malformed republishes unchanged.
#[test]
fn an_exact_no_op_still_precedes_both_audits() {
    let malformed = b"<document>".as_slice();
    let source = archive_bytes(malformed);
    let package = SourceBackedPackage::from_vec(source.clone()).expect("fixture must open");
    let mut output = Vec::new();
    package
        .write_part_overlay_to_stream(&mut output, &document_uri(), malformed.to_vec())
        .expect("an exact no-op must publish the source artifact byte for byte");
    assert_eq!(output, source, "an exact no-op copies the source artifact");
}
