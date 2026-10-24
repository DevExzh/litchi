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
    AuthoredXmlFragment, OpcError, PackURI, SourceBackedPackage, SourceLimitKind, SourceLimits,
    SourceTopologyPlan, authored_xml_requires_source_proof,
};
use litchi_opc::{SourceCacheLimits, SourceReadPolicy};
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

    let mut writer =
        StreamingArchiveWriter::new(soapberry_zip::office::StreamingArchiveLimits::default());
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
    SourceBackedPackage::from_vec(
        archive_bytes(document),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        None,
    )
    .expect("fixture package must open")
}

fn open_with_limits(document: &[u8], limits: SourceLimits) -> SourceBackedPackage {
    SourceBackedPackage::from_vec(
        archive_bytes(document),
        limits,
        SourceCacheLimits::default(),
        None,
    )
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
    let published = SourceBackedPackage::from_vec(
        output,
        SourceLimits::default(),
        SourceCacheLimits::default(),
        None,
    )
    .expect("published package must reopen");
    let published_bytes = published
        .part(&document_uri())
        .expect("published document must exist")
        .data()
        .expect("published document must be readable");
    assert_eq!(published_bytes.as_bytes(), expected.as_slice());
}

#[test]
fn empty_fragment_deletes_only_the_authorized_source_range() {
    let package = open(FORMATTED_DOCUMENT);
    let source = package.part(&document_uri()).unwrap().source_xml().unwrap();
    let range = before_range(source.bytes());
    let proof = source
        .checked_range(range.clone(), &source.bytes()[range.clone()])
        .unwrap();
    let expected = [
        &FORMATTED_DOCUMENT[..range.start],
        &FORMATTED_DOCUMENT[range.end..],
    ]
    .concat();
    let mut publication = source.into_publication().unwrap();
    publication
        .replace(proof, AuthoredXmlFragment::empty())
        .unwrap();
    let edited = publication.finish().unwrap();
    assert_eq!(edited.bytes(), expected);

    let mut plan = SourceTopologyPlan::new();
    plan.try_replace_source_xml_part(document_uri(), edited)
        .unwrap();
    let mut output = Vec::new();
    package.write_topology_to_stream(&mut output, plan).unwrap();
    let reopened = SourceBackedPackage::from_vec(
        output,
        SourceLimits::default(),
        SourceCacheLimits::default(),
        None,
    )
    .unwrap();
    assert_eq!(
        reopened
            .part(&document_uri())
            .unwrap()
            .data()
            .unwrap()
            .as_bytes(),
        expected
    );
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
    let published = SourceBackedPackage::from_vec(
        output,
        SourceLimits::default(),
        SourceCacheLimits::default(),
        None,
    )
    .expect("published package must reopen");
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

    let depth_limits = SourceLimits::builder()
        .max_depth(1)
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
        Err(OpcError::Limit {
            kind: SourceLimitKind::Depth,
            ..
        })
    ));
    assert!(depth_output.is_empty());

    let event_limits = SourceLimits::builder()
        .max_events(8)
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
        Err(OpcError::Limit {
            kind: SourceLimitKind::Events,
            ..
        })
    ));
    assert!(event_output.is_empty());
}

#[test]
fn managed_source_xml_honors_cancellation_and_releases_memory_gauge() {
    let memory = 8 * 1024 * 1024;
    let (budget, cancellation_source, context) = managed_context(memory);
    let package = SourceBackedPackage::from_read_at(
        Arc::new(OwnedSource::new(archive_bytes(FORMATTED_DOCUMENT))),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        Some(context),
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
        let package = SourceBackedPackage::from_read_at(
            Arc::new(OwnedSource::new(archive_bytes(FORMATTED_DOCUMENT))),
            SourceLimits::default(),
            SourceCacheLimits::default(),
            SourceReadPolicy::exact(),
            Some(context),
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
        let package = SourceBackedPackage::from_read_at(
            Arc::new(OwnedSource::new(archive_bytes(FORMATTED_DOCUMENT))),
            SourceLimits::default(),
            SourceCacheLimits::default(),
            SourceReadPolicy::exact(),
            Some(context),
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
    let mut writer =
        StreamingArchiveWriter::new(soapberry_zip::office::StreamingArchiveLimits::default());
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
        let package = SourceBackedPackage::from_vec(
            source.clone(),
            SourceLimits::default(),
            SourceCacheLimits::default(),
            None,
        )
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

        let before = SourceBackedPackage::from_vec(
            source,
            SourceLimits::default(),
            SourceCacheLimits::default(),
            None,
        )
        .expect("source must reopen");
        let after = SourceBackedPackage::from_vec(
            output,
            SourceLimits::default(),
            SourceCacheLimits::default(),
            None,
        )
        .expect("published package must reopen");
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
        let package = SourceBackedPackage::from_vec(
            archive_bytes(original),
            SourceLimits::default(),
            SourceCacheLimits::default(),
            None,
        )
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
    // auditor in `crates/xml-minifier/tests/audit.rs`; `SourceLimits` does not
    // feed them, so there is nothing package-level to narrow here.
}

/// An exact byte no-op is still exact: it never reaches either audit, so a
/// Part whose source payload is malformed republishes unchanged.
#[test]
fn an_exact_no_op_still_precedes_both_audits() {
    let malformed = b"<document>".as_slice();
    let source = archive_bytes(malformed);
    let package = SourceBackedPackage::from_vec(
        source.clone(),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        None,
    )
    .expect("fixture must open");
    let mut output = Vec::new();
    package
        .write_part_overlay_to_stream(&mut output, &document_uri(), malformed.to_vec())
        .expect("an exact no-op must publish the source artifact byte for byte");
    assert_eq!(output, source, "an exact no-op copies the source artifact");
}

#[test]
fn eager_replacement_preserves_a_utf8_bom_and_still_refuses_a_doctype_before_output() {
    use litchi_opc::{OpcPackage, PackageWriter};

    let original = b"<?xml version=\"1.0\"?><document><before/></document>";
    let marked = b"\xef\xbb\xbf<?xml version=\"1.0\"?><document><after/></document>";
    let mut package =
        OpcPackage::from_vec(archive_bytes(original), SourceLimits::default()).unwrap();
    package
        .get_part_mut(&document_uri())
        .unwrap()
        .set_blob(marked.to_vec());
    let published = PackageWriter::to_bytes(&package).unwrap();
    let reopened = OpcPackage::from_vec(published, SourceLimits::default()).unwrap();
    assert_eq!(reopened.get_part(&document_uri()).unwrap().blob(), marked);

    package
        .get_part_mut(&document_uri())
        .unwrap()
        .set_blob(b"\xef\xbb\xbf<!DOCTYPE document><document/>".to_vec());
    let mut output = Vec::new();
    assert!(PackageWriter::write_to(&mut output, &package).is_err());
    assert!(output.is_empty());
}

#[test]
fn marked_xml_members_in_real_packages_pass_the_source_audit_unchanged() {
    use litchi_opc::phys_pkg::PhysPkgReader;

    let fixtures = [
        "libreoffice-core/sc/qa/unit/data/xlsx/tdf167689_xmlMaps_and_xmlColumnPr.xlsx",
        "libreoffice-core/sw/qa/writerfilter/dmapper/data/alt-chunk-header.docx",
        "poi/test-data/slideshow/bug65551.pptx",
    ];
    let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../test-data");
    let mut marked = 0;
    for fixture in fixtures {
        let bytes = std::fs::read(root.join(fixture)).unwrap();
        let archive = PhysPkgReader::new(&bytes, SourceLimits::default()).unwrap();
        for name in archive.member_names().unwrap() {
            if !name.ends_with(".xml") && !name.ends_with(".rels") {
                continue;
            }
            let payload = archive.read_member(&name).unwrap();
            if !payload.starts_with(b"\xef\xbb\xbf") {
                continue;
            }
            let report = xml_minifier::audit::verify_source(
                &payload,
                xml_minifier::audit::Limits::default(),
            )
            .unwrap_or_else(|error| panic!("{fixture}:{name}: {error}"));
            assert_eq!(report.bytes(), payload.len());
            marked += 1;
        }
    }
    assert_eq!(marked, 16);
}

fn append_revision(
    source: &litchi_opc::SourceXmlPart,
    fragment: &[u8],
) -> litchi_opc::SourceXmlPart {
    let revision = source.begin_revision().unwrap();
    let insertion = insertion_range(revision.bytes());
    let proof = revision.checked_range(insertion..insertion, &[]).unwrap();
    let mut edit = revision.into_publication().unwrap();
    edit.replace(
        proof,
        AuthoredXmlFragment::markup(fragment.to_vec()).unwrap(),
    )
    .unwrap();
    edit.finish().unwrap()
}

#[test]
fn logical_xml_revisions_chain_without_reopening_the_physical_source() {
    let package = open(FORMATTED_DOCUMENT);
    let original = package.part(&document_uri()).unwrap().source_xml().unwrap();
    let first = append_revision(&original, b"<first/>");
    let second = append_revision(&first, b"<second/>");
    let third = append_revision(&second, b"<third/>");
    assert_eq!(original.bytes(), FORMATTED_DOCUMENT);
    assert!(
        !first
            .bytes()
            .windows(b"<second/>".len())
            .any(|bytes| bytes == b"<second/>")
    );
    let insertion = insertion_range(FORMATTED_DOCUMENT);
    let expected = [
        &FORMATTED_DOCUMENT[..insertion],
        b"<first/><second/><third/>",
        &FORMATTED_DOCUMENT[insertion..],
    ]
    .concat();
    assert_eq!(third.bytes(), expected);
    assert!(third.checked_range(0..0, &[]).is_err());
    assert!(third.clone().into_publication().is_err());
    let mut plan = SourceTopologyPlan::new();
    plan.try_replace_source_xml_part(document_uri(), third.clone())
        .unwrap();
    let prepared = package.clone().prepare_topology(plan).unwrap();
    prepared
        .with_candidate(|candidate| {
            assert_eq!(
                candidate.part(&document_uri())?.data()?.as_bytes(),
                expected
            );
            Ok(())
        })
        .unwrap();
    let mut output = Vec::new();
    prepared.publish_to_stream(&mut output).unwrap();
    let reopened = SourceBackedPackage::from_vec(
        output,
        SourceLimits::default(),
        SourceCacheLimits::default(),
        None,
    )
    .unwrap();
    assert_eq!(
        reopened
            .part(&document_uri())
            .unwrap()
            .data()
            .unwrap()
            .as_bytes(),
        expected
    );
    // Reversing the logical additions restores the physical no-op path.
    let revision = third.begin_revision().unwrap();
    let additions = b"<first/><second/><third/>";
    let proof = revision
        .checked_range(insertion..insertion + additions.len(), additions)
        .unwrap();
    let mut inverse = revision.into_publication().unwrap();
    inverse
        .replace(proof, AuthoredXmlFragment::empty())
        .unwrap();
    let restored = inverse.finish().unwrap();
    assert_eq!(restored.bytes(), original.bytes());
    let mut inverse_plan = SourceTopologyPlan::new();
    inverse_plan
        .try_replace_source_xml_part(document_uri(), restored)
        .unwrap();
    let mut restored_archive = Vec::new();
    package
        .write_topology_to_stream(&mut restored_archive, inverse_plan)
        .unwrap();
    assert_eq!(restored_archive, archive_bytes(FORMATTED_DOCUMENT));
    // Equal physical bytes from another open do not supply the same authority.
    let mut foreign_plan = SourceTopologyPlan::new();
    foreign_plan
        .try_replace_source_xml_part(document_uri(), third)
        .unwrap();
    let mut rejected = Vec::new();
    assert!(
        open(FORMATTED_DOCUMENT)
            .write_topology_to_stream(&mut rejected, foreign_plan)
            .is_err()
    );
    assert!(rejected.is_empty());
}

#[test]
fn range_proofs_reject_sibling_and_stale_logical_revisions_with_equal_bytes() {
    let package = open(FORMATTED_DOCUMENT);
    let original = package.part(&document_uri()).unwrap().source_xml().unwrap();
    let left = original.begin_revision().unwrap();
    let right = original.begin_revision().unwrap();
    assert_eq!(left.bytes(), right.bytes());
    let insertion = insertion_range(left.bytes());
    let proof = left.checked_range(insertion..insertion, &[]).unwrap();
    let mut sibling = right.into_publication().unwrap();
    assert!(
        sibling
            .replace(
                proof,
                AuthoredXmlFragment::markup(b"<wrong/>".to_vec()).unwrap()
            )
            .is_err()
    );
    assert_eq!(sibling.finish().unwrap().bytes(), FORMATTED_DOCUMENT);
    let stale = left.checked_range(0..0, &[]).unwrap();
    let changed = append_revision(&left, b"<changed/>");
    let next = changed.begin_revision().unwrap();
    let mut next_edit = next.into_publication().unwrap();
    assert!(
        next_edit
            .replace(
                stale,
                AuthoredXmlFragment::markup(b"<wrong/>".to_vec()).unwrap()
            )
            .is_err()
    );
    assert_eq!(next_edit.finish().unwrap().bytes(), changed.bytes());
    // A clone is the same revision and accepts its exact proof.
    let cloned = left.clone();
    let proof = left.checked_range(insertion..insertion, &[]).unwrap();
    let mut accepted = cloned.into_publication().unwrap();
    accepted
        .replace(
            proof,
            AuthoredXmlFragment::markup(b"<accepted/>".to_vec()).unwrap(),
        )
        .unwrap();
    assert!(
        accepted
            .finish()
            .unwrap()
            .bytes()
            .windows(11)
            .any(|bytes| bytes == b"<accepted/>")
    );
    // Physical-source proofs are not logical-revision proofs either.
    let proof = original.checked_range(insertion..insertion, &[]).unwrap();
    let mut edit = left.into_publication().unwrap();
    assert!(edit.replace(proof, AuthoredXmlFragment::empty()).is_err());
}

#[test]
fn revision_hint_cannot_replace_a_physical_source_capture() {
    let package = open(FORMATTED_DOCUMENT);
    let part = package.part(&document_uri()).unwrap();
    let original = part.source_xml().unwrap();
    let logical = original.begin_revision().unwrap();
    let captured = part.source_xml_with_hint(&logical).unwrap();
    let insertion = insertion_range(original.bytes());
    let physical_proof = original.checked_range(insertion..insertion, &[]).unwrap();
    let mut edit = captured.into_publication().unwrap();
    edit.replace(
        physical_proof,
        AuthoredXmlFragment::markup(b"<physical/>".to_vec()).unwrap(),
    )
    .unwrap();
    assert!(edit.finish().is_ok());
}

#[test]
fn managed_revisions_pin_payload_admissions_without_retaining_the_revision_chain() {
    let (budget, _cancel, context) = managed_context(1024 * 1024);
    let package = SourceBackedPackage::from_read_at(
        Arc::new(OwnedSource::new(archive_bytes(FORMATTED_DOCUMENT))),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        Some(context),
    )
    .unwrap();
    let mut current = package.part(&document_uri()).unwrap().source_xml().unwrap();
    let offset = before_range(current.bytes());
    let mut stable_memory = None;
    for index in 0..32 {
        let work = budget.used(Resource::Work);
        let revision = current.begin_revision().unwrap();
        assert_eq!(
            budget.used(Resource::Work),
            work,
            "starting a revision must not rescan XML"
        );
        let proof = revision
            .checked_range(offset.clone(), &revision.bytes()[offset.clone()])
            .unwrap();
        let mut edit = revision.into_publication().unwrap();
        let tag = if index % 2 == 0 {
            b"<before attr=\"two\"/>"
        } else {
            b"<before attr=\"one\"/>"
        };
        edit.replace(proof, AuthoredXmlFragment::markup(tag.to_vec()).unwrap())
            .unwrap();
        current = edit.finish().unwrap();
        let used = budget.used(Resource::Memory);
        assert_eq!(
            *stable_memory.get_or_insert(used),
            used,
            "old revision payloads must release"
        );
    }
    let revision = current.begin_revision().unwrap();
    let proof = revision.checked_range(0..0, &[]).unwrap();
    drop(revision);
    drop(current);
    drop(package);
    assert!(
        budget.used(Resource::Memory) > 0,
        "proof pins its admitted base"
    );
    drop(proof);
    assert_eq!(budget.used(Resource::Memory), 0);
    assert_eq!(budget.used(Resource::Objects), 0);
}

#[test]
fn revision_holder_refuses_exhausted_memory_and_cancelled_contexts() {
    let memory = 1024 * 1024;
    let (budget, cancel, context) = managed_context(memory);
    let package = SourceBackedPackage::from_read_at(
        Arc::new(OwnedSource::new(archive_bytes(FORMATTED_DOCUMENT))),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        Some(context),
    )
    .unwrap();
    let source = package.part(&document_uri()).unwrap().source_xml().unwrap();
    let remaining = memory - budget.used(Resource::Memory);
    let held = budget.reserve(Resource::Memory, remaining).unwrap();
    assert!(
        matches!(source.begin_revision(), Err(OpcError::Execution(litchi_core::ExecutionError::ResourceLimit(limit))) if limit.resource == Resource::Memory)
    );
    drop(held);
    assert!(source.begin_revision().is_ok());
    cancel.cancel();
    assert!(source.begin_revision().is_err());
    drop(source);
    drop(package);
    assert_eq!(budget.used(Resource::Memory), 0);
}

fn append_candidate(
    candidate: &litchi_opc::PreparedTopology,
    text: &[u8],
) -> litchi_opc::PreparedTopology {
    candidate
        .edit_xml(&document_uri(), |source, edit| {
            let at = insertion_range(source.bytes());
            edit.replace(
                source.checked_range(at..at, b"")?,
                AuthoredXmlFragment::markup(text.to_vec())?,
            )
        })
        .unwrap()
}

fn candidate_xml(candidate: &litchi_opc::PreparedTopology) -> Vec<u8> {
    candidate
        .with_candidate(|view| Ok(view.part(&document_uri())?.data()?.as_bytes().to_vec()))
        .unwrap()
}

#[test]
fn retained_candidates_chain_fork_and_restore_the_original_archive() {
    let package = open(FORMATTED_DOCUMENT);
    let root = package
        .clone()
        .prepare_topology(SourceTopologyPlan::new())
        .unwrap();
    let first = append_candidate(&root, b"<first/>");
    let second = append_candidate(&first, b"<second/>");
    let sibling = append_candidate(&first, b"<sibling/>");
    drop(package);
    assert!(
        !candidate_xml(&first)
            .windows(9)
            .any(|bytes| bytes == b"<second/>")
    );
    assert!(
        candidate_xml(&second)
            .windows(9)
            .any(|bytes| bytes == b"<second/>")
    );
    assert!(
        !candidate_xml(&sibling)
            .windows(9)
            .any(|bytes| bytes == b"<second/>")
    );
    let noop = second.edit_xml(&document_uri(), |_, _| Ok(())).unwrap();
    assert_eq!(candidate_xml(&noop), candidate_xml(&second));
    let restored = second
        .edit_xml(&document_uri(), |source, edit| {
            let start = insertion_range(FORMATTED_DOCUMENT);
            edit.replace(
                source.checked_range(start..start + 17, b"<first/><second/>")?,
                AuthoredXmlFragment::empty(),
            )
        })
        .unwrap();
    let mut bytes = Vec::new();
    restored.publish_to_stream(&mut bytes).unwrap();
    assert_eq!(bytes, archive_bytes(FORMATTED_DOCUMENT));
    let mut bytes = Vec::new();
    second.publish_to_stream(&mut bytes).unwrap();
    let reopened = SourceBackedPackage::from_vec(
        bytes,
        SourceLimits::default(),
        SourceCacheLimits::default(),
        None,
    )
    .unwrap();
    assert_eq!(
        reopened
            .part(&document_uri())
            .unwrap()
            .data()
            .unwrap()
            .as_bytes(),
        candidate_xml(&second)
    );
}

#[test]
fn retained_xml_edits_preserve_added_removed_and_relationship_topology() {
    let package = SourceBackedPackage::from_vec(
        archive_bytes_with_sibling(FORMATTED_DOCUMENT, b"<styles/>"),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        None,
    )
    .unwrap();
    let mut plan = SourceTopologyPlan::new();
    plan.try_remove_part(PackURI::new("/word/styles.xml").unwrap())
        .unwrap();
    let addition = PackURI::new("/word/extra.xml").unwrap();
    plan.try_add_part(addition.clone(), "application/xml", b"<extra/>".to_vec())
        .unwrap();
    plan.try_add_internal_relationship(document_uri(), "rExtra", "urn:extra", addition.clone())
        .unwrap();
    let root = package.prepare_topology(plan).unwrap();
    let first = append_candidate(&root, b"<first/>");
    let second = append_candidate(&first, b"<second/>");
    drop(root);
    drop(first);
    second
        .with_candidate(|view| {
            assert!(
                view.part(&PackURI::new("/word/styles.xml").unwrap())
                    .is_err()
            );
            assert_eq!(view.part(&addition)?.data()?.as_bytes(), b"<extra/>");
            assert!(
                view.part(&document_uri())?
                    .relationships()
                    .get("rExtra")
                    .is_some()
            );
            Ok(())
        })
        .unwrap();
    assert!(second.edit_xml(&addition, |_, _| Ok(())).is_err());
    let mut bytes = Vec::new();
    second.publish_to_stream(&mut bytes).unwrap();
    let reopened = SourceBackedPackage::from_vec(
        bytes,
        SourceLimits::default(),
        SourceCacheLimits::default(),
        None,
    )
    .unwrap();
    assert!(
        reopened
            .part(&PackURI::new("/word/styles.xml").unwrap())
            .is_err()
    );
    assert_eq!(
        reopened.part(&addition).unwrap().data().unwrap().as_bytes(),
        b"<extra/>"
    );
}

#[test]
fn retained_xml_edits_refuse_swapped_transactions_and_roll_back_failed_callbacks() {
    let root = open(FORMATTED_DOCUMENT)
        .prepare_topology(SourceTopologyPlan::new())
        .unwrap();
    let first = append_candidate(&root, b"<first/>");
    let before = candidate_xml(&first);
    let foreign = open(FORMATTED_DOCUMENT)
        .part(&document_uri())
        .unwrap()
        .source_xml()
        .unwrap()
        .begin_revision()
        .unwrap();
    assert!(
        first
            .edit_xml(&document_uri(), |_, edit| {
                *edit = foreign.into_publication()?;
                Ok(())
            })
            .is_err()
    );
    assert!(
        first
            .edit_xml(&document_uri(), |source, edit| {
                let at = insertion_range(source.bytes());
                edit.replace(
                    source.checked_range(at..at, b"")?,
                    AuthoredXmlFragment::markup(b"<rolledBack/>".to_vec())?,
                )?;
                Err(OpcError::PartNotFound("intentional failure".to_owned()))
            })
            .is_err()
    );
    assert_eq!(candidate_xml(&first), before);
}

#[test]
fn managed_retained_candidates_release_intermediate_revisions() {
    let (budget, _, context) = managed_context(4 * 1024 * 1024);
    let package = SourceBackedPackage::from_read_at(
        Arc::new(OwnedSource::new(archive_bytes(FORMATTED_DOCUMENT))),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        Some(context),
    )
    .unwrap();
    let mut current = package.prepare_topology(SourceTopologyPlan::new()).unwrap();
    let mut stable = None;
    for iteration in 0..32 {
        current = current
            .edit_xml(&document_uri(), |source, edit| {
                let range = before_range(FORMATTED_DOCUMENT);
                let replacement = if iteration % 2 == 0 {
                    b"<before attr=\"two\"/>"
                } else {
                    b"<before attr=\"six\"/>"
                };
                edit.replace(
                    source.checked_range(range.clone(), &source.bytes()[range])?,
                    AuthoredXmlFragment::markup(replacement.to_vec())?,
                )
            })
            .unwrap();
        if iteration > 1 {
            assert_eq!(
                *stable.get_or_insert(budget.used(Resource::Memory)),
                budget.used(Resource::Memory)
            );
        }
    }
    let clone = current.clone();
    drop(current);
    assert!(budget.used(Resource::Memory) > 0);
    drop(clone);
    assert_eq!(budget.used(Resource::Memory), 0);
    assert_eq!(budget.used(Resource::Objects), 0);
}

#[test]
fn retained_xml_revisions_enforce_aggregate_limits_and_memory_admission() {
    let limit = FORMATTED_DOCUMENT.len() as u64 + b"<first/>".len() as u64;
    let limits = SourceLimits::builder()
        .max_total_part_bytes(limit)
        .unwrap()
        .build()
        .unwrap();
    let package = SourceBackedPackage::from_vec(
        archive_bytes(FORMATTED_DOCUMENT),
        limits,
        SourceCacheLimits::default(),
        None,
    )
    .unwrap();
    let root = package.prepare_topology(SourceTopologyPlan::new()).unwrap();
    let first = append_candidate(&root, b"<first/>");
    assert!(matches!(
        first.edit_xml(&document_uri(), |source, edit| {
            let at = insertion_range(source.bytes());
            edit.replace(
                source.checked_range(at..at, b"")?,
                AuthoredXmlFragment::markup(b"<second/>".to_vec())?,
            )
        }),
        Err(OpcError::Limit {
            kind: SourceLimitKind::TotalPartBytes,
            ..
        })
    ));
    assert_eq!(candidate_xml(&first).len() as u64, limit);

    let memory = 4 * 1024 * 1024;
    let (budget, cancel, context) = managed_context(memory);
    let package = SourceBackedPackage::from_read_at(
        Arc::new(OwnedSource::new(archive_bytes(FORMATTED_DOCUMENT))),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        Some(context),
    )
    .unwrap();
    let first = append_candidate(
        &package.prepare_topology(SourceTopologyPlan::new()).unwrap(),
        b"<first/>",
    );
    let before = candidate_xml(&first);
    let held = budget
        .reserve(Resource::Memory, memory - budget.used(Resource::Memory))
        .unwrap();
    assert!(matches!(
        first.edit_xml(&document_uri(), |_, _| Ok(())),
        Err(OpcError::Execution(
            litchi_core::ExecutionError::ResourceLimit(_)
        ))
    ));
    drop(held);
    assert_eq!(candidate_xml(&first), before);
    cancel.cancel();
    assert!(first.edit_xml(&document_uri(), |_, _| Ok(())).is_err());
    drop(first);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn retained_xml_revisions_enforce_effective_archive_byte_limits() {
    let bytes = archive_bytes(FORMATTED_DOCUMENT);
    let archive =
        litchi_opc::phys_pkg::PhysPkgReader::new(&bytes, SourceLimits::default()).unwrap();
    let total: u64 = archive
        .member_names()
        .unwrap()
        .iter()
        .map(|name| archive.read_member(name).unwrap().len() as u64)
        .sum();
    let limits = SourceLimits::builder()
        .max_aggregate_bytes(total + 8)
        .unwrap()
        .build()
        .unwrap();
    let package =
        SourceBackedPackage::from_vec(bytes, limits, SourceCacheLimits::default(), None).unwrap();
    let root = package.prepare_topology(SourceTopologyPlan::new()).unwrap();
    let first = append_candidate(&root, b"<first/>");
    assert!(matches!(
        first.edit_xml(&document_uri(), |source, edit| {
            let at = insertion_range(source.bytes());
            edit.replace(
                source.checked_range(at..at, b"")?,
                AuthoredXmlFragment::markup(b"<second/>".to_vec())?,
            )
        }),
        Err(OpcError::Limit {
            kind: SourceLimitKind::AggregateBytes,
            ..
        })
    ));
}
