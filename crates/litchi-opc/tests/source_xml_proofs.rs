#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "focused proof assertions intentionally fail on fixture errors"
)]

//! Record 0317: the source XML audit proofs a source-backed package keeps.
//!
//! * An effective revision keeps its replacement's proof, with the boundaries
//!   derived for it, so the next edit of a committed revision is proved from
//!   its own bytes; every verdict, error and published byte is a fresh
//!   package's.
//! * Publishing Part overlays audits each against the remembered original, so
//!   a package written again is not scanned again.
//! * PresentationML slides are indexed three elements deep.

use std::sync::Arc;

use litchi_core::{Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits};
use litchi_opc::SourceReadPolicy;
use litchi_opc::{
    BlobPart, OpcError, OpcPackage, PackURI, PackageWriter, PreparedTopology, SourceBackedPackage,
    SourceCacheLimits, SourceLimits, SourceTopologyPlan,
};
use xml_minifier::audit::IndexDepth;

const DOCUMENT: &str = "/word/document.xml";
const SLIDE: &str = "/ppt/slides/slide1.xml";
const WML_MAIN: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
const PML_SLIDE: &str = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml";

fn pack(name: &str) -> PackURI {
    PackURI::new(name).unwrap()
}

fn paragraph(index: usize) -> String {
    format!("<w:p><w:r><w:rPr><w:i/></w:rPr><w:t>Paragraph {index}</w:t></w:r></w:p>")
}

fn document(count: usize) -> String {
    let mut text = String::from(r#"<w:document xmlns:w="urn:w"><w:body>"#);
    for index in 0..count {
        text.push_str(&paragraph(index));
    }
    text.push_str("</w:body></w:document>");
    text
}

fn shape(index: usize) -> String {
    format!(
        "<p:sp><p:nvSpPr><p:cNvPr id=\"{}\" name=\"Box {index}\"/></p:nvSpPr><p:txBody><a:p><a:r><a:t>Shape {index}</a:t></a:r></a:p></p:txBody></p:sp>",
        index + 2
    )
}

fn slide(count: usize) -> String {
    let mut text =
        String::from(r#"<p:sld xmlns:a="urn:a" xmlns:p="urn:p"><p:cSld><p:spTree><p:nvGrpSpPr/>"#);
    for index in 0..count {
        text.push_str(&shape(index));
    }
    text.push_str("</p:spTree></p:cSld></p:sld>");
    text
}

/// A package of `parts`: name, content type, payload.
fn package_bytes(parts: &[(&str, &str, String)]) -> Vec<u8> {
    let mut package = OpcPackage::new();
    for (name, content_type, payload) in parts {
        package
            .try_add_part(Box::new(BlobPart::new(
                pack(name),
                (*content_type).to_owned(),
                payload.clone().into_bytes(),
            )))
            .unwrap();
        package.relate_to(&name[1..], "http://example.invalid/relationships/part");
    }
    PackageWriter::to_bytes(&package).unwrap()
}

fn open(bytes: &[u8]) -> SourceBackedPackage {
    SourceBackedPackage::from_vec(
        bytes.to_vec(),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        None,
    )
    .unwrap()
}

fn published(topology: &PreparedTopology) -> Vec<u8> {
    let mut bytes = Vec::new();
    topology.publish_to_stream(&mut bytes).unwrap();
    bytes
}

fn data(topology: &PreparedTopology, name: &str) -> Vec<u8> {
    topology
        .part(&pack(name))
        .unwrap()
        .data()
        .unwrap()
        .as_bytes()
        .to_vec()
}

#[test]
fn a_committed_revision_keeps_its_proof_and_chains_edits() {
    let original = document(40);
    let bytes = package_bytes(&[(DOCUMENT, WML_MAIN, original.clone())]);
    let fresh = || {
        open(&bytes)
            .prepare_topology(SourceTopologyPlan::new())
            .unwrap()
    };
    let physical = open(&bytes);
    let base = physical
        .clone()
        .prepare_topology(SourceTopologyPlan::new())
        .unwrap();
    // Nothing is audited yet, so nothing is retained.
    assert!(physical.xml_audit_proof(&pack(DOCUMENT)).is_none());
    assert!(
        base.part(&pack(DOCUMENT))
            .unwrap()
            .xml_audit_proof()
            .is_none()
    );

    let edits = [
        original.replacen("Paragraph 2<", "Paragraph two<", 1),
        String::new(),
        String::new(),
    ];
    let first = base
        .apply_xml_bytes(
            &pack(DOCUMENT),
            original.as_bytes(),
            Arc::new(edits[0].clone().into_bytes()),
        )
        .unwrap();
    // The physical original is remembered with its indexed proof, and the
    // revision keeps the replacement's proof over its own allocation.
    let remembered = physical.xml_audit_proof(&pack(DOCUMENT)).unwrap();
    assert_eq!(remembered.bytes().as_slice(), original.as_bytes());
    assert_eq!(remembered.index_depth(), Some(IndexDepth::CHILDREN));
    let view = first.part(&pack(DOCUMENT)).unwrap();
    let kept = view
        .xml_audit_proof()
        .expect("the revision keeps its proof");
    assert!(core::ptr::eq(
        kept.bytes().as_slice(),
        view.data().unwrap().as_bytes()
    ));
    assert_eq!(kept.index_depth(), Some(IndexDepth::CHILDREN));

    // A chain of edits far apart, then an insertion; each publishes what a
    // fresh package publishes for the same final bytes.
    let mut current = first;
    let mut text = edits[0].clone();
    let second = text.replacen("Paragraph 38<", "Paragraph thirty-eight<", 1);
    let at = second.find(&paragraph(20)).unwrap();
    let third = format!(
        "{}<w:p><w:r><w:t>Inserted</w:t></w:r></w:p>{}",
        &second[..at],
        &second[at..]
    );
    for next in [second, third] {
        let chained = current
            .apply_xml_bytes(
                &pack(DOCUMENT),
                text.as_bytes(),
                Arc::new(next.clone().into_bytes()),
            )
            .unwrap();
        let direct = fresh()
            .apply_xml_bytes(
                &pack(DOCUMENT),
                original.as_bytes(),
                Arc::new(next.clone().into_bytes()),
            )
            .unwrap();
        assert_eq!(published(&chained), published(&direct));
        assert_eq!(data(&chained, DOCUMENT), next.as_bytes());
        let kept = chained
            .part(&pack(DOCUMENT))
            .unwrap()
            .xml_audit_proof()
            .unwrap();
        assert_eq!(kept.bytes().as_slice(), next.as_bytes());
        assert!(kept.index_depth().is_some());
        current = chained;
        text = next;
    }

    // Every edit of the chained revision, defects included, has the verdict
    // and error a fresh package reports for the same bytes.
    let at = text.find(&paragraph(10)).unwrap();
    for fragment in [
        "<w:p>",
        "</w:body><w:body>",
        "<x:p/>",
        "<w:p><w:t>&bogus;</w:t></w:p>",
        "<w:p><w:t>x]]>y</w:t></w:p>",
        "<w:p a=\"1\" a=\"2\"/>",
        "<w:p><w:t>a &amp; b</w:t></w:p>",
    ] {
        let next = Arc::new(format!("{}{fragment}{}", &text[..at], &text[at..]).into_bytes());
        let chained = current.apply_xml_bytes(&pack(DOCUMENT), text.as_bytes(), Arc::clone(&next));
        let direct = fresh().apply_xml_bytes(&pack(DOCUMENT), original.as_bytes(), next);
        match (chained, direct) {
            (Ok(chained), Ok(direct)) => assert_eq!(published(&chained), published(&direct)),
            (Err(chained), Err(direct)) => {
                assert!(
                    matches!(chained, OpcError::XmlPublication { .. }),
                    "{chained}"
                );
                assert_eq!(chained.to_string(), direct.to_string(), "{fragment}");
            },
            (chained, direct) => {
                panic!("{fragment}: {:?} against {:?}", chained.err(), direct.err())
            },
        }
    }
}

#[test]
fn a_revision_keeps_at_most_four_replacement_proofs() {
    let names: Vec<String> = (0..6)
        .map(|index| format!("/custom/part{index}.xml"))
        .collect();
    let parts: Vec<(&str, &str, String)> = names
        .iter()
        .map(|name| (name.as_str(), "application/xml", document(4)))
        .collect();
    let bytes = package_bytes(&parts);
    let mut current = open(&bytes)
        .prepare_topology(SourceTopologyPlan::new())
        .unwrap();
    let original = document(4);
    let edited = original.replacen("Paragraph 1<", "Paragraph one<", 1);
    for (edits, name) in names.iter().enumerate() {
        current = current
            .apply_xml_bytes(
                &pack(name),
                original.as_bytes(),
                Arc::new(edited.clone().into_bytes()),
            )
            .unwrap();
        // The Part just edited keeps its proof; at most four replaced Parts
        // keep one.
        assert!(
            current
                .part(&pack(name))
                .unwrap()
                .xml_audit_proof()
                .is_some()
        );
        let kept = names[..=edits]
            .iter()
            .filter(|name| {
                current
                    .part(&pack(name))
                    .unwrap()
                    .xml_audit_proof()
                    .is_some()
            })
            .count();
        assert_eq!(kept, (edits + 1).min(4), "after {} edits", edits + 1);
    }
    // A Part that gave its proof up is still edited and published exactly.
    let again = edited.replacen("Paragraph 3<", "Paragraph three<", 1);
    let chained = current
        .apply_xml_bytes(
            &pack(&names[0]),
            edited.as_bytes(),
            Arc::new(again.clone().into_bytes()),
        )
        .unwrap();
    assert_eq!(data(&chained, &names[0]), again.as_bytes());
}

#[test]
fn a_managed_package_keeps_no_proof() {
    let bytes = package_bytes(&[(DOCUMENT, WML_MAIN, document(10))]);
    let budget = Budget::root(
        "source-xml-proofs-managed",
        Limits::new(
            16 * 1024 * 1024,
            u64::MAX,
            u64::MAX,
            u64::MAX,
            u64::MAX,
            u64::MAX,
        ),
    );
    let (_cancellation_source, cancellation) = CancellationSource::pair();
    let execution_limits = ExecutionLimits::new(
        std::num::NonZeroUsize::new(1).unwrap(),
        std::num::NonZeroUsize::new(1).unwrap(),
        std::num::NonZeroU64::new(16 * 1024 * 1024).unwrap(),
        0,
    )
    .unwrap();
    let context = ExecutionContext::new(budget, cancellation, execution_limits);
    let package = SourceBackedPackage::from_read_at(
        Arc::new(litchi_core::OwnedSource::new(bytes)),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        Some(context),
    )
    .unwrap();
    let mut plan = SourceTopologyPlan::new();
    plan.try_replace_part(
        pack(DOCUMENT),
        document(10)
            .replacen("Paragraph 1<", "Paragraph one<", 1)
            .into_bytes(),
    )
    .unwrap();
    let topology = package.clone().prepare_topology(plan).unwrap();
    assert!(
        topology
            .part(&pack(DOCUMENT))
            .unwrap()
            .xml_audit_proof()
            .is_none()
    );
    assert!(package.xml_audit_proof(&pack(DOCUMENT)).is_none());
}

#[test]
fn written_overlays_are_audited_against_the_remembered_original() {
    let original = slide(60);
    let bytes = package_bytes(&[
        (SLIDE, PML_SLIDE, original.clone()),
        (DOCUMENT, WML_MAIN, document(10)),
    ]);
    let write = |package: &SourceBackedPackage, replacement: &str| {
        let mut output = Vec::new();
        package
            .clone()
            .write_part_overlays_shared_to_stream(
                &mut output,
                vec![(pack(SLIDE), Arc::new(replacement.as_bytes().to_vec()))],
            )
            .map(|()| output)
    };
    let package = open(&bytes);
    let set = original.replacen(">Shape 30<", ">Shape thirty<", 1);
    let first = write(&package, &set).unwrap();
    assert_eq!(first, write(&open(&bytes), &set).unwrap());
    // The slide is remembered with a three-deep index.
    let remembered = package.xml_audit_proof(&pack(SLIDE)).unwrap();
    assert_eq!(remembered.index_depth(), Some(IndexDepth::GRANDCHILDREN));
    assert!(package.xml_audit_proof(&pack(DOCUMENT)).is_none());
    // Later writes are proved from it: an added shape, and defects refused
    // with the error a fresh package reports, every time.
    let end = original.find("</p:spTree>").unwrap();
    let added = format!("{}{}{}", &original[..end], shape(60), &original[end..]);
    assert_eq!(
        write(&package, &added).unwrap(),
        write(&open(&bytes), &added).unwrap()
    );
    for defect in [
        "<p:sp>",
        "<z:sp/>",
        "<p:sp a=\"1\" a=\"1\"/>",
        "</p:spTree>",
    ] {
        let broken = format!("{}{defect}{}", &original[..end], &original[end..]);
        for _ in 0..2 {
            let remembered = write(&package, &broken).expect_err("refused");
            let fresh = write(&open(&bytes), &broken).expect_err("refused");
            assert!(
                matches!(remembered, OpcError::XmlPublication { .. }),
                "{remembered}"
            );
            assert_eq!(remembered.to_string(), fresh.to_string(), "{defect}");
        }
    }
}

#[test]
fn an_invalid_original_is_reported_on_every_write() {
    // The writer refuses to produce such a Part, so it is put in place in
    // the archive afterwards.
    let generated = package_bytes(&[(SLIDE, PML_SLIDE, slide(1))]);
    let reader =
        litchi_opc::phys_pkg::PhysPkgReader::new(&generated, SourceLimits::default()).unwrap();
    let mut writer = soapberry_zip::office::StreamingArchiveWriter::new(
        soapberry_zip::office::StreamingArchiveLimits::default(),
    );
    for name in reader.member_names().unwrap() {
        let member = if name == SLIDE[1..] {
            b"<!DOCTYPE a><a/>".to_vec()
        } else {
            reader.read_member(&name).unwrap()
        };
        writer.write_stored(&name, &member).unwrap();
    }
    let bytes = writer.finish_to_bytes().unwrap();
    let package = open(&bytes);
    let mut errors = Vec::new();
    for _ in 0..2 {
        let mut output = Vec::new();
        errors.push(
            package
                .clone()
                .write_part_overlays_shared_to_stream(
                    &mut output,
                    vec![(pack(SLIDE), Arc::new(b"<b/>".to_vec()))],
                )
                .expect_err("refused")
                .to_string(),
        );
    }
    assert_eq!(errors[0], errors[1]);
    assert!(package.xml_audit_proof(&pack(SLIDE)).is_none());
}
