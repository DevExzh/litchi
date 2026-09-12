#![allow(
    clippy::unwrap_used,
    reason = "Fixed in-memory raw ZIP fixtures keep publication assertions concise."
)]

use litchi_odf_common::{
    constants,
    core::{
        AuthoredXmlFragment, OwnedPackage, XmlSourcePart, XmlSplicePublication,
        rebuild_package_with_xml_splices,
    },
};
use std::{collections::HashMap, io::Cursor, io::Write, ops::Range};

const CONTENT: &[u8] = b"<?xml version=\"1.0\"?>\n<document>\n  <leaf/>\n</document>";
const RDF: &[u8] = b"<rdf>\n <leaf/>\n</rdf>";
const DECLARED: &[u8] = b"<declared>\n <leaf/>\n</declared>";
const PLUS_XML: &[u8] = b"<vendor>\n <leaf/>\n</vendor>";
const SIGNATURE: &[u8] = b"<signatures>\n <leaf/>\n</signatures>";

fn raw_package(marker: &str) -> Vec<u8> {
    raw_package_with_content(marker, CONTENT)
}

fn raw_package_with_content(marker: &str, content: &[u8]) -> Vec<u8> {
    let manifest = format!(
        r#"<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><manifest:file-entry manifest:full-path="/" manifest:media-type="{}"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/><manifest:file-entry manifest:full-path="graph.rdf" manifest:media-type=""/><manifest:file-entry manifest:full-path="declared-part" manifest:media-type="text/xml"/><manifest:file-entry manifest:full-path="vendor-part" manifest:media-type="application/vnd.example+xml"/><manifest:file-entry manifest:full-path="META-INF/documentsignatures.xml" manifest:media-type="application/vnd.oasis.opendocument.digital-signature"/><manifest:file-entry manifest:full-path="marker.bin" manifest:media-type="application/octet-stream"/></manifest:manifest>"#,
        constants::ODF_DATABASE
    );
    let mut output = Cursor::new(Vec::new());
    {
        let mut zip = zip::ZipWriter::new(&mut output);
        let stored = zip::write::SimpleFileOptions::default()
            .compression_method(zip::CompressionMethod::Stored);
        let deflated = zip::write::SimpleFileOptions::default()
            .compression_method(zip::CompressionMethod::Deflated);
        zip.start_file("mimetype", stored).unwrap();
        zip.write_all(constants::ODF_DATABASE.as_bytes()).unwrap();
        for (path, bytes) in [
            ("META-INF/manifest.xml", manifest.as_bytes()),
            ("content.xml", content),
            ("graph.rdf", RDF),
            ("declared-part", DECLARED),
            ("vendor-part", PLUS_XML),
            ("META-INF/documentsignatures.xml", SIGNATURE),
            ("marker.bin", marker.as_bytes()),
        ] {
            zip.start_file(path, deflated).unwrap();
            zip.write_all(bytes).unwrap();
        }
        zip.finish().unwrap();
    }
    output.into_inner()
}

fn root_opening(part: &XmlSourcePart) -> Range<usize> {
    let start = part
        .bytes()
        .windows(b"<document".len())
        .position(|window| window == b"<document")
        .unwrap();
    let end = part.bytes()[start..]
        .iter()
        .position(|byte| *byte == b'>')
        .map(|offset| start + offset + 1)
        .unwrap();
    start..end
}

fn root_candidate(
    part: &XmlSourcePart,
    replacement: &[u8],
) -> (litchi_odf_common::core::XmlSourceRange, Vec<u8>) {
    let range = root_opening(part);
    let proof = part
        .checked_range(range.clone(), &part.bytes()[range.clone()])
        .unwrap();
    let mut candidate = Vec::new();
    candidate.extend_from_slice(&part.bytes()[..range.start]);
    candidate.extend_from_slice(replacement);
    candidate.extend_from_slice(&part.bytes()[range.end..]);
    candidate.shrink_to_fit();
    (proof, candidate)
}

fn insertion(part: &XmlSourcePart) -> Range<usize> {
    let position = part
        .bytes()
        .windows(2)
        .rposition(|window| window == b"</")
        .unwrap();
    position..position
}

#[test]
fn every_raw_xml_part_class_requires_an_audited_fragment() {
    let package = OwnedPackage::from_bytes(raw_package("one")).unwrap();
    for path in [
        "content.xml",
        "graph.rdf",
        "declared-part",
        "vendor-part",
        "META-INF/documentsignatures.xml",
    ] {
        let part = XmlSourcePart::load(&package, path).unwrap();
        let proof = part.checked_range(insertion(&part), b"").unwrap();
        let mut publication = XmlSplicePublication::new(part);
        let fragment = AuthoredXmlFragment::markup(b"\n <authored/>".to_vec());
        assert!(fragment.is_err(), "noncompact fragment accepted for {path}");
        let unclassified = AuthoredXmlFragment::markup(b"plain text".to_vec());
        assert!(
            unclassified.is_err(),
            "unclassified fragment accepted for {path}"
        );
        publication
            .replace(
                proof,
                AuthoredXmlFragment::markup(b"<authored/>".to_vec()).unwrap(),
            )
            .unwrap();
    }
}

#[test]
fn source_identity_stale_ranges_and_overlaps_are_rejected() {
    let identical_bytes = raw_package("same-bytes");
    let first = OwnedPackage::from_bytes(identical_bytes.clone()).unwrap();
    let second = OwnedPackage::from_bytes(identical_bytes).unwrap();
    let first_part = XmlSourcePart::load(&first, "content.xml").unwrap();
    let second_part = XmlSourcePart::load(&second, "content.xml").unwrap();
    assert!(first_part.checked_range(0..5, b"stale").is_err());

    let foreign = first_part
        .checked_range(insertion(&first_part), b"")
        .unwrap();
    let mut publication = XmlSplicePublication::new(second_part.clone());
    assert!(
        publication
            .replace(foreign, AuthoredXmlFragment::deletion())
            .is_err()
    );

    let range = 32..39;
    let expected = &second_part.bytes()[range.clone()];
    publication
        .replace(
            second_part.checked_range(range.clone(), expected).unwrap(),
            AuthoredXmlFragment::markup(b"<new/>".to_vec()).unwrap(),
        )
        .unwrap();
    assert!(
        publication
            .replace(
                second_part
                    .checked_range((range.start + 1)..range.end, &expected[1..])
                    .unwrap(),
                AuthoredXmlFragment::deletion(),
            )
            .is_err()
    );

    let foreign_part = XmlSourcePart::load(&first, "graph.rdf").unwrap();
    let foreign_publication = XmlSplicePublication::new(foreign_part);
    assert!(
        rebuild_package_with_xml_splices(&second, vec![foreign_publication], 2 * 1024 * 1024)
            .is_err()
    );
}

#[test]
fn rebuild_preserves_source_bytes_outside_splices_and_enumerates_every_member() {
    let source = OwnedPackage::from_bytes(raw_package("opaque-marker")).unwrap();
    let part = XmlSourcePart::load(&source, "content.xml").unwrap();
    let insert = insertion(&part);
    let before = part.bytes()[..insert.start].to_vec();
    let after = part.bytes()[insert.end..].to_vec();
    let proof = part.checked_range(insert, b"").unwrap();
    let mut publication = XmlSplicePublication::new(part);
    publication
        .replace(
            proof,
            AuthoredXmlFragment::markup(b"<authored/>".to_vec()).unwrap(),
        )
        .unwrap();

    let rebuilt = OwnedPackage::from_bytes(
        rebuild_package_with_xml_splices(&source, vec![publication], 2 * 1024 * 1024).unwrap(),
    )
    .unwrap();
    let content = rebuilt.get_file("content.xml").unwrap();
    assert!(content.starts_with(&before));
    assert!(content.ends_with(&after));
    assert_eq!(rebuilt.get_file("graph.rdf").unwrap(), RDF);
    assert_eq!(rebuilt.get_file("declared-part").unwrap(), DECLARED);
    assert_eq!(rebuilt.get_file("vendor-part").unwrap(), PLUS_XML);
    assert_eq!(rebuilt.get_file("marker.bin").unwrap(), b"opaque-marker");
    assert!(!rebuilt.has_file("META-INF/documentsignatures.xml").unwrap());

    let files: HashMap<_, _> = rebuilt
        .files()
        .unwrap()
        .into_iter()
        .map(|path| (path.clone(), rebuilt.get_file(&path).unwrap()))
        .collect();
    assert_eq!(files.len(), 7);
    for path in [
        "mimetype",
        "META-INF/manifest.xml",
        "content.xml",
        "graph.rdf",
        "declared-part",
        "vendor-part",
        "marker.bin",
    ] {
        assert!(files.contains_key(path), "missing rebuilt member {path}");
    }
}

#[test]
fn explicit_fragment_classes_accept_compact_bytes_only() {
    assert!(AuthoredXmlFragment::start_tag(b"<node value=\"one\">".to_vec()).is_ok());
    assert!(AuthoredXmlFragment::text(b"one &amp; two".to_vec()).is_ok());
    let _deletion = AuthoredXmlFragment::deletion();
    assert!(AuthoredXmlFragment::start_tag(b"<node  value=\"one\">".to_vec()).is_err());
    assert!(AuthoredXmlFragment::text(b"   ".to_vec()).is_err());
}

#[test]
fn bounded_rebuild_refuses_to_materialize_an_oversized_archive() {
    let source = OwnedPackage::from_bytes(raw_package("bounded")).unwrap();
    let part = XmlSourcePart::load(&source, "content.xml").unwrap();
    let publication = XmlSplicePublication::new(part);
    assert!(rebuild_package_with_xml_splices(&source, vec![publication], 64).is_err());
}

#[test]
fn source_candidate_preserves_noncompact_lexical_content_and_opaque_events() {
    let opaque_content = br#"<?xml version="1.0"?>
<document>
  <!-- <?opaque?> -->
  <leaf><![CDATA[<?opaque?>]]></leaf>
  <?future instruction?>
</document>"#;
    let source =
        OwnedPackage::from_bytes(raw_package_with_content("opaque-source", opaque_content))
            .unwrap();
    let part = XmlSourcePart::load(&source, "content.xml").unwrap();
    let (proof, candidate) = root_candidate(
        &part,
        br#"<document xmlns:q="urn:example:q"   q:root = "one">"#,
    );
    let expected_candidate = candidate.clone();
    let publication = XmlSplicePublication::from_source_start_tag_candidate_with_limit(
        part,
        proof,
        candidate,
        2 * 1024 * 1024,
    )
    .expect("complete source candidate should validate");
    let output = rebuild_package_with_xml_splices(&source, vec![publication], 2 * 1024 * 1024)
        .expect("bounded source publication");
    let rebuilt = OwnedPackage::from_bytes(output).unwrap();
    assert_eq!(rebuilt.get_file("content.xml").unwrap(), expected_candidate);
    assert_eq!(rebuilt.get_file("marker.bin").unwrap(), b"opaque-source");

    let escaped_part = XmlSourcePart::load(&source, "content.xml").unwrap();
    let (escaped_proof, escaped_candidate) =
        root_candidate(&escaped_part, br#"<document value="&lt;legal&amp;value">"#);
    XmlSplicePublication::from_source_start_tag_candidate_with_limit(
        escaped_part,
        escaped_proof,
        escaped_candidate,
        2 * 1024 * 1024,
    )
    .expect("escaped XML attribute values are legal");
}

#[test]
fn source_candidate_rejects_namespace_and_xml_grammar_defects() {
    let source = OwnedPackage::from_bytes(raw_package("candidate-defects")).unwrap();
    let invalid_openings: &[&[u8]] = &[
        br#"<document q:item="one">"#,
        br#"<1document>"#,
        br#"<document value="a<b">"#,
        b"<document value=\"\xEF\xBF\xBE\">",
        b"<document value=\"\xEF\xBF\xBF\">",
        br#"<document value="&#xFFFE;">"#,
        br#"<document value="&#xFFFF;">"#,
        br#"<document xmlns="http://www.w3.org/XML/1998/namespace">"#,
        br#"<document xmlns="http://www.w3.org/2000/xmlns/">"#,
        br#"<document xmlns:q="http://www.w3.org/XML/1998/namespace" q:item="one">"#,
        br#"<document xmlns:q="http://www.w3.org/XML/1998&#x2F;namespace" q:item="one">"#,
        br#"<document xmlns="http://www.w3.org/XML/1998&#x2F;namespace">"#,
        br#"<document xmlns:a="urn:u" xmlns:b="urn:&#x75;" a:x="1" b:x="2">"#,
        br#"<?xml version="1.1"?>"#,
        br#"<?xml version="1.0"?><document>"#,
        br#"<?xml?>"#,
    ];
    for opening in invalid_openings {
        let part = XmlSourcePart::load(&source, "content.xml").unwrap();
        let (proof, candidate) = root_candidate(&part, opening);
        assert!(
            XmlSplicePublication::from_source_start_tag_candidate_with_limit(
                part,
                proof,
                candidate,
                2 * 1024 * 1024,
            )
            .is_err(),
            "invalid source candidate was accepted: {:?}",
            String::from_utf8_lossy(opening)
        );
    }

    let part = XmlSourcePart::load(&source, "content.xml").unwrap();
    let (proof, candidate) = root_candidate(
        &part,
        br#"<document xmlns:q="urn:example&#x3A;q" q:item="one">"#,
    );
    XmlSplicePublication::from_source_start_tag_candidate_with_limit(
        part,
        proof,
        candidate,
        2 * 1024 * 1024,
    )
    .expect("escaped namespace aliases are legal XML");

    let part = XmlSourcePart::load(&source, "content.xml").unwrap();
    let (proof, candidate) = root_candidate(&part, br#"<document>"#);
    let mut candidate = candidate;
    candidate.insert(0, b' ');
    assert!(
        XmlSplicePublication::from_source_start_tag_candidate_with_limit(
            part,
            proof,
            candidate,
            2 * 1024 * 1024,
        )
        .is_err(),
        "candidate changes bytes before the proven opening tag"
    );
}

#[test]
fn source_candidate_rejects_invalid_comment_and_processing_instruction_characters() {
    for content in [
        b"<?xml version=\"1.0\"?><document><!-- \xEF\xBF\xBE --><leaf/></document>".as_slice(),
        b"<?xml version=\"1.0\"?><document><?future \xEF\xBF\xBF ?><leaf/></document>".as_slice(),
    ] {
        let source =
            OwnedPackage::from_bytes(raw_package_with_content("invalid-event-character", content))
                .unwrap();
        let part = XmlSourcePart::load(&source, "content.xml").unwrap();
        let (proof, candidate) = root_candidate(&part, br#"<document>"#);
        assert!(
            XmlSplicePublication::from_source_start_tag_candidate_with_limit(
                part,
                proof,
                candidate,
                2 * 1024 * 1024,
            )
            .is_err(),
            "invalid comment or processing-instruction character was accepted"
        );
    }
}

#[test]
fn source_candidate_output_limit_is_checked_before_bounded_publication() {
    let source = OwnedPackage::from_bytes(raw_package("exact-source-cap")).unwrap();
    let first_part = XmlSourcePart::load(&source, "content.xml").unwrap();
    let (first_proof, first_candidate) = root_candidate(
        &first_part,
        br#"<document xmlns:q="urn:example:q" q:value="changed">"#,
    );
    let first_publication = XmlSplicePublication::from_source_start_tag_candidate_with_limit(
        first_part,
        first_proof,
        first_candidate,
        2 * 1024 * 1024,
    )
    .unwrap();
    let output =
        rebuild_package_with_xml_splices(&source, vec![first_publication], 2 * 1024 * 1024)
            .unwrap();

    let exact_part = XmlSourcePart::load(&source, "content.xml").unwrap();
    let (exact_proof, exact_candidate) = root_candidate(
        &exact_part,
        br#"<document xmlns:q="urn:example:q" q:value="changed">"#,
    );
    let exact_candidate_len = exact_candidate.len();
    let exact_publication = XmlSplicePublication::from_source_start_tag_candidate_with_limit(
        exact_part,
        exact_proof,
        exact_candidate,
        exact_candidate_len,
    )
    .unwrap_or_else(|_| panic!("candidate should fit its part limit"));
    assert_eq!(
        rebuild_package_with_xml_splices(&source, vec![exact_publication], output.len()).unwrap(),
        output
    );

    let under_part = XmlSourcePart::load(&source, "content.xml").unwrap();
    let (under_proof, under_candidate) = root_candidate(
        &under_part,
        br#"<document xmlns:q="urn:example:q" q:value="changed">"#,
    );
    let under_publication = XmlSplicePublication::from_source_start_tag_candidate_with_limit(
        under_part,
        under_proof,
        under_candidate,
        exact_candidate_len - 1,
    );
    assert!(under_publication.is_err());

    let under_part = XmlSourcePart::load(&source, "content.xml").unwrap();
    let (under_proof, under_candidate) = root_candidate(
        &under_part,
        br#"<document xmlns:q="urn:example:q" q:value="changed">"#,
    );
    let under_publication = XmlSplicePublication::from_source_start_tag_candidate_with_limit(
        under_part,
        under_proof,
        under_candidate,
        2 * 1024 * 1024,
    )
    .unwrap();
    assert!(
        rebuild_package_with_xml_splices(&source, vec![under_publication], output.len() - 1)
            .is_err()
    );

    let capacity_part = XmlSourcePart::load(&source, "content.xml").unwrap();
    let (capacity_proof, mut capacity_candidate) = root_candidate(
        &capacity_part,
        br#"<document xmlns:q="urn:example:q" q:value="changed">"#,
    );
    capacity_candidate.reserve(1024);
    assert!(
        XmlSplicePublication::from_source_start_tag_candidate_with_limit(
            capacity_part,
            capacity_proof,
            capacity_candidate,
            exact_candidate_len,
        )
        .is_err(),
        "candidate retained capacity beyond the caller cap was accepted"
    );
}

#[test]
fn source_candidate_publication_rejects_mixed_splice_edits() {
    let source = OwnedPackage::from_bytes(raw_package("mixed-source-publication")).unwrap();
    let part = XmlSourcePart::load(&source, "content.xml").unwrap();
    let expected = part.bytes().first().copied().unwrap();
    let proof = part.checked_range(0..1, &[expected]).unwrap();
    let (opening_proof, candidate) = root_candidate(
        &part,
        br#"<document xmlns:q="urn:example:q" q:value="changed">"#,
    );
    let mut publication = XmlSplicePublication::from_source_start_tag_candidate_with_limit(
        part,
        opening_proof,
        candidate,
        2 * 1024 * 1024,
    )
    .unwrap();
    assert!(
        publication
            .replace(proof, AuthoredXmlFragment::deletion())
            .is_err()
    );
}

#[test]
fn source_candidate_requires_exact_archive_provenance() {
    let bytes = raw_package("source-provenance");
    let first = OwnedPackage::from_bytes(bytes.clone()).unwrap();
    let second = OwnedPackage::from_bytes(bytes).unwrap();
    let part = XmlSourcePart::load(&first, "content.xml").unwrap();
    let (proof, candidate) = root_candidate(
        &part,
        br#"<document xmlns:q="urn:example:q" q:value="changed">"#,
    );
    let publication = XmlSplicePublication::from_source_start_tag_candidate_with_limit(
        part,
        proof,
        candidate,
        2 * 1024 * 1024,
    )
    .unwrap();
    assert!(rebuild_package_with_xml_splices(&second, vec![publication], 2 * 1024 * 1024).is_err());
}

#[test]
fn source_empty_expansion_preserves_inherited_root_attributes() {
    let content = br#"<?xml version="1.0"?>
<document xmlns:q="urn:example:q" q:root="one"/>"#;
    let source =
        OwnedPackage::from_bytes(raw_package_with_content("empty-expansion", content)).unwrap();
    let part = XmlSourcePart::load(&source, "content.xml").unwrap();
    let (proof, candidate) = root_candidate(
        &part,
        br#"<document xmlns:q="urn:example:q" q:root="one"><q:child/></document>"#,
    );
    let publication = XmlSplicePublication::from_source_empty_expansion_candidate_with_limit(
        part,
        proof,
        candidate.clone(),
        candidate.len(),
    )
    .expect("empty root expansion should validate");
    let output = rebuild_package_with_xml_splices(&source, vec![publication], 2 * 1024 * 1024)
        .expect("empty root expansion should publish");
    let rebuilt = OwnedPackage::from_bytes(output).unwrap();
    assert_eq!(rebuilt.get_file("content.xml").unwrap(), candidate);

    let part = XmlSourcePart::load(&source, "content.xml").unwrap();
    let (proof, changed) = root_candidate(
        &part,
        br#"<document xmlns:q="urn:example:q" q:root="changed"><q:child/></document>"#,
    );
    assert!(
        XmlSplicePublication::from_source_empty_expansion_candidate_with_limit(
            part,
            proof,
            changed,
            2 * 1024 * 1024,
        )
        .is_err(),
        "empty expansion must preserve the original root attributes"
    );
}

#[test]
fn source_empty_expansion_accepts_inherited_prefixes_and_noncompact_opening() {
    let content = br#"<?xml version="1.0"?>
<office:document xmlns:office="urn:office"><office:automatic-styles  office:version = "1"/></office:document>"#;
    let source =
        OwnedPackage::from_bytes(raw_package_with_content("inherited-expansion", content)).unwrap();
    let part = XmlSourcePart::load(&source, "content.xml").unwrap();
    let start = part
        .bytes()
        .windows(b"<office:automatic-styles".len())
        .position(|window| window == b"<office:automatic-styles")
        .unwrap();
    let end = part.bytes()[start..]
        .iter()
        .position(|byte| *byte == b'>')
        .map(|offset| start + offset + 1)
        .unwrap();
    let range = start..end;
    let proof = part
        .checked_range(range.clone(), &part.bytes()[range.clone()])
        .unwrap();
    let replacement =
        br#"<office:automatic-styles  office:version = "1"><office:style/></office:automatic-styles>"#;
    let mut candidate = Vec::new();
    candidate.extend_from_slice(&part.bytes()[..range.start]);
    candidate.extend_from_slice(replacement);
    candidate.extend_from_slice(&part.bytes()[range.end..]);

    let publication = XmlSplicePublication::from_source_empty_expansion_candidate_with_limit(
        part,
        proof,
        candidate.clone(),
        candidate.len(),
    )
    .expect("inherited namespace expansion should validate");
    let output = rebuild_package_with_xml_splices(&source, vec![publication], 2 * 1024 * 1024)
        .expect("inherited namespace expansion should publish");
    let rebuilt = OwnedPackage::from_bytes(output).unwrap();
    assert_eq!(rebuilt.get_file("content.xml").unwrap(), candidate);
}

#[test]
fn source_candidate_rejects_depth_beyond_the_common_xml_bound() {
    let depth = 4_096usize;
    let mut content = Vec::new();
    content.extend_from_slice(b"<?xml version=\"1.0\"?><document>");
    for _ in 0..depth {
        content.extend_from_slice(b"<node>");
    }
    for _ in 0..depth {
        content.extend_from_slice(b"</node>");
    }
    content.extend_from_slice(b"</document>");
    let source =
        OwnedPackage::from_bytes(raw_package_with_content("deep-source", &content)).unwrap();
    let part = XmlSourcePart::load(&source, "content.xml").unwrap();
    let (proof, candidate) = root_candidate(&part, b"<document>");
    assert!(
        XmlSplicePublication::from_source_start_tag_candidate_with_limit(
            part,
            proof,
            candidate,
            2 * 1024 * 1024,
        )
        .is_err(),
        "candidate depth beyond the bounded XML profile must fail closed"
    );
}

#[test]
fn source_candidate_rejects_empty_element_at_the_depth_boundary() {
    let depth = 4_095usize;
    let mut content = Vec::new();
    content.extend_from_slice(b"<?xml version=\"1.0\"?><document>");
    for _ in 0..depth {
        content.extend_from_slice(b"<node>");
    }
    content.extend_from_slice(b"<leaf/>");
    for _ in 0..depth {
        content.extend_from_slice(b"</node>");
    }
    content.extend_from_slice(b"</document>");
    let source =
        OwnedPackage::from_bytes(raw_package_with_content("empty-depth-boundary", &content))
            .unwrap();
    let part = XmlSourcePart::load(&source, "content.xml").unwrap();
    let (proof, candidate) = root_candidate(&part, b"<document>");
    assert!(
        XmlSplicePublication::from_source_start_tag_candidate_with_limit(
            part,
            proof,
            candidate,
            2 * 1024 * 1024,
        )
        .is_err(),
        "empty element at the depth boundary was accepted"
    );
}
