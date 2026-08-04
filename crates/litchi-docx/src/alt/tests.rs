use super::*;
use litchi_opc::constants::relationship_type;
use litchi_opc::packuri::PackURI;
use litchi_opc::part::BlobPart;

use super::model::STRICT_RELATIONSHIP;

fn find(haystack: &[u8], needle: &[u8]) -> Option<usize> {
    if needle.is_empty() {
        return Some(0);
    }
    haystack
        .windows(needle.len())
        .position(|window| window == needle)
}

#[test]
fn typed_payloads_cover_word_supported_media_types_and_move_bytes() {
    let cases = [
        (Data::Docx(vec![1]), Kind::Docx, "docx"),
        (Data::Docm(vec![2]), Kind::Docm, "docm"),
        (Data::Dotx(vec![3]), Kind::Dotx, "dotx"),
        (Data::Dotm(vec![4]), Kind::Dotm, "dotm"),
        (Data::Mime(vec![5]), Kind::Mime, "eml"),
        (Data::Html(vec![6]), Kind::Html, "html"),
        (Data::Xhtml(vec![7]), Kind::Xhtml, "xhtml"),
        (Data::Rtf(vec![8]), Kind::Rtf, "rtf"),
        (Data::Text(vec![9]), Kind::Text, "txt"),
        (Data::Xml(vec![10]), Kind::Xml, "xml"),
    ];
    for (data, kind, extension) in cases {
        assert_eq!(Kind::from_media_type(data.media_type()), kind);
        assert_eq!(data.extension(), extension);
    }

    let bytes = vec![11, 12, 13];
    let pointer = bytes.as_ptr();
    let moved = Data::Html(bytes).into_bytes();
    assert_eq!(moved.as_ptr(), pointer);
}

#[test]
fn media_classification_is_parameter_tolerant_and_preserves_unknowns() {
    assert_eq!(
        Kind::from_media_type(" Text/HTML ; charset=utf-8"),
        Kind::Html
    );
    assert_eq!(Kind::from_media_type("text/rtf"), Kind::Rtf);
    assert_eq!(
        Kind::from_media_type("application/x-vendor-opaque"),
        Kind::Unknown
    );
}

#[test]
fn identifiers_and_external_targets_are_validated_once() {
    assert!(Rel::new("rId42").is_ok());
    assert!(Rel::new("bad&value").is_err());
    assert!(Import::link("https://example.invalid/chunk.html").is_ok());
    assert!(Import::link("https://example.invalid/\nchunk").is_err());
}

#[test]
fn scans_strict_and_transitional_anchors_in_source_order() {
    let xml = br#"<s:document xmlns:s="http://purl.oclc.org/ooxml/wordprocessingml/main" xmlns:q="http://purl.oclc.org/ooxml/officeDocument/relationships"><s:body><s:altChunk q:id="first"><s:altChunkPr><s:matchSrc s:val="0"/></s:altChunkPr></s:altChunk><s:altChunk q:id="second"/></s:body></s:document>"#;
    let chunks = scan(xml).unwrap().into_values().collect::<Vec<_>>();
    assert_eq!(chunks.len(), 2);
    assert_eq!(chunks[0].relationship().as_str(), "first");
    assert_eq!(chunks[0].match_source(), Some(false));
    assert_eq!(chunks[1].relationship().as_str(), "second");
    assert_eq!(chunks[1].match_source(), None);
}

#[test]
fn enforces_conformance_specific_match_source_values() {
    const STRICT_WORD: &str = "http://purl.oclc.org/ooxml/wordprocessingml/main";
    const TRANSITIONAL_WORD: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    const STRICT_RELATIONSHIPS: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships";
    const TRANSITIONAL_RELATIONSHIPS: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    let document = |word_namespace: &str, relationship_namespace: &str, value: &str| {
        format!(
            r#"<w:document xmlns:w="{word_namespace}" xmlns:r="{relationship_namespace}"><w:body><w:altChunk r:id="chunk"><w:altChunkPr><w:matchSrc w:val="{value}"/></w:altChunkPr></w:altChunk></w:body></w:document>"#
        )
    };

    for value in ["true", "1", "false", "0"] {
        assert!(
            scan(document(STRICT_WORD, STRICT_RELATIONSHIPS, value).as_bytes()).is_ok(),
            "{value}"
        );
    }
    for value in ["on", "off"] {
        assert!(
            scan(document(STRICT_WORD, STRICT_RELATIONSHIPS, value).as_bytes()).is_err(),
            "Strict unexpectedly accepted {value}"
        );
        assert!(
            scan(document(TRANSITIONAL_WORD, TRANSITIONAL_RELATIONSHIPS, value).as_bytes()).is_ok(),
            "Transitional rejected {value}"
        );
    }
}

#[test]
fn mce_selects_one_branch_without_rewriting_source_offsets() {
    let xml = br#"<q:document xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:rel="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:unsupported"><q:body><mc:AlternateContent><mc:Choice Requires="x"><q:altChunk rel:id="inactive"/></mc:Choice><mc:Fallback><q:altChunk rel:id="fallback"/></mc:Fallback></mc:AlternateContent><mc:AlternateContent><mc:Choice Requires="q"><q:altChunk rel:id="choice"/></mc:Choice><mc:Fallback><q:altChunk rel:id="inactive-fallback"/></mc:Fallback></mc:AlternateContent></q:body></q:document>"#;
    let fallback = find(xml, br#"<q:altChunk rel:id="fallback"/>"#).unwrap();
    let choice = find(xml, br#"<q:altChunk rel:id="choice"/>"#).unwrap();

    let chunks = scan(xml).unwrap();
    assert_eq!(chunks.len(), 2);
    assert_eq!(
        chunks.keys().copied().collect::<Vec<_>>(),
        vec![
            u32::try_from(fallback).unwrap(),
            u32::try_from(choice).unwrap()
        ]
    );
    assert_eq!(
        chunks
            .values()
            .map(|chunk| chunk.relationship().as_str())
            .collect::<Vec<_>>(),
        vec!["fallback", "choice"]
    );
}

#[test]
fn rejects_missing_duplicate_or_unsafe_relationships_and_invalid_children() {
    let wrapper = |anchor: &str| {
        format!(
            r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:q="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:bad="urn:bad"><w:body>{anchor}</w:body></w:document>"#
        )
    };
    for anchor in [
        r#"<w:altChunk/>"#,
        r#"<w:altChunk bad:id="x"/>"#,
        r#"<w:altChunk r:id="x" q:id="y"/>"#,
        r#"<w:altChunk r:id="bad&amp;id"/>"#,
        r#"<w:altChunk r:id="x"><w:altChunkPr/><w:altChunkPr/></w:altChunk>"#,
        r#"<w:altChunk r:id="x"><w:altChunkPr><w:matchSrc w:val="maybe"/></w:altChunkPr></w:altChunk>"#,
    ] {
        assert!(scan(wrapper(anchor).as_bytes()).is_err(), "{anchor}");
    }
}

#[test]
fn rejects_anchor_count_resource_exhaustion() {
    let mut xml = String::from(
        r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>"#,
    );
    for index in 0..=MAX_CHUNKS {
        xml.push_str(&format!(r#"<w:altChunk r:id="rId{index}"/>"#));
    }
    xml.push_str("</w:body></w:document>");
    assert!(scan(xml.as_bytes()).is_err());
}

#[test]
fn rejects_xml_byte_and_depth_resource_exhaustion() {
    let oversized = vec![b' '; MAX_XML_BYTES + 1];
    assert!(scan(&oversized).is_err());

    let mut deep = "<x>".repeat(MAX_XML_DEPTH + 1);
    deep.push_str(&"</x>".repeat(MAX_XML_DEPTH + 1));
    assert!(scan(deep.as_bytes()).is_err());
}

#[test]
fn emitted_anchors_round_trip_both_conformance_families() {
    let chunk = Chunk::new(Rel::new("rIdAlt1").unwrap(), Some(true));
    for conformance in [Conformance::Transitional, Conformance::Strict] {
        let xml = chunk.xml(conformance);
        let parsed = scan(xml.as_bytes()).unwrap().into_values().next().unwrap();
        assert_eq!(parsed, chunk);
        assert!(xml.contains(conformance.relationship_namespace()));
    }
}

#[test]
fn part_lends_original_bytes_and_classifies_without_interpreting() {
    let bytes = b"opaque foreign payload".to_vec();
    let pointer = bytes.as_ptr();
    let raw = BlobPart::new(
        PackURI::new("/word/chunk.vendor").unwrap(),
        "application/x-vendor-opaque".into(),
        bytes,
    );
    let part = Part::new(&raw);
    assert_eq!(part.name().as_str(), "/word/chunk.vendor");
    assert_eq!(part.media_type(), "application/x-vendor-opaque");
    assert_eq!(part.kind(), Kind::Unknown);
    assert_eq!(part.bytes().as_ptr(), pointer);
}

#[test]
fn recognizes_iso_word_and_strict_relationship_dialects() {
    assert!(is_relationship(
        relationship_type::ALTERNATIVE_FORMAT_IMPORT
    ));
    assert!(is_relationship(
        relationship_type::MS_ALTERNATIVE_FORMAT_IMPORT
    ));
    assert!(is_relationship(STRICT_RELATIONSHIP));
    assert!(!is_relationship(relationship_type::IMAGE));
    assert_eq!(
        Conformance::Transitional.relationship(),
        relationship_type::MS_ALTERNATIVE_FORMAT_IMPORT
    );
}
