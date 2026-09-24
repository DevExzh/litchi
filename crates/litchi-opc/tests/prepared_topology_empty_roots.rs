#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "integration fixture and invariant assertions"
)]

//! First additions must work for valid empty source containers.

use std::sync::Arc;

use litchi_core::OwnedSource;
use litchi_opc::{PackURI, SourceBackedPackage, SourceTopologyPlan};
use soapberry_zip::office::{ArchiveReader, StreamingArchiveWriter};

const CT: &str = "http://schemas.openxmlformats.org/package/2006/content-types";
const REL: &str = "http://schemas.openxmlformats.org/package/2006/relationships";

fn uri(value: &str) -> PackURI {
    PackURI::new(value).unwrap()
}

fn empty_package(types: &str, relationships: Option<&str>) -> Vec<u8> {
    let mut zip = StreamingArchiveWriter::new();
    zip.write_stored("[Content_Types].xml", types.as_bytes())
        .unwrap();
    if let Some(relationships) = relationships {
        zip.write_stored("_rels/.rels", relationships.as_bytes())
            .unwrap();
    }
    zip.finish_to_bytes().unwrap()
}

fn publish_first_part(source: Vec<u8>, add_relationship: bool) -> Vec<u8> {
    let package = SourceBackedPackage::from_read_at(Arc::new(OwnedSource::new(source))).unwrap();
    let mut plan = SourceTopologyPlan::new();
    plan.try_add_part(
        uri("/first.xml"),
        "application/example+xml",
        b"<first/>".to_vec(),
    )
    .unwrap();
    if add_relationship {
        plan.try_add_internal_relationship(
            uri("/"),
            "first",
            "urn:example:first",
            uri("/first.xml"),
        )
        .unwrap();
    }
    let prepared = package
        .prepare_topology(plan)
        .expect("first addition is supported");
    prepared
        .with_candidate(|candidate| {
            assert_eq!(
                candidate.part(&uri("/first.xml"))?.data()?.as_bytes(),
                b"<first/>"
            );
            if add_relationship {
                assert!(candidate.relationships(&uri("/"))?.get("first").is_some());
            }
            Ok(())
        })
        .unwrap();
    let mut output = Vec::new();
    prepared.publish_to_stream(&mut output).unwrap();
    let reopened =
        SourceBackedPackage::from_read_at(Arc::new(OwnedSource::new(output.clone()))).unwrap();
    assert_eq!(
        reopened
            .part(&uri("/first.xml"))
            .unwrap()
            .data()
            .unwrap()
            .as_bytes(),
        b"<first/>"
    );
    output
}

#[test]
fn first_mapping_expands_empty_default_and_prefixed_content_types_roots() {
    for (root, namespace) in [("Types", "xmlns"), ("ct:Types", "xmlns:ct")] {
        let types = format!(
            "<?xml version='1.0'?>\r\n<!--before--><{root} {namespace}='{CT}'  /><!--after-->"
        );
        let output = publish_first_part(empty_package(&types, None), false);
        let xml = ArchiveReader::new(output.as_slice())
            .unwrap()
            .read("[Content_Types].xml")
            .unwrap();
        let xml = std::str::from_utf8(&xml).unwrap();
        assert!(xml.starts_with(&format!(
            "<?xml version='1.0'?>\r\n<!--before--><{root} {namespace}='{CT}'  >"
        )));
        assert!(xml.ends_with(&format!("</{root}><!--after-->")));
    }
}

#[test]
fn first_mapping_preserves_prefixed_paired_root_and_comments() {
    let types = format!("<c:Types xmlns:c='{CT}'>\r\n<!--inside--></c:Types><!--tail-->");
    let output = publish_first_part(empty_package(&types, None), false);
    let xml = ArchiveReader::new(output.as_slice())
        .unwrap()
        .read("[Content_Types].xml")
        .unwrap();
    let xml = std::str::from_utf8(&xml).unwrap();
    assert!(xml.starts_with(&format!("<c:Types xmlns:c='{CT}'>\r\n<!--inside-->")));
    assert!(xml.contains("<c:Override "));
    assert!(xml.ends_with("</c:Types><!--tail-->"));
}

#[test]
fn first_relationship_expands_empty_default_and_prefixed_roots() {
    let types = format!(
        "<Types xmlns='{CT}'><Default Extension='rels' ContentType='application/vnd.openxmlformats-package.relationships+xml'/></Types>"
    );
    for (root, namespace) in [("Relationships", "xmlns"), ("r:Relationships", "xmlns:r")] {
        let relationships = format!(
            "<?xml version='1.0'?>\n<!--before--><{root} {namespace}='{REL}' /><!--after-->"
        );
        let output = publish_first_part(empty_package(&types, Some(&relationships)), true);
        let xml = ArchiveReader::new(output.as_slice())
            .unwrap()
            .read("_rels/.rels")
            .unwrap();
        let xml = std::str::from_utf8(&xml).unwrap();
        assert!(xml.starts_with(&format!(
            "<?xml version='1.0'?>\n<!--before--><{root} {namespace}='{REL}' >"
        )));
        assert!(xml.ends_with(&format!("</{root}><!--after-->")));
    }
}
