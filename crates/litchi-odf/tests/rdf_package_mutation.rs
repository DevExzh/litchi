#![allow(
    clippy::unwrap_used,
    reason = "tests are expected to panic on unexpected errors"
)]

use litchi_odf::{odp, ods, odt};
use litchi_odf_common::core::{OwnedPackage, PackageWriter};
use litchi_odf_common::{
    constants,
    rdf::{Object, Subject, Triple},
};

const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const TEXT: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const TABLE: &str = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
const DRAW: &str = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";

fn package(mimetype: &str, family: &str, rdf: &[(&str, &[u8])], signed: bool) -> Vec<u8> {
    let body = match family {
        "text" => "<office:text><text:p xml:id=\"anchor\">body</text:p></office:text>",
        "spreadsheet" => {
            "<office:spreadsheet><table:table table:name=\"Sheet1\" xml:id=\"anchor\"/></office:spreadsheet>"
        },
        "presentation" => {
            "<office:presentation><draw:page draw:name=\"Slide1\" xml:id=\"anchor\"/></office:presentation>"
        },
        _ => unreachable!(),
    };
    let content = format!(
        "<?xml version=\"1.0\"?><office:document-content xmlns:office=\"{OFFICE}\" xmlns:text=\"{TEXT}\" xmlns:table=\"{TABLE}\" xmlns:draw=\"{DRAW}\"><office:body>{body}</office:body></office:document-content>"
    );
    let mut writer = PackageWriter::new();
    writer.set_mimetype(mimetype).unwrap();
    writer
        .add_file(constants::ODF_CONTENT, content.as_bytes())
        .unwrap();
    for (path, bytes) in rdf {
        writer
            .add_file_with_media_type(path, bytes, constants::ODF_MANIFEST_RDF_TYPE)
            .unwrap();
    }
    if signed {
        writer
            .add_file_with_media_type(
                "META-INF/documentsignatures.xml",
                b"<signatures/>",
                "text/xml",
            )
            .unwrap();
    }
    writer.finish_to_bytes().unwrap()
}

fn literal(subject: &str, predicate: &str, value: &str) -> Triple {
    Triple {
        subject: Subject::Iri(subject.to_string()),
        predicate: predicate.to_string(),
        object: Object::Literal {
            value: value.to_string(),
            datatype: None,
            language: Some("en".to_string()),
        },
    }
}

#[test]
fn generated_graph_triple_crud_manifest_and_atomic_refs() {
    let mut document =
        odt::Document::from_bytes(package(constants::ODF_TEXT, "text", &[], true)).unwrap();
    let first = literal("#anchor", "https://example.invalid/schema#label", "first");
    let second = Triple {
        subject: Subject::Iri("#anchor".to_string()),
        predicate: "https://example.invalid/schema#related".to_string(),
        object: Object::Iri("https://example.invalid/resource".to_string()),
    };
    let path = document
        .add_rdf_graph(None, &[first.clone(), second.clone()])
        .unwrap();
    assert_eq!(path, "Metadata/metadata_1.rdf");
    let graph = &document.rdf_graphs().unwrap()[0];
    assert_eq!(graph.triples.len(), 2);
    assert_eq!(graph.triples[0], first);

    let archive = OwnedPackage::from_bytes(document.to_bytes().unwrap()).unwrap();
    assert_eq!(
        archive.package().unwrap().manifest().get_media_type(&path),
        Some(constants::ODF_MANIFEST_RDF_TYPE)
    );
    assert!(!archive.has_file("META-INF/documentsignatures.xml").unwrap());

    let third = literal("#anchor", "https://example.invalid/schema#label", "third");
    assert_eq!(document.add_rdf_triple(&path, &third).unwrap(), 2);
    document.move_rdf_triple(&path, 2, 0).unwrap();
    assert!(
        matches!(&document.rdf_graphs().unwrap()[0].triples[0].object, Object::Literal { value, .. } if value == "third")
    );
    let replacement = literal("#anchor", "https://example.invalid/schema#label", "changed");
    document.replace_rdf_triple(&path, 0, &replacement).unwrap();
    document.remove_rdf_triple(&path, 1).unwrap();

    let before = document.to_bytes().unwrap();
    let dangling = literal("#missing", "https://example.invalid/schema#label", "bad");
    assert!(document.add_rdf_triple(&path, &dangling).is_err());
    assert_eq!(document.to_bytes().unwrap(), before);
    assert!(document.add_rdf_graph(Some("../escape.rdf"), &[]).is_err());
    assert_eq!(document.to_bytes().unwrap(), before);
}

#[test]
fn shared_graph_references_block_dangling_removal() {
    let mut document =
        odt::Document::from_bytes(package(constants::ODF_TEXT, "text", &[], false)).unwrap();
    let target = document
        .add_rdf_graph(Some("Metadata/target.rdf"), &[])
        .unwrap();
    let reference = Triple {
        subject: Subject::Iri(String::new()),
        predicate: "http://docs.oasis-open.org/ns/office/1.2/meta/pkg#hasPart".to_string(),
        object: Object::Iri(target.clone()),
    };
    let owner = document
        .add_rdf_graph(Some("manifest.rdf"), &[reference])
        .unwrap();
    let before = document.to_bytes().unwrap();
    assert!(document.remove_rdf_graph(&target).is_err());
    assert_eq!(document.to_bytes().unwrap(), before);
    document.remove_rdf_graph(&owner).unwrap();
    document.remove_rdf_graph(&target).unwrap();
    assert!(document.rdf_graphs().unwrap().is_empty());
}

#[test]
fn libreoffice_rdf_and_malformed_xml_discovery() {
    let fixture = std::fs::read(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../test-data/libreoffice-core/extras/source/autotext/lang/szl/standard/FN/manifest.rdf"
    ))
    .unwrap();
    let document = odt::Document::from_bytes(package(
        constants::ODF_TEXT,
        "text",
        &[("manifest.rdf", fixture.as_slice())],
        false,
    ))
    .unwrap();
    let graph = &document.rdf_graphs().unwrap()[0];
    assert!(!graph.triples.is_empty());
    assert!(
        graph
            .prefixes
            .iter()
            .any(|(prefix, value)| prefix == "rdf" && value.contains("rdf-syntax"))
    );

    let hostile = br#"<?xml version="1.0"?><!DOCTYPE rdf:RDF [<!ENTITY x "bad">]><rdf:RDF xmlns:rdf="http://www.w3.org/1999/02/22-rdf-syntax-ns#"/>"#;
    let malformed = odt::Document::from_bytes(package(
        constants::ODF_TEXT,
        "text",
        &[("manifest.rdf", hostile)],
        false,
    ))
    .unwrap();
    assert!(malformed.rdf_graphs().is_err());
}

#[test]
fn ods_and_odp_facades_roundtrip_blank_nodes_and_datatypes() {
    let triple = Triple {
        subject: Subject::BlankNode("node_1".to_string()),
        predicate: "https://example.invalid/schema#count".to_string(),
        object: Object::Literal {
            value: "42".to_string(),
            datatype: Some("http://www.w3.org/2001/XMLSchema#integer".to_string()),
            language: None,
        },
    };
    let mut sheet = ods::facade::Spreadsheet::from_bytes(package(
        constants::ODF_SPREADSHEET,
        "spreadsheet",
        &[],
        false,
    ))
    .unwrap();
    sheet
        .add_rdf_graph(Some("manifest.rdf"), std::slice::from_ref(&triple))
        .unwrap();
    assert_eq!(sheet.rdf_graphs().unwrap()[0].triples[0], triple);
    sheet.replace_rdf_graph("manifest.rdf", &[]).unwrap();
    assert!(sheet.rdf_graphs().unwrap()[0].triples.is_empty());

    let mut slides = odp::facade::Presentation::from_bytes(package(
        constants::ODF_PRESENTATION,
        "presentation",
        &[],
        false,
    ))
    .unwrap();
    slides
        .add_rdf_graph(None, std::slice::from_ref(&triple))
        .unwrap();
    assert_eq!(slides.rdf_graphs().unwrap()[0].triples[0], triple);
}
