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
use odt::transaction::{OperationResult, Position};
use std::io::Write;
use zip::write::SimpleFileOptions;

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

fn package_with_unchecked_rdf(mimetype: &str, family: &str, path: &str, rdf: &[u8]) -> Vec<u8> {
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
    let manifest = format!(
        "<?xml version=\"1.0\"?><manifest:manifest xmlns:manifest=\"urn:oasis:names:tc:opendocument:xmlns:manifest:1.0\"><manifest:file-entry manifest:full-path=\"/\" manifest:media-type=\"{mimetype}\"/><manifest:file-entry manifest:full-path=\"content.xml\" manifest:media-type=\"text/xml\"/><manifest:file-entry manifest:full-path=\"{path}\" manifest:media-type=\"{}\"/></manifest:manifest>",
        constants::ODF_MANIFEST_RDF_TYPE,
    );
    let options = SimpleFileOptions::default().compression_method(zip::CompressionMethod::Stored);
    let mut bytes = Vec::new();
    let mut archive = zip::ZipWriter::new(std::io::Cursor::new(&mut bytes));
    for (archive_path, archive_content) in [
        ("mimetype", mimetype.as_bytes()),
        (constants::ODF_CONTENT, content.as_bytes()),
        (path, rdf),
        ("META-INF/manifest.xml", manifest.as_bytes()),
    ] {
        archive.start_file(archive_path, options).unwrap();
        archive.write_all(archive_content).unwrap();
    }
    archive.finish().unwrap();
    bytes
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

fn apply_edit(
    document: &mut odt::Document,
    edit: odt::transaction::Edit,
) -> odt::transaction::Commit {
    let commit = edit.commit().unwrap();
    *document = commit.snapshot().document().unwrap();
    commit
}

#[test]
fn generated_graph_triple_crud_manifest_and_atomic_refs() {
    let signed =
        odt::Document::from_bytes(package(constants::ODF_TEXT, "text", &[], true)).unwrap();
    let first = literal("#anchor", "https://example.invalid/schema#label", "first");
    let signed_before = signed.to_bytes().unwrap();
    let mut signed_edit = signed.edit().unwrap();
    signed_edit
        .add_rdf_graph(None, std::slice::from_ref(&first))
        .unwrap();
    assert!(signed_edit.commit().is_err());
    assert_eq!(signed.to_bytes().unwrap(), signed_before);

    let mut document =
        odt::Document::from_bytes(package(constants::ODF_TEXT, "text", &[], false)).unwrap();
    let second = Triple {
        subject: Subject::Iri("#anchor".to_string()),
        predicate: "https://example.invalid/schema#related".to_string(),
        object: Object::Iri("https://example.invalid/resource".to_string()),
    };
    let mut graph_edit = document.edit().unwrap();
    graph_edit
        .add_rdf_graph(None, &[first.clone(), second.clone()])
        .unwrap();
    let graph_commit = apply_edit(&mut document, graph_edit);
    let path = match graph_commit.results() {
        [OperationResult::Path(path)] => path.clone(),
        _ => unreachable!(),
    };
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
    let mut triple_edit = document.edit().unwrap();
    triple_edit.add_rdf_triple(&path, &third).unwrap();
    let triple_commit = apply_edit(&mut document, triple_edit);
    assert_eq!(triple_commit.results(), &[OperationResult::Index(2)]);
    let mut move_edit = document.edit().unwrap();
    move_edit
        .move_rdf_triple_to(&path, Position::new(2), Position::new(0))
        .unwrap();
    apply_edit(&mut document, move_edit);
    assert!(
        matches!(&document.rdf_graphs().unwrap()[0].triples[0].object, Object::Literal { value, .. } if value == "third")
    );
    let replacement = literal("#anchor", "https://example.invalid/schema#label", "changed");
    let mut replacement_edit = document.edit().unwrap();
    replacement_edit
        .replace_rdf_triple(&path, 0, &replacement)
        .unwrap();
    apply_edit(&mut document, replacement_edit);
    let mut removal_edit = document.edit().unwrap();
    removal_edit
        .remove_rdf_triple_at(&path, Position::new(1))
        .unwrap();
    apply_edit(&mut document, removal_edit);

    let before = document.to_bytes().unwrap();
    let dangling = literal("#missing", "https://example.invalid/schema#label", "bad");
    let mut dangling_edit = document.edit().unwrap();
    dangling_edit.add_rdf_triple(&path, &dangling).unwrap();
    assert!(dangling_edit.commit().is_err());
    assert_eq!(document.to_bytes().unwrap(), before);
    let mut escape_edit = document.edit().unwrap();
    escape_edit
        .add_rdf_graph(Some("../escape.rdf"), &[])
        .unwrap();
    assert!(escape_edit.commit().is_err());
    assert_eq!(document.to_bytes().unwrap(), before);
}

#[test]
fn shared_graph_references_block_dangling_removal() {
    let mut document =
        odt::Document::from_bytes(package(constants::ODF_TEXT, "text", &[], false)).unwrap();
    let mut target_edit = document.edit().unwrap();
    target_edit
        .add_rdf_graph(Some("Metadata/target.rdf"), &[])
        .unwrap();
    let target_commit = apply_edit(&mut document, target_edit);
    let target = match target_commit.results() {
        [OperationResult::Path(path)] => path.clone(),
        _ => unreachable!(),
    };
    let reference = Triple {
        subject: Subject::Iri(String::new()),
        predicate: "http://docs.oasis-open.org/ns/office/1.2/meta/pkg#hasPart".to_string(),
        object: Object::Iri(target.clone()),
    };
    let mut owner_edit = document.edit().unwrap();
    owner_edit
        .add_rdf_graph(Some("manifest.rdf"), &[reference])
        .unwrap();
    let owner_commit = apply_edit(&mut document, owner_edit);
    let owner = match owner_commit.results() {
        [OperationResult::Path(path)] => path.clone(),
        _ => unreachable!(),
    };
    let before = document.to_bytes().unwrap();
    let mut blocked_removal_edit = document.edit().unwrap();
    blocked_removal_edit.remove_rdf_graph(&target).unwrap();
    assert!(blocked_removal_edit.commit().is_err());
    assert_eq!(document.to_bytes().unwrap(), before);
    let mut owner_removal_edit = document.edit().unwrap();
    owner_removal_edit.remove_rdf_graph(&owner).unwrap();
    apply_edit(&mut document, owner_removal_edit);
    let mut target_removal_edit = document.edit().unwrap();
    target_removal_edit.remove_rdf_graph(&target).unwrap();
    apply_edit(&mut document, target_removal_edit);
    assert!(document.rdf_graphs().unwrap().is_empty());
}

#[test]
fn libreoffice_rdf_and_malformed_xml_discovery() {
    let fixture = std::fs::read(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../test-data/libreoffice-core/extras/source/autotext/lang/szl/standard/FN/manifest.rdf"
    ))
    .unwrap();
    let document = odt::Document::from_bytes(package_with_unchecked_rdf(
        constants::ODF_TEXT,
        "text",
        "manifest.rdf",
        &fixture,
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
    let malformed = odt::Document::from_bytes(package_with_unchecked_rdf(
        constants::ODF_TEXT,
        "text",
        "manifest.rdf",
        hostile,
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
