use super::{Graph, Object, Subject, Triple, codec};

#[test]
fn codec_round_trips_contextual_rdf_values() {
    let triple = Triple {
        subject: Subject::Iri("#anchor".to_string()),
        predicate: "https://example.invalid/schema#label".to_string(),
        object: Object::Literal {
            value: "value".to_string(),
            datatype: None,
            language: Some("en".to_string()),
        },
    };
    let xml = codec::serialize_graph(std::slice::from_ref(&triple))
        .unwrap_or_else(|error| panic!("RDF fixture serialization must succeed: {error}"));
    let parsed = codec::parse("metadata.rdf", &xml)
        .unwrap_or_else(|error| panic!("RDF fixture parsing must succeed: {error}"));

    assert_eq!(
        parsed.graph,
        Graph {
            path: "metadata.rdf".to_string(),
            base: None,
            prefixes: vec![("rdf".to_string(), codec::RDF.to_string())],
            triples: vec![triple],
        }
    );
}

#[test]
fn codec_rejects_entity_expansion() {
    let xml = r#"<?xml version="1.0"?><!DOCTYPE rdf:RDF [<!ENTITY x "bad">]><rdf:RDF xmlns:rdf="http://www.w3.org/1999/02/22-rdf-syntax-ns#"/>"#;
    assert!(codec::parse("metadata.rdf", xml).is_err());
}
