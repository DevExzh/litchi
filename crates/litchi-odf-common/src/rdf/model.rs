//! `RDF` values exposed by the `OpenDocument` facade.

/// An RDF subject identified by an IRI or a blank-node identifier.
#[derive(Clone, Debug, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Subject {
    Iri(String),
    BlankNode(String),
}

/// An RDF object identified by an IRI, a blank node, or a literal value.
#[derive(Clone, Debug, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Object {
    Iri(String),
    BlankNode(String),
    Literal {
        value: String,
        datatype: Option<String>,
        language: Option<String>,
    },
}

/// One RDF predicate assertion.
#[derive(Clone, Debug, PartialEq, Eq, Hash)]
pub struct Triple {
    pub subject: Subject,
    pub predicate: String,
    pub object: Object,
}

/// One inert RDF/XML metadata graph stored in an ODF package.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Graph {
    pub path: String,
    pub base: Option<String>,
    pub prefixes: Vec<(String, String)>,
    pub triples: Vec<Triple>,
}
