//! Change 0764: the opaque-section attribute check and the settings MCE
//! directive facts stay bounded on hostile start tags and namespace scopes.

use std::fmt::Write as _;

use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, NamespaceResolver, PrefixDeclaration};
use quick_xml::reader::NsReader;

use super::{
    MCE_NAMESPACE, ScanError, settings_mce_directive_facts, settings_prefix_namespace,
    validate_opaque_attributes,
};

/// ` n00000=""` to ` n19999=""`, then 20,000 repeats of the last name: a tag
/// quick-xml's checked iterator reads in `O(n²)` when every item is read.
fn repeated_attributes() -> String {
    let mut attributes = String::new();
    for index in 0..20_000 {
        write!(attributes, " n{index:05}=\"\"").unwrap();
    }
    for _ in 0..20_000 {
        attributes.push_str(" n19999=\"\"");
    }
    attributes
}

/// Apply `check` to the element named `local` in `xml`, with the reader that
/// resolved its namespace scope.
fn at_element<'x, T>(
    xml: &'x str,
    local: &[u8],
    check: impl FnOnce(&NsReader<&'x [u8]>, &BytesStart<'x>) -> T,
) -> T {
    let mut reader = NsReader::from_reader(xml.as_bytes());
    loop {
        match reader.read_event().expect("the fixture is well formed") {
            Event::Start(element) | Event::Empty(element)
                if element.local_name().as_ref() == local =>
            {
                return check(&reader, &element);
            },
            Event::Eof => panic!("the fixture has the element"),
            _ => {},
        }
    }
}

fn opaque(attributes: &str) -> Result<(), ScanError> {
    let xml = format!(r#"<v:opaque xmlns:v="urn:v" xmlns:u="urn:v"{attributes}/>"#);
    at_element(&xml, b"opaque", validate_opaque_attributes)
}

#[test]
fn opaque_attributes_refuse_repeats_as_before_and_accept_many_distinct_names() {
    assert!(opaque(r#" v:a="1" a="2""#).is_ok());
    assert!(matches!(
        opaque(r#" v:a="1" v:a="2""#),
        Err(ScanError::Parser)
    ));
    assert!(matches!(
        opaque(r#" v:a="1" u:a="2""#),
        Err(ScanError::Semantic(
            "opaque section contains duplicate expanded attributes"
        ))
    ));
    assert!(matches!(
        opaque(&repeated_attributes()),
        Err(ScanError::Parser)
    ));

    let mut distinct = String::new();
    for index in 0..50_000 {
        write!(distinct, " a{index:05}=\"\"").unwrap();
    }
    assert!(opaque(&distinct).is_ok());
    distinct.push_str(r#" v:a="1" u:a="2""#);
    assert!(matches!(
        opaque(&distinct),
        Err(ScanError::Semantic(
            "opaque section contains duplicate expanded attributes"
        ))
    ));
}

/// The prefix lookup `settings_mce_directive_facts` made before 0764.
fn bindings_lookup<'r>(resolver: &'r NamespaceResolver, prefix: &[u8]) -> Option<Namespace<'r>> {
    resolver
        .bindings()
        .find_map(|(candidate, namespace)| match candidate {
            PrefixDeclaration::Named(value) if value == prefix => Some(namespace),
            PrefixDeclaration::Default | PrefixDeclaration::Named(_) => None,
        })
}

#[test]
fn prefix_lookup_matches_the_bindings_scan_in_every_scope() {
    // Redeclared, undeclared (`xmlns:q=""`), twice-declared and default
    // bindings, the predefined `xml` prefix, and prefixes never declared.
    let xml = r#"<a xmlns="urn:default" xmlns:p="urn:p1" xmlns:q="urn:q1" xmlns:xml="http://www.w3.org/XML/1998/namespace"><b xmlns:p="urn:p2" xmlns:r="urn:r" xmlns=""><c xmlns:q="" xmlns:s="urn:s"><d xmlns:q="urn:q2"/></c><e xmlns:t="urn:t1" xmlns:t="urn:t2"/></b><f/></a>"#;
    let prefixes: [&[u8]; 11] = [
        b"p", b"q", b"r", b"s", b"t", b"u", b"xml", b"xmlns", b"", b"a", b"urn",
    ];
    let mut reader = NsReader::from_reader(xml.as_bytes());
    let mut scratch = Vec::new();
    let mut scopes = 0;
    loop {
        match reader.read_event().unwrap() {
            Event::Start(_) | Event::Empty(_) | Event::End(_) => {
                scopes += 1;
                for prefix in prefixes {
                    assert_eq!(
                        settings_prefix_namespace(reader.resolver(), prefix, &mut scratch).unwrap(),
                        bindings_lookup(reader.resolver(), prefix),
                        "scope {scopes}, prefix {:?}",
                        String::from_utf8_lossy(prefix)
                    );
                }
            },
            Event::Eof => break,
            _ => {},
        }
    }
    // Three start, three empty and three end events.
    assert_eq!(scopes, 9);
}

fn directive_facts(xml: &str) -> (usize, usize) {
    at_element(xml, b"settings", |reader, element| {
        settings_mce_directive_facts(element, reader.resolver(), reader.decoder())
    })
    .unwrap()
}

#[test]
fn directive_facts_count_tokens_and_their_bound_namespaces() {
    let xml = format!(
        r#"<w:settings xmlns:w="urn:w" xmlns:mc="{MCE_NAMESPACE}" xmlns:w14="urn:w14" mc:Ignorable="w14 w15 xml" mc:ProcessContent="w14:a w15:b :c"/>"#
    );
    // Six tokens of 3 + 3 + 3 + 5 + 5 + 2 bytes; `w14` adds its namespace
    // twice and the local name `a` once.
    assert_eq!(directive_facts(&xml), (6, 21 + 7 + 7 + 1));
}

#[test]
fn directive_prefix_lookups_scan_the_bindings_in_scope_once() {
    // Sixteen levels of 256 declarations put 4,100 bindings in scope at
    // `w:settings`. A search through `NamespaceResolver::bindings` costs
    // O(B²) comparisons for every token whose prefix is undeclared.
    let mut xml = String::new();
    for level in 0..16 {
        xml.push_str("<e");
        for index in 0..256 {
            write!(xml, r#" xmlns:p{level}x{index}="urn:{level}:{index}""#).unwrap();
        }
        xml.push('>');
    }
    let mut tokens = String::new();
    let mut token_bytes = 0;
    for index in 0..2_000 {
        let token = format!("t{index}");
        token_bytes += token.len();
        tokens.push_str(&token);
        tokens.push(' ');
    }
    tokens.push_str("p15x255");
    write!(
        xml,
        r#"<w:settings xmlns:w="urn:w" xmlns:mc="{MCE_NAMESPACE}" mc:Ignorable="{tokens}"/>"#
    )
    .unwrap();
    for _ in 0..16 {
        xml.push_str("</e>");
    }
    assert_eq!(
        directive_facts(&xml),
        (2_001, token_bytes + "p15x255".len() + "urn:15:255".len())
    );
}
