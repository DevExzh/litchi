//! Record 0771: the MCE stream's names share their namespace URI, so its
//! event-name sharing and directive checks avoid per-name URI copies and
//! hashing. Registered-extension checks and observer work are outside this
//! property. Checks through namespace identities preserve URI-text equality.
//!
//! Work is bounded by counting, not timing: the namespace-URI bytes the scope
//! hashes, compares or examines for facts, the bytes the stream copies
//! (names, values and URIs), and the scope's index lookups.

use std::{cell::RefCell, collections::HashSet, convert::Infallible, io::Cursor};

use super::codec::process_markup_compatibility;
use super::model::{Capabilities, ExpandedName, Limits, NAMESPACE, NamespaceUri, XMLNS_NAMESPACE};
use super::scope::counter::{counted, counted_copies, counted_uri_bytes};
use super::stream::{
    SemanticEvent, StreamError, StreamLimits, process_markup_compatibility_stream_with_observers,
};

const MC: &str = NAMESPACE;
const XML: &str = "http://www.w3.org/XML/1998/namespace";

/// The stream's raw and semantic events as text, in which a namespace of
/// `long` bytes reads `{U}`, and whether every name in that namespace shared
/// the first such name's copy of the URI.
struct Transcript {
    lines: Vec<String>,
    long: usize,
    first: Option<NamespaceUri>,
    shared: bool,
}

impl Transcript {
    fn name(&mut self, name: &ExpandedName) -> String {
        if name.namespace.len() == self.long {
            match &self.first {
                None => self.first = Some(name.namespace.clone()),
                Some(first) => self.shared &= first.shares_storage_with(&name.namespace),
            }
            format!("{{U}}{}", name.local_name)
        } else {
            format!("{{{}}}{}", name.namespace, name.local_name)
        }
    }
}

/// What the stream did with one document.
#[derive(Debug, PartialEq, Eq)]
struct Work {
    lines: Vec<String>,
    outcome: String,
    shared: bool,
    uri_read: usize,
    copied: usize,
    lookups: usize,
}

fn stream_work(xml: &str, capabilities: &Capabilities, long: usize) -> Work {
    let transcript = RefCell::new(Transcript {
        lines: Vec::new(),
        long,
        first: None,
        shared: true,
    });
    let run = || {
        let mut input = Cursor::new(xml.as_bytes());
        process_markup_compatibility_stream_with_observers(
            &mut input,
            capabilities,
            &StreamLimits::default(),
            |element| {
                let mut transcript = transcript.borrow_mut();
                let mut line = format!(
                    "raw {:?} {}",
                    element.kind,
                    transcript.name(&element.expanded_name)
                );
                for attribute in element.attrs() {
                    line.push(' ');
                    line.push_str(&transcript.name(&attribute.expanded_name));
                }
                transcript.lines.push(line);
                Ok::<(), Infallible>(())
            },
            |event| {
                let mut transcript = transcript.borrow_mut();
                let line = match &event {
                    SemanticEvent::Start(element) | SemanticEvent::Empty(element) => {
                        let mut line = format!("start {}", transcript.name(&element.expanded_name));
                        for attribute in element.attrs() {
                            line.push(' ');
                            line.push_str(&transcript.name(&attribute.expanded_name));
                        }
                        line
                    },
                    SemanticEvent::End(end) => {
                        format!("end {}", transcript.name(&end.expanded_name))
                    },
                    _ => "other".to_owned(),
                };
                transcript.lines.push(line);
                Ok::<(), Infallible>(())
            },
        )
    };
    let (((result, lookups), uri_read), copied) =
        counted_copies(|| counted_uri_bytes(|| counted(run)));
    let transcript = transcript.into_inner();
    Work {
        lines: transcript.lines,
        outcome: match result {
            Ok(report) => format!("ok {report:?}"),
            Err(error) => format!("err {error}"),
        },
        shared: transcript.shared,
        uri_read,
        copied,
        lookups,
    }
}

/// `urn:` and `count` `u`s: a URI of `count + 4` bytes.
fn uri(count: usize) -> String {
    format!("urn:{}", "u".repeat(count))
}

/// The 0764 review's stream probe input (harness case
/// `mce_stream_review_long_uri`): `p` bound to a URI of `u_count + 4` bytes
/// beside an ignorable `x`, around elements of at most 1,000 attributes in
/// `p` until `attributes` are written.
fn review_long_uri(u_count: usize, attributes: usize) -> String {
    let mut xml = format!(
        r#"<r xmlns:mc="{MC}" xmlns:x="urn:x" xmlns:p="{}" mc:Ignorable="x">"#,
        uri(u_count)
    );
    let per = attributes.min(1_000);
    let mut written = 0;
    while written < attributes {
        xml.push_str("<e");
        for index in 0..per {
            xml.push_str(&format!(r#" p:a{index}="""#));
        }
        xml.push_str("/>");
        written += per;
    }
    xml.push_str("</r>");
    xml
}

/// The harness's `mce_stream_long_uri` family: `z` bound on the root to a URI
/// of `uri_bytes` bytes, with `directives`, around `elements` elements of
/// `attributes` attributes in `z`.
fn long_uri(uri_bytes: usize, directives: &str, elements: usize, attributes: usize) -> String {
    let names: String = (0..attributes)
        .map(|index| format!(r#" z:a{index}="""#))
        .collect();
    format!(
        r#"<r xmlns:mc="{MC}" xmlns:z="{}"{directives}>{}</r>"#,
        uri(uri_bytes - 4),
        format!("<e{names}/>").repeat(elements)
    )
}

/// The harness's `mce_stream_long_uri_elements` family: `elements` empty
/// elements in `z`.
fn long_uri_elements(uri_bytes: usize, directives: &str, elements: usize) -> String {
    format!(
        r#"<r xmlns:mc="{MC}" xmlns:z="{}"{directives}>{}</r>"#,
        uri(uri_bytes - 4),
        "<z:e/>".repeat(elements)
    )
}

/// The harness's `mce_stream_long_uri_tokens`: `elements` elements that each
/// make `z` ignorable and preserve its elements.
fn long_uri_tokens(uri_bytes: usize, elements: usize) -> String {
    format!(
        r#"<r xmlns:mc="{MC}" xmlns:z="{}">{}</r>"#,
        uri(uri_bytes - 4),
        r#"<e mc:Ignorable="z" mc:PreserveElements="z:*"/>"#.repeat(elements)
    )
}

/// Check that the stream does the same work on `long_xml`, whose one long
/// namespace URI has `long` bytes, as on `short_xml`, the same document with
/// a URI of `short` bytes, apart from a constant number of passes over the
/// URI for its one declaration.
fn assert_work_per_declaration(
    label: &str,
    capabilities: &Capabilities,
    (long_xml, long): (&str, usize),
    (short_xml, short): (&str, usize),
) {
    let long_work = stream_work(long_xml, capabilities, long);
    let short_work = stream_work(short_xml, capabilities, short);
    assert!(
        long_work.outcome.starts_with("ok"),
        "{label}: {}",
        long_work.outcome
    );
    assert_eq!(long_work.outcome, short_work.outcome, "{label}");
    assert_eq!(long_work.lines, short_work.lines, "{label}");
    assert!(
        long_work.shared,
        "{label}: names in one namespace copied its URI"
    );
    assert_eq!(long_work.lookups, short_work.lookups, "{label}");
    let extra = long - short;
    // The scope hashes the declared URI once and computes its facts once.
    let read = long_work.uri_read - short_work.uri_read;
    assert!(read <= 4 * extra, "{label}: {read} URI bytes read");
    // The declaration's value is copied as the attribute's raw and decoded
    // value, twice as a declaration (raw and semantic chains) and once into
    // the scope: a constant number of copies per declaration, none per name.
    let copied = long_work.copied - short_work.copied;
    assert!(copied <= 8 * extra, "{label}: {copied} bytes copied");
}

#[test]
fn the_reviewers_long_uri_inputs_cost_the_stream_per_declaration_not_per_name() {
    // Before this change each of these 4,000 attributes copied the
    // 1,040,004-byte URI into its name, again for the raw observer, and
    // hashed it for the duplicate and ignorable checks: 3.7 s in a release
    // build.
    let long = review_long_uri(1_040_000, 4_000);
    let short = review_long_uri(60, 4_000);
    assert_work_per_declaration(
        "review",
        &Capabilities::ooxml_baseline(),
        (&long, 1_040_004),
        (&short, 64),
    );
    let long_bytes = (1 << 20) - 64;
    for directives in ["", r#" mc:Ignorable="z""#] {
        let long = long_uri(long_bytes, directives, 4, 1_000);
        let short = long_uri(64, directives, 4, 1_000);
        assert_work_per_declaration(
            directives,
            &Capabilities::new(),
            (&long, long_bytes),
            (&short, 64),
        );
    }
}

#[test]
fn long_uri_element_names_and_directive_tokens_cost_the_stream_per_declaration() {
    let long_bytes = (1 << 20) - 64;
    for directives in ["", r#" mc:Ignorable="z""#] {
        let long = long_uri_elements(long_bytes, directives, 4_000);
        let short = long_uri_elements(64, directives, 4_000);
        assert_work_per_declaration(
            directives,
            &Capabilities::new(),
            (&long, long_bytes),
            (&short, 64),
        );
    }
    let long = long_uri_tokens(long_bytes, 4_000);
    let short = long_uri_tokens(64, 4_000);
    assert_work_per_declaration(
        "tokens",
        &Capabilities::new(),
        (&long, long_bytes),
        (&short, 64),
    );
}

#[test]
fn the_fixed_namespaces_are_shared_without_a_copy() {
    let xml = format!(
        r#"<r xmlns:x="{XML}" xmlns:n="{XMLNS_NAMESPACE}" xml:space="preserve" x:lang="en" n:q=""><e xmlns="" a=""/></r>"#
    );
    let mut names = Vec::new();
    let mut input = Cursor::new(xml.as_bytes());
    process_markup_compatibility_stream_with_observers(
        &mut input,
        &Capabilities::new(),
        &StreamLimits::default(),
        |element| {
            names.push(element.expanded_name);
            names.extend(element.attributes.into_iter().map(|a| a.expanded_name));
            Ok::<(), Infallible>(())
        },
        |_| Ok::<(), Infallible>(()),
    )
    .expect("the document is valid");
    let expected = [
        ExpandedName::new("", "r"),
        ExpandedName::new(XMLNS_NAMESPACE, "x"),
        ExpandedName::new(XMLNS_NAMESPACE, "n"),
        ExpandedName::new(XML, "space"),
        ExpandedName::new(XML, "lang"),
        ExpandedName::new(XMLNS_NAMESPACE, "q"),
        ExpandedName::new("", "e"),
        ExpandedName::new(XMLNS_NAMESPACE, "xmlns"),
        ExpandedName::new("", "a"),
    ];
    assert_eq!(names, expected);
    // A name in the `xml` namespace, through either prefix, shares the
    // library's static text, and a reset default namespace is no namespace.
    assert!(names[3].namespace.shares_storage_with(&names[4].namespace));
    assert!(names[6].namespace.is_empty());
    assert!(names[6].namespace.shares_storage_with(&NamespaceUri::NONE));
}

/// The expanded names of a stream's semantic events, and its outcome.
fn semantic_events(xml: &str, capabilities: &Capabilities) -> (Vec<String>, String) {
    let mut events = Vec::new();
    let mut input = Cursor::new(xml.as_bytes());
    let result = super::stream::process_markup_compatibility_stream(
        &mut input,
        capabilities,
        &StreamLimits::default(),
        |event| {
            let line = match &event {
                SemanticEvent::Start(element) | SemanticEvent::Empty(element) => {
                    let mut line = format!(
                        "start {{{}}}{}",
                        element.expanded_name.namespace, element.expanded_name.local_name
                    );
                    for attribute in element.attrs() {
                        line.push_str(&format!(
                            " {{{}}}{}={}",
                            attribute.expanded_name.namespace,
                            attribute.expanded_name.local_name,
                            attribute.value()
                        ));
                    }
                    line
                },
                SemanticEvent::End(end) => format!(
                    "end {{{}}}{}",
                    end.expanded_name.namespace, end.expanded_name.local_name
                ),
                _ => "other".to_owned(),
            };
            events.push(line);
            Ok::<(), Infallible>(())
        },
    );
    let outcome = match result {
        Ok(report) => format!("ok {report:?}"),
        Err(error) => format!("err {error}"),
    };
    (events, outcome)
}

fn codec_output(xml: &str, capabilities: &Capabilities) -> String {
    match process_markup_compatibility(xml.as_bytes(), capabilities, &Limits::default()) {
        Ok(output) => format!(
            "ok {} {:?}",
            String::from_utf8_lossy(output.xml.as_ref()),
            output.report
        ),
        Err(error) => format!("err {error}"),
    }
}

#[test]
fn aliased_prefixes_name_one_attribute() {
    for xml in [
        r#"<r xmlns:a="urn:a" xmlns:b="urn:a" a:k="1" b:k="2"/>"#.to_owned(),
        format!(r#"<r xmlns:x="{XML}" x:lang="en" xml:lang="fr"/>"#),
        format!(r#"<r xmlns:m="{XMLNS_NAMESPACE}" xmlns:n="{XMLNS_NAMESPACE}" m:q="" n:q=""/>"#),
        r#"<r xmlns:a="urn:a"><e xmlns:b="urn:a" b:k="1" a:k="2"/></r>"#.to_owned(),
    ] {
        let (_, outcome) = semantic_events(&xml, &Capabilities::new());
        assert!(outcome.ends_with("duplicate attribute"), "{xml}: {outcome}");
    }
    // A prefix bound to another namespace names another attribute.
    let (_, outcome) = semantic_events(
        r#"<r xmlns:a="urn:a" xmlns:b="urn:b" a:k="1" b:k="2"/>"#,
        &Capabilities::new(),
    );
    assert!(outcome.starts_with("ok"), "{outcome}");
}

/// Pairs of documents that differ only in the prefix a compatibility
/// directive or name uses for one namespace: through an alias, and directly.
/// Both processors must treat the two alike, since a namespace is its URI.
fn aliased_pairs() -> Vec<(String, String)> {
    let root = format!(r#"xmlns:mc="{MC}" xmlns:x="{XML}" xmlns:a="urn:a" xmlns:b="urn:a""#);
    let pair = |aliased: &str, direct: &str, body: &str| {
        (
            format!("<r {root} {aliased}>{body}</r>"),
            format!("<r {root} {direct}>{body}</r>"),
        )
    };
    vec![
        // `x` is the xml namespace: making it ignorable makes `xml:space`
        // ignorable.
        pair(
            r#"mc:Ignorable="x""#,
            r#"mc:Ignorable="xml""#,
            r#"<e xml:space="preserve"/>"#,
        ),
        pair(
            r#"mc:Ignorable="x" mc:PreserveAttributes="xml:space""#,
            r#"mc:Ignorable="xml" mc:PreserveAttributes="xml:space""#,
            r#"<e xml:space="preserve" xml:lang="en"/>"#,
        ),
        pair(
            r#"mc:Ignorable="xml" mc:PreserveAttributes="x:*""#,
            r#"mc:Ignorable="xml" mc:PreserveAttributes="xml:*""#,
            r#"<e xml:space="preserve"/>"#,
        ),
        // `b` is `a`'s namespace.
        pair(
            r#"mc:Ignorable="b""#,
            r#"mc:Ignorable="a""#,
            r#"<a:e a:k="1"><f/></a:e><b:e/>"#,
        ),
        pair(
            r#"mc:Ignorable="a" mc:ProcessContent="b:e""#,
            r#"mc:Ignorable="a" mc:ProcessContent="a:e""#,
            r#"<a:e><f/></a:e>"#,
        ),
        pair(
            r#"mc:Ignorable="a" mc:PreserveElements="b:*""#,
            r#"mc:Ignorable="a" mc:PreserveElements="a:*""#,
            r#"<a:e><a:f/></a:e>"#,
        ),
        pair(
            r#"mc:Ignorable="a" mc:PreserveAttributes="b:k a:k""#,
            r#"mc:Ignorable="a" mc:PreserveAttributes="a:k a:k""#,
            r#"<e a:k="1"/>"#,
        ),
        pair(r#"mc:Ignorable="a b""#, r#"mc:Ignorable="a""#, r#"<a:e/>"#),
    ]
}

#[test]
fn directives_through_aliased_prefixes_decide_as_their_namespaces_do() {
    let mut understand_a = Capabilities::new();
    understand_a.understand_namespace("urn:a");
    for capabilities in [
        Capabilities::new(),
        Capabilities::ooxml_baseline(),
        understand_a,
    ] {
        for (aliased, direct) in aliased_pairs() {
            let (aliased_events, aliased_outcome) = semantic_events(&aliased, &capabilities);
            let (direct_events, direct_outcome) = semantic_events(&direct, &capabilities);
            assert_eq!(aliased_outcome, direct_outcome, "stream: {aliased}");
            assert_eq!(aliased_events, direct_events, "stream: {aliased}");
            assert_eq!(
                codec_output(&aliased, &capabilities),
                codec_output(&direct, &capabilities),
                "processor: {aliased}"
            );
        }
    }
}

#[test]
fn a_namespace_uri_reads_compares_and_hashes_as_its_text() {
    let shared = NamespaceUri::from("urn:a");
    let copy = NamespaceUri::from("urn:a".to_owned());
    let clone = shared.clone();
    assert_eq!(shared, copy);
    assert!(!shared.shares_storage_with(&copy));
    assert!(shared.shares_storage_with(&clone));
    assert_eq!(shared, "urn:a");
    assert_eq!("urn:a", shared);
    assert_eq!(shared, "urn:a".to_owned());
    assert_eq!(&*shared, "urn:a");
    assert_eq!(shared.as_str(), "urn:a");
    assert_eq!(
        format!("{shared}|{shared:?}|{shared:>7}"),
        r#"urn:a|"urn:a"|  urn:a"#
    );
    assert_eq!(String::from(shared.clone()), "urn:a");
    let later = NamespaceUri::from("urn:b");
    assert!(shared < later);
    assert_eq!(shared.cmp(&copy), core::cmp::Ordering::Equal);

    // Hashing and lookups by text agree with the text's.
    let set: HashSet<NamespaceUri> = [shared.clone(), NamespaceUri::from_static(XML)]
        .into_iter()
        .collect();
    assert!(set.contains("urn:a"));
    assert!(set.contains(XML));
    assert!(set.contains(&copy));
    assert!(!set.contains("urn:b"));

    // The empty URI is no namespace, however it is made.
    for empty in [
        NamespaceUri::NONE,
        NamespaceUri::default(),
        NamespaceUri::from(""),
        NamespaceUri::from(String::new()),
        NamespaceUri::from_static(""),
    ] {
        assert!(empty.is_empty());
        assert_eq!(empty, "");
        assert!(empty.shares_storage_with(&NamespaceUri::NONE));
    }
    assert!(!shared.is_empty());
    assert_eq!(
        ExpandedName::new("urn:a", "k"),
        ExpandedName {
            namespace: copy,
            local_name: "k".to_owned(),
        }
    );
}

#[test]
fn a_stream_limited_below_the_fixed_namespaces_still_refuses_their_names() {
    // The stream checked each namespace URI against its name limit where it
    // copied one; it checks the same bound where it now shares the copy.
    let limits = StreamLimits {
        max_name_bytes: 16,
        ..StreamLimits::default()
    };
    for xml in [r#"<r xmlns:p="urn:p"/>"#, r#"<r xml:space="preserve"/>"#] {
        let mut input = Cursor::new(xml.as_bytes());
        let error = super::stream::process_markup_compatibility_stream(
            &mut input,
            &Capabilities::new(),
            &limits,
            |_| Ok::<(), Infallible>(()),
        )
        .expect_err(xml);
        assert!(
            matches!(&error, StreamError::Mce { error: super::model::Error::LimitExceeded(message), .. } if message == "stream name bytes"),
            "{xml}: {error:?}"
        );
    }
}
