//! Public contract tests for the source-preserving ODF namespace reader.
//!
//! These tests exercise the adapter's namespace boundary: lexical event bytes
//! stay available for source ranges while declaration values are normalized
//! once before namespace resolution.  Generic document validation (root
//! cardinality, matching end tags, and application vocabulary) belongs to the
//! caller and is intentionally outside this fixture.

use litchi_odf_common::ResolvedReader;
use litchi_odf_common::namespace::MAX_NAMESPACE_URI_BYTES;
use quick_xml::events::Event;
use quick_xml::name::{DEFAULT_MAX_DECLARATIONS_PER_ELEMENT, Namespace, ResolveResult};

#[derive(Clone, Copy)]
enum ElementEvent {
    Start,
    Empty,
    End,
}

fn expect_element(
    reader: &mut ResolvedReader<'_>,
    expected_event: ElementEvent,
    expected_name: &[u8],
    expected_namespace: ExpectedNamespace<'_>,
    expected_level: u16,
) -> quick_xml::Result<()> {
    let (namespace, event) = reader.read_resolved_event()?;
    assert_namespace(&namespace, expected_namespace);
    match (expected_event, event) {
        (ElementEvent::Start, Event::Start(element))
        | (ElementEvent::Empty, Event::Empty(element)) => {
            assert_eq!(element.name().as_ref(), expected_name);
        },
        (ElementEvent::End, Event::End(element)) => {
            assert_eq!(element.name().as_ref(), expected_name);
        },
        (ElementEvent::Start, event)
        | (ElementEvent::Empty, event)
        | (ElementEvent::End, event) => {
            panic!("expected an element event for {expected_name:?}, got {event:?}");
        },
    }
    drop(namespace);
    assert_eq!(reader.resolver().level(), expected_level);
    Ok(())
}

#[derive(Debug)]
enum ExpectedNamespace<'a> {
    Unbound,
    Bound(&'a [u8]),
    Unknown(&'a [u8]),
}

fn assert_namespace(namespace: &ResolveResult<'_>, expected: ExpectedNamespace<'_>) {
    match (namespace, expected) {
        (ResolveResult::Unbound, ExpectedNamespace::Unbound) => {},
        (ResolveResult::Bound(Namespace(actual)), ExpectedNamespace::Bound(wanted)) => {
            assert_eq!(*actual, wanted);
        },
        (ResolveResult::Unknown(actual), ExpectedNamespace::Unknown(wanted)) => {
            assert_eq!(actual.as_slice(), wanted);
        },
        (actual, expected) => panic!("namespace mismatch: expected {expected:?}, got {actual:?}"),
    }
}

fn first_error(source: &str) -> String {
    let mut reader = ResolvedReader::from_xml(source);
    match reader.read_resolved_event() {
        Ok(_) => panic!("invalid namespace source was accepted"),
        Err(error) => error.to_string(),
    }
}

#[test]
fn source_events_and_ranges_are_retained_while_namespace_values_normalize() -> quick_xml::Result<()>
{
    let source = r#"<root xmlns:p="urn:a&amp;amp;"><p:item attr="&amp;"/></root>"#;
    let source_bytes = source.as_bytes();
    let mut reader = ResolvedReader::from_xml(source);

    let start = reader.buffer_position() as usize;
    let (namespace, event) = reader.read_resolved_event()?;
    assert_namespace(&namespace, ExpectedNamespace::Unbound);
    match event {
        Event::Start(element) => {
            let raw: &[u8] = element.as_ref();
            assert_eq!(raw, b"root xmlns:p=\"urn:a&amp;amp;\"");
        },
        event => panic!("expected root start event, got {event:?}"),
    }
    drop(namespace);
    let end = reader.buffer_position() as usize;
    assert_eq!(
        &source_bytes[start..end],
        b"<root xmlns:p=\"urn:a&amp;amp;\">"
    );

    let start = reader.buffer_position() as usize;
    let (namespace, event) = reader.read_resolved_event()?;
    assert_namespace(&namespace, ExpectedNamespace::Bound(b"urn:a&amp;"));
    match event {
        Event::Empty(element) => {
            let raw: &[u8] = element.as_ref();
            assert_eq!(raw, b"p:item attr=\"&amp;\"");
        },
        event => panic!("expected child empty event, got {event:?}"),
    }
    drop(namespace);
    let end = reader.buffer_position() as usize;
    assert_eq!(&source_bytes[start..end], b"<p:item attr=\"&amp;\"/>");

    let start = reader.buffer_position() as usize;
    let (namespace, event) = reader.read_resolved_event()?;
    assert_namespace(&namespace, ExpectedNamespace::Unbound);
    match event {
        Event::End(element) => assert_eq!(element.as_ref(), b"root"),
        event => panic!("expected root end event, got {event:?}"),
    }
    drop(namespace);
    let end = reader.buffer_position() as usize;
    assert_eq!(&source_bytes[start..end], b"</root>");

    let before_eof = reader.buffer_position();
    let (namespace, event) = reader.read_resolved_event()?;
    assert_namespace(&namespace, ExpectedNamespace::Unbound);
    assert!(matches!(event, Event::Eof));
    drop(namespace);
    assert_eq!(reader.buffer_position(), before_eof);
    assert_eq!(before_eof as usize, source.len());
    Ok(())
}

#[test]
fn buffered_events_retain_the_same_lexical_source_range() -> quick_xml::Result<()> {
    let source = r#"<root xmlns:p="urn:a&amp;amp;"><p:item/></root>"#;
    let mut reader = ResolvedReader::from_xml(source);
    let mut buffer = Vec::new();
    let start = reader.buffer_position() as usize;

    {
        let (namespace, event) = reader.read_resolved_event_into(&mut buffer)?;
        assert_namespace(&namespace, ExpectedNamespace::Unbound);
        match event {
            Event::Start(element) => {
                let raw: &[u8] = element.as_ref();
                assert_eq!(raw, b"root xmlns:p=\"urn:a&amp;amp;\"");
            },
            event => panic!("expected buffered root start event, got {event:?}"),
        }
        drop(namespace);
    }

    let end = reader.buffer_position() as usize;
    assert_eq!(
        &source.as_bytes()[start..end],
        b"<root xmlns:p=\"urn:a&amp;amp;\">"
    );
    Ok(())
}

#[test]
fn borrowed_resolution_keeps_the_raw_event_and_re_resolves_the_next_event() -> quick_xml::Result<()>
{
    let source = r#"<root xmlns:p="urn:p"><p:item/></root>"#;
    let mut reader = ResolvedReader::from_xml(source);

    let (namespace, event) = reader.read_resolved_event()?;
    assert_namespace(&namespace, ExpectedNamespace::Unbound);
    match event {
        Event::Start(element) => {
            let raw: &[u8] = element.as_ref();
            assert_eq!(raw, b"root xmlns:p=\"urn:p\"");
        },
        event => panic!("expected root start event, got {event:?}"),
    }
    drop(namespace);

    let (namespace, event) = reader.read_resolved_event()?;
    assert_eq!(namespace, ResolveResult::Bound(Namespace(b"urn:p")));
    assert!(matches!(event, Event::Empty(_)));
    Ok(())
}

#[test]
fn escaped_xml_binding_is_admitted_and_literal_entity_text_decodes_once() -> quick_xml::Result<()> {
    let source = concat!(
        r#"<root xmlns:xml="http:&#x2F;&#x2F;www.w3.org&#x2F;XML&#x2F;1998&#x2F;namespace" "#,
        r#"xmlns:p="urn:a&amp;amp;"><p:item/></root>"#,
    );
    let mut reader = ResolvedReader::from_xml(source);

    let (namespace, event) = reader.read_resolved_event()?;
    assert_namespace(&namespace, ExpectedNamespace::Unbound);
    match event {
        Event::Start(element) => {
            let raw: &[u8] = element.as_ref();
            assert_eq!(
                raw,
                b"root xmlns:xml=\"http:&#x2F;&#x2F;www.w3.org&#x2F;XML&#x2F;1998&#x2F;namespace\" xmlns:p=\"urn:a&amp;amp;\"",
            );
        },
        event => panic!("expected root start event, got {event:?}"),
    }
    drop(namespace);

    let bindings = reader.resolver().bindings().collect::<Vec<_>>();
    assert_eq!(
        bindings,
        vec![(
            quick_xml::name::PrefixDeclaration::Named(b"p"),
            Namespace(b"urn:a&amp;"),
        )]
    );

    let (namespace, event) = reader.read_resolved_event()?;
    assert_namespace(&namespace, ExpectedNamespace::Bound(b"urn:a&amp;"));
    assert!(matches!(event, Event::Empty(_)));
    drop(namespace);
    Ok(())
}

#[test]
fn inherited_shadowed_and_popped_scopes_apply_to_start_and_empty_elements() -> quick_xml::Result<()>
{
    let source = concat!(
        r#"<root xmlns="urn:root" xmlns:p="urn:root:p">"#,
        r#"<branch xmlns="urn:branch" xmlns:p="urn:branch:p">"#,
        r#"<leaf/><p:leaf/></branch>"#,
        r#"<p:root-sibling/>"#,
        r#"<plain xmlns=""><p:inner/></plain>"#,
        r#"<p:after/>"#,
        r#"<empty-shadow xmlns:p="urn:empty:p"/>"#,
        r#"<p:after-empty/></root>"#,
    );
    let mut reader = ResolvedReader::from_xml(source);

    expect_element(
        &mut reader,
        ElementEvent::Start,
        b"root",
        ExpectedNamespace::Bound(b"urn:root"),
        1,
    )?;
    assert_eq!(reader.resolver().bindings().count(), 2);

    expect_element(
        &mut reader,
        ElementEvent::Start,
        b"branch",
        ExpectedNamespace::Bound(b"urn:branch"),
        2,
    )?;
    assert_eq!(reader.resolver().bindings().count(), 2);

    expect_element(
        &mut reader,
        ElementEvent::Empty,
        b"leaf",
        ExpectedNamespace::Bound(b"urn:branch"),
        3,
    )?;
    expect_element(
        &mut reader,
        ElementEvent::Empty,
        b"p:leaf",
        ExpectedNamespace::Bound(b"urn:branch:p"),
        3,
    )?;

    expect_element(
        &mut reader,
        ElementEvent::End,
        b"branch",
        ExpectedNamespace::Bound(b"urn:branch"),
        2,
    )?;
    expect_element(
        &mut reader,
        ElementEvent::Empty,
        b"p:root-sibling",
        ExpectedNamespace::Bound(b"urn:root:p"),
        2,
    )?;

    expect_element(
        &mut reader,
        ElementEvent::Start,
        b"plain",
        ExpectedNamespace::Unbound,
        2,
    )?;
    expect_element(
        &mut reader,
        ElementEvent::Empty,
        b"p:inner",
        ExpectedNamespace::Bound(b"urn:root:p"),
        3,
    )?;
    expect_element(
        &mut reader,
        ElementEvent::End,
        b"plain",
        ExpectedNamespace::Unbound,
        2,
    )?;
    expect_element(
        &mut reader,
        ElementEvent::Empty,
        b"p:after",
        ExpectedNamespace::Bound(b"urn:root:p"),
        2,
    )?;

    expect_element(
        &mut reader,
        ElementEvent::Empty,
        b"empty-shadow",
        ExpectedNamespace::Bound(b"urn:root"),
        2,
    )?;
    expect_element(
        &mut reader,
        ElementEvent::Empty,
        b"p:after-empty",
        ExpectedNamespace::Bound(b"urn:root:p"),
        2,
    )?;

    expect_element(
        &mut reader,
        ElementEvent::End,
        b"root",
        ExpectedNamespace::Bound(b"urn:root"),
        1,
    )?;
    let (namespace, event) = reader.read_resolved_event()?;
    assert_namespace(&namespace, ExpectedNamespace::Unbound);
    assert!(matches!(event, Event::Eof));
    drop(namespace);
    assert_eq!(reader.resolver().level(), 0);
    assert_eq!(reader.resolver().bindings().count(), 0);
    Ok(())
}

#[test]
fn unknown_prefixes_are_reported_without_changing_source_events() -> quick_xml::Result<()> {
    let source = r#"<root xmlns:p="urn:p"><q:item/></root>"#;
    let mut reader = ResolvedReader::from_xml(source);

    let (namespace, event) = reader.read_resolved_event()?;
    drop(namespace);
    assert!(matches!(event, Event::Start(_)));
    let (namespace, event) = reader.read_resolved_event()?;
    assert_namespace(&namespace, ExpectedNamespace::Unknown(b"q"));
    match event {
        Event::Empty(element) => {
            let raw: &[u8] = element.as_ref();
            assert_eq!(raw, b"q:item");
        },
        event => panic!("expected unknown-prefix empty event, got {event:?}"),
    }
    drop(namespace);
    Ok(())
}

#[test]
fn reserved_and_empty_namespace_bindings_are_rejected_after_normalization() {
    let invalid_sources = [
        (
            "default XML namespace",
            r#"<root xmlns="http://www.w3.org/XML/1998/namespace"/>"#,
        ),
        (
            "default XMLNS namespace",
            r#"<root xmlns="http://www.w3.org/2000/xmlns/"/>"#,
        ),
        (
            "escaped default XML namespace",
            r#"<root xmlns="http:&#x2F;&#x2F;www.w3.org&#x2F;XML&#x2F;1998&#x2F;namespace"/>"#,
        ),
        (
            "escaped default XMLNS namespace",
            r#"<root xmlns="http:&#x2F;&#x2F;www.w3.org&#x2F;2000&#x2F;xmlns&#x2F;"/>"#,
        ),
        (
            "named XML namespace",
            r#"<root xmlns:p="http://www.w3.org/XML/1998/namespace"/>"#,
        ),
        (
            "named XMLNS namespace",
            r#"<root xmlns:p="http://www.w3.org/2000/xmlns/"/>"#,
        ),
        ("named empty namespace", r#"<root xmlns:p=""/>"#),
        ("empty named prefix", r#"<root xmlns:="urn:empty-prefix"/>"#),
    ];

    for (label, source) in invalid_sources {
        let error = first_error(source);
        assert!(!error.is_empty(), "{label} produced an empty error");
    }
}

#[test]
fn unknown_entities_raw_delimiters_and_numeric_plus_forms_are_rejected() {
    let invalid_sources = [
        ("unknown entity", r#"<root xmlns:p="urn:a&unknown;"/>"#),
        ("unterminated entity", r#"<root xmlns:p="urn:a&unknown"/>"#),
        ("raw less-than delimiter", r#"<root xmlns:p="urn:a<raw"/>"#),
        (
            "empty hexadecimal reference",
            r#"<root xmlns:p="urn:a&#x;"/>"#,
        ),
        (
            "plus-prefixed hexadecimal reference",
            r#"<root xmlns:p="urn:a&#x+41;"/>"#,
        ),
        (
            "plus-prefixed decimal reference",
            r#"<root xmlns:p="urn:a&#+65;"/>"#,
        ),
    ];

    for (label, source) in invalid_sources {
        let error = first_error(source);
        assert!(!error.is_empty(), "{label} produced an empty error");
    }
}

#[test]
fn duplicate_and_malformed_namespace_attributes_are_rejected() {
    let invalid_sources = [
        (
            "duplicate named namespace",
            r#"<root xmlns:p="urn:first" xmlns:p="urn:second"/>"#,
        ),
        (
            "duplicate default namespace",
            r#"<root xmlns="urn:first" xmlns="urn:second"/>"#,
        ),
        ("namespace without a value", r#"<root xmlns:p/>"#),
        (
            "namespace with an unquoted value",
            r#"<root xmlns:p=urn:value/>"#,
        ),
        (
            "namespace with a missing closing quote",
            r#"<root xmlns:p="urn:value/>"#,
        ),
    ];

    for (label, source) in invalid_sources {
        let error = first_error(source);
        assert!(!error.is_empty(), "{label} produced an empty error");
    }
}

#[test]
fn namespace_declarations_have_a_bounded_per_element_profile() -> quick_xml::Result<()> {
    let mut accepted = String::from("<root");
    for index in 0..DEFAULT_MAX_DECLARATIONS_PER_ELEMENT {
        accepted.push_str(" xmlns:p");
        accepted.push_str(&index.to_string());
        accepted.push_str("=\"urn:");
        accepted.push_str(&index.to_string());
        accepted.push('"');
    }
    accepted.push_str("/>");

    let mut reader = ResolvedReader::from_xml(&accepted);
    let (namespace, event) = reader.read_resolved_event()?;
    assert_namespace(&namespace, ExpectedNamespace::Unbound);
    assert!(matches!(event, Event::Empty(_)));
    drop(namespace);
    assert_eq!(
        reader.resolver().bindings().count(),
        DEFAULT_MAX_DECLARATIONS_PER_ELEMENT
    );

    let mut rejected = String::from("<root");
    for index in 0..=DEFAULT_MAX_DECLARATIONS_PER_ELEMENT {
        rejected.push_str(" xmlns:p");
        rejected.push_str(&index.to_string());
        rejected.push_str("=\"urn:");
        rejected.push_str(&index.to_string());
        rejected.push('"');
    }
    rejected.push_str("/>");
    let error = first_error(&rejected);
    assert!(error.contains("namespace") || error.contains("declaration"));
    Ok(())
}

#[test]
fn nested_namespace_scopes_stop_at_the_bounded_depth_profile() -> quick_xml::Result<()> {
    // The adapter's retained scope profile is four times quick-xml's
    // per-element declaration profile.  Keep this as a stream test: matching
    // end tags are the caller's structural-validation responsibility.
    let max_depth = DEFAULT_MAX_DECLARATIONS_PER_ELEMENT * 4;
    let mut source = String::with_capacity((max_depth + 1) * 3);
    for _ in 0..=max_depth {
        source.push_str("<e>");
    }

    let mut reader = ResolvedReader::from_xml(&source);
    for expected_level in 1..=max_depth {
        let (namespace, event) = reader.read_resolved_event()?;
        assert_namespace(&namespace, ExpectedNamespace::Unbound);
        assert!(matches!(event, Event::Start(_)));
        drop(namespace);
        assert_eq!(usize::from(reader.resolver().level()), expected_level);
    }

    let error = match reader.read_resolved_event() {
        Ok((namespace, event)) => {
            panic!("depth-over-limit event was accepted: {namespace:?}, {event:?}");
        },
        Err(error) => error.to_string(),
    };
    assert!(error.contains("namespace") || error.contains("declaration"));
    Ok(())
}

#[test]
fn active_namespace_bytes_are_bounded_before_resolver_growth() {
    let uri = "u".repeat(MAX_NAMESPACE_URI_BYTES);
    let mut source = String::with_capacity(
        (DEFAULT_MAX_DECLARATIONS_PER_ELEMENT * (uri.len() + 24)).saturating_add(16),
    );
    source.push_str("<root");
    for index in 0..DEFAULT_MAX_DECLARATIONS_PER_ELEMENT {
        source.push_str(" xmlns:p");
        source.push_str(&index.to_string());
        source.push_str("=\"");
        source.push_str(&uri);
        source.push('"');
    }
    source.push_str("/>");

    let error = first_error(&source);
    assert!(error.contains("active namespace") || error.contains("namespace byte"));
    assert!(
        error.len() <= 4096,
        "active-scope error appeared to retain an input payload: {} bytes",
        error.len()
    );
}

#[test]
fn namespace_uri_and_prefix_errors_do_not_echo_oversized_payloads() {
    let oversized_uri = "u".repeat(MAX_NAMESPACE_URI_BYTES + 1);
    let uri_source = format!(r#"<root xmlns:p="{oversized_uri}"/>"#);
    let uri_error = first_error(&uri_source);
    assert!(
        uri_error.len() <= 4096,
        "oversized URI appeared in the error payload: {} bytes",
        uri_error.len()
    );

    let oversized_prefix = "p".repeat(MAX_NAMESPACE_URI_BYTES + 1);
    let prefix_source = format!(r#"<root xmlns:{oversized_prefix}="urn:p"/>"#);
    let prefix_error = first_error(&prefix_source);
    assert!(
        prefix_error.len() <= 4096,
        "oversized prefix appeared in the error payload: {} bytes",
        prefix_error.len()
    );
}
