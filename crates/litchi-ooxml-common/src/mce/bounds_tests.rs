//! Record 0764: the per-element attribute limit and the bounded namespace and
//! directive work of both MCE processors on hostile documents.
//!
//! Work is bounded by counting operations rather than by measuring time: each
//! prefix-index lookup (a logarithmic B-tree operation) and each declaration a
//! lookup walks before it asks the index, at most
//! [`WALKED_DECLARATIONS`] per lookup.

use std::{convert::Infallible, io::Cursor};

use super::codec::process_markup_compatibility;
use super::model::{Capabilities, DEFAULT_MAX_ATTRIBUTES_PER_ELEMENT, Error, Limits, Report};
use super::scope::WALKED_DECLARATIONS;
use super::scope::counter::{counted, counted_uri_bytes};
use super::stream::{
    SemanticEvent, StreamError, StreamLimits, process_markup_compatibility_stream,
    process_markup_compatibility_stream_with_observers,
};

const MC: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";

/// `count` distinct attributes, each preceded by a space.
fn attributes(count: usize) -> String {
    (0..count).map(|index| format!(r#" a{index}="""#)).collect()
}

/// A document the codec parses (it names the MCE namespace) whose second
/// element carries `attributes`.
fn with_tag(attributes: &str) -> String {
    format!(r#"<r xmlns:mc="{MC}"><t{attributes}/></r>"#)
}

fn codec(xml: &str, limits: &Limits) -> Result<(String, Report), Error> {
    let output = process_markup_compatibility(xml.as_bytes(), &Capabilities::new(), limits)?;
    Ok((
        String::from_utf8(output.xml.into_owned()).expect("MCE output stays UTF-8"),
        output.report,
    ))
}

/// The expanded names of the start and empty events the stream reports.
fn stream(
    xml: &str,
    limits: &StreamLimits,
) -> Result<Vec<String>, StreamError<Infallible, Infallible>> {
    let mut input = Cursor::new(xml.as_bytes());
    let mut names = Vec::new();
    process_markup_compatibility_stream(&mut input, &Capabilities::new(), limits, |event| {
        if let SemanticEvent::Start(element) | SemanticEvent::Empty(element) = &event {
            names.push(format!(
                "{{{}}}{}",
                element.expanded_name.namespace, element.expanded_name.local_name
            ));
        }
        Ok::<(), Infallible>(())
    })?;
    Ok(names)
}

fn is_limit(error: &Error, label: &str) -> bool {
    matches!(error, Error::LimitExceeded(message) if message == label)
}

#[test]
fn the_codec_admits_a_tag_at_the_limit_and_refuses_one_more() {
    let limits = Limits {
        max_attributes_per_element: 8,
        ..Limits::default()
    };
    let (output, _) = codec(&with_tag(&attributes(8)), &limits).expect("eight attributes fit");
    assert!(output.contains(r#"a7="""#), "{output}");
    let error = codec(&with_tag(&attributes(9)), &limits).expect_err("nine do not");
    assert!(is_limit(&error, "attributes per element"), "{error:?}");

    let defaults = Limits::default();
    assert_eq!(
        defaults.max_attributes_per_element,
        DEFAULT_MAX_ATTRIBUTES_PER_ELEMENT
    );
    codec(
        &with_tag(&attributes(DEFAULT_MAX_ATTRIBUTES_PER_ELEMENT)),
        &defaults,
    )
    .expect("a tag at the default limit is admitted");
    let error = codec(
        &with_tag(&attributes(DEFAULT_MAX_ATTRIBUTES_PER_ELEMENT + 1)),
        &defaults,
    )
    .expect_err("one attribute more is refused");
    assert!(is_limit(&error, "attributes per element"), "{error:?}");
}

#[test]
fn a_duplicate_flood_is_refused_at_the_limit_before_its_duplicates() {
    // 20,000 distinct names and 20,000 repeats of the last: the checked
    // iterator would scan the distinct names for each repeat. The limit
    // refuses the tag at its 1,025th attribute instead.
    let mut tag = attributes(20_000);
    for _ in 0..20_000 {
        tag.push_str(r#" a19999="""#);
    }
    let error = codec(&with_tag(&tag), &Limits::default()).expect_err("refused");
    assert!(is_limit(&error, "attributes per element"), "{error:?}");
    let error = stream(&with_tag(&tag), &StreamLimits::default()).expect_err("refused");
    assert!(
        matches!(&error, StreamError::Mce { error, .. } if is_limit(error, "attributes per element")),
        "{error:?}"
    );
}

/// `depth` nested elements that each re-declare the prefixes `q0..q{width}`,
/// under a root binding `z`, around an element of `attributes` attributes in
/// the `z` namespace. Resolving `z` by walking the declarations in scope costs
/// `depth * width` comparisons per attribute.
fn shadowed_chain(depth: usize, width: usize, attributes: usize) -> String {
    let mut xml = format!(r#"<r xmlns:mc="{MC}" xmlns:z="urn:z">"#);
    for level in 0..depth {
        xml.push_str("<s");
        for prefix in 0..width {
            xml.push_str(&format!(r#" xmlns:q{prefix}="urn:{level}:{prefix}""#));
        }
        xml.push('>');
    }
    xml.push_str("<z:e");
    for index in 0..attributes {
        xml.push_str(&format!(r#" z:a{index}="""#));
    }
    xml.push_str("/>");
    for _ in 0..depth {
        xml.push_str("</s>");
    }
    xml.push_str("</r>");
    xml
}

#[test]
fn lookups_do_not_walk_shadowed_declarations() {
    let (depth, width, attributes) = (64, 500, 1_000);
    let xml = shadowed_chain(depth, width, attributes);
    // Every declaration is resolved once in the index to count new bindings;
    // every element and attribute name, at most three times, by a walk of at
    // most WALKED_DECLARATIONS declarations and one index lookup. A chain walk
    // would compare `depth * width` declarations for each attribute name.
    let names = attributes + depth + 4;
    let bound = depth * width + 3 * (WALKED_DECLARATIONS + 1) * names;
    let walked = depth * width * attributes;
    assert!(bound * 100 < walked);

    let (result, operations) = counted(|| codec(&xml, &Limits::default()));
    let (output, _) = result.expect("the document is valid");
    assert!(output.ends_with("</s></r>"), "{output}");
    assert!(
        operations <= bound,
        "{operations} operations, bound {bound}"
    );

    let (result, operations) = counted(|| stream(&xml, &StreamLimits::default()));
    let events = result.expect("the stream accepts the document");
    assert_eq!(events.last().map(String::as_str), Some("{urn:z}e"));
    assert!(
        operations <= bound,
        "{operations} operations, bound {bound}"
    );
}

#[test]
fn a_declaration_flood_at_the_limit_costs_a_lookup_per_declaration() {
    let count = DEFAULT_MAX_ATTRIBUTES_PER_ELEMENT;
    let declarations: String = (0..count)
        .map(|index| format!(r#" xmlns:p{index}="urn:{index}""#))
        .collect();
    let xml = with_tag(&declarations);
    let (result, lookups) = counted(|| codec(&xml, &Limits::default()));
    let (output, _) = result.expect("declarations up to the limit are admitted");
    assert!(output.contains(&format!(r#"xmlns:p{}="urn:{}""#, count - 1, count - 1)));
    assert!(lookups <= 3 * count, "{lookups} lookups");

    let over = with_tag(&format!(r#"{declarations} xmlns:extra="urn:extra""#));
    let error = codec(&over, &Limits::default()).expect_err("one more is refused");
    assert!(is_limit(&error, "attributes per element"), "{error:?}");
}

#[test]
fn hoisting_many_shadowed_declarations_onto_many_children_is_output_bound() {
    // Five nested AlternateContent wrappers whose Fallback branch is
    // selected; each wrapper and each Fallback re-declares p0..p99, and 500
    // emitted children re-declare the innermost bindings.
    let (levels, width, children) = (5, 100, 500);
    let declarations = |layer: usize| -> String {
        (0..width)
            .map(|prefix| format!(r#" xmlns:p{prefix}="urn:{layer}:{prefix}""#))
            .collect()
    };
    let mut xml = format!(r#"<r xmlns:mc="{MC}" xmlns:i="urn:i" xmlns:w="urn:w">"#);
    for level in 0..levels {
        xml.push_str(&format!(
            r#"<mc:AlternateContent{}><mc:Choice Requires="i"/><mc:Fallback{}>"#,
            declarations(2 * level),
            declarations(2 * level + 1)
        ));
    }
    for _ in 0..children {
        xml.push_str("<w:x/>");
    }
    for _ in 0..levels {
        xml.push_str("</mc:Fallback></mc:AlternateContent>");
    }
    xml.push_str("</r>");

    let innermost = 2 * levels - 1;
    let hoisted: String = (0..width)
        .map(|prefix| format!(r#" xmlns:p{prefix}="urn:{innermost}:{prefix}""#))
        .collect();
    let child = format!("<w:x{hoisted}></w:x>");
    let expected = format!(
        r#"<r xmlns:mc="{MC}" xmlns:i="urn:i" xmlns:w="urn:w">{}</r>"#,
        child.repeat(children)
    );

    let (result, lookups) = counted(|| codec(&xml, &Limits::default()));
    let (output, report) = result.expect("the document is valid");
    assert_eq!(output, expected);
    assert_eq!(report.alternate_content_count, levels);
    assert_eq!(report.selected_fallbacks, levels);
    // One own-declaration check per hoisted binding per child, plus the
    // declarations and names themselves; walking the dropped scopes for
    // every child would visit `2 * levels * width` declarations each.
    let layers = 2 * levels;
    let bound = 2 * children * width + 4 * layers * width + 4 * children + 64;
    assert!(lookups <= bound, "{lookups} lookups, bound {bound}");
    assert!(bound * 4 < children * layers * width);
}

#[test]
fn directive_targets_at_scale_match_exact_names_and_wildcards() {
    // 3,000 exact PreserveAttributes targets and one wildcard namespace.
    let targets: Vec<String> = (0..3_000).map(|index| format!("i:k{index}")).collect();
    let xml = format!(
        r#"<r xmlns:mc="{MC}" xmlns:i="urn:i" xmlns:j="urn:j" mc:Ignorable="i j" mc:PreserveAttributes="{} j:*"><e i:k0="1" i:k2999="2" i:other="3" j:any="4"/></r>"#,
        targets.join(" ")
    );
    let (output, report) = codec(&xml, &Limits::default()).expect("valid directives");
    assert!(output.contains(r#"i:k0="1""#), "{output}");
    assert!(output.contains(r#"i:k2999="2""#), "{output}");
    assert!(!output.contains("i:other"), "{output}");
    assert!(output.contains(r#"j:any="4""#), "{output}");
    assert_eq!(report.preserved_attributes, 3);

    let events = stream(&xml, &StreamLimits::default()).expect("the stream accepts it");
    assert_eq!(events, vec!["{}r".to_owned(), "{}e".to_owned()]);
}

#[test]
fn the_stream_applies_the_tighter_of_its_two_attribute_limits() {
    let processing = Limits {
        max_attributes_per_element: 4,
        ..Limits::default()
    };
    let limits = StreamLimits::new(processing);
    stream(&with_tag(&attributes(4)), &limits).expect("four attributes fit");
    let error = stream(&with_tag(&attributes(5)), &limits).expect_err("five do not");
    assert!(
        matches!(&error, StreamError::Mce { error, .. } if is_limit(error, "attributes per element")),
        "{error:?}"
    );

    let event_limited = StreamLimits {
        max_attributes_per_event: 3,
        ..StreamLimits::default()
    };
    let error = stream(&with_tag(&attributes(4)), &event_limited).expect_err("refused");
    assert!(
        matches!(&error, StreamError::Mce { error, .. } if is_limit(error, "stream attributes per event")),
        "{error:?}"
    );
}

#[test]
fn an_empty_elements_declarations_leave_the_scope_when_the_raw_flow_recovers() {
    // The empty AlternateContent is refused, the raw flow recovers and reads
    // on, and `p`, which only the empty element declared, is unbound after it.
    let xml = format!(r#"<r xmlns:mc="{MC}"><mc:AlternateContent xmlns:p="urn:p"/><p:x/></r>"#);
    let mut input = Cursor::new(xml.as_bytes());
    let mut raw = Vec::new();
    let error = process_markup_compatibility_stream_with_observers(
        &mut input,
        &Capabilities::new(),
        &StreamLimits::default(),
        |element| {
            raw.push(String::from_utf8_lossy(element.name()).into_owned());
            Ok::<(), Infallible>(())
        },
        |_event| Ok::<(), Infallible>(()),
    )
    .expect_err("the document is refused");
    assert_eq!(raw, ["r", "mc:AlternateContent"]);
    assert!(
        matches!(
            &error,
            StreamError::Mce {
                error: Error::NonConformant(message),
                prior_mce_error: Some(Error::NonConformant(prior)),
                ..
            } if message == "unbound prefix" && prior == "empty AlternateContent"
        ),
        "{error:?}"
    );
}

#[test]
fn no_configuration_admits_more_than_the_ceiling() {
    use super::model::ATTRIBUTES_PER_ELEMENT_CEILING;

    let unbounded = Limits {
        max_attributes_per_element: usize::MAX,
        ..Limits::default()
    };
    // The in-memory processor applies the ceiling whatever the field holds.
    codec(
        &with_tag(&attributes(ATTRIBUTES_PER_ELEMENT_CEILING)),
        &unbounded,
    )
    .expect("the ceiling is admitted");
    let error = codec(
        &with_tag(&attributes(ATTRIBUTES_PER_ELEMENT_CEILING + 1)),
        &unbounded,
    )
    .expect_err("one attribute more is refused");
    assert!(is_limit(&error, "attributes per element"), "{error:?}");

    // The stream refuses such a configuration before reading.
    let stream_limits = StreamLimits::new(unbounded);
    let error = stream(&with_tag(&attributes(1)), &stream_limits).expect_err("refused");
    assert!(
        matches!(&error, StreamError::Mce { error, .. } if is_limit(error, "MCE attributes per element")),
        "{error:?}"
    );
    let per_event = StreamLimits {
        max_attributes_per_event: ATTRIBUTES_PER_ELEMENT_CEILING + 1,
        ..StreamLimits::default()
    };
    let error = stream(&with_tag(&attributes(1)), &per_event).expect_err("refused");
    assert!(
        matches!(&error, StreamError::Mce { error, .. } if is_limit(error, "stream attributes per event")),
        "{error:?}"
    );
    // At the ceiling both limits are valid.
    let widest = StreamLimits {
        max_attributes_per_event: ATTRIBUTES_PER_ELEMENT_CEILING,
        ..StreamLimits::new(Limits {
            max_attributes_per_element: ATTRIBUTES_PER_ELEMENT_CEILING,
            ..Limits::default()
        })
    };
    stream(
        &with_tag(&attributes(ATTRIBUTES_PER_ELEMENT_CEILING)),
        &widest,
    )
    .expect("the ceiling is admitted by the stream");
}

/// A root binding `z` to a URI of `uri_bytes` bytes, ignorable when
/// `directives` says so, around one element of `attributes` attributes in
/// `z`.
fn long_uri(uri_bytes: usize, directives: &str, attributes: usize) -> String {
    let uri = format!("urn:{}", "u".repeat(uri_bytes - 4));
    let names: String = (0..attributes)
        .map(|index| format!(r#" z:a{index}="""#))
        .collect();
    format!(r#"<r xmlns:mc="{MC}" xmlns:z="{uri}"{directives}><e{names}/></r>"#)
}

#[test]
fn a_long_namespace_uri_is_read_once_per_declaration_not_per_name() {
    // Before namespace identities, an attribute in an ignorable namespace
    // hashed its URI in each directive and capability check: 1,000
    // attributes under a 4 MiB URI took 1.4 s. Now the URI is hashed and its
    // facts computed once, when it is declared.
    let attributes = 1_000;
    for directives in [
        "",
        r#" mc:Ignorable="z""#,
        r#" mc:Ignorable="z" mc:PreserveAttributes="z:*""#,
    ] {
        let mut outputs = Vec::new();
        let mut operations = Vec::new();
        for uri_bytes in [64, 64 * 1024] {
            let xml = long_uri(uri_bytes, directives, attributes);
            let ((result, lookups), uri_read) =
                counted_uri_bytes(|| counted(|| codec(&xml, &Limits::default())));
            let (output, report) = result.expect("the document is valid");
            // The declaration of `z` is interned (hashed, compared with the
            // same-hash URIs in scope, of which there are none) and its facts
            // computed: a constant number of passes over the URI, whatever
            // the number of attributes in its namespace.
            assert!(
                uri_read <= 4 * uri_bytes + 1024,
                "{directives}: {uri_read} bytes"
            );
            outputs.push((output.replace(&"u".repeat(uri_bytes - 4), "U"), report));
            operations.push(lookups);
        }
        // The same work and the same output, the URI aside.
        assert_eq!(operations[0], operations[1], "{directives}");
        assert_eq!(outputs[0], outputs[1], "{directives}");
    }
}
