//! The per-element attribute limit ([`Resource::ElementAttributes`]).
//!
//! Every policy, slice and stream alike, refuses a start or empty-element tag
//! at its first attribute beyond the limit, with a typed limit error that
//! names the resource, the limit, the observed count and the offset of the
//! surplus attribute. The limit is independent of the document-wide attribute
//! budget, does not count an XML declaration's pseudo-attributes, and a
//! replacement audit returns the verdict of the two complete audits.
#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "each test states one fixed expected verdict"
)]

use std::io::Cursor;

use xml_minifier::audit::{
    self, Error, Limits, ReplacementError, Resource, StreamError, verify_source,
    verify_source_replacement,
};

/// `count` distinct attributes, each separated by one space.
fn attributes(count: usize) -> String {
    (0..count)
        .map(|index| format!(r#"a{index}="""#))
        .collect::<Vec<_>>()
        .join(" ")
}

/// A root element with `count` attributes, as a start tag or an empty tag.
fn root(count: usize, empty: bool) -> String {
    let separator = if count == 0 { "" } else { " " };
    let attributes = attributes(count);
    if empty {
        format!("<root{separator}{attributes}/>")
    } else {
        format!("<root{separator}{attributes}></root>")
    }
}

fn limits(maximum: usize) -> Limits {
    Limits::builder()
        .element_attributes(maximum)
        .expect("a limit at or below the ceiling")
        .build()
}

fn stream(result: Result<audit::Report, StreamError>) -> Result<audit::Report, Error> {
    match result {
        Ok(report) => Ok(report),
        Err(StreamError::Audit(error)) => Err(error),
        Err(other) => panic!("unexpected stream failure: {other}"),
    }
}

/// Every auditor's verdict on `xml`, in a fixed order.
fn verdicts(xml: &[u8], limits: Limits) -> [Result<audit::Report, Error>; 5] {
    [
        verify_source(xml, limits),
        audit::verify_authored(xml, limits),
        audit::verify(xml, limits),
        stream(audit::verify_authored_reader(Cursor::new(xml), limits)),
        stream(audit::verify_reader(Cursor::new(xml), limits)),
    ]
}

fn assert_refused_at(result: &Result<audit::Report, Error>, maximum: usize, offset: usize) {
    assert!(
        matches!(
            result,
            Err(Error::Limit {
                resource: Resource::ElementAttributes,
                limit,
                actual,
                offset: at,
            }) if *limit == maximum && *actual == maximum + 1 && *at == offset
        ),
        "expected a per-element limit refusal at {offset}, got {result:?}"
    );
}

#[test]
fn default_limit_is_documented_and_within_the_ceiling() {
    let defaults = Limits::default();
    assert_eq!(
        defaults.max_element_attributes(),
        Limits::DEFAULT_ELEMENT_ATTRIBUTES
    );
    const { assert!(Limits::DEFAULT_ELEMENT_ATTRIBUTES <= Limits::ELEMENT_ATTRIBUTE_CEILING) };
    assert_eq!(
        Limits::ceiling(Resource::ElementAttributes),
        Limits::ELEMENT_ATTRIBUTE_CEILING
    );
    // The positional constructor keeps its six parameters and takes the
    // per-element default.
    let positional = Limits::new(1024, 8, 64, 64, 1024, 1024).unwrap();
    assert_eq!(
        positional.max_element_attributes(),
        Limits::DEFAULT_ELEMENT_ATTRIBUTES
    );
}

#[test]
fn a_tag_at_the_limit_is_accepted_and_one_more_is_refused_at_the_surplus_attribute() {
    for maximum in [0, 1, 2, 31, 32, 33, 64, Limits::DEFAULT_ELEMENT_ATTRIBUTES] {
        for empty in [false, true] {
            let exact = root(maximum, empty);
            for verdict in verdicts(exact.as_bytes(), limits(maximum)) {
                let report = verdict.expect("a tag at the limit is accepted");
                assert_eq!(report.attributes(), maximum);
            }

            let over = root(maximum + 1, empty);
            let surplus = format!("a{maximum}=");
            let offset = over.find(&surplus).expect("surplus attribute");
            for verdict in verdicts(over.as_bytes(), limits(maximum)) {
                assert_refused_at(&verdict, maximum, offset);
            }
        }
    }
}

#[test]
fn the_limit_applies_to_every_tag_and_not_to_the_document_total() {
    // Three children with the limit's attributes each: accepted, although
    // the document holds three times the per-element limit.
    let maximum = 8;
    let child = format!("<c {}/>", attributes(maximum));
    let document = format!("<root>{child}{child}{child}</root>");
    for verdict in verdicts(document.as_bytes(), limits(maximum)) {
        assert_eq!(
            verdict.expect("each tag is within the limit").attributes(),
            3 * maximum
        );
    }

    // The third child carries one attribute more.
    let surplus_child = format!("<c {} extra=\"\"/>", attributes(maximum));
    let document = format!("<root>{child}{child}{surplus_child}</root>");
    let offset = document.find("extra=").expect("surplus attribute");
    for verdict in verdicts(document.as_bytes(), limits(maximum)) {
        assert_refused_at(&verdict, maximum, offset);
    }
}

#[test]
fn namespace_declarations_count_as_attributes() {
    let declarations = (0..4)
        .map(|index| format!(r#"xmlns:p{index}="urn:{index}""#))
        .collect::<Vec<_>>()
        .join(" ");
    let document = format!("<root {declarations}/>");
    for verdict in verdicts(document.as_bytes(), limits(4)) {
        assert_eq!(verdict.expect("four declarations fit").attributes(), 4);
    }
    let offset = document.find("xmlns:p3").expect("fourth declaration");
    for verdict in verdicts(document.as_bytes(), limits(3)) {
        assert_refused_at(&verdict, 3, offset);
    }
}

#[test]
fn the_declaration_pseudo_attributes_are_not_counted() {
    let document = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><root/>"#;
    for verdict in verdicts(document, limits(0)) {
        assert_eq!(
            verdict
                .expect("a declaration is not an element")
                .attributes(),
            0
        );
    }
}

#[test]
fn a_huge_tag_is_refused_at_the_first_surplus_attribute_before_its_duplicates() {
    // One tag of 200,000 attributes whose later half repeats one name. The
    // refusal is the limit at attribute 1,025, not the duplicate further on:
    // nothing after the surplus attribute is read.
    let distinct = 100_000;
    let mut tag = String::from("<root ");
    tag.push_str(&attributes(distinct));
    for _ in 0..distinct {
        tag.push_str(r#" a0="""#);
    }
    tag.push_str("/>");
    let maximum = Limits::DEFAULT_ELEMENT_ATTRIBUTES;
    let surplus = format!(" a{maximum}=");
    let offset = tag.find(&surplus).expect("surplus attribute") + 1;
    for verdict in verdicts(tag.as_bytes(), Limits::default()) {
        assert_refused_at(&verdict, maximum, offset);
    }
}

#[test]
fn builder_ceiling_and_narrowing_are_typed() {
    let ceiling = Limits::ELEMENT_ATTRIBUTE_CEILING;
    let exact = Limits::builder()
        .limit(Resource::ElementAttributes, ceiling)
        .unwrap()
        .build();
    assert_eq!(exact.max_element_attributes(), ceiling);
    let error = Limits::builder()
        .element_attributes(ceiling + 1)
        .unwrap_err();
    assert_eq!(error.resource(), Resource::ElementAttributes);
    assert_eq!(error.requested(), ceiling + 1);
    assert_eq!(error.ceiling(), ceiling);

    let narrowed = Limits::default().narrow(Resource::ElementAttributes, 3);
    assert_eq!(narrowed.max_element_attributes(), 3);
    assert_eq!(
        narrowed
            .narrow(Resource::ElementAttributes, usize::MAX)
            .max_element_attributes(),
        3
    );
    // Narrowing one resource leaves the others unchanged.
    assert_eq!(
        narrowed.max_attributes(),
        Limits::default().max_attributes()
    );
}

#[test]
fn a_replacement_that_adds_a_surplus_attribute_gets_the_complete_audits_verdict() {
    let maximum = 4;
    let within = format!(
        "<root><a {}/><b {}/></root>",
        attributes(maximum),
        attributes(2)
    );
    let over = format!(
        "<root><a {}/><b {} x=\"\"/></root>",
        attributes(maximum),
        attributes(maximum)
    );
    let limits = limits(maximum);
    let pair = verify_source_replacement(within.as_bytes(), over.as_bytes(), limits);
    let complete = verify_source(over.as_bytes(), limits).unwrap_err();
    assert_eq!(pair, Err(ReplacementError::Replacement(complete)));
    let offset = over.rfind("x=").expect("surplus attribute");
    assert_refused_at(&verify_source(over.as_bytes(), limits), maximum, offset);

    // And the other way round: the original is refused first.
    let pair = verify_source_replacement(over.as_bytes(), within.as_bytes(), limits);
    assert!(matches!(
        pair,
        Err(ReplacementError::Original(Error::Limit {
            resource: Resource::ElementAttributes,
            ..
        }))
    ));
}
