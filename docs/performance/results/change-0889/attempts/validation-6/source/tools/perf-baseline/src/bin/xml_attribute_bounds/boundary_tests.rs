//! Independent exact and one-over boundary qualification for the MCE APIs.
//!
//! The fixtures in this module are deliberately small and their expected
//! outcomes are chosen from the public contract, rather than from a result
//! produced by the implementation under test.  The tests exercise both the
//! byte-buffer processor and the stream processor where both expose the same
//! policy; stream-only tests cover the event-facing budgets.

use std::convert::Infallible;
use std::io::Cursor;

use litchi_ooxml_common::mce::{
    Capabilities, Error, Limits, StreamError, StreamLimits, StreamReport,
    process_markup_compatibility, process_markup_compatibility_stream,
};

const MC: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";

fn codec(input: &[u8], limits: &Limits) -> Result<Vec<u8>, Error> {
    process_markup_compatibility(input, &Capabilities::new(), limits)
        .map(|output| output.xml.into_owned())
}

fn stream(
    input: &[u8],
    limits: &StreamLimits,
) -> Result<StreamReport, StreamError<Infallible, Infallible>> {
    let mut input = Cursor::new(input);
    process_markup_compatibility_stream(&mut input, &Capabilities::new(), limits, |_| {
        Ok::<(), Infallible>(())
    })
}

fn assert_codec_limit(result: Result<Vec<u8>, Error>, label: &str) {
    assert!(
        matches!(&result, Err(Error::LimitExceeded(message)) if message == label),
        "expected codec limit {label:?}, got {result:?}"
    );
}

fn assert_stream_limit(
    result: Result<StreamReport, StreamError<Infallible, Infallible>>,
    label: &str,
) {
    assert!(
        matches!(
            &result,
            Err(StreamError::Mce {
                error: Error::LimitExceeded(message),
                ..
            }) if message == label
        ),
        "expected stream limit {label:?}, got {result:?}"
    );
}

#[test]
fn namespace_budget_counts_new_prefixes_but_not_same_prefix_rebindings() {
    // The implicit `xml` binding counts as one.  The root adds `mc` and `a`,
    // so the exact budget is three.  Rebinding `a` changes its URI but does
    // not add a prefix identity; introducing `b` does.
    let rebind = format!(r#"<r xmlns:mc="{MC}" xmlns:a="urn:a"><s xmlns:a="urn:b"></s></r>"#);
    let new_prefix = format!(r#"<r xmlns:mc="{MC}" xmlns:a="urn:a"><s xmlns:b="urn:b"></s></r>"#);

    let exact = Limits {
        max_namespace_bindings: 3,
        ..Limits::default()
    };
    codec(rebind.as_bytes(), &exact).expect("same-prefix rebind fits the exact budget");
    assert_codec_limit(codec(new_prefix.as_bytes(), &exact), "namespace bindings");

    let stream_exact = StreamLimits {
        processing: exact.clone(),
        ..StreamLimits::default()
    };
    stream(rebind.as_bytes(), &stream_exact).expect("stream rebind fits the exact budget");
    assert_stream_limit(
        stream(new_prefix.as_bytes(), &stream_exact),
        "namespace bindings",
    );

    let four = Limits {
        max_namespace_bindings: 4,
        ..exact
    };
    codec(new_prefix.as_bytes(), &four).expect("the new prefix fits one larger budget");
    let stream_four = StreamLimits {
        processing: four,
        ..StreamLimits::default()
    };
    stream(new_prefix.as_bytes(), &stream_four)
        .expect("the stream new prefix fits one larger budget");
}

#[test]
fn directive_token_budget_is_cumulative_and_exact() {
    let source =
        format!(r#"<r xmlns:mc="{MC}" xmlns:a="urn:a" xmlns:b="urn:b" mc:Ignorable="a b"/>"#);

    let exact = Limits {
        max_directive_tokens: 2,
        ..Limits::default()
    };
    codec(source.as_bytes(), &exact).expect("two directive tokens fit");
    assert_codec_limit(
        codec(
            source.as_bytes(),
            &Limits {
                max_directive_tokens: 1,
                ..exact.clone()
            },
        ),
        "directive tokens",
    );

    let stream_exact = StreamLimits {
        processing: exact.clone(),
        ..StreamLimits::default()
    };
    stream(source.as_bytes(), &stream_exact).expect("stream accepts two directive tokens");
    assert_stream_limit(
        stream(
            source.as_bytes(),
            &StreamLimits {
                processing: Limits {
                    max_directive_tokens: 1,
                    ..exact
                },
                ..StreamLimits::default()
            },
        ),
        "directive tokens",
    );
}

#[test]
fn codec_and_stream_depth_limits_admit_exact_nesting_only() {
    let source = format!(r#"<r xmlns:mc="{MC}"><a></a></r>"#);

    let exact = Limits {
        max_depth: 2,
        ..Limits::default()
    };
    codec(source.as_bytes(), &exact).expect("two element levels fit");
    assert_codec_limit(
        codec(
            source.as_bytes(),
            &Limits {
                max_depth: 1,
                ..exact.clone()
            },
        ),
        "depth",
    );

    let stream_exact = StreamLimits {
        processing: exact.clone(),
        ..StreamLimits::default()
    };
    stream(source.as_bytes(), &stream_exact).expect("stream admits two element levels");
    assert_stream_limit(
        stream(
            source.as_bytes(),
            &StreamLimits {
                processing: Limits {
                    max_depth: 1,
                    ..exact
                },
                ..StreamLimits::default()
            },
        ),
        "depth",
    );
}

#[test]
fn codec_input_and_output_byte_limits_are_exact() {
    let input_source = format!(r#"<r xmlns:mc="{MC}"/>"#);
    let input_exact = Limits {
        max_input_bytes: input_source.len(),
        ..Limits::default()
    };
    codec(input_source.as_bytes(), &input_exact).expect("input at the exact bound fits");
    assert_codec_limit(
        codec(
            input_source.as_bytes(),
            &Limits {
                max_input_bytes: input_source.len() - 1,
                ..input_exact.clone()
            },
        ),
        "input bytes",
    );

    // The processor expands the empty child to an explicit start/end pair.
    // These lengths are properties of this literal fixture, not observations
    // obtained from the processor being qualified.
    let source = format!(r#"<r xmlns:mc="{MC}"><a/></r>"#);
    let expected = format!(r#"<r xmlns:mc="{MC}"><a></a></r>"#);
    assert_eq!(source.len(), 82);
    assert_eq!(expected.len(), 85);
    assert_eq!(
        codec(source.as_bytes(), &Limits::default()).unwrap(),
        expected.as_bytes()
    );

    codec(
        source.as_bytes(),
        &Limits {
            max_output_bytes: 85,
            ..Limits::default()
        },
    )
    .expect("output at the exact bound fits");
    assert_codec_limit(
        codec(
            source.as_bytes(),
            &Limits {
                max_output_bytes: 84,
                ..Limits::default()
            },
        ),
        "output bytes",
    );
}

#[test]
fn stream_input_bytes_are_distinct_from_event_bytes() {
    let source = b"<r/>";

    let input_exact = StreamLimits {
        processing: Limits {
            max_input_bytes: source.len(),
            ..Limits::default()
        },
        ..StreamLimits::default()
    };
    stream(source, &input_exact).expect("four input bytes fit");
    assert_stream_limit(
        stream(
            source,
            &StreamLimits {
                processing: Limits {
                    max_input_bytes: source.len() - 1,
                    ..Limits::default()
                },
                ..StreamLimits::default()
            },
        ),
        "input bytes",
    );

    let event_exact = StreamLimits {
        max_event_bytes: 4,
        ..StreamLimits::default()
    };
    let report = stream(source, &event_exact).expect("the one four-byte event fits");
    assert_eq!(report.events, 1);
    assert_stream_limit(
        stream(
            source,
            &StreamLimits {
                max_event_bytes: 3,
                ..StreamLimits::default()
            },
        ),
        "stream event bytes",
    );
}

#[test]
fn stream_event_count_and_raw_attribute_bytes_are_exact() {
    let events = b"<r><a/></r>";
    let exact = StreamLimits {
        max_events: 3,
        ..StreamLimits::default()
    };
    let report = stream(events, &exact).expect("three non-EOF events fit");
    assert_eq!(report.events, 3);
    assert_stream_limit(
        stream(
            events,
            &StreamLimits {
                max_events: 2,
                ..exact
            },
        ),
        "stream event count",
    );

    // The raw attribute budget counts `a` plus the five source bytes in
    // `&amp;`; it does not count the one decoded ampersand byte.
    let attribute = br#"<r a="&amp;"/>"#;
    let exact = StreamLimits {
        max_attribute_bytes_per_event: 6,
        ..StreamLimits::default()
    };
    stream(attribute, &exact).expect("raw attribute name and value fit");
    assert_stream_limit(
        stream(
            attribute,
            &StreamLimits {
                max_attribute_bytes_per_event: 5,
                ..exact
            },
        ),
        "stream attribute bytes",
    );
}

#[test]
fn stream_qualified_and_expanded_name_bytes_share_one_exact_bound() {
    let namespace = "urn:qualified-namespace-longer";
    let source = format!(r#"<p:value xmlns:p="{namespace}"/>"#);
    // The expanded element name is the namespace (30) plus local name (5).
    // This is larger than the fixed XMLNS namespace used by the declaration,
    // so the fixture reaches the element's expanded-name check.
    assert_eq!(namespace.len(), 30);
    let exact_name_bytes = 35;
    let exact = StreamLimits {
        max_name_bytes: exact_name_bytes,
        ..StreamLimits::default()
    };
    stream(source.as_bytes(), &exact).expect("qualified and expanded names fit");
    assert_stream_limit(
        stream(
            source.as_bytes(),
            &StreamLimits {
                max_name_bytes: exact_name_bytes - 1,
                ..exact
            },
        ),
        "stream name bytes",
    );
}

#[test]
fn stream_context_bytes_are_exact_across_namespace_and_frame_state() {
    let source = br#"<r xmlns:p="urn:p"><p:e></p:e></r>"#;
    // Semantic context at the deepest open frame is 19 bytes:
    // root declaration (1 + 5 + 2), root expanded/qualified name (1 + 0 + 1),
    // and child expanded/qualified name (3 + 5 + 1).
    let exact = StreamLimits {
        max_context_bytes: 19,
        ..StreamLimits::default()
    };
    stream(source, &exact).expect("the deepest frame fits the exact context bound");
    assert_stream_limit(
        stream(
            source,
            &StreamLimits {
                max_context_bytes: 18,
                ..exact
            },
        ),
        "stream context bytes",
    );
}

// Deliberately outside this qualification module: allocator-failure
// boundaries cannot be forced deterministically through the public API, and
// the streaming API has no output-byte budget because it emits callbacks
// rather than an owned output buffer.  Those cases remain covered by the
// production unit tests and typed-oracle reviews where their available
// evidence is explicit.
