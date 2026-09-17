#![allow(
    clippy::unwrap_used,
    reason = "focused audit tests fail at the unexpected result"
)]

use std::io::{self, BufRead, Read};

use xml_minifier::audit::{self, Error, Kind, Limits, Report, Resource, StreamError};

fn limits_for(input: &[u8]) -> Limits {
    Limits::new(input.len(), 32, 128, 64, input.len().max(1), input.len()).unwrap()
}

fn stream_report<R: BufRead>(reader: R, limits: Limits, authored: bool) -> Result<Report, Error> {
    let result = if authored {
        audit::verify_authored_reader(reader, limits)
    } else {
        audit::verify_reader(reader, limits)
    };
    result.map_err(|error| match error {
        StreamError::Audit(error) => error,
        StreamError::Input(error) => panic!("unexpected input error: {error}"),
        _ => panic!("unexpected non-exhaustive stream error"),
    })
}

fn assert_same_result(input: &[u8], authored: bool) {
    let expected = if authored {
        audit::verify_authored(input, limits_for(input))
    } else {
        audit::verify(input, limits_for(input))
    };

    for chunks in [
        &[1usize][..],
        &[2usize][..],
        &[3usize][..],
        &[7usize][..],
        &[31usize, 1, 5, 2][..],
    ] {
        let actual = if authored {
            audit::verify_authored_reader(Chunked::new(input, chunks), limits_for(input))
        } else {
            audit::verify_reader(Chunked::new(input, chunks), limits_for(input))
        };
        match (&expected, &actual) {
            (Ok(expected), Ok(actual)) => assert_eq!(expected, actual, "chunks {chunks:?}"),
            (Err(expected), Err(StreamError::Audit(actual))) => {
                assert_error_kind_and_offset(expected, actual, chunks);
            },
            (Err(expected), Err(StreamError::Input(actual))) => {
                panic!("slice error {expected:?} became input error {actual}")
            },
            (Err(_), Err(_)) => panic!("slice error became an unknown stream error"),
            (Ok(expected), Err(actual)) => {
                panic!("slice success {expected:?} became stream error {actual:?}")
            },
            (Err(expected), Ok(actual)) => {
                panic!("slice error {expected:?} became stream success {actual:?}")
            },
        }
    }
}

fn assert_authored_with_limits(input: &[u8], limits: Limits) {
    let expected = audit::verify_authored(input, limits);

    for chunks in [
        &[1usize][..],
        &[2usize][..],
        &[3usize][..],
        &[7usize][..],
        &[31usize, 1, 5, 2][..],
    ] {
        let actual = audit::verify_authored_reader(Chunked::new(input, chunks), limits);
        match (&expected, actual) {
            (Ok(expected), Ok(actual)) => assert_eq!(expected, &actual, "chunks {chunks:?}"),
            (Err(expected), Err(StreamError::Audit(actual))) => {
                assert_error_kind_and_offset(expected, &actual, chunks);
            },
            (Err(expected), Err(StreamError::Input(actual))) => {
                panic!("slice error {expected:?} became input error {actual}")
            },
            (Err(_), Err(_)) => panic!("slice error became an unknown stream error"),
            (Ok(expected), Err(actual)) => {
                panic!("slice success {expected:?} became stream error {actual:?}")
            },
            (Err(expected), Ok(actual)) => {
                panic!("slice error {expected:?} became stream success {actual:?}")
            },
        }
    }
}

fn authored_limits(input: &[u8], max_attributes: usize) -> Limits {
    Limits::new(
        input.len(),
        32,
        128,
        max_attributes,
        input.len().max(1),
        input.len(),
    )
    .unwrap()
}

fn assert_same_category(input: &[u8], authored: bool) {
    let expected = if authored {
        audit::verify_authored(input, limits_for(input))
    } else {
        audit::verify(input, limits_for(input))
    };
    for chunks in [&[1usize][..], &[3usize][..]] {
        let actual = if authored {
            audit::verify_authored_reader(Chunked::new(input, chunks), limits_for(input))
        } else {
            audit::verify_reader(Chunked::new(input, chunks), limits_for(input))
        };
        match (&expected, actual) {
            (Ok(expected), Ok(actual)) => assert_eq!(expected, &actual, "chunks {chunks:?}"),
            (Err(expected), Err(StreamError::Audit(actual))) => {
                assert_eq!(
                    error_category(expected),
                    error_category(&actual),
                    "chunks {chunks:?}"
                );
            },
            (Err(expected), Err(StreamError::Input(actual))) => {
                panic!("slice error {expected:?} became input error {actual}")
            },
            (Err(_), Err(_)) => panic!("slice error became an unknown stream error"),
            (Ok(expected), Err(actual)) => {
                panic!("slice success {expected:?} became stream error {actual:?}")
            },
            (Err(expected), Ok(actual)) => {
                panic!("slice error {expected:?} became stream success {actual:?}")
            },
        }
    }
}

fn error_category(error: &Error) -> &'static str {
    match error {
        Error::Limit { .. } => "limit",
        Error::Encoding { .. } => "encoding",
        Error::NotCompact(_) => "not-compact",
        Error::Doctype { .. } => "doctype",
        Error::Malformed { .. } => "malformed",
        Error::Allocation => "allocation",
        _ => "unknown",
    }
}

fn assert_error_kind_and_offset(expected: &Error, actual: &Error, chunks: &[usize]) {
    match (expected, actual) {
        (
            Error::Limit {
                resource: expected_resource,
                limit: expected_limit,
                actual: expected_actual,
                offset: expected_offset,
            },
            Error::Limit {
                resource: actual_resource,
                limit: actual_limit,
                actual: actual_actual,
                offset: actual_offset,
            },
        ) => assert_eq!(
            (
                expected_resource,
                expected_limit,
                expected_actual,
                expected_offset,
            ),
            (actual_resource, actual_limit, actual_actual, actual_offset),
            "chunks {chunks:?}"
        ),
        (
            Error::Encoding {
                valid_up_to: expected,
            },
            Error::Encoding {
                valid_up_to: actual,
            },
        ) => {
            assert_eq!(expected, actual, "chunks {chunks:?}");
        },
        (Error::NotCompact(expected), Error::NotCompact(actual)) => {
            assert_eq!(expected, actual, "chunks {chunks:?}")
        },
        (Error::Doctype { offset: expected }, Error::Doctype { offset: actual }) => {
            assert_eq!(expected, actual, "chunks {chunks:?}");
        },
        (
            Error::Malformed {
                offset: expected, ..
            },
            Error::Malformed { offset: actual, .. },
        ) => assert_eq!(expected, actual, "chunks {chunks:?}"),
        (Error::Allocation, Error::Allocation) => {},
        (expected, actual) => panic!(
            "error category differs for chunks {chunks:?}: expected {expected:?}, actual {actual:?}"
        ),
    }
}

#[test]
fn successful_reports_match_slice_for_arbitrary_tiny_chunks() {
    let xml = br#"<?xml version="1.0"?><root a='a&quot;b' b="2"><child xml:space="preserve">
</child><!--comment--><?keep x?><![CDATA[  ]]>&amp;&#32;&#x20;</root>"#;
    assert_same_result(xml, false);
    assert_same_result(xml, true);

    let mixed = b"<p><b>a</b> <i>b</i></p>";
    assert_same_result(mixed, false);
    assert_same_result(b"<p>boxed &lt;text&gt; &amp; more</p>", true);
    assert_same_result(b"<a>&unknown;&#xZZ;&#;</a>", false);
}

#[test]
fn raw_layout_cases_match_slice_for_arbitrary_tiny_chunks() {
    for xml in [
        b"<a x=\"1\"/>".as_slice(),
        b"<a  x=\"1\"/>",
        b"<a x =\"1\"/>",
        b"<a x= \"1\"/>",
        b"<a x=\"1\"  y=\"2\"/>",
        b"<a\nx=\"1\"/>",
        b"<a />",
        b"<a  />",
        b"<a x=\">\"/>",
        b"<a x=\"/>\"/>",
        b"<a >x</a>",
        b"<a></a >",
        b"<a></a\t>",
        b"<a><!--x--y--></a>",
        b"<?xml version=\"1.0\"?><a/>",
        b"<?xml  version=\"1.0\"?><a/>",
        b"<?xml version =\"1.0\"?><a/>",
        b"<?xml version= \"1.0\"?><a/>",
        b"<?xml version=\"1.0\" ?><a/>",
    ] {
        assert_same_result(xml, false);
    }
}

#[test]
fn deterministic_ascii_mutations_keep_reader_transport_parity() {
    let seeds: &[&[u8]] = &[
        b"<a/>",
        b"<a>text</a>",
        b"<a x=\"1\"/>",
        b"<a><b>text&amp;</b></a>",
        b"<?xml version=\"1.0\"?><a/>",
        b"<a><!--x--><?p?><b/></a>",
        b"<a><![CDATA[x]]>&amp;</a>",
    ];
    let replacements = b"<>/?!=-&;_'\" \t\r\nabc012";
    let mut random = 0x0482_5eed_u64;
    for case_index in 0..1_000 {
        let seed = seeds[next_random(&mut random) as usize % seeds.len()];
        let mut candidate = seed.to_vec();
        let mutations = 1 + (next_random(&mut random) as usize % 3);
        for _ in 0..mutations {
            let position = next_random(&mut random) as usize % candidate.len();
            let replacement = replacements[next_random(&mut random) as usize % replacements.len()];
            candidate[position] = replacement;
        }
        let result = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
            assert_same_category(&candidate, false);
            let mut marked = b"\xef\xbb\xbf".to_vec();
            marked.extend_from_slice(&candidate);
            assert_same_category(&marked, false);
            if case_index % 7 == 0 {
                assert_same_category(&candidate, true);
            }
        }));
        if result.is_err() {
            panic!("ASCII differential failed for case {case_index}: {candidate:?}");
        }
    }
}

#[test]
fn bom_raw_event_boundaries_match_slice_for_all_prefix_chunkings() {
    // Both auditors address physical input bytes while quick-xml excludes the
    // initial BOM from its event positions. Exactly one leading BOM is framing.
    for xml in [
        b"\xef\xbb\xbf<a/>".as_slice(),
        b"\xef\xbb\xbf<?xml version=\"1.0\"?><a/>",
        b"\xef\xbb\xbf<!--c--><a/>",
        b"\xef\xbb\xbf<?p?><a/>",
        b"\xef\xbb\xbfx<a/>",
        b"\xef\xbb\xbf <a/>",
        b"\xef\xbb\xbf<![CDATA[x]]><a/>",
        b"\xef\xbb\xbf<!DOCTYPE a><a/>",
        b"\xef\xbb\xbf\xef\xbb\xbf<a/>",
        b"\xef\xbb\xbf<a>\xff</a>",
        b"\xef\xbb\xbf\xff<a/>",
        b"\xef\xbb\xbf",
    ] {
        assert_same_result(xml, false);
    }
}

#[test]
fn bom_token_limit_offsets_match_slice_for_all_prefix_chunkings() {
    let xml = b"\xef\xbb\xbf<a/>";
    let limits = Limits::new(xml.len(), 32, 128, 64, 3, xml.len()).unwrap();
    let expected = audit::verify(xml, limits).unwrap_err();
    assert!(matches!(
        expected,
        Error::Limit {
            resource: Resource::TokenBytes,
            limit: 3,
            actual: 4,
            offset: 3,
        }
    ));
    for chunks in [
        &[1usize][..],
        &[2usize][..],
        &[3usize][..],
        &[1usize, 2, 1][..],
        &[2usize, 1, 3][..],
    ] {
        let actual = stream_report(Chunked::new(xml, chunks), limits, false).unwrap_err();
        assert_error_kind_and_offset(&expected, &actual, chunks);
    }
}

fn next_random(state: &mut u64) -> u64 {
    *state ^= state.wrapping_shl(7);
    *state ^= state.wrapping_shr(9);
    *state
}

#[test]
fn authored_space_and_entity_state_survive_every_chunk_boundary() {
    let explicit = b"<a xml:space=\"pre&#115;erve\">\n<b> </b></a>";
    assert_same_result(explicit, false);
    assert_same_result(explicit, true);

    let ambiguous = b"<p><b>a</b> <i>b</i></p>";
    let expected = audit::verify_authored(ambiguous, limits_for(ambiguous)).unwrap_err();
    assert!(
        matches!(expected, Error::NotCompact(violation) if violation.kind() == Kind::AmbiguousWhitespace)
    );
    for chunks in [&[1usize][..], &[2usize][..], &[3usize][..], &[7usize][..]] {
        let actual =
            audit::verify_authored_reader(Chunked::new(ambiguous, chunks), limits_for(ambiguous))
                .unwrap_err();
        let actual = match actual {
            StreamError::Audit(error) => error,
            StreamError::Input(error) => panic!("unexpected input error: {error}"),
            _ => panic!("unexpected non-exhaustive stream error"),
        };
        assert_error_kind_and_offset(&expected, &actual, chunks);
    }
}

#[test]
fn authored_attribute_probe_covers_zero_one_many_and_namespace_attributes() {
    let zero = b"<root><child/></root>";
    let one = br#"<root a="1"/>"#;
    let many = br#"<root a="1" b="2" c="3"/>"#;
    let namespaced = br#"<p:root xmlns:p="urn:test" p:value="v"/>"#;

    for (xml, attributes) in [
        (zero.as_slice(), 0),
        (one.as_slice(), 1),
        (many.as_slice(), 3),
        (namespaced.as_slice(), 2),
    ] {
        let report =
            audit::verify_authored(xml, limits_for(xml)).expect("valid authored attribute fixture");
        assert_eq!(report.attributes(), attributes, "fixture {xml:?}");
        assert_same_result(xml, true);
    }
}

#[test]
fn authored_attribute_probe_preserves_xml_space_state_and_limit_precedence() {
    let preserved = b"<root xml:space=\"preserve\">\n<child/></root>";
    let report = audit::verify_authored(preserved, limits_for(preserved))
        .expect("xml:space=preserve must carry into the child text run");
    assert_eq!(report.attributes(), 1);
    assert_same_result(preserved, true);

    let many_preserved = b"<root a=\"1\" xml:space=\"preserve\" b=\"2\">\n</root>";
    let report = audit::verify_authored(many_preserved, limits_for(many_preserved))
        .expect("checked multi-attribute path must carry xml:space state");
    assert_eq!(report.attributes(), 3);
    assert_same_result(many_preserved, true);

    let reset = b"<root xml:space=\"preserve\"><child xml:space=\"default\">\n</child></root>";
    let error = audit::verify_authored(reset, limits_for(reset)).unwrap_err();
    assert!(matches!(
        error,
        Error::NotCompact(violation) if violation.kind() == Kind::FormattingWhitespace
    ));
    assert_same_result(reset, true);

    let normalized = br#"<root xml:space="pre&#115;erve">
</root>"#;
    let report = audit::verify_authored(normalized, limits_for(normalized))
        .expect("normalized xml:space value must preserve the text run");
    assert_eq!(report.attributes(), 1);
    assert_same_result(normalized, true);

    let invalid = br#"<root xml:space="keep"/>"#;
    let error = audit::verify_authored(invalid, limits_for(invalid)).unwrap_err();
    assert!(matches!(error, Error::Malformed { ref detail, .. } if detail.contains("xml:space")));
    assert_same_result(invalid, true);

    let limited = b"<root xml:space=\"default\">\n</root>";
    let limits = authored_limits(limited, 0);
    let error = audit::verify_authored(limited, limits).unwrap_err();
    assert!(matches!(
        error,
        Error::Limit {
            resource: Resource::Attributes,
            limit: 0,
            actual: 1,
            ..
        }
    ));
    assert_authored_with_limits(limited, limits);
}

#[test]
fn authored_attribute_probe_keeps_duplicate_and_value_error_order() {
    let cases = [
        (
            br#"<root xml:space="preserve" xml:space="keep"/>"#.as_slice(),
            "duplicated attribute",
        ),
        (
            br#"<root xml:space="keep"/>"#.as_slice(),
            "xml:space must be 'default' or 'preserve'",
        ),
        (
            br#"<root a="ok" xml:space="keep"/>"#.as_slice(),
            "xml:space must be 'default' or 'preserve'",
        ),
    ];

    for (xml, detail) in cases {
        let error = audit::verify_authored(xml, limits_for(xml)).unwrap_err();
        assert!(
            matches!(error, Error::Malformed { detail: ref actual, .. } if actual.contains(detail)),
            "unexpected malformed ordering for {xml:?}: {error:?}"
        );
        assert_authored_with_limits(xml, limits_for(xml));
    }
}

#[test]
fn authored_attribute_limits_remain_aggregate_and_inclusive() {
    let same_tag = br#"<root a="1" b="2"/>"#;
    let accepted = audit::verify_authored(same_tag, authored_limits(same_tag, 2))
        .expect("the inclusive attribute limit must accept its boundary");
    assert_eq!(accepted.attributes(), 2);
    assert_authored_with_limits(same_tag, authored_limits(same_tag, 2));

    let error = audit::verify_authored(same_tag, authored_limits(same_tag, 1)).unwrap_err();
    assert!(matches!(
        error,
        Error::Limit {
            resource: Resource::Attributes,
            limit: 1,
            actual: 2,
            ..
        }
    ));
    assert_authored_with_limits(same_tag, authored_limits(same_tag, 1));

    let one = br#"<root a="1"/>"#;
    let error = audit::verify_authored(one, authored_limits(one, 0)).unwrap_err();
    assert!(matches!(
        error,
        Error::Limit {
            resource: Resource::Attributes,
            limit: 0,
            actual: 1,
            ..
        }
    ));
    assert_authored_with_limits(one, authored_limits(one, 0));

    let aggregate = br#"<root a="1"><child b="2"/><leaf c="3"/></root>"#;
    let report = audit::verify_authored(aggregate, authored_limits(aggregate, 3))
        .expect("aggregate attributes across one-attribute tags must be counted");
    assert_eq!(report.attributes(), 3);
    assert_authored_with_limits(aggregate, authored_limits(aggregate, 3));

    let error = audit::verify_authored(aggregate, authored_limits(aggregate, 2)).unwrap_err();
    assert!(matches!(
        error,
        Error::Limit {
            resource: Resource::Attributes,
            limit: 2,
            actual: 3,
            ..
        }
    ));
    assert_authored_with_limits(aggregate, authored_limits(aggregate, 2));
}

#[test]
fn accumulated_text_and_entity_bytes_match() {
    let xml = b"<a>abc&amp;def<![CDATA[ghi]]></a>";
    let expected = audit::verify(xml, limits_for(xml)).unwrap();
    assert_eq!(expected.text_bytes(), 14);
    assert_eq!(expected.events(), 7);
    for chunks in [
        &[1usize][..],
        &[2usize][..],
        &[3usize][..],
        &[5usize, 1, 4][..],
    ] {
        let actual = stream_report(Chunked::new(xml, chunks), limits_for(xml), false).unwrap();
        assert_eq!(actual, expected, "chunks {chunks:?}");
    }
}

#[test]
fn exact_resource_boundaries_match_and_one_under_fails() {
    let xml = b"<a x=\"1\"><b>text&amp;</b></a>";
    let exact = audit::verify(xml, limits_for(xml)).unwrap();
    for chunks in [
        &[1usize][..],
        &[2usize][..],
        &[7usize][..],
        &[31usize, 1, 3][..],
    ] {
        assert_eq!(
            stream_report(Chunked::new(xml, chunks), limits_for(xml), false).unwrap(),
            exact,
            "chunks {chunks:?}"
        );
    }

    let cases = [
        (
            Resource::Bytes,
            Limits::new(xml.len() - 1, 32, 128, 64, xml.len(), xml.len()).unwrap(),
        ),
        (
            Resource::Events,
            Limits::new(xml.len(), 32, exact.events() - 1, 64, xml.len(), xml.len()).unwrap(),
        ),
        (
            Resource::Depth,
            Limits::new(
                xml.len(),
                exact.max_depth() - 1,
                128,
                64,
                xml.len(),
                xml.len(),
            )
            .unwrap(),
        ),
        (
            Resource::Attributes,
            Limits::new(
                xml.len(),
                32,
                128,
                exact.attributes() - 1,
                xml.len(),
                xml.len(),
            )
            .unwrap(),
        ),
        (
            Resource::TextBytes,
            Limits::new(xml.len(), 32, 128, 64, xml.len(), exact.text_bytes() - 1).unwrap(),
        ),
    ];
    for (resource, limits) in cases {
        let expected = audit::verify(xml, limits).unwrap_err();
        assert!(matches!(expected, Error::Limit { resource: actual, .. } if actual == resource));
        let actual = stream_report(Chunked::new(xml, &[1]), limits, false).unwrap_err();
        assert!(matches!(actual, Error::Limit { resource: found, .. } if found == resource));
    }

    let token_xml = b"<a>text</a>";
    let exact_token = Limits::new(token_xml.len(), 32, 128, 64, 4, token_xml.len()).unwrap();
    assert!(audit::verify(token_xml, exact_token).is_ok());
    assert!(stream_report(Chunked::new(token_xml, &[1]), exact_token, false).is_ok());
    let short_token = Limits::new(token_xml.len(), 32, 128, 64, 3, token_xml.len()).unwrap();
    let error = stream_report(Chunked::new(token_xml, &[1]), short_token, false).unwrap_err();
    assert!(matches!(
        error,
        Error::Limit {
            resource: Resource::TokenBytes,
            actual: 4,
            offset: 3,
            ..
        }
    ));
}

#[test]
fn depth_and_event_boundaries_match_slice() {
    let nested = b"<a><b><c/></b></a>";
    let exact_depth = Limits::new(nested.len(), 3, 64, 64, nested.len(), nested.len()).unwrap();
    assert_same_result(nested, false);
    assert_eq!(
        stream_report(Chunked::new(nested, &[1]), exact_depth, false).unwrap(),
        audit::verify(nested, exact_depth).unwrap()
    );

    let shallow = Limits::new(nested.len(), 2, 64, 64, nested.len(), nested.len()).unwrap();
    let expected = audit::verify(nested, shallow).unwrap_err();
    assert!(matches!(
        expected,
        Error::Limit {
            resource: Resource::Depth,
            limit: 2,
            actual: 3,
            offset: 6,
        }
    ));
    for chunks in [&[1usize][..], &[2usize][..], &[3usize][..], &[7usize][..]] {
        let actual = stream_report(Chunked::new(nested, chunks), shallow, false).unwrap_err();
        assert_error_kind_and_offset(&expected, &actual, chunks);
    }

    let events = b"<a/>";
    let exact_events = Limits::new(events.len(), 8, 2, 8, events.len(), events.len()).unwrap();
    assert_eq!(
        stream_report(Chunked::new(events, &[1]), exact_events, false).unwrap(),
        audit::verify(events, exact_events).unwrap()
    );
    let one_short = Limits::new(events.len(), 8, 1, 8, events.len(), events.len()).unwrap();
    let expected = audit::verify(events, one_short).unwrap_err();
    let actual = stream_report(Chunked::new(events, &[1]), one_short, false).unwrap_err();
    assert_error_kind_and_offset(&expected, &actual, &[1]);
}

#[test]
fn resource_limit_offsets_are_absolute_and_stream_ordered() {
    let bytes_xml = b"<a><b>text</b></a>";
    let bytes_limits = Limits::new(
        bytes_xml.len() - 1,
        32,
        128,
        64,
        bytes_xml.len(),
        bytes_xml.len(),
    )
    .unwrap();
    let expected = audit::verify(bytes_xml, bytes_limits).unwrap_err();
    assert!(matches!(
        expected,
        Error::Limit {
            resource: Resource::Bytes,
            offset: 0,
            ..
        }
    ));
    for chunks in [&[1usize][..], &[2usize][..], &[3usize][..], &[7usize][..]] {
        let actual =
            stream_report(Chunked::new(bytes_xml, chunks), bytes_limits, false).unwrap_err();
        let offset = match actual {
            Error::Limit {
                resource: Resource::Bytes,
                offset,
                ..
            } => offset,
            other => panic!("wrong byte-limit error for chunks {chunks:?}: {other:?}"),
        };
        // The slice API rejects the complete input before parsing and reports
        // offset zero. The reader reports the first source position at which
        // the byte window is observed, the closing </a> event at byte 14.
        assert_eq!(offset, 14, "chunks {chunks:?}");
    }

    let token_xml = b"<a>0123456789</a>";
    let token_limits = Limits::new(token_xml.len(), 32, 128, 64, 8, token_xml.len()).unwrap();
    let expected = audit::verify(token_xml, token_limits).unwrap_err();
    let expected_offset = match expected {
        Error::Limit {
            resource: Resource::TokenBytes,
            offset,
            ..
        } => offset,
        other => panic!("wrong token-limit slice error: {other:?}"),
    };
    for chunks in [&[1usize][..], &[2usize][..], &[3usize][..], &[7usize][..]] {
        let actual =
            stream_report(Chunked::new(token_xml, chunks), token_limits, false).unwrap_err();
        let offset = match actual {
            Error::Limit {
                resource: Resource::TokenBytes,
                offset,
                ..
            } => offset,
            other => panic!("wrong token-limit error for chunks {chunks:?}: {other:?}"),
        };
        assert_eq!(offset, expected_offset, "chunks {chunks:?}");
    }
}

#[test]
fn oversized_token_is_rejected_at_bounded_lookahead() {
    let xml = b"<root>0123456789</root>";
    let limits = Limits::new(xml.len(), 32, 128, 64, 8, xml.len()).unwrap();
    for chunks in [
        &[1usize][..],
        &[2usize][..],
        &[3usize][..],
        &[7usize][..],
        &[31usize, 1, 2][..],
    ] {
        let error = stream_report(Chunked::new(xml, chunks), limits, false).unwrap_err();
        assert!(
            matches!(
                error,
                Error::Limit {
                    resource: Resource::TokenBytes,
                    limit: 8,
                    actual: 9,
                    offset: 6,
                }
            ),
            "chunks {chunks:?}: {error:?}"
        );
    }
}

#[test]
fn truncated_token_kinds_are_rejected_before_unbounded_growth() {
    let inputs = [
        b"<root>0123456789".as_slice(),
        b"<root attr=\"0123456789".as_slice(),
        b"<root><!--0123456789".as_slice(),
        b"<root><![CDATA[0123456789".as_slice(),
        b"<root><?p 0123456789".as_slice(),
        b"<root>&0123456789".as_slice(),
        b"<?xml version=\"0123456789".as_slice(),
    ];
    for input in inputs {
        let limits = Limits::new(input.len(), 32, 128, 64, 8, input.len()).unwrap();
        for chunks in [&[1usize][..], &[2usize][..], &[3usize][..], &[7usize][..]] {
            let error = stream_report(Chunked::new(input, chunks), limits, false).unwrap_err();
            assert!(
                matches!(
                    error,
                    Error::Limit {
                        resource: Resource::TokenBytes,
                        limit: 8,
                        actual: 9,
                        ..
                    }
                ),
                "chunks {chunks:?}, input {input:?}: {error:?}"
            );
        }
    }
}

#[test]
fn invalid_utf8_offsets_match_slice_for_each_event_kind() {
    for xml in [
        b"<a>\xff</a>".as_slice(),
        b"<a x=\"\xff\"/>",
        b"<\xff/>",
        b"<a><!--\xff--></a>",
        b"<a><![CDATA[\xff]]></a>",
        b"<?xml \xff?><a/>",
        b"<a><?p \xff?></a>",
        b"<a>&\xff;</a>",
        b"\xff<a/>",
    ] {
        assert_same_result(xml, false);
    }
}

#[test]
fn malformed_inputs_keep_typed_category() {
    for xml in [
        b"<a>".as_slice(),
        b"<a",
        b"<a x=\"1",
        b"<a></b>",
        b"<a/><b/>",
        b"x<a/>",
        b"<a>&</a>",
        b"<a>&name",
        b"<a><!--",
        b"<a><![CDATA[x</a>",
        b"<a><![CDATA[",
        b"<a><!--x</a>",
        b"<a><?p",
        b"<?xml",
        b"<a x=y/>",
        b"<a x=\"1\" x=\"2\"/>",
        b"<a xml:space=\"preserve\" xml:space=\"default\">x</a>",
        b"<a>\xff</a>",
    ] {
        let expected = audit::verify(xml, limits_for(xml)).unwrap_err();
        for chunks in [&[1usize][..], &[2usize][..], &[3usize][..], &[7usize][..]] {
            let actual = audit::verify_reader(Chunked::new(xml, chunks), limits_for(xml));
            let actual = match actual {
                Err(StreamError::Audit(error)) => error,
                Err(StreamError::Input(error)) => panic!("unexpected input error: {error}"),
                Err(_) => panic!("unexpected non-exhaustive stream error"),
                Ok(report) => panic!("expected {expected:?}, got {report:?}"),
            };
            match (&expected, &actual) {
                (Error::Encoding { .. }, Error::Encoding { .. })
                | (Error::Malformed { .. }, Error::Malformed { .. }) => {},
                (expected, actual) => {
                    panic!(
                        "category differs for {xml:?}, chunks {chunks:?}: {expected:?} vs {actual:?}"
                    )
                },
            }
        }
    }
}

#[test]
fn interrupted_input_retries_and_permanent_input_error_is_typed() {
    let xml = b"<a><b>text</b></a>";
    let limits = limits_for(xml);
    let interrupted_during_prefix =
        audit::verify_reader(Faulty::interrupted_once(xml, 0, &[1, 2, 1]), limits).unwrap();
    assert_eq!(
        interrupted_during_prefix,
        audit::verify(xml, limits).unwrap()
    );

    let interrupted =
        audit::verify_reader(Faulty::interrupted_once(xml, 3, &[1, 2, 1]), limits).unwrap();
    assert_eq!(interrupted, audit::verify(xml, limits).unwrap());

    let error = audit::verify_reader(Faulty::fail_after(xml, 3, &[1, 2, 1]), limits).unwrap_err();
    match error {
        StreamError::Input(error) => assert_eq!(error.kind(), io::ErrorKind::Other),
        StreamError::Audit(error) => panic!("I/O failure became audit error: {error:?}"),
        _ => panic!("unexpected non-exhaustive stream error"),
    }

    let error = audit::verify_reader(Faulty::fail_after(xml, 7, &[1, 2, 1]), limits).unwrap_err();
    assert!(matches!(error, StreamError::Input(error) if error.kind() == io::ErrorKind::Other));

    let error =
        audit::verify_reader(Faulty::fail_after(xml, xml.len(), &[1, 2, 1]), limits).unwrap_err();
    assert!(matches!(error, StreamError::Input(error) if error.kind() == io::ErrorKind::Other));
}

#[test]
fn streaming_memory_bound_is_finite_and_monotonic() {
    let small = Limits::new(1024, 1, 16, 4, 8, 16).unwrap();
    let large = Limits::new(1024, 4, 16, 4, 32, 16).unwrap();
    assert!(small.streaming_memory_upper_bound().is_some());
    assert!(
        large.streaming_memory_upper_bound().unwrap()
            > small.streaming_memory_upper_bound().unwrap()
    );
}

struct Chunked<'a> {
    input: &'a [u8],
    offset: usize,
    chunks: &'a [usize],
    chunk_index: usize,
    exposed_end: Option<usize>,
}

impl<'a> Chunked<'a> {
    fn new(input: &'a [u8], chunks: &'a [usize]) -> Self {
        assert!(!chunks.is_empty());
        assert!(chunks.iter().all(|size| *size != 0));
        Self {
            input,
            offset: 0,
            chunks,
            chunk_index: 0,
            exposed_end: None,
        }
    }
}

impl Read for Chunked<'_> {
    fn read(&mut self, output: &mut [u8]) -> io::Result<usize> {
        if output.is_empty() || self.offset == self.input.len() {
            return Ok(0);
        }
        let size = self.chunks[self.chunk_index % self.chunks.len()];
        let amount = output
            .len()
            .min(size)
            .min(self.input.len().saturating_sub(self.offset));
        output[..amount].copy_from_slice(&self.input[self.offset..self.offset + amount]);
        self.offset += amount;
        Ok(amount)
    }
}

impl BufRead for Chunked<'_> {
    fn fill_buf(&mut self) -> io::Result<&[u8]> {
        if self.offset == self.input.len() {
            return Ok(&[]);
        }
        let end = *self.exposed_end.get_or_insert_with(|| {
            let size = self.chunks[self.chunk_index % self.chunks.len()];
            (self.offset + size).min(self.input.len())
        });
        Ok(&self.input[self.offset..end])
    }

    fn consume(&mut self, amount: usize) {
        let end = self.exposed_end.unwrap_or(self.offset);
        assert!(amount <= end.saturating_sub(self.offset));
        self.offset += amount;
        if self.offset == end {
            self.exposed_end = None;
            self.chunk_index += 1;
        }
    }
}

struct Faulty<'a> {
    inner: Chunked<'a>,
    fail_at: usize,
    interrupt_at: Option<usize>,
    interrupted: bool,
}

impl<'a> Faulty<'a> {
    fn interrupted_once(input: &'a [u8], interrupt_at: usize, chunks: &'a [usize]) -> Self {
        Self {
            inner: Chunked::new(input, chunks),
            fail_at: usize::MAX,
            interrupt_at: Some(interrupt_at),
            interrupted: false,
        }
    }

    fn fail_after(input: &'a [u8], fail_at: usize, chunks: &'a [usize]) -> Self {
        Self {
            inner: Chunked::new(input, chunks),
            fail_at,
            interrupt_at: None,
            interrupted: false,
        }
    }

    fn should_fail(&mut self) -> Option<io::Error> {
        if !self.interrupted
            && self
                .interrupt_at
                .is_some_and(|offset| self.inner.offset >= offset)
        {
            self.interrupted = true;
            return Some(io::Error::new(io::ErrorKind::Interrupted, "retry"));
        }
        if self.inner.offset >= self.fail_at {
            return Some(io::Error::other("injected"));
        }
        None
    }
}

impl Read for Faulty<'_> {
    fn read(&mut self, output: &mut [u8]) -> io::Result<usize> {
        if let Some(error) = self.should_fail() {
            return Err(error);
        }
        self.inner.read(output)
    }
}

impl BufRead for Faulty<'_> {
    fn fill_buf(&mut self) -> io::Result<&[u8]> {
        if let Some(error) = self.should_fail() {
            return Err(error);
        }
        self.inner.fill_buf()
    }

    fn consume(&mut self, amount: usize) {
        self.inner.consume(amount);
    }
}
