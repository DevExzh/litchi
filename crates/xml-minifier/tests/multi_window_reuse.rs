//! Differential and boundary cases for the multi-window source audit.
//!
//! These cases deliberately exercise more than one replacement window in one
//! call.  The complete source/replacement audits remain the oracle: a window
//! parser may prove success, but it must never publish a different refusal.
#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "the fixed fixtures state the expected audit outcome"
)]

use core::ops::Range;

use xml_minifier::audit::{
    Error, Limits, ReplacementError, ReplacementProof, ReplacementSpan, Resource, verify_source,
    verify_source_replacement_with_spans,
};

fn oracle(original: &[u8], replacement: &[u8], limits: Limits) -> Result<(), ReplacementError> {
    let _ = verify_source(original, limits).map_err(ReplacementError::Original)?;
    let _ = verify_source(replacement, limits).map_err(ReplacementError::Replacement)?;
    Ok(())
}

fn checked(
    original: &[u8],
    replacement: &[u8],
    limits: Limits,
    spans: &[ReplacementSpan],
) -> Result<ReplacementProof, ReplacementError> {
    let expected = oracle(original, replacement, limits);
    let actual = verify_source_replacement_with_spans(original, replacement, limits, spans);
    match (&actual, &expected) {
        (Ok(_), Ok(())) => {},
        (Err(actual), Err(expected)) => assert_eq!(actual, expected),
        _ => panic!(
            "multi-window result differs from complete-audit oracle\nactual: {actual:?}\nexpected: {expected:?}\noriginal: {}\nreplacement: {}",
            String::from_utf8_lossy(original),
            String::from_utf8_lossy(replacement),
        ),
    }
    actual
}

fn range(document: &str, needle: &str) -> Range<usize> {
    let start = document.find(needle).expect("fixture needle");
    start..start + needle.len()
}

fn span(
    original: &str,
    replacement: &str,
    original_needle: &str,
    replacement_needle: &str,
) -> ReplacementSpan {
    ReplacementSpan {
        original: range(original, original_needle),
        replacement: range(replacement, replacement_needle),
    }
}

fn assert_windows(original: &str, replacement: &str, limits: Limits, spans: &[ReplacementSpan]) {
    let proof = checked(original.as_bytes(), replacement.as_bytes(), limits, spans)
        .expect("the complete replacement audit accepts the fixture");
    assert!(
        matches!(proof, ReplacementProof::Windows { count, .. } if count == spans.len()),
        "expected one proof for every window, got {proof:?}"
    );
}

#[test]
fn empty_element_first_window_then_nested_namespace_window() {
    let original = r#"<r xmlns:p="urn:p"><a/><nest><b p:k="0"/></nest><pad>padding padding padding padding</pad></r>"#;
    let replacement = r#"<r xmlns:p="urn:p"><x/><nest><b p:k="123"><p:leaf/></b></nest><pad>padding padding padding padding</pad></r>"#;
    let spans = [
        span(original, replacement, "<a/>", "<x/>"),
        span(
            original,
            replacement,
            r#"<b p:k="0"/>"#,
            r#"<b p:k="123"><p:leaf/></b>"#,
        ),
    ];
    assert_windows(original, replacement, Limits::default(), &spans);
}

#[test]
fn varying_depths_and_namespace_rebindings_survive_three_windows() {
    let original = r#"<r xmlns="urn:r" xmlns:p="urn:p"><a/><g xmlns:p="urn:q"><b p:k="0"/><c/></g><z/><pad>padding padding padding padding padding</pad></r>"#;
    let replacement = r#"<r xmlns="urn:r" xmlns:p="urn:p"><a xmlns:x="urn:x" x:k="1"/><g xmlns:p="urn:q"><b p:k="long"><p:q/></b><c/></g><z><p:q/></z><pad>padding padding padding padding padding</pad></r>"#;
    let spans = [
        span(
            original,
            replacement,
            "<a/>",
            r#"<a xmlns:x="urn:x" x:k="1"/>"#,
        ),
        span(
            original,
            replacement,
            r#"<b p:k="0"/>"#,
            r#"<b p:k="long"><p:q/></b>"#,
        ),
        span(original, replacement, "<z/>", "<z><p:q/></z>"),
    ];
    assert_windows(original, replacement, Limits::default(), &spans);
}

#[test]
fn final_window_can_end_with_comment_and_processing_instruction() {
    let original = "<r><a>one</a><b>two</b><pad>padding padding padding padding padding</pad></r>";
    let replacement = "<r><a>ONE</a><b>two</b><!--tail--><?keep value?><pad>padding padding padding padding padding</pad></r>";
    let spans = [
        span(original, replacement, "<a>one</a>", "<a>ONE</a>"),
        span(
            original,
            replacement,
            "<b>two</b>",
            "<b>two</b><!--tail--><?keep value?>",
        ),
    ];
    assert_windows(original, replacement, Limits::default(), &spans);
}

#[test]
fn malformed_first_middle_and_last_windows_use_canonical_errors() {
    let original = "<r><a>0</a><b>0</b><c>0</c><pad>padding padding padding padding</pad></r>";
    let cases = [
        (
            "<a>&</a>",
            "<b>1</b>",
            "<c>2</c>",
            "first dangling reference",
        ),
        ("<a>1</a>", "<b></a>", "<c>2</c>", "middle mismatched end"),
        ("<a>1</a>", "<b>2</b>", "<c><x", "last truncated markup"),
    ];

    for (a, b, c, label) in cases {
        let replacement = format!("<r>{a}{b}{c}<pad>padding padding padding padding</pad></r>");
        let spans = [
            span(original, &replacement, "<a>0</a>", a),
            span(original, &replacement, "<b>0</b>", b),
            span(original, &replacement, "<c>0</c>", c),
        ];
        let actual = checked(
            original.as_bytes(),
            replacement.as_bytes(),
            Limits::default(),
            &spans,
        )
        .expect_err(label);
        assert!(
            matches!(actual, ReplacementError::Replacement(_)),
            "{label}: {actual:?}"
        );
    }

    // This window ends in character data containing the forbidden literal
    // `]]>`.  Its local parser must not expose a relative error; the complete
    // replacement audit owns the canonical offset.
    let replacement = "<r><a/>x]]><b>1</b><c>2</c><pad>padding padding padding padding</pad></r>";
    let spans = [
        span(original, replacement, "<a>0</a>", "<a/>x]]>"),
        span(original, replacement, "<b>0</b>", "<b>1</b>"),
        span(original, replacement, "<c>0</c>", "<c>2</c>"),
    ];
    let actual = checked(
        original.as_bytes(),
        replacement.as_bytes(),
        Limits::default(),
        &spans,
    )
    .expect_err("literal ]]> must use the complete replacement error");
    assert!(matches!(actual, ReplacementError::Replacement(_)));
}

#[test]
fn later_original_error_wins_after_provisional_replacement_failure() {
    let original = r#"<r><a>0</a><b>0</b><tail attr="unterminated></r>"#;
    let replacement = r#"<r><a>1</a><b>&</b></r>"#;
    let spans = [
        span(original, replacement, "<a>0</a>", "<a>1</a>"),
        span(original, replacement, "<b>0</b>", "<b>&</b>"),
    ];
    let actual = checked(
        original.as_bytes(),
        replacement.as_bytes(),
        Limits::default(),
        &spans,
    )
    .expect_err("the malformed original must win");
    assert!(
        matches!(actual, ReplacementError::Original(_)),
        "{actual:?}"
    );
}

#[test]
fn empty_element_hint_cannot_extend_past_its_exact_end() {
    let original = "<r><g><a/></g><b/><pad>padding padding padding padding</pad></r>";
    // The replacement deliberately omits `</g>`.  The first hint starts at
    // the empty `<a/>` but extends through that enclosing close; accepting it
    // as a window would hide the malformed replacement from the canonical
    // complete audit.
    let replacement = "<r><g><x/><y/><pad>padding padding padding padding</pad></r>";
    let original_start = original.find("<a/>").expect("empty source element");
    let original_end = original.find("</g>").expect("source group close") + "</g>".len();
    let replacement_start = replacement.find("<x/>").expect("empty replacement element");
    let replacement_end = replacement_start + "<x/>".len();
    let spans = [
        ReplacementSpan {
            original: original_start..original_end,
            replacement: replacement_start..replacement_end,
        },
        span(original, replacement, "<b/>", "<y/>"),
    ];
    let actual = checked(
        original.as_bytes(),
        replacement.as_bytes(),
        Limits::default(),
        &spans,
    )
    .expect_err("missing group close must use the canonical replacement error");
    assert!(
        matches!(actual, ReplacementError::Replacement(_)),
        "{actual:?}"
    );
}

#[test]
fn bom_is_framing_and_declarations_inside_windows_keep_precedence() {
    let original = "\u{feff}<r><a/><b/><pad>padding padding padding padding</pad></r>";
    let replacement =
        "\u{feff}<r><a x=\"1\"/><b y=\"2\"/><pad>padding padding padding padding</pad></r>";
    let spans = [
        span(original, replacement, "<a/>", "<a x=\"1\"/>"),
        span(original, replacement, "<b/>", "<b y=\"2\"/>"),
    ];
    assert_windows(original, replacement, Limits::default(), &spans);

    let declaration = "<?xml version=\"1.0\"?><a/>";
    let replacement =
        format!("\u{feff}<r>{declaration}<b/><pad>padding padding padding padding</pad></r>");
    let spans = [
        span(original, &replacement, "<a/>", declaration),
        span(original, &replacement, "<b/>", "<b/>"),
    ];
    let actual = checked(
        original.as_bytes(),
        replacement.as_bytes(),
        Limits::default(),
        &spans,
    )
    .expect_err("embedded declaration");
    assert!(
        matches!(actual, ReplacementError::Replacement(_)),
        "embedded declaration: {actual:?}"
    );

    // A second U+FEFF is legal character data.  It is not a window boundary,
    // so the local window attempt must fall back to the complete oracle.
    let embedded_bom = "\u{feff}<a/>";
    let replacement =
        format!("\u{feff}<r>{embedded_bom}<b/><pad>padding padding padding padding</pad></r>");
    let spans = [
        span(original, &replacement, "<a/>", embedded_bom),
        span(original, &replacement, "<b/>", "<b/>"),
    ];
    let proof = checked(
        original.as_bytes(),
        replacement.as_bytes(),
        Limits::default(),
        &spans,
    )
    .expect("interior BOM is accepted by the complete source oracle");
    assert_eq!(proof, ReplacementProof::Complete);
}

#[test]
fn aggregate_event_attribute_text_and_depth_limits_include_all_windows() {
    let pad = format!("<pad>{}</pad>", "padding ".repeat(128));
    let original = format!("<r><a/><b/><c/>{pad}</r>");
    let replacement = format!(
        "<r><a x=\"1\">alpha<!--a--></a><b y=\"2\"><x>beta</x><?p?></b><c z=\"3\"><![CDATA[gamma]]></c>{pad}</r>"
    );
    let spans = [
        span(
            &original,
            &replacement,
            "<a/>",
            "<a x=\"1\">alpha<!--a--></a>",
        ),
        span(
            &original,
            &replacement,
            "<b/>",
            "<b y=\"2\"><x>beta</x><?p?></b>",
        ),
        span(
            &original,
            &replacement,
            "<c/>",
            "<c z=\"3\"><![CDATA[gamma]]></c>",
        ),
    ];
    let replacement_report = verify_source(replacement.as_bytes(), Limits::default())
        .expect("the aggregate-limit replacement is valid");

    for (resource, value) in [
        (Resource::Events, replacement_report.events()),
        (Resource::Attributes, replacement_report.attributes()),
        (Resource::Bytes, replacement_report.bytes()),
        (Resource::TextBytes, replacement_report.text_bytes()),
        (Resource::Depth, replacement_report.max_depth()),
    ] {
        let exact = Limits::default().narrow(resource, value);
        assert_windows(&original, &replacement, exact, &spans);

        let below = Limits::default().narrow(resource, value - 1);
        let actual = checked(original.as_bytes(), replacement.as_bytes(), below, &spans)
            .expect_err("one below the replacement total must fail");
        assert!(
            matches!(actual, ReplacementError::Replacement(Error::Limit { resource: found, .. }) if found == resource),
            "{resource:?}: {actual:?}"
        );
    }
}

#[test]
fn invalid_layouts_never_publish_a_partial_window_proof() {
    let original = "<r><a>0</a><b>0</b><c>0</c><pad>padding padding padding padding</pad></r>";
    let replacement = "<r><a>1</a><b>2</b><c>3</c><pad>padding padding padding padding</pad></r>";
    let valid_a = span(original, replacement, "<a>0</a>", "<a>1</a>");
    let valid_b = span(original, replacement, "<b>0</b>", "<b>2</b>");
    let valid_c = span(original, replacement, "<c>0</c>", "<c>3</c>");
    let cases = [
        vec![valid_b.clone(), valid_a.clone()],
        vec![
            ReplacementSpan {
                original: valid_a.original.start..valid_b.original.end,
                replacement: valid_a.replacement.start..valid_b.replacement.end,
            },
            valid_b.clone(),
        ],
        vec![
            ReplacementSpan {
                original: 0..0,
                replacement: 0..0,
            },
            valid_a.clone(),
        ],
        vec![valid_a.clone(), valid_c.clone()],
    ];

    for spans in cases {
        let actual = verify_source_replacement_with_spans(
            original.as_bytes(),
            replacement.as_bytes(),
            Limits::default(),
            &spans,
        )
        .expect("the complete replacement is valid");
        assert_eq!(actual, ReplacementProof::Complete);
    }
}

fn generated_case(case: usize) -> (String, String, Vec<ReplacementSpan>) {
    let mut original = String::from(r#"<r xmlns:p="urn:p">"#);
    let mut replacement = original.clone();
    let first = case % 6;
    let second = (case * 3 + 1) % 6;
    let third = (case * 5 + 2) % 6;
    let changed = |index| index == first || index == second || (case % 3 == 0 && index == third);
    let mut spans = Vec::new();

    for index in 0..6 {
        let original_item = if index % 2 == 0 {
            format!(r#"<p:item id="{index}">v{index}</p:item>"#)
        } else {
            format!(r#"<p:empty id="{index}"/>"#)
        };
        let replacement_item = if !changed(index) {
            original_item.clone()
        } else {
            match (case + index) % 3 {
                0 => format!(r#"<p:item id="{index}">value-{case}-{index}</p:item>"#),
                1 => format!(r#"<p:item id="{index}"><p:inner>value-{case}</p:inner></p:item>"#),
                _ => format!(r#"<p:item id="{index}"><!--c--><?p value-{case}?></p:item>"#),
            }
        };
        let original_start = original.len();
        original.push_str(&original_item);
        let replacement_start = replacement.len();
        replacement.push_str(&replacement_item);
        if changed(index) {
            spans.push(ReplacementSpan {
                original: original_start..original.len(),
                replacement: replacement_start..replacement.len(),
            });
        }
    }

    let pad = format!("<pad>{}</pad>", "padding ".repeat(128));
    original.push_str(&pad);
    replacement.push_str(&pad);
    original.push_str("</r>");
    replacement.push_str("</r>");
    (original, replacement, spans)
}

#[test]
fn generated_disjoint_multi_pairs_match_two_complete_audits() {
    for case in 0..96 {
        let (original, replacement, spans) = generated_case(case);
        let proof = checked(
            original.as_bytes(),
            replacement.as_bytes(),
            Limits::default(),
            &spans,
        )
        .unwrap_or_else(|error| panic!("case {case} unexpectedly refused: {error:?}"));
        assert!(
            matches!(proof, ReplacementProof::Windows { count, .. } if count == spans.len()),
            "case {case} did not use every independent window: {proof:?}"
        );
    }
}

#[test]
fn a_large_bounded_window_set_remains_one_public_proof() {
    let count = 128;
    let mut original = String::from("<r>");
    let mut replacement = String::from("<r>");
    let mut spans = Vec::with_capacity(count);
    for index in 0..count {
        let original_start = original.len();
        original.push_str("<a/>");
        let replacement_start = replacement.len();
        replacement.push_str(if index % 2 == 0 { "<b/>" } else { "<c/>" });
        spans.push(ReplacementSpan {
            original: original_start..original.len(),
            replacement: replacement_start..replacement.len(),
        });
    }
    let pad = format!("<pad>{}</pad>", "padding ".repeat(512));
    original.push_str(&pad);
    replacement.push_str(&pad);
    original.push_str("</r>");
    replacement.push_str("</r>");

    let proof = checked(
        original.as_bytes(),
        replacement.as_bytes(),
        Limits::default(),
        &spans,
    )
    .expect("bounded independent windows are valid");
    assert!(matches!(
        proof,
        ReplacementProof::Windows { count: 128, .. }
    ));
}
