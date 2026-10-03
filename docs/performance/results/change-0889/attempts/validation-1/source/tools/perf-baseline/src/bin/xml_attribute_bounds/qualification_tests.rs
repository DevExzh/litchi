//! Independent semantic qualification checks for the MCE differential oracle.
//!
//! These tests intentionally live outside the benchmark implementation.  They
//! pin a few small, hand-derived records so a change that makes the baseline
//! and candidate agree on an incomplete transcript cannot be called a
//! correctness pass.

use std::convert::Infallible;
use std::io::{self, BufRead, Cursor, Read};

use litchi_ooxml_common::mce::{
    Capabilities, Error as MceError, Limits, Name, StreamError, StreamLimits,
    process_markup_compatibility, process_markup_compatibility_stream_with_observers,
};
use serde_json::{Value, json};

use super::{codec_oracle, stream_oracle};

const MC: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";

fn field<'a>(record: &'a Value, path: &[&str]) -> &'a Value {
    path.iter().fold(record, |value, key| {
        value
            .get(*key)
            .unwrap_or_else(|| panic!("missing qualification record field {path:?}"))
    })
}

#[test]
fn codec_oracle_matches_a_hand_derived_output_and_report() {
    let source = format!(r#"<x:item xmlns:x="urn:x" xmlns:mc="{MC}"/>"#);
    let capabilities = Capabilities::new();
    let (_, record) = codec_oracle(source.as_bytes(), &capabilities, "bare");

    assert_eq!(field(&record, &["kind"]), "mce_codec");
    assert_eq!(field(&record, &["capabilities"]), "bare");
    assert_eq!(field(&record, &["success"]), true);
    assert_eq!(field(&record, &["input", "bytes"]), 96);
    assert_eq!(
        field(&record, &["output", "bytes"]),
        104,
        "the empty element is emitted as a start/end pair"
    );
    assert_eq!(
        field(&record, &["output", "sha256"]),
        "c43564ef91cd97fe0f983f6a941357c9fd945ab80daa40ff787116f42f1d67ce"
    );
    assert_eq!(
        field(&record, &["report"]),
        &json!({
            "alternate_content_count": 0,
            "selected_choices": 0,
            "selected_fallbacks": 0,
            "ignored_elements": 0,
            "ignored_attributes": 0,
            "preserved_elements": 0,
            "preserved_attributes": 0,
            "unwrapped_elements": 0,
        })
    );
    assert_eq!(field(&record, &["error"]), &Value::Null);
}

#[test]
fn stream_oracle_matches_hand_derived_raw_semantic_transcripts() {
    let source = format!(r#"<x:item xmlns:x="urn:x" xmlns:mc="{MC}"/>"#);
    let (_, record) = stream_oracle(source.as_bytes(), &Capabilities::new(), "bare");

    assert_eq!(field(&record, &["success"]), true);
    assert_eq!(field(&record, &["raw", "events"]), 1);
    assert_eq!(
        field(&record, &["raw", "sha256"]),
        "149e03e47f4fe3f12c0db649aa636aa30e923e1d22d41c12f3a156eb89db8b58"
    );
    assert_eq!(field(&record, &["semantic", "events"]), 1);
    assert_eq!(field(&record, &["semantic", "unknown_events"]), 0);
    assert_eq!(
        field(&record, &["semantic", "sha256"]),
        "e87d5496c37a97d2204a4d4a6373d35f69df10348f8430de6b435255e8ed88cb"
    );
    assert_eq!(
        field(&record, &["report"]),
        &json!({
            "alternate_content_count": 0,
            "selected_choices": 0,
            "selected_fallbacks": 0,
            "ignored_elements": 0,
            "ignored_attributes": 0,
            "preserved_elements": 0,
            "preserved_attributes": 0,
            "unwrapped_elements": 0,
            "events": 1,
        })
    );
    assert_eq!(field(&record, &["error"]), &Value::Null);
}

#[test]
fn stream_oracle_keeps_raw_lexical_values_observable() {
    let escaped = format!(r#"<r xmlns:mc="{MC}" value="&amp;"/>"#);
    let numeric = format!(r#"<r xmlns:mc="{MC}" value="&#38;"/>"#);
    let (_, escaped_record) = stream_oracle(escaped.as_bytes(), &Capabilities::new(), "bare");
    let (_, numeric_record) = stream_oracle(numeric.as_bytes(), &Capabilities::new(), "bare");

    assert_eq!(
        field(&escaped_record, &["semantic", "events"]),
        field(&numeric_record, &["semantic", "events"])
    );
    assert_eq!(
        field(&escaped_record, &["report"]),
        field(&numeric_record, &["report"])
    );
    assert_ne!(
        field(&escaped_record, &["raw", "sha256"]),
        field(&numeric_record, &["raw", "sha256"]),
        "raw_value must distinguish source spellings that decode equally"
    );
    assert_ne!(
        field(&escaped_record, &["semantic", "sha256"]),
        field(&numeric_record, &["semantic", "sha256"]),
        "semantic attributes retain raw source values"
    );
}

#[test]
fn codec_oracle_retains_typed_must_understand_refusal() {
    let source = format!(r#"<r xmlns:mc="{MC}" xmlns:x="urn:unsupported" mc:MustUnderstand="x"/>"#);
    let (_, record) = codec_oracle(source.as_bytes(), &Capabilities::new(), "bare");

    assert_eq!(field(&record, &["success"]), false);
    assert_eq!(
        field(&record, &["error"]),
        &json!({"kind": "must_understand", "message": "urn:unsupported"})
    );
    assert_eq!(field(&record, &["report"]), &Value::Null);
}

#[test]
fn codec_and_stream_oracles_match_hand_derived_directive_semantics() {
    let source = format!(
        r#"<r xmlns:mc="{MC}" xmlns:x="urn:ext" mc:Ignorable="x" mc:PreserveElements="x:keep" mc:PreserveAttributes="x:flag" mc:ProcessContent="x:process"><x:keep plain="yes" x:flag="yes" x:drop="no"><x:drop/><known/></x:keep><x:process><known2/></x:process><x:discard/></r>"#
    );
    let expected = format!(
        r#"<r xmlns:mc="{MC}" xmlns:x="urn:ext"><x:keep plain="yes" x:flag="yes"><known></known></x:keep><known2></known2></r>"#
    );

    let output =
        process_markup_compatibility(source.as_bytes(), &Capabilities::new(), &Limits::default())
            .expect("hand-derived directive fixture must be valid");
    assert_eq!(
        String::from_utf8(output.xml.into_owned()).unwrap(),
        expected
    );
    let (_, codec_record) = codec_oracle(source.as_bytes(), &Capabilities::new(), "bare");
    assert_eq!(field(&codec_record, &["success"]), true);
    assert_eq!(field(&codec_record, &["input", "bytes"]), 318);
    assert_eq!(field(&codec_record, &["output", "bytes"]), 170);
    assert_eq!(
        field(&codec_record, &["output", "sha256"]),
        "5d5a5dc1ccb6f3ccf2e4dca7946543e312b50738e3e89a490ac691c2aa7793d1"
    );
    assert_eq!(
        field(&codec_record, &["report"]),
        &json!({
            "alternate_content_count": 0,
            "selected_choices": 0,
            "selected_fallbacks": 0,
            "ignored_elements": 2,
            "ignored_attributes": 5,
            "preserved_elements": 1,
            "preserved_attributes": 1,
            "unwrapped_elements": 1,
        })
    );

    let (_, stream_record) = stream_oracle(source.as_bytes(), &Capabilities::new(), "bare");
    assert_eq!(field(&stream_record, &["qualification_complete"]), true);
    assert_eq!(field(&stream_record, &["raw", "events"]), 7);
    assert_eq!(field(&stream_record, &["semantic", "events"]), 6);
    assert_eq!(field(&stream_record, &["semantic", "unknown_events"]), 0);
    assert_eq!(
        field(&stream_record, &["report"]),
        &json!({
            "alternate_content_count": 0,
            "selected_choices": 0,
            "selected_fallbacks": 0,
            "ignored_elements": 2,
            "ignored_attributes": 5,
            "preserved_elements": 1,
            "preserved_attributes": 1,
            "unwrapped_elements": 1,
            "events": 10,
        })
    );
}

#[test]
fn codec_and_stream_oracles_match_alias_choice_and_namespace_reset_semantics() {
    let source = format!(
        r#"<r xmlns:mc="{MC}" xmlns:a="urn:x" xmlns:b="urn:x" xmlns="urn:root" mc:Ignorable="a"><mc:AlternateContent><mc:Choice Requires="b"><a:choice/></mc:Choice><mc:Fallback><fallback/></mc:Fallback></mc:AlternateContent><scope xmlns="urn:inner"><child/></scope><reset xmlns=""><child/></reset></r>"#
    );
    let bare_expected = format!(
        r#"<r xmlns:mc="{MC}" xmlns:a="urn:x" xmlns:b="urn:x" xmlns="urn:root"><fallback></fallback><scope xmlns="urn:inner"><child></child></scope><reset xmlns=""><child></child></reset></r>"#
    );
    let custom_expected = format!(
        r#"<r xmlns:mc="{MC}" xmlns:a="urn:x" xmlns:b="urn:x" xmlns="urn:root"><a:choice></a:choice><scope xmlns="urn:inner"><child></child></scope><reset xmlns=""><child></child></reset></r>"#
    );

    let bare = Capabilities::new();
    let mut custom = Capabilities::new();
    custom.understand_namespace("urn:x");
    for (
        capabilities,
        expected,
        output_sha256,
        selected_choices,
        selected_fallbacks,
        ignored_elements,
    ) in [
        (
            &bare,
            bare_expected,
            "f7574be5315ba794d9cf496e1bd4bdb6ebf35ccf1049537d79aaefe977b2d0fd",
            0,
            1,
            1,
        ),
        (
            &custom,
            custom_expected,
            "f2f720a1b9b121376f254221074f5cc7a0c9c338a8aa85e79c84b2a568bed86e",
            1,
            0,
            0,
        ),
    ] {
        let output =
            process_markup_compatibility(source.as_bytes(), capabilities, &Limits::default())
                .expect("choice/fallback namespace fixture must be valid");
        assert_eq!(
            String::from_utf8(output.xml.into_owned()).unwrap(),
            expected
        );

        let (_, codec_record) = codec_oracle(source.as_bytes(), capabilities, "fixture");
        assert_eq!(field(&codec_record, &["success"]), true);
        assert_eq!(field(&codec_record, &["input", "bytes"]), 345);
        assert_eq!(field(&codec_record, &["output", "bytes"]), 235);
        assert_eq!(field(&codec_record, &["output", "sha256"]), output_sha256);
        assert_eq!(
            field(&codec_record, &["report"]),
            &json!({
                "alternate_content_count": 1,
                "selected_choices": selected_choices,
                "selected_fallbacks": selected_fallbacks,
                "ignored_elements": ignored_elements,
                "ignored_attributes": 1,
                "preserved_elements": 0,
                "preserved_attributes": 0,
                "unwrapped_elements": 0,
            })
        );

        let (_, stream_record) = stream_oracle(source.as_bytes(), capabilities, "fixture");
        assert_eq!(field(&stream_record, &["qualification_complete"]), true);
        assert_eq!(field(&stream_record, &["raw", "events"]), 10);
        assert_eq!(field(&stream_record, &["semantic", "events"]), 9);
        assert_eq!(field(&stream_record, &["semantic", "unknown_events"]), 0);
        assert_eq!(
            field(&stream_record, &["report"]),
            &json!({
                "alternate_content_count": 1,
                "selected_choices": selected_choices,
                "selected_fallbacks": selected_fallbacks,
                "ignored_elements": ignored_elements,
                "ignored_attributes": 1,
                "preserved_elements": 0,
                "preserved_attributes": 0,
                "unwrapped_elements": 0,
                "events": 16,
            })
        );
    }
}

#[test]
fn registered_opaque_extension_preserves_lexical_mce_content() {
    let source = format!(
        r#"<r xmlns:mc="{MC}" xmlns:ext="urn:ext" mc:Ignorable="ext"><ext:opaque><mc:AlternateContent>plain<![CDATA[cdata]]></mc:AlternateContent></ext:opaque></r>"#
    );
    let expected = format!(
        r#"<r xmlns:mc="{MC}" xmlns:ext="urn:ext"><ext:opaque><mc:AlternateContent>plain<![CDATA[cdata]]></mc:AlternateContent></ext:opaque></r>"#
    );
    let mut capabilities = Capabilities::new();
    capabilities.preserve_extension_element(Name {
        namespace: "urn:ext".to_owned(),
        local_name: "opaque".to_owned(),
    });

    let output = process_markup_compatibility(source.as_bytes(), &capabilities, &Limits::default())
        .expect("registered extension fixture must be valid");
    assert_eq!(
        String::from_utf8(output.xml.into_owned()).unwrap(),
        expected
    );

    let (_, codec_record) = codec_oracle(source.as_bytes(), &capabilities, "extension:opaque");
    assert_eq!(field(&codec_record, &["success"]), true);
    assert_eq!(field(&codec_record, &["input", "bytes"]), 207);
    assert_eq!(field(&codec_record, &["output", "bytes"]), 188);
    assert_eq!(
        field(&codec_record, &["output", "sha256"]),
        "ede3ff7567340ce0cc399a867b74cb02de58554c0359515ceef11fe52bee14ac"
    );
    assert_eq!(
        field(&codec_record, &["report"]),
        &json!({
            "alternate_content_count": 0,
            "selected_choices": 0,
            "selected_fallbacks": 0,
            "ignored_elements": 0,
            "ignored_attributes": 1,
            "preserved_elements": 0,
            "preserved_attributes": 0,
            "unwrapped_elements": 0,
        })
    );

    let (_, stream_record) = stream_oracle(source.as_bytes(), &capabilities, "extension:opaque");
    assert_eq!(field(&stream_record, &["qualification_complete"]), true);
    assert_eq!(field(&stream_record, &["raw", "events"]), 3);
    assert_eq!(field(&stream_record, &["semantic", "events"]), 8);
    assert_eq!(field(&stream_record, &["semantic", "unknown_events"]), 0);
    assert_eq!(
        field(&stream_record, &["report"]),
        &json!({
            "alternate_content_count": 0,
            "selected_choices": 0,
            "selected_fallbacks": 0,
            "ignored_elements": 0,
            "ignored_attributes": 1,
            "preserved_elements": 0,
            "preserved_attributes": 0,
            "unwrapped_elements": 0,
            "events": 8,
        })
    );
}

#[test]
fn registered_opaque_extension_retains_typed_name_and_duplicate_refusals() {
    let duplicate = format!(
        r#"<r xmlns:mc="{MC}" xmlns:e="urn:extension" xmlns:a="urn:shared" xmlns:b="urn:shared"><e:opaque><child a:k="one" b:k="two"/></e:opaque></r>"#
    );
    let invalid_name = format!(
        r#"<r xmlns:mc="{MC}" xmlns:ext="urn:ext" xmlns:x="urn:x"><ext:opaque><x:bad:name/></ext:opaque></r>"#
    );
    let mut duplicate_capabilities = Capabilities::new();
    duplicate_capabilities.preserve_extension_element(Name {
        namespace: "urn:extension".to_owned(),
        local_name: "opaque".to_owned(),
    });
    let mut invalid_name_capabilities = Capabilities::new();
    invalid_name_capabilities.preserve_extension_element(Name {
        namespace: "urn:ext".to_owned(),
        local_name: "opaque".to_owned(),
    });

    let cases = [
        (
            duplicate.as_bytes(),
            &duplicate_capabilities,
            "duplicate attribute",
            "duplicate attribute",
        ),
        (
            invalid_name.as_bytes(),
            &invalid_name_capabilities,
            "invalid QName",
            "invalid QName: invalid XML QName 'x:bad:name'",
        ),
    ];
    for (input, capabilities, stream_message, codec_message) in cases {
        let (_, codec_record) = codec_oracle(input, capabilities, "extension:opaque");
        assert_eq!(field(&codec_record, &["success"]), false);
        assert_eq!(
            field(&codec_record, &["error"]),
            &json!({"kind": "non_conformant", "message": codec_message})
        );
        assert_eq!(field(&codec_record, &["report"]), &Value::Null);

        let (_, stream_record) = stream_oracle(input, capabilities, "extension:opaque");
        assert_eq!(field(&stream_record, &["success"]), false);
        assert_eq!(field(&stream_record, &["qualification_complete"]), true);
        assert_eq!(
            field(&stream_record, &["error"]),
            &json!({
                "kind": "mce",
                "error": {"kind": "non_conformant", "message": stream_message},
                "prior_mce_error": Value::Null,
                "raw_callback_error": false,
                "active_callback_error": false,
            })
        );
        assert_eq!(field(&stream_record, &["report"]), &Value::Null);
    }
}

/// A reader that delivers the complete source, then reports a transport
/// failure instead of EOF.  This makes the stream's recoverable MCE error and
/// later primary input error observable together.
struct FailAtEof {
    input: Cursor<Vec<u8>>,
    failed: bool,
}

impl FailAtEof {
    fn new(input: Vec<u8>) -> Self {
        Self {
            input: Cursor::new(input),
            failed: false,
        }
    }

    fn failure(&mut self) -> io::Error {
        self.failed = true;
        io::Error::other("qualification input sentinel")
    }
}

impl Read for FailAtEof {
    fn read(&mut self, output: &mut [u8]) -> io::Result<usize> {
        if self.failed {
            return Err(self.failure());
        }
        let count = self.input.read(output)?;
        if count == 0 {
            return Err(self.failure());
        }
        Ok(count)
    }
}

impl BufRead for FailAtEof {
    fn fill_buf(&mut self) -> io::Result<&[u8]> {
        if self.failed {
            return Err(self.failure());
        }
        // Check before borrowing the cursor's buffer so reporting the
        // sentinel cannot overlap the returned slice's borrow.
        if self.input.position() >= self.input.get_ref().len() as u64 {
            return Err(self.failure());
        }
        self.input.fill_buf()
    }

    fn consume(&mut self, amount: usize) {
        self.input.consume(amount);
    }
}

#[test]
fn stream_error_record_preserves_typed_primary_and_prior_mce_failures() {
    let source = format!(
        r#"<r xmlns:mc="{MC}"><mc:AlternateContent><mc:Choice Requires="missing"><bad/></mc:Choice></mc:AlternateContent><later/></r>"#
    );
    let mut input = FailAtEof::new(source.into_bytes());
    let result = process_markup_compatibility_stream_with_observers(
        &mut input,
        &Capabilities::new(),
        &StreamLimits::default(),
        |_| Ok::<(), Infallible>(()),
        |_| Ok::<(), Infallible>(()),
    );

    let error = match &result {
        Err(error) => error,
        Ok(report) => panic!("expected typed input primary with prior MCE error: {report:?}"),
    };
    assert_eq!(
        super::stream_error_value(error),
        json!({
            "kind": "input",
            "error_kind": "Other",
            "message": "qualification input sentinel",
            "prior_mce_error": {
                "kind": "non_conformant",
                "message": "unbound Requires prefix",
            },
            "raw_callback_error": false,
            "active_callback_error": false,
        })
    );

    match result {
        Err(StreamError::Input {
            error,
            prior_mce_error: Some(MceError::NonConformant(message)),
            raw_error: None,
            active_error: None,
            ..
        }) => {
            assert_eq!(error.kind(), io::ErrorKind::Other);
            assert_eq!(error.to_string(), "qualification input sentinel");
            assert_eq!(message, "unbound Requires prefix");
        },
        other => panic!("expected typed input primary with prior MCE error: {other:?}"),
    }
}
