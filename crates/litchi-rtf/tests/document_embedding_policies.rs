#![allow(
    clippy::expect_used,
    clippy::shadow_reuse,
    clippy::shadow_same,
    clippy::shadow_unrelated,
    clippy::unwrap_used,
    reason = "test assertions panic on failure by design and rebind fixture names across steps"
)]

use litchi_rtf::{DocumentEmbeddingPolicies, RtfDocument, RtfWriter};

fn write(document: &RtfDocument<'_>) -> Vec<u8> {
    let mut output = Vec::new();
    RtfWriter::new(&mut output)
        .write_document(document)
        .unwrap();
    output
}

#[test]
fn parses_all_explicit_producer_combinations() {
    for (system, linguistic) in [(false, false), (false, true), (true, false), (true, true)] {
        let source = format!(
            r"{{\rtf1\donotembedsysfont{}\donotembedlingdata{} Body}}",
            i32::from(system),
            i32::from(linguistic)
        );
        let document = RtfDocument::parse(&source).unwrap();
        assert_eq!(
            *document.embedding_policies(),
            DocumentEmbeddingPolicies {
                do_not_embed_system_fonts: Some(system),
                do_not_embed_linguistic_data: Some(linguistic),
            }
        );
        assert_eq!(document.text(), "Body");
    }
}

#[test]
fn omission_preserves_unspecified_linguistic_policy_and_system_font_default() {
    let document = RtfDocument::parse(r"{\rtf1 Body}").unwrap();
    assert!(document.embedding_policies().is_empty());
    assert!(
        document
            .embedding_policies()
            .effective_do_not_embed_system_fonts()
    );
    assert_eq!(
        document.embedding_policies().do_not_embed_linguistic_data,
        None
    );
    let serialized = String::from_utf8(write(&document)).unwrap();
    assert!(!serialized.contains("donotembedsysfont"));
    assert!(!serialized.contains("donotembedlingdata"));
}

#[test]
fn typed_api_round_trips_in_stable_order_and_clears_without_embedding() {
    let mut document = RtfDocument::parse(r"{\rtf1 Body}").unwrap();
    document.set_embedding_policies(DocumentEmbeddingPolicies {
        do_not_embed_system_fonts: Some(false),
        do_not_embed_linguistic_data: Some(true),
    });
    let output = write(&document);
    let serialized = String::from_utf8(output.clone()).unwrap();
    assert!(
        serialized.find("\\donotembedsysfont0").unwrap()
            < serialized.find("\\donotembedlingdata1").unwrap()
    );
    let reparsed = RtfDocument::parse_bytes(&output).unwrap();
    assert_eq!(reparsed.embedding_policies(), document.embedding_policies());
    assert_eq!(reparsed.text(), "Body");

    document.clear_embedding_policies();
    assert!(document.embedding_policies().is_empty());
    assert_eq!(document.text(), "Body");
}

#[test]
fn coexists_with_embedded_font_and_xml_policy_metadata_without_activation() {
    let document = RtfDocument::parse(concat!(
        r#"{\rtf1\donotembedlingdata0\validatexml1"#,
        r#"\donotembedsysfont1\relyonvml0 Body}"#,
    ))
    .unwrap();
    let output = write(&document);
    let reparsed = RtfDocument::parse_bytes(&output).unwrap();
    assert_eq!(reparsed.embedding_policies(), document.embedding_policies());
    assert_eq!(reparsed.xml_policies(), document.xml_policies());
    assert_eq!(reparsed.text(), "Body");
}

#[test]
fn rejects_missing_non_boolean_overflow_and_duplicate_values() {
    for name in ["donotembedsysfont", "donotembedlingdata"] {
        for suffix in ["", "-1", "2", "32767", "99999999999"] {
            let source = format!(r"{{\rtf1\{name}{suffix} Body}}");
            assert!(
                RtfDocument::parse(&source).is_err(),
                "accepted malformed {source}"
            );
        }
        let source = format!(r"{{\rtf1\{name}0\{name}1 Body}}");
        assert!(
            RtfDocument::parse(&source).is_err(),
            "accepted duplicate {source}"
        );
    }
}

#[test]
fn rejects_every_starred_grouped_and_late_embedding_policy() {
    for control in [r"\donotembedsysfont1", r"\donotembedlingdata0"] {
        for source in [
            format!(r"{{\rtf1{{\*{control}}}Body}}"),
            format!(r"{{\rtf1{{{control}}}Body}}"),
            format!(r"{{\rtf1 Body{control}}}"),
        ] {
            assert!(
                RtfDocument::parse(&source).is_err(),
                "accepted malformed {source}"
            );
        }
    }
}
