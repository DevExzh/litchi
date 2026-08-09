#![allow(
    clippy::expect_used,
    clippy::shadow_reuse,
    clippy::shadow_same,
    clippy::shadow_unrelated,
    clippy::unwrap_used,
    reason = "test assertions panic on failure by design and rebind fixture names across steps"
)]

use litchi_rtf::{FitText, RtfDocument, RtfWriter};

fn write(document: &RtfDocument<'_>) -> Vec<u8> {
    let mut output = Vec::new();
    RtfWriter::new(&mut output)
        .write_document(document)
        .unwrap();
    output
}

#[test]
fn parses_fittext_runs_with_continuation_and_round_trips() {
    let document =
        RtfDocument::parse(r"{\rtf1\ansi plain {\fittext1000 Fit this} {\fittext-1 text} done}")
            .unwrap();
    assert_eq!(document.text(), "plain Fit this text done");
    let runs = document.runs();
    let fit_runs: Vec<_> = runs
        .iter()
        .filter(|run| run.formatting.fit_text != FitText::None)
        .collect();
    assert_eq!(fit_runs.len(), 2);
    assert_eq!(fit_runs[0].formatting.fit_text, FitText::Fixed(1000));
    assert_eq!(fit_runs[0].text, "Fit this");
    assert_eq!(fit_runs[1].formatting.fit_text, FitText::Continue);
    assert_eq!(fit_runs[1].text, "text");

    let output = write(&document);
    let serialized = String::from_utf8(output.clone()).unwrap();
    assert!(serialized.contains("\\fittext1000"));
    assert!(serialized.contains("\\fittext-1"));
    let reparsed = RtfDocument::parse_bytes(&output).unwrap();
    assert_eq!(reparsed.text(), document.text());
    let reparsed_runs = reparsed.runs();
    let reparsed_fit: Vec<_> = reparsed_runs
        .iter()
        .filter(|run| run.formatting.fit_text != FitText::None)
        .map(|run| run.formatting.fit_text)
        .collect();
    assert_eq!(reparsed_fit, [FitText::Fixed(1000), FitText::Continue]);
}

#[test]
fn rejects_malformed_fittext_controls() {
    for rtf in [
        // Missing parameter.
        r"{\rtf1{\fittext x}}",
        // Out-of-domain parameters.
        r"{\rtf1{\fittext-2 x}}",
        r"{\rtf1{\fittext1048577 x}}",
    ] {
        assert!(RtfDocument::parse(rtf).is_err(), "accepted malformed {rtf}");
    }
    // Zero is a valid (degenerate) width.
    assert!(RtfDocument::parse(r"{\rtf1{\fittext0 x}}").is_ok());
}

#[test]
fn typed_fittext_domain_conversion() {
    assert_eq!(FitText::from_rtf(-1), Some(FitText::Continue));
    assert_eq!(FitText::from_rtf(0), Some(FitText::Fixed(0)));
    assert_eq!(FitText::from_rtf(1000), Some(FitText::Fixed(1000)));
    assert_eq!(FitText::from_rtf(-2), None);
    assert_eq!(FitText::from_rtf(FitText::MAX_TWIPS + 1), None);
    assert_eq!(FitText::None.rtf_value(), None);
    assert_eq!(FitText::Continue.rtf_value(), Some(-1));
    assert_eq!(FitText::Fixed(42).rtf_value(), Some(42));
}
