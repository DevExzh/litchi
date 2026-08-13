#![allow(
    clippy::unwrap_used,
    reason = "integration-test assertions panic on failure by design"
)]

use litchi_core::{TextOutputError, TextOutputLimitKind, TextOutputOptions};
use litchi_rtf::Document;
use std::io::{self, Write};

struct FailAfter {
    accepted: Vec<u8>,
    limit: usize,
}

impl Write for FailAfter {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        if self.accepted.len() == self.limit {
            return Err(io::Error::other("injected failure"));
        }
        let accepted = self
            .limit
            .saturating_sub(self.accepted.len())
            .min(bytes.len());
        self.accepted.extend_from_slice(&bytes[..accepted]);
        Ok(accepted)
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

#[test]
fn writes_formatted_utf8_paragraphs_without_flattening_the_document() {
    let document =
        Document::parse(r"{\rtf1\ansi First \b caf\u233?\b0\line\u28450?\u23383?\par Second\par}")
            .unwrap();
    let shared = document.clone();
    let options = TextOutputOptions::new("|", "--", 128, 8);
    let mut output = Vec::new();

    let report = document.write_text_to(&mut output, options).unwrap();

    assert_eq!(
        String::from_utf8(output).unwrap(),
        "First café\n漢字|Second"
    );
    assert_eq!(report.objects_written(), 2);
    assert_eq!(report.bytes_written(), 25);
    assert!(document.same_snapshot(&shared));
}

#[test]
fn empty_and_protected_readable_documents_are_deterministic() {
    let empty = Document::parse(r"{\rtf1\ansi}").unwrap();
    let mut output = Vec::new();
    let report = empty
        .write_text_to(&mut output, TextOutputOptions::default())
        .unwrap();
    assert!(output.is_empty());
    assert_eq!(report.objects_written(), 0);

    let protected = Document::parse(r"{\rtf1\ansi\formprot readable\par}").unwrap();
    protected
        .write_text_to(&mut output, TextOutputOptions::default())
        .unwrap();
    assert_eq!(output, b"readable");
}

#[test]
fn object_limits_and_partial_sink_failures_report_truthful_progress() {
    let document = Document::parse(r"{\rtf1\ansi alpha\par beta\par}").unwrap();
    let options = TextOutputOptions::new("|", "--", 128, 1);
    let mut limited = Vec::new();
    let error = document.write_text_to(&mut limited, options).unwrap_err();
    assert_eq!(limited, b"alpha");
    assert_eq!(error.progress().objects_written(), 1);
    assert_eq!(error.limit().unwrap().kind(), TextOutputLimitKind::Objects);

    let mut failed = FailAfter {
        accepted: Vec::new(),
        limit: 3,
    };
    let error = document
        .write_text_to(&mut failed, TextOutputOptions::default())
        .unwrap_err();
    assert_eq!(failed.accepted, b"alp");
    assert_eq!(error.progress().bytes_written(), 3);
    assert_eq!(error.progress().objects_written(), 0);
    assert!(matches!(error, TextOutputError::Sink { .. }));
}

#[test]
fn empty_paragraphs_use_the_requested_separator_and_inert_destinations_stay_hidden() {
    let document =
        Document::parse(r"{\rtf1\ansi alpha\par\par{\*\litchiunknown hidden}beta\par}").unwrap();
    let mut output = Vec::new();
    let options = TextOutputOptions::new("<P>", "--", 128, 8);

    let report = document.write_text_to(&mut output, options).unwrap();

    assert_eq!(output, b"alpha<P><P>beta");
    assert_eq!(report.objects_written(), 3);
    assert_eq!(report.bytes_written(), 15);

    output.clear();
    let options = options.with_empty_objects(false);
    let report = document.write_text_to(&mut output, options).unwrap();
    assert_eq!(output, b"alpha<P>beta");
    assert_eq!(report.objects_written(), 2);
}
