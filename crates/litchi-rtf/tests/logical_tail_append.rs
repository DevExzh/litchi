#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "these tests intentionally panic on an unexpected transaction result"
)]

use litchi_rtf::tail_append::{
    DurableTailAppendPatch, PlainParagraph, PlainRun, TailAppendError, TailAppendLimits,
    TailSelector,
};
use litchi_rtf::{Document, ParseLimits, ProtectionType};
use std::io::{self, Write};

fn patch_limits() -> TailAppendLimits {
    TailAppendLimits::new(8, 32, 128, 1024, 4096, 16 * 1024)
}

fn retained_bytes(document: &Document) -> Vec<u8> {
    let commit = document
        .tail_append(TailSelector::Body)
        .commit()
        .expect("empty append should be valid for this source");
    let mut output = Vec::new();
    commit
        .write_to(&mut output, patch_limits())
        .expect("in-memory sink should accept the retained source");
    output
}

fn replace_once(mut bytes: Vec<u8>, from: &[u8], to: &[u8]) -> Vec<u8> {
    assert_eq!(from.len(), to.len());
    let start = bytes
        .windows(from.len())
        .position(|window| window == from)
        .expect("test wire field should be present");
    bytes
        .get_mut(start..start + from.len())
        .expect("test wire field span should be valid")
        .copy_from_slice(to);
    bytes
}

fn uppercase_after_digest(mut bytes: Vec<u8>) -> Vec<u8> {
    let marker = br#""after":""#;
    let marker_start = bytes
        .windows(marker.len())
        .position(|window| window == marker)
        .expect("after digest field should be present");
    let digest_start = marker_start + marker.len();
    let digest_end = digest_start + 64;
    let digest = bytes
        .get_mut(digest_start..digest_end)
        .expect("after digest span should be valid");
    let byte = digest
        .iter_mut()
        .find(|byte| matches!(**byte, b'a'..=b'f'))
        .expect("SHA-256 test digest should contain a letter");
    *byte = byte.to_ascii_uppercase();
    bytes
}

fn mutate_numeric_field(mut bytes: Vec<u8>, field: &[u8]) -> Vec<u8> {
    let marker = [b'"']
        .into_iter()
        .chain(field.iter().copied())
        .collect::<Vec<_>>();
    let start = bytes
        .windows(marker.len())
        .position(|window| window == marker)
        .expect("numeric field should be present")
        + marker.len();
    let end = bytes
        .get(start..)
        .and_then(|tail| tail.iter().position(|byte| *byte == b',' || *byte == b'}'))
        .map(|offset| start + offset)
        .expect("numeric field should have a delimiter");
    let digit = bytes
        .get_mut(start..end)
        .and_then(|digits| digits.last_mut())
        .expect("numeric field should contain a digit");
    *digit = if *digit == b'9' {
        b'8'
    } else {
        digit.saturating_add(1)
    };
    bytes
}

#[test]
fn empty_append_is_an_exact_noop_and_shares_snapshot() {
    let source = Document::parse("{\\rtf1\\ansi  A\\par B}\n").unwrap();
    let exact = retained_bytes(&source);
    let mut edit = source.tail_append(TailSelector::Body);
    edit.append_text_paragraphs(&[]).unwrap();
    let commit = edit.commit().unwrap();
    assert!(!commit.diagnostics().changed());
    assert!(commit.snapshot().same_snapshot(&source));
    assert_eq!(retained_bytes(commit.snapshot()), exact);
    assert_eq!(commit.diagnostics().inserted_bytes(), 0);

    let mut empty_runs = source.tail_append(TailSelector::Body);
    empty_runs.append_runs(&[]).unwrap();
    let empty_run_commit = empty_runs.commit().unwrap();
    assert!(!empty_run_commit.diagnostics().changed());
    assert!(empty_run_commit.snapshot().same_snapshot(&source));

    let durable = commit.patch().to_durable(patch_limits()).unwrap();
    let durable_inverse = durable.inverse();
    let durable_inverse_json = durable_inverse.to_deterministic_json().unwrap();
    let durable_inverse =
        DurableTailAppendPatch::from_deterministic_json(&durable_inverse_json, patch_limits())
            .unwrap();
    let restored = durable_inverse.apply(&source).unwrap();
    assert!(restored.same_snapshot(&source));
}

#[test]
fn appends_plain_paragraphs_before_the_exact_root_close() {
    let source_bytes = b"{\\rtf1\\ansi  A\\par {\\b B}\\par C} \r\n";
    let source = Document::from_bytes(source_bytes).unwrap();
    let first = [PlainRun::new("D"), PlainRun::new(" + E")];
    let second = [PlainRun::new("F{G}\\H\tI")];
    let paragraphs = [PlainParagraph::new(&first), PlainParagraph::new(&second)];

    let mut edit = source.tail_append_with_limits(TailSelector::Body, patch_limits());
    edit.append_paragraphs(&paragraphs).unwrap();
    let commit = edit.commit().unwrap();
    let mut output = Vec::new();
    commit.write_to(&mut output, patch_limits()).unwrap();

    assert_eq!(commit.snapshot().text(), " A\nB\nC\nD + E\nF{G}\\H\tI\n");
    assert!(output.starts_with(br"{\rtf1\ansi  A\par {\b B}\par C"));
    assert!(output.ends_with(b"} \r\n"));
    assert!(
        output
            .windows(br"{\b B}".len())
            .any(|window| window == br"{\b B}")
    );
    assert!(
        output
            .windows(b"\\plain".len())
            .any(|window| window == b"\\plain")
    );
    assert_eq!(commit.diagnostics().paragraphs(), 2);
    assert_eq!(commit.diagnostics().runs(), 3);
}

#[test]
fn appending_after_an_existing_paragraph_break_does_not_add_an_empty_paragraph() {
    let source = Document::parse(r"{\rtf1\ansi A\par }").unwrap();
    let mut edit = source.tail_append(TailSelector::Body);
    edit.append_text_paragraphs(&["B"]).unwrap();
    let commit = edit.commit().unwrap();
    assert_eq!(commit.snapshot().text(), "A\nB\n");
    assert_eq!(commit.snapshot().paragraph_count(), 2);
}

#[test]
fn unicode_appends_reset_the_source_unicode_fallback_state_locally() {
    let source = Document::parse(r"{\rtf1\ansi\uc0 A}").unwrap();
    let mut edit = source.tail_append(TailSelector::Body);
    edit.append_text_paragraphs(&["é😀"]).unwrap();
    let commit = edit.commit().unwrap();
    assert_eq!(commit.snapshot().text(), "A\né😀\n");
}

#[test]
fn exact_source_patch_and_durable_inverse_round_trip_bytes() {
    let source = Document::parse(r"{\rtf1\ansi Before}").unwrap();
    let mut edit = source.tail_append(TailSelector::Body);
    edit.append_text_paragraphs(&["After"]).unwrap();
    let commit = edit.commit().unwrap();
    let patch = commit.patch();

    let applied = patch.apply(&source).unwrap();
    assert_eq!(retained_bytes(&applied), retained_bytes(commit.snapshot()));
    let restored = patch.inverse().apply(&applied).unwrap();
    assert_eq!(retained_bytes(&restored), retained_bytes(&source));

    let durable = patch.to_durable(patch_limits()).unwrap();
    let encoded = durable.to_deterministic_json().unwrap();
    let decoded =
        DurableTailAppendPatch::from_deterministic_json(&encoded, patch_limits()).unwrap();
    assert_eq!(decoded.to_deterministic_json().unwrap(), encoded);
    let durable_applied = decoded.apply(&source).unwrap();
    assert_eq!(
        retained_bytes(&durable_applied),
        retained_bytes(commit.snapshot())
    );
    let inverse = decoded.inverse();
    let inverse_encoded = inverse.to_deterministic_json().unwrap();
    let inverse_decoded =
        DurableTailAppendPatch::from_deterministic_json(&inverse_encoded, patch_limits()).unwrap();
    let durable_restored = inverse_decoded.apply(&durable_applied).unwrap();
    assert_eq!(retained_bytes(&durable_restored), retained_bytes(&source));

    let foreign = Document::parse(r"{\rtf1\ansi Foreign}").unwrap();
    assert!(matches!(
        patch.apply(&foreign),
        Err(TailAppendError::PatchConflict)
    ));
}

#[test]
fn durable_wire_rejects_noncanonical_digests_forged_noops_and_bounds() {
    let source = Document::parse(r"{\rtf1\ansi Before} ").unwrap();
    let mut edit = source.tail_append(TailSelector::Body);
    edit.append_text_paragraphs(&["After"]).unwrap();
    let commit = edit.commit().unwrap();
    let patch = commit.patch();
    let durable = patch.to_durable(patch_limits()).unwrap();
    let encoded = durable.to_deterministic_json().unwrap();

    let uppercase = uppercase_after_digest(encoded.clone());
    assert!(matches!(
        DurableTailAppendPatch::from_deterministic_json(&uppercase, patch_limits()),
        Err(TailAppendError::DurablePatch(_))
    ));

    let forged_sizes = mutate_numeric_field(encoded.clone(), b"after_bytes\":");
    assert!(matches!(
        DurableTailAppendPatch::from_deterministic_json(&forged_sizes, patch_limits()),
        Err(TailAppendError::DurablePatch(_))
    ));

    let no_op = source
        .tail_append(TailSelector::Body)
        .commit()
        .unwrap()
        .patch()
        .to_durable(patch_limits())
        .unwrap()
        .to_deterministic_json()
        .unwrap();
    let forged_no_op = replace_once(
        no_op,
        br#""direction":"append""#,
        br#""direction":"remove""#,
    );
    assert!(matches!(
        DurableTailAppendPatch::from_deterministic_json(&forged_no_op, patch_limits()),
        Err(TailAppendError::DurablePatch(_))
    ));

    let tiny_output = TailAppendLimits::new(8, 32, 128, 1024, 1, 16 * 1024);
    assert!(matches!(
        DurableTailAppendPatch::from_deterministic_json(&encoded, tiny_output),
        Err(TailAppendError::LimitExceeded {
            resource: "patch offset" | "output bytes",
            ..
        })
    ));

    let too_small_patch = TailAppendLimits::new(8, 32, 128, 1024, 4096, encoded.len() - 1);
    assert!(matches!(
        patch.to_durable(too_small_patch),
        Err(TailAppendError::LimitExceeded {
            resource: "patch bytes",
            ..
        })
    ));
}

#[test]
fn durable_remove_checks_output_bound_before_publish_and_noop_proof() {
    let source = Document::parse(r"{\rtf1\ansi Before} ").unwrap();
    let mut edit = source.tail_append(TailSelector::Body);
    edit.append_text_paragraphs(&["After"]).unwrap();
    let commit = edit.commit().unwrap();
    let inverse = commit.patch().inverse();
    let encoded = inverse
        .to_durable(patch_limits())
        .unwrap()
        .to_deterministic_json()
        .unwrap();
    let low_output = TailAppendLimits::new(
        8,
        32,
        128,
        1024,
        retained_bytes(&source).len() - 1,
        16 * 1024,
    );
    assert!(matches!(
        DurableTailAppendPatch::from_deterministic_json(&encoded, low_output),
        Err(TailAppendError::LimitExceeded {
            resource: "output bytes",
            ..
        })
    ));

    let protected = Document::parse(r"{\rtf1\ansi\readprot\enforceprot1 A}").unwrap();
    let protected_noop = protected
        .tail_append(TailSelector::Body)
        .commit()
        .unwrap()
        .patch()
        .to_durable(patch_limits())
        .unwrap()
        .to_deterministic_json()
        .unwrap();
    let protected_noop =
        DurableTailAppendPatch::from_deterministic_json(&protected_noop, patch_limits()).unwrap();
    assert!(matches!(
        protected_noop.apply(&protected),
        Err(TailAppendError::ProtectedDocument(ProtectionType::ReadOnly))
    ));
}

#[test]
fn durable_append_checks_source_limit_before_reserve() {
    let source_bytes = br"{\rtf1\ansi A}";
    let source = Document::from_bytes(source_bytes).unwrap();
    let mut edit = source.tail_append(TailSelector::Body);
    edit.append_text_paragraphs(&["B"]).unwrap();
    let commit = edit.commit().unwrap();
    let durable = commit.patch().to_durable(patch_limits()).unwrap();
    let encoded = durable.to_deterministic_json().unwrap();
    let strict_source = Document::from_bytes_with_limits(
        source_bytes,
        ParseLimits::default().with_max_source_bytes(source_bytes.len()),
    )
    .unwrap();
    let decoded =
        DurableTailAppendPatch::from_deterministic_json(&encoded, patch_limits()).unwrap();
    assert!(matches!(
        decoded.apply(&strict_source),
        Err(TailAppendError::LimitExceeded {
            resource: "source bytes",
            ..
        })
    ));
}

#[test]
fn staging_and_commit_limits_fail_atomically() {
    let source = Document::parse(r"{\rtf1\ansi A}").unwrap();
    let limits = TailAppendLimits::new(1, 1, 2, 4, 128, 256);
    let mut edit = source.tail_append_with_limits(TailSelector::Body, limits);
    let runs = [PlainRun::new("a"), PlainRun::new("b")];
    let paragraphs = [PlainParagraph::new(&runs)];
    assert!(matches!(
        edit.append_paragraphs(&paragraphs),
        Err(TailAppendError::LimitExceeded {
            resource: "runs",
            observed: 2,
            limit: 1
        })
    ));
    assert_eq!(edit.paragraph_count(), 0);
    edit.append_text_paragraphs(&["ab"]).unwrap();
    assert!(matches!(
        edit.commit(),
        Err(TailAppendError::LimitExceeded {
            resource: "inserted bytes",
            ..
        })
    ));
}

#[test]
fn unsupported_active_opaque_protected_compressed_and_binary_sources_refuse() {
    let active = Document::parse(
        r#"{\rtf1\ansi {\field{\*\fldinst INCLUDETEXT "file:///tmp/x"}{\fldrslt x}}}"#,
    )
    .unwrap();
    let mut active_edit = active.tail_append(TailSelector::Body);
    active_edit.append_text_paragraphs(&["x"]).unwrap();
    assert!(matches!(
        active_edit.commit(),
        Err(TailAppendError::UnsupportedSource(_))
    ));

    let opaque = Document::parse(r"{\rtf1\ansi A\future42 B}").unwrap();
    let mut opaque_edit = opaque.tail_append(TailSelector::Body);
    opaque_edit.append_text_paragraphs(&["x"]).unwrap();
    let opaque_commit = opaque_edit.commit().unwrap();
    let mut opaque_output = Vec::new();
    opaque_commit
        .write_to(&mut opaque_output, patch_limits())
        .unwrap();
    assert!(
        opaque_output
            .windows(br"\future42".len())
            .any(|window| window == br"\future42")
    );
    assert_eq!(opaque_commit.snapshot().text(), "AB\nx\n");

    let protected = Document::parse(r"{\rtf1\ansi\readprot\enforceprot1 A}").unwrap();
    let mut protected_edit = protected.tail_append(TailSelector::Body);
    protected_edit.append_text_paragraphs(&["x"]).unwrap();
    assert!(matches!(
        protected_edit.commit(),
        Err(TailAppendError::ProtectedDocument(ProtectionType::ReadOnly))
    ));

    let compressed = litchi_rtf::transport::compress(br"{\rtf1\ansi A}", true).unwrap();
    let compressed = Document::from_bytes(&compressed).unwrap();
    let mut compressed_edit = compressed.tail_append(TailSelector::Body);
    compressed_edit.append_text_paragraphs(&["x"]).unwrap();
    assert!(matches!(
        compressed_edit.commit(),
        Err(TailAppendError::UnsupportedSource(_))
    ));

    let binary = Document::parse(r"{\rtf1\ansi {\pict\pngblip\bin1 X} A}");
    if let Ok(binary) = binary {
        let mut binary_edit = binary.tail_append(TailSelector::Body);
        binary_edit.append_text_paragraphs(&["x"]).unwrap();
        assert!(matches!(
            binary_edit.commit(),
            Err(TailAppendError::UnsupportedSource(_))
        ));
    }
}

#[test]
fn malformed_durable_envelopes_and_plain_text_controls_refuse() {
    let source = Document::parse(r"{\rtf1\ansi A}").unwrap();
    let mut edit = source.tail_append(TailSelector::Body);
    assert!(matches!(
        edit.append_text_paragraphs(&["bad\ntext"]),
        Err(TailAppendError::InvalidText(_))
    ));
    assert_eq!(edit.paragraph_count(), 0);

    assert!(matches!(
        DurableTailAppendPatch::from_deterministic_json(b"{}", patch_limits()),
        Err(TailAppendError::DurablePatch(_))
    ));
    let limits = TailAppendLimits::new(8, 32, 128, 1024, 4096, 4);
    assert!(matches!(
        DurableTailAppendPatch::from_deterministic_json(br#"{"x":1}"#, limits),
        Err(TailAppendError::LimitExceeded {
            resource: "patch bytes",
            ..
        })
    ));
}

struct PartialSink {
    remaining: usize,
    accepted: usize,
}

impl Write for PartialSink {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        if self.remaining == 0 {
            return Err(io::Error::new(io::ErrorKind::BrokenPipe, "test sink"));
        }
        let count = bytes.len().min(self.remaining).min(3);
        self.remaining -= count;
        self.accepted += count;
        Ok(count)
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

#[test]
fn sequential_sink_reports_partial_progress() {
    let source = Document::parse(r"{\rtf1\ansi A}").unwrap();
    let mut edit = source.tail_append(TailSelector::Body);
    edit.append_text_paragraphs(&["B"]).unwrap();
    let commit = edit.commit().unwrap();
    let mut sink = PartialSink {
        remaining: 7,
        accepted: 0,
    };
    let error = commit.write_to(&mut sink, patch_limits()).unwrap_err();
    assert!(matches!(
        error,
        TailAppendError::Sink {
            kind: io::ErrorKind::BrokenPipe,
            written: 7
        }
    ));
    assert_eq!(sink.accepted, 7);
}
