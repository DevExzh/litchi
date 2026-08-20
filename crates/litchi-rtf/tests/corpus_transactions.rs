#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "test assertions panic on failure by design"
)]

use litchi_rtf::{
    Document,
    edit::{Error, Limits, TextSpan, TransferPlan},
};
use std::path::{Path, PathBuf};

fn corpus(relative: &str) -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../..")
        .join(relative)
}

#[test]
fn large_libreoffice_producer_artifacts_round_trip_exactly() {
    let fixtures = [
        "test-data/libreoffice-core/sw/qa/extras/rtfexport/data/tdf158982.rtf",
        "test-data/libreoffice-core/sw/qa/extras/rtfexport/data/watermark.rtf",
        "test-data/libreoffice-core/sw/qa/extras/rtfexport/data/all_gaps_word.rtf",
        "test-data/libreoffice-core/sw/qa/extras/rtfexport/data/tdf167569-2.rtf",
        "test-data/libreoffice-core/sw/qa/extras/rtfexport/data/tdf158762.rtf",
        "test-data/libreoffice-core/sw/qa/extras/rtfexport/data/tdf161878.rtf",
        "test-data/libreoffice-core/sw/qa/extras/rtfexport/data/text-with-comment.rtf",
        "test-data/libreoffice-core/sw/qa/extras/rtfexport/data/tdf158830.rtf",
        "test-data/libreoffice-core/sw/qa/extras/rtfimport/data/tblrepeat.rtf",
        "test-data/libreoffice-core/sw/qa/extras/rtfimport/data/tdf148544.rtf",
        "test-data/libreoffice-core/sw/qa/extras/rtfimport/data/tdf163003.rtf",
        "test-data/libreoffice-core/sw/qa/extras/rtfimport/data/tdf165923.rtf",
    ];
    let mut accepted = 0usize;
    for fixture in fixtures {
        let bytes = std::fs::read(corpus(fixture)).unwrap();
        if let Ok(document) = Document::from_bytes(&bytes) {
            accepted = accepted.saturating_add(1);
            assert_eq!(document.to_bytes().unwrap(), bytes, "fixture: {fixture}");
            let commit = document.edit().commit().unwrap();
            assert!(commit.snapshot().same_snapshot(&document));
            assert_eq!(commit.snapshot().to_bytes().unwrap(), bytes);
        }
    }
    assert!(accepted > 0, "no large producer fixture was accepted");
}

#[test]
fn hostile_large_corpus_is_bounded_and_never_normalized_on_open() {
    let fixtures = [
        "test-data/libreoffice-core/sw/qa/core/data/rtf/fail/forcepoint-4.rtf",
        "test-data/libreoffice-core/sw/qa/core/data/rtf/fail/forcepoint-5.rtf",
        "test-data/libreoffice-core/sw/qa/writerfilter/filters-test/data/pass/TCI-TN65GP-DDRHDLL-partial.rtf",
        "test-data/libreoffice-core/sw/qa/core/data/rtf/pass/forcepoint-1.rtf",
        "test-data/libreoffice-core/sw/qa/core/data/rtf/pass/forcepoint-3.rtf",
        "test-data/libreoffice-core/sw/qa/core/data/rtf/pass/fdo78900.rtf",
        "test-data/libreoffice-core/sw/qa/core/data/rtf/pass/fdo80924.rtf",
        "test-data/libreoffice-core/sw/qa/core/data/rtf/pass/tdf116851.rtf",
    ];
    for fixture in fixtures {
        let bytes = std::fs::read(corpus(fixture)).unwrap();
        if let Ok(document) = Document::from_bytes(&bytes) {
            assert_eq!(document.to_bytes().unwrap(), bytes, "fixture: {fixture}");
        }
    }
}

#[test]
fn lossy_multibyte_tail_retains_text_and_exact_transport() {
    let mut source = br"{\rtf1\ansi\ansicpg932 ordinary ".to_vec();
    source.push(0x82);
    source.push(b'}');

    let document = Document::from_bytes(&source).unwrap();
    assert_eq!(document.text(), "ordinary \u{fffd}");
    assert_eq!(document.to_bytes().unwrap(), source);
}

#[test]
fn hostile_operation_fanout_stops_at_the_caller_bound() {
    let body = "x".repeat(600);
    let source = Document::parse(&format!(r"{{\rtf1\ansi {body}}}")).unwrap();
    let mut edit = source.edit_with_limits(Limits::new(256));
    for position in 0..256 {
        edit.replace_text(TextSpan::new(position, position).unwrap(), "y")
            .unwrap();
    }
    assert!(matches!(
        edit.replace_text(TextSpan::new(300, 300).unwrap(), "z"),
        Err(Error::OperationLimit {
            observed: 257,
            limit: 256
        })
    ));
}

#[test]
fn microsoft_and_libreoffice_corpus_policy_and_output_reopen() {
    let microsoft_bytes = std::fs::read(corpus("test-data/rtf/testNegativeUnicode.rtf")).unwrap();
    let microsoft = Document::from_bytes(&microsoft_bytes).unwrap();
    let microsoft_target = Document::parse(r"{\rtf1\ansi Microsoft interop target}").unwrap();
    assert!(matches!(
        TransferPlan::field(&microsoft, 0, &microsoft_target),
        Err(Error::UnsupportedSource(_))
    ));
    let microsoft_reopened = Document::from_bytes(&microsoft.to_bytes().unwrap()).unwrap();
    assert_eq!(microsoft_reopened.fields().len(), 1);
    assert_eq!(
        microsoft_reopened.fields()[0].instruction,
        microsoft.fields()[0].instruction
    );

    let libreoffice = Document::from_bytes(
        &std::fs::read(corpus(
            "test-data/libreoffice-core/sw/qa/extras/rtfimport/data/ole-inline.rtf",
        ))
        .unwrap(),
    )
    .unwrap();
    let target = Document::parse(r"{\rtf1\ansi Interop target}").unwrap();
    let transferred = TransferPlan::object(&libreoffice, 0, &target)
        .unwrap()
        .commit()
        .unwrap()
        .into_snapshot();
    assert_ne!(transferred.to_bytes().unwrap(), target.to_bytes().unwrap());
    let reopened = Document::from_bytes(&transferred.to_bytes().unwrap()).unwrap();
    assert_eq!(reopened.objects().len(), 1);
    assert_eq!(reopened.objects()[0].data, libreoffice.objects()[0].data);
}

#[test]
fn one_real_libreoffice_run_accepts_checked_italic_edit_when_closure_holds() {
    let fixture = "test-data/libreoffice-core/sw/qa/extras/rtfexport/data/margmirror.rtf";
    let bytes = std::fs::read(corpus(fixture)).unwrap();
    let Ok(document) = Document::from_bytes(&bytes) else {
        panic!("real producer fixture was not accepted: {fixture}");
    };
    let mut body_position = 0usize;
    for paragraph in document.body().paragraphs() {
        let mut run_position = body_position;
        for run in paragraph.runs() {
            let end = run_position.saturating_add(run.text().len());
            if run_position < end
                && let Ok(span) = TextSpan::new(run_position, end)
            {
                let mut edit = document.edit();
                if edit.set_text_italic(span, !run.format().italic()).is_ok()
                    && let Ok(commit) = edit.commit()
                {
                    let reopened =
                        Document::from_bytes(&commit.snapshot().to_bytes().unwrap()).unwrap();
                    assert_eq!(reopened.text(), document.text(), "fixture: {fixture}");
                    return;
                }
            }
            run_position = end;
        }
        body_position = body_position
            .saturating_add(paragraph.len())
            .saturating_add(1);
    }
    panic!("no selected run satisfied the narrow italic closure: {fixture}");
}
