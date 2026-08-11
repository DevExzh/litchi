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
fn changed_microsoft_and_libreoffice_corpus_output_reopens() {
    let microsoft_bytes = std::fs::read(corpus("test-data/rtf/testNegativeUnicode.rtf")).unwrap();
    let microsoft = Document::from_bytes(&microsoft_bytes).unwrap();
    let microsoft_target = Document::parse(r"{\rtf1\ansi Microsoft interop target}").unwrap();
    let changed = TransferPlan::field(&microsoft, 0, &microsoft_target)
        .unwrap()
        .commit()
        .unwrap()
        .into_snapshot();
    assert_ne!(
        changed.to_bytes().unwrap(),
        microsoft_target.to_bytes().unwrap()
    );
    let reopened = Document::from_bytes(&changed.to_bytes().unwrap()).unwrap();
    assert_eq!(reopened.fields().len(), 1);
    assert_eq!(
        reopened.fields()[0].instruction,
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
