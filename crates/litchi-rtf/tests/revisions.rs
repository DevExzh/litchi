#![allow(
    clippy::expect_used,
    clippy::shadow_reuse,
    clippy::shadow_same,
    clippy::shadow_unrelated,
    clippy::unwrap_used,
    reason = "test assertions panic on failure by design and rebind fixture names across steps"
)]

use litchi_rtf::{Revision, RevisionAuthor, RevisionType, RtfDocument, RtfWriter};
use std::borrow::Cow;
use std::fs;

fn write(document: &RtfDocument<'_>) -> Vec<u8> {
    let mut output = Vec::new();
    RtfWriter::new(&mut output)
        .write_document(document)
        .unwrap();
    output
}

#[test]
fn mixed_unicode_revisions_preserve_visible_and_deleted_semantics() {
    let rtf = concat!(
        r#"{\rtf1\ansi{\*\revtbl {Unused;}{Ada \u20320?;}}"#,
        r#"A{\deleted\revauthdel1\revdttmdel12 old \u20320?}B"#,
        r#"{\revised\revauth1\revdttm13 new \u20320?}C}"#,
    );
    let document = RtfDocument::parse(rtf).unwrap();

    assert_eq!(document.text(), "ABnew 你C");
    assert_eq!(
        document
            .revision_authors()
            .iter()
            .map(|author| author.name.as_ref())
            .collect::<Vec<_>>(),
        ["Unused", "Ada 你"]
    );
    assert_eq!(document.revisions().len(), 2);

    let deletion = document
        .revisions()
        .iter()
        .find(|revision| revision.revision_type == RevisionType::Deletion)
        .unwrap();
    assert_eq!(deletion.content, "old 你");
    assert_eq!(deletion.position, 1);
    assert_eq!(deletion.range_end, deletion.position);

    let insertion = document
        .revisions()
        .iter()
        .find(|revision| revision.revision_type == RevisionType::Insertion)
        .unwrap();
    assert_eq!(insertion.content, "new 你");
    assert_eq!(
        document.text().get(insertion.position..insertion.range_end),
        Some(insertion.content.as_ref())
    );

    let output = write(&document);
    let reparsed = RtfDocument::parse_bytes(&output).unwrap_or_else(|error| {
        panic!(
            "failed to parse writer output: {error}\n{}",
            String::from_utf8_lossy(&output)
        )
    });
    assert_eq!(reparsed.text(), document.text());
    assert_eq!(reparsed.revision_authors(), document.revision_authors());
    assert_eq!(reparsed.revisions(), document.revisions());
}

#[test]
fn mutation_uses_utf8_boundaries_and_explicit_author_table() {
    let mut document = RtfDocument::parse(r"{\rtf1\ansi A\u20320?B}").unwrap();
    document
        .push_revision_author(RevisionAuthor::new(Cow::Borrowed("Ada")).unwrap())
        .unwrap();
    document
        .push_revision(Revision {
            revision_type: RevisionType::Insertion,
            author: Cow::Borrowed("Ada"),
            date: Some(Cow::Borrowed("12")),
            id: 0,
            content: Cow::Borrowed("你"),
            position: 1,
            range_end: 4,
        })
        .unwrap();
    document
        .push_revision(Revision {
            revision_type: RevisionType::Deletion,
            author: Cow::Borrowed("Ada"),
            date: Some(Cow::Borrowed("13")),
            id: 0,
            content: Cow::Borrowed("旧"),
            position: 4,
            range_end: 4,
        })
        .unwrap();

    assert!(
        document
            .push_revision(Revision {
                revision_type: RevisionType::Deletion,
                author: Cow::Borrowed("Ada"),
                date: Some(Cow::Borrowed("14")),
                id: 0,
                content: Cow::Borrowed("x"),
                position: 2,
                range_end: 2,
            })
            .is_err()
    );
    assert!(document.clear_revision_authors().is_err());

    let reparsed = RtfDocument::parse_bytes(&write(&document)).unwrap();
    assert_eq!(reparsed.text(), "A你B");
    assert_eq!(reparsed.revisions(), document.revisions());
}

#[test]
fn malformed_revision_grammar_is_rejected() {
    let cases = [
        r"{\rtf1{\*\revtbl{A;}}{\*\revtbl{B;}}}",
        r"{\rtf1{\*\revtbl{A{\field x};}}}",
        r"{\rtf1{\*\revtbl{A;}}{\revised\deleted\revauth0\revdttm1 x}}",
        r"{\rtf1{\*\revtbl{A;}}\revauth0 x}",
        r"{\rtf1{\*\revtbl{A;}}{\revised\revauthdel0\revdttm1 x}}",
        r"{\rtf1{\*\revtbl{A;}}{\revised\revauth0\revdttm1}}",
        r"{\rtf1{\*\revtbl{A;}}{\revised\revauth0\revdttm1{\field x}}}",
        r"{\rtf1{\*\revtbl{A;}}{\deleted\revauthdel0\revdttmdel1\bin3 abc}}",
        r"{\rtf1{\revised\revauth0\revdttm1 x}}",
    ];
    for rtf in cases {
        assert!(RtfDocument::parse(rtf).is_err(), "accepted malformed {rtf}");
    }
}

#[test]
fn handles_bundled_libreoffice_revision_fixtures() {
    const FIXTURES: &[&str] = &[
        "sw/qa/extras/rtfexport/data/redline-insdel.rtf",
        "sw/qa/extras/rtfexport/data/text-change-tracking.rtf",
        "sw/qa/extras/rtfexport/data/redline.rtf",
        "sw/qa/extras/rtfexport/data/fdo55504-1-min.rtf",
        "sw/qa/extras/rtfexport/data/FWDP90_min.rtf",
        "sw/qa/extras/rtfimport/data/tdf167710.rtf",
    ];
    let root = concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../test-data/libreoffice-core"
    );
    for fixture in FIXTURES {
        let bytes = fs::read(format!("{root}/{fixture}")).unwrap();
        if *fixture == "sw/qa/extras/rtfexport/data/FWDP90_min.rtf" {
            assert!(matches!(
                RtfDocument::parse_bytes(&bytes),
                Err(litchi_rtf::RtfError::MalformedDocument(message))
                    if message.contains("trailing non-whitespace")
            ));
            continue;
        }
        let document = RtfDocument::parse_bytes(&bytes)
            .unwrap_or_else(|error| panic!("failed to parse {fixture}: {error}"));
        assert!(
            !document.revision_authors().is_empty(),
            "fixture has no revision authors: {fixture}"
        );
    }
}
