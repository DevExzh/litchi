use litchi_core::Position;
use litchi_doc::Package;
use litchi_doc::body_text::{Projection, Snapshot};
use std::io::Cursor;
use std::path::PathBuf;

fn fixture(relative: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../..")
        .join(relative)
}

#[test]
fn real_multi_generation_docs_resize_and_fully_reopen() {
    let fixtures = [
        ("test-data/ole/doc/NoHeadFoot.doc", 0x00C1),
        ("test-data/ole/doc/documentProperties.doc", 0x0101),
    ];
    for (relative, expected_nfib) in fixtures {
        let bytes = std::fs::read(fixture(relative)).expect("real DOC fixture");
        let source = Snapshot::parse(&bytes).expect("fixture has a safe edit basis");
        let paragraphs = source
            .paragraphs(Projection::All)
            .expect("fixture paragraphs");
        let target = paragraphs
            .iter()
            .find(|paragraph| !paragraph.text().is_empty())
            .expect("fixture has an ordinary paragraph");
        let replacement = format!("{} [resized]", target.text());
        let mut transaction = source.edit().expect("fixture transaction");
        transaction
            .replace_paragraph(target.position(), &replacement)
            .expect("fixture length change");
        let commit = transaction.commit().expect("fixture commit");
        let mut package = Package::from_reader(Cursor::new(commit.snapshot().finish()))
            .expect("edited CFB reopens");
        let document = package.document().expect("edited DOC reopens");
        assert_eq!(document.fib().version(), expected_nfib);
        assert_eq!(
            commit
                .snapshot()
                .paragraphs(Projection::All)
                .expect("semantic readback")[target.position().get()]
            .text(),
            replacement
        );
    }
}

#[test]
fn durable_stale_source_is_rejected_without_mutation() {
    let first = Snapshot::parse(
        &std::fs::read(fixture("test-data/ole/doc/NoHeadFoot.doc")).expect("first fixture"),
    )
    .expect("first snapshot");
    let second = Snapshot::parse(
        &std::fs::read(fixture("test-data/ole/doc/ThreeColHeadFoot.doc")).expect("second fixture"),
    )
    .expect("second snapshot");
    let paragraph = first
        .paragraphs(Projection::All)
        .expect("paragraphs")
        .into_iter()
        .find(|paragraph| !paragraph.text().is_empty())
        .expect("ordinary paragraph");
    let mut edit = first.edit().expect("edit");
    edit.replace_paragraph(
        Position::new(paragraph.position().get()),
        "durable replacement",
    )
    .expect("replacement");
    let commit = edit.commit().expect("commit");
    let limits = litchi_core::PatchLimits::new(
        litchi_core::BlobLimits::new(0, 0, 0),
        128 * 1024,
        16,
        8,
        16 * 1024,
        64 * 1024,
    );
    let patch = commit.patch().to_durable(limits).expect("durable patch");
    assert!(second.apply_durable(&patch).is_err());
}
