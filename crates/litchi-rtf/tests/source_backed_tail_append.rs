//! The source-backed RTF tail-append family (record 0359a F2): a commit
//! validates and writes nothing, the committed package writes the planned
//! tail splice, edits chain, and the patch is an exact pair.

#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "these tests intentionally panic on an unexpected result"
)]

use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits as CoreLimits,
    OwnedSource, ReadAt,
};
use litchi_rtf::source_backed::{OpenOptions, Package, tail_append};
use litchi_rtf::tail_append::{TailAppendPublicationError, TailAppendPublicationLimits};
use litchi_rtf::{Package as Document, SourceLimits};
use std::num::{NonZeroU64, NonZeroUsize};
use std::sync::Arc;

const SOURCE: &[u8] = br#"{\rtf1\ansi A\par B}"#;

fn context() -> ExecutionContext {
    let budget = Budget::root(
        "rtf-source-backed-test",
        CoreLimits::new(1 << 20, 1 << 20, 1 << 20, 1_000_000, 64, 1 << 20),
    );
    let (_source, token) = CancellationSource::pair();
    let limits = ExecutionLimits::new(
        NonZeroUsize::new(1).unwrap(),
        NonZeroUsize::new(1).unwrap(),
        NonZeroU64::new(64 * 1024).unwrap(),
        0,
    )
    .unwrap();
    ExecutionContext::new(budget, token, limits)
}

fn open(bytes: &[u8]) -> Package {
    open_with_limits(bytes, TailAppendPublicationLimits::new(3, 2))
}

fn open_with_limits(bytes: &[u8], limits: TailAppendPublicationLimits) -> Package {
    let document = Document::from_bytes(bytes).unwrap();
    let source: Arc<dyn ReadAt> = Arc::new(OwnedSource::new(bytes.to_vec()));
    Package::open_with_options(
        &document,
        source,
        context(),
        OpenOptions::new().publication_limits(limits),
    )
    .unwrap()
}

fn bytes_of(package: &Package) -> Vec<u8> {
    let mut output = Vec::new();
    package.write_to(&mut output).unwrap();
    output
}

fn text_of(bytes: &[u8]) -> String {
    Document::from_bytes(bytes).unwrap().text().to_string()
}

#[test]
fn commit_writes_the_planned_tail_and_leaves_the_source() {
    let source = open(SOURCE);
    let mut edit = source.edit_tail_append().unwrap();
    edit.append_text_paragraphs(&["C"]).unwrap();
    let commit = edit.commit().unwrap();
    assert!(!commit.patch().is_noop());
    let diagnostics = commit.diagnostics();
    assert!(diagnostics.changed());
    assert_eq!(diagnostics.source_bytes(), SOURCE.len());
    assert_eq!(diagnostics.paragraphs(), 1);
    let output = bytes_of(commit.package());
    assert_eq!(output.len(), diagnostics.output_bytes());
    assert_eq!(commit.package().len(), output.len());
    assert_eq!(text_of(&output), "A\nB\nC\n");
    assert_eq!(bytes_of(&source), SOURCE);
}

#[test]
fn a_noop_commit_is_a_share_of_the_source() {
    let source = open(SOURCE);
    let commit = source.edit_tail_append().unwrap().commit().unwrap();
    assert!(commit.patch().is_noop());
    assert!(!commit.diagnostics().changed());
    assert_eq!(bytes_of(commit.package()), SOURCE);
}

#[test]
fn edits_chain_on_a_committed_package() {
    let source = open(SOURCE);
    let mut first = source.edit_tail_append().unwrap();
    first.append_text_paragraphs(&["C"]).unwrap();
    let first = first.commit().unwrap();
    let mut second = first.package().edit_tail_append().unwrap();
    second.append_text_paragraphs(&["D"]).unwrap();
    let second = second.commit().unwrap();
    assert_eq!(text_of(&bytes_of(second.package())), "A\nB\nC\nD\n");
    assert_eq!(text_of(&bytes_of(first.package())), "A\nB\nC\n");
}

#[test]
fn chained_edit_reuses_the_document_open_profile() {
    let source_bytes = br#"{\rtf1\ansi A\par B}"#;
    let paragraphs = ["C"; 8];
    let baseline = open(source_bytes);
    let mut baseline_edit = baseline.edit_tail_append().unwrap();
    baseline_edit.append_text_paragraphs(&paragraphs).unwrap();
    let baseline_first = baseline_edit.commit().unwrap();
    let first_output = bytes_of(baseline_first.package());
    let limits = SourceLimits::builder()
        .max_input_bytes(first_output.len())
        .unwrap()
        .build()
        .unwrap();
    let document_options = litchi_rtf::read::OpenOptions::new().source_limits(limits);
    let document = Document::from_bytes_with_options(source_bytes, document_options).unwrap();
    let source: Arc<dyn ReadAt> = Arc::new(OwnedSource::new(source_bytes.to_vec()));
    let source =
        Package::open_with_options(&document, source, context(), OpenOptions::default()).unwrap();

    let mut first = source.edit_tail_append().unwrap();
    first.append_text_paragraphs(&paragraphs).unwrap();
    let first = first.commit().unwrap();
    assert_eq!(bytes_of(first.package()), first_output);
    let mut second = first.package().edit_tail_append().unwrap();
    second.append_text_paragraphs(&paragraphs).unwrap();
    assert!(second.commit().is_err());
}

#[test]
fn the_patch_is_an_exact_pair() {
    let source = open(SOURCE);
    let mut edit = source.edit_tail_append().unwrap();
    edit.append_text_paragraphs(&["C"]).unwrap();
    let commit = edit.commit().unwrap();
    let replayed = source.apply_tail_append(commit.patch()).unwrap();
    assert_eq!(bytes_of(replayed.package()), bytes_of(commit.package()));
    let restored = commit
        .package()
        .apply_tail_append(&commit.patch().inverse())
        .unwrap();
    assert_eq!(bytes_of(restored.package()), SOURCE);
    let other = open(SOURCE);
    let copied = other.apply_tail_append(commit.patch()).unwrap();
    assert_eq!(bytes_of(copied.package()), bytes_of(commit.package()));
    let independent_target = open(&bytes_of(commit.package()));
    let independently_restored = independent_target
        .apply_tail_append(&commit.patch().inverse())
        .unwrap();
    assert_eq!(bytes_of(independently_restored.package()), SOURCE);
    let strict = open_with_limits(SOURCE, TailAppendPublicationLimits::new(1, 1));
    let strict_copy = strict.apply_tail_append(commit.patch()).unwrap();
    assert_eq!(bytes_of(strict_copy.package()), bytes_of(commit.package()));
    let wrong = open(br#"{\rtf1\ansi A\par D}"#);
    assert!(matches!(
        wrong.apply_tail_append(commit.patch()),
        Err(TailAppendPublicationError::Plan(
            litchi_rtf::tail_append::Error::PatchConflict
        ))
    ));
    assert!(commit.package().apply_tail_append(commit.patch()).is_err());
}

#[test]
fn save_replaces_the_destination_with_the_committed_bytes() {
    let source = open(SOURCE);
    let mut edit = source.edit_tail_append().unwrap();
    edit.append_runs(&[litchi_rtf::PlainRun::new("C")]).unwrap();
    let commit: tail_append::Commit = edit.commit().unwrap();
    let directory = std::env::temp_dir().join(format!(
        "litchi-rtf-source-backed-{}-{}",
        std::process::id(),
        line!()
    ));
    std::fs::create_dir_all(&directory).unwrap();
    let path = directory.join("out.rtf");
    let report = commit.package().save(&path).unwrap();
    let saved = std::fs::read(&path).unwrap();
    std::fs::remove_dir_all(&directory).unwrap();
    assert_eq!(report.bytes(), saved.len());
    assert_eq!(saved, bytes_of(commit.package()));
}

#[test]
fn a_document_that_is_not_its_source_is_refused_at_open() {
    let document = Document::from_bytes(SOURCE).unwrap();
    let other: Arc<dyn ReadAt> = Arc::new(OwnedSource::new(br#"{\rtf1\ansi X\par Y}"#.to_vec()));
    assert!(Package::open(&document, other, context()).is_err());
}
