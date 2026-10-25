#![cfg(feature = "rtf")]

use litchi_core::Position;
use litchi_core::selector::{Path, Scope};
use litchi_ops::error::Source;
use litchi_ops::query::Query;
use litchi_ops::transaction::{Error, LimitKind, SemanticLimits};
use litchi_ops::{BatchItem, BatchMode, Code, File, open_bytes, props};

const INPUT: &[u8] = br"{\rtf1\ansi First\par Second\par Third\par Fourth}";
fn file() -> File {
    open_bytes(INPUT).unwrap()
}
fn paragraphs(file: &File) -> Vec<String> {
    file.query("p")
        .unwrap()
        .map(|node| node.unwrap().view().text().unwrap().unwrap().into_owned())
        .collect()
}
fn query(selector: &'static str) -> Query<'static> {
    Query::parse(selector)
        .unwrap()
        .with_scope(Scope::Path(Box::new(Path::parse("/body").unwrap())))
}
fn command(path: &'static str) -> BatchItem<'static> {
    let mut item = BatchItem::new("remove");
    item.path = Some(path.into());
    item
}

#[test]
fn scalar_removal_matches_native_bytes_and_preserves_input() {
    for (path, position) in [("/body/p[1]", 0), ("/body/p[2]", 1), ("/body/p[last()]", 3)] {
        let source = file();
        let native = litchi_rtf::Package::from_bytes(INPUT).unwrap();
        let mut edit = native.edit();
        edit.remove_paragraph(Position::new(position)).unwrap();
        let expected = edit.commit().unwrap().into_package();
        let actual = source.remove(path).unwrap();
        assert_eq!(actual.to_bytes().unwrap(), expected.to_bytes().unwrap());
        assert_eq!(
            paragraphs(&open_bytes(actual.to_bytes().unwrap()).unwrap()),
            paragraphs(&actual)
        );
        assert_eq!(source.to_bytes().unwrap(), INPUT);
    }
}

#[test]
fn scoped_query_removes_original_positions_in_one_atomic_selection() {
    let source = file();
    let removed = source
        .remove(query("p[text~=Second], p[text~=Fourth]"))
        .unwrap();
    assert_eq!(paragraphs(&removed), ["First", "Third"]);
    let all = source.remove(query("p")).unwrap();
    assert!(paragraphs(&all).is_empty());
    assert!(paragraphs(&open_bytes(all.to_bytes().unwrap()).unwrap()).is_empty());
    assert_eq!(source.to_bytes().unwrap(), INPUT);
    assert_eq!(
        source.remove(query("p[text~=missing]")).unwrap_err().code(),
        Code::ZeroMatches
    );
    assert_eq!(
        source
            .remove(Query::parse("p").unwrap())
            .unwrap_err()
            .code(),
        Code::BareSelectorRejected
    );
}

#[test]
fn adjacent_empty_and_last_paragraph_removal_preserves_survivors() {
    let source = open_bytes(br"{\rtf1\ansi First\par\par Third\par}".as_slice()).unwrap();
    let removed = source.remove("/body/p[2]").unwrap();
    assert_eq!(paragraphs(&removed), ["First", "Third"]);
    let only = open_bytes(br"{\rtf1\ansi Only}".as_slice())
        .unwrap()
        .remove("/body/p[1]")
        .unwrap();
    assert!(paragraphs(&only).is_empty());
}

#[test]
fn mixed_edits_are_sequential_reversible_and_content_bound() {
    let source = file();
    let commit = source
        .edit(|edit| {
            edit.remove("/body/p[2]")?;
            assert_eq!(
                edit.get("/body/p[2]")?.view().text()?.as_deref(),
                Some("Third")
            );
            edit.set("/body/p[2]", props! {"text" => "changed"})?;
            edit.remove("/body/p[3]")
        })
        .unwrap();
    assert_eq!(paragraphs(commit.file()), ["First", "changed"]);
    assert_eq!(commit.diagnostics().operations(), 3);
    assert!(commit.diagnostics().changed());
    assert_eq!(
        commit.patch().apply(&source).unwrap().to_bytes().unwrap(),
        commit.file().to_bytes().unwrap()
    );
    assert_eq!(
        commit
            .patch()
            .inverse()
            .apply(commit.file())
            .unwrap()
            .to_bytes()
            .unwrap(),
        INPUT
    );
    assert_eq!(
        commit.patch().apply(&file()).unwrap().to_bytes().unwrap(),
        commit.file().to_bytes().unwrap()
    );
    let reopened = open_bytes(commit.file().to_bytes().unwrap()).unwrap();
    assert_eq!(
        commit
            .patch()
            .inverse()
            .apply(&reopened)
            .unwrap()
            .to_bytes()
            .unwrap(),
        INPUT
    );
    assert!(matches!(
        commit.patch().apply(commit.file()).unwrap_err().leaf(),
        Source::Transaction(Error::PatchConflict)
    ));
}

#[test]
fn atomic_and_best_effort_batches_keep_sequential_addressing() {
    let source = file();
    let items = || {
        [
            command("/body/p[2]"),
            command("/body/p[99]"),
            command("/body/p[2]"),
        ]
    };
    assert_eq!(
        source.batch(items(), BatchMode::Atomic).unwrap_err().item(),
        Some(1)
    );
    let best = source.batch(items(), BatchMode::BestEffort).unwrap();
    assert_eq!(paragraphs(best.file()), ["First", "Fourth"]);
    assert!(
        matches!(&best.value()[1], litchi_ops::ItemResult::Failed(error) if error.item() == Some(1))
    );
    assert_eq!(source.to_bytes().unwrap(), INPUT);
}

#[test]
fn expansion_operation_and_snapshot_limits_are_typed_and_atomic() {
    let source = file();
    for (limits, kind) in [
        (
            SemanticLimits::builder()
                .max_expanded_targets(2)
                .unwrap()
                .build()
                .unwrap(),
            LimitKind::ExpandedTargets,
        ),
        (
            SemanticLimits::builder()
                .max_operations(2)
                .unwrap()
                .build()
                .unwrap(),
            LimitKind::Operations,
        ),
        (
            SemanticLimits::builder()
                .max_staged_bytes(1)
                .unwrap()
                .build()
                .unwrap(),
            LimitKind::StagedBytes,
        ),
        (
            SemanticLimits::builder()
                .max_snapshot_bytes(1)
                .unwrap()
                .build()
                .unwrap(),
            LimitKind::SnapshotBytes,
        ),
    ] {
        let error = source
            .edit_with_options(
                litchi_ops::EditOptions::new().semantic_limits(limits),
                |edit| edit.remove(query("p")),
            )
            .unwrap_err();
        assert_eq!(error.code(), Code::LimitExceeded);
        assert!(
            matches!(error.leaf(), Source::Transaction(Error::Limit {kind: actual, ..}) if *actual == kind),
            "{error:?}"
        );
    }
    assert_eq!(source.to_bytes().unwrap(), INPUT);
}

#[test]
fn unsupported_targets_and_protected_sources_keep_native_causes() {
    let source = file();
    assert_eq!(
        source.remove("/body").unwrap_err().code(),
        Code::UnsupportedCapability
    );
    assert!(source.remove("/body/p[99]").is_err());
    let source =
        open_bytes(br"{\rtf1\ansi\readprot\enforceprot1 First\par Second}".as_slice()).unwrap();
    let before = source.to_bytes().unwrap();
    let error = source.remove("/body/p[1]").unwrap_err();
    assert_eq!(error.code(), Code::DocumentProtected);
    assert!(matches!(error.leaf(), Source::RtfEdit(_)));
    assert_eq!(source.to_bytes().unwrap(), before);
}

#[test]
fn removal_preserves_origin_and_publishes_reopenable_output() {
    let dir = tempfile::tempdir().unwrap();
    let path = dir.path().join("remove.rtf");
    std::fs::write(&path, INPUT).unwrap();
    let source = litchi_ops::open(&path).unwrap();
    let result = source.remove("/body/p[2]").unwrap();
    result.save().unwrap();
    assert_eq!(
        paragraphs(&litchi_ops::open(&path).unwrap()),
        ["First", "Third", "Fourth"]
    );
    assert_eq!(std::fs::read_dir(dir.path()).unwrap().count(), 1);
}

#[test]
fn removing_neighbors_keeps_explicit_empty_survivors() {
    let source = open_bytes(br"{\rtf1\ansi First\par\par Third}".as_slice()).unwrap();
    let tail = source.remove("/body/p[3]").unwrap();
    assert_eq!(paragraphs(&tail), ["First", ""]);
    let neighbors = source
        .remove(query("p[text~=First], p[text~=Third]"))
        .unwrap();
    assert_eq!(paragraphs(&neighbors), [""]);
    assert_eq!(
        paragraphs(&open_bytes(neighbors.to_bytes().unwrap()).unwrap()),
        [""]
    );
}

#[test]
fn removing_empty_paragraph_keeps_numeric_and_space_prefixed_survivors() {
    for text in ["123", "-12", " leading", "  spaces"] {
        let input = format!(r"{{\rtf1\ansi First\par\par {text}}}");
        let source = open_bytes(input.as_bytes()).unwrap();
        let result = source.remove("/body/p[2]").unwrap();
        assert_eq!(paragraphs(&result), ["First", text]);
    }
}

#[test]
fn scalar_removal_retains_explicit_native_repair_policy() {
    let input = br"{\rtf1\ansi First\par caf\'e9 bad\'g1\par Third";
    let options = litchi_ops::OpenOptions::new().rtf_options(
        litchi_rtf::OpenOptions::new()
            .leniency(litchi_rtf::leniency::Leniency::TolerateSyntaxDefects),
    );
    let source = litchi_ops::open_bytes_with(input.as_slice(), &options).unwrap();
    let result = source.remove("/body/p[1]").unwrap();
    assert_eq!(paragraphs(&result), ["caf\u{e9} bad\u{1}", "Third"]);
    assert_eq!(source.to_bytes().unwrap(), input);
}
