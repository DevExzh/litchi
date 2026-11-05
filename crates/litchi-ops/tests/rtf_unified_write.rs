#![cfg(feature = "rtf")]

use litchi_core::Position;
use litchi_core::selector::{Path, Scope};
use litchi_ops::error::Source;
use litchi_ops::query::Query;
use litchi_ops::transaction::{Error, LimitKind, SemanticLimits};
use litchi_ops::{BatchItem, BatchMode, Code, File, Props, open_bytes, props};

const INPUT: &[u8] = br"{\rtf1\ansi First\line second\par\par Third}";
fn file() -> File {
    open_bytes(INPUT).unwrap()
}
fn text(file: &File, path: &str) -> String {
    file.get(path)
        .unwrap()
        .view()
        .text()
        .unwrap()
        .unwrap()
        .into_owned()
}
fn command(path: &'static str, props: Props<'static>) -> BatchItem<'static> {
    let mut item = BatchItem::new("set");
    item.path = Some(path.into());
    item.props = Some(props);
    item
}

#[test]
fn paragraph_set_preserves_source_and_native_boundaries() {
    let source = file();
    let changed = source
        .set("/body/p[2]", props! {"text" => "雪 é \\ {}"})
        .unwrap();
    assert_eq!(text(&changed, "/body/p[1]"), "First\nsecond");
    assert_eq!(text(&changed, "/body/p[2]"), "雪 é \\ {}");
    assert_eq!(text(&changed, "/body/p[3]"), "Third");
    assert_eq!(source.to_bytes().unwrap(), INPUT);
    let reopened = open_bytes(changed.to_bytes().unwrap()).unwrap();
    assert_eq!(text(&reopened, "/body/p[2]"), "雪 é \\ {}");
}

#[test]
fn query_fanout_matches_one_native_batch() {
    let source = file();
    let changed = source
        .set(paragraph_query(), props! {"text" => "same"})
        .unwrap();
    let native = litchi_rtf::Package::from_bytes(INPUT).unwrap();
    let mut edit = native.edit();
    edit.replace_body_paragraph_texts(&[0, 1, 2].map(|position| {
        litchi_rtf::edit::ParagraphTextReplacement::new(Position::new(position), "same")
    }))
    .unwrap();
    let expected = edit.commit().unwrap().into_package();
    assert_eq!(changed.to_bytes().unwrap(), expected.to_bytes().unwrap());
    assert_eq!(changed.query("p").unwrap().count(), 3);
    assert_eq!(
        source
            .set(
                Query::parse("p[text~=missing]")
                    .unwrap()
                    .with_scope(Scope::Path(Box::new(Path::parse("/body").unwrap()))),
                props! {"text" => "x"}
            )
            .unwrap_err()
            .code(),
        Code::ZeroMatches
    );
}

#[test]
fn history_replays_exact_endpoints_and_accepts_reopened_copies() {
    let source = file();
    let commit = source
        .edit(|edit| {
            edit.set("/body/p[1]", props! {"text" => "PRIVATE replacement"})?;
            assert_eq!(
                edit.get("/body/p[1]")?.view().text()?.as_deref(),
                Some("PRIVATE replacement")
            );
            edit.set("/body/p[3]", props! {"text" => "last"})
        })
        .unwrap();
    assert!(commit.diagnostics().changed());
    assert_eq!(commit.diagnostics().operations(), 2);
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
        commit
            .patch()
            .apply(&file())
            .expect("a byte-identical copy takes the patch")
            .to_bytes()
            .unwrap(),
        commit.file().to_bytes().unwrap()
    );
    assert!(!format!("{:?} {:?}", commit.patch(), commit).contains("PRIVATE"));
    let noop = source
        .edit(|edit| edit.set("/body/p[3]", props! {"text" => "Third"}))
        .unwrap();
    assert!(!noop.diagnostics().changed());
    assert!(noop.patch().is_noop());
    assert_eq!(noop.file().to_bytes().unwrap(), INPUT);
    assert!(commit.patch().apply(noop.file()).is_ok());
}

#[test]
fn failed_edits_are_atomic_and_best_effort_keeps_successes() {
    let source = file();
    let commands = || {
        [
            command("/body/p[1]", Props::new().with("text", "first")),
            command("/body/p[2]", Props::new().with("bold", true)),
            command("/body/p[3]", Props::new().with("text", "last")),
        ]
    };
    assert_eq!(
        source
            .batch(commands(), BatchMode::Atomic)
            .unwrap_err()
            .item(),
        Some(1)
    );
    let result = source.batch(commands(), BatchMode::BestEffort).unwrap();
    assert_eq!(text(result.file(), "/body/p[1]"), "first");
    assert_eq!(text(result.file(), "/body/p[2]"), "");
    assert_eq!(text(result.file(), "/body/p[3]"), "last");
    assert!(
        matches!(&result.value()[1], litchi_ops::ItemResult::Failed(error) if error.item() == Some(1))
    );
    assert!(
        source
            .edit(|edit| {
                edit.set("/body/p[1]", props! {"text" => "discard"})?;
                let _ = edit.set("/body/p[2]", props! {"text" => true});
                Ok(())
            })
            .is_err()
    );
    assert_eq!(source.to_bytes().unwrap(), INPUT);
}

#[test]
fn expansion_staging_and_snapshot_limits_refuse_candidates() {
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
                .max_staged_bytes(5)
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
                |edit| edit.set(paragraph_query(), props! {"text" => "long"}),
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
fn protected_publication_refuses_changes_but_accepts_exact_noops() {
    let input = br"{\rtf1\ansi\readprot\enforceprot1 Same}";
    let source = open_bytes(input.as_slice()).unwrap();
    let error = source
        .set("/body/p[1]", props! {"text" => "different"})
        .unwrap_err();
    assert_eq!(error.code(), Code::DocumentProtected);
    assert!(matches!(
        error.leaf(),
        Source::RtfEdit(litchi_rtf::edit::Error::Refused(refusal)) if refusal.kind() == litchi_core::edit::RefusalKind::Protected
    ));
    assert_eq!(
        source
            .set("/body/p[1]", props! {"text" => "Same"})
            .unwrap()
            .to_bytes()
            .unwrap(),
        input
    );
}

#[test]
fn edited_handle_keeps_origin_for_atomic_save() {
    let dir = tempfile::tempdir().unwrap();
    let path = dir.path().join("document.rtf");
    std::fs::write(&path, INPUT).unwrap();
    let source = litchi_ops::open(&path).unwrap();
    let changed = source
        .set("/body/p[2]", props! {"text" => "saved"})
        .unwrap();
    changed.save().unwrap();
    assert_eq!(
        text(&litchi_ops::open(&path).unwrap(), "/body/p[2]"),
        "saved"
    );
    assert_eq!(source.to_bytes().unwrap(), INPUT);
    assert_eq!(std::fs::read_dir(dir.path()).unwrap().count(), 1);
}

#[test]
fn retained_native_source_policy_rejects_oversized_replacement() {
    let limits = litchi_rtf::SourceLimits::builder()
        .max_input_bytes(INPUT.len())
        .unwrap()
        .build()
        .unwrap();
    let source = litchi_ops::open_bytes_with(
        INPUT,
        &litchi_ops::OpenOptions::new()
            .rtf_options(litchi_rtf::OpenOptions::new().source_limits(limits)),
    )
    .unwrap();
    let error = source
        .set("/body/p[1]", props! {"text" => "x".repeat(1024)})
        .unwrap_err();
    assert_eq!(error.code(), Code::LimitExceeded);
    assert!(matches!(error.leaf(), Source::RtfEdit(_)), "{error:?}");
    assert_eq!(source.to_bytes().unwrap(), INPUT);
}

#[test]
fn snapshot_admission_matches_reported_retained_charge() {
    let source = file();
    let commit = source
        .edit(|edit| edit.set("/body/p[1]", props! {"text" => "x"}))
        .unwrap();
    let weight = commit.diagnostics().snapshot_bytes();
    assert!(weight > source.to_bytes().unwrap().len() + commit.file().to_bytes().unwrap().len());
    for (maximum, expected) in [(weight, true), (weight - 1, false)] {
        let limits = SemanticLimits::builder()
            .max_snapshot_bytes(maximum)
            .unwrap()
            .build()
            .unwrap();
        let result = source.edit_with_options(
            litchi_ops::EditOptions::new().semantic_limits(limits),
            |edit| edit.set("/body/p[1]", props! {"text" => "x"}),
        );
        assert_eq!(result.is_ok(), expected);
    }
}

#[test]
fn restoring_text_does_not_hide_changes_to_the_native_artifact() {
    let source = open_bytes(br"{\rtf1\ansi ABC}".as_slice()).unwrap();
    let first = source
        .set("/body/p[1]", props! {"text" => "temporary"})
        .unwrap();
    let expected = first.set("/body/p[1]", props! {"text" => "ABC"}).unwrap();
    let commit = source
        .edit(|edit| {
            edit.set("/body/p[1]", props! {"text" => "temporary"})?;
            edit.set("/body/p[1]", props! {"text" => "ABC"})
        })
        .unwrap();
    assert_eq!(
        commit.file().to_bytes().unwrap(),
        expected.to_bytes().unwrap()
    );
    assert_eq!(
        commit.diagnostics().changed(),
        source.to_bytes().unwrap() != expected.to_bytes().unwrap()
    );
    assert_eq!(
        commit
            .patch()
            .inverse()
            .apply(commit.file())
            .unwrap()
            .to_bytes()
            .unwrap(),
        source.to_bytes().unwrap()
    );
}

fn paragraph_query() -> Query<'static> {
    Query::parse("p")
        .unwrap()
        .with_scope(Scope::Path(Box::new(Path::parse("/body").unwrap())))
}

#[test]
fn mixed_formatting_refusal_keeps_exact_input_and_native_cause() {
    let input = br"{\rtf1\ansi A{\b B}C}";
    let source = open_bytes(input.as_slice()).unwrap();
    let error = source
        .set("/body/p[1]", props! {"text" => "replacement"})
        .unwrap_err();
    assert_eq!(error.code(), Code::UnsupportedCapability);
    assert!(matches!(error.leaf(), Source::RtfEdit(_)));
    assert_eq!(source.to_bytes().unwrap(), input);
}

#[test]
fn capabilities_publish_text_writes_and_structural_gaps() {
    use litchi_ops::capabilities::Verb;
    let source = file();
    let capabilities = source.capabilities();
    for verb in [
        Verb::Set,
        Verb::Remove,
        Verb::Move,
        Verb::Add,
        Verb::Swap,
        Verb::Edit,
        Verb::Batch,
    ] {
        assert!(capabilities.verb(verb).is_supported());
    }
    assert!(
        capabilities
            .element("p")
            .unwrap()
            .property("text")
            .unwrap()
            .set()
    );
}

#[test]
fn empty_paragraph_insertions_preserve_leading_spaces_and_signs() {
    for replacement in [" leading", "-12", "123", "\t tab", ""] {
        let source = file();
        let changed = source
            .set("/body/p[2]", props! {"text" => replacement})
            .unwrap();
        assert_eq!(text(&changed, "/body/p[2]"), replacement);
        assert_eq!(text(&changed, "/body/p[1]"), "First\nsecond");
    }
}
