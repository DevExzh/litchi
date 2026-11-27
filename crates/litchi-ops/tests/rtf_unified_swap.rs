#![cfg(feature = "rtf")]

use litchi_core::Position;
use litchi_ops::error::Source;
use litchi_ops::transaction::{Error, LimitKind, SemanticLimits};
use litchi_ops::{BatchItem, BatchMode, Code, File, NewNode, open_bytes, props};

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

#[test]
fn every_pair_matches_native_bytes_and_original_position_order() {
    for first in 0..4 {
        for second in 0..4 {
            let source = file();
            let native = litchi_rtf::Package::from_bytes(INPUT).unwrap();
            let mut edit = native.edit();
            edit.swap_paragraphs(Position::new(first), Position::new(second))
                .unwrap();
            let expected = edit.commit().unwrap().into_package();
            let a = format!("/body/p[{}]", first + 1);
            let b = format!("/body/p[{}]", second + 1);
            let result = source.swap(a.as_str(), b.as_str()).unwrap();
            assert_eq!(result.to_bytes().unwrap(), expected.to_bytes().unwrap());
            let mut texts = vec!["First", "Second", "Third", "Fourth"];
            texts.swap(first, second);
            assert_eq!(paragraphs(&result), texts);
            assert_eq!(
                paragraphs(&open_bytes(result.to_bytes().unwrap()).unwrap()),
                texts
            );
            assert_eq!(source.to_bytes().unwrap(), INPUT);
        }
    }
}

#[test]
fn swapping_empty_and_terminal_paragraphs_preserves_count_and_exact_undo() {
    for input in [
        br"{\rtf1\ansi First\par\par Third}".as_slice(),
        br"{\rtf1\ansi First\par\par Third\par}".as_slice(),
    ] {
        for (a, b) in [(0, 1), (1, 2), (0, 2), (2, 1)] {
            let source = open_bytes(input).unwrap();
            let a_path = format!("/body/p[{}]", a + 1);
            let b_path = format!("/body/p[{}]", b + 1);
            let commit = source
                .edit(|edit| edit.swap(a_path.as_str(), b_path.as_str()))
                .unwrap();
            let mut expected = vec!["First", "", "Third"];
            expected.swap(a, b);
            assert_eq!(paragraphs(commit.file()), expected);
            assert_eq!(
                paragraphs(&open_bytes(commit.file().to_bytes().unwrap()).unwrap()),
                expected
            );
            assert_eq!(
                commit
                    .patch()
                    .inverse()
                    .apply(commit.file())
                    .unwrap()
                    .to_bytes()
                    .unwrap(),
                input
            );
        }
    }
}

#[test]
fn same_position_is_an_exact_noop_even_on_protected_or_formatted_sources() {
    for input in [
        INPUT,
        br"{\rtf1\ansi\readprot\enforceprot1 First\par Second}".as_slice(),
        br"{\rtf1\ansi\b First\par Second}".as_slice(),
    ] {
        let source = open_bytes(input).unwrap();
        let commit = source
            .edit(|edit| edit.swap("/body/p[1]", "/body/paragraph[1]"))
            .unwrap();
        assert!(!commit.diagnostics().changed());
        assert_eq!(commit.diagnostics().operations(), 1);
        assert_eq!(commit.file().to_bytes().unwrap(), input);
    }
}

#[test]
fn swapping_identical_plain_paragraphs_is_an_exact_noop() {
    let bytes = br"{\rtf1\ansi Same\par Same}";
    let source = open_bytes(bytes.as_slice()).unwrap();
    let commit = source
        .edit(|edit| edit.swap("/body/p[1]", "/body/p[2]"))
        .unwrap();
    assert!(!commit.diagnostics().changed());
    assert_eq!(commit.file().to_bytes().unwrap(), bytes);
}

#[test]
fn mixed_edits_keep_sequential_paths_and_content_bound_history() {
    let source = file();
    let commit = source
        .edit(|edit| {
            edit.swap("/body/p[1]", "/body/p[4]")?;
            edit.set("/body/p[1]", props! {"text" => "changed"})?;
            edit.add("/body", NewNode::new("p").prop("text", "last"))?;
            edit.remove("/body/p[2]")
        })
        .unwrap();
    assert_eq!(
        paragraphs(commit.file()),
        ["changed", "Third", "First", "last"]
    );
    assert_eq!(commit.diagnostics().operations(), 4);
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
fn atomic_and_best_effort_batches_attribute_swap_failures() {
    let command = |a: &'static str, b: &'static str| {
        let mut item = BatchItem::new("swap");
        item.path = Some(a.into());
        item.path2 = Some(b.into());
        item
    };
    let items = || {
        [
            command("/body/p[1]", "/body/p[4]"),
            command("/body/p[2]", "/body/p[99]"),
            command("/body/p[1]", "/body/p[2]"),
        ]
    };
    let source = file();
    assert_eq!(
        source.batch(items(), BatchMode::Atomic).unwrap_err().item(),
        Some(1)
    );
    let best = source.batch(items(), BatchMode::BestEffort).unwrap();
    assert_eq!(
        paragraphs(best.file()),
        ["Second", "Fourth", "Third", "First"]
    );
    assert!(
        matches!(&best.value()[1], litchi_ops::ItemResult::Failed(error) if error.item() == Some(1))
    );
    assert_eq!(source.to_bytes().unwrap(), INPUT);
}

#[test]
fn resource_limits_count_distinct_targets_and_preserve_input() {
    let source = file();
    for (limits, kind) in [
        (
            SemanticLimits::builder()
                .max_expanded_targets(1)
                .unwrap()
                .build()
                .unwrap(),
            LimitKind::ExpandedTargets,
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
                |edit| edit.swap("/body/p[1]", "/body/p[2]"),
            )
            .unwrap_err();
        assert!(
            matches!(error.leaf(), Source::Transaction(Error::Limit {kind: actual, ..}) if *actual == kind),
            "{error:?}"
        );
    }
    let limits = SemanticLimits::builder()
        .max_expanded_targets(1)
        .unwrap()
        .build()
        .unwrap();
    let same = source
        .edit_with_options(
            litchi_ops::EditOptions::new().semantic_limits(limits),
            |edit| edit.swap("/body/p[1]", "/body/p[1]"),
        )
        .unwrap();
    assert!(!same.diagnostics().changed());
    let limits = SemanticLimits::builder()
        .max_operations(1)
        .unwrap()
        .build()
        .unwrap();
    let error = source
        .edit_with_options(
            litchi_ops::EditOptions::new().semantic_limits(limits),
            |edit| {
                edit.swap("/body/p[1]", "/body/p[2]")?;
                edit.swap("/body/p[1]", "/body/p[2]")
            },
        )
        .unwrap_err();
    assert_eq!(error.code(), Code::LimitExceeded);
    assert_eq!(source.to_bytes().unwrap(), INPUT);
}

#[test]
fn invalid_targets_and_protection_keep_native_typed_causes() {
    let source = file();
    for (a, b) in [
        ("/body", "/body/p[1]"),
        ("/body/p[1]", "/"),
        ("/body/p[1]", "/body/p[99]"),
    ] {
        assert!(source.swap(a, b).is_err());
    }
    let protected =
        open_bytes(br"{\rtf1\ansi\readprot\enforceprot1 First\par Second}".as_slice()).unwrap();
    let error = protected.swap("/body/p[1]", "/body/p[2]").unwrap_err();
    assert_eq!(error.code(), Code::DocumentProtected);
    assert!(matches!(error.leaf(), Source::RtfEdit(_)));
    assert_eq!(source.to_bytes().unwrap(), INPUT);
}

#[test]
fn swapped_owner_retains_origin_for_atomic_save() {
    let dir = tempfile::tempdir().unwrap();
    let path = dir.path().join("swap.rtf");
    std::fs::write(&path, INPUT).unwrap();
    litchi_ops::open(&path)
        .unwrap()
        .swap("/body/p[1]", "/body/p[4]")
        .unwrap()
        .save()
        .unwrap();
    assert_eq!(
        paragraphs(&litchi_ops::open(&path).unwrap()),
        ["Fourth", "Second", "Third", "First"]
    );
    assert_eq!(std::fs::read_dir(dir.path()).unwrap().count(), 1);
}
