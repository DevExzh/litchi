#![cfg(feature = "rtf")]

use litchi_core::Position;
use litchi_ops::error::Source;
use litchi_ops::transaction::{Error, LimitKind, SemanticLimits};
use litchi_ops::{BatchItem, BatchMode, Code, File, Placement, open_bytes, props};

const INPUT: &[u8] = br"{\rtf1\ansi First\par Second\par Third\par Fourth}";
fn file() -> File {
    open_bytes(INPUT).unwrap()
}
fn paragraphs(file: &File) -> Vec<String> {
    file.query("p")
        .unwrap()
        .map(|n| n.unwrap().view().text().unwrap().unwrap().into_owned())
        .collect()
}

#[test]
fn all_placements_match_native_final_positions() {
    let cases = [
        (1, Placement::Start, 0),
        (1, Placement::End, 3),
        (0, Placement::Position(Position::new(2)), 2),
        (0, Placement::before("/body/p[3]").unwrap(), 1),
        (0, Placement::after("/body/p[3]").unwrap(), 2),
        (3, Placement::before("/body/p[2]").unwrap(), 1),
        (3, Placement::after("/body/p[2]").unwrap(), 2),
        (1, Placement::inside("/body").unwrap(), 3),
        (
            2,
            Placement::inside_at("/body", Position::new(0)).unwrap(),
            0,
        ),
    ];
    for (from, placement, to) in cases {
        let source = file();
        let native = litchi_rtf::Package::from_bytes(INPUT).unwrap();
        let mut edit = native.edit();
        edit.move_paragraph(Position::new(from), Position::new(to))
            .unwrap();
        let expected = edit.commit().unwrap().into_package();
        let path = format!("/body/p[{}]", from + 1);
        let actual = source.move_to(path.as_str(), placement).unwrap();
        assert_eq!(actual.to_bytes().unwrap(), expected.to_bytes().unwrap());
        let mut ordered = vec!["First", "Second", "Third", "Fourth"];
        let moved = ordered.remove(from);
        ordered.insert(to, moved);
        assert_eq!(paragraphs(&actual), ordered);
        assert_eq!(
            paragraphs(&open_bytes(actual.to_bytes().unwrap()).unwrap()),
            ordered
        );
        assert_eq!(source.to_bytes().unwrap(), INPUT);
    }
}

#[test]
fn self_and_adjacent_anchors_are_exact_noops_even_on_protected_source() {
    for input in [
        INPUT,
        br"{\rtf1\ansi\readprot\enforceprot1 First\par Second\par Third\par Fourth}".as_slice(),
    ] {
        let source = open_bytes(input).unwrap();
        for placement in [
            Placement::before("/body/p[2]").unwrap(),
            Placement::after("/body/p[2]").unwrap(),
            Placement::before("/body/p[3]").unwrap(),
            Placement::after("/body/p[1]").unwrap(),
            Placement::Position(Position::new(1)),
        ] {
            let commit = source
                .edit(|edit| edit.move_to("/body/p[2]", placement))
                .unwrap();
            assert!(!commit.diagnostics().changed());
            assert_eq!(commit.file().to_bytes().unwrap(), input);
            assert_eq!(commit.diagnostics().operations(), 1);
        }
    }
}

#[test]
fn explicit_empty_paragraphs_survive_moves_and_reopen() {
    for input in [
        br"{\rtf1\ansi First\par\par Third}".as_slice(),
        br"{\rtf1\ansi First\par\par Third\par}".as_slice(),
    ] {
        let source = open_bytes(input).unwrap();
        for (path, placement, expected) in [
            ("/body/p[2]", Placement::Start, vec!["", "First", "Third"]),
            ("/body/p[2]", Placement::End, vec!["First", "Third", ""]),
            ("/body/p[3]", Placement::Start, vec!["Third", "First", ""]),
        ] {
            let commit = source.edit(|edit| edit.move_to(path, placement)).unwrap();
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
            let result = commit.file();
            assert_eq!(paragraphs(result), expected);
            assert_eq!(
                paragraphs(&open_bytes(result.to_bytes().unwrap()).unwrap()),
                expected
            );
        }
    }
}

#[test]
fn mixed_sequential_edits_have_exact_content_bound_history() {
    let source = file();
    let commit = source
        .edit(|edit| {
            edit.move_to("/body/p[1]", Placement::End)?;
            edit.set("/body/p[1]", props! {"text" => "changed"})?;
            edit.remove("/body/p[2]")
        })
        .unwrap();
    assert_eq!(paragraphs(commit.file()), ["changed", "Fourth", "First"]);
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
fn batches_attribute_failure_and_continue_from_successful_moves() {
    let command = |path: &'static str, index| {
        let mut item = BatchItem::new("move");
        item.path = Some(path.into());
        item.index = Some(index);
        item
    };
    let items = || {
        [
            command("/body/p[1]", 3),
            command("/body/p[99]", 0),
            command("/body/p[2]", 0),
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
        ["Third", "Second", "Fourth", "First"]
    );
    assert!(matches!(&best.value()[1], litchi_ops::ItemResult::Failed(e) if e.item() == Some(1)));
    assert_eq!(source.to_bytes().unwrap(), INPUT);
}

#[test]
fn limits_and_invalid_destinations_leave_source_unchanged() {
    let source = file();
    for (limits, kind) in [
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
                |edit| edit.move_to("/body/p[1]", Placement::End),
            )
            .unwrap_err();
        assert!(
            matches!(error.leaf(), Source::Transaction(Error::Limit {kind: actual, ..}) if *actual == kind),
            "{error:?}"
        );
    }
    let limits = SemanticLimits::builder()
        .max_operations(1)
        .unwrap()
        .build()
        .unwrap();
    let error = source
        .edit_with_options(
            litchi_ops::EditOptions::new().semantic_limits(limits),
            |edit| {
                edit.move_to("/body/p[1]", Placement::End)?;
                edit.move_to("/body/p[1]", Placement::End)
            },
        )
        .unwrap_err();
    assert_eq!(error.code(), Code::LimitExceeded);
    for placement in [
        Placement::Position(Position::new(4)),
        Placement::inside("/").unwrap(),
        Placement::inside("/body/p[2]").unwrap(),
        Placement::before("/body").unwrap(),
        Placement::after("/body/p[99]").unwrap(),
    ] {
        assert!(source.move_to("/body/p[1]", placement).is_err());
    }
    assert!(source.move_to("/body", Placement::End).is_err());
    assert_eq!(source.to_bytes().unwrap(), INPUT);
}

#[test]
fn changed_protected_move_retains_native_error() {
    let source =
        open_bytes(br"{\rtf1\ansi\readprot\enforceprot1 First\par Second}".as_slice()).unwrap();
    let error = source.move_to("/body/p[1]", Placement::End).unwrap_err();
    assert_eq!(error.code(), Code::DocumentProtected);
    assert!(matches!(error.leaf(), Source::RtfEdit(_)));
}

#[test]
fn moved_owner_retains_origin_for_atomic_save() {
    let dir = tempfile::tempdir().unwrap();
    let path = dir.path().join("move.rtf");
    std::fs::write(&path, INPUT).unwrap();
    let source = litchi_ops::open(&path).unwrap();
    source
        .move_to("/body/p[1]", Placement::End)
        .unwrap()
        .save()
        .unwrap();
    assert_eq!(
        paragraphs(&litchi_ops::open(&path).unwrap()),
        ["Second", "Third", "Fourth", "First"]
    );
    assert_eq!(std::fs::read_dir(dir.path()).unwrap().count(), 1);
}
