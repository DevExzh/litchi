#![cfg(feature = "rtf")]

use litchi_core::{FileFormat, Position};
use litchi_ops::error::Source;
use litchi_ops::transaction::{Error, LimitKind, SemanticLimits};
use litchi_ops::{BatchItem, BatchMode, Code, File, NewNode, Placement, create, open_bytes, props};

const INPUT: &[u8] = br"{\rtf1\ansi First\par Second\par Third}";
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
fn insertion_placements_match_native_bytes_and_report_created_path() {
    for (placement, position) in [
        (Placement::Start, 0),
        (Placement::End, 3),
        (Placement::Position(Position::new(1)), 1),
        (Placement::before("/body/p[2]").unwrap(), 1),
        (Placement::after("/body/p[2]").unwrap(), 2),
        (Placement::inside("/body").unwrap(), 3),
        (Placement::inside_at("/body", Position::new(0)).unwrap(), 0),
    ] {
        let source = file();
        let native = litchi_rtf::Package::from_bytes(INPUT).unwrap();
        let mut edit = native.edit();
        edit.insert_paragraph(Position::new(position), "added")
            .unwrap();
        let expected = edit.commit().unwrap().into_package();
        let result = source
            .add(
                "/body",
                NewNode::new("paragraph")
                    .prop("text", "added")
                    .at(placement),
            )
            .unwrap();
        assert_eq!(
            result.file.to_bytes().unwrap(),
            expected.to_bytes().unwrap()
        );
        assert_eq!(
            result.path.to_string(),
            format!("/body/p[{}]", position + 1)
        );
        assert_eq!(
            result
                .file
                .get(&result.path)
                .unwrap()
                .view()
                .text()
                .unwrap()
                .as_deref(),
            Some("added")
        );
        let mut texts = vec!["First", "Second", "Third"];
        texts.insert(position, "added");
        assert_eq!(paragraphs(&result.file), texts);
        assert_eq!(
            paragraphs(&open_bytes(result.file.to_bytes().unwrap()).unwrap()),
            texts
        );
        assert_eq!(source.to_bytes().unwrap(), INPUT);
    }
}

#[test]
fn empty_owner_and_empty_paragraphs_are_real_structural_insertions() {
    for text in ["", "Hello", "café 水 🦀", " leading", "123", "-12"] {
        let source = create(FileFormat::Rtf).unwrap();
        let result = source
            .add("/body", NewNode::new("p").prop("text", text))
            .unwrap_or_else(|error| {
                panic!("text={text:?}, error={error:?}, native={:?}", error.leaf())
            });
        assert_eq!(result.path.to_string(), "/body/p[1]");
        assert_eq!(paragraphs(&result.file), [text]);
        assert_eq!(
            paragraphs(&open_bytes(result.file.to_bytes().unwrap()).unwrap()),
            [text]
        );
        assert!(paragraphs(&source).is_empty());
    }
    let source = file();
    for (placement, expected) in [
        (Placement::Start, vec!["", "First", "Second", "Third"]),
        (
            Placement::Position(Position::new(1)),
            vec!["First", "", "Second", "Third"],
        ),
        (Placement::End, vec!["First", "Second", "Third", ""]),
    ] {
        let commit = source
            .edit(|edit| edit.add("/body", NewNode::new("p").at(placement)))
            .unwrap();
        assert!(commit.diagnostics().changed());
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
            INPUT
        );
    }
}

#[test]
fn existing_empty_and_terminal_paragraph_boundaries_are_preserved() {
    for input in [
        br"{\rtf1\ansi First\par\par Third}".as_slice(),
        br"{\rtf1\ansi First\par\par Third\par}".as_slice(),
    ] {
        for position in 0..=3 {
            let source = open_bytes(input).unwrap();
            let result = source
                .add(
                    "/body",
                    NewNode::new("p")
                        .prop("text", "new")
                        .at(Placement::Position(Position::new(position))),
                )
                .unwrap_or_else(|error| {
                    panic!(
                        "input={input:?}, position={position}, native={:?}",
                        error.leaf()
                    )
                });
            let mut expected = vec!["First", "", "Third"];
            expected.insert(position, "new");
            assert_eq!(paragraphs(&result.file), expected);
            assert_eq!(
                paragraphs(&open_bytes(result.file.to_bytes().unwrap()).unwrap()),
                expected
            );
        }
    }
}

#[test]
fn add_set_move_remove_resolve_sequentially_and_undo_exactly() {
    let source = create(FileFormat::Rtf).unwrap();
    let before = source.to_bytes().unwrap();
    let commit = source
        .edit(|edit| {
            let a = edit.add("/body", NewNode::new("p").prop("text", "one"))?;
            edit.set(&a, props! {"text" => "changed"})?;
            edit.add("/body", NewNode::new("p").prop("text", "two"))?;
            edit.move_to("/body/p[1]", Placement::End)?;
            edit.remove("/body/p[1]")
        })
        .unwrap();
    assert_eq!(paragraphs(commit.file()), ["changed"]);
    assert_eq!(commit.diagnostics().operations(), 5);
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
        before
    );
    let independent = create(FileFormat::Rtf).unwrap();
    assert_eq!(independent.to_bytes().unwrap(), before);
    assert_eq!(
        commit
            .patch()
            .apply(&independent)
            .unwrap()
            .to_bytes()
            .unwrap(),
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
        before
    );
    assert!(matches!(
        commit.patch().apply(commit.file()).unwrap_err().leaf(),
        Source::Transaction(Error::PatchConflict)
    ));
}

#[test]
fn batches_keep_created_paths_and_attribute_failure() {
    let command = |parent: &'static str, text: &'static str| {
        let mut item = BatchItem::new("add");
        item.parent = Some(parent.into());
        item.kind = Some("paragraph".into());
        item.props = Some(litchi_ops::Props::from_props(props! {"text" => text}).unwrap());
        item
    };
    let items = || {
        [
            command("/body", "new"),
            command("/body/p[1]", "wrong"),
            command("/body", "last"),
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
        ["First", "Second", "Third", "new", "last"]
    );
    assert!(matches!(&best.value()[1], litchi_ops::ItemResult::Failed(e) if e.item() == Some(1)));
    for (index, path) in [(0, "/body/p[4]"), (2, "/body/p[5]")] {
        let litchi_ops::ItemResult::Applied(ref applied) = best.value()[index] else {
            panic!("addition should succeed");
        };
        assert_eq!(applied.path.as_ref().unwrap().to_string(), path);
        assert!(applied.changed);
    }
    assert_eq!(source.to_bytes().unwrap(), INPUT);
}

#[test]
fn invalid_properties_and_destinations_are_atomic() {
    let source = file();
    for node in [
        NewNode::new("run"),
        NewNode::new("p").prop("bold", true),
        NewNode::new("p").prop("text", 7),
        NewNode::new("p")
            .prop("text", "first")
            .prop("text", "second"),
        NewNode::new("p").prop("text", "one\ntwo"),
        NewNode::new("p").at(Placement::Position(Position::new(4))),
        NewNode::new("p").at(Placement::inside("/").unwrap()),
        NewNode::new("p").at(Placement::before("/body").unwrap()),
    ] {
        assert!(source.add("/body", node).is_err());
    }
    assert!(source.add("/body/p[1]", "p").is_err());
    assert_eq!(source.to_bytes().unwrap(), INPUT);
}

#[test]
fn typed_limits_and_protection_are_preserved() {
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
                |edit| edit.add("/body", NewNode::new("p").prop("text", "new")),
            )
            .unwrap_err();
        assert!(
            matches!(error.leaf(), Source::Transaction(Error::Limit {kind: actual, ..}) if *actual == kind),
            "{error:?}"
        );
    }
    let protected = open_bytes(br"{\rtf1\ansi\readprot\enforceprot1 First}".as_slice()).unwrap();
    let error = protected.add("/body", "p").unwrap_err();
    assert_eq!(error.code(), Code::DocumentProtected);
    assert!(matches!(error.leaf(), Source::RtfEdit(_)));
    assert_eq!(source.to_bytes().unwrap(), INPUT);
}

#[test]
fn added_owner_keeps_origin_for_atomic_save() {
    let dir = tempfile::tempdir().unwrap();
    let path = dir.path().join("add.rtf");
    std::fs::write(&path, INPUT).unwrap();
    let source = litchi_ops::open(&path).unwrap();
    source
        .add("/body", NewNode::new("p").prop("text", "saved"))
        .unwrap()
        .file
        .save()
        .unwrap();
    assert_eq!(
        paragraphs(&litchi_ops::open(&path).unwrap()),
        ["First", "Second", "Third", "saved"]
    );
    assert_eq!(std::fs::read_dir(dir.path()).unwrap().count(), 1);
}

#[test]
fn insertion_delimits_controls_without_changing_authored_spaces() {
    for (input, initial) in [
        (br"{\rtf1\ansi}".as_slice(), vec![]),
        (br"{\rtf1\ansi First\par}".as_slice(), vec!["First"]),
        (br"{\rtf1\ansi First\par }".as_slice(), vec!["First"]),
        (br"{\rtf1\ansi\par}".as_slice(), vec![""]),
    ] {
        for text in ["added", "123", "-12", " leading", ""] {
            let source = open_bytes(input).unwrap();
            assert_eq!(paragraphs(&source), initial);
            for position in 0..=initial.len() {
                let result = source
                    .add(
                        "/body",
                        NewNode::new("p")
                            .prop("text", text)
                            .at(Placement::Position(Position::new(position))),
                    )
                    .unwrap_or_else(|error| {
                        panic!(
                            "input={input:?}, position={position}, text={text:?}, native={:?}",
                            error.leaf()
                        )
                    });
                let mut expected = initial.clone();
                expected.insert(position, text);
                assert_eq!(
                    paragraphs(&result.file),
                    expected,
                    "input={input:?}, position={position}, text={text:?}"
                );
            }
        }
    }
}

#[test]
fn encoded_source_empty_insert_never_reports_a_phantom_created_node() {
    let source = create(FileFormat::Rtf)
        .unwrap()
        .add("/body", NewNode::new("p").prop("text", "café 水 🦀"))
        .unwrap()
        .file;
    let before = source.to_bytes().unwrap();
    let added = source.add("/body", "p").unwrap();
    assert_eq!(paragraphs(&added.file), ["café 水 🦀", ""]);
    assert_eq!(
        added
            .file
            .get(&added.path)
            .unwrap()
            .view()
            .text()
            .unwrap()
            .as_deref(),
        Some("")
    );
    assert_eq!(source.to_bytes().unwrap(), before);
}
