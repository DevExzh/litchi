//! Source-backed PPTX slide insertion through the unified API.
#![cfg(feature = "pptx")]

use litchi_core::{FileFormat, Position};
use litchi_ops::transaction::SemanticLimits;
use litchi_ops::{
    BatchItem, BatchMode, Code, File, ItemResult, NewNode, OpenOptions, Placement, Props,
    SourceMode,
};

fn fixture() -> Vec<u8> {
    let mut builder = litchi_pptx::builder::Builder::new().unwrap();
    for text in ["A", "B", "C"] {
        let slide = builder.presentation().unwrap().add_slide().unwrap();
        slide.add_text_box(text, 0, 0, 914_400, 914_400);
    }
    builder.build().unwrap().to_bytes().unwrap()
}

fn empty_fixture() -> Vec<u8> {
    litchi_pptx::builder::Builder::new()
        .unwrap()
        .build()
        .unwrap()
        .to_bytes()
        .unwrap()
}

fn open_source(bytes: Vec<u8>) -> File {
    litchi_ops::open_bytes_with(bytes, &OpenOptions::new().source_mode(SourceMode::Auto)).unwrap()
}

fn text(file: &File, position: usize) -> String {
    let path = format!("/slide[{position}]/shape[1]");
    file.get(&path)
        .unwrap()
        .view()
        .text()
        .unwrap()
        .unwrap()
        .into_owned()
}

fn slide(layout: &'static str) -> NewNode<'static> {
    NewNode::new("slide").prop("layout", layout)
}

fn add_item(layout: &'static str) -> BatchItem<'static> {
    let mut item = BatchItem::new("add");
    item.parent = Some("/".into());
    item.kind = Some("slide".into());
    item.props = Some(Props::new().with("layout", layout));
    item
}

#[test]
fn source_add_resolves_exact_layout_and_all_slide_placements() {
    let bytes = fixture();
    let source = open_source(bytes.clone());
    assert_eq!(source.capabilities().retains_source(), Some(true));
    assert!(
        source
            .capabilities()
            .element("slide")
            .unwrap()
            .operations()
            .add
    );
    assert!(
        source
            .capabilities()
            .element("shape")
            .unwrap()
            .operations()
            .add
    );

    for (placement, expected) in [
        (Placement::Start, "/slide[1]"),
        (Placement::End, "/slide[4]"),
        (Placement::Position(Position::new(1)), "/slide[2]"),
        (Placement::before("/slide[2]").unwrap(), "/slide[2]"),
        (Placement::after("/slide[2]").unwrap(), "/slide[3]"),
        (Placement::inside("/").unwrap(), "/slide[4]"),
    ] {
        let added = source.add("/", slide("Blank").at(placement)).unwrap();
        assert_eq!(added.path.to_string(), expected);
        assert!(added.file.get(expected).is_ok());
    }

    let before = source.to_bytes().unwrap();
    let error = source
        .add("/", slide("blank"))
        .expect_err("layout matching is exact and case-sensitive");
    assert_eq!(error.code(), Code::NotFound);
    assert_eq!(source.to_bytes().unwrap(), before);

    let shape = source
        .add(
            "/slide[1]",
            NewNode::new("shape")
                .prop("width", "1in")
                .prop("height", "1in"),
        )
        .expect("source shape insertion uses the native source transaction");
    assert_eq!(shape.path.to_string(), "/slide[1]/shape[2]");
    assert_eq!(source.to_bytes().unwrap(), before);
}

#[test]
fn source_add_works_for_an_empty_reopened_source_and_after_source_removal() {
    let empty_bytes = empty_fixture();
    let empty = open_source(empty_bytes.clone());
    let added = empty
        .add("/", slide("Blank"))
        .expect("insert into an empty reopened source");
    assert_eq!(added.path.to_string(), "/slide[1]");
    assert!(added.file.get("/slide[1]").is_ok());
    assert_eq!(empty.to_bytes().unwrap(), empty_bytes);

    let reopened = open_source(added.file.to_bytes().unwrap());
    let appended = reopened
        .add("/", slide("Blank"))
        .expect("append to the reopened source revision");
    assert_eq!(appended.path.to_string(), "/slide[2]");

    let original = open_source(fixture());
    let removed = original
        .remove("/slide[2]")
        .expect("source removal before insertion");
    let inserted = removed
        .add("/", slide("Blank"))
        .expect("insert after source removal");
    assert_eq!(inserted.path.to_string(), "/slide[3]");
    assert_eq!(text(&inserted, 1), "A");
    assert_eq!(text(&inserted, 2), "C");
    assert!(inserted.get("/slide[3]").is_ok());
}

#[test]
fn source_add_composes_with_edit_and_replays_inverse() {
    let source = open_source(fixture());
    let before = source.to_bytes().unwrap();
    let commit = source
        .edit(|edit| {
            let path = edit.add("/", slide("Blank").at(Placement::Start))?;
            assert_eq!(path.to_string(), "/slide[1]");
            edit.set(
                "/slide[2]/shape[1]",
                litchi_ops::props! { "text" => "edited after insertion" },
            )?;
            Ok(())
        })
        .expect("source edit with insertion and a following text edit");
    assert!(commit.diagnostics().changed());
    assert_eq!(text(commit.file(), 2), "edited after insertion");
    assert!(commit.file().get("/slide[1]").is_ok());

    let result = commit.file().to_bytes().unwrap();
    assert_eq!(
        commit.patch().apply(&source).unwrap().to_bytes().unwrap(),
        result
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
    assert_eq!(source.to_bytes().unwrap(), before);
}

#[test]
fn source_add_batches_are_atomic_or_best_effort_with_diagnostics() {
    let source = open_source(fixture());
    let before = source.to_bytes().unwrap();
    let commands = || [add_item("Blank"), add_item("missing"), add_item("Blank")];

    let atomic = source
        .batch(commands(), BatchMode::Atomic)
        .expect_err("atomic source insertion must roll back on a later failure");
    assert_eq!(atomic.item(), Some(1));
    assert_eq!(atomic.code(), Code::NotFound);
    assert_eq!(source.to_bytes().unwrap(), before);

    let best_effort = source
        .batch(commands(), BatchMode::BestEffort)
        .expect("best-effort source insertion");
    assert!(matches!(
        &best_effort.value()[0],
        ItemResult::Applied(applied) if applied.changed
    ));
    assert!(matches!(
        &best_effort.value()[1],
        ItemResult::Failed(error) if error.item() == Some(1) && error.code() == Code::NotFound
    ));
    assert!(matches!(
        &best_effort.value()[2],
        ItemResult::Applied(applied) if applied.changed
    ));
    assert!(best_effort.diagnostics().changed());
    assert!(best_effort.file().get("/slide[5]").is_ok());
    assert_eq!(source.to_bytes().unwrap(), before);
}

#[test]
fn source_add_honors_retention_and_staging_limits_before_mutation() {
    let bytes = fixture();
    let released = litchi_ops::open_bytes_with(
        bytes.clone(),
        &OpenOptions::new()
            .retain_source(false)
            .source_mode(SourceMode::Auto),
    )
    .unwrap();
    let error = released
        .add("/", slide("Blank"))
        .expect_err("released source cannot start a native edit");
    assert_eq!(error.code(), Code::UnsupportedCapability);
    assert_eq!(released.to_bytes().unwrap(), bytes);

    let source = open_source(bytes.clone());
    let limits = SemanticLimits::builder()
        .max_staged_bytes(4)
        .unwrap()
        .build()
        .unwrap();
    let error = source
        .edit_with_options(
            litchi_ops::EditOptions::new().semantic_limits(limits),
            |edit| edit.add("/", slide("Blank")),
        )
        .expect_err("the staged action charge must be checked before insertion");
    assert_eq!(error.code(), Code::LimitExceeded);
    assert_eq!(source.to_bytes().unwrap(), bytes);
}

#[test]
fn source_add_refuses_out_of_range_and_bad_parent_atomically() {
    let source = open_source(fixture());
    let before = source.to_bytes().unwrap();
    let out_of_range = source
        .add(
            "/",
            slide("Blank").at(Placement::Position(Position::new(99))),
        )
        .expect_err("position beyond the catalog must be refused");
    assert_eq!(out_of_range.code(), Code::OutOfRange);
    let bad_parent = source
        .add("/slide[1]", slide("Blank"))
        .expect_err("slide insertion requires the presentation root");
    assert_eq!(bad_parent.code(), Code::UnsupportedType);
    assert_eq!(source.to_bytes().unwrap(), before);
}

#[test]
fn source_add_uses_the_same_layout_name_after_reopening_a_revision() {
    let source = open_source(fixture());
    let first = source.add("/", slide("Blank")).unwrap();
    let bytes = first.file.to_bytes().unwrap();
    let reopened = open_source(bytes);
    let second = reopened
        .add("/", slide("Blank").at(Placement::End))
        .expect("the registered layout remains resolvable after reopening");
    assert_eq!(second.path.to_string(), "/slide[5]");
    assert_eq!(text(&second.file, 1), "A");
}

#[test]
fn source_add_keeps_format_and_source_diagnostics() {
    let source = open_source(fixture());
    let error = source
        .add("/", NewNode::new("slide"))
        .expect_err("layout is required");
    assert_eq!(error.code(), Code::MissingProperty);
    assert_eq!(error.format(), Some(FileFormat::Pptx));
    assert!(error.property().is_some());
}

fn slide_property(file: &File, path: &str, key: &str) -> Option<String> {
    file.get(path)
        .unwrap()
        .view()
        .format()
        .unwrap()
        .find(|property| property.key() == key)
        .map(|property| property.value().to_string())
}

#[test]
fn slide_creation_properties_round_trip_and_replay_for_both_owners() {
    for mode in [SourceMode::Auto, SourceMode::Eager] {
        let before = empty_fixture();
        let source =
            litchi_ops::open_bytes_with(before.clone(), &OpenOptions::new().source_mode(mode))
                .unwrap();
        let name = "Roadmap & Q4\nDraft";
        let commit = source
            .edit(|edit| {
                edit.add("/", slide("Blank").prop("NaMe", name).prop("HIDDEN", true))?;
                Ok(())
            })
            .unwrap();
        for file in [
            commit.file(),
            &open_source(commit.file().to_bytes().unwrap()),
        ] {
            assert_eq!(
                slide_property(file, "/slide[1]", "name").as_deref(),
                Some(name)
            );
            assert_eq!(
                slide_property(file, "/slide[1]", "hidden").as_deref(),
                Some("true")
            );
        }
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
        assert_eq!(source.to_bytes().unwrap(), before);
        for key in ["name", "hidden"] {
            assert!(
                source
                    .capabilities()
                    .element("slide")
                    .unwrap()
                    .property(key)
                    .unwrap()
                    .add()
            );
        }
        let blank = source
            .add(
                "/",
                slide("Blank")
                    .prop("name", litchi_ops::PropValue::Null)
                    .prop("hidden", false),
            )
            .unwrap();
        assert_eq!(
            slide_property(&blank.file, "/slide[1]", "name").as_deref(),
            Some("")
        );
        assert!(matches!(
            slide_property(&blank.file, "/slide[1]", "hidden").as_deref(),
            None | Some("false")
        ));
    }
}

#[test]
fn slide_creation_properties_fail_atomically_before_publication() {
    for mode in [SourceMode::Auto, SourceMode::Eager] {
        let source =
            litchi_ops::open_bytes_with(fixture(), &OpenOptions::new().source_mode(mode)).unwrap();
        let before = source.to_bytes().unwrap();
        for (node, code) in [
            (
                slide("Blank").prop("name", "Draft").prop("hidden", "maybe"),
                Code::InvalidValue,
            ),
            (slide("Blank").prop("name", "bad\u{1}"), Code::InvalidValue),
            (
                slide("Blank").prop("name", "A").prop("NAME", "B"),
                Code::InvalidArgument,
            ),
            (
                slide("Blank").prop("hidden", true).prop("HIDDEN", false),
                Code::InvalidArgument,
            ),
            (
                NewNode::new("slide").prop("name", "Missing layout"),
                Code::MissingProperty,
            ),
        ] {
            assert_eq!(source.add("/", node).unwrap_err().code(), code);
            assert_eq!(source.to_bytes().unwrap(), before);
        }
        let limits = SemanticLimits::builder()
            .max_staged_bytes(8)
            .unwrap()
            .build()
            .unwrap();
        let error = source
            .edit_with_options(
                litchi_ops::EditOptions::new().semantic_limits(limits),
                |edit| {
                    edit.add(
                        "/",
                        slide("Blank").prop("name", "Draft").prop("hidden", true),
                    )
                },
            )
            .unwrap_err();
        assert_eq!(error.code(), Code::LimitExceeded);
        assert_eq!(source.to_bytes().unwrap(), before);
    }
}

#[test]
fn batches_preserve_creation_properties_and_isolate_invalid_items() {
    for mode in [SourceMode::Auto, SourceMode::Eager] {
        let source =
            litchi_ops::open_bytes_with(fixture(), &OpenOptions::new().source_mode(mode)).unwrap();
        let before = source.to_bytes().unwrap();
        let items = || {
            let mut valid = add_item("Blank");
            valid.props = Some(
                Props::new()
                    .with("layout", "Blank")
                    .with("name", "Batch Draft")
                    .with("hidden", true),
            );
            let mut invalid = add_item("Blank");
            invalid.props = Some(Props::new().with("layout", "Blank").with("hidden", "maybe"));
            [valid, invalid]
        };
        assert!(source.batch(items(), BatchMode::Atomic).is_err());
        assert_eq!(source.to_bytes().unwrap(), before);
        let result = source.batch(items(), BatchMode::BestEffort).unwrap();
        assert!(matches!(&result.value()[0], ItemResult::Applied(applied) if applied.changed));
        assert!(
            matches!(&result.value()[1], ItemResult::Failed(error) if error.code() == Code::InvalidValue)
        );
        assert_eq!(
            slide_property(result.file(), "/slide[4]", "name").as_deref(),
            Some("Batch Draft")
        );
        assert_eq!(
            slide_property(result.file(), "/slide[4]", "hidden").as_deref(),
            Some("true")
        );
        assert!(result.file().get("/slide[5]").is_err());
        assert_eq!(source.to_bytes().unwrap(), before);
    }
}
