#![cfg(feature = "rtf")]

use litchi_core::Position;
use litchi_ops::error::Source;
use litchi_ops::transaction::{Error, LimitKind, SemanticLimits};
use litchi_ops::{BatchItem, BatchMode, Code, File, NewNode, Placement, open_bytes, props};
use litchi_rtf::edit::RunFormatPatch;

const INPUT: &[u8] = br"{\rtf1\ansi First{\i Italic}{\b Bold}\par\par Tail}";

fn text(file: &File, path: &str) -> String {
    file.get(path)
        .unwrap()
        .view()
        .text()
        .unwrap()
        .unwrap()
        .into_owned()
}

fn paragraphs(file: &File) -> Vec<String> {
    file.query("p")
        .unwrap()
        .map(|node| node.unwrap().view().text().unwrap().unwrap().into_owned())
        .collect()
}

fn property(file: &File, path: &str, key: &str) -> String {
    file.get(path)
        .unwrap()
        .view()
        .format()
        .unwrap()
        .find(|value| value.key() == key)
        .unwrap()
        .to_string()
}

fn native_run_add(
    input: &[u8],
    paragraph: usize,
    run: usize,
    text: &str,
    format: RunFormatPatch,
) -> Vec<u8> {
    let native = litchi_rtf::Package::from_bytes(input).unwrap();
    let mut edit = native.edit();
    edit.insert_run(Position::new(paragraph), Position::new(run), text, format)
        .unwrap();
    edit.commit().unwrap().into_package().to_bytes().unwrap()
}

fn run_node<'a>(text: &'a str) -> NewNode<'a> {
    NewNode::new("r").prop("text", text)
}

#[test]
fn insertion_placements_match_native_bytes_and_report_created_paths() {
    for (placement, paragraph, run, expected) in [
        (Placement::Start, 0, 0, "addedFirstItalicBold"),
        (Placement::End, 0, 3, "FirstItalicBoldadded"),
        (
            Placement::Position(Position::new(1)),
            0,
            1,
            "FirstaddedItalicBold",
        ),
        (
            Placement::before("/body/p[1]/r[2]").unwrap(),
            0,
            1,
            "FirstaddedItalicBold",
        ),
        (
            Placement::after("/body/p[1]/r[2]").unwrap(),
            0,
            2,
            "FirstItalicaddedBold",
        ),
        (
            Placement::inside_at("/body/p[1]", Position::new(0)).unwrap(),
            0,
            0,
            "addedFirstItalicBold",
        ),
    ] {
        let source = open_bytes(INPUT).unwrap();
        let result = source
            .add("/body/p[1]", run_node("added").at(placement))
            .unwrap();
        assert_eq!(
            result.file.to_bytes().unwrap(),
            native_run_add(INPUT, paragraph, run, "added", RunFormatPatch::new())
        );
        assert_eq!(
            result.path.to_string(),
            format!("/body/p[1]/r[{}]", run + 1)
        );
        assert_eq!(text(&result.file, "/body/p[1]"), expected);
        assert_eq!(text(&result.file, &result.path.to_string()), "added");
        assert_eq!(source.to_bytes().unwrap(), INPUT);
    }
}

#[test]
fn empty_paragraphs_are_real_runs_and_unicode_source_fragments_keep_paths_valid() {
    let input = br"{\rtf1\ansi First\par\par Tail}";
    let source = open_bytes(input).unwrap();
    let added = source
        .add("/body/p[2]", run_node("middle").at(Placement::End))
        .unwrap();
    assert_eq!(added.path.to_string(), "/body/p[2]/r[1]");
    assert_eq!(paragraphs(&added.file), ["First", "middle", "Tail"]);
    assert_eq!(text(&added.file, "/body/p[2]/r[1]"), "middle");
    assert_eq!(
        added.file.to_bytes().unwrap(),
        native_run_add(input, 1, 0, "middle", RunFormatPatch::new())
    );

    let unicode = br"{\rtf1\ansi\uc1 A\u233?B{\i C\u38634?D}\tab end}";
    let native = litchi_rtf::Package::from_bytes(unicode).unwrap();
    let native_runs = native
        .body()
        .paragraphs()
        .next()
        .unwrap()
        .runs()
        .map(|run| run.text())
        .collect::<Vec<_>>();
    assert_eq!(
        native_runs,
        vec!["A", "é", "B", "C", "雪", "D", "\tend"],
        "the insertion boundary must follow the native source run axis"
    );
    let before_b = native_runs
        .iter()
        .position(|run| *run == "B")
        .expect("the fixture must expose B as its own native run");
    let before_b_text = native_runs
        .iter()
        .take(before_b)
        .copied()
        .collect::<String>();
    let after_b_text = native_runs
        .iter()
        .skip(before_b)
        .copied()
        .collect::<String>();
    let source = open_bytes(unicode).unwrap();
    let added = source
        .add(
            "/body/p[1]",
            run_node("inserted").at(Placement::Position(Position::new(before_b))),
        )
        .unwrap();
    assert_eq!(
        added.path.to_string(),
        format!("/body/p[1]/r[{}]", before_b + 1)
    );
    assert_eq!(text(&added.file, &added.path.to_string()), "inserted");
    assert_eq!(
        text(&added.file, "/body/p[1]"),
        format!("{before_b_text}inserted{after_b_text}")
    );
    assert_eq!(
        added.file.to_bytes().unwrap(),
        native_run_add(unicode, 0, before_b, "inserted", RunFormatPatch::new())
    );
    assert_eq!(
        open_bytes(added.file.to_bytes().unwrap())
            .unwrap()
            .get(&added.path)
            .unwrap()
            .view()
            .text()
            .unwrap()
            .as_deref(),
        Some("inserted")
    );
}

#[test]
fn formatting_defaults_inherit_and_compound_properties_match_native_patch() {
    let input = br"{\rtf1\ansi\fs28\b Prefix{\i Italic}Suffix}";
    let source = open_bytes(input).unwrap();
    let inherited = source
        .add(
            "/body/p[1]",
            run_node("inherited").at(Placement::Position(Position::new(1))),
        )
        .unwrap();
    assert_eq!(
        text(&inherited.file, "/body/p[1]"),
        "PrefixinheritedItalicSuffix"
    );
    assert_eq!(
        property(&inherited.file, &inherited.path.to_string(), "size"),
        "14pt"
    );
    assert_eq!(
        property(&inherited.file, &inherited.path.to_string(), "bold"),
        "true"
    );
    assert_eq!(
        property(&inherited.file, &inherited.path.to_string(), "italic"),
        "true"
    );
    assert_eq!(
        inherited.file.to_bytes().unwrap(),
        native_run_add(input, 0, 1, "inherited", RunFormatPatch::new())
    );

    let compound_text = "new \\ {text}\té雪😀\tCR:\rVT:\u{b}NUL:\0";
    let format = RunFormatPatch::new()
        .with_bold(true)
        .with_italic(true)
        .with_underline(litchi_rtf::text::Underline::Double)
        .with_size(std::num::NonZeroU16::new(25).unwrap())
        .with_strike(true)
        .with_double_strike(true)
        .with_hidden(true)
        .with_all_caps(true)
        .with_small_caps(true)
        .with_outline(true);
    let changed = source
        .add(
            "/body/p[1]",
            NewNode::new("run")
                .prop("text", compound_text)
                .prop("bold", true)
                .prop("italic", true)
                .prop("underline", "double")
                .prop("size", "12.5pt")
                .prop("strike", true)
                .prop("dstrike", true)
                .prop("vanish", true)
                .prop("caps", true)
                .prop("smallcaps", true)
                .prop("outline", true)
                .at(Placement::End),
        )
        .unwrap();
    assert_eq!(
        changed.file.to_bytes().unwrap(),
        native_run_add(input, 0, 3, compound_text, format)
    );
    for (key, value) in [
        ("bold", "true"),
        ("italic", "true"),
        ("underline", "double"),
        ("size", "12.5pt"),
        ("strike", "true"),
        ("dstrike", "true"),
        ("vanish", "true"),
        ("caps", "true"),
        ("smallcaps", "true"),
        ("outline", "true"),
    ] {
        assert_eq!(
            property(&changed.file, &changed.path.to_string(), key),
            value,
            "{key}"
        );
    }
    assert_eq!(
        text(&changed.file, "/body/p[1]"),
        format!("PrefixItalicSuffix{compound_text}")
    );
    assert_eq!(
        text(&changed.file, &changed.path.to_string()),
        compound_text,
        "the created path must resolve to the complete inserted run"
    );
}

#[test]
fn opaque_source_is_byte_exact_and_patch_inverse_reopens() {
    let input = br"{\rtf1\ansi{\fonttbl{\f0\fnil Example;}}{\info{\title Keep title}}{\header Keep header}{\*\vendor before}Before{\i Italic}{\*\vendor after}Tail}";
    let source = open_bytes(input).unwrap();
    let result = source
        .edit(|edit| {
            edit.add(
                "/body/p[1]",
                NewNode::new("r")
                    .prop("text", "inserted")
                    .prop("bold", true)
                    .at(Placement::after("/body/p[1]/r[1]").unwrap()),
            )
        })
        .unwrap();
    let expected = native_run_add(
        input,
        0,
        1,
        "inserted",
        RunFormatPatch::new().with_bold(true),
    );
    let bytes = result.file().to_bytes().unwrap();
    assert_eq!(bytes, expected);
    assert!(
        bytes
            .windows(b"{\\*\\vendor before}".len())
            .any(|part| part == b"{\\*\\vendor before}")
    );
    assert!(
        bytes
            .windows(b"{\\*\\vendor after}".len())
            .any(|part| part == b"{\\*\\vendor after}")
    );
    assert!(
        bytes
            .windows(b"Keep title".len())
            .any(|part| part == b"Keep title")
    );
    assert!(
        bytes
            .windows(b"Keep header".len())
            .any(|part| part == b"Keep header")
    );
    let reopened = open_bytes(bytes.clone()).unwrap();
    assert_eq!(text(&reopened, "/body/p[1]"), "BeforeinsertedItalicTail");
    assert_eq!(
        property(&reopened, &result.value().to_string(), "bold"),
        "true"
    );
    assert_eq!(
        result
            .patch()
            .inverse()
            .apply(result.file())
            .unwrap()
            .to_bytes()
            .unwrap(),
        input
    );
    assert_eq!(source.to_bytes().unwrap(), input);
}

#[test]
fn batch_and_sequential_adds_keep_paths_and_atomic_history() {
    let source = open_bytes(INPUT).unwrap();
    let mut item = BatchItem::new("add");
    item.parent = Some("/body/p[1]".into());
    item.kind = Some("r".into());
    item.props = Some(litchi_ops::Props::from_props(props! {"text" => "new"}).unwrap());
    let mut bad = item.clone();
    bad.props = Some(litchi_ops::Props::from_props(props! {"text" => true}).unwrap());
    assert_eq!(
        source
            .batch([item.clone(), bad], BatchMode::Atomic)
            .unwrap_err()
            .item(),
        Some(1)
    );
    assert_eq!(source.to_bytes().unwrap(), INPUT);
    let best = source.batch([item], BatchMode::BestEffort).unwrap();
    assert_eq!(text(best.file(), "/body/p[1]"), "FirstItalicBoldnew");
    assert!(matches!(
        best.value()[0],
        litchi_ops::ItemResult::Applied(_)
    ));

    let commit = source
        .edit(|edit| {
            let first = edit.add("/body/p[1]", run_node("one").at(Placement::Start))?;
            edit.set(&first, props! {"bold" => true})?;
            edit.add("/body/p[1]", run_node("two").at(Placement::End))?;
            Ok(())
        })
        .unwrap();
    assert_eq!(text(commit.file(), "/body/p[1]"), "oneFirstItalicBoldtwo");
    assert_eq!(property(commit.file(), "/body/p[1]/r[1]", "bold"), "true");
    assert_eq!(commit.diagnostics().operations(), 3);
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

#[test]
fn invalid_run_adds_and_limits_are_typed_and_atomic() {
    let source = open_bytes(INPUT).unwrap();
    let before = source.to_bytes().unwrap();
    for node in [
        NewNode::new("p").prop("text", "wrong"),
        NewNode::new("r"),
        NewNode::new("r").prop("text", 7),
        NewNode::new("r").prop("text", "one\ntwo"),
        NewNode::new("r").prop("bold", "maybe"),
    ] {
        assert!(source.add("/body/p[1]", node).is_err());
        assert_eq!(source.to_bytes().unwrap(), before);
    }
    assert_eq!(
        source
            .add("/body/p[99]", run_node("bad"))
            .unwrap_err()
            .code(),
        Code::PathNotFound
    );
    assert_eq!(
        source
            .add(
                "/body/p[1]",
                run_node("bad").at(Placement::Position(Position::new(99)))
            )
            .unwrap_err()
            .code(),
        Code::OutOfRange
    );

    let limits = SemanticLimits::builder()
        .max_operations(1)
        .unwrap()
        .build()
        .unwrap();
    let error = source
        .edit_with_options(
            litchi_ops::EditOptions::new().semantic_limits(limits),
            |edit| {
                edit.add("/body/p[1]", run_node("one"))?;
                edit.add("/body/p[1]", run_node("two"))
            },
        )
        .unwrap_err();
    assert!(
        matches!(
            error.leaf(),
            Source::Transaction(Error::Limit {
                kind: LimitKind::Operations,
                ..
            })
        ),
        "{error:?}"
    );
    assert_eq!(source.to_bytes().unwrap(), before);

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
                |edit| edit.add("/body/p[1]", run_node("added")),
            )
            .unwrap_err();
        assert!(
            matches!(
                error.leaf(),
                Source::Transaction(Error::Limit { kind: actual, .. }) if *actual == kind
            ),
            "{error:?}"
        );
        assert_eq!(source.to_bytes().unwrap(), before);
    }
}

#[test]
fn run_add_preserves_protection_and_source_origin_for_atomic_save() {
    let protected = open_bytes(br"{\rtf1\ansi\readprot\enforceprot1 Text}").unwrap();
    assert_eq!(
        protected
            .add("/body/p[1]", run_node("new"))
            .unwrap_err()
            .code(),
        Code::DocumentProtected
    );

    let dir = tempfile::tempdir().unwrap();
    let path = dir.path().join("run-add.rtf");
    std::fs::write(&path, INPUT).unwrap();
    let source = litchi_ops::open(&path).unwrap();
    source
        .add("/body/p[1]", run_node("saved").at(Placement::End))
        .unwrap()
        .file
        .save()
        .unwrap();
    let reopened = litchi_ops::open(&path).unwrap();
    assert_eq!(text(&reopened, "/body/p[1]"), "FirstItalicBoldsaved");
    assert_eq!(std::fs::read_dir(dir.path()).unwrap().count(), 1);
}

#[test]
fn run_capability_admits_add_and_reports_remaining_structural_gaps() {
    let source = open_bytes(INPUT).unwrap();
    let run = source.capabilities().element("r").unwrap();
    assert!(run.operations().get && run.operations().query);
    assert!(run.operations().add);
    assert!(run.operations().remove);
    for name in [
        "text",
        "size",
        "bold",
        "italic",
        "underline",
        "strike",
        "dstrike",
        "vanish",
        "caps",
        "smallcaps",
        "outline",
    ] {
        assert!(run.property(name).unwrap().add(), "{name}");
    }
}
