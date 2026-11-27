//! The unified creation-to-publication loop uses native text-box authoring.
#![cfg(feature = "pptx")]
use litchi_core::{FileFormat, Position};
use litchi_ops::node::Value;
use litchi_ops::transaction::SemanticLimits;
use litchi_ops::{Code, File, NewNode, OpenOptions, Placement, SourceMode, props};

fn shape() -> NewNode<'static> {
    NewNode::new("shape")
        .prop("width", "2in")
        .prop("height", "1in")
}
fn slide() -> File {
    litchi_ops::create(FileFormat::Pptx)
        .unwrap()
        .add("/", NewNode::new("slide").prop("layout", "Blank"))
        .unwrap()
        .file
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
fn length(file: &File, path: &str, key: &str) -> i64 {
    let node = file.get(path).unwrap();
    let prop = node
        .view()
        .format()
        .unwrap()
        .find(|prop| prop.key() == key)
        .unwrap();
    let Value::Length(value) = prop.value() else {
        panic!("length property")
    };
    value.emu()
}
#[test]
fn create_populate_edit_save_and_reopen_without_native_escape_hatches() {
    let source = litchi_ops::create(FileFormat::Pptx).unwrap();
    let before = source.to_bytes().unwrap();
    let commit = source
        .edit(|edit| {
            let slide = edit.add("/", NewNode::new("slide").prop("layout", "Blank"))?;
            let first = edit.add(
                &slide,
                shape()
                    .prop("text", "<&> first")
                    .prop("x", "-1in")
                    .prop("y", "36pt"),
            )?;
            assert_eq!(first.to_string(), "/slide[1]/shape[1]");
            edit.set(&first, props! { "text" => "edited & <escaped>" })?;
            let second = edit.add(
                &slide,
                NewNode::new("textbox")
                    .prop("width", "2in")
                    .prop("height", "1in"),
            )?;
            assert_eq!(second.to_string(), "/slide[1]/shape[2]");
            Ok(())
        })
        .unwrap();
    let file = commit.file();
    assert_eq!(text(file, "/slide[1]/shape[1]"), "edited & <escaped>");
    assert_eq!(text(file, "/slide[1]/shape[2]"), "");
    for (key, value) in [
        ("x", -914_400),
        ("y", 457_200),
        ("width", 1_828_800),
        ("height", 914_400),
    ] {
        assert_eq!(length(file, "/slide[1]/shape[1]", key), value);
    }
    assert_eq!(length(file, "/slide[1]/shape[2]", "x"), 0);
    assert_eq!(length(file, "/slide[1]/shape[2]", "y"), 0);
    assert_eq!(source.to_bytes().unwrap(), before);
    assert_eq!(
        commit.patch().apply(&source).unwrap().to_bytes().unwrap(),
        file.to_bytes().unwrap()
    );
    assert_eq!(
        commit
            .patch()
            .inverse()
            .apply(file)
            .unwrap()
            .to_bytes()
            .unwrap(),
        before
    );
    let dir = tempfile::tempdir().unwrap();
    let path = dir.path().join("authored.pptx");
    file.save_as(&path).unwrap();
    let reopened = litchi_ops::open(&path).unwrap();
    assert_eq!(text(&reopened, "/slide[1]/shape[1]"), "edited & <escaped>");
}
#[test]
fn shape_paths_count_only_direct_shape_siblings() {
    let source = litchi_pptx::Package::new().unwrap();
    let mut edit = source.edit().unwrap();
    edit.insert_slide(Position::new(0), "Blank").unwrap();
    edit.add_table(
        Position::new(0),
        &[vec!["table\r\ncell".into()]],
        (0, 0, 914400, 914400),
    )
    .unwrap();
    edit.add_rectangle(Position::new(0), (0, 0, 914400, 914400), None)
        .unwrap();
    let bytes = edit.commit().unwrap().into_package().to_bytes().unwrap();
    let source =
        litchi_ops::open_bytes_with(bytes, &OpenOptions::new().source_mode(SourceMode::Eager))
            .unwrap();
    assert_eq!(
        text(&source, "/slide[1]/table[1]/tr[1]/tc[1]"),
        "table\r\ncell"
    );
    let before = source.to_bytes().unwrap();
    let added = source
        .add(
            "/slide[1]",
            shape().prop("text", "after table and rectangle"),
        )
        .unwrap();
    assert_eq!(added.path.to_string(), "/slide[1]/shape[2]");
    assert_eq!(
        text(&added.file, "/slide[1]/shape[2]"),
        "after table and rectangle"
    );
    assert!(added.file.get("/slide[1]/table[1]").is_ok());
    let changed = added
        .file
        .set(&added.path, props! { "text" => "selected new shape" })
        .unwrap();
    assert_eq!(text(&changed, "/slide[1]/shape[2]"), "selected new shape");
    assert_eq!(source.to_bytes().unwrap(), before);
}
#[test]
fn shape_addition_has_strict_properties_placement_and_atomic_limits() {
    let source = slide();
    let before = source.to_bytes().unwrap();
    for (node, code) in [
        (NewNode::new("shape"), Code::MissingProperty),
        (
            NewNode::new("shape").prop("width", 0).prop("height", "1in"),
            Code::InvalidValue,
        ),
        (
            NewNode::new("shape")
                .prop("width", "-1in")
                .prop("height", "1in"),
            Code::InvalidValue,
        ),
        (shape().prop("text", "bad\u{0}"), Code::InvalidValue),
        (shape().prop("name", "ignored?"), Code::UnsupportedProperty),
        (shape().prop("opacity", "0.5"), Code::UnsupportedProperty),
        (shape().prop("x", 1).prop("x", 2), Code::InvalidArgument),
        (shape().at(Placement::Start), Code::UnsupportedCapability),
    ] {
        assert_eq!(source.add("/slide[1]", node).unwrap_err().code(), code);
        assert_eq!(source.to_bytes().unwrap(), before);
    }
    assert_eq!(
        source.add("/", shape()).unwrap_err().code(),
        Code::UnsupportedType
    );
    let limits = SemanticLimits::builder()
        .max_staged_bytes(35)
        .unwrap()
        .build()
        .unwrap();
    assert!(
        source
            .edit_with_options(
                litchi_ops::EditOptions::new().semantic_limits(limits),
                |edit| edit.add("/slide[1]", shape().prop("text", "abc"))
            )
            .is_ok()
    );
    let error = source
        .edit_with_options(
            litchi_ops::EditOptions::new().semantic_limits(limits),
            |edit| edit.add("/slide[1]", shape().prop("text", "abcd")),
        )
        .unwrap_err();
    assert_eq!(error.code(), Code::LimitExceeded);
    assert_eq!(source.to_bytes().unwrap(), before);
    let deferred = litchi_ops::open_bytes(before.clone()).unwrap();
    let source_added = deferred
        .add("/slide[1]", shape())
        .expect("the source-backed owner admits top-level text-box insertion");
    assert_eq!(source_added.path.to_string(), "/slide[1]/shape[1]");
    assert_eq!(text(&source_added.file, "/slide[1]/shape[1]"), "");
    let added = source.add("/slide[1]", shape()).unwrap();
    let resized = added
        .file
        .set(&added.path, props! { "width" => "3in" })
        .unwrap();
    assert_eq!(length(&resized, "/slide[1]/shape[1]", "width"), 2_743_200);
}

#[test]
fn source_shape_add_is_reversible_and_keeps_source_adoption() {
    let eager = slide();
    let source = litchi_ops::open_bytes_with(
        eager.to_bytes().unwrap(),
        &OpenOptions::new().source_mode(SourceMode::Auto),
    )
    .unwrap();
    let before = source.to_bytes().unwrap();
    let added = source
        .add("/slide[1]", shape().prop("text", "source & <text>"))
        .expect("source-backed text-box insertion");
    assert_eq!(added.path.to_string(), "/slide[1]/shape[1]");
    assert_eq!(
        text(&added.file, &added.path.to_string()),
        "source & <text>"
    );
    assert_eq!(source.to_bytes().unwrap(), before);
    let commit = source
        .edit(|edit| edit.add("/slide[1]", shape().prop("text", "transaction shape")))
        .expect("source-backed managed shape insertion");
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
    let updated = added
        .file
        .set(&added.path, props! { "text" => "edited source shape" })
        .unwrap();
    assert_eq!(text(&updated, "/slide[1]/shape[1]"), "edited source shape");
    let reopened = litchi_ops::open_bytes(updated.to_bytes().unwrap()).unwrap();
    assert_eq!(text(&reopened, "/slide[1]/shape[1]"), "edited source shape");
}

#[test]
fn shape_batches_preserve_atomic_and_best_effort_boundaries() {
    use litchi_ops::{BatchItem, BatchMode, ItemResult, Props};
    let source = slide();
    let before = source.to_bytes().unwrap();
    let commands = || {
        [false, true, false].into_iter().map(|invalid| {
            let mut item = BatchItem::new("add");
            item.parent = Some("/slide[1]".into());
            item.kind = Some("shape".into());
            let props = Props::new()
                .with("text", "batch text")
                .with("width", "1in")
                .with("height", "1in");
            item.props = Some(if invalid {
                props.with("opacity", "0.5")
            } else {
                props
            });
            item
        })
    };
    assert_eq!(
        source
            .batch(commands(), BatchMode::Atomic)
            .unwrap_err()
            .item(),
        Some(1)
    );
    let result = source.batch(commands(), BatchMode::BestEffort).unwrap();
    assert!(matches!(&result.value()[1], ItemResult::Failed(error) if error.item() == Some(1)));
    assert_eq!(text(result.file(), "/slide[1]/shape[2]"), "batch text");
    assert!(result.file().get("/slide[1]/shape[3]").is_err());
    assert_eq!(source.to_bytes().unwrap(), before);
}

#[test]
fn authored_text_preserves_escaping_whitespace_and_unicode() {
    let source = slide();
    for value in [
        "<&> \"quotes\" 'apostrophe'",
        " leading\tand\ntrailing ",
        "line\rbreak\r\nend",
        "你好 🙂",
    ] {
        let added = source
            .add("/slide[1]", shape().prop("text", value))
            .unwrap();
        assert_eq!(text(&added.file, "/slide[1]/shape[1]"), value);
        let reopened = litchi_ops::open_bytes(added.file.to_bytes().unwrap()).unwrap();
        assert_eq!(text(&reopened, "/slide[1]/shape[1]"), value);
        for mode in [SourceMode::Auto, SourceMode::Eager] {
            let original = litchi_ops::open_bytes_with(
                added.file.to_bytes().unwrap(),
                &OpenOptions::new().source_mode(mode),
            )
            .unwrap();
            let replacement = format!("updated\r{value}");
            let commit = original
                .edit(|edit| {
                    edit.set(
                        "/slide[1]/shape[1]",
                        props! { "text" => replacement.as_str() },
                    )
                })
                .unwrap();
            assert_eq!(text(commit.file(), "/slide[1]/shape[1]"), replacement);
            assert_eq!(
                commit
                    .patch()
                    .inverse()
                    .apply(commit.file())
                    .unwrap()
                    .to_bytes()
                    .unwrap(),
                original.to_bytes().unwrap()
            );
        }
    }
}
