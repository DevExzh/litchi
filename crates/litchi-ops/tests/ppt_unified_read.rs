//! Public unified-read coverage for the legacy PowerPoint adapter.
//!
//! The fixture is deliberately authored through the native PPT builder.  The
//! assertions then stay on `File` and its generic node surface, so the test
//! covers both the eager and source-backed owners without duplicating CFB
//! records in the facade test suite.

#![cfg(feature = "ppt")]

use std::io::Cursor;

use litchi_ops::node::{Node, ViewOptions};
use litchi_ops::view::Mode;
use litchi_ops::{Code, File, OpenOptions, SourceMode};
use litchi_ppt::Shape as _;
use litchi_ppt::writer::Builder;

fn presentation_bytes() -> Vec<u8> {
    let mut builder = Builder::new();
    let first = builder.add_slide().expect("first slide");
    builder
        .add_textbox(first, 10, 10, 240, 40, "Quarterly Review")
        .expect("first title");
    builder
        .add_textbox(first, 10, 70, 240, 40, "Revenue grew 25%")
        .expect("first body");

    let second = builder.add_slide().expect("second slide");
    builder
        .add_textbox(second, 10, 10, 240, 40, "Risks")
        .expect("second title");
    builder
        .add_textbox(second, 10, 70, 240, 40, "Follow-up detail")
        .expect("second body");

    builder
        .build()
        .expect("build PPT")
        .to_bytes()
        .expect("encode PPT")
}

fn signed_presentation_bytes() -> Vec<u8> {
    let mut ole = litchi_cfb::OleFile::open(
        Cursor::new(presentation_bytes()),
        litchi_cfb::OleFileLimits::default(),
    )
    .expect("open native PPT package");
    let document = ole
        .open_stream(&["PowerPoint Document"])
        .expect("PowerPoint Document stream");
    let current_user = ole
        .open_stream(&["Current User"])
        .expect("Current User stream");

    let mut writer = litchi_cfb::OleWriter::new();
    writer
        .create_stream(&["PowerPoint Document"], &document)
        .expect("copy document stream");
    writer
        .create_stream(&["Current User"], &current_user)
        .expect("copy current-user stream");
    writer
        .create_stream(&["_signatures"], b"opaque-signature")
        .expect("add signature stream");

    let mut output = Cursor::new(Vec::new());
    writer
        .write_to(&mut output)
        .expect("write signed PPT package");
    output.into_inner()
}

fn eager(bytes: &[u8]) -> File {
    litchi_ops::open_bytes_with(
        bytes.to_vec(),
        &OpenOptions::new().source_mode(SourceMode::Eager),
    )
    .expect("open eager PPT")
}

fn source_file(bytes: &[u8]) -> (tempfile::NamedTempFile, File) {
    let temporary = tempfile::NamedTempFile::new().expect("temporary PPT");
    std::fs::write(temporary.path(), bytes).expect("write temporary PPT");
    let file = litchi_ops::open(temporary.path()).expect("open source-backed PPT");
    (temporary, file)
}

fn child_paths(node: &Node<'_>) -> Vec<String> {
    node.children()
        .map(|child| child.expect("child node").positional_path().to_string())
        .collect()
}

fn query_paths(file: &File, selector: &str) -> Vec<String> {
    file.query(selector)
        .unwrap_or_else(|error| panic!("{selector}: {error}"))
        .map(|node| {
            node.unwrap_or_else(|error| panic!("{selector}: {error}"))
                .positional_path()
                .to_string()
        })
        .collect()
}

fn text(node: &Node<'_>) -> Option<String> {
    node.view()
        .text()
        .expect("node text")
        .map(|text| text.into_owned())
}

fn assert_code(error: &litchi_ops::Error, code: Code) {
    assert_eq!(error.code(), code, "{error}");
}

fn first_shape_id(bytes: &[u8]) -> u32 {
    let package = litchi_ppt::Package::from_bytes(bytes).expect("parse native PPT");
    let presentation = package.presentation().expect("native presentation");
    let slide = presentation
        .slide_at(0)
        .expect("native first slide")
        .expect("first slide exists");
    slide
        .shapes()
        .expect("native first-slide shapes")
        .first()
        .expect("first shape exists")
        .as_textbox()
        .expect("textbox")
        .properties()
        .id
}

fn assert_navigation_and_views(file: &File) {
    assert_eq!(file.format(), Some(litchi_core::FileFormat::Ppt));
    assert_eq!(
        file.root().expect("presentation root").kind(),
        "presentation"
    );

    let root = file.get("/").expect("root");
    assert_eq!(root.positional_path().to_string(), "/");
    assert_eq!(child_paths(&root), ["/slide[1]", "/slide[2]"]);

    let first = file.get("/slide[1]").expect("first slide");
    assert_eq!(first.kind(), "slide");
    assert_eq!(first.positional_path().to_string(), "/slide[1]");
    assert_eq!(
        child_paths(&first),
        ["/slide[1]/shape[1]", "/slide[1]/shape[2]"]
    );

    let title = file.get("/slide[1]/shape[1]").expect("first title shape");
    assert_eq!(title.kind(), "shape");
    assert_eq!(title.positional_path().to_string(), "/slide[1]/shape[1]");
    assert_eq!(text(&title).as_deref(), Some("Quarterly Review"));

    let body = file
        .find("/slide[1]/shape[2]")
        .expect("find body shape")
        .expect("body shape exists");
    assert_eq!(text(&body).as_deref(), Some("Revenue grew 25%"));
    assert!(
        file.find("/slide[9]")
            .expect("find missing slide")
            .is_none()
    );
    assert!(
        file.find("/slide[1]/shape[9]")
            .expect("find missing shape")
            .is_none()
    );

    assert_eq!(query_paths(file, "slide"), ["/slide[1]", "/slide[2]"]);
    assert_eq!(
        query_paths(file, "shape"),
        [
            "/slide[1]/shape[1]",
            "/slide[1]/shape[2]",
            "/slide[2]/shape[1]",
            "/slide[2]/shape[2]",
        ]
    );

    let node_json = serde_json::to_value(title.view_with(ViewOptions::default().depth(1)))
        .expect("serialize shape node view");
    assert_eq!(node_json["type"], "shape");
    assert_eq!(node_json["path"], "/slide[1]/shape[1]");
    assert!(node_json.to_string().contains("Quarterly Review"));

    for mode in [Mode::Text, Mode::Annotated, Mode::Outline, Mode::Stats] {
        let rendered = file
            .view(mode)
            .unwrap_or_else(|error| panic!("{mode:?}: {error}"));
        assert!(!rendered.is_empty(), "empty {mode:?} view");
        let json = serde_json::to_value(
            file.view_json(mode)
                .unwrap_or_else(|error| panic!("{mode:?} JSON: {error}")),
        )
        .unwrap_or_else(|error| panic!("serialize {mode:?} JSON: {error}"));
        assert!(!json.is_null(), "null {mode:?} JSON view");
    }

    let text_view = file.view(Mode::Text).expect("text view");
    assert!(text_view.contains("Quarterly Review"));
    assert!(text_view.contains("Follow-up detail"));
}

#[test]
fn navigation_query_views_and_owner_parity_are_public() {
    let bytes = presentation_bytes();
    let (temporary, source) = source_file(&bytes);
    let eager = eager(&bytes);
    let _keep_source_alive = temporary;

    assert_navigation_and_views(&source);
    assert_navigation_and_views(&eager);
    assert_eq!(
        source.text().expect("source text"),
        eager.text().expect("eager text")
    );
    assert_eq!(
        source.view(Mode::Text).expect("source text view"),
        eager.view(Mode::Text).expect("eager text view")
    );
}

#[test]
fn native_shape_id_requires_the_explicit_raw_path_escape() {
    let bytes = presentation_bytes();
    let file = eager(&bytes);
    let shape_id = first_shape_id(&bytes);
    let identity_path = format!("/slide[1]/shape[@id={shape_id}]");

    let ordinary = file
        .get(identity_path.as_str())
        .expect_err("native identity must not be ordinary path syntax");
    assert_code(&ordinary, Code::InvalidPath);

    let raw = litchi_ops::path::raw::path(identity_path.as_str()).expect("raw native-id path");
    let shape = file.get(raw).expect("raw shape lookup");
    assert_eq!(shape.positional_path().to_string(), "/slide[1]/shape[1]");
    assert_eq!(text(&shape).as_deref(), Some("Quarterly Review"));
}

#[test]
fn released_ppt_sources_keep_read_navigation_and_views() {
    let file = litchi_ops::open_bytes_with(
        presentation_bytes(),
        &OpenOptions::new().retain_source(false),
    )
    .expect("open released PPT");

    assert_navigation_and_views(&file);
    assert_eq!(
        file.get("/slide[2]/shape[2]")
            .expect("released shape")
            .view()
            .text()
            .expect("released shape text")
            .as_deref(),
        Some("Follow-up detail")
    );
}

#[test]
fn signed_ppt_remains_readable_through_the_unified_surface() {
    let file = litchi_ops::open_bytes(signed_presentation_bytes()).expect("open signed PPT");
    assert!(file.is_signed());
    assert!(matches!(file, File::Signed(_)));

    assert_eq!(
        file.get("/slide[1]/shape[1]")
            .expect("signed shape")
            .view()
            .text()
            .expect("signed shape text")
            .as_deref(),
        Some("Quarterly Review")
    );
    assert_eq!(
        query_paths(&file, "shape"),
        [
            "/slide[1]/shape[1]",
            "/slide[1]/shape[2]",
            "/slide[2]/shape[1]",
            "/slide[2]/shape[2]",
        ]
    );
    assert!(
        file.view(Mode::Text)
            .expect("signed text view")
            .contains("Risks")
    );
}

#[cfg(any(unix, windows))]
#[test]
fn source_nodes_queries_and_views_refuse_a_changed_ppt_source() {
    let bytes = presentation_bytes();
    let (temporary, file) = source_file(&bytes);
    let node = file.get("/slide[1]/shape[1]").expect("source shape");
    let node_view = node.view();
    let mut query = file.query("shape").expect("source query");
    let root = file.root().expect("source root");
    let mut children = root.children();

    std::fs::write(temporary.path(), b"changed PPT source").expect("change source");

    assert_code(
        &node_view.text().expect_err("stale node view"),
        Code::IoError,
    );
    assert_code(
        &query
            .next()
            .expect("stale query reports one item")
            .expect_err("stale query item"),
        Code::IoError,
    );
    assert_code(
        &children
            .next()
            .expect("stale child reports one item")
            .expect_err("stale child item"),
        Code::IoError,
    );
    assert!(query.next().is_none(), "query fuses after a source error");
    assert!(
        children.next().is_none(),
        "children fuse after a source error"
    );
    assert_code(
        &file.view(Mode::Text).expect_err("stale presentation view"),
        Code::IoError,
    );
}

#[cfg(any(unix, windows))]
#[test]
fn released_positional_projection_still_fences_cached_reads() {
    let temporary = tempfile::NamedTempFile::new().expect("PPT file");
    std::fs::write(temporary.path(), presentation_bytes()).expect("fixture");
    let file = litchi_ops::open_with(temporary.path(), &OpenOptions::new().retain_source(false))
        .expect("released projection");
    let node = file.get("/slide[1]/shape[1]").expect("cached shape");
    let root = file.root().expect("root");
    std::fs::write(temporary.path(), b"changed source").expect("change source");
    assert_code(
        &node
            .view()
            .text()
            .expect_err("cached text must fence source"),
        Code::IoError,
    );
    assert_code(
        &root
            .view()
            .child_count()
            .expect_err("cached catalog must fence source"),
        Code::IoError,
    );
    assert_code(
        &file.view(Mode::Stats).expect_err("stats must fence source"),
        Code::IoError,
    );
}

#[test]
fn selectors_paths_and_raw_properties_share_the_read_registry() {
    let file = eager(&presentation_bytes());
    assert_eq!(
        query_paths(&file, "shape:contains(revenue)"),
        ["/slide[1]/shape[2]"]
    );
    assert_eq!(
        query_paths(&file, "slide > shape[2]"),
        ["/slide[1]/shape[2]"]
    );
    assert_eq!(
        file.get("/slide[last()]/shape[last()]")
            .expect("last shape")
            .positional_path()
            .to_string(),
        "/slide[2]/shape[2]"
    );
    assert_code(
        &file.get("/slide[9]").expect_err("missing"),
        Code::PathNotFound,
    );
    assert_code(
        &file.get("/body/p[1]").expect_err("word axis"),
        Code::UnsupportedPath,
    );
    let node = file.get("/slide[1]/shape[1]").expect("shape");
    let json = serde_json::to_value(node.view()).expect("view");
    assert!(
        json["format"].get("id").is_none(),
        "ordinary node views omit native identity"
    );
    assert!(
        !format!("{node:?}").contains("Quarterly Review"),
        "Debug must redact document text"
    );
}

#[test]
fn native_geometry_retains_typed_emu_values_and_query_units() {
    let file = eager(&presentation_bytes());
    let node = file.get("/slide[1]/shape[1]").expect("shape");
    let view = node.view();
    let geometry: Vec<_> = view.format().expect("geometry").collect();
    for (key, emu) in [
        ("x", 127_000),
        ("y", 127_000),
        ("width", 3_048_000),
        ("height", 508_000),
    ] {
        let property = geometry
            .iter()
            .find(|property| property.key() == key)
            .expect("geometry property");
        assert_eq!(
            property.value(),
            &litchi_ops::node::Value::Length(litchi_ops::value::Length::from_emu(emu))
        );
    }
    assert_eq!(
        query_paths(&file, "shape[x=10pt]"),
        [
            "/slide[1]/shape[1]",
            "/slide[1]/shape[2]",
            "/slide[2]/shape[1]",
            "/slide[2]/shape[2]"
        ]
    );
}

#[test]
fn typed_ppt_nodes_are_shareable_cheap_clones() {
    fn assert_traits<T: Send + Sync + Clone>() {}
    assert_traits::<litchi_ops::format::ppt::Node<'static>>();
    let file = eager(&presentation_bytes());
    let node = file.get("/slide[1]/shape[1]").expect("shape");
    let typed = node.as_ppt().expect("PPT node");
    let copy = typed.clone();
    assert!(std::ptr::eq(
        typed.shape().expect("shape"),
        copy.shape().expect("cloned shape")
    ));
}

#[test]
fn placeholder_queries_preserve_physical_shape_paths() {
    let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/poi/test-data/slideshow/basic_test_ppt_file.ppt");
    let bytes = std::fs::read(path).expect("POI fixture");
    let package = litchi_ppt::Package::from_bytes(&bytes).expect("native PPT");
    let presentation = package.presentation().expect("presentation");
    let mut expected = Vec::new();
    for position in 0..presentation.slide_count() {
        let slide = presentation
            .slide_at(position)
            .expect("slide")
            .expect("present");
        let mut shape_ordinal = 0;
        for shape in slide.shapes().expect("shapes") {
            use litchi_ppt::shapes::ShapeEnum;
            if matches!(
                shape,
                ShapeEnum::TextBox(_) | ShapeEnum::AutoShape(_) | ShapeEnum::Placeholder(_)
            ) {
                shape_ordinal += 1;
            }
            if matches!(shape, ShapeEnum::Placeholder(_)) {
                expected.push(format!("/slide[{}]/shape[{shape_ordinal}]", position + 1));
            }
        }
    }
    assert!(!expected.is_empty(), "fixture must contain placeholders");
    let file = eager(&bytes);
    assert_eq!(query_paths(&file, "slide > placeholder"), expected);
    for path in expected {
        let node = file.get(&path).expect("physical path");
        assert!(matches!(
            node.as_ppt().expect("PPT").shape(),
            Some(litchi_ppt::shapes::ShapeEnum::Placeholder(_))
        ));
    }
}
