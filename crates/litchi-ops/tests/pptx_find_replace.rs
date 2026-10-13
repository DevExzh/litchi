//! PPTX find/replace through `set` with `find=` (record 0273), on both
//! native owners: cross-run matches collapse into the run where they start,
//! format properties apply to exactly the replaced or matched text, the
//! scope is the target's subtree, zero matches are an exact no-op, and a
//! table cell that would match is refused rather than skipped.
#![cfg(feature = "pptx")]

use litchi_core::Position;
use litchi_ops::{BatchItem, Code, File, OpenOptions, Props, SourceMode, props};
use litchi_pptx::opened::{TextEdit, TextProperties, TextTarget};

/// Slide 1: "draft one, dra" + bold "ft two", and "Draft notes".
/// Slide 2: a table whose first cell reads "draft".
fn fixture() -> Vec<u8> {
    let mut builder = litchi_pptx::builder::Builder::new().unwrap();
    let presentation = builder.presentation().unwrap();
    let slide = presentation.add_slide().unwrap();
    slide.add_text_box("draft one, draft two", 0, 0, 1_828_800, 914_400);
    slide.add_text_box("Draft notes", 0, 1_000_000, 1_828_800, 914_400);
    presentation.add_slide().unwrap();
    let package = builder.build().unwrap();
    let mut edit = package.edit().unwrap();
    edit.set_text_properties(
        0usize,
        &[TextEdit::new(
            TextTarget::RunSpan {
                shape: Position::new(0),
                paragraph: Position::new(0),
                run: Position::new(0),
                start: 14,
                end: 20,
            },
            TextProperties::default().with_bold(Some(true)),
        )],
    )
    .unwrap();
    edit.add_table(
        1usize,
        &[vec!["draft".to_owned(), "b1".to_owned()]],
        (0, 1_000_000, 2_000_000, 600_000),
    )
    .unwrap();
    edit.commit().unwrap().into_package().to_bytes().unwrap()
}

fn open(bytes: &[u8], eager: bool) -> File {
    litchi_ops::open_bytes_with(
        bytes.to_vec(),
        &OpenOptions::new().source_mode(if eager {
            SourceMode::Eager
        } else {
            SourceMode::Auto
        }),
    )
    .unwrap()
}

fn find(file: &File, path: &str, props: Props<'_>) -> litchi_ops::Result<(File, Option<usize>)> {
    let mut item = BatchItem::new("set");
    item.path = Some(path.into());
    item.props = Some(props);
    let commit = file.apply(&item)?;
    Ok((commit.file, commit.value.matched()))
}

fn text(file: &File, path: &str) -> String {
    file.get(path)
        .unwrap()
        .view()
        .text()
        .unwrap()
        .map(std::borrow::Cow::into_owned)
        .unwrap_or_default()
}

fn bold(file: &File, path: &str) -> bool {
    file.get(path)
        .unwrap()
        .view()
        .format()
        .unwrap()
        .any(|property| property.key() == "bold" && property.value().to_string() == "true")
}

#[test]
fn a_cross_run_match_takes_the_run_where_it_starts() {
    let bytes = fixture();
    for eager in [false, true] {
        let file = open(&bytes, eager);
        assert!(bold(&file, "/slide[1]/shape[1]/paragraph[1]/run[2]"));
        let (after, matched) = find(
            &file,
            "/slide[1]/shape[1]",
            Props::new()
                .with("find", "draft two")
                .with("replace", "final 2"),
        )
        .unwrap();
        assert_eq!(matched, Some(1));
        assert_eq!(text(&after, "/slide[1]/shape[1]"), "draft one, final 2");
        assert_eq!(
            text(&after, "/slide[1]/shape[1]/paragraph[1]/run[1]"),
            "draft one, final 2"
        );
        assert!(!bold(&after, "/slide[1]/shape[1]/paragraph[1]/run[1]"));
        let reopened = open(&after.to_bytes().unwrap(), eager);
        assert_eq!(text(&reopened, "/slide[1]/shape[1]"), "draft one, final 2");
    }
}

#[test]
fn scopes_follow_the_target() {
    let bytes = fixture();
    for eager in [false, true] {
        let file = open(&bytes, eager);
        let (after, matched) = find(
            &file,
            "/slide[1]",
            Props::new().with("find", "draft").with("replace", "final"),
        )
        .unwrap();
        assert_eq!(matched, Some(2));
        assert_eq!(text(&after, "/slide[1]/shape[1]"), "final one, final two");
        assert_eq!(text(&after, "/slide[1]/shape[2]"), "Draft notes");
        let (after, matched) = find(
            &file,
            "/slide[1]",
            Props::new()
                .with("find", "draft")
                .with("replace", "final")
                .with("matchCase", false),
        )
        .unwrap();
        assert_eq!(matched, Some(3));
        assert_eq!(text(&after, "/slide[1]/shape[2]"), "final notes");
        // A run scope keeps only matches inside the run.
        let (_, matched) = find(
            &file,
            "/slide[1]/shape[1]/paragraph[1]/run[2]",
            Props::new().with("find", "draft").with("replace", "x"),
        )
        .unwrap();
        assert_eq!(matched, Some(0));
        let (_, matched) = find(
            &file,
            "/slide[1]/shape[1]/paragraph[1]/run[2]",
            Props::new().with("find", "two").with("replace", "2"),
        )
        .unwrap();
        assert_eq!(matched, Some(1));
    }
}

#[test]
fn format_properties_cover_exactly_the_replacement() {
    let bytes = fixture();
    for eager in [false, true] {
        let file = open(&bytes, eager);
        let (after, matched) = find(
            &file,
            "/slide[1]/shape[1]",
            Props::new()
                .with("find", "one")
                .with("replace", "ONE")
                .with("italic", true),
        )
        .unwrap();
        assert_eq!(matched, Some(1));
        assert_eq!(text(&after, "/slide[1]/shape[1]"), "draft ONE, draft two");
        assert_eq!(
            text(&after, "/slide[1]/shape[1]/paragraph[1]/run[1]"),
            "draft "
        );
        assert_eq!(
            text(&after, "/slide[1]/shape[1]/paragraph[1]/run[2]"),
            "ONE"
        );
        let italic = after
            .get("/slide[1]/shape[1]/paragraph[1]/run[2]")
            .unwrap()
            .view()
            .format()
            .unwrap()
            .any(|property| property.key() == "italic");
        assert!(italic);
        assert_eq!(
            text(&after, "/slide[1]/shape[1]/paragraph[1]/run[3]"),
            ", dra"
        );
        // Format only: every matched slice, across the run boundary.
        let (after, matched) = find(
            &file,
            "/slide[1]/shape[1]",
            Props::new().with("find", "draft two").with("italic", true),
        )
        .unwrap();
        assert_eq!(matched, Some(1));
        assert_eq!(text(&after, "/slide[1]/shape[1]"), "draft one, draft two");
        assert_eq!(
            text(&after, "/slide[1]/shape[1]/paragraph[1]/run[2]"),
            "dra"
        );
        assert_eq!(
            text(&after, "/slide[1]/shape[1]/paragraph[1]/run[3]"),
            "ft two"
        );
        assert!(bold(&after, "/slide[1]/shape[1]/paragraph[1]/run[3]"));
    }
}

#[test]
fn zero_matches_are_an_exact_no_op() {
    let bytes = fixture();
    for eager in [false, true] {
        let file = open(&bytes, eager);
        let before = file.to_bytes().unwrap();
        let (after, matched) = find(
            &file,
            "/slide[1]",
            Props::new().with("find", "absent").with("replace", "x"),
        )
        .unwrap();
        assert_eq!(matched, Some(0));
        assert_eq!(after.to_bytes().unwrap(), before);
        let set = file
            .set("/", props! { "find" => "absent", "replace" => "x" })
            .unwrap();
        assert_eq!(set.to_bytes().unwrap(), before);
    }
}

#[test]
fn a_matching_table_cell_is_refused_not_skipped() {
    let file = open(&fixture(), true);
    let error = find(
        &file,
        "/",
        Props::new().with("find", "draft").with("replace", "final"),
    )
    .unwrap_err();
    assert_eq!(error.code(), Code::UnsupportedCapability);
    assert_eq!(error.property(), Some("find"));
    // A table without a match does not block the find.
    let (_, matched) = find(
        &file,
        "/",
        Props::new().with("find", "two").with("replace", "2"),
    )
    .unwrap();
    assert_eq!(matched, Some(1));
}
