//! Unified PPTX paragraph and run `add`/`remove` (record 0333) in shape,
//! table-cell and notes bodies, through both native owner routes, read back
//! through `get` and after a byte round trip.
#![cfg(feature = "pptx")]

use litchi_ops::{Code, File, NewNode, OpenOptions, Placement, SourceMode};

const SHAPE: &str = "/slide[1]/shape[1]";
const CELL: &str = "/slide[1]/table[1]/tr[2]/tc[1]";
const NOTES: &str = "/slide[1]/notes";

fn fixture() -> Vec<u8> {
    let mut builder = litchi_pptx::builder::Builder::new().unwrap();
    builder
        .presentation()
        .unwrap()
        .add_slide()
        .unwrap()
        .add_text_box("alpha", 0, 0, 1_828_800, 914_400);
    let package = builder.build().unwrap();
    let mut edit = package.edit().unwrap();
    edit.add_table(
        0usize,
        &[
            vec!["a1".to_owned(), "b1".to_owned()],
            vec!["a2".to_owned(), "b2".to_owned()],
        ],
        (0, 1_000_000, 2_000_000, 600_000),
    )
    .unwrap();
    edit.set_notes_text(0usize, "note one").unwrap();
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

fn reopen(file: &File, eager: bool) -> File {
    open(&file.to_bytes().unwrap(), eager)
}

fn text(file: &File, path: &str) -> String {
    file.get(path)
        .unwrap()
        .view()
        .text()
        .unwrap()
        .unwrap_or_default()
        .into_owned()
}

fn prop(file: &File, path: &str, key: &str) -> Option<String> {
    let node = file.get(path).unwrap();
    node.view()
        .format()
        .unwrap()
        .find(|property| property.key() == key)
        .map(|property| property.value().to_string())
}

fn source_backed(file: &File) -> bool {
    matches!(file, File::Pptx(handle) if handle.is_source_backed())
}

fn code(result: litchi_ops::Result<impl Sized>) -> Code {
    match result {
        Ok(_) => panic!("expected a refusal"),
        Err(error) => error.code(),
    }
}

#[test]
fn shape_paragraphs_are_added_with_placement_and_formatting() {
    let bytes = fixture();
    for eager in [false, true] {
        let file = open(&bytes, eager);
        let added = file
            .add(
                SHAPE,
                NewNode::new("paragraph")
                    .prop("text", "omega")
                    .prop("bold", true)
                    .prop("align", "center")
                    .prop("size", "20pt"),
            )
            .unwrap();
        assert_eq!(added.path.to_string(), "/slide[1]/shape[1]/paragraph[2]");
        // A source handle stays source-backed for slide bodies.
        assert_eq!(source_backed(&added.file), !eager, "{eager}");
        let before = added
            .file
            .add(
                SHAPE,
                NewNode::new("p")
                    .prop("text", "first")
                    .at(Placement::before("/slide[1]/shape[1]/paragraph[1]").unwrap()),
            )
            .unwrap();
        assert_eq!(before.path.to_string(), "/slide[1]/shape[1]/paragraph[1]");
        let middle = before
            .file
            .add(
                SHAPE,
                NewNode::new("paragraph")
                    .prop("text", "middle")
                    .at(Placement::after("/slide[1]/shape[1]/paragraph[2]").unwrap()),
            )
            .unwrap();
        assert_eq!(middle.path.to_string(), "/slide[1]/shape[1]/paragraph[3]");
        let past = middle
            .file
            .add(
                SHAPE,
                NewNode::new("paragraph")
                    .prop("text", "last")
                    .at(Placement::Position(litchi_core::Position::new(40))),
            )
            .unwrap();
        assert_eq!(past.path.to_string(), "/slide[1]/shape[1]/paragraph[5]");
        for file in [&past.file, &reopen(&past.file, eager)] {
            assert_eq!(text(file, SHAPE), "first\nalpha\nmiddle\nomega\nlast");
            let omega = "/slide[1]/shape[1]/paragraph[4]";
            assert_eq!(text(file, omega), "omega");
            assert_eq!(prop(file, omega, "align").as_deref(), Some("center"));
            assert_eq!(
                prop(file, &format!("{omega}/run[1]"), "bold").as_deref(),
                Some("true")
            );
            assert_eq!(
                prop(file, &format!("{omega}/run[1]"), "size").as_deref(),
                Some("20pt")
            );
        }
        // The input handle is unchanged.
        assert_eq!(text(&file, SHAPE), "alpha");
    }
}

#[test]
fn runs_are_added_and_removed_inside_a_paragraph() {
    let bytes = fixture();
    for eager in [false, true] {
        let file = open(&bytes, eager);
        let paragraph = "/slide[1]/shape[1]/paragraph[1]";
        let end = file
            .add(
                paragraph,
                NewNode::new("run")
                    .prop("text", " omega")
                    .prop("italic", true)
                    .prop("color", "C00000"),
            )
            .unwrap();
        assert_eq!(
            end.path.to_string(),
            "/slide[1]/shape[1]/paragraph[1]/run[2]"
        );
        let start = end
            .file
            .add(
                paragraph,
                NewNode::new("r").prop("text", "pre ").at(Placement::Start),
            )
            .unwrap();
        assert_eq!(
            start.path.to_string(),
            "/slide[1]/shape[1]/paragraph[1]/run[1]"
        );
        let file = reopen(&start.file, eager);
        assert_eq!(text(&file, paragraph), "pre alpha omega");
        let omega = "/slide[1]/shape[1]/paragraph[1]/run[3]";
        assert_eq!(prop(&file, omega, "italic").as_deref(), Some("true"));
        assert_eq!(prop(&file, omega, "color").as_deref(), Some("#C00000"));
        let removed = file
            .remove("/slide[1]/shape[1]/paragraph[1]/run[2]")
            .unwrap();
        assert_eq!(text(&removed, paragraph), "pre  omega");
        // A run cannot hold a line break or take paragraph alignment.
        assert_eq!(
            code(file.add(paragraph, NewNode::new("run").prop("text", "a\nb"))),
            Code::InvalidValue
        );
        assert_eq!(
            code(file.add(paragraph, NewNode::new("run").prop("align", "left"))),
            Code::UnsupportedProperty
        );
        // A run's parent is a paragraph.
        assert_eq!(
            code(file.add(SHAPE, NewNode::new("run").prop("text", "x"))),
            Code::UnsupportedType
        );
    }
}

#[test]
fn paragraph_removal_keeps_one_paragraph() {
    let bytes = fixture();
    for eager in [false, true] {
        let file = open(&bytes, eager);
        let added = file
            .add(SHAPE, NewNode::new("paragraph").prop("text", "beta"))
            .unwrap();
        let removed = added
            .file
            .remove("/slide[1]/shape[1]/paragraph[1]")
            .unwrap();
        assert_eq!(text(&reopen(&removed, eager), SHAPE), "beta");
        assert_eq!(
            code(removed.remove("/slide[1]/shape[1]/paragraph[1]")),
            Code::UnsupportedCapability
        );
        assert_eq!(
            code(removed.remove("/slide[1]/shape[1]/paragraph[4]")),
            Code::PathNotFound
        );
    }
}

#[test]
fn cell_paragraphs_and_runs_are_addressable_and_writable() {
    let bytes = fixture();
    for eager in [false, true] {
        let file = open(&bytes, eager);
        assert_eq!(text(&file, &format!("{CELL}/paragraph[1]")), "a2");
        assert_eq!(text(&file, &format!("{CELL}/paragraph[1]/run[1]")), "a2");
        let added = file
            .add(CELL, NewNode::new("paragraph").prop("text", "second"))
            .unwrap();
        assert_eq!(
            added.path.to_string(),
            "/slide[1]/table[1]/tr[2]/tc[1]/paragraph[2]"
        );
        let run = added
            .file
            .add(
                &format!("{CELL}/paragraph[1]"),
                NewNode::new("run").prop("text", "!").prop("bold", true),
            )
            .unwrap();
        let file = reopen(&run.file, eager);
        assert_eq!(text(&file, &format!("{CELL}/paragraph[1]")), "a2!");
        assert_eq!(text(&file, &format!("{CELL}/paragraph[2]")), "second");
        assert_eq!(
            prop(&file, &format!("{CELL}/paragraph[1]/run[2]"), "bold").as_deref(),
            Some("true")
        );
        // A cell paragraph takes formatting through `set`.
        let set = file
            .set(
                &format!("{CELL}/paragraph[2]"),
                litchi_ops::props! { "align" => "right", "italic" => true },
            )
            .unwrap();
        assert_eq!(
            prop(&set, &format!("{CELL}/paragraph[2]"), "align").as_deref(),
            Some("right")
        );
        assert_eq!(
            prop(&set, &format!("{CELL}/paragraph[2]/run[1]"), "italic").as_deref(),
            Some("true")
        );
        let removed = set.remove(&format!("{CELL}/paragraph[1]")).unwrap();
        assert_eq!(text(&removed, &format!("{CELL}/paragraph[1]")), "second");
        // Other cells are untouched; unscoped queries do not list cell text.
        assert_eq!(text(&removed, "/slide[1]/table[1]/tr[1]/tc[1]"), "a1");
        assert_eq!(
            removed.query("paragraph").unwrap().count(),
            open(&bytes, eager).query("paragraph").unwrap().count()
        );
    }
}

#[test]
fn notes_paragraphs_are_addressable_and_writable() {
    let bytes = fixture();
    for eager in [false, true] {
        let file = open(&bytes, eager);
        assert_eq!(text(&file, &format!("{NOTES}/paragraph[1]")), "note one");
        let added = file
            .add(NOTES, NewNode::new("paragraph").prop("text", "note two"))
            .unwrap();
        assert_eq!(added.path.to_string(), "/slide[1]/notes/paragraph[2]");
        let file = reopen(&added.file, eager);
        assert_eq!(text(&file, NOTES), "note one\nnote two");
        assert_eq!(text(&file, &format!("{NOTES}/paragraph[2]")), "note two");
        let run = file
            .add(
                &format!("{NOTES}/paragraph[2]"),
                NewNode::new("run").prop("text", "!"),
            )
            .unwrap();
        let bold = run
            .file
            .set(
                &format!("{NOTES}/paragraph[2]/run[2]"),
                litchi_ops::props! { "bold" => true },
            )
            .unwrap();
        assert_eq!(
            prop(&bold, &format!("{NOTES}/paragraph[2]/run[2]"), "bold").as_deref(),
            Some("true")
        );
        let removed = bold.remove(&format!("{NOTES}/paragraph[1]")).unwrap();
        assert_eq!(text(&reopen(&removed, eager), NOTES), "note two!");
        // The slide's own text is untouched.
        assert_eq!(text(&removed, SHAPE), "alpha");
    }
}

#[test]
fn edits_compose_in_one_atomic_closure() {
    let bytes = fixture();
    for eager in [false, true] {
        let file = open(&bytes, eager);
        let commit = file
            .edit(|tx| {
                tx.add(SHAPE, NewNode::new("paragraph").prop("text", "two"))?;
                tx.add(
                    "/slide[1]/shape[1]/paragraph[2]",
                    NewNode::new("run").prop("text", "!"),
                )?;
                tx.remove("/slide[1]/shape[1]/paragraph[1]")?;
                Ok::<_, litchi_ops::Error>(())
            })
            .unwrap();
        assert_eq!(text(commit.file(), SHAPE), "two!");
        // A failing item leaves the input untouched.
        let failed = file.edit(|tx| {
            tx.add(SHAPE, NewNode::new("paragraph").prop("text", "two"))?;
            tx.remove("/slide[1]/shape[1]/paragraph[9]")?;
            Ok::<_, litchi_ops::Error>(())
        });
        assert!(failed.is_err());
        assert_eq!(text(&file, SHAPE), "alpha");
    }
}
