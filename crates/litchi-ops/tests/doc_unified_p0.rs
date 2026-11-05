//! DOC P0 surface through the unified API (record 0258): run, table,
//! header/footer, note and comment reads; run and paragraph formatting;
//! paragraph add/remove/move/swap; exact no-ops, signature and encryption
//! policies, and read-back after reopen.
#![cfg(feature = "doc")]
#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    clippy::indexing_slicing,
    reason = "integration tests favor concise panic-driven assertions"
)]

use std::io::Cursor;

use litchi_doc::writer::{CharacterFormatting, CommentEntry, FootnoteEntry, ParagraphFormatting};
use litchi_ops::{Code, File, NewNode, Placement, open_bytes, props};

/// Paragraph 1 is plain, paragraph 2 has three runs ("plain ", bold "bold",
/// italic " tail"), then "after" and "last". The builder places the 2x2
/// table after the paragraphs and closes the story with an empty paragraph.
fn rich_doc() -> Vec<u8> {
    let mut builder = litchi_doc::Builder::new();
    builder.add_paragraph("first").unwrap();
    builder
        .add_paragraph_runs(
            vec![
                (
                    "plain ".to_string(),
                    CharacterFormatting {
                        font_name: Some("Arial".to_string()),
                        ..CharacterFormatting::default()
                    },
                ),
                (
                    "bold".to_string(),
                    CharacterFormatting {
                        bold: Some(true),
                        font_name: Some("Courier New".to_string()),
                        ..CharacterFormatting::default()
                    },
                ),
                (
                    " tail".to_string(),
                    CharacterFormatting {
                        italic: Some(true),
                        font_name: Some("Arial".to_string()),
                        ..CharacterFormatting::default()
                    },
                ),
            ],
            ParagraphFormatting::default(),
        )
        .unwrap();
    let table = builder.add_table(2, 2).unwrap();
    for (row, column, text) in [(0, 0, "a1"), (0, 1, "b1"), (1, 0, "a2"), (1, 1, "b2")] {
        builder
            .set_table_cell_text(table, row, column, text)
            .unwrap();
    }
    builder.add_paragraph("after").unwrap();
    builder.add_paragraph("last").unwrap();
    let mut output = Cursor::new(Vec::new());
    builder.build().unwrap().write_to(&mut output).unwrap();
    output.into_inner()
}

/// Two body paragraphs, the second anchoring a comment and a footnote, and
/// an odd header.
fn story_doc() -> Vec<u8> {
    let mut builder = litchi_doc::Builder::new();
    builder.add_paragraph("Body").unwrap();
    builder.add_paragraph("noted").unwrap();
    builder.set_odd_header("Header text");
    builder.add_footnote(FootnoteEntry::new(7, "Footnote text", 1));
    builder.add_comment(CommentEntry::new(6, "A remark", "Reviewer", "RV"));
    let mut output = Cursor::new(Vec::new());
    builder.build().unwrap().write_to(&mut output).unwrap();
    output.into_inner()
}

fn text(file: &File, path: &str) -> String {
    file.get(path)
        .unwrap_or_else(|error| panic!("{path}: {error}"))
        .view()
        .text()
        .unwrap()
        .unwrap_or_default()
        .into_owned()
}

fn prop(file: &File, path: &str, key: &str) -> Option<String> {
    let node = file
        .get(path)
        .unwrap_or_else(|error| panic!("{path}: {error}"));
    let view = node.view();
    view.format()
        .unwrap()
        .find(|prop| prop.key() == key)
        .map(|prop| prop.to_string())
}

fn body_texts(file: &File) -> Vec<String> {
    file.query("/body > p")
        .unwrap()
        .map(|node| {
            node.unwrap()
                .view()
                .text()
                .unwrap()
                .unwrap_or_default()
                .into_owned()
        })
        .collect()
}

fn reopen(file: &File) -> File {
    open_bytes(file.to_bytes().unwrap()).unwrap()
}

#[test]
fn reads_reach_runs_tables_headers_notes_and_comments() {
    let file = open_bytes(rich_doc()).unwrap();
    assert_eq!(text(&file, "/body/p[2]"), "plain bold tail");
    assert_eq!(text(&file, "/body/p[2]/r[2]"), "bold");
    assert_eq!(
        prop(&file, "/body/p[2]/r[2]", "bold").as_deref(),
        Some("true")
    );
    assert_eq!(
        prop(&file, "/body/p[2]/r[2]", "font").as_deref(),
        Some("Courier New")
    );
    assert_eq!(
        prop(&file, "/body/p[2]/r[3]", "italic").as_deref(),
        Some("true")
    );
    assert_eq!(prop(&file, "/body/p[2]/r[1]", "bold"), None);
    assert_eq!(prop(&file, "/body/p[1]", "align").as_deref(), Some("left"));

    assert_eq!(text(&file, "/body/tbl[1]/tr[1]/tc[2]"), "b1");
    assert_eq!(text(&file, "/body/tbl[1]/tr[2]/tc[1]/p[1]"), "a2");
    assert_eq!(prop(&file, "/body/tbl[1]", "rows").as_deref(), Some("2"));
    assert_eq!(prop(&file, "/body/tbl[1]", "cols").as_deref(), Some("2"));

    let stories = open_bytes(story_doc()).unwrap();
    assert_eq!(text(&stories, "/header[1]"), "Header text");
    assert_eq!(
        prop(&stories, "/header[1]", "type").as_deref(),
        Some("default")
    );
    assert_eq!(text(&stories, "/header[@type=default]"), "Header text");
    assert!(text(&stories, "/footnote[1]").contains("Footnote text"));
    assert!(text(&stories, "/comments/comment[1]").contains("A remark"));
    assert_eq!(
        prop(&stories, "/comments/comment[1]", "author").as_deref(),
        Some("Reviewer")
    );

    // The body lists paragraphs and the table in document order.
    let body = file.get("/body").unwrap();
    let kinds = body
        .view()
        .children()
        .map(|child| child.unwrap().view().kind().to_owned())
        .collect::<Vec<_>>();
    assert_eq!(
        kinds,
        [
            "paragraph",
            "paragraph",
            "paragraph",
            "paragraph",
            "table",
            "paragraph"
        ]
    );
    // `/body/p[N]` counts paragraphs outside tables.
    assert_eq!(text(&file, "/body/p[3]"), "after");

    let bold_runs = file
        .query("r[bold=true]")
        .unwrap()
        .map(|node| node.unwrap().path().to_string())
        .collect::<Vec<_>>();
    assert_eq!(bold_runs, ["/body/p[2]/r[2]"]);
    assert_eq!(file.query("tc").unwrap().count(), 4);
    assert_eq!(
        file.get("/body/tbl[2]").unwrap_err().code(),
        Code::PathNotFound
    );
    // Record 0324 projects hyperlinks: this paragraph has none.
    assert_eq!(
        file.get("/body/p[2]/hyperlink[1]").unwrap_err().code(),
        Code::PathNotFound
    );
    assert_eq!(
        file.get("/body/p[2]/picture[1]").unwrap_err().code(),
        Code::UnsupportedCapability
    );
}

#[test]
fn run_and_paragraph_formatting_round_trip() {
    let file = open_bytes(rich_doc()).unwrap();
    let edited = file
        .set(
            "/body/p[2]/r[2]",
            props! {
                "text" => "BOLDER",
                "italic" => true,
                "underline" => "single",
                "size" => "18pt",
                "color" => "FF0000",
                "font" => "Arial",
            },
        )
        .unwrap();
    for file in [edited.clone(), reopen(&edited)] {
        assert_eq!(text(&file, "/body/p[2]"), "plain BOLDER tail");
        let run = "/body/p[2]/r[2]";
        assert_eq!(text(&file, run), "BOLDER");
        assert_eq!(prop(&file, run, "bold").as_deref(), Some("true"));
        assert_eq!(prop(&file, run, "italic").as_deref(), Some("true"));
        assert_eq!(prop(&file, run, "underline").as_deref(), Some("single"));
        assert_eq!(prop(&file, run, "size").as_deref(), Some("18pt"));
        assert_eq!(prop(&file, run, "font").as_deref(), Some("Arial"));
        assert_eq!(prop(&file, run, "color").as_deref(), Some("#FF0000"));
        // Neighbours are untouched.
        assert_eq!(prop(&file, "/body/p[2]/r[3]", "size"), None);
        assert_eq!(text(&file, "/body/p[2]/r[3]"), " tail");
    }
    // The source handle is unchanged.
    assert_eq!(text(&file, "/body/p[2]/r[2]"), "bold");

    let whole = file
        .set("/body/p[1]", props! { "bold" => true, "size" => "14pt" })
        .unwrap();
    let whole = reopen(&whole);
    assert_eq!(
        prop(&whole, "/body/p[1]/r[1]", "bold").as_deref(),
        Some("true")
    );
    assert_eq!(
        prop(&whole, "/body/p[1]/r[1]", "size").as_deref(),
        Some("14pt")
    );

    // A scoped query formats every match in one native commit.
    let all = file
        .set("/body/p[2] > r", props! { "color" => "0000FF" })
        .unwrap();
    let all = reopen(&all);
    for run in 1..=3 {
        assert_eq!(
            prop(&all, &format!("/body/p[2]/r[{run}]"), "color").as_deref(),
            Some("#0000FF")
        );
    }
}

#[test]
fn unchanged_assignments_are_exact_noops() {
    let file = open_bytes(rich_doc()).unwrap();
    let source = file.to_bytes().unwrap();
    let cases: [(&str, litchi_ops::Props<'_>); 5] = [
        (
            "/body/p[2]/r[2]",
            litchi_ops::Props::from_props(props! { "bold" => true }).unwrap(),
        ),
        (
            "/body/p[2]/r[2]",
            litchi_ops::Props::from_props(props! { "text" => "bold", "font" => "Courier New" })
                .unwrap(),
        ),
        (
            "/body/p[2]/r[1]",
            litchi_ops::Props::from_props(props! { "italic" => false, "underline" => "none" })
                .unwrap(),
        ),
        (
            "/body/p[1]",
            litchi_ops::Props::from_props(props! { "text" => "first", "bold" => false }).unwrap(),
        ),
        (
            "/body/p[2]/r[1]",
            litchi_ops::Props::from_props(props! { "color" => "auto" }).unwrap(),
        ),
    ];
    for (path, props) in cases {
        let edited = file.set(path, props).unwrap();
        assert_eq!(edited.to_bytes().unwrap(), source, "{path}");
    }
    let moved = file.move_to(
        "/body/p[2]",
        Placement::Position(litchi_core::Position::new(1)),
    );
    assert_eq!(moved.unwrap().to_bytes().unwrap(), source);
    let swapped = file.swap("/body/p[3]", "/body/p[3]").unwrap();
    assert_eq!(swapped.to_bytes().unwrap(), source);
}

#[test]
fn paragraphs_add_remove_move_and_swap() {
    let file = open_bytes(rich_doc()).unwrap();
    assert_eq!(
        body_texts(&file),
        ["first", "plain bold tail", "after", "last", ""]
    );

    let added = file
        .add("/body", NewNode::new("p").prop("text", "appended"))
        .unwrap();
    assert_eq!(added.path.to_string(), "/body/p[6]");
    assert_eq!(
        body_texts(&reopen(&added.file)),
        ["first", "plain bold tail", "after", "last", "", "appended"]
    );
    let front = file
        .add(
            "/body",
            NewNode::new("paragraph")
                .prop("text", "front")
                .at(Placement::Start),
        )
        .unwrap();
    assert_eq!(
        body_texts(&reopen(&front.file)),
        ["front", "first", "plain bold tail", "after", "last", ""]
    );

    let removed = file.remove("/body/p[1]").unwrap();
    assert_eq!(
        body_texts(&reopen(&removed)),
        ["plain bold tail", "after", "last", ""]
    );
    let run_removed = file.remove("/body/p[2]/r[2]").unwrap();
    assert_eq!(text(&reopen(&run_removed), "/body/p[2]"), "plain  tail");

    let moved = file.move_to("/body/p[2]", Placement::End).unwrap();
    let moved = reopen(&moved);
    assert_eq!(
        body_texts(&moved),
        ["first", "after", "last", "", "plain bold tail"]
    );
    assert_eq!(
        prop(&moved, "/body/p[5]/r[2]", "bold").as_deref(),
        Some("true")
    );
    // The table stays where it was.
    assert_eq!(text(&moved, "/body/tbl[1]/tr[1]/tc[1]"), "a1");

    let swapped = file.swap("/body/p[1]", "/body/p[4]").unwrap();
    assert_eq!(
        body_texts(&reopen(&swapped)),
        ["last", "plain bold tail", "after", "first", ""]
    );

    // The source handle never changes.
    assert_eq!(
        body_texts(&file),
        ["first", "plain bold tail", "after", "last", ""]
    );
}

#[test]
fn refusals_are_typed() {
    let file = open_bytes(rich_doc()).unwrap();
    let cases = [
        file.set("/body/p[2]/r[2]", props! { "underline" => "double" })
            .unwrap_err(),
        file.set("/body/p[2]/r[2]", props! { "font" => "No Such Font" })
            .unwrap_err(),
        file.set("/body/tbl[1]/tr[1]/tc[1]/p[1]", props! { "text" => "x" })
            .unwrap_err(),
        open_bytes(story_doc())
            .unwrap()
            .set("/header[1]/p[1]", props! { "bold" => true })
            .unwrap_err(),
        file.set("/body/p[2]/r[2]", props! { "strike" => true })
            .unwrap_err(),
        file.add("/body/p[1]", NewNode::new("r").prop("text", "x"))
            .unwrap_err(),
    ];
    let codes = cases
        .iter()
        .map(litchi_ops::Error::code)
        .collect::<Vec<_>>();
    assert_eq!(
        codes,
        [
            Code::UnsupportedCapability,
            Code::NotFound,
            Code::UnsupportedCapability,
            Code::UnsupportedCapability,
            Code::UnsupportedCapability,
            Code::UnsupportedElement,
        ]
    );
    assert_eq!(
        file.set("/body/p[2]/r[2]", props! { "size" => "0.25pt" })
            .unwrap_err()
            .code(),
        Code::InvalidValue
    );
    let single = open_bytes({
        let mut builder = litchi_doc::Builder::new();
        builder.add_paragraph("only").unwrap();
        let mut output = Cursor::new(Vec::new());
        builder.build().unwrap().write_to(&mut output).unwrap();
        output.into_inner()
    })
    .unwrap();
    assert_eq!(
        single.remove("/body/p[1]").unwrap_err().code(),
        Code::InvalidArgument
    );
}

#[test]
fn signed_documents_refuse_every_write() {
    let plain = rich_doc();
    let signed = open_bytes(signed_compound(plain)).unwrap();
    assert!(matches!(signed, File::Signed(_)));
    let errors = [
        signed
            .set("/body/p[2]/r[2]", props! { "bold" => false })
            .unwrap_err(),
        signed
            .add("/body", NewNode::new("p").prop("text", "x"))
            .unwrap_err(),
        signed.remove("/body/p[1]").unwrap_err(),
        signed.move_to("/body/p[1]", Placement::End).unwrap_err(),
        signed.swap("/body/p[1]", "/body/p[2]").unwrap_err(),
    ];
    for error in errors {
        assert_eq!(error.code(), Code::DocumentProtected);
    }
    let File::Signed(wrapper) = signed else {
        unreachable!()
    };
    let unsigned = wrapper.unsign();
    let edited = unsigned.remove("/body/p[1]").unwrap();
    assert_eq!(body_texts(&reopen(&edited)).len(), 4);
}

#[cfg(feature = "encryption")]
#[test]
fn encrypted_documents_refuse_writes_without_downgrade() {
    let mut builder = litchi_doc::Builder::new();
    builder.add_paragraph("secret body").unwrap();
    builder.add_paragraph("second").unwrap();
    builder
        .set_password("pw", litchi_doc::writer::EncryptionProfile::OfficeBinaryRc4)
        .unwrap();
    let mut output = Cursor::new(Vec::new());
    builder.build().unwrap().write_to(&mut output).unwrap();
    let File::Locked(locked) = open_bytes(output.into_inner()).unwrap() else {
        panic!("an encrypted DOC opens locked")
    };
    let file = locked
        .unlock(litchi_ops::Password::new("pw".to_owned()))
        .unwrap();
    assert_eq!(text(&file, "/body/p[1]"), "secret body");
    let errors = [
        file.set("/body/p[1]", props! { "bold" => true })
            .unwrap_err(),
        file.set("/body/p[1]", props! { "text" => "x" })
            .unwrap_err(),
        file.add("/body", NewNode::new("p")).unwrap_err(),
        file.remove("/body/p[2]").unwrap_err(),
        file.move_to("/body/p[1]", Placement::End).unwrap_err(),
    ];
    for error in errors {
        assert_eq!(error.code(), Code::DocumentProtected, "{}", error.message());
    }
}

fn signed_compound(bytes: Vec<u8>) -> Vec<u8> {
    use litchi_cfb::{OleFile, OleWriter};

    let mut source = OleFile::open(Cursor::new(bytes), litchi_cfb::OleFileLimits::default())
        .expect("open DOC source");
    let mut writer = OleWriter::new();
    for path in source.list_streams() {
        let path = path.iter().map(String::as_str).collect::<Vec<_>>();
        for depth in 1..path.len() {
            writer.create_storage(&path[..depth]).unwrap();
        }
        let data = source.open_stream(&path).unwrap();
        writer.create_stream(&path, &data).unwrap();
    }
    writer
        .create_stream(&["_signatures"], b"test signature")
        .unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}

#[test]
fn transactions_and_batches_stage_structure_and_format_atomically() {
    let file = open_bytes(rich_doc()).unwrap();
    let edited = file
        .edit(|tx| {
            tx.set("/body/p[2]/r[1]", props! { "bold" => true })?;
            tx.add("/body", NewNode::new("p").prop("text", "added"))?;
            tx.remove("/body/p[1]")?;
            tx.swap("/body/p[1]", "/body/p[2]")?;
            Ok(())
        })
        .unwrap();
    let edited = reopen(edited.file());
    assert_eq!(
        body_texts(&edited),
        ["after", "plain bold tail", "last", "", "added"]
    );
    assert_eq!(
        prop(&edited, "/body/p[2]/r[1]", "bold").as_deref(),
        Some("true")
    );

    // A failing step rolls the whole transaction back.
    let error = file
        .edit(|tx| {
            tx.remove("/body/p[1]")?;
            tx.set("/body/p[9]", props! { "bold" => true })?;
            Ok(())
        })
        .unwrap_err();
    assert_eq!(error.code(), Code::PathNotFound);
    assert_eq!(
        body_texts(&file),
        ["first", "plain bold tail", "after", "last", ""]
    );
}

/// Over the DOC corpus: every story node reads, and the runs of every plain
/// body paragraph spell exactly its text, so run targets and paragraph
/// targets address the same characters.
#[test]
fn corpus_reads_agree_between_paragraphs_and_runs() {
    let directory =
        std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../test-data/ole/doc");
    let mut names = std::fs::read_dir(&directory)
        .unwrap()
        .map(|entry| entry.unwrap().file_name().into_string().unwrap())
        .filter(|name| name.ends_with(".doc"))
        .collect::<Vec<_>>();
    names.sort();
    let (mut files, mut paragraphs) = (0usize, 0usize);
    for name in names {
        let Ok(file @ File::Doc(_)) = open_bytes(std::fs::read(directory.join(&name)).unwrap())
        else {
            continue;
        };
        let Ok(root) = file.get("/") else { continue };
        files += 1;
        // Every top-level child reads or fails with a typed error.
        for child in root.view().children() {
            let Ok(child) = child else { continue };
            let _ = child.view().text();
            let _ = child.view().format().map(Iterator::count);
        }
        let Ok(body) = file.query("/body > p") else {
            continue;
        };
        for (index, node) in body.take(40).enumerate() {
            let Ok(node) = node else { break };
            let Ok(Some(text)) = node.view().text() else {
                continue;
            };
            if text.chars().any(char::is_control) {
                continue;
            }
            let path = format!("/body/p[{}]", index + 1);
            let paragraph = file.get(path.as_str()).unwrap();
            let runs = paragraph
                .view()
                .children()
                .map(|run| {
                    run.unwrap()
                        .view()
                        .text()
                        .unwrap()
                        .unwrap_or_default()
                        .into_owned()
                })
                .collect::<String>();
            assert_eq!(runs, text, "{name} {path}");
            paragraphs += 1;
        }
    }
    assert!(files > 25, "too few DOC files read: {files}");
    eprintln!("compared {paragraphs} paragraphs over {files} DOC files");
    assert!(
        paragraphs > 200,
        "too few paragraphs compared: {paragraphs}"
    );
}
