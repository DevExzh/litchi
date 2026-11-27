//! A source-backed DOCX promotes itself to an owning package for header,
//! footer, footnote and endnote writes. The promoted package must carry every
//! edit the handle already committed: each sequence below is applied to a
//! source-backed open (bytes and path) and to an eager open, as single
//! operations, as one edit closure and as one atomic batch, and the saved
//! packages must hold the same parts with the same bytes.
#![cfg(feature = "docx")]

use std::collections::BTreeMap;
use std::io::{Cursor, Read as _};
use std::path::PathBuf;

use litchi_core::FileFormat;
use litchi_ops::{BatchMode, File, NewNode, OpenOptions, SourceMode, props};

#[derive(Clone, Copy, Debug)]
enum Step {
    /// Add a node of `kind` with `text` under `parent`.
    Add(&'static str, &'static str, &'static str),
    /// Set the text of the node at the path.
    Set(&'static str, &'static str),
    /// Remove the node at the path.
    Remove(&'static str),
}

#[derive(Clone, Copy, Debug)]
enum Form {
    Single,
    Closure,
    Batch,
}

#[derive(Clone, Copy, Debug)]
enum Route {
    Bytes,
    Path,
}

fn fixture(name: &str) -> Vec<u8> {
    let path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/ooxml/docx")
        .join(name);
    std::fs::read(&path).unwrap_or_else(|error| panic!("{}: {error}", path.display()))
}

/// A document with a body paragraph, a header, a footer, a footnote and an
/// endnote, written by the eager owner so a later open can be source-backed.
fn all_stories() -> Vec<u8> {
    litchi_ops::create(FileFormat::Docx)
        .unwrap()
        .add("/body", NewNode::new("paragraph").prop("text", "Body"))
        .unwrap()
        .file
        .add("/", NewNode::new("header").prop("text", "Top"))
        .unwrap()
        .file
        .add("/", NewNode::new("footer").prop("text", "Bottom"))
        .unwrap()
        .file
        .add("/body/p[1]", NewNode::new("footnote").prop("text", "Foot"))
        .unwrap()
        .file
        .add(
            "/body/p[1]/r[1]",
            NewNode::new("endnote").prop("text", "End"),
        )
        .unwrap()
        .file
        .to_bytes()
        .unwrap()
}

fn open_eager(bytes: &[u8]) -> File {
    litchi_ops::open_bytes_with(
        bytes.to_vec(),
        &OpenOptions::new().source_mode(SourceMode::Eager),
    )
    .unwrap()
}

fn open_source(bytes: &[u8], route: Route, directory: &tempfile::TempDir) -> File {
    let file = match route {
        Route::Bytes => litchi_ops::open_bytes(bytes.to_vec()).unwrap(),
        Route::Path => {
            let path = directory.path().join("input.docx");
            std::fs::write(&path, bytes).unwrap();
            litchi_ops::open(&path).unwrap()
        },
    };
    assert!(
        matches!(&file, File::Docx(docx) if docx.is_source_backed()),
        "the default open of a DOCX is source-backed"
    );
    file
}

fn apply(file: &File, steps: &[Step], form: Form) -> File {
    match form {
        Form::Single => steps.iter().fold(file.clone(), |file, step| match *step {
            Step::Add(parent, kind, text) => {
                file.add(parent, NewNode::new(kind).prop("text", text))
                    .unwrap_or_else(|error| panic!("{step:?}: {error}"))
                    .file
            },
            Step::Set(path, text) => file
                .set(path, props! { "text" => text })
                .unwrap_or_else(|error| panic!("{step:?}: {error}")),
            Step::Remove(path) => file
                .remove(path)
                .unwrap_or_else(|error| panic!("{step:?}: {error}")),
        }),
        Form::Closure => file
            .edit(|tx| {
                for step in steps {
                    match *step {
                        Step::Add(parent, kind, text) => {
                            tx.add(parent, NewNode::new(kind).prop("text", text))?;
                        },
                        Step::Set(path, text) => tx.set(path, props! { "text" => text })?,
                        Step::Remove(path) => tx.remove(path)?,
                    }
                }
                Ok(())
            })
            .unwrap_or_else(|error| panic!("{steps:?}: {error}"))
            .into_file(),
        Form::Batch => {
            let items: Vec<String> = steps
                .iter()
                .map(|step| match *step {
                    Step::Add(parent, kind, text) => format!(
                        r#"{{"command":"add","parent":"{parent}","type":"{kind}","props":{{"text":"{text}"}}}}"#
                    ),
                    Step::Set(path, text) => format!(
                        r#"{{"command":"set","path":"{path}","props":{{"text":"{text}"}}}}"#
                    ),
                    Step::Remove(path) => format!(r#"{{"command":"remove","path":"{path}"}}"#),
                })
                .collect();
            let json = format!("[{}]", items.join(","));
            let items = litchi_ops::batch::from_bytes(
                json.as_bytes(),
                &litchi_ops::batch::SourceLimits::default(),
            )
            .unwrap();
            file.batch(items, BatchMode::Atomic)
                .unwrap_or_else(|error| panic!("{steps:?}: {error}"))
                .into_file()
        },
    }
}

/// Every member of a saved package, by name.
fn members(bytes: &[u8]) -> BTreeMap<String, Vec<u8>> {
    let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).unwrap();
    (0..archive.len())
        .map(|index| {
            let mut member = archive.by_index(index).unwrap();
            let mut data = Vec::new();
            member.read_to_end(&mut data).unwrap();
            (member.name().unwrap().into_owned(), data)
        })
        .collect()
}

/// The text of every XML member of a saved package, headers, footers and
/// notes included.
fn text(bytes: &[u8]) -> String {
    members(bytes)
        .into_values()
        .map(|data| String::from_utf8_lossy(&data).into_owned())
        .collect()
}

/// Apply `steps` on every source-backed route and in every form, and require
/// the saved package of each to equal the eager owner's, and to contain
/// `expected`.
fn check(input: &[u8], steps: &[Step], expected: &[&str]) {
    let directory = tempfile::tempdir().unwrap();
    for form in [Form::Single, Form::Closure, Form::Batch] {
        let oracle = apply(&open_eager(input), steps, form).to_bytes().unwrap();
        let oracle_text = text(&oracle);
        for needle in expected {
            assert!(
                oracle_text.contains(needle),
                "{form:?} {steps:?}: the eager output lacks {needle:?}"
            );
        }
        for route in [Route::Bytes, Route::Path] {
            let saved = apply(&open_source(input, route, &directory), steps, form)
                .to_bytes()
                .unwrap();
            let saved_text = text(&saved);
            for needle in expected {
                assert!(
                    saved_text.contains(needle),
                    "{route:?} {form:?} {steps:?}: the saved document lost {needle:?}"
                );
            }
            assert_eq!(
                members(&saved),
                members(&oracle),
                "{route:?} {form:?} {steps:?}: the saved package differs from the eager one"
            );
        }
    }
}

/// Apply `prefix` once, then each of `branches` to that same handle, so one
/// source-backed handle carrying edits is promoted for several writes. Each
/// branch's saved package must equal the eager owner's and contain the
/// branch's `expected` texts.
fn check_branches(input: &[u8], prefix: &[Step], branches: &[(&[Step], &[&str])]) {
    let directory = tempfile::tempdir().unwrap();
    for form in [Form::Single, Form::Closure, Form::Batch] {
        let eager = apply(&open_eager(input), prefix, form);
        for route in [Route::Bytes, Route::Path] {
            let source = apply(&open_source(input, route, &directory), prefix, form);
            // Twice over, so later branches reuse the handle's promotion.
            for (steps, expected) in branches.iter().chain(branches) {
                let oracle = apply(&eager, steps, form).to_bytes().unwrap();
                let saved = apply(&source, steps, form).to_bytes().unwrap();
                let saved_text = text(&saved);
                for needle in *expected {
                    assert!(
                        saved_text.contains(needle),
                        "{route:?} {form:?} {prefix:?} {steps:?}: the saved document lost {needle:?}"
                    );
                }
                assert_eq!(
                    members(&saved),
                    members(&oracle),
                    "{route:?} {form:?} {prefix:?} {steps:?}: the saved package differs from the eager one"
                );
            }
        }
    }
}

const BODY: Step = Step::Add("/body", "paragraph", "BODY_ADDED");
/// Changes the paragraph [`BODY`] added.
const BODY_SET: Step = Step::Set("/body/p[2]", "BODY_CHANGED");
const HEADER: Step = Step::Add("/header[1]", "paragraph", "HEADER_ADDED");
const FOOTER: Step = Step::Add("/footer[1]", "paragraph", "FOOTER_ADDED");

/// Both inputs: a Word-authored document with headers and footers, and a
/// document holding every story kind.
fn inputs() -> [Vec<u8>; 2] {
    [fixture("headerFooter.docx"), all_stories()]
}

#[test]
fn a_header_write_keeps_earlier_body_edits() {
    for input in inputs() {
        check(&input, &[BODY, HEADER], &["BODY_ADDED", "HEADER_ADDED"]);
    }
    check(
        &all_stories(),
        &[
            BODY,
            BODY_SET,
            Step::Set("/header[1]/p[1]", "HEADER_CHANGED"),
        ],
        &["BODY_CHANGED", "HEADER_CHANGED"],
    );
}

#[test]
fn a_footer_write_keeps_earlier_body_edits() {
    for input in inputs() {
        check(&input, &[BODY, FOOTER], &["BODY_ADDED", "FOOTER_ADDED"]);
    }
    check(
        &all_stories(),
        &[
            BODY,
            BODY_SET,
            Step::Set("/footer[1]/p[1]", "FOOTER_CHANGED"),
        ],
        &["BODY_CHANGED", "FOOTER_CHANGED"],
    );
}

#[test]
fn a_note_write_keeps_earlier_body_edits() {
    for input in inputs() {
        check(
            &input,
            &[BODY, Step::Add("/body/p[2]", "footnote", "FOOTNOTE_ADDED")],
            &["BODY_ADDED", "FOOTNOTE_ADDED"],
        );
        check(
            &input,
            &[
                BODY,
                Step::Add("/body/p[2]/r[1]", "endnote", "ENDNOTE_ADDED"),
            ],
            &["BODY_ADDED", "ENDNOTE_ADDED"],
        );
    }
    check(
        &all_stories(),
        &[
            BODY,
            BODY_SET,
            Step::Set("/footnote[1]/p[1]", "FOOTNOTE_CHANGED"),
        ],
        &["BODY_CHANGED", "FOOTNOTE_CHANGED"],
    );
    check(
        &all_stories(),
        &[
            BODY,
            BODY_SET,
            Step::Set("/endnote[1]/p[1]", "ENDNOTE_CHANGED"),
        ],
        &["BODY_CHANGED", "ENDNOTE_CHANGED"],
    );
}

#[test]
fn body_edits_between_story_writes_survive() {
    for input in inputs() {
        check(
            &input,
            &[HEADER, BODY, FOOTER],
            &["HEADER_ADDED", "BODY_ADDED", "FOOTER_ADDED"],
        );
        check(
            &input,
            &[BODY, BODY_SET, HEADER, BODY, FOOTER],
            &["BODY_CHANGED", "HEADER_ADDED", "BODY_ADDED", "FOOTER_ADDED"],
        );
    }
}

#[test]
fn a_story_removal_keeps_earlier_body_edits() {
    for input in inputs() {
        check(
            &input,
            &[BODY, BODY_SET, Step::Remove("/header[1]")],
            &["BODY_CHANGED"],
        );
        check(&input, &[BODY, Step::Remove("/footer[1]")], &["BODY_ADDED"]);
    }
}

#[test]
fn several_promotions_in_one_handle_keep_every_edit() {
    for input in inputs() {
        check(
            &input,
            &[
                BODY,
                HEADER,
                Step::Add("/body", "paragraph", "SECOND_BODY"),
                FOOTER,
                BODY_SET,
                Step::Add("/body/p[2]", "footnote", "FOOTNOTE_ADDED"),
            ],
            &[
                "HEADER_ADDED",
                "SECOND_BODY",
                "FOOTER_ADDED",
                "BODY_CHANGED",
                "FOOTNOTE_ADDED",
            ],
        );
        check_branches(
            &input,
            &[BODY, Step::Add("/body", "paragraph", "SECOND_BODY")],
            &[
                (&[HEADER], &["BODY_ADDED", "SECOND_BODY", "HEADER_ADDED"]),
                (&[FOOTER], &["BODY_ADDED", "SECOND_BODY", "FOOTER_ADDED"]),
                (
                    &[Step::Add("/body/p[2]", "footnote", "FOOTNOTE_ADDED")],
                    &["BODY_ADDED", "SECOND_BODY", "FOOTNOTE_ADDED"],
                ),
                (
                    &[Step::Set("/body/p[3]", "THIRD_CHANGED"), HEADER],
                    &["BODY_ADDED", "THIRD_CHANGED", "HEADER_ADDED"],
                ),
            ],
        );
    }
}
