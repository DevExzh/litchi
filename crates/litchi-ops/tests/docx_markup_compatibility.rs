//! DOCX run-level markup compatibility (record 0307).
//!
//! Word writes drawings and text boxes as an `mc:AlternateContent` inside a
//! run: a DrawingML `Choice` and a VML `Fallback`, each with its own copy of
//! any text-box text. Readers show the selected branch; the document
//! transaction edits the source bytes, which hold both. The transaction
//! admits such a document, proves that block and run positions agree with
//! the visible view (the wrapper never leaves its run), and treats each
//! wrapper as one opaque inline object:
//!
//! - body `set`, `add`, `remove` and `move` work, and every wrapper keeps its
//!   bytes;
//! - a `text` set keeps a wrapper without text of its own and refuses, as a
//!   typed `unsupported_capability`, one whose branches hold text;
//! - find/replace and template merge never read or change a wrapper's text,
//!   and a find whose match would span a wrapper is refused.
//!
//! Every case runs on both owners and survives a save and reopen.
#![cfg(feature = "docx")]

use std::io::{Cursor, Read as _, Write as _};

use litchi_core::Position;
use litchi_ops::merge::{Options, Record};
use litchi_ops::{
    BatchItem, Code, File, NewNode, OpenOptions, Placement, Props, SourceMode, props,
};

/// A text box (`wps` with a VML fallback, `<Text>` in both branches) in the
/// first run of paragraph 27.
const TEXT_BOX: &str = "libreoffice-core/sw/qa/core/objectpositioning/data/do-not-capture-draw-objs-on-page-draw-wrap-none.docx";
/// One paragraph: a text run, then a run holding a shape without text.
const SHAPE: &str = "libreoffice-core/sw/qa/extras/ooxmlexport/data/tdf90153.docx";
/// One paragraph of two runs, each holding a text box.
const TWO_BOXES: &str = "libreoffice-core/sw/qa/extras/layout/data/inline-textbox-position.docx";
/// A group shape inside a tracked insertion, beside a check-box control.
const REDLINED_GROUP: &str =
    "libreoffice-core/sw/qa/writerfilter/dmapper/data/redlined-shape-sdt.docx";
/// Two paragraphs, each holding a drawing canvas; the second inside a
/// bookmark.
const CANVAS: &str = "libreoffice-core/sw/qa/extras/ooxmlexport/data/tdf103389.docx";

const OWNERS: [(&str, SourceMode); 2] =
    [("source", SourceMode::Auto), ("eager", SourceMode::Eager)];

fn fixture(relative: &str) -> Vec<u8> {
    let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data")
        .join(relative);
    std::fs::read(&path).unwrap_or_else(|error| panic!("{}: {error}", path.display()))
}

fn open(bytes: Vec<u8>, mode: SourceMode) -> File {
    litchi_ops::open_bytes_with(bytes, &OpenOptions::new().source_mode(mode)).unwrap()
}

/// Publish and reopen, so the assertions read the saved bytes.
fn reopen(file: &File, mode: SourceMode) -> File {
    open(file.to_bytes().unwrap(), mode)
}

fn document_xml(file: &File) -> String {
    let mut archive = zip::ZipArchive::new(Cursor::new(file.to_bytes().unwrap())).unwrap();
    let mut xml = String::new();
    archive
        .by_name("word/document.xml")
        .unwrap()
        .read_to_string(&mut xml)
        .unwrap();
    xml
}

/// `bytes` with `word/document.xml` rewritten by `change`.
fn rewrite(bytes: &[u8], change: impl Fn(&str) -> String) -> Vec<u8> {
    let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).unwrap();
    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let options =
        zip::write::SimpleFileOptions::default().compression_method(zip::CompressionMethod::Stored);
    for index in 0..archive.len() {
        let mut entry = archive.by_index(index).unwrap();
        let name = entry.name().to_owned();
        if name.ends_with('/') {
            continue;
        }
        let mut contents = Vec::new();
        entry.read_to_end(&mut contents).unwrap();
        if name == "word/document.xml" {
            contents = change(std::str::from_utf8(&contents).unwrap()).into_bytes();
        }
        zip.start_file(name, options).unwrap();
        zip.write_all(&contents).unwrap();
    }
    zip.finish().unwrap().into_inner()
}

/// Every `mc:AlternateContent` element of the main part, without the
/// whitespace between its tags: a changed paragraph is published compacted,
/// its wrappers included.
fn wrappers(xml: &str) -> Vec<String> {
    let close = "</mc:AlternateContent>";
    let mut found = Vec::new();
    let mut offset = 0;
    while let Some(start) = xml[offset..].find("<mc:AlternateContent") {
        let start = offset + start;
        let end = xml[start..].find(close).unwrap() + start + close.len();
        found.push(compact(&xml[start..end]));
        offset = end;
    }
    found
}

/// `xml` without whitespace-only text between tags.
fn compact(xml: &str) -> String {
    let mut output = String::with_capacity(xml.len());
    let mut rest = xml;
    while let Some(close) = rest.find('>') {
        output.push_str(&rest[..=close]);
        rest = &rest[close + 1..];
        let trimmed = rest.trim_start();
        if trimmed.starts_with('<') {
            rest = trimmed;
        }
    }
    output.push_str(rest);
    output
}

fn text(file: &File, path: &str) -> String {
    file.get(path)
        .unwrap_or_else(|error| panic!("{path}: {error}"))
        .view()
        .text()
        .unwrap()
        .map(std::borrow::Cow::into_owned)
        .unwrap_or_default()
}

fn paragraphs(file: &File) -> Vec<String> {
    let count = file.query("/body/p").unwrap().count();
    (1..=count)
        .map(|index| text(file, &format!("/body/p[{index}]")))
        .collect()
}

fn find(file: &File, path: &str, props: Props<'_>) -> litchi_ops::Result<(File, Option<usize>)> {
    let mut item = BatchItem::new("set");
    item.path = Some(path.into());
    item.props = Some(props);
    file.apply(&item)
        .map(|commit| (commit.file().clone(), commit.value().matched()))
}

/// The refusal of `edit`, after checking it left `file` untouched.
fn refused(file: &File, edit: impl FnOnce(&File) -> litchi_ops::Result<File>) -> litchi_ops::Error {
    let before = file.to_bytes().unwrap();
    let error = match edit(file) {
        Ok(_) => panic!("the edit was admitted"),
        Err(error) => error,
    };
    assert_eq!(file.to_bytes().unwrap(), before);
    error
}

#[test]
fn body_structure_edits_keep_every_wrapper_byte_for_byte() {
    for (owner, mode) in OWNERS {
        let file = open(fixture(TEXT_BOX), mode);
        let before = document_xml(&file);
        let kept = wrappers(&before);
        assert_eq!(kept.len(), 1);
        let texts = paragraphs(&file);
        assert_eq!(texts.len(), 31);
        assert!(
            texts[26].starts_with("<Text>To make"),
            "{owner}: {}",
            texts[26]
        );

        // Set the text of the paragraph after the wrapper's.
        let changed = file
            .set("/body/p[28]", props! { "text" => "Changed" })
            .unwrap_or_else(|error| panic!("{owner}: {error}"));
        // Remove the first paragraph: the wrapper's paragraph moves up.
        let removed = changed
            .remove("/body/p[1]")
            .unwrap_or_else(|error| panic!("{owner}: {error}"));
        // Add one at the start: it moves back.
        let added = removed
            .add(
                "/body",
                NewNode::new("paragraph")
                    .prop("text", "Added")
                    .at(Placement::Start),
            )
            .unwrap_or_else(|error| panic!("{owner}: {error}"));
        assert_eq!(added.path.to_string(), "/body/p[1]");
        let added = added.into_file();
        // Move the wrapper's paragraph after the next one, and set its
        // alignment.
        let moved = added
            .move_to("/body/p[27]", Placement::after("/body/p[28]").unwrap())
            .unwrap_or_else(|error| panic!("{owner}: {error}"));
        let aligned = moved
            .set("/body/p[28]", props! { "align" => "center" })
            .unwrap_or_else(|error| panic!("{owner}: {error}"));

        for file in [&aligned, &reopen(&aligned, mode)] {
            let texts = paragraphs(file);
            assert_eq!(texts.len(), 31, "{owner}");
            assert_eq!(texts[0], "Added");
            assert_eq!(texts[26], "Changed");
            assert_eq!(texts[27], paragraphs(&open(fixture(TEXT_BOX), mode))[26]);
            assert_eq!(wrappers(&document_xml(file)), kept, "{owner}");
        }
        // The source handle is untouched.
        assert_eq!(document_xml(&file), before);
    }
}

#[test]
fn run_edits_place_new_runs_beside_a_wrapper_run() {
    for (owner, mode) in OWNERS {
        let file = open(fixture(TEXT_BOX), mode);
        let kept = wrappers(&document_xml(&file));
        for (position, expected) in [
            (0, "New<Text>To make"),
            (1, "<Text>NewTo make"),
            (2, "<Text>To make"),
        ] {
            let added = file
                .add(
                    "/body/p[27]",
                    NewNode::new("run")
                        .prop("text", "New")
                        .at(Placement::Position(Position::new(position))),
                )
                .unwrap_or_else(|error| panic!("{owner} run {position}: {error}"));
            assert_eq!(
                added.path.to_string(),
                format!("/body/p[27]/r[{}]", position + 1)
            );
            let added = reopen(&added.into_file(), mode);
            assert!(
                text(&added, "/body/p[27]").starts_with(expected),
                "{owner} run {position}: {}",
                text(&added, "/body/p[27]")
            );
            assert_eq!(
                text(&added, &format!("/body/p[27]/r[{}]", position + 1)),
                "New"
            );
            assert_eq!(
                wrappers(&document_xml(&added)),
                kept,
                "{owner} run {position}"
            );
        }
        // The wrapper's run takes run properties and keeps the wrapper; the
        // text run beside it takes a text set.
        let bold = file
            .set("/body/p[27]/r[1]", props! { "bold" => true })
            .unwrap_or_else(|error| panic!("{owner}: {error}"));
        let retexted = bold
            .set("/body/p[27]/r[2]", props! { "text" => "Beside" })
            .unwrap_or_else(|error| panic!("{owner}: {error}"));
        let retexted = reopen(&retexted, mode);
        assert_eq!(text(&retexted, "/body/p[27]"), "<Text>Beside");
        assert_eq!(wrappers(&document_xml(&retexted)), kept, "{owner}");
        // Removing the wrapper's run is an explicit removal of the object.
        let removed = file.remove("/body/p[27]/r[1]").unwrap();
        assert!(wrappers(&document_xml(&removed)).is_empty());
        assert!(text(&removed, "/body/p[27]").starts_with("To make"));
    }
}

#[test]
fn a_text_set_keeps_a_wrapper_without_text() {
    for (owner, mode) in OWNERS {
        let file = open(fixture(SHAPE), mode);
        let kept = wrappers(&document_xml(&file));
        assert_eq!(kept.len(), 1);
        let changed = file
            .set("/body/p[1]", props! { "text" => "Changed" })
            .unwrap_or_else(|error| panic!("{owner}: {error}"));
        let changed = reopen(&changed, mode);
        assert_eq!(text(&changed, "/body/p[1]"), "Changed");
        assert_eq!(wrappers(&document_xml(&changed)), kept, "{owner}");
    }
}

#[test]
fn a_text_set_over_a_wrapper_with_text_is_refused() {
    for (owner, mode) in OWNERS {
        for (fixture_name, path) in [
            (TEXT_BOX, "/body/p[27]"),
            (TEXT_BOX, "/body/p[27]/r[1]"),
            (TWO_BOXES, "/body/p[1]"),
            (TWO_BOXES, "/body/p[1]/r[2]"),
        ] {
            let file = open(fixture(fixture_name), mode);
            let error = refused(&file, |file| file.set(path, props! { "text" => "Lost" }));
            assert_eq!(error.code(), Code::UnsupportedCapability, "{owner} {path}");
            let cause = std::error::Error::source(&error).map(ToString::to_string);
            assert!(
                cause
                    .as_deref()
                    .is_some_and(|cause| cause.contains("alternate content")),
                "{owner} {path}: {cause:?}"
            );
        }
    }
}

#[test]
fn find_never_reads_or_changes_a_wrapper() {
    for (owner, mode) in OWNERS {
        let file = open(fixture(TEXT_BOX), mode);
        let before = file.to_bytes().unwrap();
        let kept = wrappers(&document_xml(&file));
        // The text-box text is the wrapper's: no match, an exact no-op.
        let (unchanged, matched) = find(
            &file,
            "/",
            Props::new().with("find", "Text").with("replace", "TEXT"),
        )
        .unwrap_or_else(|error| panic!("{owner}: {error}"));
        assert_eq!(matched, Some(0));
        assert_eq!(unchanged.to_bytes().unwrap(), before, "{owner}");
        // Text beside the wrapper, in the same paragraph and elsewhere.
        for scope in ["/", "/body", "/body/p[27]"] {
            let (changed, matched) = find(
                &file,
                scope,
                Props::new()
                    .with("find", "To make")
                    .with("replace", "TO MAKE")
                    .with("bold", true),
            )
            .unwrap_or_else(|error| panic!("{owner} {scope}: {error}"));
            assert_eq!(matched, Some(1), "{owner} {scope}");
            let changed = reopen(&changed, mode);
            assert!(text(&changed, "/body/p[27]").starts_with("<Text>TO MAKE your"));
            assert_eq!(wrappers(&document_xml(&changed)), kept, "{owner} {scope}");
        }
    }
}

/// [`SHAPE`] with `before` and `after` text around the wrapper, in its run.
fn shape_between(before: &str, after: &str) -> Vec<u8> {
    rewrite(&fixture(SHAPE), |xml| {
        let start = xml.find("<mc:AlternateContent>").unwrap();
        let close = "</mc:AlternateContent>";
        let end = xml.find(close).unwrap() + close.len();
        format!(
            "{}<w:t>{before}</w:t>{}<w:t>{after}</w:t>{}",
            &xml[..start],
            &xml[start..end],
            &xml[end..]
        )
    })
}

#[test]
fn a_find_inside_a_wrapper_run_splits_text_and_keeps_the_wrapper_whole() {
    for (owner, mode) in OWNERS {
        let file = open(shape_between("ab", "cd"), mode);
        let kept = wrappers(&document_xml(&file));
        assert!(text(&file, "/body/p[1]").ends_with("abcd"));
        for (needle, replacement, expected) in [("b", "B", "aBcd"), ("cd", "CD", "abCD")] {
            let (changed, matched) = find(
                &file,
                "/body/p[1]",
                Props::new()
                    .with("find", needle)
                    .with("replace", replacement)
                    .with("italic", true),
            )
            .unwrap_or_else(|error| panic!("{owner} {needle}: {error}"));
            assert_eq!(matched, Some(1));
            let changed = reopen(&changed, mode);
            assert!(
                text(&changed, "/body/p[1]").ends_with(expected),
                "{owner} {needle}: {}",
                text(&changed, "/body/p[1]")
            );
            assert_eq!(wrappers(&document_xml(&changed)), kept, "{owner} {needle}");
        }
    }
}

#[test]
fn a_find_whose_match_would_span_a_wrapper_is_refused() {
    for (owner, mode) in OWNERS {
        let file = open(shape_between("ab", "cd"), mode);
        for needle in ["bc", "abcd"] {
            let error = refused(&file, |file| {
                find(
                    file,
                    "/",
                    Props::new().with("find", needle).with("replace", "x"),
                )
                .map(|(file, _)| file)
            });
            assert_eq!(
                error.code(),
                Code::UnsupportedCapability,
                "{owner} {needle}"
            );
            assert!(error.to_string().contains("alternate content"), "{error}");
        }
    }
}

#[test]
fn merge_fills_placeholders_beside_wrappers_and_never_inside_them() {
    for (owner, mode) in OWNERS {
        let template = open(
            rewrite(&fixture(TEXT_BOX), |xml| {
                xml.replace("&lt;Text&gt;", "{{name}}")
                    .replacen("To make", "{{name}} makes", 1)
            }),
            mode,
        );
        let kept = wrappers(&document_xml(&template));
        assert_eq!(kept[0].matches("{{name}}").count(), 2);
        let record: Record = [("name", "Ada")].into_iter().collect();
        let commit = template
            .merge(&record, &Options::new())
            .unwrap_or_else(|error| panic!("{owner}: {error}"));
        assert_eq!(commit.value().replacements(), 1, "{owner}");
        let merged = reopen(commit.file(), mode);
        assert!(text(&merged, "/body/p[27]").starts_with("{{name}}Ada makes"));
        assert_eq!(wrappers(&document_xml(&merged)), kept, "{owner}");
    }
}

#[test]
fn other_producer_wrappers_are_admitted() {
    for (owner, mode) in OWNERS {
        // Two canvases, one inside a bookmark, and two text boxes.
        for (fixture_name, count) in [(CANVAS, 2), (TWO_BOXES, 2)] {
            let file = open(fixture(fixture_name), mode);
            let kept = wrappers(&document_xml(&file));
            assert_eq!(kept.len(), count, "{fixture_name}");
            let before = paragraphs(&file);
            let added = file
                .add(
                    "/body",
                    NewNode::new("paragraph")
                        .prop("text", "Added")
                        .at(Placement::Start),
                )
                .unwrap_or_else(|error| panic!("{owner} {fixture_name}: {error}"))
                .into_file();
            let added = reopen(&added, mode);
            let after = paragraphs(&added);
            assert_eq!(after[0], "Added");
            assert_eq!(after[1..], before[..], "{owner} {fixture_name}");
            assert_eq!(
                wrappers(&document_xml(&added)),
                kept,
                "{owner} {fixture_name}"
            );
            let removed = added.remove("/body/p[1]").unwrap();
            assert_eq!(paragraphs(&removed), before, "{owner} {fixture_name}");
            // An empty find is an exact no-op over the wrappers.
            let (unchanged, matched) = find(
                &file,
                "/",
                Props::new().with("find", "\u{2603}").with("replace", "x"),
            )
            .unwrap();
            assert_eq!(matched, Some(0));
            assert_eq!(unchanged.to_bytes().unwrap(), file.to_bytes().unwrap());
        }
    }
}

#[test]
fn a_wrapper_in_a_tracked_insertion_reaches_the_revision_policy() {
    // A group inside a tracked insertion beside a check-box control, in a
    // document that tracks revisions: the markup-compatibility admission
    // passes, and the publication refuses for the tracked revisions, as it
    // does for any such document.
    for (owner, mode) in OWNERS {
        let file = open(fixture(REDLINED_GROUP), mode);
        assert_eq!(wrappers(&document_xml(&file)).len(), 1);
        let error = refused(&file, |file| {
            file.add(
                "/body",
                NewNode::new("paragraph")
                    .prop("text", "Added")
                    .at(Placement::Start),
            )
            .map(litchi_ops::Added::into_file)
        });
        let cause = std::error::Error::source(&error).map(ToString::to_string);
        assert!(
            cause
                .as_deref()
                .is_some_and(|cause| cause.contains("tracked revisions")),
            "{owner}: {cause:?}"
        );
    }
}
