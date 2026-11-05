//! Unified RTF parts (record 0328): sections, bookmarks, fields, tables of
//! contents, tab stops, text boxes and shapes, read through the DOCX paths
//! and property names, and written by native source splicing.

#![cfg(feature = "rtf")]

use litchi_ops::{Code, File, NewNode, open_bytes, props};

fn fixture() -> Vec<u8> {
    std::fs::read(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/unified-ops/0294/parts.rtf"),
    )
    .expect("read fixture")
}

fn parts() -> File {
    open_bytes(fixture()).expect("open RTF")
}

fn reopen(file: &File) -> File {
    open_bytes(file.to_bytes().expect("publish")).expect("reopen")
}

fn prop(file: &File, path: &str, key: &str) -> Option<String> {
    let node = file
        .get(path)
        .unwrap_or_else(|error| panic!("{path}: {error}"));
    node.view()
        .format()
        .expect("format")
        .into_iter()
        .find(|prop| prop.key() == key)
        .map(|prop| prop.to_string())
}

fn text(file: &File, path: &str) -> String {
    file.get(path)
        .unwrap_or_else(|error| panic!("{path}: {error}"))
        .view()
        .text()
        .expect("text")
        .unwrap_or_default()
        .into_owned()
}

fn count(file: &File, query: &str) -> usize {
    file.query(query)
        .unwrap_or_else(|error| panic!("{query}: {error}"))
        .count()
}

/// The source bytes with exactly one replacement.
fn assert_one_splice(before: &File, after: &File, old: &str, new: &str) {
    let before = String::from_utf8(before.to_bytes().unwrap()).unwrap();
    let after = String::from_utf8(after.to_bytes().unwrap()).unwrap();
    assert_eq!(before.matches(old).count(), 1, "{old} occurs once");
    assert_eq!(after, before.replacen(old, new, 1));
}

#[test]
fn parts_read_through_docx_paths() {
    let file = parts();
    assert_eq!(count(&file, "section"), 2);
    assert_eq!(
        prop(&file, "/section[2]", "type").as_deref(),
        Some("continuous")
    );
    assert_eq!(
        prop(&file, "/section[2]", "orientation").as_deref(),
        Some("landscape")
    );
    assert_eq!(prop(&file, "/section[2]", "columns").as_deref(), Some("2"));
    assert_eq!(count(&file, "bookmark"), 2);
    assert_eq!(
        prop(&file, "/bookmark[1]", "name").as_deref(),
        Some("Intro")
    );
    assert_eq!(text(&file, "/bookmark[@name=Second]"), "Second");
    assert_eq!(count(&file, "field"), 3);
    assert_eq!(
        prop(&file, "/field[1]", "fieldType").as_deref(),
        Some("page")
    );
    assert_eq!(count(&file, "toc"), 1);
    assert_eq!(prop(&file, "/toc[1]", "levels").as_deref(), Some("1-3"));
    assert_eq!(
        prop(&file, "/toc[1]", "hyperlinks").as_deref(),
        Some("true")
    );
    assert_eq!(count(&file, "/body/p[1]/tab"), 2);
    assert_eq!(
        prop(&file, "/body/p[1]/tab[2]", "val").as_deref(),
        Some("right")
    );
    assert_eq!(
        prop(&file, "/body/p[1]/tab[2]", "leader").as_deref(),
        Some("dot")
    );
    assert_eq!(text(&file, "/body/textbox[1]"), "Boxed text");
    assert_eq!(
        prop(&file, "/body/shape[1]", "name").as_deref(),
        Some("Box1")
    );
    assert_eq!(
        prop(&file, "/body/shape[1]", "geometry").as_deref(),
        Some("rect")
    );
}

#[test]
fn field_and_toc_sets_rewrite_the_instruction_only() {
    let file = parts();
    let edited = file
        .set("/field[2]", props! { "instruction" => "SECTIONPAGES" })
        .unwrap();
    assert_one_splice(&file, &edited, "{ NUMPAGES }", "{ SECTIONPAGES }");
    let edited = reopen(&edited);
    assert_eq!(
        prop(&edited, "/field[2]", "fieldType").as_deref(),
        Some("sectionpages")
    );
    let toc = file
        .set(
            "/toc[1]",
            props! { "levels" => "1-2", "hyperlinks" => false, "pageNumbers" => false },
        )
        .unwrap();
    assert_one_splice(
        &file,
        &toc,
        r#"{ TOC \\o "1-3" \\h \\z }"#,
        r#"{ TOC \\o "1-2" \\z \\n }"#,
    );
    let toc = reopen(&toc);
    assert_eq!(prop(&toc, "/toc[1]", "levels").as_deref(), Some("1-2"));
    assert_eq!(
        prop(&toc, "/toc[1]", "hyperlinks").as_deref(),
        Some("false")
    );
    assert_eq!(
        prop(&toc, "/toc[1]", "pageNumbers").as_deref(),
        Some("false")
    );
    // Read-only keys are refused.
    let error = file
        .set("/field[1]", props! { "fieldType" => "date" })
        .unwrap_err();
    assert_eq!(error.code(), Code::UnsupportedCapability);
    let error = file.remove("/field[1]").unwrap_err();
    assert_eq!(error.code(), Code::UnsupportedCapability);
}

#[test]
fn bookmarks_rename_and_remove() {
    let file = parts();
    let renamed = file
        .set("/bookmark[1]", props! { "name" => "Opening" })
        .unwrap();
    assert_one_splice(
        &file,
        &renamed,
        r"{\*\bkmkstart Intro}Intro{\*\bkmkend Intro}",
        r"{\*\bkmkstart Opening}Intro{\*\bkmkend Opening}",
    );
    let renamed = reopen(&renamed);
    assert_eq!(text(&renamed, "/bookmark[@name=Opening]"), "Intro");
    let error = file
        .set("/bookmark[1]", props! { "name" => "Second" })
        .unwrap_err();
    assert_eq!(error.code(), Code::InvalidValue, "{error}");
    let removed = reopen(&file.remove("/bookmark[@name=Second]").unwrap());
    assert_eq!(count(&removed, "bookmark"), 1);
    assert_eq!(text(&removed, "/body/p[4]"), text(&file, "/body/p[4]"));
}

#[test]
fn tab_stops_set_add_and_remove() {
    let file = parts();
    let edited = file
        .set(
            "/body/p[1]/tab[2]",
            props! { "val" => "center", "leader" => "none" },
        )
        .unwrap();
    assert_one_splice(&file, &edited, r"\tqr\tldot\tx5760", r"\tqc\tx5760");
    let edited = reopen(&edited);
    assert_eq!(
        prop(&edited, "/body/p[1]/tab[2]", "val").as_deref(),
        Some("center")
    );
    assert_eq!(
        prop(&edited, "/body/p[1]/tab[1]", "pos").as_deref(),
        Some("1440")
    );
    // Setting what a stop holds is an exact no-op.
    let same = file
        .set(
            "/body/p[1]/tab[1]",
            props! { "pos" => 1440, "val" => "left" },
        )
        .unwrap();
    assert_eq!(same.to_bytes().unwrap(), file.to_bytes().unwrap());
    let added = file
        .add(
            "/body/p[1]",
            NewNode::new("tab").prop("pos", 2880).prop("val", "decimal"),
        )
        .unwrap();
    assert_eq!(added.path.to_string(), "/body/p[1]/tab[2]");
    let added = reopen(&added.file);
    assert_eq!(count(&added, "/body/p[1]/tab"), 3);
    assert_eq!(
        prop(&added, "/body/p[1]/tab[2]", "val").as_deref(),
        Some("decimal")
    );
    let error = file
        .add("/body/p[1]", NewNode::new("tab").prop("val", "left"))
        .unwrap_err();
    assert_eq!(error.code(), Code::MissingProperty);
    let error = file
        .add("/body/p[1]", NewNode::new("tab").prop("pos", 1440))
        .unwrap_err();
    assert_eq!(error.code(), Code::InvalidValue);
    let removed = reopen(&file.remove("/body/p[1]/tab[1]").unwrap());
    assert_eq!(count(&removed, "/body/p[1]/tab"), 1);
    assert_eq!(
        prop(&removed, "/body/p[1]/tab[1]", "pos").as_deref(),
        Some("5760")
    );
    let first = reopen(
        &file
            .add("/body/p[2]", NewNode::new("tab").prop("pos", 720))
            .unwrap()
            .file,
    );
    assert_eq!(
        prop(&first, "/body/p[2]/tab[1]", "pos").as_deref(),
        Some("720")
    );
    assert_eq!(text(&first, "/body/p[2]"), text(&file, "/body/p[2]"));
}

#[test]
fn section_sets_rewrite_page_setup_controls() {
    let file = parts();
    let edited = file
        .set(
            "/section[2]",
            props! { "orientation" => "portrait", "columns" => 1, "type" => "nextPage" },
        )
        .unwrap();
    let edited = reopen(&edited);
    assert_eq!(
        prop(&edited, "/section[2]", "orientation").as_deref(),
        Some("portrait")
    );
    assert_eq!(
        prop(&edited, "/section[2]", "columns").as_deref(),
        Some("1")
    );
    assert_eq!(
        prop(&edited, "/section[2]", "type").as_deref(),
        Some("nextPage")
    );
    assert_eq!(
        prop(&edited, "/section[1]", "pageWidth"),
        prop(&file, "/section[1]", "pageWidth")
    );
    let margin = file
        .set("/section[1]", props! { "marginLeft" => "1.5in" })
        .unwrap();
    assert_one_splice(&file, &margin, r"\marglsxn1440", r"\marglsxn2160");
    let error = file
        .set("/section[1]", props! { "pageStart" => 3 })
        .unwrap_err();
    assert_eq!(error.code(), Code::UnsupportedCapability);
}

#[test]
fn read_only_parts_refuse_writes() {
    let file = parts();
    for (path, key) in [("/body/textbox[1]", "width"), ("/body/shape[1]", "width")] {
        let error = file.set(path, props! { key => "1in" }).unwrap_err();
        assert_eq!(error.code(), Code::UnsupportedCapability, "{path}");
        let error = file.remove(path).unwrap_err();
        assert_eq!(error.code(), Code::UnsupportedCapability, "{path}");
    }
    let error = file.remove("/section[1]").unwrap_err();
    assert_eq!(error.code(), Code::UnsupportedCapability);
}
