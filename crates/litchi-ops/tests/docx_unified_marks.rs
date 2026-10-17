//! DOCX tab stops, breaks, bookmarks, fields, tables of contents, the
//! watermark and text boxes through the `File` facade (record 0319), on both
//! owners.
#![cfg(feature = "docx")]

use std::path::PathBuf;

use litchi_ops::{File, NewNode, OpenOptions, SourceMode, props};

fn fixture(name: &str) -> Vec<u8> {
    let path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data")
        .join(name);
    std::fs::read(&path).unwrap_or_else(|error| panic!("{}: {error}", path.display()))
}

fn open(eager: bool, bytes: Vec<u8>) -> File {
    if eager {
        litchi_ops::open_bytes_with(bytes, &OpenOptions::new().source_mode(SourceMode::Eager))
            .expect("eager fixture")
    } else {
        litchi_ops::open_bytes(bytes).expect("source fixture")
    }
}

fn reopen(file: &File, eager: bool) -> File {
    open(eager, file.to_bytes().expect("publish"))
}

fn property(file: &File, path: &str, key: &str) -> Option<String> {
    file.get(path)
        .unwrap_or_else(|error| panic!("{path}: {error}"))
        .view()
        .format()
        .unwrap_or_else(|error| panic!("{path} format: {error}"))
        .find(|property| property.key() == key)
        .map(|property| property.to_string())
}

fn paths(file: &File, selector: &str) -> Vec<String> {
    file.query(selector)
        .unwrap_or_else(|error| panic!("{selector}: {error}"))
        .map(|node| {
            node.unwrap_or_else(|error| panic!("{selector}: {error}"))
                .view()
                .path()
                .to_string()
        })
        .collect()
}

#[test]
fn tab_stops_are_read_set_added_and_removed() {
    for eager in [false, true] {
        let file = open(eager, fixture("poi/test-data/document/Bug54849.docx"));
        let tabs = paths(&file, "tab");
        assert!(!tabs.is_empty(), "the fixture has tab stops");
        let first = &tabs[0];
        assert!(first.contains("/tab[1]"), "{first}");
        assert!(property(&file, first, "pos").is_some());
        assert!(property(&file, first, "val").is_some());

        let file = file
            .set(
                first.as_str(),
                props! { "val" => "center", "leader" => "dot", "pos" => 4320 },
            )
            .unwrap_or_else(|error| panic!("tab set: {error}"));
        let file = reopen(&file, eager);
        assert_eq!(property(&file, first, "val").as_deref(), Some("center"));
        assert_eq!(property(&file, first, "leader").as_deref(), Some("dot"));
        assert_eq!(property(&file, first, "pos").as_deref(), Some("4320"));

        let paragraph = first.rsplit_once("/tab[").unwrap().0.to_owned();
        let count = paths(&file, &format!("{paragraph} > tab")).len();
        let added = file
            .add(
                paragraph.as_str(),
                NewNode::new("tab").prop("pos", 9360).prop("val", "right"),
            )
            .unwrap_or_else(|error| panic!("tab add: {error}"));
        assert_eq!(
            added.path.to_string(),
            format!("{paragraph}/tab[{}]", count + 1)
        );
        let file = added.into_file();
        assert_eq!(paths(&file, &format!("{paragraph} > tab")).len(), count + 1);
        let file = file.remove(first.as_str()).expect("tab remove");
        assert_eq!(paths(&file, &format!("{paragraph} > tab")).len(), count);
    }
}

#[test]
fn breaks_are_read_added_and_removed() {
    for eager in [false, true] {
        let file = open(eager, fixture("ooxml/docx/FancyFoot.docx"));
        let breaks = paths(&file, "break");
        assert!(!breaks.is_empty());
        assert_eq!(
            property(&file, &breaks[0], "breakType").as_deref(),
            Some("page")
        );

        let added = file
            .add(
                "/body/p[1]",
                NewNode::new("pagebreak").prop("type", "column"),
            )
            .unwrap_or_else(|error| panic!("break add: {error}"));
        let path = added.path.to_string();
        assert_eq!(path, "/body/p[1]/break[1]");
        let file = reopen(&added.into_file(), eager);
        assert_eq!(
            property(&file, &path, "breakType").as_deref(),
            Some("column")
        );
        let before = paths(&file, "break").len();
        let file = file.remove(path.as_str()).expect("break remove");
        assert_eq!(paths(&file, "break").len(), before - 1);
    }
}

#[test]
fn bookmarks_fields_and_tables_of_contents_are_read() {
    for eager in [false, true] {
        let file = open(
            eager,
            fixture("libreoffice-core/sw/qa/extras/ooxmlexport/data/lvlPicBulletId.docx"),
        );
        let bookmarks = paths(&file, "bookmark");
        assert!(!bookmarks.is_empty());
        let name = property(&file, &bookmarks[0], "name").expect("bookmark name");
        let by_name = file
            .get(format!("/bookmark[@name={name}]").as_str())
            .expect("by name");
        assert_eq!(by_name.view().path().to_string(), bookmarks[0]);
        assert!(property(&file, &bookmarks[0], "id").is_some());

        let file = open(
            eager,
            fixture("libreoffice-core/sw/qa/extras/ooxmlexport/data/multi-page-toc.docx"),
        );
        let fields = paths(&file, "field");
        assert!(!fields.is_empty());
        assert!(property(&file, &fields[0], "instruction").is_some());
        let tocs = paths(&file, "toc");
        assert_eq!(tocs[0], "/toc[1]");
        assert_eq!(
            property(&file, "/toc[1]", "pageNumbers").as_deref(),
            Some("true")
        );
        assert!(property(&file, "/toc[1]", "levels").is_some());

        // The root view still lists only the body and the parts.
        let root = file.get("/").expect("root");
        for child in root.view().children() {
            let path = child.expect("child").view().path().to_string();
            assert!(
                !path.starts_with("/field")
                    && !path.starts_with("/toc")
                    && !path.starts_with("/bookmark"),
                "{path}"
            );
        }
    }
}

#[test]
fn the_watermark_and_text_boxes_are_read() {
    let mut builder = litchi_docx::Builder::new().expect("builder");
    builder
        .document()
        .expect("document")
        .add_paragraph_with_text("Body");
    builder
        .document()
        .expect("document")
        .set_watermark(litchi_docx::Watermark::text("DRAFT"));
    let mut package = builder.build().expect("build");
    let bytes = package.to_bytes().expect("bytes");
    for eager in [false, true] {
        let file = open(eager, bytes.clone());
        let watermark = file.get("/watermark").expect("watermark");
        assert_eq!(
            watermark.view().text().expect("text").as_deref(),
            Some("DRAFT")
        );
        assert_eq!(
            property(&file, "/watermark", "text").as_deref(),
            Some("DRAFT")
        );

        let file = open(
            eager,
            fixture("libreoffice-core/sw/qa/extras/ooxmlexport/data/textWatermark.docx"),
        );
        assert_eq!(
            property(&file, "/watermark", "text").as_deref(),
            Some("Airplane")
        );

        let file = open(
            eager,
            fixture(
                "libreoffice-core/sw/qa/core/objectpositioning/data/do-not-capture-draw-objs-on-page-draw-wrap-none.docx",
            ),
        );
        let boxes = paths(&file, "textbox");
        assert_eq!(boxes.first().map(String::as_str), Some("/body/textbox[1]"));
        let text_box = file.get("/body/textbox[1]").expect("text box");
        assert_eq!(
            text_box.view().text().expect("text").as_deref(),
            Some("<Text>")
        );
    }
}
