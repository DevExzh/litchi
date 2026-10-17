//! DOC parts and marks through the unified API (record 0324): styles,
//! numbering, paragraph style and list keys, hyperlinks, bookmarks, fields,
//! tables of contents, sections, tab stops, text boxes and shapes. Reads
//! answer the DOCX paths and keys; writes round-trip through a reopen, equal
//! values are exact no-ops, and the generic views list what they listed.
#![cfg(feature = "doc")]
#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    clippy::indexing_slicing,
    reason = "integration tests favor concise panic-driven assertions"
)]

use litchi_ops::{Code, File, NewNode, Placement, open_bytes, props};

fn fixture(name: &str) -> Vec<u8> {
    std::fs::read(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data")
            .join(name),
    )
    .unwrap()
}

fn open(name: &str) -> File {
    open_bytes(fixture(name)).unwrap()
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

fn text(file: &File, path: &str) -> String {
    file.get(path)
        .unwrap_or_else(|error| panic!("{path}: {error}"))
        .view()
        .text()
        .unwrap()
        .unwrap_or_default()
        .into_owned()
}

fn count(file: &File, query: &str) -> usize {
    file.query(query).unwrap().count()
}

fn paths(file: &File, query: &str) -> Vec<String> {
    file.query(query)
        .unwrap()
        .map(|node| node.unwrap().view().path().to_string())
        .collect()
}

fn reopen(file: &File) -> File {
    open_bytes(file.to_bytes().unwrap()).unwrap()
}

fn code(result: litchi_ops::Result<File>) -> Code {
    match result {
        Ok(_) => panic!("the write was admitted"),
        Err(error) => error.code(),
    }
}

/// Every DOC fixture of the tracked corpus.
fn corpus() -> Vec<String> {
    let mut names = Vec::new();
    for directory in ["ole/doc", "poi/test-data/document", "poi/test-data/hpsf"] {
        let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data")
            .join(directory);
        for entry in std::fs::read_dir(root).unwrap() {
            let name = entry.unwrap().file_name().into_string().unwrap();
            if name.ends_with(".doc") {
                names.push(format!("{directory}/{name}"));
            }
        }
    }
    names.sort();
    names
}

#[test]
fn styles_answer_the_docx_style_keys() {
    let file = open("ole/doc/Lists.doc");
    assert_eq!(prop(&file, "/styles", "count").as_deref(), Some("7"));
    assert_eq!(
        paths(&file, "style"),
        [
            "/styles/Normal",
            "/styles/Heading1",
            "/styles/DefaultParagraphFont",
            "/styles/TableNormal",
            "/styles/NoList",
            "/styles/Heading1Char",
            "/styles/ListParagraph",
        ]
    );
    let heading = "/styles/Heading1";
    for (key, value) in [
        ("id", "Heading1"),
        ("name", "Heading 1"),
        ("type", "paragraph"),
        ("basedOn", "Normal"),
        ("next", "Normal"),
        ("linked", "Heading1Char"),
        ("qFormat", "true"),
        ("uiPriority", "9"),
        ("font", "Cambria"),
        ("size", "14pt"),
        ("bold", "true"),
        ("color", "#365F91"),
        ("spaceBefore", "24pt"),
        ("outlineLvl", "0"),
        ("keepNext", "true"),
    ] {
        assert_eq!(prop(&file, heading, key).as_deref(), Some(value), "{key}");
    }
    assert_eq!(
        prop(&file, "/styles/Normal", "default").as_deref(),
        Some("true")
    );
    assert_eq!(
        prop(&file, "/styles/Heading1Char", "customStyle").as_deref(),
        Some("true")
    );
    // A style is also found by its display name or by a query predicate.
    assert_eq!(text(&file, "/styles/ListParagraph"), "");
    assert_eq!(count(&file, "style[type=character]"), 2);
    assert_eq!(
        file.get("/styles/Nope").unwrap_err().code(),
        Code::PathNotFound
    );
}

#[test]
fn numbering_and_paragraph_lists_answer_the_docx_keys() {
    let file = open("ole/doc/Lists.doc");
    assert_eq!(
        prop(&file, "/numbering", "abstractNumCount").as_deref(),
        Some("5")
    );
    assert_eq!(prop(&file, "/numbering", "numCount").as_deref(), Some("5"));
    assert_eq!(
        prop(&file, "/numbering/abstractNum[1]", "type").as_deref(),
        Some("hybridMultilevel")
    );
    let level = "/numbering/abstractNum[4]/level[2]";
    for (key, value) in [
        ("start", "1"),
        ("format", "decimal"),
        ("lvlText", "%1.%2."),
        ("justification", "left"),
        ("suff", "tab"),
        ("indent", "792"),
        ("hanging", "432"),
    ] {
        assert_eq!(prop(&file, level, key).as_deref(), Some(value), "{key}");
    }
    assert_eq!(
        prop(&file, "/numbering/num[2]", "abstractNumId").as_deref(),
        Some("2")
    );
    // The native identities are raw: absent from the ordinary format.
    assert_eq!(prop(&file, "/numbering/abstractNum[1]", "id"), None);
    assert_eq!(prop(&file, level, "ilvl"), None);
    assert_eq!(count(&file, "abstractNum"), 5);
    assert_eq!(count(&file, "/numbering/abstractNum[1] > level"), 9);

    for (key, value) in [
        ("styleId", "ListParagraph"),
        ("styleName", "List Paragraph"),
        ("numFmt", "decimal"),
        ("listStyle", "ordered"),
        ("start", "1"),
        ("numLevel", "1"),
    ] {
        assert_eq!(
            prop(&file, "/body/p[28]", key).as_deref(),
            Some(value),
            "{key}"
        );
    }
    assert_eq!(
        prop(&file, "/body/p[3]", "listStyle").as_deref(),
        Some("bullet")
    );
    // A Normal paragraph has no style key, as DOCX has no `w:pStyle` there.
    assert_eq!(prop(&file, "/body/p[2]", "styleId"), None);
    assert_eq!(prop(&file, "/body/p[2]", "numFmt"), None);
}

#[test]
fn paragraph_and_run_styles_are_set_by_id_or_name() {
    let file = open("ole/doc/Lists.doc");
    let by_name = file
        .set("/body/p[2]", props! { "styleName" => "Heading 1" })
        .unwrap();
    let by_id = file
        .set("/body/p[2]", props! { "style" => "Heading1" })
        .unwrap();
    assert_eq!(by_name.to_bytes().unwrap(), by_id.to_bytes().unwrap());
    let reopened = reopen(&by_id);
    assert_eq!(
        prop(&reopened, "/body/p[2]", "styleId").as_deref(),
        Some("Heading1")
    );
    assert_eq!(text(&reopened, "/body/p[2]"), text(&file, "/body/p[2]"));
    // Setting the style a paragraph has publishes the source bytes.
    let same = file
        .set("/body/p[3]", props! { "styleId" => "ListParagraph" })
        .unwrap();
    assert_eq!(same.to_bytes().unwrap(), file.to_bytes().unwrap());

    let run = reopen(
        &file
            .set("/body/p[2]/r[1]", props! { "rStyle" => "Heading1Char" })
            .unwrap(),
    );
    assert_eq!(
        prop(&run, "/body/p[2]/r[1]", "rStyle").as_deref(),
        Some("Heading1Char")
    );
    let cleared = reopen(
        &run.set(
            "/body/p[2]/r[1]",
            props! { "rStyle" => "Default Paragraph Font" },
        )
        .unwrap(),
    );
    assert_eq!(prop(&cleared, "/body/p[2]/r[1]", "rStyle"), None);

    assert_eq!(
        code(file.set("/body/p[2]", props! { "style" => "Missing" })),
        Code::NotFound
    );
    assert_eq!(
        code(file.set("/body/p[2]", props! { "style" => "Heading1Char" })),
        Code::InvalidValue
    );
    assert_eq!(
        code(file.set("/body/p[2]/r[1]", props! { "rStyle" => "Heading1" })),
        Code::InvalidValue
    );
    assert_eq!(
        code(file.set("/body/p[2]/r[1]", props! { "style" => "Heading1" })),
        Code::UnsupportedCapability
    );
}

#[test]
fn hyperlinks_are_read_set_added_and_removed() {
    let file = open("ole/doc/hyperlink.doc");
    assert_eq!(paths(&file, "hyperlink"), ["/body/p[1]/hyperlink[1]"]);
    let link = "/body/p[1]/hyperlink[1]";
    assert_eq!(text(&file, link), "Hyperlink text");
    assert_eq!(
        prop(&file, link, "url").as_deref(),
        Some("http://testuri.org/")
    );
    assert_eq!(prop(&file, link, "underline").as_deref(), Some("single"));
    assert_eq!(prop(&file, link, "color").as_deref(), Some("#000080"));

    let file = open("ole/doc/Lists.doc");
    let added = file
        .add(
            "/body/p[2]",
            NewNode::new("hyperlink")
                .prop("url", "http://example.com/")
                .prop("text", "Example")
                .prop("tooltip", "Tip"),
        )
        .unwrap();
    assert_eq!(added.path.to_string(), "/body/p[2]/hyperlink[1]");
    let linked = reopen(&added.file);
    let link = "/body/p[2]/hyperlink[1]";
    assert_eq!(text(&linked, link), "Example");
    assert_eq!(
        prop(&linked, link, "url").as_deref(),
        Some("http://example.com/")
    );
    assert_eq!(prop(&linked, link, "tooltip").as_deref(), Some("Tip"));
    assert_eq!(prop(&linked, link, "underline").as_deref(), Some("single"));
    assert_eq!(prop(&linked, link, "color").as_deref(), Some("#0563C1"));
    // A second link before the first run is the paragraph's first.
    let first = linked
        .add(
            "/body/p[2]",
            NewNode::new("hyperlink")
                .prop("anchor", "Top")
                .at(Placement::Start),
        )
        .unwrap();
    assert_eq!(first.path.to_string(), "/body/p[2]/hyperlink[1]");
    let both = reopen(&first.file);
    assert_eq!(count(&both, "/body/p[2] > hyperlink"), 2);
    assert_eq!(
        prop(&both, "/body/p[2]/hyperlink[1]", "anchor").as_deref(),
        Some("Top")
    );
    assert_eq!(text(&both, "/body/p[2]/hyperlink[1]"), "Top");

    let changed = reopen(
        &linked
            .set(
                link,
                props! { "url" => "http://example.org/", "text" => "Other" },
            )
            .unwrap(),
    );
    assert_eq!(
        prop(&changed, link, "url").as_deref(),
        Some("http://example.org/")
    );
    assert_eq!(prop(&changed, link, "tooltip").as_deref(), Some("Tip"));
    assert_eq!(text(&changed, link), "Other");
    let same = linked.set(link, props! { "text" => "Example" }).unwrap();
    assert_eq!(same.to_bytes().unwrap(), linked.to_bytes().unwrap());

    let removed = reopen(&changed.remove(link).unwrap());
    assert_eq!(count(&removed, "hyperlink"), 0);
    assert_eq!(text(&removed, "/body/p[2]"), text(&file, "/body/p[2]"));

    assert_eq!(
        code(
            file.add("/body", NewNode::new("hyperlink").prop("url", "x"))
                .map(|added| added.file)
        ),
        Code::UnsupportedType
    );
    assert_eq!(
        code(
            file.add("/body/p[2]", NewNode::new("hyperlink").prop("text", "x"))
                .map(|added| added.file)
        ),
        Code::InvalidArgument
    );
    assert_eq!(
        code(linked.set(link, props! { "bold" => true })),
        Code::UnsupportedCapability
    );
}

#[test]
fn bookmarks_fields_and_tables_of_contents_are_read() {
    let file = open("ole/doc/watermark.doc");
    assert_eq!(count(&file, "bookmark"), 8);
    assert_eq!(
        prop(&file, "/bookmark[1]", "name").as_deref(),
        Some("_Ref189041054")
    );
    assert_eq!(
        prop(&file, "/bookmark[@name=__RefHeading__8_476954814]", "name").as_deref(),
        Some("__RefHeading__8_476954814")
    );
    assert_eq!(
        prop(&file, "/field[1]", "fieldType").as_deref(),
        Some("toc")
    );
    assert_eq!(
        prop(&file, "/field[2]", "fieldType").as_deref(),
        Some("hyperlink")
    );
    assert_eq!(
        prop(&file, "/field[2]", "evaluated").as_deref(),
        Some("true")
    );
    assert_eq!(text(&file, "/field[2]"), "1. TERM\t1");
    assert_eq!(count(&file, "toc"), 1);
    assert_eq!(
        prop(&file, "/toc[1]", "pageNumbers").as_deref(),
        Some("true")
    );
    // The TOC's text is its result without the nested instructions.
    assert!(text(&file, "/toc[1]").starts_with("1. TERM\t1"));
    assert!(!text(&file, "/toc[1]").contains("HYPERLINK"));
}

#[test]
fn bookmarks_are_added_renamed_and_removed() {
    let file = open("ole/doc/Lists.doc");
    let added = file
        .add(
            "/body/p[2]",
            NewNode::new("bookmark")
                .prop("name", "Mark1")
                .prop("text", "document"),
        )
        .unwrap();
    assert_eq!(added.path.to_string(), "/bookmark[1]");
    let marked = reopen(&added.file);
    assert_eq!(
        prop(&marked, "/bookmark[@name=Mark1]", "name").as_deref(),
        Some("Mark1")
    );
    let body = file
        .add(
            "/body",
            NewNode::new("bookmark")
                .prop("name", "Intro")
                .prop("text", "Introduction"),
        )
        .unwrap();
    let body = reopen(&body.file);
    assert_eq!(text(&body, "/body/p[41]"), "Introduction");
    assert_eq!(count(&body, "bookmark[name=Intro]"), 1);

    let renamed = reopen(
        &marked
            .set("/bookmark[1]", props! { "name" => "Mark2" })
            .unwrap(),
    );
    assert_eq!(paths(&renamed, "bookmark"), ["/bookmark[1]"]);
    assert_eq!(
        prop(&renamed, "/bookmark[1]", "name").as_deref(),
        Some("Mark2")
    );
    let removed = reopen(&renamed.remove("/bookmark[@name=Mark2]").unwrap());
    assert_eq!(count(&removed, "bookmark"), 0);
    assert_eq!(text(&removed, "/body/p[2]"), text(&file, "/body/p[2]"));

    assert_eq!(
        code(
            marked
                .add("/body/p[3]", NewNode::new("bookmark").prop("name", "mark1"))
                .map(|added| added.file)
        ),
        Code::InvalidValue
    );
    assert_eq!(
        code(
            file.add(
                "/body/p[2]",
                NewNode::new("bookmark").prop("name", "bad name")
            )
            .map(|added| added.file)
        ),
        Code::InvalidValue
    );
}

#[test]
fn sections_read_and_set_page_setup() {
    let file = open("poi/test-data/document/Bug53453Section.doc");
    assert_eq!(count(&file, "section"), 2);
    assert_eq!(
        prop(&file, "/section[2]", "type").as_deref(),
        Some("continuous")
    );
    assert_eq!(prop(&file, "/section[2]", "columns").as_deref(), Some("3"));
    assert_eq!(
        prop(&file, "/section[1]", "orientation").as_deref(),
        Some("portrait")
    );
    let set = reopen(
        &file
            .set(
                "/section[1]",
                props! {
                    "orientation" => "landscape",
                    "pageWidth" => "29.7cm",
                    "pageHeight" => "21cm",
                    "marginLeft" => "2cm",
                    "columns" => 2,
                    "titlePage" => true,
                },
            )
            .unwrap(),
    );
    for (key, value) in [
        ("orientation", "landscape"),
        // 29.7cm, 21cm and 2cm round to whole twips, which print in points.
        ("pageWidth", "841.9pt"),
        ("pageHeight", "595.3pt"),
        ("marginLeft", "56.7pt"),
        ("columns", "2"),
        ("titlePage", "true"),
    ] {
        assert_eq!(
            prop(&set, "/section[1]", key).as_deref(),
            Some(value),
            "{key}"
        );
    }
    assert_eq!(
        prop(&set, "/section[2]", "columns").as_deref(),
        Some("3"),
        "the other section is unchanged"
    );
    let same = file.set("/section[2]", props! { "columns" => 3 }).unwrap();
    assert_eq!(same.to_bytes().unwrap(), file.to_bytes().unwrap());
    assert_eq!(
        code(file.set("/section[1]", props! { "pageNumFmt" => "upperRoman" })),
        Code::UnsupportedCapability
    );
    assert_eq!(
        code(file.set("/section[1]", props! { "columns" => 0 })),
        Code::InvalidValue
    );
}

#[test]
fn tab_stops_are_read_added_set_and_removed() {
    let file = open("ole/doc/HeaderFooterUnicode.doc");
    assert_eq!(
        prop(&file, "/header[1]/p[1]/tab[1]", "pos").as_deref(),
        Some("4680")
    );
    assert_eq!(
        prop(&file, "/header[1]/p[1]/tab[2]", "val").as_deref(),
        Some("right")
    );

    let file = open("ole/doc/Lists.doc");
    let added = file
        .add(
            "/body/p[2]",
            NewNode::new("tab")
                .prop("pos", 2880)
                .prop("val", "right")
                .prop("leader", "dot"),
        )
        .unwrap();
    assert_eq!(added.path.to_string(), "/body/p[2]/tab[1]");
    let second = added
        .file
        .add("/body/p[2]", NewNode::new("tab").prop("pos", 1440))
        .unwrap();
    assert_eq!(second.path.to_string(), "/body/p[2]/tab[1]");
    let tabs = reopen(&second.file);
    assert_eq!(paths(&tabs, "/body/p[2] > tab").len(), 2);
    assert_eq!(
        prop(&tabs, "/body/p[2]/tab[2]", "pos").as_deref(),
        Some("2880")
    );
    assert_eq!(
        prop(&tabs, "/body/p[2]/tab[2]", "leader").as_deref(),
        Some("dot")
    );
    assert_eq!(
        prop(&tabs, "/body/p[2]/tab[1]", "val").as_deref(),
        Some("left")
    );

    let set = reopen(
        &tabs
            .set(
                "/body/p[2]/tab[1]",
                props! { "val" => "center", "pos" => 1000 },
            )
            .unwrap(),
    );
    assert_eq!(
        prop(&set, "/body/p[2]/tab[1]", "pos").as_deref(),
        Some("1000")
    );
    assert_eq!(
        prop(&set, "/body/p[2]/tab[1]", "val").as_deref(),
        Some("center")
    );
    let removed = reopen(&set.remove("/body/p[2]/tab[2]").unwrap());
    assert_eq!(paths(&removed, "/body/p[2] > tab"), ["/body/p[2]/tab[1]"]);
    assert_eq!(text(&removed, "/body/p[2]"), text(&file, "/body/p[2]"));

    assert_eq!(
        code(
            tabs.add("/body/p[2]", NewNode::new("tab").prop("pos", 2880))
                .map(|added| added.file)
        ),
        Code::DuplicateName
    );
    assert_eq!(
        code(
            file.add("/body/p[2]", NewNode::new("tab").prop("val", "left"))
                .map(|added| added.file)
        ),
        Code::InvalidArgument
    );
}

#[test]
fn text_boxes_and_shapes_are_read() {
    use litchi_doc::writer::{FloatingPosition, Kind as DrawingKind, Shape as DrawingShape};
    let mut builder = litchi_doc::Builder::new();
    builder.add_paragraph("before").unwrap();
    builder
        .insert_floating_text_box(
            DrawingShape::new(DrawingKind::Rectangle, 3600, 1800).unwrap(),
            FloatingPosition::new(2500, 1200),
            "Boxed text",
        )
        .unwrap();
    builder.add_paragraph("after").unwrap();
    let mut bytes = std::io::Cursor::new(Vec::new());
    builder.build().unwrap().write_to(&mut bytes).unwrap();
    let file = open_bytes(bytes.into_inner()).unwrap();
    assert_eq!(paths(&file, "textbox"), ["/body/textbox[1]"]);
    assert_eq!(text(&file, "/body/textbox[1]"), "Boxed text");
    assert_eq!(
        prop(&file, "/body/textbox[1]", "width").as_deref(),
        Some("6.35cm")
    );
    assert_eq!(
        prop(&file, "/body/textbox[1]", "height").as_deref(),
        Some("90pt")
    );

    let file = open("ole/doc/FloatingPictures.doc");
    assert_eq!(count(&file, "shape"), 4);
    for (key, value) in [
        ("width", "75pt"),
        ("height", "56.25pt"),
        ("anchor.x", "153pt"),
        ("wrap", "square"),
        ("hRelative", "column"),
        ("vRelative", "paragraph"),
    ] {
        assert_eq!(
            prop(&file, "/body/shape[1]", key).as_deref(),
            Some(value),
            "{key}"
        );
    }
    assert_eq!(
        prop(&file, "/body/shape[2]", "behindDoc").as_deref(),
        Some("true")
    );
}

/// The parts and marks are reached by path and query; the generic views of
/// the root, the body and a paragraph list what they listed before.
#[test]
fn views_list_only_the_structural_children() {
    for name in [
        "ole/doc/Lists.doc",
        "ole/doc/watermark.doc",
        "ole/doc/hyperlink.doc",
    ] {
        let file = open(name);
        for path in ["/", "/body", "/body/p[1]"] {
            let node = file.get(path).unwrap();
            let view = node.view();
            for child in view.children() {
                let child = child.unwrap();
                assert!(
                    matches!(
                        child.view().kind(),
                        "body"
                            | "header"
                            | "footer"
                            | "footnote"
                            | "endnote"
                            | "comments"
                            | "paragraph"
                            | "table"
                            | "run"
                    ),
                    "{name} {path}: {}",
                    child.view().kind()
                );
            }
        }
    }
}

/// Over the corpus, every new kind is readable or a typed error, and setting
/// each body paragraph's own style is an exact no-op.
#[test]
fn the_corpus_reads_every_part_and_mark() {
    let mut opened = 0usize;
    for name in corpus() {
        let Ok(file) = open_bytes(fixture(&name)) else {
            continue;
        };
        opened += 1;
        for query in [
            "style",
            "abstractNum",
            "num",
            "level",
            "hyperlink",
            "bookmark",
            "field",
            "toc",
            "section",
            "tab",
            "textbox",
            "shape",
        ] {
            let Ok(matches) = file.query(query) else {
                continue;
            };
            for node in matches {
                let Ok(node) = node else {
                    break;
                };
                let view = node.view();
                let _ = view.format().map(Iterator::count);
                let _ = view.text();
                let _ = view.path().to_string();
            }
        }
        if let Ok(id) = file.get("/body/p[1]").map(|node| {
            node.view()
                .format()
                .ok()
                .and_then(|mut format| format.find(|prop| prop.key() == "styleId"))
                .map(|prop| prop.to_string())
        }) && let Some(id) = id
            && let Ok(same) = file.set("/body/p[1]", props! { "styleId" => id.as_str() })
        {
            assert_eq!(same.to_bytes().unwrap(), file.to_bytes().unwrap(), "{name}");
        }
    }
    assert!(opened > 40, "{opened}");
}

#[test]
#[ignore = "writes LibreOffice probe files to DOC_WRITE_DIR"]
fn libreoffice_probe_files() {
    let dir = std::env::var("DOC_WRITE_DIR").unwrap();
    let write = |name: &str, file: &File| {
        std::fs::write(format!("{dir}/{name}.doc"), file.to_bytes().unwrap()).unwrap();
    };
    let equation = open("ole/doc/equation.doc");
    let linked = equation
        .add(
            "/body/p[1]",
            NewNode::new("hyperlink")
                .prop("url", "http://example.com/")
                .prop("text", "probe"),
        )
        .unwrap()
        .file;
    write("equation-link", &linked);
    let toc = open("ole/doc/watermark.doc");
    let first = paths(&toc, "hyperlink").remove(0);
    let retitled = match toc.set(first.as_str(), props! { "text" => "Retitled entry" }) {
        Ok(file) => file,
        Err(error) => {
            let mut source: Option<&dyn std::error::Error> = Some(&error);
            while let Some(current) = source {
                println!("error: {current}");
                source = current.source();
            }
            toc.clone()
        },
    };
    println!("{first} -> {}", text(&reopen(&retitled), first.as_str()));
    write("toc-retitled", &retitled);
    let lists = open("ole/doc/Lists.doc");
    let edited = lists
        .set("/body/p[2]", props! { "style" => "ListParagraph" })
        .unwrap()
        .add(
            "/body/p[2]",
            NewNode::new("tab")
                .prop("pos", 4320)
                .prop("val", "right")
                .prop("leader", "dot"),
        )
        .unwrap()
        .file
        .add(
            "/body/p[2]",
            NewNode::new("bookmark")
                .prop("name", "Probe")
                .prop("text", "different"),
        )
        .unwrap()
        .file
        .set(
            "/section[1]",
            props! { "marginLeft" => "1in", "columns" => 2 },
        )
        .unwrap();
    write("lists-edited", &edited);
}
