//! DOCX styles, numbering, headers, footers, footnotes and endnotes through
//! the format-neutral `File` facade (record 0319): every element is read on
//! both owners, written, read back and reopened.
#![cfg(feature = "docx")]

use std::path::PathBuf;

use litchi_ops::{Code, File, NewNode, OpenOptions, SourceMode, props};

fn fixture(name: &str) -> Vec<u8> {
    let path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/ooxml/docx")
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

fn text(file: &File, path: &str) -> String {
    file.get(path)
        .unwrap_or_else(|error| panic!("{path}: {error}"))
        .view()
        .text()
        .unwrap_or_else(|error| panic!("{path} text: {error}"))
        .map(std::borrow::Cow::into_owned)
        .unwrap_or_default()
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
fn styles_are_read_by_id_and_name_and_written() {
    for eager in [false, true] {
        let file = open(eager, fixture("FancyFoot.docx"));
        let styles = paths(&file, "style");
        assert!(
            styles.iter().any(|path| path == "/styles/Normal"),
            "{styles:?}"
        );
        assert_eq!(
            property(&file, "/styles/Normal", "type").as_deref(),
            Some("paragraph")
        );
        assert_eq!(
            property(&file, "/styles/Normal", "id").as_deref(),
            Some("Normal")
        );
        assert_eq!(
            property(&file, "/styles/Normal", "default").as_deref(),
            Some("true")
        );
        // A display name resolves to the same style.
        let heading = file.get("/styles/heading 1").expect("by display name");
        assert_eq!(heading.view().path().to_string(), "/styles/Heading1");
        assert!(property(&file, "/styles", "count").is_some());

        let file = file
            .set(
                "/styles/Normal",
                props! { "bold" => true, "spaceAfter" => "6pt", "qFormat" => true },
            )
            .unwrap_or_else(|error| panic!("style set: {error}"));
        assert_eq!(
            property(&file, "/styles/Normal", "bold").as_deref(),
            Some("true")
        );
        assert_eq!(
            property(&file, "/styles/Normal", "spaceAfter").as_deref(),
            Some("6pt")
        );
        assert_eq!(
            property(&file, "/styles/Normal", "qFormat").as_deref(),
            Some("true")
        );

        let added = file
            .add(
                "/styles",
                NewNode::new("style")
                    .prop("id", "Accent")
                    .prop("name", "Accent Text")
                    .prop("type", "character")
                    .prop("italic", true)
                    .prop("basedOn", "Default Paragraph Font"),
            )
            .unwrap_or_else(|error| panic!("style add: {error}"));
        assert_eq!(added.path.to_string(), "/styles/Accent");
        let file = reopen(&added.into_file(), eager);
        assert_eq!(
            property(&file, "/styles/Accent", "name").as_deref(),
            Some("Accent Text")
        );
        assert_eq!(
            property(&file, "/styles/Accent", "italic").as_deref(),
            Some("true")
        );
        assert_eq!(
            property(&file, "/styles/Accent", "customStyle").as_deref(),
            Some("true")
        );
        assert_eq!(
            property(&file, "/styles/Accent", "basedOn").as_deref(),
            Some("DefaultParagraphFont")
        );
        let error = file
            .add("/styles", NewNode::new("style").prop("id", "Accent"))
            .expect_err("a taken identifier");
        assert_eq!(error.code(), Code::DuplicateName);

        // A paragraph style set by display name writes the identifier.
        let file = file
            .set("/body/p[1]", props! { "style" => "heading 1" })
            .unwrap_or_else(|error| panic!("paragraph style: {error}"));
        assert_eq!(
            property(&file, "/body/p[1]", "style").as_deref(),
            Some("Heading1")
        );
        assert_eq!(
            property(&file, "/body/p[1]", "styleName").as_deref(),
            Some("heading 1")
        );

        let removed = file.remove("/styles/Accent").expect("style remove");
        assert!(removed.get("/styles/Accent").is_err());
    }
}

#[test]
fn numbering_definitions_instances_and_levels_round_trip() {
    for eager in [false, true] {
        let file = open(eager, fixture("ComplexNumberedLists.docx"));
        let definitions = paths(&file, "abstractNum");
        assert!(!definitions.is_empty());
        assert_eq!(definitions[0], "/numbering/abstractNum[1]");
        let levels = paths(&file, "/numbering/abstractNum[1] > level");
        assert!(!levels.is_empty(), "{levels:?}");
        assert!(property(&file, "/numbering/abstractNum[1]/level[1]", "format").is_some());
        // Native identifiers are raw-path only.
        assert!(property(&file, "/numbering/abstractNum[1]/level[1]", "ilvl").is_none());
        let raw = litchi_ops::path::raw::path("/numbering/abstractNum[1]/level[@ilvl=0]").unwrap();
        assert_eq!(
            file.get(raw).expect("raw ilvl").view().path().to_string(),
            "/numbering/abstractNum[1]/level[1]"
        );
        let numbered = file
            .query("paragraph[numId]")
            .unwrap()
            .next()
            .expect("a numbered paragraph")
            .unwrap();
        let path = numbered.view().path().to_string();
        assert!(property(&file, &path, "numFmt").is_some(), "{path}");
        assert!(matches!(
            property(&file, &path, "listStyle").as_deref(),
            Some("bullet" | "ordered")
        ));

        let file = file
            .set(
                "/numbering/abstractNum[1]/level[1]",
                props! { "format" => "upperRoman", "lvlText" => "%1)", "indent" => 1080, "bold" => true },
            )
            .unwrap_or_else(|error| panic!("level set: {error}"));
        let level = "/numbering/abstractNum[1]/level[1]";
        assert_eq!(
            property(&file, level, "format").as_deref(),
            Some("upperRoman")
        );
        assert_eq!(property(&file, level, "lvlText").as_deref(), Some("%1)"));
        assert_eq!(property(&file, level, "indent").as_deref(), Some("1080"));
        assert_eq!(property(&file, level, "bold").as_deref(), Some("true"));

        let added = file
            .add(
                "/numbering",
                NewNode::new("abstractNum")
                    .prop("id", 77)
                    .prop("format", "bullet")
                    .prop("name", "Dots"),
            )
            .unwrap_or_else(|error| panic!("abstractNum add: {error}"));
        let definition = added.path.to_string();
        let file = added.into_file();
        assert_eq!(
            property(&file, &definition, "name").as_deref(),
            Some("Dots")
        );
        let raw = litchi_ops::path::raw::path("/numbering/abstractNum[@id=77]").unwrap();
        assert_eq!(
            file.get(raw).expect("raw id").view().path().to_string(),
            definition
        );
        let id = "77".to_owned();
        let added = file
            .add(
                "/numbering",
                NewNode::new("num").prop("abstractNumId", id.parse::<i64>().unwrap()),
            )
            .unwrap_or_else(|error| panic!("num add: {error}"));
        let num = added.path.to_string();
        let file = reopen(&added.into_file(), eager);
        assert_eq!(
            property(&file, &num, "abstractNumId").as_deref(),
            Some(id.as_str())
        );
        assert_eq!(
            property(&file, &num, "startOverride.0").as_deref(),
            Some("1")
        );
        assert_eq!(text(&file, &format!("{definition}/level[2]")), "");
        assert_eq!(
            property(&file, &format!("{definition}/level[2]"), "lvlText").as_deref(),
            Some("\u{25E6}")
        );

        // Removing a definition removes its instances.
        let file = file.remove(definition.as_str()).expect("definition remove");
        assert!(file.get(num.as_str()).is_err());
    }
}

#[test]
fn level_keys_build_definitions_and_instances_in_one_call() {
    // Record 0325: OfficeCLI's `level<N>.*` keys ("Mode A"), as
    // numbering.sh lines 46, 160, 200 and 222 use them.
    let file = litchi_ops::create(litchi_core::FileFormat::Docx).unwrap();
    let added = file
        .add(
            "/numbering",
            NewNode::new("abstractNum")
                .prop("id", 100)
                .prop("type", "hybridMultilevel")
                .prop("level0.format", "decimal")
                .prop("level0.text", "%1.")
                .prop("level0.bold", true)
                .prop("level0.color", "FF0000")
                .prop("level0.size", 14)
                .prop("level1.format", "lowerLetter")
                .prop("level1.text", "%2)")
                .prop("level1.italic", true)
                .prop("level1.indent", 1440)
                .prop("level2.format", "bullet")
                .prop("level2.font", "Arial")
                .prop("level2.hanging", 300)
                .prop("level2.justification", "right")
                .prop("level2.suff", "space"),
        )
        .unwrap_or_else(|error| panic!("abstractNum add: {error}"));
    let definition = added.path.to_string();
    for eager in [false, true] {
        let file = reopen(&added.file, eager);
        let level = |n: usize, key: &str| property(&file, &format!("{definition}/level[{n}]"), key);
        assert_eq!(level(1, "format").as_deref(), Some("decimal"));
        assert_eq!(level(1, "bold").as_deref(), Some("true"));
        assert_eq!(level(1, "color").as_deref(), Some("#FF0000"));
        assert_eq!(level(2, "lvlText").as_deref(), Some("%2)"));
        assert_eq!(level(2, "indent").as_deref(), Some("1440"));
        assert_eq!(level(2, "hanging").as_deref(), Some("360"));
        assert_eq!(level(2, "italic").as_deref(), Some("true"));
        assert_eq!(level(3, "format").as_deref(), Some("bullet"));
        assert_eq!(level(3, "lvlText").as_deref(), Some("\u{25AA}"));
        assert_eq!(level(3, "indent").as_deref(), Some("2160"));
        assert_eq!(level(3, "hanging").as_deref(), Some("300"));
        assert_eq!(level(3, "justification").as_deref(), Some("right"));
        assert_eq!(level(3, "suff").as_deref(), Some("space"));
        // Levels without keys keep the decimal, letter, Roman cycle.
        assert_eq!(level(4, "format").as_deref(), Some("decimal"));
        assert_eq!(level(4, "lvlText").as_deref(), Some("%4."));
    }
    // A bare level-0 key wins over its `level0.*` spelling.
    let bare = file
        .add(
            "/numbering",
            NewNode::new("abstractNum")
                .prop("format", "upperRoman")
                .prop("level0.format", "decimal")
                .prop("level1.start", 3),
        )
        .unwrap();
    let path = bare.path.to_string();
    assert_eq!(
        property(&bare.file, &format!("{path}/level[1]"), "format").as_deref(),
        Some("upperRoman")
    );
    assert_eq!(
        property(&bare.file, &format!("{path}/level[2]"), "start").as_deref(),
        Some("3")
    );
    // `add num` without abstractNumId creates the definition too.
    let num = added
        .file
        .add(
            "/numbering",
            NewNode::new("num")
                .prop("level0.format", "upperRoman")
                .prop("level0.text", "%1.")
                .prop("level0.bold", true)
                .prop("start", 4),
        )
        .unwrap_or_else(|error| panic!("num Mode A: {error}"));
    let path = num.path.to_string();
    let file = reopen(&num.file, true);
    assert_eq!(
        property(&file, &path, "startOverride.0").as_deref(),
        Some("4")
    );
    let before = paths(&added.file, "abstractNum").len();
    let definitions = paths(&file, "abstractNum");
    assert_eq!(definitions.len(), before + 1);
    let created = definitions.last().unwrap();
    assert_eq!(
        property(&file, &format!("{created}/level[1]"), "format").as_deref(),
        Some("upperRoman")
    );
    assert_eq!(
        property(&file, &format!("{created}/level[1]"), "start").as_deref(),
        Some("4")
    );
    // Naming a definition and describing one is refused.
    let error = added
        .file
        .add(
            "/numbering",
            NewNode::new("num")
                .prop("abstractNumId", 100)
                .prop("level0.format", "decimal"),
        )
        .unwrap_err();
    assert_eq!(error.code(), Code::InvalidArgument);
    let error = added
        .file
        .add(
            "/numbering",
            NewNode::new("abstractNum").prop("level9.format", "decimal"),
        )
        .unwrap_err();
    assert_eq!(error.code(), Code::UnsupportedProperty);
}

#[test]
fn headers_and_footers_are_read_written_added_and_removed() {
    for eager in [false, true] {
        let file = open(eager, fixture("headerFooter.docx"));
        let headers = paths(&file, "header");
        assert_eq!(headers.len(), 3, "{headers:?}");
        let types: Vec<Option<String>> = headers
            .iter()
            .map(|path| property(&file, path, "type"))
            .collect();
        assert!(types.iter().all(Option::is_some), "{types:?}");
        let by_type = file.get("/header[@type=default]").expect("default header");
        let default_path = by_type.view().path().to_string();

        let file = file
            .set(
                default_path.as_str(),
                props! { "text" => "Replaced", "align" => "center", "bold" => true },
            )
            .unwrap_or_else(|error| panic!("header set: {error}"));
        assert_eq!(text(&file, &default_path), "Replaced");
        assert_eq!(
            property(&file, &default_path, "align").as_deref(),
            Some("center")
        );
        assert_eq!(
            property(&file, &default_path, "bold").as_deref(),
            Some("true")
        );
        // An unscoped paragraph query stays in the main story.
        let main = paths(&open(eager, fixture("headerFooter.docx")), "p").len();
        assert_eq!(paths(&file, "p").len(), main);

        let error = file
            .add("/", NewNode::new("header").prop("text", "Twice"))
            .expect_err("the section has a default header");
        assert_eq!(error.code(), Code::DuplicateName);

        let removed = reopen(&file, eager)
            .remove(default_path.as_str())
            .unwrap_or_else(|error| panic!("header remove: {error}"));
        assert_eq!(paths(&removed, "header").len(), 2);
        let added = removed
            .add(
                "/",
                NewNode::new("header")
                    .prop("text", "Fresh")
                    .prop("type", "default")
                    .prop("field", "page"),
            )
            .unwrap_or_else(|error| panic!("header add: {error}"));
        let path = added.path.to_string();
        let file = reopen(&added.into_file(), eager);
        assert_eq!(paths(&file, "header").len(), 3);
        assert_eq!(property(&file, &path, "type").as_deref(), Some("default"));
        assert_eq!(text(&file, &path), "Fresh1");
    }
}

#[test]
fn footnotes_and_endnotes_are_read_added_written_and_removed() {
    for eager in [false, true] {
        let file = open(eager, fixture("footnotes.docx"));
        let notes = paths(&file, "footnote");
        assert!(!notes.is_empty(), "{notes:?}");
        assert!(notes[0].starts_with("/footnotes/footnote["), "{notes:?}");
        let id = property(&file, &notes[0], "id").expect("note id");
        let by_id = file
            .get(format!("/footnote[{id}]").as_str())
            .expect("by note id");
        assert_eq!(by_id.view().path().to_string(), notes[0]);

        let added = file
            .add(
                "/body/p[1]",
                NewNode::new("footnote").prop("text", "Added note"),
            )
            .unwrap_or_else(|error| panic!("footnote add: {error}"));
        let path = added.path.to_string();
        let file = reopen(&added.into_file(), eager);
        assert_eq!(text(&file, &path), "Added note");

        let file = file
            .set(path.as_str(), props! { "text" => "Changed note" })
            .unwrap_or_else(|error| panic!("footnote set: {error}"));
        assert_eq!(text(&file, &path), "Changed note");

        let before = paths(&file, "footnote").len();
        let file = file.remove(path.as_str()).expect("footnote remove");
        assert_eq!(paths(&file, "footnote").len(), before - 1);

        let added = file
            .add(
                "/body/p[1]/r[1]",
                NewNode::new("endnote").prop("text", "End"),
            )
            .unwrap_or_else(|error| panic!("endnote add: {error}"));
        let endnote = added.path.to_string();
        assert!(endnote.starts_with("/endnotes/endnote["), "{endnote}");
        let file = added.into_file();
        assert_eq!(text(&file, &endnote), "End");
        let error = file
            .add("/body/p[1]", NewNode::new("footnote"))
            .expect_err("a note needs its text");
        assert_eq!(error.code(), Code::MissingProperty);
    }
}

#[test]
fn the_root_lists_its_parts_in_officecli_order() {
    let file = open(true, fixture("Headers.docx"));
    let root = file.get("/").expect("root");
    let children: Vec<String> = root
        .view()
        .children()
        .map(|child| child.expect("child").view().path().to_string())
        .collect();
    let position = |name: &str| children.iter().position(|path| path.starts_with(name));
    let body = position("/body").expect("body");
    let styles = position("/styles").expect("styles");
    let header = position("/header[").expect("headers");
    assert!(body < styles && styles < header, "{children:?}");
}

#[test]
fn story_paragraphs_and_runs_are_read_and_written() {
    // Record 0325: `/header[N]/p[M]`, `/footer[N]/p[M]`, `/footnote[N]/p[M]`
    // and `/endnote[N]/p[M]` with the body paragraph machinery.
    let base = litchi_ops::create(litchi_core::FileFormat::Docx)
        .unwrap()
        .add("/body", NewNode::new("paragraph").prop("text", "Body"))
        .unwrap()
        .file
        .add("/", NewNode::new("header").prop("text", "Top"))
        .unwrap()
        .file
        .add("/", NewNode::new("footer").prop("text", "Bottom"))
        .unwrap()
        .file;
    let added = base
        .add(
            "/body/p[1]",
            NewNode::new("footnote").prop("text", "Note text"),
        )
        .unwrap();
    let note = added.path.to_string();
    let base = added.file;
    for eager in [false, true] {
        let file = reopen(&base, eager);
        assert_eq!(text(&file, "/header[1]/p[1]"), "Top");
        assert_eq!(text(&file, "/header[1]/p[1]/r[1]"), "Top");
        assert_eq!(text(&file, "/footer[1]/p[1]"), "Bottom");
        assert!(text(&file, &format!("{note}/p[1]")).contains("Note text"));
        // Unscoped queries stay in the main story.
        assert_eq!(paths(&file, "paragraph").len(), 1);

        // set: text and formatting of a story paragraph and run.
        let set = file
            .set(
                "/header[1]/p[1]",
                props! { "text" => "Header text", "align" => "center" },
            )
            .unwrap_or_else(|error| panic!("header paragraph set: {error}"))
            .set("/footer[1]/p[1]/r[1]", props! { "bold" => true })
            .unwrap_or_else(|error| panic!("footer run set: {error}"));
        let set = reopen(&set, eager);
        assert_eq!(text(&set, "/header[1]/p[1]"), "Header text");
        assert_eq!(text(&set, "/header[1]"), "Header text");
        assert_eq!(
            property(&set, "/header[1]/p[1]", "align").as_deref(),
            Some("center")
        );
        assert_eq!(
            property(&set, "/footer[1]/p[1]/r[1]", "bold").as_deref(),
            Some("true")
        );
        // The main document is untouched.
        assert_eq!(text(&set, "/body/p[1]"), text(&file, "/body/p[1]"));

        // add: a paragraph under a header and a note, a run under a story
        // paragraph.
        let added = set
            .add(
                "/header[1]",
                NewNode::new("paragraph").prop("text", "Second"),
            )
            .unwrap_or_else(|error| panic!("header paragraph add: {error}"));
        assert_eq!(added.path.to_string(), "/header[1]/p[2]");
        let added = added
            .file
            .add("/header[1]/p[2]", NewNode::new("run").prop("text", " line"))
            .unwrap_or_else(|error| panic!("header run add: {error}"));
        assert_eq!(added.path.to_string(), "/header[1]/p[2]/r[2]");
        let added = added
            .file
            .add(
                note.as_str(),
                NewNode::new("paragraph").prop("text", "More"),
            )
            .unwrap_or_else(|error| panic!("note paragraph add: {error}"));
        let file = reopen(&added.file, eager);
        assert_eq!(text(&file, "/header[1]/p[2]"), "Second line");
        assert_eq!(text(&file, &format!("{note}/p[2]")), "More");

        // remove: a run and a paragraph; a story keeps its last paragraph.
        let removed = file
            .remove("/header[1]/p[2]/r[2]")
            .unwrap()
            .remove("/header[1]/p[1]")
            .unwrap();
        assert_eq!(text(&removed, "/header[1]/p[1]"), "Second");
        let error = removed.remove("/header[1]/p[1]").unwrap_err();
        assert_eq!(error.code(), Code::InvalidArgument);
        // A note's first paragraph keeps its reference-mark run.
        let noted = removed
            .set(
                format!("{note}/p[1]").as_str(),
                props! { "text" => "Replaced" },
            )
            .unwrap();
        assert_eq!(text(&noted, &format!("{note}/p[1]")), "Replaced");
        assert_eq!(
            noted
                .get(format!("{note}/p[1]").as_str())
                .unwrap()
                .view()
                .children()
                .count(),
            2
        );
        // An equal set is an exact no-op.
        let bytes = removed.to_bytes().unwrap();
        let same = removed
            .set("/header[1]/p[1]", props! { "text" => "Second" })
            .unwrap();
        assert_eq!(same.to_bytes().unwrap(), bytes);
    }
}
