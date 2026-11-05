//! Exercise every DOCX property advertised for `add` through the public API.
#![cfg(feature = "docx")]

use std::io::{Cursor, Write as _};

use litchi_core::FileFormat;
use litchi_ops::capabilities::Capabilities;
use litchi_ops::value::{Color, Length};
use litchi_ops::{File, NewNode, OpenOptions, Props, SourceMode};

const WORD_MAIN: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";

const INITIAL_RUN_KEYS: &[&str] = &[
    "font.latin",
    "font.ascii",
    "font.hAnsi",
    "font.ea",
    "font.asciiTheme",
    "font.hAnsiTheme",
    "font.eaTheme",
    "font.csTheme",
    "font.hint",
    "lang.latin",
    "lang.ea",
    "lang.cs",
    "size",
    "bold",
    "italic",
    "font.cs",
    "size.cs",
    "bold.cs",
    "italic.cs",
    "color",
    "underline",
    "underline.color",
    "strike",
    "highlight",
    "caps",
    "smallcaps",
    "dstrike",
    "vanish",
    "outline",
    "shadow",
    "emboss",
    "imprint",
    "noproof",
    "superscript",
    "subscript",
    "position",
    "charSpacing",
    "rStyle",
];

fn package(document: &str) -> Vec<u8> {
    let content_types = format!(
        r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="{WORD_MAIN}"/></Types>"#
    );
    let root_relationships = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#;
    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let options =
        zip::write::SimpleFileOptions::default().compression_method(zip::CompressionMethod::Stored);
    for (name, contents) in [
        ("[Content_Types].xml", content_types.as_str()),
        ("_rels/.rels", root_relationships),
        ("word/document.xml", document),
    ] {
        zip.start_file(name, options).expect("start fixture member");
        zip.write_all(contents.as_bytes())
            .expect("write fixture member");
    }
    zip.finish().expect("finish fixture package").into_inner()
}

fn fixture() -> Vec<u8> {
    package(
        r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Existing</w:t></w:r></w:p><w:sectPr/></w:body></w:document>"#,
    )
}

fn open_source() -> File {
    litchi_ops::open_bytes(fixture()).expect("source-backed fixture")
}

fn open_eager() -> File {
    litchi_ops::open_bytes_with(
        fixture(),
        &OpenOptions::new().source_mode(SourceMode::Eager),
    )
    .expect("eager fixture")
}

fn reopen(bytes: Vec<u8>, eager: bool) -> File {
    if eager {
        litchi_ops::open_bytes_with(bytes, &OpenOptions::new().source_mode(SourceMode::Eager))
            .expect("reopen eager result")
    } else {
        litchi_ops::open_bytes(bytes).expect("reopen source result")
    }
}

/// Keep this table keyed by the capability names.  A newly advertised add
/// key without a representative value fails loudly instead of being silently
/// skipped by the test.
fn sample_props(kind: &str, key: &'static str) -> Props<'static> {
    if kind == "paragraph" {
        return match key {
            "text" => Props::new().with(key, "added paragraph"),
            // Record 0319: `styleId` and `styleName` write the style too.
            "style" | "styleId" | "styleName" => Props::new().with(key, "Title"),
            "align" => Props::new().with(key, "right"),
            "spaceBefore" | "spaceAfter" => Props::new().with(key, Length::from_points(6)),
            "lineSpacing" => Props::new().with(key, "1.15x"),
            "lineRule" => Props::new().with(key, "exact"),
            "spaceBeforeLines" | "spaceAfterLines" => Props::new().with(key, 1),
            "spaceBeforeAuto" | "spaceAfterAuto" => Props::new().with(key, true),
            "firstLineIndent" | "hangingIndent" | "indent" | "rightIndent" => {
                Props::new().with(key, Length::from_points(1))
            },
            "firstLineChars" | "hangingChars" | "leftChars" | "rightChars" => {
                Props::new().with(key, 1)
            },
            "keepNext" | "keepLines" | "pageBreakBefore" | "widowControl" => {
                Props::new().with(key, true)
            },
            "direction" => Props::new().with(key, "rtl"),
            "contextualSpacing" => Props::new().with(key, true),
            "numId" => Props::new().with(key, "9"),
            "numLevel" => Props::new().with(key, "1"),
            "outlineLvl" => Props::new().with(key, 1),
            // Record 0342: a new paragraph joins a new list.
            "listStyle" => Props::new().with(key, "bullet"),
            "start" => Props::new().with(key, 3),
            _ if INITIAL_RUN_KEYS.contains(&key) => sample_run_props(key),
            _ => panic!("unhandled advertised DOCX paragraph add property {key}"),
        };
    }
    assert_eq!(
        kind, "run",
        "the add liveness table only covers paragraph/run"
    );
    sample_run_props(key)
}

fn sample_run_props(key: &'static str) -> Props<'static> {
    match key {
        "text" => Props::new().with(key, "added run"),
        "font.latin" | "font.ascii" | "font.hAnsi" | "font.ea" | "font.cs" => {
            Props::new().with(key, "Arial")
        },
        "font.asciiTheme" | "font.hAnsiTheme" | "font.eaTheme" | "font.csTheme" => {
            Props::new().with(key, "minorAscii")
        },
        "font.hint" => Props::new().with(key, "default"),
        "lang.latin" | "lang.ea" | "lang.cs" => Props::new().with(key, "en-US"),
        "size" | "size.cs" => Props::new().with(key, Length::from_points(14)),
        "bold" | "italic" | "bold.cs" | "italic.cs" | "strike" | "caps" | "smallcaps"
        | "dstrike" | "vanish" | "outline" | "shadow" | "emboss" | "imprint" | "noproof"
        | "superscript" | "subscript" => Props::new().with(key, true),
        "color" | "underline.color" => Props::new().with(key, Color::rgb(0x1F, 0x4E, 0x79)),
        "underline" => Props::new().with(key, "single"),
        "highlight" => Props::new().with(key, "yellow"),
        "direction" => Props::new().with(key, "rtl"),
        "position" | "charSpacing" => Props::new().with(key, Length::from_points(1)),
        "rStyle" => Props::new().with(key, "Emphasis"),
        _ => panic!("unhandled advertised DOCX run add property {key}"),
    }
}

fn format_contains(file: &File, path: &str, key: &str) -> bool {
    file.get(path)
        .unwrap_or_else(|error| panic!("{path}: {error}"))
        .view()
        .format()
        .unwrap_or_else(|error| panic!("{path} format: {error}"))
        .any(|property| property.key() == key)
}

fn assert_added(file: &File, path: &str, kind: &str, key: &str) {
    let node = file
        .get(path)
        .unwrap_or_else(|error| panic!("{path} after adding {key}: {error}"));
    if key == "text" {
        let actual = node
            .view()
            .text()
            .unwrap_or_else(|error| panic!("{path} text after adding {key}: {error}"))
            .map(|value| value.into_owned())
            .unwrap_or_default();
        let expected = if kind == "paragraph" {
            "added paragraph"
        } else {
            "added run"
        };
        assert_eq!(actual, expected, "{path} text after adding {key}");
    } else if kind == "paragraph" && INITIAL_RUN_KEYS.contains(&key) {
        let run_path = format!("{path}/r[1]");
        assert!(
            format_contains(file, &run_path, key),
            "initial-run property {key} was not emitted at {run_path}"
        );
    } else if matches!(key, "style" | "styleId" | "styleName") {
        // `styleName` reads a display name only through a styles part,
        // which the fixture lacks; the style id is written either way.
        assert_eq!(
            node.view()
                .style()
                .unwrap_or_else(|error| panic!("{path} style: {error}")),
            Some("Title".into()),
            "paragraph style after adding {key}"
        );
    } else {
        assert!(
            format_contains(file, path, key),
            "property {key} was not emitted at {path}"
        );
    }
}

#[test]
fn every_advertised_docx_add_property_is_live_for_source_and_eager_owners() {
    let capabilities = Capabilities::of(FileFormat::Docx);
    let paragraph = capabilities
        .element("paragraph")
        .expect("DOCX paragraph capability");
    for key in INITIAL_RUN_KEYS {
        let property = paragraph
            .property(key)
            .unwrap_or_else(|| panic!("missing paragraph add-only property {key}"));
        assert!(property.add(), "paragraph {key} is not add-capable");
        assert!(!property.get(), "paragraph {key} leaked into get");
        assert!(!property.set(), "paragraph {key} leaked into set");
    }

    let openers: [(&str, fn() -> File, bool); 2] =
        [("source", open_source, false), ("eager", open_eager, true)];
    for (mode, open, eager) in openers {
        for (kind, parent) in [("paragraph", "/body"), ("run", "/body/p[1]")] {
            let element = capabilities
                .element(kind)
                .unwrap_or_else(|| panic!("missing DOCX {kind} capability"));
            let addable: Vec<_> = element
                .properties()
                .filter(|property| property.add())
                .collect();
            assert!(!addable.is_empty(), "DOCX {kind} has no add properties");
            for property in addable {
                let key = property.name();
                let original = open();
                let added = original
                    .add(
                        parent,
                        NewNode::new(kind)
                            .with_props(sample_props(kind, key))
                            .unwrap_or_else(|error| panic!("{mode} {kind} {key}: {error}")),
                    )
                    .unwrap_or_else(|error| panic!("{mode} {kind} {key}: add failed: {error}"));
                let path = added.path.to_string();
                assert_eq!(
                    path,
                    if kind == "paragraph" {
                        "/body/p[2]"
                    } else {
                        "/body/p[1]/r[2]"
                    },
                    "{mode} {kind} {key}: canonical path"
                );
                let updated = added.into_file();
                assert_added(&updated, &path, kind, key);
                let reopened = reopen(
                    updated
                        .to_bytes()
                        .unwrap_or_else(|error| panic!("{mode} {kind} {key}: publish: {error}")),
                    eager,
                );
                assert_added(&reopened, &path, kind, key);
            }
        }
    }
}
