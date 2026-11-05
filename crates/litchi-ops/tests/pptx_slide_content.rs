//! Unified PPTX notes, pictures and layout placeholders (record 0252).
#![cfg(feature = "pptx")]

use litchi_ops::{Code, File, NewNode, OpenOptions, PropValue, SourceMode, props};

/// A 1×1 transparent PNG.
const PNG: &[u8] = &[
    0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52,
    0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1F, 0x15, 0xC4,
    0x89, 0x00, 0x00, 0x00, 0x0A, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9C, 0x63, 0x00, 0x01, 0x00, 0x00,
    0x05, 0x00, 0x01, 0x0D, 0x0A, 0x2D, 0xB4, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, 0x44, 0xAE,
    0x42, 0x60, 0x82,
];

fn fixture() -> Vec<u8> {
    let mut builder = litchi_pptx::builder::Builder::new().unwrap();
    builder
        .presentation()
        .unwrap()
        .add_slide()
        .unwrap()
        .add_text_box("alpha", 0, 0, 914_400, 914_400);
    builder.build().unwrap().to_bytes().unwrap()
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

fn reopen(file: &File) -> File {
    open(&file.to_bytes().unwrap(), false)
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

fn data_uri() -> String {
    use std::fmt::Write as _;
    // Minimal standard base64 for the fixture bytes.
    const TABLE: &[u8; 64] = b"ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
    let mut out = String::from("data:image/png;base64,");
    for chunk in PNG.chunks(3) {
        let b = [
            chunk[0],
            *chunk.get(1).unwrap_or(&0),
            *chunk.get(2).unwrap_or(&0),
        ];
        let n = (u32::from(b[0]) << 16) | (u32::from(b[1]) << 8) | u32::from(b[2]);
        for (index, shift) in [18, 12, 6, 0].into_iter().enumerate() {
            if index <= chunk.len() {
                out.push(TABLE[((n >> shift) & 63) as usize] as char);
            } else {
                out.push('=');
            }
        }
    }
    let _ = write!(out, "");
    out
}

#[test]
fn notes_are_created_replaced_and_read_back() {
    let bytes = fixture();
    for eager in [false, true] {
        let file = open(&bytes, eager);
        assert!(file.get("/slide[1]/notes").is_err());
        let created = file
            .set("/slide[1]", props! { "notes" => "Remember the demo" })
            .unwrap();
        let reread = reopen(&created);
        assert_eq!(text(&reread, "/slide[1]/notes"), "Remember the demo");
        let replaced = reread
            .set("/slide[1]/notes", props! { "text" => "Second" })
            .unwrap();
        assert_eq!(text(&reopen(&replaced), "/slide[1]/notes"), "Second");
        // The slide's shapes are untouched.
        assert_eq!(text(&replaced, "/slide[1]/shape[1]"), "alpha");
        let cleared = replaced
            .set("/slide[1]/notes", props! { "text" => PropValue::Null })
            .unwrap();
        // A notes page without text reads as no notes.
        assert!(reopen(&cleared).get("/slide[1]/notes").is_err());
    }
}

#[test]
fn notes_refusals_are_typed() {
    let file = open(&fixture(), false);
    let code = |result: litchi_ops::Result<File>| result.err().map(|error| error.code());
    assert_eq!(
        code(file.set("/slide[1]", props! { "notes" => "x", "hidden" => true })),
        Some(Code::UnsupportedProperty)
    );
    assert_eq!(
        code(file.set("/slide[9]", props! { "notes" => "x" })),
        Some(Code::PathNotFound)
    );
}

#[test]
fn pictures_are_added_from_data_uris_and_files() {
    let bytes = fixture();
    let directory = tempfile::tempdir().unwrap();
    let path = directory.path().join("dot.png");
    std::fs::write(&path, PNG).unwrap();
    for eager in [false, true] {
        let file = open(&bytes, eager);
        let added = file
            .add(
                "/slide[1]",
                NewNode::new("picture")
                    .prop("src", data_uri())
                    .prop("x", "1in")
                    .prop("width", "2in")
                    .prop("height", "1in"),
            )
            .unwrap();
        assert_eq!(added.path.to_string(), "/slide[1]/picture[1]");
        let second = added
            .file
            .add(
                "/slide[1]",
                NewNode::new("picture").prop("path", path.to_str().unwrap()),
            )
            .unwrap();
        assert_eq!(second.path.to_string(), "/slide[1]/picture[2]");
        let reread = reopen(&second.file);
        let picture = reread.get("/slide[1]/picture[1]").unwrap();
        let x = picture
            .view()
            .format()
            .unwrap()
            .find(|property| property.key() == "x")
            .map(|property| property.value().to_string());
        assert_eq!(x.as_deref(), Some("2.54cm"));
        assert!(reread.get("/slide[1]/picture[2]").is_ok());
        assert_eq!(text(&reread, "/slide[1]/shape[1]"), "alpha");
    }
}

#[test]
fn picture_refusals_are_typed() {
    let file = open(&fixture(), false);
    let code = |node: NewNode<'_>| file.add("/slide[1]", node).err().map(|error| error.code());
    assert_eq!(code(NewNode::new("picture")), Some(Code::MissingProperty));
    assert_eq!(
        code(NewNode::new("picture").prop("src", "data:text/plain;base64,aGk=")),
        Some(Code::UnsupportedCapability)
    );
    assert_eq!(
        code(NewNode::new("picture").prop("src", "data:image/png,raw")),
        Some(Code::UnsupportedCapability)
    );
    assert_eq!(
        code(NewNode::new("picture").prop("src", "/nonexistent/litchi.png")),
        Some(Code::FileNotFound)
    );
}

#[test]
fn layout_slides_get_fillable_placeholders() {
    let mut builder = litchi_pptx::builder::Builder::new().unwrap();
    builder.presentation().unwrap().add_slide().unwrap();
    let package = builder.build().unwrap();
    let master = package.presentation().unwrap().slide_masters().unwrap()[0]
        .part()
        .part()
        .partname()
        .clone();
    let mut edit = package.edit().unwrap();
    edit.add_slide_layout(
        &master,
        litchi_pptx::master_layout::SlideLayoutKind::Object,
        "Agenda Layout",
        &[
            litchi_pptx::master_layout::PlaceholderSpec::new(
                litchi_pptx::master_layout::PlaceholderKind::Title,
            ),
            litchi_pptx::master_layout::PlaceholderSpec::new(
                litchi_pptx::master_layout::PlaceholderKind::Body,
            )
            .with_index(1),
        ],
    )
    .unwrap();
    let bytes = edit.commit().unwrap().into_package().to_bytes().unwrap();
    for eager in [false, true] {
        let file = open(&bytes, eager);
        let added = file
            .add("/", NewNode::new("slide").prop("layout", "Agenda Layout"))
            .unwrap()
            .file;
        let title = "/slide[2]/placeholder[1]";
        assert!(added.get(title).is_ok(), "{eager}");
        assert_eq!(text(&added, title), "");
        let filled = added
            .set(
                "/slide[2]/shape[1]",
                props! { "text" => "Agenda", "bold" => true },
            )
            .unwrap()
            .set("/slide[2]/shape[2]", props! { "text" => "One\nTwo" })
            .unwrap();
        let reread = reopen(&filled);
        assert_eq!(text(&reread, "/slide[2]/shape[1]"), "Agenda");
        assert_eq!(text(&reread, "/slide[2]/shape[2]"), "One\nTwo");
    }
}

fn prop(file: &File, path: &str, key: &str) -> Option<String> {
    file.get(path)
        .unwrap()
        .view()
        .format()
        .unwrap()
        .find(|property| property.key() == key)
        .map(|property| property.value().to_string())
}

#[test]
fn comments_are_added_and_set() {
    let bytes = fixture();
    for eager in [false, true] {
        let file = open(&bytes, eager);
        let added = file
            .add(
                "/slide[1]",
                NewNode::new("comment")
                    .prop("text", "Check this")
                    .prop("author", "Ada Lovelace")
                    .prop("date", "2026-10-13T00:00:00.000")
                    .prop("x", "1in"),
            )
            .unwrap();
        assert_eq!(added.path.to_string(), "/slide[1]/comment[1]");
        let second = added
            .file
            .add(
                "/slide[1]",
                NewNode::new("comment")
                    .prop("text", "Second")
                    .prop("author", "Ada Lovelace"),
            )
            .unwrap();
        assert_eq!(second.path.to_string(), "/slide[1]/comment[2]");
        let reread = reopen(&second.file);
        assert_eq!(text(&reread, "/slide[1]/comment[1]"), "Check this");
        assert_eq!(
            prop(&reread, "/slide[1]/comment[1]", "author").as_deref(),
            Some("Ada Lovelace")
        );
        assert_eq!(
            prop(&reread, "/slide[1]/comment[1]", "initials").as_deref(),
            Some("AL")
        );
        assert_eq!(text(&reread, "/slide[1]/comment[2]"), "Second");
        let edited = reread
            .set(
                "/slide[1]/comment[2]",
                props! { "text" => "Edited", "y" => "2in" },
            )
            .unwrap();
        let reread = reopen(&edited);
        assert_eq!(text(&reread, "/slide[1]/comment[2]"), "Edited");
        assert_eq!(
            prop(&reread, "/slide[1]/comment[2]", "y").as_deref(),
            Some("5.08cm")
        );
        assert_eq!(text(&reread, "/slide[1]/comment[1]"), "Check this");
    }
    let file = open(&bytes, false);
    let code = file
        .add(
            "/slide[1]",
            NewNode::new("comment").prop("text", "anonymous"),
        )
        .err()
        .map(|error| error.code());
    assert_eq!(code, Some(Code::MissingProperty));
    // A date that is not an XML dateTime is a typed value refusal, not a
    // corrupt comments part (record 0330).
    let Err(error) = file.add(
        "/slide[1]",
        NewNode::new("comment")
            .prop("author", "Ada")
            .prop("date", "yesterday"),
    ) else {
        panic!("an invalid comment date is refused");
    };
    assert_eq!(error.code(), Code::InvalidValue, "{error:?}");
    assert_eq!(error.property(), Some("date"));
}
