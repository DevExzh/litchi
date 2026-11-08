//! Unified PPTX picture `set` (record 0333): frame geometry, alternative
//! text and image replacement, alone and with a click action, on both
//! owners and after a byte round trip.
#![cfg(feature = "pptx")]

use std::io::{Cursor, Read as _};

use litchi_ops::{Code, File, NewNode, OpenOptions, PropValue, SourceMode, props};

const PICTURE: &str = "/slide[1]/picture[1]";

fn fixture() -> Vec<u8> {
    let mut builder = litchi_pptx::builder::Builder::new().unwrap();
    builder.presentation().unwrap().add_slide().unwrap();
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

fn reopen(file: &File, eager: bool) -> File {
    open(&file.to_bytes().unwrap(), eager)
}

fn png(tag: &[u8]) -> String {
    let mut bytes = b"\x89PNG\r\n\x1a\n".to_vec();
    bytes.extend_from_slice(tag);
    use base64::Engine as _;
    format!(
        "data:image/png;base64,{}",
        base64::engine::general_purpose::STANDARD.encode(bytes)
    )
}

fn prop(file: &File, path: &str, key: &str) -> Option<String> {
    let node = file.get(path).unwrap();
    node.view()
        .format()
        .unwrap()
        .find(|property| property.key() == key)
        .map(|property| property.value().to_string())
}

fn media(file: &File) -> Vec<Vec<u8>> {
    let mut archive = zip::ZipArchive::new(Cursor::new(file.to_bytes().unwrap())).unwrap();
    let mut out = Vec::new();
    for index in 0..archive.len() {
        let mut member = archive.by_index(index).unwrap();
        if member.name().unwrap().starts_with("ppt/media/") {
            let mut bytes = Vec::new();
            member.read_to_end(&mut bytes).unwrap();
            out.push(bytes);
        }
    }
    out
}

fn with_picture(eager: bool) -> File {
    open(&fixture(), eager)
        .add(
            "/slide[1]",
            NewNode::new("picture")
                .prop("src", png(b"first"))
                .prop("alt", "first image"),
        )
        .unwrap()
        .file
}

#[test]
fn picture_frames_are_set() {
    for eager in [false, true] {
        let file = with_picture(eager);
        assert_eq!(prop(&file, PICTURE, "alt").as_deref(), Some("first image"));
        let moved = file
            .set(PICTURE, props! { "x" => "1in", "width" => "3in" })
            .unwrap();
        let moved = reopen(&moved, eager);
        assert_eq!(prop(&moved, PICTURE, "x").as_deref(), Some("2.54cm"));
        assert_eq!(prop(&moved, PICTURE, "width").as_deref(), Some("7.62cm"));
        assert_eq!(prop(&moved, PICTURE, "y"), prop(&file, PICTURE, "y"));
        assert_eq!(
            file.set(PICTURE, props! { "width" => "0in" })
                .unwrap_err()
                .code(),
            Code::InvalidValue
        );
        // A picture is named at creation only.
        assert_eq!(
            file.set(PICTURE, props! { "name" => "Logo" })
                .unwrap_err()
                .code(),
            Code::UnsupportedProperty
        );
    }
}

#[test]
fn alternative_text_is_set_and_cleared() {
    for eager in [false, true] {
        let file = with_picture(eager);
        let set = file
            .set(PICTURE, props! { "altText" => "a <logo> & mark" })
            .unwrap();
        let set = reopen(&set, eager);
        assert_eq!(
            prop(&set, PICTURE, "alt").as_deref(),
            Some("a <logo> & mark")
        );
        let cleared = set
            .set(PICTURE, props! { "alt" => PropValue::Null })
            .unwrap();
        assert_eq!(prop(&cleared, PICTURE, "alt"), None);
        // An equal value is an exact no-op.
        let same = set
            .set(PICTURE, props! { "alt" => "a <logo> & mark" })
            .unwrap();
        assert_eq!(same.to_bytes().unwrap(), set.to_bytes().unwrap());
    }
}

#[test]
fn the_image_is_replaced_and_the_old_one_dropped() {
    for eager in [false, true] {
        let file = with_picture(eager);
        let replaced = file
            .set(
                PICTURE,
                props! { "src" => png(b"second"), "link" => "https://example.com" },
            )
            .unwrap();
        let replaced = reopen(&replaced, eager);
        let media = media(&replaced);
        assert_eq!(media.len(), 1);
        assert!(media[0].ends_with(b"second"));
        assert_eq!(
            prop(&replaced, PICTURE, "link").as_deref(),
            Some("https://example.com")
        );
        assert_eq!(
            file.set(PICTURE, props! { "src" => "data:text/plain;base64,aGk=" })
                .unwrap_err()
                .code(),
            Code::UnsupportedCapability
        );
        // Crop is a declared gap (the leaf does not model a:srcRect).
        assert_eq!(
            file.set(PICTURE, props! { "crop" => "10" })
                .unwrap_err()
                .code(),
            Code::UnsupportedProperty
        );
    }
}
