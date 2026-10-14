//! Unified OTH (OpenDocument HTML template, `text-web`) documents (record
//! 0282): reads, writes, save and create through the shared ODT adapter.
//! The package keeps its `text-web` MIME type.

#![cfg(feature = "oth")]

use litchi_core::FileFormat;
use litchi_ops::capabilities::{Capabilities, Verb};
use litchi_ops::{File, NewNode, create, open, open_bytes, props};

fn template() -> Vec<u8> {
    let package = litchi_oth::Builder::new()
        .paragraph(litchi_oth::paragraph::Paragraph::new("First"))
        .paragraph(litchi_oth::paragraph::Paragraph::new("Second"))
        .build()
        .expect("build OTH");
    package.to_bytes().expect("bytes")
}

fn text(file: &File, path: &str) -> String {
    file.get(path)
        .expect("path")
        .view()
        .text()
        .expect("text")
        .unwrap_or_default()
        .into_owned()
}

fn mimetype(bytes: &[u8]) -> String {
    let start = 30 + "mimetype".len();
    let end = bytes[start..]
        .windows(2)
        .position(|pair| pair == b"PK")
        .map_or(bytes.len(), |offset| start + offset);
    String::from_utf8_lossy(&bytes[start..end]).into_owned()
}

#[test]
fn templates_open_as_oth_and_read_their_body() {
    let file = open_bytes(template()).expect("open OTH");
    assert_eq!(file.format(), Some(FileFormat::Oth));
    assert!(matches!(file, File::Oth(_)));
    assert_eq!(text(&file, "/body/p[1]"), "First");
    assert_eq!(file.get("/body/p[2]").unwrap().format(), FileFormat::Oth);
    assert_eq!(file.query("p").expect("query").count(), 2);
}

#[test]
fn edits_keep_the_template_package() {
    let file = open_bytes(template()).expect("open OTH");
    let file = file
        .set("/body/p[1]", props! { "text" => "Changed", "bold" => true })
        .expect("set")
        .remove("/body/p[2]")
        .expect("remove")
        .add("/body", NewNode::new("heading").prop("text", "Title"))
        .expect("add")
        .file;
    let bytes = file.to_bytes().unwrap();
    assert_eq!(
        mimetype(&bytes),
        "application/vnd.oasis.opendocument.text-web"
    );
    let reopened = open_bytes(bytes.clone()).expect("reopen");
    assert_eq!(reopened.format(), Some(FileFormat::Oth));
    assert_eq!(text(&reopened, "/body/p[1]"), "Changed");
    // The leaf owner reads what the unified adapter wrote.
    let leaf = litchi_oth::Package::from_bytes(&bytes).expect("leaf reopen");
    assert!(leaf.document().content_xml().contains("Changed"));
}

#[test]
fn edit_batches_are_atomic() {
    let file = open_bytes(template()).expect("open OTH");
    let before = file.to_bytes().unwrap();
    let result = file.edit(|tx| {
        tx.set("/body/p[1]", props! { "text" => "One" })?;
        tx.set("/body/p[99]", props! { "text" => "Missing" })?;
        Ok(())
    });
    assert!(result.is_err());
    assert_eq!(file.to_bytes().unwrap(), before);
}

#[test]
fn save_and_create_round_trip() {
    let directory = tempfile::tempdir().unwrap();
    let path = directory.path().join("page.oth");
    let file = create(FileFormat::Oth).expect("create");
    assert_eq!(file.format(), Some(FileFormat::Oth));
    let file = file
        .add("/body", NewNode::new("paragraph").prop("text", "Hello"))
        .expect("add")
        .file;
    file.save_as(&path).expect("save");
    let reopened = open(&path).expect("reopen");
    assert_eq!(reopened.format(), Some(FileFormat::Oth));
    assert_eq!(text(&reopened, "/body/p[last()]"), "Hello");
    let reopened = reopened
        .set("/body/p[last()]", props! { "text" => "Again" })
        .expect("set");
    reopened.save().expect("save in place");
    assert_eq!(text(&open(&path).unwrap(), "/body/p[last()]"), "Again");
}

#[test]
fn capabilities_report_the_shared_adapter() {
    let report = Capabilities::of(FileFormat::Oth);
    for verb in [
        Verb::Get,
        Verb::Query,
        Verb::Set,
        Verb::Add,
        Verb::Save,
        Verb::Create,
    ] {
        assert!(report.verb(verb).is_supported(), "{verb:?}");
    }
}
