//! Unified ODM master documents (record 0282): reads, writes, save and
//! create through the shared ODT adapter; the package keeps its master MIME
//! type and its linked sections stay links.

#![cfg(feature = "odm")]

use litchi_core::FileFormat;
use litchi_ops::capabilities::{Capabilities, Verb};
use litchi_ops::{Code, File, NewNode, create, open, open_bytes, props};

fn fixture(path: &str) -> std::path::PathBuf {
    std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/libreoffice-core")
        .join(path)
}

const MASTER: &str = "vcl/qa/cppunit/pdfexport/data/master.odm";

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
    // The stored `mimetype` member is the first entry, uncompressed.
    let start = 30 + "mimetype".len();
    let end = bytes[start..]
        .windows(2)
        .position(|pair| pair == b"PK")
        .map_or(bytes.len(), |offset| start + offset);
    String::from_utf8_lossy(&bytes[start..end]).into_owned()
}

#[test]
fn master_documents_open_as_odm_and_read_their_body() {
    let file = open(fixture(MASTER)).expect("open ODM");
    assert_eq!(file.format(), Some(FileFormat::Odm));
    assert!(matches!(file, File::Odm(_)));
    let first = file.get("/body/p[1]").expect("first paragraph");
    assert_eq!(first.format(), FileFormat::Odm);
    assert!(text(&file, "/body/p[1]").starts_with("Writer Guide"));
    assert!(file.query("section").expect("query").count() > 0);
    assert!(file.text().expect("text").contains("Writer Guide"));
    assert!(file.metadata().expect("metadata").is_some());
}

#[test]
fn edits_keep_the_master_package_and_round_trip() {
    let file = open_bytes(std::fs::read(fixture(MASTER)).unwrap()).expect("open ODM");
    let before = file.to_bytes().unwrap();
    let file = file
        .set("/body/p[1]", props! { "text" => "Edited master" })
        .expect("set");
    let file = file
        .add("/body", NewNode::new("paragraph").prop("text", "Appended"))
        .expect("add")
        .file;
    let bytes = file.to_bytes().unwrap();
    assert_ne!(bytes, before);
    assert_eq!(
        mimetype(&bytes),
        "application/vnd.oasis.opendocument.text-master"
    );
    let reopened = open_bytes(bytes).expect("reopen");
    assert_eq!(reopened.format(), Some(FileFormat::Odm));
    assert_eq!(text(&reopened, "/body/p[1]"), "Edited master");
    assert_eq!(text(&reopened, "/body/p[last()]"), "Appended");
}

#[test]
fn an_unchanged_value_is_an_exact_no_op() {
    let file = open_bytes(std::fs::read(fixture(MASTER)).unwrap()).expect("open ODM");
    let before = file.to_bytes().unwrap();
    // A section's current name is a value the ODT adapter compares exactly.
    let name = file
        .get("/body/section[1]")
        .expect("section")
        .view()
        .format()
        .expect("format")
        .into_iter()
        .find(|prop| prop.key() == "name")
        .map(|prop| prop.to_string())
        .expect("section name");
    let file = file
        .set("/body/section[1]", props! { "name" => name.as_str() })
        .expect("set");
    assert_eq!(file.to_bytes().unwrap(), before);
    let file = file.edit(|_| Ok(())).expect("empty edit").into_file();
    assert_eq!(file.to_bytes().unwrap(), before);
}

#[test]
fn errors_name_the_master_format() {
    let file = open(fixture(MASTER)).expect("open ODM");
    let error = file.get("/body/p[9999]").expect_err("missing");
    assert_eq!(error.format(), Some(FileFormat::Odm));
    let error = file
        .set("/body/p[1]", props! { "noSuchKey" => true })
        .expect_err("unknown key");
    assert_eq!(error.format(), Some(FileFormat::Odm));
    assert_ne!(error.code(), Code::InternalError);
}

#[test]
fn save_writes_back_to_the_origin() {
    let directory = tempfile::tempdir().unwrap();
    let path = directory.path().join("master.odm");
    std::fs::copy(fixture(MASTER), &path).unwrap();
    let file = open(&path).expect("open");
    let file = file
        .set("/body/p[1]", props! { "text" => "Saved" })
        .expect("set");
    file.save().expect("save");
    let reopened = open(&path).expect("reopen");
    assert_eq!(reopened.format(), Some(FileFormat::Odm));
    assert_eq!(text(&reopened, "/body/p[1]"), "Saved");
}

#[test]
fn create_makes_an_empty_master_document() {
    let file = create(FileFormat::Odm).expect("create");
    assert_eq!(file.format(), Some(FileFormat::Odm));
    let bytes = file.to_bytes().unwrap();
    assert_eq!(
        mimetype(&bytes),
        "application/vnd.oasis.opendocument.text-master"
    );
    let file = open_bytes(bytes)
        .expect("reopen")
        .add("/body", NewNode::new("paragraph").prop("text", "Hello"))
        .expect("add")
        .file;
    let reopened = open_bytes(file.to_bytes().unwrap()).expect("reopen");
    assert_eq!(text(&reopened, "/body/p[last()]"), "Hello");
    assert_eq!(
        create(FileFormat::Odm).unwrap().to_bytes().unwrap(),
        create(FileFormat::Odm).unwrap().to_bytes().unwrap(),
        "fresh creation is deterministic"
    );
}

#[test]
fn capabilities_report_the_shared_adapter() {
    let report = Capabilities::of(FileFormat::Odm);
    for verb in [
        Verb::Get,
        Verb::Query,
        Verb::Set,
        Verb::Add,
        Verb::Remove,
        Verb::Save,
        Verb::Create,
    ] {
        assert!(report.verb(verb).is_supported(), "{verb:?}");
    }
    assert!(report.element("section").is_some());
}
