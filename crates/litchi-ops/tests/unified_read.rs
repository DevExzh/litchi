//! Public File dispatch: lazy queries, raw admission, source revisions and views.
#![cfg(feature = "docx")]

use litchi_ops::query::{Plan, Query};
use litchi_ops::view::Mode;
use litchi_ops::{Code, File, OpenOptions, SourceMode};
use std::io::{Cursor, Write as _};

fn fixture() -> Vec<u8> {
    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let options = zip::write::SimpleFileOptions::default();
    for (name, content) in [
        (
            "[Content_Types].xml",
            r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#,
        ),
        (
            "_rels/.rels",
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#,
        ),
        (
            "word/document.xml",
            r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body><w:p w14:paraId="12345678"><w:r><w:t>first</w:t></w:r></w:p><w:p><w:r><w:t>second</w:t></w:r></w:p></w:body></w:document>"#,
        ),
    ] {
        zip.start_file(name, options).unwrap();
        zip.write_all(content.as_bytes()).unwrap();
    }
    zip.finish().unwrap().into_inner()
}

fn file() -> File {
    litchi_ops::open_bytes(fixture()).unwrap()
}

#[test]
fn file_reads_typed_and_text_paths_and_reuses_plans() {
    let file = file();
    assert_eq!(file.root().unwrap().path().to_string(), "/");
    let parsed = litchi_ops::path::Path::parse("/body/p[2]").unwrap();
    let node = file.get(&parsed).unwrap();
    assert_eq!(node.view().text().unwrap().as_deref(), Some("second"));
    assert!(file.find("/body/p[3]").unwrap().is_none());
    assert_eq!(
        file.get("/body/p[3]").unwrap_err().code(),
        Code::PathNotFound
    );
    assert_eq!(
        file.find("/body/p[0]").unwrap_err().code(),
        Code::InvalidPath
    );
    let query = Query::parse("p").unwrap();
    let plan = Plan::compile(&query, litchi_core::FileFormat::Docx).unwrap();
    for _ in 0..2 {
        let matches = file.query(&plan).unwrap();
        assert!(std::ptr::eq(matches.plan(), &plan));
        let nodes = matches.collect::<litchi_ops::Result<Vec<_>>>().unwrap();
        assert_eq!(nodes.len(), 2);
        for node in nodes {
            assert_eq!(
                file.get(node.positional_path().to_path()).unwrap().path(),
                node.path()
            );
        }
    }
    assert_eq!(
        file.query("p")
            .unwrap()
            .first()
            .unwrap()
            .unwrap()
            .path()
            .to_string(),
        "/body/p[1]"
    );
    assert_eq!(file.query("p").unwrap().limit(1).count(), 1);
    assert_eq!(
        file.query("p[unknown=secret]").unwrap().diagnostics().len(),
        1
    );
}

#[test]
fn raw_paths_and_queries_preserve_positional_output() {
    let file = file();
    let path = "/body/p[@paraId=12345678]";
    assert_eq!(file.get(path).unwrap_err().code(), Code::InvalidPath);
    assert_eq!(
        file.get(litchi_ops::path::raw::path(path).unwrap())
            .unwrap()
            .path()
            .to_string(),
        "/body/p[1]"
    );
    let selector = "p[paraId=12345678]";
    assert_eq!(
        file.query(selector).unwrap_err().code(),
        Code::InvalidSelector
    );
    let node = file
        .query(litchi_ops::query::raw::query(selector).unwrap())
        .unwrap()
        .first()
        .unwrap()
        .unwrap();
    assert_eq!(node.path().to_string(), "/body/p[1]");
    let serialized = serde_json::to_string(&node.view()).unwrap();
    assert!(!serialized.contains("paraId"));
    assert!(!serialized.contains("12345678"));
}

#[test]
fn public_views_agree_across_source_modes_and_propagate_sink_errors() {
    let source = file();
    let eager = litchi_ops::open_bytes_with(
        fixture(),
        &OpenOptions::default().source_mode(SourceMode::Eager),
    )
    .unwrap();
    for mode in [Mode::Text, Mode::Annotated, Mode::Outline, Mode::Stats] {
        assert_eq!(source.view(mode).unwrap(), eager.view(mode).unwrap());
        assert_eq!(
            serde_json::to_value(source.view_json(mode).unwrap()).unwrap(),
            serde_json::to_value(eager.view_json(mode).unwrap()).unwrap()
        );
        assert!(!format!("{:?}", source.view_json(mode).unwrap()).contains("first"));
    }
    assert!(source.view(Mode::Text).unwrap().contains("first"));
    assert_eq!(
        source.view(Mode::Html).unwrap_err().code(),
        Code::UnsupportedCapability
    );
    struct Refuse;
    impl std::fmt::Write for Refuse {
        fn write_str(&mut self, _: &str) -> std::fmt::Result {
            Err(std::fmt::Error)
        }
    }
    assert_eq!(
        source.view_to(Mode::Text, &mut Refuse).unwrap_err().code(),
        Code::IoError
    );
}

#[test]
fn lazy_file_query_and_cached_nodes_refuse_changed_source() {
    let directory = tempfile::tempdir().unwrap();
    let path = directory.path().join("revision.docx");
    std::fs::write(&path, fixture()).unwrap();
    let file = litchi_ops::open(&path).unwrap();
    let node = file.get("/body/p[1]").unwrap();
    let mut matches = file.query("p").unwrap();
    let mut zero = file.query("p").unwrap().limit(0);
    std::fs::write(&path, b"changed source").unwrap();
    assert!(zero.next().is_none());
    assert_eq!(matches.next().unwrap().unwrap_err().code(), Code::IoError);
    assert!(matches.next().is_none());
    assert_eq!(node.view().text().unwrap_err().code(), Code::IoError);
    assert_eq!(file.root().unwrap_err().code(), Code::IoError);
    assert_eq!(file.view(Mode::Text).unwrap_err().code(), Code::IoError);
}

#[test]
fn signed_files_delegate_read_access() {
    let file = litchi_ops::open(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../test-data/poi/test-data/xmldsign/hello-world-signed.docx"
    ))
    .unwrap();
    assert!(file.is_signed());
    assert_eq!(file.root().unwrap().path().to_string(), "/");
    assert!(file.query("p").unwrap().first().unwrap().is_some());
    assert!(!file.view(Mode::Text).unwrap().is_empty());
}
