//! SmartArt in Word documents (record 0295, record 0284 follow-up 0284a):
//! inline insertion into an existing document through the authored spec,
//! and the document-wide `smartart` element read from the leaf inventory.
#![cfg(feature = "docx")]

use litchi_core::FileFormat;
use litchi_ops::capabilities::{Capabilities, Verb};
use litchi_ops::{Code, File, NewNode, OpenOptions, SourceMode, create, open_bytes, props};

fn prop(file: &File, path: &str, key: &str) -> Option<String> {
    let node = file.get(path).unwrap();
    let view = node.view();
    view.format()
        .unwrap()
        .find(|prop| prop.key() == key)
        .map(|prop| prop.value().to_string())
}

fn document() -> File {
    create(FileFormat::Docx)
        .unwrap()
        .add("/body", NewNode::new("paragraph").prop("text", "Steps:"))
        .unwrap()
        .file
}

#[test]
fn smartart_is_inserted_read_and_queried() {
    let file = document();
    let added = file
        .add(
            "/body/p[1]",
            NewNode::new("smartart")
                .prop("layout", "process")
                .prop("text", "Plan\nBuild\n\tTest\nShip")
                .prop("name", "Flow")
                .prop("alt", "The release flow")
                .prop("width", "12cm"),
        )
        .unwrap();
    assert_eq!(added.path.to_string(), "/smartart[1]");
    for file in [
        added.file.clone(),
        open_bytes(added.file.to_bytes().unwrap()).unwrap(),
    ] {
        assert_eq!(
            prop(&file, "/smartart[1]", "layout").as_deref(),
            Some("process")
        );
        assert_eq!(
            prop(&file, "/smartart[1]", "nodeCount").as_deref(),
            Some("4")
        );
        assert_eq!(
            prop(&file, "/smartart[1]", "text").as_deref(),
            Some("Plan\nBuild\n\tTest\nShip")
        );
        assert!(file.get("/body/smartart[1]").is_ok());
        assert_eq!(file.query("smartart").unwrap().count(), 1);
        assert_eq!(
            file.get("/body/p[1]")
                .unwrap()
                .view()
                .text()
                .unwrap()
                .as_deref(),
            Some("Steps:")
        );
    }
    // A second graphic follows in document order.
    let second = added
        .file
        .add(
            "/body/p[1]",
            NewNode::new("smartart").prop("text", "One\nTwo"),
        )
        .unwrap();
    assert_eq!(second.path.to_string(), "/smartart[2]");
    assert_eq!(
        prop(&second.file, "/smartart[2]", "layout").as_deref(),
        Some("list")
    );
    // The template is unchanged.
    assert!(file.get("/smartart[1]").is_err());
}

#[test]
fn source_backed_documents_take_and_read_smartart() {
    let bytes = document().to_bytes().unwrap();
    let directory = tempfile::tempdir().unwrap();
    let path = directory.path().join("source.docx");
    std::fs::write(&path, &bytes).unwrap();
    let file =
        litchi_ops::open_with(&path, &OpenOptions::new().source_mode(SourceMode::Auto)).unwrap();
    let added = file
        .add(
            "/body/p[1]",
            NewNode::new("smartart")
                .prop("layout", "cycle")
                .prop("text", "A\nB\nC"),
        )
        .unwrap();
    let saved = directory.path().join("saved.docx");
    added.file.save_as(&saved).unwrap();
    let reopened =
        litchi_ops::open_with(&saved, &OpenOptions::new().source_mode(SourceMode::Auto)).unwrap();
    assert_eq!(
        prop(&reopened, "/smartart[1]", "layout").as_deref(),
        Some("cycle")
    );
}

#[test]
fn refusals_leave_the_document_unchanged() {
    let file = document();
    let before = file.to_bytes().unwrap();
    for (node, code) in [
        (NewNode::new("smartart"), Code::MissingProperty),
        (
            NewNode::new("smartart")
                .prop("layout", "spiral")
                .prop("text", "A"),
            Code::InvalidValue,
        ),
        (
            NewNode::new("smartart")
                .prop("text", "A")
                .prop("color", "red"),
            Code::UnsupportedProperty,
        ),
    ] {
        let error = file.add("/body/p[1]", node).unwrap_err();
        assert_eq!(error.code(), code, "{error}");
    }
    assert_eq!(file.to_bytes().unwrap(), before);
    let added = file
        .add("/body/p[1]", NewNode::new("smartart").prop("text", "A"))
        .unwrap()
        .file;
    // `set` writes the outline (record 0335); an unknown key is refused.
    let error = added
        .set("/smartart[1]", props! { "color" => "red" })
        .unwrap_err();
    assert_eq!(error.code(), Code::UnsupportedProperty, "{error}");
    let set = added.set("/smartart[1]", props! { "text" => "B" }).unwrap();
    assert_eq!(prop(&set, "/smartart[1]", "text").as_deref(), Some("B"));
    let capabilities = Capabilities::of(FileFormat::Docx);
    let element = capabilities
        .elements()
        .find(|element| element.name() == "smartart")
        .unwrap();
    assert!(element.operations().add);
    assert!(element.operations().set && element.operations().remove);
    assert!(capabilities.verb(Verb::Add).is_supported());
}
