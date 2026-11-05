//! Unified ODT writes (record 0276): set, add, remove, move, swap,
//! metadata, create, save and coalesced batches over the leaf body edits.

#![cfg(feature = "odt")]

use litchi_core::FileFormat;
use litchi_ops::{Code, File, NewNode, Placement, open_bytes, props};

fn fixture(name: &str) -> Vec<u8> {
    std::fs::read(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/odf")
            .join(name),
    )
    .expect("read fixture")
}

fn open(name: &str) -> File {
    open_bytes(fixture(name)).expect("open ODT")
}

fn reopen(file: &File) -> File {
    open_bytes(file.to_bytes().expect("publish")).expect("reopen")
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

fn prop(file: &File, path: &str, key: &str) -> Option<String> {
    let node = file.get(path).expect("path");
    let format = node.view().format().expect("format");
    format
        .into_iter()
        .find(|prop| prop.key() == key)
        .map(|prop| prop.to_string())
}

#[test]
fn paragraph_text_and_formatting_round_trip() {
    let file = open("corpus/writer-paragraph-styles.odt");
    let before = file.to_bytes().unwrap();
    let file = file
        .set(
            "/body/p[2]",
            props! { "text" => "Replaced  text", "align" => "center", "bold" => true },
        )
        .unwrap();
    let file = reopen(&file);
    assert_eq!(text(&file, "/body/p[2]"), "Replaced  text");
    assert_eq!(
        prop(&file, "/body/p[2]", "align").as_deref(),
        Some("center")
    );
    assert_eq!(prop(&file, "/body/p[2]", "bold").as_deref(), Some("true"));
    // The first paragraph kept its text.
    let original = open_bytes(before).unwrap();
    assert_eq!(text(&file, "/body/p[1]"), text(&original, "/body/p[1]"));
}

#[test]
fn run_formatting_wraps_character_data() {
    let file = open("corpus/writer-paragraph-styles.odt");
    let file = file
        .set(
            "/body/p[1]/r[1]",
            props! { "italic" => true, "size" => "14pt", "color" => "#FF0000" },
        )
        .unwrap();
    let file = reopen(&file);
    let run = "/body/p[1]/r[1]";
    assert_eq!(prop(&file, run, "italic").as_deref(), Some("true"));
    assert_eq!(prop(&file, run, "size").as_deref(), Some("14pt"));
    assert_eq!(prop(&file, run, "color").as_deref(), Some("#FF0000"));
    assert!(file.query("r[italic=true]").unwrap().count() >= 1);
}

#[test]
fn exact_noop_returns_the_source_bytes() {
    let source = fixture("corpus/writer-table.odt");
    let file = open_bytes(source.clone()).unwrap();
    let name = prop(&file, "/body/tbl[1]", "name").expect("table name");
    let unchanged = file
        .set("/body/tbl[1]", props! { "name" => name.as_str() })
        .unwrap();
    assert_eq!(unchanged.to_bytes().unwrap(), source);
    let commit = file.edit(|_| Ok(())).unwrap();
    assert!(!commit.diagnostics().changed());
    assert_eq!(commit.file().to_bytes().unwrap(), source);
}

#[test]
fn add_every_insertable_element() {
    let file = litchi_ops::create(FileFormat::Odt).expect("create ODT");
    let added = file
        .add(
            "/body",
            NewNode::new("heading")
                .prop("text", "Title")
                .prop("level", 2),
        )
        .unwrap();
    assert_eq!(added.path.to_string(), "/body/h[1]");
    let file = added.file;
    let added = file
        .add("/body", NewNode::new("paragraph").prop("text", "First"))
        .unwrap();
    assert_eq!(added.path.to_string(), "/body/p[1]");
    let file = added.file;
    let file = file
        .add(
            "/body",
            NewNode::new("p")
                .prop("text", "Zeroth")
                .at(Placement::before("/body/p[1]").unwrap()),
        )
        .unwrap()
        .file;
    assert_eq!(text(&file, "/body/p[1]"), "Zeroth");
    assert_eq!(text(&file, "/body/p[2]"), "First");
    let file = file
        .add(
            "/body/p[2]",
            NewNode::new("r")
                .prop("text", " bold tail")
                .prop("bold", true),
        )
        .unwrap()
        .file;
    let file = file
        .add(
            "/body/p[2]",
            NewNode::new("hyperlink")
                .prop("url", "https://example.org/")
                .prop("text", "link"),
        )
        .unwrap()
        .file;
    let file = file
        .add("/body/p[2]", NewNode::new("bookmark").prop("name", "mark"))
        .unwrap()
        .file;
    let file = file
        .add(
            "/body/p[2]",
            NewNode::new("footnote").prop("text", "a note"),
        )
        .unwrap()
        .file;
    let file = file
        .add(
            "/body/p[2]",
            NewNode::new("comment")
                .prop("text", "a comment")
                .prop("author", "Reviewer"),
        )
        .unwrap()
        .file;
    let file = file
        .add("/body", NewNode::new("tbl").prop("rows", 2).prop("cols", 3))
        .unwrap()
        .file;
    let file = file.add("/body/tbl[1]", NewNode::new("tr")).unwrap().file;
    let file = file
        .add("/body", NewNode::new("list").prop("text", "one\ntwo"))
        .unwrap()
        .file;
    let file = file
        .add(
            "/body/list[1]",
            NewNode::new("listItem").prop("text", "three"),
        )
        .unwrap()
        .file;
    let file = file
        .add("/body", NewNode::new("section").prop("name", "Appendix"))
        .unwrap()
        .file;
    let file = reopen(&file);
    assert_eq!(text(&file, "/body/h[1]"), "Title");
    assert_eq!(prop(&file, "/body/h[1]", "level").as_deref(), Some("2"));
    assert_eq!(text(&file, "/body/p[2]/r[2]"), " bold tail");
    assert_eq!(
        prop(&file, "/body/p[2]/r[2]", "bold").as_deref(),
        Some("true")
    );
    assert_eq!(
        prop(&file, "/body/p[2]/hyperlink[1]", "url").as_deref(),
        Some("https://example.org/")
    );
    assert_eq!(
        prop(&file, "/body/p[2]/bookmark[@name=mark]", "name").as_deref(),
        Some("mark")
    );
    assert_eq!(text(&file, "/body/p[2]/footnote[1]"), "a note");
    assert_eq!(
        prop(&file, "/body/p[2]/comment[1]", "author").as_deref(),
        Some("Reviewer")
    );
    assert_eq!(file.query("/body/tbl[1]/tr").unwrap().count(), 3);
    assert_eq!(file.query("/body/tbl[1]/tr[3]/tc").unwrap().count(), 3);
    assert_eq!(file.query("/body/list[1]/listItem").unwrap().count(), 3);
    assert_eq!(
        prop(&file, "/body/section[1]", "name").as_deref(),
        Some("Appendix")
    );
}

#[test]
fn remove_move_and_swap() {
    let file = litchi_ops::create(FileFormat::Odt).unwrap();
    let mut file = file;
    for text in ["a", "b", "c", "d"] {
        file = file
            .add("/body", NewNode::new("p").prop("text", text))
            .unwrap()
            .file;
    }
    let moved = file
        .move_to("/body/p[1]", Placement::after("/body/p[3]").unwrap())
        .unwrap();
    let order = |file: &File| -> Vec<String> {
        file.query("/body/p")
            .unwrap()
            .map(|node| node.unwrap().view().text().unwrap().unwrap().into_owned())
            .collect()
    };
    assert_eq!(order(&moved), ["b", "c", "a", "d"]);
    let swapped = moved.swap("/body/p[1]", "/body/p[4]").unwrap();
    assert_eq!(order(&swapped), ["d", "c", "a", "b"]);
    let removed = swapped.remove("/body/p[2]").unwrap();
    assert_eq!(order(&removed), ["d", "a", "b"]);
    let start = removed.move_to("/body/p[3]", Placement::Start).unwrap();
    assert_eq!(order(&start), ["b", "d", "a"]);
    // Cells keep the table grid.
    let table = start
        .add("/body", NewNode::new("tbl").prop("rows", 1).prop("cols", 2))
        .unwrap()
        .file;
    let error = table.remove("/body/tbl[1]/tr[1]/tc[1]").unwrap_err();
    assert_eq!(error.code(), Code::InvalidValue);
}

#[test]
fn metadata_reads_and_writes() {
    let file = open("corpus/writer-paragraph-styles.odt");
    let file = file
        .set(
            "/metadata",
            props! { "title" => "Unified", "author" => "Ada", "keywords" => "a, b" },
        )
        .unwrap();
    let file = reopen(&file);
    assert_eq!(
        prop(&file, "/metadata", "title").as_deref(),
        Some("Unified")
    );
    assert_eq!(prop(&file, "/metadata", "author").as_deref(), Some("Ada"));
    let metadata = file.metadata().unwrap().unwrap();
    assert_eq!(metadata.title.as_deref(), Some("Unified"));
    let error = file
        .set("/metadata", props! { "created" => "2020-01-01" })
        .unwrap_err();
    assert_eq!(error.code(), Code::UnsupportedCapability);
    let error = file
        .set("/metadata", props! { "nonsense" => "x" })
        .unwrap_err();
    assert_eq!(error.code(), Code::UnsupportedProperty);
}

#[test]
fn property_errors_are_typed() {
    let file = open("corpus/writer-paragraph-styles.odt");
    let error = file
        .set("/body/p[1]", props! { "nonsense" => true })
        .unwrap_err();
    assert_eq!(error.code(), Code::UnsupportedProperty);
    let error = file
        .set("/body/p[1]", props! { "style" => "No Such Style" })
        .unwrap_err();
    assert_eq!(error.code(), Code::InvalidValue);
    let error = file
        .set("/body/p[1]/r[1]", props! { "style" => "Standard" })
        .unwrap_err();
    assert_eq!(error.code(), Code::UnsupportedCapability);
    let error = file
        .set("/body/p[1]", props! { "align" => "sideways" })
        .unwrap_err();
    assert_eq!(error.code(), Code::InvalidValue);
}

#[test]
fn coalesced_batch_commits_once_and_inverts() {
    let file = litchi_ops::create(FileFormat::Odt).unwrap();
    let mut file = file;
    for index in 0..20 {
        file = file
            .add("/body", NewNode::new("p").prop("text", format!("p{index}")))
            .unwrap()
            .file;
    }
    let source = file.to_bytes().unwrap();
    let commit = file
        .edit(|tx| {
            for index in 1..=20 {
                let path = format!("/body/p[{index}]");
                tx.set(path.as_str(), props! { "bold" => true })?;
                tx.set(path.as_str(), props! { "text" => format!("q{index}") })?;
            }
            tx.set("/metadata", props! { "title" => "batched" })?;
            Ok(())
        })
        .unwrap();
    assert!(commit.diagnostics().changed());
    let edited = reopen(commit.file());
    assert_eq!(text(&edited, "/body/p[7]"), "q7");
    assert_eq!(prop(&edited, "/body/p[7]", "bold").as_deref(), Some("true"));
    let restored = commit.patch().inverse().apply(commit.file()).unwrap();
    assert_eq!(restored.to_bytes().unwrap(), source);
}

#[test]
fn batch_runs_split_at_replaced_content() {
    let file = open("corpus/writer-paragraph-styles.odt");
    let commit = file
        .edit(|tx| {
            tx.set("/body/p[1]", props! { "text" => "one run" })?;
            // The run is inside the paragraph replaced above: the batch
            // resolves it against the result.
            tx.set("/body/p[1]/r[1]", props! { "italic" => true })?;
            Ok(())
        })
        .unwrap();
    let file = reopen(commit.file());
    assert_eq!(text(&file, "/body/p[1]/r[1]"), "one run");
    assert_eq!(
        prop(&file, "/body/p[1]/r[1]", "italic").as_deref(),
        Some("true")
    );
}

#[test]
fn create_with_metadata_and_save_atomically() {
    let metadata = litchi_core::Metadata {
        title: Some("Fresh".to_owned()),
        ..litchi_core::Metadata::default()
    };
    let file = litchi_ops::create_with(
        FileFormat::Odt,
        &litchi_ops::CreateOptions::new().metadata(metadata),
    )
    .unwrap();
    let file = file
        .add("/body", NewNode::new("p").prop("text", "saved"))
        .unwrap()
        .file;
    let directory = tempfile::tempdir().unwrap();
    let path = directory.path().join("out.odt");
    file.save_as(&path).unwrap();
    let reopened = litchi_ops::open(&path).unwrap();
    assert_eq!(text(&reopened, "/body/p[1]"), "saved");
    assert_eq!(
        reopened.metadata().unwrap().unwrap().title.as_deref(),
        Some("Fresh")
    );
    // The path-opened handle edits and saves back to its origin.
    let edited = reopened
        .set("/body/p[1]", props! { "text" => "again" })
        .unwrap();
    edited.save().unwrap();
    let again = litchi_ops::open(&path).unwrap();
    assert_eq!(text(&again, "/body/p[1]"), "again");
    let refused = again.save_with(
        litchi_ops::SaveTo::Path(&path),
        &litchi_ops::SaveOptions::new().overwrite(litchi_ops::Overwrite::Refuse),
    );
    assert!(refused.is_err());
}

#[test]
fn capabilities_report_the_write_surface() {
    use litchi_ops::capabilities::{Capabilities, Verb};
    let capabilities = Capabilities::of(FileFormat::Odt);
    for verb in [
        Verb::Get,
        Verb::Query,
        Verb::Set,
        Verb::Add,
        Verb::Remove,
        Verb::Move,
        Verb::Swap,
        Verb::Edit,
        Verb::Batch,
        Verb::Save,
        Verb::Create,
    ] {
        assert!(capabilities.verb(verb).is_supported(), "{verb:?}");
    }
    let paragraph = capabilities.element("paragraph").unwrap();
    assert!(paragraph.property("text").unwrap().set());
}
