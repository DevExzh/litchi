//! Unified blank-slide insertion for legacy PowerPoint (record 0260).

#![cfg(feature = "ppt")]

use std::io::Cursor;
use std::path::PathBuf;

use litchi_ops::{BatchItem, BatchMode, Code, File, NewNode, OpenOptions, Placement, SourceMode};

const SAMPLE: &str = "ole/ppt/SampleShow.ppt";

fn fixture(path: &str) -> Vec<u8> {
    std::fs::read(
        PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data")
            .join(path),
    )
    .expect("PPT fixture")
}

fn eager(bytes: &[u8]) -> File {
    litchi_ops::open_bytes_with(
        bytes.to_vec(),
        &OpenOptions::new().source_mode(SourceMode::Eager),
    )
    .expect("open eager PPT")
}

fn each_owner(bytes: &[u8], mut test: impl FnMut(File)) {
    test(eager(bytes));
    let directory = tempfile::tempdir().expect("PPT source directory");
    let path = directory.path().join("source.ppt");
    std::fs::write(&path, bytes).expect("write PPT source");
    test(
        litchi_ops::open_with(&path, &OpenOptions::new().source_mode(SourceMode::Auto))
            .expect("open positional PPT"),
    );
}

/// The preview of every slide, in order.
fn previews(file: &File) -> Vec<String> {
    file.query("slide")
        .expect("slide query")
        .map(|node| {
            let node = node.expect("slide");
            node.view()
                .preview()
                .expect("preview")
                .map(|preview| preview.to_string())
                .unwrap_or_default()
        })
        .collect()
}

#[test]
fn blank_slides_add_at_every_placement() {
    each_owner(&fixture(SAMPLE), |source| {
        let before = source.to_bytes().expect("source bytes");
        let titles = previews(&source);
        let cases: [(NewNode<'_>, usize); 5] = [
            (NewNode::new("slide"), titles.len()),
            (NewNode::new("slide").at(Placement::Start), 0),
            (
                NewNode::new("slide").at(Placement::Position(litchi_core::Position::new(1))),
                1,
            ),
            (
                NewNode::new("slide").at(Placement::after("/slide[1]").unwrap()),
                1,
            ),
            (
                NewNode::new("slide")
                    .prop("layout", "blank")
                    .at(Placement::before("/slide[2]").unwrap()),
                1,
            ),
        ];
        for (node, position) in cases {
            let added = source.add("/", node).expect("blank slide");
            assert_eq!(added.path.to_string(), format!("/slide[{}]", position + 1));
            let mut expected = titles.clone();
            expected.insert(position, "(untitled)".to_owned());
            assert_eq!(previews(&added.file), expected);
            let node = added.file.get(&added.path).expect("new slide");
            assert_eq!(node.view().child_count().unwrap(), Some(0));
            let reopened = eager(&added.file.to_bytes().expect("published"));
            assert_eq!(previews(&reopened), expected);
        }
        assert_eq!(source.to_bytes().expect("unchanged source"), before);
    });
}

#[test]
fn slide_add_composes_in_edits_and_batches_with_exact_history() {
    let source = eager(&fixture(SAMPLE));
    let before = source.to_bytes().unwrap();
    let commit = source
        .edit(|edit| {
            edit.set(
                "/slide[1]/shape[1]",
                litchi_ops::props! { "text" => "Retitled" },
            )?;
            let path = edit.add("/", NewNode::new("slide").at(Placement::Start))?;
            assert_eq!(path.to_string(), "/slide[1]");
            edit.move_to("/slide[1]", Placement::End)
        })
        .expect("transaction");
    let titles = previews(commit.file());
    assert_eq!(titles.last().map(String::as_str), Some("(untitled)"));
    assert_eq!(titles[0], "Retitled");
    assert_eq!(
        commit
            .patch()
            .inverse()
            .apply(commit.file())
            .unwrap()
            .to_bytes()
            .unwrap(),
        before
    );

    let mut add = BatchItem::new("add");
    add.parent = Some("/".into());
    add.kind = Some("slide".into());
    let batch = source
        .batch([add.clone(), add], BatchMode::Atomic)
        .expect("batch insertion");
    assert_eq!(previews(batch.file()).len(), previews(&source).len() + 2);
}

#[test]
fn invalid_protected_and_released_insertions_are_typed_refusals() {
    let bytes = fixture(SAMPLE);
    let source = eager(&bytes);
    let refusal = |node: NewNode<'_>, parent: &str| {
        source
            .add(parent, node)
            .map(|_| ())
            .expect_err("refused insertion")
            .code()
    };
    assert_eq!(
        refusal(NewNode::new("shape"), "/"),
        Code::UnsupportedElement
    );
    assert_eq!(
        refusal(NewNode::new("slide"), "/slide[1]"),
        Code::UnsupportedType
    );
    assert_eq!(
        refusal(NewNode::new("slide").prop("layout", "title"), "/"),
        Code::UnsupportedCapability
    );
    assert_eq!(
        refusal(NewNode::new("slide").prop("name", "x"), "/"),
        Code::UnsupportedProperty
    );
    assert_eq!(
        refusal(
            NewNode::new("slide").at(Placement::Position(litchi_core::Position::new(9))),
            "/"
        ),
        Code::OutOfRange
    );
    assert_eq!(source.to_bytes().unwrap(), bytes);

    let released =
        litchi_ops::open_bytes_with(bytes.clone(), &OpenOptions::new().retain_source(false))
            .expect("released PPT");
    assert_eq!(
        released
            .add("/", NewNode::new("slide"))
            .map(|_| ())
            .expect_err("released source")
            .code(),
        Code::UnsupportedCapability
    );

    let mut ole = litchi_cfb::OleFile::open(
        Cursor::new(bytes.as_slice()),
        litchi_cfb::OleFileLimits::default(),
    )
    .unwrap();
    let mut writer = litchi_cfb::OleWriter::new();
    for path in ole.list_streams() {
        let names: Vec<&str> = path.iter().map(String::as_str).collect();
        writer
            .create_stream(&names, &ole.open_stream(&names).unwrap())
            .unwrap();
    }
    writer
        .create_stream(&["_signatures"], b"signature")
        .unwrap();
    let mut signed = Cursor::new(Vec::new());
    writer.write_to(&mut signed).unwrap();
    let signed = litchi_ops::open_bytes(signed.into_inner()).unwrap();
    assert_eq!(
        signed
            .add("/", NewNode::new("slide"))
            .map(|_| ())
            .expect_err("signed source")
            .code(),
        Code::DocumentProtected
    );
}

#[test]
fn created_presentations_accept_slides() {
    let created = litchi_ops::create(litchi_core::FileFormat::Ppt).expect("create PPT");
    let count = previews(&created).len();
    let added = created
        .add("/", NewNode::new("slide"))
        .expect("blank slide");
    assert_eq!(previews(&added.file).len(), count + 1);
    let reopened = eager(&added.file.to_bytes().unwrap());
    assert_eq!(previews(&reopened).len(), count + 1);
}
