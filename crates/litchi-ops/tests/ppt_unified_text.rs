//! Unified PPT notes, paragraphs, shape identity and PowerPoint-authored
//! text (record 0260).
//!
//! The fixtures are PowerPoint-authored decks copied into `test-data`, so
//! the tests exercise the outline-backed placeholders, special-information
//! and ruler records that record 0255 found refused.

#![cfg(feature = "ppt")]

use std::io::Cursor;
use std::path::PathBuf;

use litchi_ops::{BatchItem, BatchMode, Code, File, OpenOptions, Props, SourceMode, props};

const SAMPLE: &str = "ole/ppt/SampleShow.ppt";
const NOTES: &str = "poi/test-data/slideshow/basic_test_ppt_file.ppt";

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

fn text(file: &File, path: &str) -> String {
    file.get(path)
        .unwrap_or_else(|error| panic!("resolve {path}: {error}"))
        .view()
        .text()
        .unwrap_or_else(|error| panic!("read {path}: {error}"))
        .unwrap_or_else(|| panic!("{path} has no text"))
        .into_owned()
}

fn property(file: &File, path: &str, key: &str) -> Option<String> {
    file.get(path)
        .unwrap_or_else(|error| panic!("resolve {path}: {error}"))
        .view()
        .format()
        .unwrap_or_else(|error| panic!("read {path} properties: {error}"))
        .find(|property| property.key() == key)
        .map(|property| property.value().to_string())
}

/// A mutation query scoped to the whole presentation.
fn scoped(query: &str) -> litchi_core::selector::Query<'static> {
    use litchi_core::selector::{Path, Query, Scope};
    Query::parse(query)
        .expect("query")
        .into_owned()
        .with_scope(Scope::Path(Box::new(Path::parse("/").expect("root"))))
}

fn signed(bytes: &[u8]) -> Vec<u8> {
    let mut source =
        litchi_cfb::OleFile::open(Cursor::new(bytes), litchi_cfb::OleFileLimits::default())
            .expect("open PPT");
    let mut writer = litchi_cfb::OleWriter::new();
    for path in source.list_streams() {
        let names: Vec<&str> = path.iter().map(String::as_str).collect();
        let data = source.open_stream(&names).expect("read stream");
        writer.create_stream(&names, &data).expect("copy stream");
    }
    writer
        .create_stream(&["_signatures"], b"opaque-signature")
        .expect("add signature stream");
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).expect("write signed PPT");
    output.into_inner()
}

#[test]
fn notes_paragraphs_names_and_alt_text_are_readable() {
    let deck = eager(&fixture(NOTES));
    assert_eq!(
        text(&deck, "/slide[1]/notes"),
        "These are the notes for page 1"
    );
    assert_eq!(
        property(&deck, "/slide[2]/notes", "text").as_deref(),
        Some("These are the notes on page two, again lacking formatting")
    );
    let notes = deck
        .query("notes")
        .expect("notes query")
        .map(|node| node.expect("notes node").path().to_string())
        .collect::<Vec<_>>();
    assert_eq!(notes, ["/slide[1]/notes", "/slide[2]/notes"]);

    // Placeholder text lives in the slide list; paragraphs split it at the
    // binary paragraph marks.
    assert_eq!(
        text(&deck, "/slide[1]/shape[2]"),
        "This is a test subtitle\rThis is on page 1"
    );
    assert_eq!(
        text(&deck, "/slide[1]/shape[2]/paragraph[2]"),
        "This is on page 1"
    );
    assert_eq!(
        deck.get("/slide[1]/shape[2]/paragraph[3]")
            .expect_err("two paragraphs")
            .code(),
        Code::PathNotFound
    );

    let sample = eager(&fixture(SAMPLE));
    assert_eq!(
        property(&sample, "/slide[1]/shape[1]", "name").as_deref(),
        Some("The Title")
    );
    assert_eq!(
        text(
            &sample,
            "/slide[1]/shape[@name=Another Subtitle]/paragraph[3]"
        ),
        "This bit is in italic green"
    );
}

#[test]
fn powerpoint_authored_shape_text_changes_length() {
    each_owner(&fixture(SAMPLE), |source| {
        let before = source.to_bytes().expect("source bytes");
        let changed = source
            .set(
                "/slide[1]/shape[1]",
                props! { "text" => "A much longer replacement title" },
            )
            .expect("outline placeholder text");
        assert_eq!(
            text(&changed, "/slide[1]/shape[1]"),
            "A much longer replacement title"
        );
        assert_eq!(
            text(&changed, "/slide[2]/shape[1]"),
            text(&source, "/slide[2]/shape[1]")
        );
        assert_eq!(source.to_bytes().expect("unchanged source"), before);
        let reopened = eager(&changed.to_bytes().expect("published bytes"));
        assert_eq!(
            text(&reopened, "/slide[1]/shape[1]"),
            "A much longer replacement title"
        );
    });
}

#[test]
fn notes_text_set_round_trips_with_exact_history_and_noops() {
    each_owner(&fixture(NOTES), |source| {
        let before = source.to_bytes().expect("source bytes");
        let commit = source
            .edit(|edit| edit.set("/slide[1]/notes", props! { "text" => "New notes" }))
            .expect("notes write");
        assert!(commit.diagnostics().changed());
        assert_eq!(text(commit.file(), "/slide[1]/notes"), "New notes");
        assert_eq!(
            text(commit.file(), "/slide[2]/notes"),
            text(&source, "/slide[2]/notes")
        );
        assert_eq!(
            text(
                &eager(&commit.file().to_bytes().unwrap()),
                "/slide[1]/notes"
            ),
            "New notes"
        );
        assert_eq!(
            commit
                .patch()
                .inverse()
                .apply(commit.file())
                .expect("inverse")
                .to_bytes()
                .unwrap(),
            before
        );
        assert_eq!(source.to_bytes().unwrap(), before);

        let unchanged = text(&source, "/slide[2]/notes");
        let noop = source
            .edit(|edit| edit.set("/slide[2]/notes", props! { "text" => unchanged.as_str() }))
            .expect("equal notes");
        assert!(noop.patch().is_noop());
        assert!(!noop.diagnostics().changed());
        assert_eq!(noop.file().to_bytes().unwrap(), before);

        // Every notes page of a query in one native edit.
        let all = source
            .set(scoped("notes"), props! { "text" => "Same everywhere" })
            .expect("notes fanout");
        assert_eq!(text(&all, "/slide[1]/notes"), "Same everywhere");
        assert_eq!(text(&all, "/slide[2]/notes"), "Same everywhere");
    });
}

#[test]
fn paragraph_text_set_rewrites_only_the_selected_paragraph() {
    each_owner(&fixture(NOTES), |source| {
        let changed = source
            .set(
                "/slide[1]/shape[2]/paragraph[2]",
                props! { "text" => "This line was rewritten" },
            )
            .expect("paragraph write");
        assert_eq!(
            text(&changed, "/slide[1]/shape[2]"),
            "This is a test subtitle\rThis line was rewritten"
        );
        assert_eq!(
            text(&changed, "/slide[1]/shape[2]/paragraph[1]"),
            "This is a test subtitle"
        );

        // Two paragraphs of one shape publish as one shape text.
        let both = source
            .batch(
                [
                    set_item("/slide[1]/shape[2]/paragraph[1]", "One"),
                    set_item("/slide[1]/shape[2]/paragraph[2]", "Two"),
                ],
                BatchMode::Atomic,
            )
            .expect("paragraph batch");
        assert_eq!(text(both.file(), "/slide[1]/shape[2]"), "One\rTwo");
    });
}

fn set_item(path: &'static str, value: &'static str) -> BatchItem<'static> {
    let mut item = BatchItem::new("set");
    item.path = Some(path.into());
    item.props = Some(Props::new().with("text", value));
    item
}

#[test]
fn protected_and_invalid_text_writes_are_typed_refusals() {
    let bytes = fixture(NOTES);
    let signed = litchi_ops::open_bytes(signed(&bytes)).expect("signed PPT");
    assert!(matches!(signed, File::Signed(_)));
    for path in ["/slide[1]/notes", "/slide[1]/shape[2]/paragraph[1]"] {
        assert_eq!(
            signed
                .set(path, props! { "text" => "refused" })
                .expect_err("signed source")
                .code(),
            Code::DocumentProtected
        );
    }

    let source = eager(&bytes);
    assert_eq!(
        source
            .set("/slide[1]/notes", props! { "bold" => true })
            .expect_err("notes have only text")
            .code(),
        Code::UnsupportedProperty
    );
    assert_eq!(
        source
            .edit(|edit| {
                edit.set("/slide[1]/notes", props! { "text" => "a" })?;
                edit.set("/slide[1]/shape[1]", props! { "text" => "b" })
            })
            .map(|commit| text(commit.file(), "/slide[1]/notes"))
            .expect("notes and shape in one transaction"),
        "a"
    );
    // One call cannot mix element classes.
    assert_eq!(
        source
            .set(scoped("notes, paragraph"), props! { "text" => "mixed" },)
            .map(|_| ())
            .err()
            .map(|error| error.code()),
        Some(Code::UnsupportedCapability)
    );
}

#[cfg(feature = "encryption")]
#[test]
fn unlocked_encrypted_presentations_refuse_text_writes() {
    let mut builder = litchi_ppt::Builder::new();
    let slide = builder.add_slide().unwrap();
    builder.set_slide_notes(slide, "secret notes").unwrap();
    builder
        .set_password(
            "secret",
            litchi_ppt::EncryptionProfile::CryptoApiRc4 { key_bits: 128 },
        )
        .unwrap();
    builder.set_encryption_entropy(litchi_ppt::EncryptionEntropy::from_bytes([7; 16], [9; 16]));
    let bytes = builder.build().unwrap().to_encrypted_bytes().unwrap();
    let File::Locked(locked) = litchi_ops::open_bytes(bytes).unwrap() else {
        panic!("locked presentation")
    };
    let file = locked
        .unlock(litchi_ops::Password::new("secret".to_owned()))
        .unwrap();
    assert_eq!(text(&file, "/slide[1]/notes"), "secret notes");
    assert_eq!(
        file.set("/slide[1]/notes", props! { "text" => "plain" })
            .expect_err("encrypted source")
            .code(),
        Code::DocumentProtected
    );
}
