//! The `/metadata` set path of signed legacy documents, and of legacy
//! documents whose property sets the editor cannot rewrite: every request
//! follows the signature policy or is refused with a user-facing code.
#![cfg(all(feature = "doc", feature = "xls", feature = "ppt"))]

use std::io::Cursor;

use litchi_core::FileFormat;
use litchi_ops::props::PropValue;
use litchi_ops::{
    BatchItem, BatchMode, Code, Error, File, ItemResult, Props, create, open_bytes, props,
};

const DOCUMENT_SUMMARY: &str = "\u{5}DocumentSummaryInformation";

/// `bytes` rewritten stream by stream, with the `replaced` stream set to the
/// given data and, when `sign` is set, a legacy signature stream.
fn rewrite(bytes: Vec<u8>, replaced: Option<(&str, &[u8])>, sign: bool) -> Vec<u8> {
    use litchi_cfb::{OleFile, OleFileLimits, OleWriter};

    let mut source = OleFile::open(Cursor::new(bytes), OleFileLimits::default()).unwrap();
    let mut writer = OleWriter::new();
    for path in source.list_streams() {
        let path = path.iter().map(String::as_str).collect::<Vec<_>>();
        for depth in 1..path.len() {
            writer.create_storage(&path[..depth]).unwrap();
        }
        if replaced.is_some_and(|(name, _)| path == [name]) {
            continue;
        }
        let data = source.open_stream(&path).unwrap();
        writer.create_stream(&path, &data).unwrap();
    }
    if let Some((name, data)) = replaced {
        writer.create_stream(&[name], data).unwrap();
    }
    if sign {
        writer
            .create_stream(&["_signatures"], b"test signature")
            .unwrap();
    }
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}

fn stream(bytes: &[u8], name: &str) -> Option<Vec<u8>> {
    let mut file = litchi_cfb::OleFile::open(
        Cursor::new(bytes.to_vec()),
        litchi_cfb::OleFileLimits::default(),
    )
    .unwrap();
    file.exists(&[name])
        .then(|| file.open_stream(&[name]).unwrap())
}

fn has_signature(bytes: &[u8]) -> bool {
    stream(bytes, "_signatures").is_some()
}

fn title(file: &File) -> Option<String> {
    file.metadata().unwrap().unwrap_or_default().title
}

/// A created `format` document titled "Original", as its published bytes.
fn titled(format: FileFormat) -> Vec<u8> {
    create(format)
        .unwrap()
        .set("/metadata", props! { "title" => "Original" })
        .unwrap()
        .to_bytes()
        .unwrap()
}

fn signed(format: FileFormat) -> (File, Vec<u8>) {
    let bytes = rewrite(titled(format), None, true);
    let file = open_bytes(bytes.clone()).unwrap();
    assert!(matches!(file, File::Signed(_)), "{format:?}");
    (file, bytes)
}

fn metadata_item(props: Props<'static>) -> BatchItem<'static> {
    let mut item = BatchItem::new("set");
    item.path = Some("/metadata".into());
    item.props = Some(props);
    item
}

/// Hand `check` the refusal of one metadata request through each write
/// route, or `None` where the route applied the request.
fn each_route(
    file: &File,
    key: &'static str,
    value: &'static str,
    mut check: impl FnMut(&str, Option<&Error>),
) {
    let set = file.set("/metadata", Props::new().with(key, value));
    check("set", set.as_ref().err());
    let edit = file.edit(|edit| edit.set("/metadata", Props::new().with(key, value)));
    check("edit", edit.as_ref().err());
    for (route, mode) in [
        ("atomic batch", BatchMode::Atomic),
        ("best-effort batch", BatchMode::BestEffort),
    ] {
        // A best-effort batch reports an item failure in its results.
        match file.batch([metadata_item(Props::new().with(key, value))], mode) {
            Ok(commit) => match &commit.value()[0] {
                ItemResult::Failed(error) => check(route, Some(error)),
                _ => check(route, None),
            },
            Err(error) => check(route, Some(&error)),
        }
    }
}

/// The Document Summary Information stream of this fixture carries a
/// Dictionary property, which the property-set editor refuses to rewrite.
fn unrewritable_summary() -> Vec<u8> {
    let fixture = std::fs::read(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/ole/doc/documentProperties.doc"),
    )
    .unwrap();
    stream(&fixture, DOCUMENT_SUMMARY).expect("the fixture's document summary stream")
}

#[test]
fn unrewritable_property_sets_refuse_with_a_user_facing_code() {
    let summary = unrewritable_summary();
    for format in [FileFormat::Doc, FileFormat::Xls, FileFormat::Ppt] {
        let plain = rewrite(titled(format), Some((DOCUMENT_SUMMARY, &summary)), false);
        let signed = open_bytes(rewrite(plain.clone(), None, true)).unwrap();
        let File::Signed(wrapper) = &signed else {
            panic!("{format:?}: expected a signed file");
        };
        let files = [
            ("plain", open_bytes(plain).unwrap()),
            ("unsigned", wrapper.clone().unsign()),
            ("signed", signed),
        ];
        for (state, file) in files {
            let original = file.to_bytes().unwrap();
            for (key, value) in [("title", "Changed"), ("custom.Department", "Research")] {
                each_route(&file, key, value, |route, error| {
                    let what = format!("{format:?}, {state}, {route}, {key}");
                    let error = error.unwrap_or_else(|| panic!("{what}: expected a refusal"));
                    // A batch refuses a signed file before it reads an item.
                    if state == "signed" && route.ends_with("batch") {
                        assert_signed_refusal(error, &what);
                    } else {
                        assert_eq!(error.code(), Code::CorruptFile, "{what}: {error}");
                        assert_eq!(error.format(), Some(format), "{what}");
                    }
                });
            }
            let error = file
                .set("/metadata", props! { "title" => PropValue::Null })
                .unwrap_err();
            assert_eq!(
                error.code(),
                Code::CorruptFile,
                "{format:?}, {state}: {error}"
            );
            assert_eq!(file.to_bytes().unwrap(), original, "{format:?}, {state}");
        }
    }
}

#[test]
fn signed_metadata_noop_keeps_the_signed_publication() {
    for format in [FileFormat::Doc, FileFormat::Xls, FileFormat::Ppt] {
        let (file, bytes) = signed(format);

        let direct = file
            .set("/metadata", props! { "title" => "Original" })
            .unwrap_or_else(|error| panic!("{format:?} direct no-op: {error}"));
        assert_eq!(direct.to_bytes().unwrap(), bytes, "{format:?}");

        let commit = file
            .edit(|edit| {
                edit.set("/metadata", props! { "title" => "Original" })?;
                edit.set("/metadata", props! { "title" => "Original" })
            })
            .unwrap_or_else(|error| panic!("{format:?} no-op transaction: {error}"));
        assert!(!commit.diagnostics().changed(), "{format:?}");
        assert!(commit.file().is_signed(), "{format:?}");
        assert_eq!(commit.file().to_bytes().unwrap(), bytes, "{format:?}");
        let replayed = commit
            .patch()
            .apply(&open_bytes(bytes.clone()).unwrap())
            .unwrap_or_else(|error| panic!("{format:?} no-op replay: {error}"));
        assert_eq!(replayed.to_bytes().unwrap(), bytes, "{format:?}");

        // A batch refuses a signed file before it reads an item.
        for mode in [BatchMode::Atomic, BatchMode::BestEffort] {
            let error = file
                .batch(
                    [metadata_item(Props::new().with("title", "Original"))],
                    mode,
                )
                .unwrap_err();
            assert_signed_refusal(&error, "no-op batch");
        }
        assert_eq!(file.to_bytes().unwrap(), bytes, "{format:?}");
    }
}

#[track_caller]
fn assert_signed_refusal(error: &Error, what: &str) {
    assert_eq!(error.code(), Code::DocumentProtected, "{what}: {error}");
    assert!(
        error
            .suggestion()
            .is_some_and(|suggestion| suggestion.contains("Signed::unsign")),
        "{what}: {error}"
    );
}

/// The changed half of the policy.
#[test]
fn signed_metadata_changes_refuse_until_the_signature_is_consumed() {
    for format in [FileFormat::Doc, FileFormat::Xls, FileFormat::Ppt] {
        let (file, bytes) = signed(format);
        for (key, value) in [("title", "Changed"), ("custom.Department", "Research")] {
            each_route(&file, key, value, |route, error| {
                let what = format!("{format:?}, {route}, {key}");
                let error = error.unwrap_or_else(|| panic!("{what}: expected a refusal"));
                assert_signed_refusal(error, &what);
            });
        }
        assert_signed_refusal(
            &file
                .set("/metadata", props! { "title" => PropValue::Null })
                .unwrap_err(),
            "clear",
        );
        // A no-op step does not launder a changed one.
        assert_signed_refusal(
            &file
                .edit(|edit| {
                    edit.set("/metadata", props! { "title" => "Original" })?;
                    edit.set("/metadata", props! { "subject" => "Changed" })
                })
                .unwrap_err(),
            "no-op then change",
        );
        assert_eq!(file.to_bytes().unwrap(), bytes, "{format:?}");

        let File::Signed(wrapper) = file else {
            panic!("{format:?}: expected a signed file");
        };
        let unsigned = wrapper.unsign();
        let commit = unsigned
            .edit(|edit| {
                edit.set("/metadata", props! { "title" => "Authorized" })?;
                edit.set("/metadata", props! { "custom.Department" => "Research" })
            })
            .unwrap_or_else(|error| panic!("{format:?} after unsign: {error}"));
        assert!(commit.diagnostics().changed(), "{format:?}");
        let output = commit.file().to_bytes().unwrap();
        assert!(!has_signature(&output), "{format:?}");
        let reopened = open_bytes(output).unwrap();
        assert!(!reopened.is_signed(), "{format:?}");
        assert_eq!(title(&reopened).as_deref(), Some("Authorized"));

        let cleared = unsigned
            .set("/metadata", props! { "title" => PropValue::Null })
            .unwrap_or_else(|error| panic!("{format:?} clear after unsign: {error}"));
        assert!(title(&cleared).is_none(), "{format:?}");
        assert!(!has_signature(&cleared.to_bytes().unwrap()), "{format:?}");

        let batch = unsigned
            .batch(
                [metadata_item(Props::new().with("title", "Batch"))],
                BatchMode::Atomic,
            )
            .unwrap_or_else(|error| panic!("{format:?} batch after unsign: {error}"));
        assert_eq!(title(batch.file()).as_deref(), Some("Batch"), "{format:?}");
        assert!(!has_signature(&batch.file().to_bytes().unwrap()));
    }
}

/// An unlocked encrypted workbook is not a signed one: its metadata stays
/// refused, and the refusal does not name the signature.
#[cfg(feature = "encryption")]
#[test]
fn encrypted_xls_metadata_stays_refused() {
    const SUGGESTION: &str = "export the original encrypted source explicitly";

    let mut builder = litchi_xls::Builder::new();
    builder.add_worksheet("Data").unwrap();
    builder
        .set_password("secret", litchi_xls::EncryptionProfile::OfficeBinaryRc4)
        .unwrap();
    builder.set_encryption_entropy(litchi_xls::EncryptionEntropy::generate().unwrap());
    let bytes = builder.build().unwrap().to_encrypted_bytes().unwrap();
    let File::Locked(locked) = open_bytes(bytes).unwrap() else {
        panic!("an encrypted XLS must open locked");
    };
    let file = locked
        .unlock(litchi_ops::Password::new("secret".to_owned()))
        .unwrap();
    for (key, value) in [("title", "Changed"), ("custom.Department", "Research")] {
        each_route(&file, key, value, |route, error| {
            let what = format!("{route}, {key}");
            let error = error.unwrap_or_else(|| panic!("{what}: expected a refusal"));
            assert_eq!(error.code(), Code::DocumentProtected, "{what}: {error}");
            assert_eq!(error.suggestion(), Some(SUGGESTION), "{what}: {error}");
        });
    }
    let error = file
        .set("/metadata", props! { "title" => PropValue::Null })
        .unwrap_err();
    assert_eq!(error.code(), Code::DocumentProtected, "clear: {error}");
    assert_eq!(error.suggestion(), Some(SUGGESTION), "clear: {error}");
}
