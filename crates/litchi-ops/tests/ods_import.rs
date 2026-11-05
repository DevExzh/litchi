//! CSV/TSV import into ODS (record 0295, record 0266 follow-up F8): one
//! leaf commit, native value types, an exact no-op for empty input, typed
//! refusals and an inverse patch.
#![cfg(feature = "ods")]

use litchi_core::FileFormat;
use litchi_ops::capabilities::Verb;
use litchi_ops::import::Options;
use litchi_ops::{Code, File, props};

fn text(file: &File, path: &str) -> Option<String> {
    file.get(path)
        .unwrap()
        .view()
        .text()
        .unwrap()
        .map(std::borrow::Cow::into_owned)
}

fn property(file: &File, path: &str, key: &str) -> Option<String> {
    file.get(path)
        .unwrap()
        .view()
        .format()
        .unwrap()
        .find(|property| property.key() == key)
        .map(|property| property.value().to_string())
}

fn reopen(file: &File) -> File {
    litchi_ops::open_bytes(file.to_bytes().unwrap()).unwrap()
}

const SOURCE: &[u8] =
    b"name,score,active,joined\nAda,36,TRUE,2026-10-13\n\"Lovelace, A.\",-2.5,false,1815-12-10T08:30:00\n";

#[test]
fn imports_native_values_in_one_commit() {
    let original = litchi_ops::create(FileFormat::Ods).unwrap();
    assert!(original.capabilities().verb(Verb::Import).is_supported());
    assert!(!original.capabilities().verb(Verb::Evaluate).is_supported());
    let before = original.to_bytes().unwrap();
    let commit = original
        .import("/Sheet1", SOURCE, &Options::csv().header(true))
        .unwrap();
    let summary = commit.value();
    assert_eq!((summary.rows(), summary.columns()), (3, 4));
    assert_eq!(summary.cells(), 12);
    assert_eq!(summary.numbers(), 2);
    assert_eq!(summary.booleans(), 2);
    assert_eq!(summary.dates(), 2);
    assert_eq!(summary.dates_as_text(), 0);
    assert_eq!(commit.diagnostics().commits(), 1);
    assert_eq!(original.to_bytes().unwrap(), before);
    for file in [commit.file().clone(), reopen(commit.file())] {
        assert_eq!(text(&file, "/Sheet1/A1").as_deref(), Some("name"));
        assert_eq!(text(&file, "/Sheet1/A3").as_deref(), Some("Lovelace, A."));
        assert_eq!(text(&file, "/Sheet1/B3").as_deref(), Some("-2.5"));
        assert_eq!(
            property(&file, "/Sheet1/B2", "type").as_deref(),
            Some("Number")
        );
        assert_eq!(
            property(&file, "/Sheet1/C2", "type").as_deref(),
            Some("Boolean")
        );
        assert_eq!(
            property(&file, "/Sheet1/D2", "type").as_deref(),
            Some("Date")
        );
        assert_eq!(
            text(&file, "/Sheet1/D3").as_deref(),
            Some("1815-12-10T08:30:00")
        );
    }
    let restored = commit.patch().inverse().apply(commit.file()).unwrap();
    assert_eq!(restored.to_bytes().unwrap(), before);
}

#[test]
fn origins_empty_fields_noops_and_refusals() {
    let file = litchi_ops::create(FileFormat::Ods)
        .unwrap()
        .set("/Sheet1/A1", props! { "value" => "old" })
        .unwrap();
    // An origin cell offsets the rectangle; an empty field clears a cell.
    let commit = file
        .import("/Sheet1/B2", b"x,y\n", &Options::csv())
        .unwrap();
    assert_eq!(text(commit.file(), "/Sheet1/C2").as_deref(), Some("y"));
    let cleared = file.import("/Sheet1", b",new\n", &Options::csv()).unwrap();
    assert_eq!(cleared.value().empty(), 1);
    let reopened = reopen(cleared.file());
    assert_eq!(text(&reopened, "/Sheet1/B1").as_deref(), Some("new"));
    assert!(
        text(&reopened, "/Sheet1/A1").is_none_or(|value| value.is_empty() || value == "(empty)")
    );
    // Empty input is an exact no-op that still resolves the target.
    assert!(
        file.import("/Sheet1", b"", &Options::csv())
            .unwrap()
            .patch()
            .is_noop()
    );
    assert_eq!(
        file.import("/Missing", b"", &Options::csv())
            .unwrap_err()
            .code(),
        Code::PathNotFound
    );
    // Formulas need OpenFormula translation.
    let error = file
        .import("/Sheet1", b"=1+1\n", &Options::csv().formulas(true))
        .unwrap_err();
    assert_eq!(error.code(), Code::UnsupportedCapability);
    assert_eq!(error.format(), Some(FileFormat::Ods));
    // A formula-looking field is text by default.
    let plain = file.import("/Sheet1", b"=1+1\n", &Options::csv()).unwrap();
    assert_eq!(text(plain.file(), "/Sheet1/A1").as_deref(), Some("=1+1"));
}
