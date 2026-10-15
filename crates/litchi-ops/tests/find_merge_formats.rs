//! Find/replace and template merge on the legacy, RTF and ODF formats
//! (record 0295): the shared text-surface engine over each format's native
//! run, paragraph or cell edits; one atomic commit, exact no-ops, round
//! trips and typed refusals.

#![allow(clippy::unwrap_used, reason = "test assertions")]

use litchi_core::FileFormat;
#[allow(unused_imports, reason = "used by a subset of the format features")]
use litchi_ops::merge::{Missing, Options, Record};
#[allow(unused_imports, reason = "used by a subset of the format features")]
use litchi_ops::{BatchItem, Code, File, NewNode, create, open_bytes, props};

#[allow(dead_code, reason = "used by a subset of the format features")]
fn record(pairs: &[(&str, &str)]) -> Record {
    pairs.iter().copied().collect()
}

#[allow(dead_code, reason = "used by a subset of the format features")]
fn text(file: &File, path: &str) -> String {
    file.get(path)
        .unwrap()
        .view()
        .text()
        .unwrap()
        .map(|text| text.into_owned())
        .unwrap_or_default()
}

/// Reopen `file` from its published bytes.
#[allow(dead_code, reason = "used by a subset of the format features")]
fn reopen(file: &File) -> File {
    open_bytes(file.to_bytes().unwrap()).unwrap()
}

#[allow(dead_code, reason = "used by a subset of the format features")]
fn find(file: &File, target: &str, pattern: &str, replacement: &str) -> (File, usize) {
    let mut item = BatchItem::new("set");
    item.path = Some(target.into());
    item.props = Some(
        litchi_ops::Props::new()
            .with("find", pattern)
            .with("replace", replacement),
    );
    let commit = file.apply(&item).unwrap();
    (commit.file, commit.value.matched().unwrap())
}

/// The value of format property `key` of the node at `path`.
#[allow(dead_code, reason = "used by a subset of the format features")]
fn prop(file: &File, path: &str, key: &str) -> Option<String> {
    let node = file.get(path).unwrap();
    let view = node.view();
    view.format()
        .unwrap()
        .find(|prop| prop.key() == key)
        .map(|prop| prop.to_string())
}

#[cfg(feature = "doc")]
mod doc {
    use std::io::Cursor;

    use litchi_doc::{CharacterFormatting, ParagraphFormatting};

    use super::*;

    fn runs_doc() -> File {
        let mut builder = litchi_doc::Builder::new();
        let bold = CharacterFormatting {
            bold: Some(true),
            ..CharacterFormatting::default()
        };
        builder
            .add_paragraph_runs(
                vec![
                    ("Dear {{na".to_owned(), CharacterFormatting::default()),
                    ("me}}, hello world".to_owned(), bold),
                ],
                ParagraphFormatting::default(),
            )
            .unwrap();
        builder.add_paragraph("hello again").unwrap();
        let mut output = Cursor::new(Vec::new());
        builder.build().unwrap().write_to(&mut output).unwrap();
        open_bytes(output.into_inner()).unwrap()
    }

    #[test]
    fn find_replaces_runs_and_keeps_their_formatting() {
        let file = runs_doc();
        let (after, matched) = find(&file, "/", "hello", "bye");
        assert_eq!(matched, 2);
        let after = reopen(&after);
        assert_eq!(text(&after, "/body/p[1]"), "Dear {{name}}, bye world");
        assert_eq!(text(&after, "/body/p[2]"), "bye again");
        assert_eq!(
            prop(&after, "/body/p[1]/r[2]", "bold").as_deref(),
            Some("true")
        );
        // The original is untouched; a scoped find touches only its scope.
        assert_eq!(text(&file, "/body/p[2]"), "hello again");
        let (scoped, matched) = find(&file, "/body/p[2]", "hello", "bye");
        assert_eq!(matched, 1);
        assert_eq!(text(&scoped, "/body/p[1]"), "Dear {{name}}, hello world");
        // No match: an exact no-op.
        let (same, matched) = find(&file, "/", "absent", "x");
        assert_eq!(matched, 0);
        assert_eq!(same.to_bytes().unwrap(), file.to_bytes().unwrap());
        // Formatting needs a run split.
        let error = file
            .set("/", props! { "find" => "hello", "bold" => true })
            .unwrap_err();
        assert_eq!(error.code(), Code::UnsupportedProperty);
    }

    #[test]
    fn merge_spans_runs() {
        let file = runs_doc();
        let commit = file
            .merge(&record(&[("name", "Ada")]), &Options::new())
            .unwrap();
        assert_eq!(commit.value.replacements(), 1);
        let after = reopen(&commit.file);
        assert_eq!(text(&after, "/body/p[1]"), "Dear Ada, hello world");
        let none = file
            .merge(&record(&[("other", "x")]), &Options::new())
            .unwrap();
        assert!(none.patch.is_noop());
        assert_eq!(none.value.unresolved(), ["name"]);
        let empty = file
            .merge(&Record::new(), &Options::new().missing(Missing::Empty))
            .unwrap();
        assert_eq!(text(&empty.file, "/body/p[1]"), "Dear , hello world");
    }
}

#[cfg(feature = "rtf")]
mod rtf {
    use super::*;

    fn template() -> File {
        open_bytes(
            br"{\rtf1\ansi{\fonttbl{\f0 Arial;}}\f0 Dear \{\{na{\b me\}\}, hello}\par hello again\par}"
                .to_vec(),
        )
        .unwrap()
    }

    #[test]
    fn merge_fills_runs() {
        let file = template();
        let before = text(&file, "/body/p[1]");
        assert!(before.contains("{{name}}"), "{before}");
        let commit = file
            .merge(&record(&[("name", "Ada")]), &Options::new())
            .unwrap();
        assert_eq!(commit.value.replacements(), 1);
        let after = reopen(&commit.file);
        assert_eq!(text(&after, "/body/p[1]"), "Dear Ada, hello");
        let none = file.merge(&Record::new(), &Options::new()).unwrap();
        assert!(none.patch.is_noop());
        let error = file
            .merge(&record(&[("name", "a\nb")]), &Options::new())
            .unwrap_err();
        assert_eq!(error.code(), Code::InvalidValue);
    }
}

#[test]
fn formats_are_reported() {
    let _ = FileFormat::Doc;
}
