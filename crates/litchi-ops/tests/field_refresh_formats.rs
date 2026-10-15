//! Field refresh on the legacy and ODF formats (record 0295, record 0279
//! follow-up F1): computable fields rewritten in one native commit,
//! page-number fields kept (`needs_layout`), exact no-ops when current.

#![allow(clippy::unwrap_used, reason = "test assertions")]

#[allow(unused_imports, reason = "used by a subset of the format features")]
use litchi_ops::refresh::{Clock, Options, Status};
#[allow(unused_imports, reason = "used by a subset of the format features")]
use litchi_ops::{File, open_bytes};

#[allow(dead_code, reason = "used by a subset of the format features")]
fn text(file: &File, path: &str) -> String {
    file.get(path)
        .unwrap()
        .view()
        .text()
        .unwrap()
        .map(std::borrow::Cow::into_owned)
        .unwrap_or_default()
}

#[allow(dead_code, reason = "used by a subset of the format features")]
fn clock() -> Options {
    Options::new().clock(Clock::parse("2026-10-15T09:30:00").unwrap())
}

#[cfg(feature = "doc")]
mod doc {
    use std::io::Cursor;

    use litchi_doc::{CharacterFormatting, ParagraphFormatting};

    use super::*;

    /// A paragraph holding one field with `instruction` and `result`.
    fn field(instruction: &str, result: &str) -> Vec<(String, CharacterFormatting)> {
        let special = CharacterFormatting {
            special: Some(true),
            ..CharacterFormatting::default()
        };
        let hidden = CharacterFormatting {
            field_vanish: Some(true),
            ..CharacterFormatting::default()
        };
        vec![
            ("Value: ".to_owned(), CharacterFormatting::default()),
            ("\u{13}".to_owned(), special.clone()),
            (instruction.to_owned(), hidden),
            ("\u{14}".to_owned(), special.clone()),
            (result.to_owned(), CharacterFormatting::default()),
            ("\u{15}".to_owned(), special),
        ]
    }

    fn document() -> File {
        let mut builder = litchi_doc::Builder::new();
        for (instruction, result) in [
            (" DATE \\@ \"yyyy-MM-dd\" ", "old date"),
            (" PAGE ", "7"),
            (" NUMWORDS ", "0"),
            (" SEQ Figure ", "9"),
        ] {
            builder
                .add_paragraph_runs(field(instruction, result), ParagraphFormatting::default())
                .unwrap();
        }
        let mut output = Cursor::new(Vec::new());
        builder.build().unwrap().write_to(&mut output).unwrap();
        open_bytes(output.into_inner()).unwrap()
    }

    #[test]
    fn computable_fields_are_rewritten_once() {
        let file = document();
        // Without a clock, DATE needs one; the others still compute.
        let report = file.refresh(&Options::new()).unwrap();
        let statuses: Vec<(String, Status)> = report
            .value
            .fields()
            .iter()
            .map(|field| (field.kind().to_owned(), field.status()))
            .collect();
        assert_eq!(
            statuses,
            [
                ("DATE".to_owned(), Status::NeedsClock),
                ("PAGE".to_owned(), Status::NeedsLayout),
                ("NUMWORDS".to_owned(), Status::Refreshed),
                ("SEQ".to_owned(), Status::Refreshed),
            ]
        );
        let commit = file.refresh(&clock()).unwrap();
        assert_eq!(commit.value.refreshed(), 3);
        assert_eq!(commit.diagnostics.commits(), 1);
        let after = open_bytes(commit.file.to_bytes().unwrap()).unwrap();
        // The paragraph text shows the field: instruction and result.
        assert!(text(&after, "/body/p[1]").ends_with("\u{14}2026-10-15\u{15}"));
        assert!(text(&after, "/body/p[2]").ends_with("\u{14}7\u{15}"));
        // Counted before the refresh: "Value:" four times and the old
        // results ("old date", "7", "0", "9"), instructions left out.
        let words = text(&after, "/body/p[3]");
        assert!(words.ends_with("\u{14}9\u{15}"), "{words:?}");
        assert!(text(&after, "/body/p[4]").ends_with("\u{14}1\u{15}"));
        // The new DATE result has one word where the old had two: NUMWORDS
        // changes once more, then every result is current, an exact no-op.
        let again = after.refresh(&clock()).unwrap();
        assert_eq!(again.value.refreshed(), 1);
        let current = again.file.refresh(&clock()).unwrap();
        assert!(current.patch.is_noop(), "{:?}", current.value);
        assert_eq!(current.value.count(Status::Unchanged), 3);
        let restored = commit.patch.inverse().apply(&commit.file).unwrap();
        assert_eq!(restored.to_bytes().unwrap(), file.to_bytes().unwrap());
    }
}

#[cfg(feature = "odt")]
mod odt {
    use litchi_ops::props;

    use super::*;

    /// odfdo's `variable.odt`: fixed date and time fields, a title field,
    /// and page fields.
    fn fixture() -> File {
        open_bytes(
            std::fs::read(
                std::path::PathBuf::from(env!("CARGO_MANIFEST_DIR"))
                    .join("../../test-data/odfdo/tests/samples/variable.odt"),
            )
            .unwrap(),
        )
        .unwrap()
    }

    fn statuses(summary: &litchi_ops::refresh::Summary) -> Vec<(String, Status)> {
        summary
            .fields()
            .iter()
            .map(|field| (field.kind().to_owned(), field.status()))
            .collect()
    }

    #[test]
    fn metadata_fields_follow_the_document_and_page_fields_need_layout() {
        let file = fixture();
        // The stored title is current: nothing to write.
        let report = file.refresh(&clock()).unwrap();
        assert!(report.patch.is_noop());
        let expected = [
            ("DATE".to_owned(), Status::Locked),
            ("TIME".to_owned(), Status::Locked),
            ("TITLE".to_owned(), Status::Unchanged),
            ("PAGE-NUMBER".to_owned(), Status::NeedsLayout),
            ("PAGE-COUNT".to_owned(), Status::NeedsLayout),
        ];
        let found = statuses(&report.value);
        for entry in &expected {
            assert!(found.contains(entry), "{entry:?} in {found:?}");
        }
        // A new title is written into the title field, once.
        let retitled = file
            .set("/metadata", props! { "title" => "Rapport" })
            .unwrap();
        let commit = retitled.refresh(&clock()).unwrap();
        assert_eq!(commit.value.refreshed(), 1);
        let after = open_bytes(commit.file.to_bytes().unwrap()).unwrap();
        let text = after.view(litchi_ops::view::Mode::Text).unwrap();
        assert!(text.contains("Rapport"), "{text}");
        assert!(!text.contains("Intitulé"), "{text}");
        assert!(after.refresh(&clock()).unwrap().patch.is_noop());
        let restored = commit.patch.inverse().apply(&commit.file).unwrap();
        assert_eq!(restored.to_bytes().unwrap(), retitled.to_bytes().unwrap());
    }
}
