//! Unified reads and writes of the Apple iWork formats (record 0296, owner
//! decision D1): Pages documents, Keynote presentations and Numbers
//! workbooks open as `File::Pages`, `File::Keynote` and `File::Numbers`, read
//! through the iWork object-tree adapter and write through the leaves'
//! focused transactions.

#![cfg(any(feature = "pages", feature = "keynote", feature = "numbers"))]
#![cfg_attr(
    not(all(feature = "pages", feature = "keynote", feature = "numbers")),
    allow(dead_code, unused_imports)
)]

use litchi_core::FileFormat;
use litchi_ops::capabilities::{Capabilities, Support, Verb};
use litchi_ops::view::Mode;
use litchi_ops::{Code, File, Placement, create, open, open_bytes, props};

fn fixture(path: &str) -> std::path::PathBuf {
    std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/iwork")
        .join(path)
}

fn read(path: &str) -> File {
    open_bytes(std::fs::read(fixture(path)).expect("read fixture")).expect("open")
}

fn text(file: &File, path: &str) -> Option<String> {
    file.get(path)
        .expect("path")
        .view()
        .text()
        .expect("text")
        .map(std::borrow::Cow::into_owned)
}

fn prop(file: &File, path: &str, key: &str) -> Option<String> {
    let node = file.get(path).expect("path");
    let format = node.view().format().expect("format");
    format
        .into_iter()
        .find(|prop| prop.key() == key)
        .map(|prop| prop.to_string())
}

fn reopen(file: &File) -> File {
    open_bytes(file.to_bytes().expect("publish")).expect("reopen")
}

#[cfg(feature = "pages")]
mod pages {
    use super::*;

    #[test]
    fn opens_as_pages_from_bytes_and_path() {
        let file = read("pages/basic.pages");
        assert!(matches!(file, File::Pages(_)));
        assert_eq!(file.format(), Some(FileFormat::Pages));
        let from_path = open(fixture("pages/basic.pages")).expect("open path");
        assert!(matches!(from_path, File::Pages(_)));
        assert_eq!(from_path.text().unwrap(), file.text().unwrap());
    }

    #[test]
    fn reads_body_paragraphs_and_sections() {
        let file = read("pages/body-sections-unicode.pages");
        assert_eq!(prop(&file, "/", "sectionCount").as_deref(), Some("2"));
        assert_eq!(
            text(&file, "/body/p[1]").as_deref(),
            Some("Pages borrowed body 😀")
        );
        assert_eq!(
            text(&file, "/body/p[3]").as_deref(),
            Some("Second section marker — end.")
        );
        assert_eq!(prop(&file, "/body/p[3]", "section").as_deref(), Some("2"));
        assert_eq!(
            text(&file, "/section[2]").as_deref(),
            Some("Second section marker — end.")
        );
        let paragraphs: Vec<_> = file
            .query("p")
            .expect("query")
            .map(|node| node.expect("node").kind())
            .collect();
        assert_eq!(paragraphs.len(), 3);
        assert_eq!(
            file.get("/body/p[9]").unwrap_err().code(),
            Code::PathNotFound
        );
        let outline = file.view(Mode::Outline).expect("outline");
        assert!(outline.contains("/section[2]"), "{outline}");
    }

    #[test]
    fn sets_paragraph_text_and_round_trips() {
        let file = read("pages/body-sections-unicode.pages");
        let edited = file
            .set("/body/p[2]", props! { "text" => "Edited marker ✓" })
            .expect("set");
        let reopened = reopen(&edited);
        assert_eq!(
            text(&reopened, "/body/p[2]").as_deref(),
            Some("Edited marker ✓")
        );
        assert_eq!(
            text(&reopened, "/body/p[1]").as_deref(),
            Some("Pages borrowed body 😀")
        );
        assert_eq!(
            text(&reopened, "/body/p[3]").as_deref(),
            Some("Second section marker — end.")
        );
        // The source snapshot is untouched.
        assert_eq!(
            text(&file, "/body/p[2]").as_deref(),
            Some("First section marker.")
        );
    }

    #[test]
    fn sets_section_text() {
        let file = read("pages/body-sections-unicode.pages");
        let edited = file
            .set("/section[2]", props! { "text" => "Replaced section" })
            .expect("set");
        let reopened = reopen(&edited);
        assert_eq!(
            text(&reopened, "/section[2]").as_deref(),
            Some("Replaced section")
        );
    }

    #[test]
    fn refuses_a_newline_in_paragraph_text() {
        let file = read("pages/basic.pages");
        let error = file
            .set("/body/p[1]", props! { "text" => "two\nlines" })
            .unwrap_err();
        assert_eq!(error.code(), Code::InvalidValue);
    }

    #[test]
    fn exact_noop_set_returns_the_same_bytes() {
        let file = read("pages/basic.pages");
        let current = text(&file, "/body/p[1]").expect("text");
        let edited = file
            .set("/body/p[1]", props! { "text" => current.as_str() })
            .expect("set");
        assert_eq!(edited.to_bytes().unwrap(), file.to_bytes().unwrap());
    }

    #[test]
    fn section_lifecycle_is_a_typed_gap() {
        // Record 0299: creation and paragraph structure are supported; the
        // section lifecycle (request 11) stays a typed gap.
        assert!(create(FileFormat::Pages).is_ok());
        let file = read("pages/basic.pages");
        let error = file
            .add("/", litchi_ops::NewNode::new("section"))
            .unwrap_err();
        assert_eq!(error.code(), Code::UnsupportedCapability);
        let capabilities = Capabilities::of(FileFormat::Pages);
        assert!(capabilities.verb(Verb::Get).is_supported());
        assert!(capabilities.verb(Verb::Set).is_supported());
        assert!(capabilities.verb(Verb::Save).is_supported());
        assert!(!matches!(
            capabilities.verb(Verb::Create),
            Support::Unsupported { .. }
        ));
    }

    #[test]
    fn saves_atomically_to_a_path() {
        let directory = tempfile::tempdir().expect("tempdir");
        let path = directory.path().join("copy.pages");
        std::fs::copy(fixture("pages/basic.pages"), &path).expect("copy");
        let file = open(&path).expect("open");
        let edited = file
            .set("/body/p[1]", props! { "text" => "Saved through File" })
            .expect("set");
        edited.save().expect("save");
        let reopened = open(&path).expect("reopen");
        assert_eq!(
            text(&reopened, "/body/p[1]").as_deref(),
            Some("Saved through File")
        );
    }
}

#[cfg(feature = "keynote")]
mod keynote {
    use super::*;

    #[test]
    fn reads_slides_and_placeholders() {
        let file = read("keynote/basic.key");
        assert!(matches!(file, File::Keynote(_)));
        assert_eq!(prop(&file, "/", "slideCount").as_deref(), Some("1"));
        assert_eq!(
            text(&file, "/slide[1]/placeholder[title]").as_deref(),
            Some("Litchi native Keynote fixture")
        );
        assert_eq!(
            text(&file, "/slide[1]/placeholder[body]").as_deref(),
            Some("Buffa lazy-view migration verification")
        );
        assert_eq!(prop(&file, "/slide[1]", "hidden").as_deref(), Some("false"));
    }

    #[test]
    fn sets_title_body_and_hidden() {
        let file = read("keynote/basic.key");
        let edited = file
            .set(
                "/slide[1]/placeholder[title]",
                props! { "text" => "New title" },
            )
            .expect("title")
            .set(
                "/slide[1]/placeholder[body]",
                props! { "text" => "New body" },
            )
            .expect("body")
            .set("/slide[1]", props! { "hidden" => true })
            .expect("hidden");
        let reopened = reopen(&edited);
        assert_eq!(
            text(&reopened, "/slide[1]/placeholder[title]").as_deref(),
            Some("New title")
        );
        assert_eq!(
            text(&reopened, "/slide[1]/placeholder[body]").as_deref(),
            Some("New body")
        );
        assert_eq!(
            prop(&reopened, "/slide[1]", "hidden").as_deref(),
            Some("true")
        );
    }

    #[test]
    fn refuses_removing_the_last_slide_and_moves_in_place() {
        let file = read("keynote/basic.key");
        let error = file.remove("/slide[1]").unwrap_err();
        assert_eq!(error.code(), Code::InvalidArgument);
        let error = file
            .move_to(
                "/slide[1]",
                Placement::Position(litchi_core::Position::new(0)),
            )
            .map(|_| ());
        assert!(error.is_ok());
    }
}

#[cfg(feature = "numbers")]
mod numbers {
    use super::*;

    #[test]
    fn reads_sheets_tables_and_cells() {
        let file = read("numbers/basic.numbers");
        assert!(matches!(file, File::Numbers(_)));
        assert_eq!(prop(&file, "/sheet[1]", "name").as_deref(), Some("Sheet 1"));
        assert_eq!(
            prop(&file, "/sheet[1]/table[1]", "rowCount").as_deref(),
            Some("22")
        );
        assert_eq!(
            text(&file, "/sheet[1]/table[1]/B2").as_deref(),
            Some("Litchi native Numbers fixture")
        );
        assert_eq!(
            prop(&file, "/sheet[1]/table[1]/B3", "value").as_deref(),
            Some("42")
        );
        assert_eq!(
            prop(&file, "/sheet[1]/table[1]/B3", "type").as_deref(),
            Some("number")
        );
        let cells = file
            .query("cell")
            .expect("query")
            .collect::<Result<Vec<_>, _>>()
            .expect("cells");
        assert_eq!(cells.len(), 2);
        assert_eq!(cells[0].path().to_string(), "/sheet[1]/table[1]/B2");
    }

    #[test]
    fn sets_stored_and_empty_cells() {
        let file = read("numbers/basic.numbers");
        let edited = file
            .set("/sheet[1]/table[1]/B3", props! { "value" => 7.5 })
            .expect("number");
        let reopened = reopen(&edited);
        assert_eq!(
            prop(&reopened, "/sheet[1]/table[1]/B3", "value").as_deref(),
            Some("7.5")
        );
        let edited = reopened
            .set("/sheet[1]/table[1]/C4", props! { "value" => true })
            .expect("empty cell");
        let reopened = reopen(&edited);
        assert_eq!(
            prop(&reopened, "/sheet[1]/table[1]/C4", "type").as_deref(),
            Some("boolean")
        );
        let error = file
            .set("/sheet[1]/table[1]/Z99", props! { "value" => 1 })
            .unwrap_err();
        assert_eq!(error.code(), Code::PathNotFound);
    }
}
