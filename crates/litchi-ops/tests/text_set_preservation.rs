//! No silent content loss on `text` sets (record 0304).
//!
//! Setting the `text` of a container (paragraph, run, shape, cell, note)
//! replaces its text. Content inside the container that is not text — a
//! picture, a frame, a field, a bookmark, a comment anchor, a note
//! reference — must either survive the set or make it fail with a typed
//! error naming it. Each module sets text on a container holding such
//! content in one format and asserts which of the two happens.

#![allow(clippy::unwrap_used, reason = "tests")]

#[cfg(any(
    feature = "odt",
    feature = "docx",
    feature = "pptx",
    feature = "xlsx",
    feature = "xlsb",
    feature = "ods",
    feature = "doc",
    feature = "rtf",
    feature = "ppt",
    feature = "odp",
    feature = "odg",
    feature = "pages",
    feature = "keynote",
    feature = "numbers"
))]
#[allow(dead_code, reason = "each format uses a subset of the helpers")]
mod support {
    use litchi_ops::{File, open_bytes};

    pub(crate) fn fixture(relative: &str) -> Vec<u8> {
        let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data")
            .join(relative);
        std::fs::read(&path).unwrap_or_else(|error| panic!("{}: {error}", path.display()))
    }

    pub(crate) fn open(relative: &str) -> File {
        open_bytes(fixture(relative)).unwrap_or_else(|error| panic!("{relative}: {error}"))
    }

    /// Publish and reopen, so the assertions read the saved bytes.
    pub(crate) fn reopen(file: &File) -> File {
        open_bytes(file.to_bytes().unwrap()).unwrap()
    }

    pub(crate) fn text(file: &File, path: &str) -> String {
        file.get(path)
            .unwrap_or_else(|error| panic!("{path}: {error}"))
            .view()
            .text()
            .unwrap()
            .unwrap_or_default()
            .into_owned()
    }

    /// One member of a ZIP package, as text.
    pub(crate) fn member(file: &File, name: &str) -> String {
        use std::io::Read as _;
        let bytes = file.to_bytes().unwrap();
        let mut archive = zip::ZipArchive::new(std::io::Cursor::new(bytes)).unwrap();
        let mut entry = archive.by_name(name).unwrap();
        let mut text = String::new();
        entry.read_to_string(&mut text).unwrap();
        text
    }

    pub(crate) fn count(file: &File, query: &str) -> usize {
        file.query(query)
            .unwrap_or_else(|error| panic!("{query}: {error}"))
            .count()
    }
}

/// ODT, OTH and ODM share the ODT body edit: frames, notes, bookmarks and
/// other elements without text keep their character offset; text fields are
/// refused.
#[cfg(feature = "odt")]
mod odt_family {
    use litchi_ops::{Code, props};

    use super::support::{count, member, open, reopen, text};

    #[test]
    fn odt_paragraph_text_keeps_its_frames() {
        let file = open("odf/corpus/writer-images-frames.odt");
        let frames = count(&file, "frame");
        assert_eq!(frames, 5);
        let old = text(&file, "/body/p[2]");
        let changed = reopen(
            &file
                .set("/body/p[2]", props! { "text" => "New text" })
                .unwrap(),
        );
        assert_eq!(text(&changed, "/body/p[2]"), "New text");
        assert_eq!(count(&changed, "frame"), frames);
        assert_eq!(count(&changed, "/body/p[2]/frame"), 1);
        // Setting the old text back restores the source content.
        let restored = changed.set("/body/p[2]", props! { "text" => old }).unwrap();
        assert_eq!(
            member(&restored, "content.xml"),
            member(&file, "content.xml")
        );
    }

    #[cfg(all(feature = "oth", feature = "odm"))]
    #[test]
    fn oth_and_odm_paragraph_text_keeps_inline_frames() {
        for name in ["unified-ops/0294/frames.oth", "unified-ops/0294/frames.odm"] {
            let file = open(name);
            assert_eq!(count(&file, "/body/p[1]/frame"), 3, "{name}");
            let changed = reopen(
                &file
                    .set("/body/p[1]", props! { "text" => "Other text" })
                    .unwrap(),
            );
            assert_eq!(text(&changed, "/body/p[1]"), "Other text", "{name}");
            assert_eq!(count(&changed, "/body/p[1]/frame"), 3, "{name}");
            assert!(
                member(&changed, "content.xml").contains("Inside the box"),
                "{name}"
            );
            // The frames sit after "Frames", at offset 6, in both texts.
            let restored = changed
                .set("/body/p[1]", props! { "text" => "Frames" })
                .unwrap();
            assert_eq!(
                member(&restored, "content.xml"),
                member(&file, "content.xml"),
                "{name}: the old text restores the source"
            );
        }
    }

    #[test]
    fn odt_paragraph_text_refuses_text_fields() {
        let file = open("odfdo/tests/samples/variable.odt");
        // "Today is <text:date>…</text:date> <text:time>…</text:time>."
        let error = file
            .set("/body/p[1]", props! { "text" => "Today" })
            .unwrap_err();
        assert_eq!(error.code(), Code::UnsupportedCapability);
        assert!(error.message().contains("text:date"), "{}", error.message());
        // The same refusal inside a batch leaves the snapshot untouched.
        let before = file.to_bytes().unwrap();
        let result = file.edit(|tx| {
            tx.set("/body/p[1]", props! { "text" => "Today" })?;
            Ok(())
        });
        assert!(result.is_err());
        assert_eq!(file.to_bytes().unwrap(), before);
        // A paragraph whose only element is a zero-text reference mark is
        // set, and the mark is kept; the field after it is refused.
        let error = file
            .set("/body/p[4]", props! { "text" => "Page" })
            .unwrap_err();
        assert!(
            error.message().contains("text:reference-ref"),
            "{}",
            error.message()
        );
    }
}
