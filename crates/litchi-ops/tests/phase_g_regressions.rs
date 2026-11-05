//! Regressions for the claimed-supported cells record 0289's CRUD checklist
//! found failing: each was a single verb that worked through `File::set`
//! but not through `File::edit`/`File::batch`, or a capability claim the
//! format does not honour.

#[cfg(feature = "docx")]
mod docx {
    use litchi_core::FileFormat;
    use litchi_ops::{BatchItem, BatchMode, NewNode, Props, create, props};

    /// A producer file opened from bytes: a source-backed owner, whose
    /// formatting run stages on the source editor.
    fn with_link() -> litchi_ops::File {
        litchi_ops::open_bytes(
            std::fs::read(concat!(
                env!("CARGO_MANIFEST_DIR"),
                "/../../test-data/ooxml/docx/Hyperlink.docx"
            ))
            .unwrap(),
        )
        .unwrap()
    }

    const LINK: &str = "/body/p[2]/hyperlink[1]";

    fn url(file: &litchi_ops::File) -> String {
        let node = file.get(LINK).unwrap();
        node.view()
            .format()
            .unwrap()
            .into_iter()
            .find(|prop| prop.key() == "url")
            .map(|prop| prop.to_string())
            .unwrap()
    }

    /// A hyperlink target changes a relationship, which only the owning
    /// package editor publishes; the formatting run must not take it.
    #[test]
    fn hyperlink_targets_change_through_edit_and_batch() {
        let file = with_link();
        let set = file
            .set(LINK, props! { "url" => "https://example.test/new" })
            .unwrap();
        let edit = file
            .edit(|tx| tx.set(LINK, props! { "url" => "https://example.test/new" }))
            .unwrap()
            .into_file();
        let mut item = BatchItem::new("set");
        item.path = Some(LINK.into());
        item.props = Some(Props::new().with("url", "https://example.test/new"));
        let batch = file.batch([item], BatchMode::Atomic).unwrap().into_file();
        for changed in [&set, &edit, &batch] {
            assert_eq!(url(changed), "https://example.test/new");
        }
        assert_eq!(edit.to_bytes().unwrap(), set.to_bytes().unwrap());
        assert_eq!(batch.to_bytes().unwrap(), set.to_bytes().unwrap());
    }

    /// Charts have their own setter; an edit closure reaches it.
    #[test]
    fn chart_sets_run_in_edit_closures() {
        let file = create(FileFormat::Docx)
            .unwrap()
            .add("/body", NewNode::new("paragraph").prop("text", "Lead"))
            .unwrap()
            .file
            .add(
                "/body/p[1]",
                NewNode::new("chart")
                    .prop("chartType", "column")
                    .prop("data", "A:1,2;B:3,4"),
            )
            .unwrap()
            .file;
        let set = file
            .set("/body/p[1]/chart[1]", props! { "title" => "Sales" })
            .unwrap();
        let edit = file
            .edit(|tx| tx.set("/body/p[1]/chart[1]", props! { "title" => "Sales" }))
            .unwrap()
            .into_file();
        assert_eq!(edit.to_bytes().unwrap(), set.to_bytes().unwrap());
    }
}

#[cfg(feature = "doc")]
mod doc {
    use litchi_ops::{open_bytes, props};

    /// Run text takes the ordinary setter path inside an edit closure; the
    /// paragraph-text run admits paragraphs only.
    #[test]
    fn run_text_sets_run_in_edit_closures() {
        let file = open_bytes(
            std::fs::read(concat!(
                env!("CARGO_MANIFEST_DIR"),
                "/../../test-data/ole/doc/documentProperties.doc"
            ))
            .unwrap(),
        )
        .unwrap();
        let set = file
            .set("/body/p[1]/r[1]", props! { "text" => "Run" })
            .unwrap();
        let edit = file
            .edit(|tx| tx.set("/body/p[1]/r[1]", props! { "text" => "Run" }))
            .unwrap()
            .into_file();
        assert_eq!(edit.to_bytes().unwrap(), set.to_bytes().unwrap());
    }
}

#[cfg(feature = "rtf")]
mod rtf {
    use litchi_core::FileFormat;
    use litchi_ops::capabilities::Capabilities;

    /// The body takes added paragraphs; it is not itself an element `add`
    /// creates, so the report must not claim it.
    #[test]
    fn the_body_is_not_reported_addable() {
        let caps = Capabilities::of(FileFormat::Rtf);
        assert!(!caps.element("body").unwrap().operations().add);
        assert!(caps.element("paragraph").unwrap().operations().add);
        assert!(caps.element("run").unwrap().operations().add);
    }
}

#[cfg(all(feature = "odt", feature = "ods", feature = "odp", feature = "odg"))]
mod odf {
    use litchi_ops::{Props, open};

    /// The same write on the same input publishes the same bytes.
    /// LibreOffice packages list manifest directory entries the writer
    /// copies; they came out in hash-map order, so two identical writes
    /// differed.
    #[test]
    fn identical_writes_publish_identical_bytes() {
        for (relative, path, key) in [
            (
                "odf/corpus/writer-definition-lists.odt",
                "/metadata",
                "author",
            ),
            ("odf/corpus/calc-two-sheets.ods", "/metadata", "title"),
            ("odf/corpus/impress-basic.odp", "/metadata", "title"),
            (
                "libreoffice-core/svx/qa/unit/data/FontWork.odg",
                "/page[1]",
                "name",
            ),
        ] {
            let file = open(
                std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
                    .join("../../test-data")
                    .join(relative),
            )
            .unwrap();
            let write = || {
                file.set(path, Props::new().with(key, "Litchi"))
                    .unwrap_or_else(|error| panic!("{relative}: {error}"))
                    .to_bytes()
                    .unwrap()
            };
            let first = write();
            for _ in 0..4 {
                assert!(
                    write() == first,
                    "{relative}: {path} {key} is not deterministic"
                );
            }
        }
    }
}
