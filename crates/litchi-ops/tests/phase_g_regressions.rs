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
            .file;
        let mut item = BatchItem::new("set");
        item.path = Some(LINK.into());
        item.props = Some(Props::new().with("url", "https://example.test/new"));
        let batch = file.batch([item], BatchMode::Atomic).unwrap().file;
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
            .file;
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
            .file;
        assert_eq!(edit.to_bytes().unwrap(), set.to_bytes().unwrap());
    }
}
