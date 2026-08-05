use litchi_docx::Package;
use std::io::Cursor;

#[test]
fn removing_notes_strips_references_and_keeps_ids_unique() {
    let mut package = Package::new().unwrap();
    {
        let doc = package.document_mut().unwrap();
        doc.add_paragraph_with_text("body");
        let (first, _) = doc.add_footnote();
        let (second, second_note) = doc.add_footnote();
        second_note.add_paragraph().add_run().set_text("kept note");
        doc.add_endnote();
        assert_eq!((first, second), (1, 2));

        let paragraph = doc.paragraph(0).unwrap();
        paragraph.add_run().add_footnote_reference(first);
        paragraph.add_run().add_footnote_reference(second);

        let removed = doc.remove_footnote(first).unwrap();
        assert_eq!(removed.id(), 1);
        assert!(doc.has_footnotes());
        assert!(doc.remove_footnote(99).is_err());

        // References to the removed note are stripped; the kept one stays.
        let xml = doc.to_xml().unwrap();
        assert!(!xml.contains(r#"<w:footnoteReference w:id="1"/>"#));
        assert!(xml.contains(r#"<w:footnoteReference w:id="2"/>"#));

        // New notes never reuse an ID, even after removals.
        let (third, _) = doc.add_footnote();
        assert_eq!(third, 3);

        let removed_endnote = doc.remove_endnote(1).unwrap();
        assert_eq!(removed_endnote.id(), 1);
        assert!(!doc.has_endnotes());
    }

    let mut bytes = Cursor::new(Vec::new());
    package.to_stream(&mut bytes).unwrap();
    bytes.set_position(0);
    let reopened = Package::from_reader(bytes).unwrap();
    let document = reopened.document().unwrap();
    assert_eq!(document.footnote_count().unwrap(), 2);
    assert_eq!(document.endnote_count().unwrap(), 0);
    let footnotes = document.footnotes().unwrap();
    assert!(footnotes[0].text().unwrap().contains("kept note"));
}

#[test]
fn removing_comments_updates_the_comments_part() {
    let mut package = Package::new().unwrap();
    {
        let doc = package.document_mut().unwrap();
        doc.add_paragraph_with_text("annotated");
        let (first, _) = doc.add_comment("Alice", "drop me");
        doc.add_comment("Bob", "keep me");

        let removed = doc.remove_comment(first).unwrap();
        assert_eq!(removed.author(), "Alice");
        assert_eq!(doc.comment_count(), 1);
        assert!(doc.remove_comment(42).is_err());

        // Comment IDs stay unique after removals.
        let (third, _) = doc.add_comment("Carol", "fresh");
        assert_eq!(third, 3);
    }

    let mut bytes = Cursor::new(Vec::new());
    package.to_stream(&mut bytes).unwrap();
    bytes.set_position(0);
    let reopened = Package::from_reader(bytes).unwrap();
    let document = reopened.document().unwrap();
    assert_eq!(document.comment_count().unwrap(), 2);
    let comments = document.comments().unwrap();
    let authors: Vec<&str> = comments.iter().map(|comment| comment.author()).collect();
    assert!(authors.contains(&"Bob"));
    assert!(authors.contains(&"Carol"));
    assert!(!authors.contains(&"Alice"));
}
