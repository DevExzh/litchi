use litchi_odt::{Document, mutable::MutableDocument, transaction::ParagraphSelector};

fn source() -> Document {
    let mut document = MutableDocument::new();
    document.add_paragraph("Unique paragraph").unwrap();
    Document::from_bytes(document.to_bytes().unwrap()).unwrap()
}

#[test]
fn packaged_transaction_is_source_checked_reversible_and_exact_for_noop() {
    let source = source();
    let snapshot = litchi_odt::transaction::Snapshot::from_document(&source).unwrap();

    let no_op = snapshot.edit().commit().unwrap();
    assert_eq!(no_op.snapshot().as_bytes(), snapshot.as_bytes());

    let mut edit = snapshot.edit();
    edit.append_line_break(ParagraphSelector::exact_text("Unique paragraph"))
        .unwrap();
    let commit = edit.commit().unwrap();
    assert!(
        commit.snapshot().document().unwrap().paragraphs().unwrap()[0]
            .text()
            .unwrap()
            .contains('\n')
    );
    assert_eq!(
        commit
            .patch()
            .inverse()
            .apply(commit.snapshot())
            .unwrap()
            .as_bytes(),
        snapshot.as_bytes()
    );
    assert!(commit.patch().apply(commit.snapshot()).is_err());
}

#[test]
fn packaged_transaction_rejects_ambiguous_text_selectors() {
    let mut document = MutableDocument::new();
    document.add_paragraph("duplicate").unwrap();
    document.add_paragraph("duplicate").unwrap();
    let source = Document::from_bytes(document.to_bytes().unwrap()).unwrap();
    let snapshot = litchi_odt::transaction::Snapshot::from_document(&source).unwrap();

    assert!(
        snapshot
            .edit()
            .append_line_break(ParagraphSelector::exact_text("duplicate"))
            .is_err()
    );
}
