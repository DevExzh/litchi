use litchi_core::Position;
use litchi_docx::{Package, document};
use litchi_opc::PackURI;

fn source() -> Package {
    let bytes = include_bytes!(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../test-data/ooxml/docx/documentProperties.docx"
    ));
    Package::from_reader(bytes.as_slice()).unwrap()
}

#[test]
fn commit_retains_complete_package_and_chains_without_reopening() {
    fn assert_shared<T: Clone + Send + Sync + 'static>() {}
    assert_shared::<document::Commit>();
    let source = source();
    let original = source.document_snapshot().unwrap();
    let core = PackURI::new("/docProps/core.xml").unwrap();
    let properties = source
        .opc_package()
        .get_part(&core)
        .unwrap()
        .blob()
        .to_vec();
    let mut edit = source.edit().unwrap();
    edit.replace_paragraph_text(Position::new(0), "private authored first")
        .unwrap();
    let mut fork = edit.try_clone().unwrap();
    fork.replace_paragraph_text(Position::new(0), "independent fork")
        .unwrap();
    assert_eq!(
        source.document_snapshot().unwrap().xml_bytes(),
        original.xml_bytes()
    );
    assert!(!format!("{edit:?}").contains("private authored first"));
    drop(source);

    let commit: document::Commit = edit.commit().unwrap();
    assert_eq!(
        commit.package().document_snapshot().unwrap().xml_bytes(),
        commit.patch().target().xml_bytes()
    );
    assert!(commit.diagnostics().changed());
    assert_eq!(commit.diagnostics().operations(), 1);
    assert!(!format!("{commit:?}").contains("private authored first"));
    assert_eq!(
        commit
            .package()
            .opc_package()
            .get_part(&core)
            .unwrap()
            .blob(),
        properties
    );
    assert_eq!(
        commit
            .package()
            .document_snapshot()
            .unwrap()
            .paragraph(Position::new(0))
            .unwrap()
            .text()
            .unwrap(),
        "private authored first"
    );
    let shared = commit.clone();
    let (first, patch) = commit.into_parts();
    assert!(std::ptr::eq(
        first.opc_package(),
        shared.package().opc_package()
    ));
    let restored = first.apply(&patch.inverse()).unwrap().into_package();
    assert_eq!(
        restored.document_snapshot().unwrap().xml_bytes(),
        original.xml_bytes()
    );
    let mut next = first.edit().unwrap();
    next.replace_paragraph_text(Position::new(0), "second edit")
        .unwrap();
    let next = next.commit().unwrap().into_package();
    assert_eq!(
        next.document_snapshot()
            .unwrap()
            .paragraph(Position::new(0))
            .unwrap()
            .text()
            .unwrap(),
        "second edit"
    );
    assert_eq!(
        first
            .document_snapshot()
            .unwrap()
            .paragraph(Position::new(0))
            .unwrap()
            .text()
            .unwrap(),
        "private authored first"
    );
    let fork = fork.commit().unwrap().into_package();
    assert_eq!(
        fork.document_snapshot()
            .unwrap()
            .paragraph(Position::new(0))
            .unwrap()
            .text()
            .unwrap(),
        "independent fork"
    );
}

#[test]
fn no_op_commit_shares_the_owner_and_failed_staging_rolls_back() {
    let source = source();
    let before = source.document_snapshot().unwrap();
    let mut edit = source.edit().unwrap();
    assert!(
        edit.replace_paragraph_text(Position::new(usize::MAX), "invalid")
            .is_err()
    );
    let commit = edit.commit().unwrap();
    assert!(!commit.diagnostics().changed());
    assert_eq!(commit.diagnostics().operations(), 0);
    assert!(std::ptr::eq(
        source.opc_package(),
        commit.package().opc_package()
    ));
    let mut abandoned = source.edit().unwrap();
    abandoned
        .replace_paragraph_text(Position::new(0), "abandoned")
        .unwrap();
    drop(abandoned);
    assert_eq!(
        source.document_snapshot().unwrap().xml_bytes(),
        before.xml_bytes()
    );
}

#[test]
fn captured_edit_is_independent_of_later_source_writer_mutation() {
    let source = source();
    let original = source.document_snapshot().unwrap();
    let mut edit = source.edit().unwrap();
    edit.replace_paragraph_text(Position::new(0), "captured edit")
        .unwrap();
    let mut later = source.edit().unwrap();
    later
        .insert_paragraph(
            Position::new(original.paragraph_count()),
            "later writer change",
        )
        .unwrap();
    let source = later.commit().unwrap().into_package();
    let commit = edit.commit().unwrap();
    assert_eq!(
        commit
            .package()
            .document_snapshot()
            .unwrap()
            .paragraph_count(),
        original.paragraph_count()
    );
    assert!(
        !commit
            .package()
            .document()
            .unwrap()
            .text()
            .unwrap()
            .contains("later writer change")
    );
    let mut output = Vec::new();
    source.write_to(&mut output).unwrap();
    let reopened = Package::from_reader(output.as_slice()).unwrap();
    assert!(
        reopened
            .document()
            .unwrap()
            .text()
            .unwrap()
            .contains("later writer change")
    );
    assert!(
        !reopened
            .document()
            .unwrap()
            .text()
            .unwrap()
            .contains("captured edit")
    );
    let original_package = Package::from_reader(
        include_bytes!(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../test-data/ooxml/docx/documentProperties.docx"
        ))
        .as_slice(),
    )
    .unwrap();
    let applied = original_package
        .apply(commit.patch())
        .unwrap()
        .into_package();
    assert_eq!(
        applied.document_snapshot().unwrap().xml_bytes(),
        commit.package().document_snapshot().unwrap().xml_bytes()
    );
}
