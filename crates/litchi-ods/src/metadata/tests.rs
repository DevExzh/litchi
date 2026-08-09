use super::{Snapshot, Transaction};

const META: &str = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-meta
    xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
    xmlns:dc="http://purl.org/dc/elements/1.1/"
    xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0"
    xmlns:vendor="urn:example:vendor">
  <office:meta>
    <vendor:opaque vendor:flag="keep"><vendor:value>untouched</vendor:value></vendor:opaque>
    <dc:title>Before</dc:title>
    <dc:creator>Author</dc:creator>
  </office:meta>
</office:document-meta>"#;

#[test]
fn no_op_metadata_transaction_preserves_source_exactly() {
    let snapshot = Snapshot::from_source(Some(META.to_string()))
        .expect("test fixture or operation should succeed");
    let commit = snapshot
        .transaction()
        .commit()
        .expect("test fixture or operation should succeed");
    assert!(!commit.changed());
    assert_eq!(commit.xml(), Some(META));
    assert_eq!(commit.metadata().title.as_deref(), Some("Before"));
}

#[test]
fn metadata_patch_preserves_bounded_unknown_markup() {
    let snapshot = Snapshot::from_source(Some(META.to_string()))
        .expect("test fixture or operation should succeed");
    let mut transaction = snapshot.transaction();
    transaction
        .editor()
        .set_title("After")
        .expect("test fixture or operation should succeed");
    let commit = transaction
        .commit()
        .expect("test fixture or operation should succeed");
    let xml = commit
        .xml()
        .expect("test fixture or operation should succeed");
    assert!(xml.contains("vendor:opaque"));
    assert!(xml.contains("vendor:flag=\"keep\""));
    assert!(xml.contains("<dc:title>After</dc:title>"));
    assert_eq!(
        Snapshot::from_source(Some(xml.to_string()))
            .expect("test fixture or operation should succeed")
            .value()
            .title
            .as_deref(),
        Some("After")
    );
}

#[test]
fn unsupported_metadata_fields_fail_before_publication() {
    let snapshot = Snapshot::from_source(Some(META.to_string()))
        .expect("test fixture or operation should succeed");
    let mut transaction = snapshot.transaction();
    let mut value = transaction.metadata().clone();
    value.identifier = Some("id".to_string());
    assert!(transaction.replace(value).is_err());
    assert_eq!(transaction.metadata().identifier, None);
}

#[test]
fn metadata_can_be_created_and_explicitly_removed() {
    let snapshot = Snapshot::from_source(None).expect("test fixture or operation should succeed");
    let mut transaction = snapshot.transaction();
    transaction
        .editor()
        .set_title("New")
        .expect("test fixture or operation should succeed");
    let created = transaction
        .commit()
        .expect("test fixture or operation should succeed")
        .into_owned_xml()
        .expect("test fixture or operation should succeed");
    assert!(created.contains("<dc:title>New</dc:title>"));

    let snapshot =
        Snapshot::from_source(Some(created)).expect("test fixture or operation should succeed");
    let mut transaction: Transaction<'_> = snapshot.transaction();
    transaction.remove();
    let removed = transaction
        .commit()
        .expect("test fixture or operation should succeed");
    assert!(removed.changed());
    assert_eq!(removed.xml(), None);
}
