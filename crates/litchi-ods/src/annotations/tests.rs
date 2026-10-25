//! Focused ODS cell-annotation owner regressions.

use super::*;

const XML: &str = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content
 xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
 xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
 xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
 xmlns:dc="http://purl.org/dc/elements/1.1/"
 xmlns:vendor="urn:example:vendor" office:version="1.3">
 <office:body><office:spreadsheet>
  <vendor:before value="keep"/>
  <table:table table:name="Data">
   <table:table-row>
    <table:table-cell office:value-type="string"><text:p>value</text:p>
      <office:annotation office:name="first"><dc:creator>Ada</dc:creator><text:p>hello <text:span>rich</text:span></text:p></office:annotation>
    </table:table-cell>
    <table:table-cell/>
   </table:table-row>
  </table:table>
  <vendor:after value="keep"/>
 </office:spreadsheet></office:body>
</office:document-content>"#;

fn named(name: &str, text: &str) -> Annotation {
    let mut annotation = Annotation::new(text);
    annotation.set_name(Some(name));
    annotation
}

#[test]
fn parses_rich_annotation_with_contextual_cell_selection() {
    let snapshot = Snapshot::parse(
        XML,
        &litchi_odf_common::XmlStreamLimits::owner_default_value(),
    )
    .expect("test fixture or operation should succeed");
    assert_eq!(snapshot.entries().len(), 1);
    let entry = snapshot
        .cell("Data", 0, 0)
        .expect("test fixture or operation should succeed")
        .expect("test fixture or operation should succeed");
    assert_eq!(entry.cell().sheet(), "Data");
    assert_eq!(entry.cell().row(), 0);
    assert_eq!(entry.cell().column(), 0);
    assert_eq!(entry.annotation().creator().as_deref(), Some("Ada"));
    assert_eq!(entry.annotation().text(), "hello rich");
    assert_eq!(
        snapshot
            .named("first")
            .expect("test fixture or operation should succeed")
            .expect("test fixture or operation should succeed")
            .index(),
        0
    );
}

#[test]
fn no_op_commit_keeps_exact_content_xml_and_empty_patch() {
    let snapshot = Snapshot::parse(
        XML,
        &litchi_odf_common::XmlStreamLimits::owner_default_value(),
    )
    .expect("test fixture or operation should succeed");
    let commit = snapshot
        .edit()
        .commit()
        .expect("test fixture or operation should succeed");
    assert!(!commit.changed());
    assert_eq!(commit.content_xml(), XML);
}

#[test]
fn add_replace_remove_preserve_unrelated_xml() {
    let snapshot = Snapshot::parse(
        XML,
        &litchi_odf_common::XmlStreamLimits::owner_default_value(),
    )
    .expect("test fixture or operation should succeed");
    let mut transaction = snapshot.edit();
    transaction
        .add_at(
            Cell::new("Data", 0, 1).expect("test fixture or operation should succeed"),
            named("second", "added"),
        )
        .expect("test fixture or operation should succeed");
    let added = transaction
        .commit()
        .expect("test fixture or operation should succeed");
    assert!(added.changed());
    assert!(added.content_xml().contains("vendor:before value=\"keep\""));
    assert!(added.content_xml().contains("vendor:after value=\"keep\""));
    assert_eq!(
        added
            .snapshot()
            .cell("Data", 0, 1)
            .expect("test fixture or operation should succeed")
            .expect("test fixture or operation should succeed")
            .annotation()
            .text(),
        "added"
    );

    let mut transaction = added.snapshot().edit();
    transaction
        .replace(0, named("renamed", "changed"))
        .expect("test fixture or operation should succeed");
    let replaced = transaction
        .commit()
        .expect("test fixture or operation should succeed");
    assert!(replaced.content_xml().contains("office:name=\"renamed\""));
    assert!(
        replaced
            .content_xml()
            .contains("vendor:before value=\"keep\"")
    );

    let mut transaction = replaced.snapshot().edit();
    assert_eq!(
        transaction
            .remove(0)
            .expect("test fixture or operation should succeed")
            .text(),
        "changed"
    );
    let removed = transaction
        .commit()
        .expect("test fixture or operation should succeed");
    assert!(
        removed
            .snapshot()
            .cell("Data", 0, 0)
            .expect("test fixture or operation should succeed")
            .is_none()
    );
    assert!(
        removed
            .content_xml()
            .contains("vendor:after value=\"keep\"")
    );
}

#[test]
fn duplicate_names_and_invalid_coordinates_are_failure_atomic() {
    let snapshot = Snapshot::parse(
        XML,
        &litchi_odf_common::XmlStreamLimits::owner_default_value(),
    )
    .expect("test fixture or operation should succeed");
    let mut transaction = snapshot.edit();
    let before = transaction.current().source_xml().to_owned();
    assert!(
        transaction
            .add("Data", 0, 1, named("first", "duplicate"))
            .is_err()
    );
    assert_eq!(transaction.current().source_xml(), before);
    assert!(
        transaction
            .add("Data", usize::MAX, 0, Annotation::new("bad"))
            .is_err()
    );
    assert_eq!(transaction.current().source_xml(), before);
    assert!(
        Snapshot::parse(
            XML,
            &litchi_odf_common::XmlStreamLimits::owner_default_value()
        )
        .expect("test fixture or operation should succeed")
        .cell("Data", 1, 0)
        .is_ok()
    );
}

#[test]
fn caller_xml_budget_covers_annotation_events_depth_and_attributes() {
    let event_limit =
        litchi_odf_common::XmlStreamLimits::new(XML.len() as u64 + 1, 256, 1, 100_000, 64 * 1024)
            .unwrap();
    assert!(Snapshot::parse(XML, &event_limit).is_err());
    let depth_limit = litchi_odf_common::XmlStreamLimits::new(
        XML.len() as u64 + 1,
        4,
        100_000,
        100_000,
        64 * 1024,
    )
    .unwrap();
    assert!(Snapshot::parse(XML, &depth_limit).is_err());
    let attribute_limit =
        litchi_odf_common::XmlStreamLimits::new(XML.len() as u64 + 1, 256, 100_000, 1, 64 * 1024)
            .unwrap();
    assert!(Snapshot::parse(XML, &attribute_limit).is_err());
}
