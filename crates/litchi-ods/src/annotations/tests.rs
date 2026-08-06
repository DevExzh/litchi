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
    let snapshot = Snapshot::parse(XML).unwrap();
    assert_eq!(snapshot.entries().len(), 1);
    let entry = snapshot.cell("Data", 0, 0).unwrap().unwrap();
    assert_eq!(entry.cell().sheet(), "Data");
    assert_eq!(entry.cell().row(), 0);
    assert_eq!(entry.cell().column(), 0);
    assert_eq!(entry.annotation().creator().as_deref(), Some("Ada"));
    assert_eq!(entry.annotation().text(), "hello rich");
    assert_eq!(snapshot.named("first").unwrap().unwrap().index(), 0);
}

#[test]
fn no_op_commit_keeps_exact_content_xml_and_empty_patch() {
    let snapshot = Snapshot::parse(XML).unwrap();
    let commit = snapshot.edit().commit().unwrap();
    assert!(!commit.changed());
    assert_eq!(commit.content_xml(), XML);
    assert!(commit.patch().is_empty());
}

#[test]
fn add_replace_remove_preserve_unrelated_xml() {
    let snapshot = Snapshot::parse(XML).unwrap();
    let mut transaction = snapshot.edit();
    transaction
        .add_at(Cell::new("Data", 0, 1).unwrap(), named("second", "added"))
        .unwrap();
    let added = transaction.commit().unwrap();
    assert!(added.changed());
    assert!(added.content_xml().contains("vendor:before value=\"keep\""));
    assert!(added.content_xml().contains("vendor:after value=\"keep\""));
    assert_eq!(
        added
            .snapshot()
            .cell("Data", 0, 1)
            .unwrap()
            .unwrap()
            .annotation()
            .text(),
        "added"
    );

    let mut transaction = added.snapshot().edit();
    transaction.replace(0, named("renamed", "changed")).unwrap();
    let replaced = transaction.commit().unwrap();
    assert!(replaced.content_xml().contains("office:name=\"renamed\""));
    assert!(
        replaced
            .content_xml()
            .contains("vendor:before value=\"keep\"")
    );

    let mut transaction = replaced.snapshot().edit();
    assert_eq!(transaction.remove(0).unwrap().text(), "changed");
    let removed = transaction.commit().unwrap();
    assert!(removed.snapshot().cell("Data", 0, 0).unwrap().is_none());
    assert!(
        removed
            .content_xml()
            .contains("vendor:after value=\"keep\"")
    );
}

#[test]
fn duplicate_names_and_invalid_coordinates_are_failure_atomic() {
    let snapshot = Snapshot::parse(XML).unwrap();
    let mut transaction = snapshot.edit();
    let before = transaction.snapshot().source_xml().to_owned();
    assert!(
        transaction
            .add("Data", 0, 1, named("first", "duplicate"))
            .is_err()
    );
    assert_eq!(transaction.snapshot().source_xml(), before);
    assert!(
        transaction
            .add("Data", usize::MAX, 0, Annotation::new("bad"))
            .is_err()
    );
    assert_eq!(transaction.snapshot().source_xml(), before);
    assert!(Snapshot::parse(XML).unwrap().cell("Data", 1, 0).is_ok());
}

#[test]
fn patch_is_source_checked_and_reversible_semantically() {
    let snapshot = Snapshot::parse(XML).unwrap();
    let mut transaction = snapshot.edit();
    transaction
        .add("Data", 0, 1, named("second", "added"))
        .unwrap();
    let commit = transaction.commit().unwrap();
    let applied = commit.patch().apply(&snapshot).unwrap();
    assert_eq!(applied.content_xml(), commit.content_xml());

    let inverse = commit.patch().inverse();
    let restored = inverse.apply(commit.snapshot()).unwrap();
    assert!(restored.snapshot().cell("Data", 0, 1).unwrap().is_none());
}
