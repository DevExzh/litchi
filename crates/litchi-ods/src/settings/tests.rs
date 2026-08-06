use super::{Settings, Snapshot};

const XML: &str = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content
    xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
    xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
    xmlns:vendor="urn:example:vendor">
  <office:body><office:spreadsheet>
    <vendor:extension vendor:flag="keep"><vendor:value>opaque</vendor:value></vendor:extension>
    <table:calculation-settings table:case-sensitive="false"/>
    <table:table table:name="Data"/>
  </office:spreadsheet></office:body>
</office:document-content>"#;

#[test]
fn no_op_commit_borrows_original_xml_and_preserves_unknown_content() {
    let snapshot = Snapshot::from_content_xml(XML).unwrap();
    let commit = snapshot.transaction().commit().unwrap();
    assert!(!commit.changed());
    assert_eq!(commit.content_xml(), XML);
    assert!(commit.content_xml().contains("vendor:extension"));
}

#[test]
fn update_replaces_only_owned_element_and_keeps_unknown_siblings() {
    let snapshot = Snapshot::from_content_xml(XML).unwrap();
    let mut transaction = snapshot.transaction();
    transaction
        .editor()
        .update(|settings| {
            settings.case_sensitive = Some(true);
            Ok(())
        })
        .unwrap();
    let commit = transaction.commit().unwrap();
    assert!(commit.changed());
    assert!(commit.content_xml().contains("vendor:extension"));
    assert!(
        commit
            .content_xml()
            .contains("table:case-sensitive=\"true\"")
    );
    let reparsed = Snapshot::from_content_xml(commit.content_xml()).unwrap();
    assert_eq!(reparsed.calculation().unwrap().case_sensitive, Some(true));
}

#[test]
fn insertion_and_removal_are_atomic_and_contextual() {
    let source = XML.replace(
        "    <table:calculation-settings table:case-sensitive=\"false\"/>\n",
        "",
    );
    let snapshot = Snapshot::from_content_xml(&source).unwrap();
    assert!(snapshot.calculation().is_none());

    let mut transaction = snapshot.transaction();
    transaction
        .replace(Settings {
            precision_as_shown: Some(true),
            ..Settings::default()
        })
        .unwrap();
    let inserted = transaction.commit().unwrap().into_owned();
    let inserted_snapshot = Snapshot::from_content_xml(&inserted).unwrap();
    assert_eq!(
        inserted_snapshot.calculation().unwrap().precision_as_shown,
        Some(true)
    );

    let mut removal = inserted_snapshot.transaction();
    removal.editor().remove();
    let removed = removal.commit().unwrap();
    assert!(removed.changed());
    assert!(!removed.content_xml().contains("calculation-settings"));
    assert!(removed.content_xml().contains("vendor:extension"));
}

#[test]
fn empty_spreadsheet_hosts_a_new_calculation_element() {
    let source = r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><o:body><o:spreadsheet/></o:body></o:document-content>"#;
    let snapshot = Snapshot::from_content_xml(source).unwrap();
    let mut transaction = snapshot.transaction();
    transaction.replace(Settings::default()).unwrap();
    let output = transaction.commit().unwrap().into_owned();
    assert!(output.contains("<o:spreadsheet><table:calculation-settings"));
    assert!(output.contains("</o:spreadsheet>"));
    assert!(
        Snapshot::from_content_xml(&output)
            .unwrap()
            .calculation()
            .is_some()
    );
}

#[test]
fn invalid_typed_replacement_does_not_mutate_transaction() {
    let snapshot = Snapshot::from_content_xml(XML).unwrap();
    let original = snapshot.calculation().cloned();
    let mut transaction = snapshot.transaction();
    let result = transaction.editor().update(|settings| {
        settings
            .iteration
            .get_or_insert_with(Default::default)
            .maximum_difference = Some("not-a-double".to_string());
        Ok(())
    });
    assert!(result.is_err());
    assert_eq!(transaction.calculation(), original.as_ref());
}
