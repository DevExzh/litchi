//! Focused protection snapshot and transaction regressions.

use super::*;

fn source() -> String {
    r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content
 xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
 xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
 xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
 xmlns:loext="urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0"
 office:version="1.3">
 <office:automatic-styles>
  <style:style style:name="locked" style:family="table-cell"><style:table-cell-properties style:cell-protect="protected"/></style:style>
  <draw:unknown xmlns:draw="urn:example:draw" draw:name="keep"/>
 </office:automatic-styles>
 <office:body><office:spreadsheet table:structure-protected="false" table:protection-key="doc">
  <table:table table:name="Sheet1" table:protected="true" table:protection-key="sheet">
   <loext:table-protection loext:insert-rows="false" loext:use-pivot="true"/>
  </table:table>
 </office:spreadsheet></office:body>
</office:document-content>"#
        .to_string()
}

#[test]
fn no_op_snapshot_retains_exact_content_xml() {
    let source = source();
    let snapshot = Snapshot::parse(source.clone(), None).unwrap();
    let commit = snapshot.edit().commit().unwrap();
    assert!(!commit.changed());
    assert_eq!(commit.content_xml(), source);
}

#[test]
fn document_and_sheet_edits_are_source_checked_and_atomic() {
    let snapshot = Snapshot::parse(source(), None).unwrap();
    assert_eq!(snapshot.document().structure_protected, Some(false));
    assert_eq!(
        snapshot.sheet("Sheet1").unwrap().permissions.insert_rows,
        Some(false)
    );

    let mut transaction = snapshot.edit();
    transaction.document_mut().structure_protected = Some(true);
    transaction.sheets_mut()[0].protected = Some(false);
    transaction.sheets_mut()[0].permissions.insert_rows = Some(true);
    let commit = transaction.commit().unwrap();
    assert!(commit.changed());
    assert!(
        commit
            .content_xml()
            .contains("table:structure-protected=\"true\"")
    );
    assert!(commit.content_xml().contains("loext:insert-rows=\"true\""));
    assert!(commit.content_xml().contains("draw:name=\"keep\""));
    assert_eq!(
        commit.snapshot().sheet("Sheet1").unwrap().protected,
        Some(false)
    );
}

#[test]
fn style_edits_replace_only_managed_automatic_styles() {
    let snapshot = Snapshot::parse(source(), None).unwrap();
    let mut transaction = snapshot.edit();
    transaction
        .styles_mut()
        .set_automatic(vec![Style::new("locked", Protection::HiddenAndProtected)]);
    let commit = transaction.commit().unwrap();
    assert!(
        commit
            .content_xml()
            .contains("cell-protect=\"hidden-and-protected\"")
    );
    assert!(commit.content_xml().contains("draw:name=\"keep\""));
}

#[test]
fn patch_is_exact_source_checked_and_reversible() -> litchi_core::Result<()> {
    let snapshot = Snapshot::parse(source(), None)?;
    let mut edit = snapshot.edit();
    edit.document_mut().structure_protected = Some(true);
    let commit = edit.commit()?;
    let patch = commit.patch().clone();
    let restored = patch.inverse().apply(commit.snapshot())?;
    assert_eq!(restored.content_xml(), snapshot.source_xml());

    let other = Snapshot::parse(
        "<office:document-content xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" xmlns:table=\"urn:oasis:names:tc:opendocument:xmlns:table:1.0\"><office:body><office:spreadsheet/></office:body></office:document-content>",
        None,
    )?;
    assert!(patch.apply(&other).is_err());
    Ok(())
}
