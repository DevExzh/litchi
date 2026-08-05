// End-to-end cases require the full content parser facade.
#![cfg(any())]

use super::*;
use litchi_core::Result;

const PREFIX: &str = r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:dc="http://purl.org/dc/elements/1.1/"><office:body><office:spreadsheet>"#;
const SUFFIX: &str = "</office:spreadsheet></office:body></office:document-content>";

fn parse(fragment: &str) -> Result<Option<Changes>> {
    parse_tracked_changes(&format!("{PREFIX}{fragment}{SUFFIX}"))
}

#[test]
fn parses_complete_spreadsheet_change_graph() {
    let xml = r#"<table:tracked-changes table:track-changes="true">
      <table:insertion table:id="i1" table:type="row" table:position="2" table:count="3" table:table="0"><office:change-info><dc:creator>A</dc:creator><dc:date>2026-07-17</dc:date><text:p>insert</text:p></office:change-info><table:dependencies><table:dependency table:id="m1"/></table:dependencies></table:insertion>
      <table:deletion table:id="d1" table:type="column" table:position="4" table:acceptance-state="rejected" table:multi-deletion-spanned="2"><office:change-info/><table:deletions><table:change-deletion table:id="i1"/><table:cell-content-deletion><table:cell-address table:table="0" table:column="4" table:row="1"/><table:change-track-table-cell office:value-type="float" office:value="12.5" table:formula="of:=1+1"><text:p>12.5</text:p></table:change-track-table-cell></table:cell-content-deletion></table:deletions><table:cut-offs><table:insertion-cut-off table:id="i1" table:position="1"/><table:movement-cut-off table:start-position="2" table:end-position="5"/></table:cut-offs></table:deletion>
      <table:movement table:id="m1"><table:source-range-address table:start-table="0" table:start-column="1" table:start-row="2" table:end-table="0" table:end-column="3" table:end-row="4"/><table:target-range-address table:table="0" table:column="5" table:row="6"/><office:change-info><dc:creator>Mover</dc:creator></office:change-info></table:movement>
      <table:cell-content-change table:id="c1" table:acceptance-state="accepted"><table:cell-address table:table="0" table:column="1" table:row="2"/><office:change-info><text:p>A &amp;&#x20;B</text:p></office:change-info><table:previous table:id="old"><table:change-track-table-cell office:value-type="string" office:string-value="old" table:matrix-covered="false"><text:p>old</text:p></table:change-track-table-cell></table:previous></table:cell-content-change>
    </table:tracked-changes>"#;
    let tracked = parse(xml).unwrap().unwrap();
    assert!(tracked.enabled);
    assert_eq!(tracked.changes.len(), 4);
    let Change::Insertion(insertion) = &tracked.changes[0] else {
        panic!("expected insertion")
    };
    assert_eq!(insertion.count.get(), 3);
    assert_eq!(insertion.metadata.info.comments, ["insert"]);
    let Change::Deletion(deletion) = &tracked.changes[1] else {
        panic!("expected deletion")
    };
    assert_eq!(deletion.cut_offs.len(), 2);
    assert_eq!(deletion.metadata.deletions.len(), 2);
    let Change::Movement(movement) = &tracked.changes[2] else {
        panic!("expected movement")
    };
    assert!(matches!(movement.source, RangeAddress::Range { .. }));
    let Change::CellContent(change) = &tracked.changes[3] else {
        panic!("expected cell change")
    };
    assert_eq!(change.metadata.acceptance, Acceptance::Accepted);
    assert_eq!(change.metadata.info.comments, ["A & B"]);
    assert_eq!(change.previous.value, CellValue::Text("old".into()));
}

#[test]
fn applies_defaults_and_rejects_malformed_change_graphs() {
    let empty = parse("<table:tracked-changes/>").unwrap().unwrap();
    assert!(!empty.enabled);
    assert!(empty.changes.is_empty());
    let info = "<office:change-info/>";
    for fragment in [
        r#"<table:tracked-changes table:track-changes="yes"/>"#.to_string(),
        format!(
            r#"<table:tracked-changes><table:insertion table:id="x" table:type="bad" table:position="0">{info}</table:insertion></table:tracked-changes>"#
        ),
        format!(
            r#"<table:tracked-changes><table:insertion table:id="x" table:type="row" table:position="0" table:count="0">{info}</table:insertion></table:tracked-changes>"#
        ),
        format!(
            r#"<table:tracked-changes><table:movement table:id="x"><table:source-range-address table:table="0" table:column="1"/><table:target-range-address table:table="0" table:column="1" table:row="1"/>{info}</table:movement></table:tracked-changes>"#
        ),
        format!(
            r#"<table:tracked-changes><table:cell-content-change table:id="x"><table:cell-address table:table="0" table:column="1" table:row="1"/>{info}<table:previous><table:change-track-table-cell office:value="1"/></table:previous></table:cell-content-change></table:tracked-changes>"#
        ),
        format!(
            r#"<table:tracked-changes><table:insertion table:id="x" table:type="row" table:position="0">{info}</table:insertion><table:deletion table:id="x" table:type="row" table:position="0">{info}</table:deletion></table:tracked-changes>"#
        ),
        r#"<table:tracked-changes><fake:insertion xmlns:fake="urn:fake"/></table:tracked-changes>"#
            .to_string(),
    ] {
        assert!(parse(&fragment).is_err(), "accepted {fragment}");
    }
    assert!(parse_tracked_changes(
        r#"<table:tracked-changes xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"/>"#
    )
    .is_err());
}

#[test]
fn parses_libreoffice_change_tracking_fixtures() {
    let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../..");
    let tracked_path =
        root.join("test-data/libreoffice-core/sc/qa/unit/data/ods/change-tracking.ods");
    let tracked = crate::Spreadsheet::open(tracked_path).unwrap();
    let changes = tracked.tracked_changes().unwrap();
    assert_eq!(changes.changes.len(), 2);
    assert!(
        changes
            .changes
            .iter()
            .all(|change| matches!(change, Change::CellContent(_)))
    );

    let protected_path = root
        .join("test-data/libreoffice-core/sc/qa/extras/testdocuments/RecordChangesProtected.ods");
    let protected = crate::Spreadsheet::open(protected_path).unwrap();
    assert!(protected.tracked_changes().unwrap().changes.is_empty());
}
