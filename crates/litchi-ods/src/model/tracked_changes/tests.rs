use super::codec::{
    insert_tracked_change_into_owner, insert_tracked_owner_into_spreadsheet,
    inspect_tracked_changes_source, rewrite_owner_tracking,
};
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
      <table:insertion table:id="i1" table:type="row" table:position="2" table:count="3" table:table="0"><office:change-info><dc:creator>A</dc:creator><dc:date>2026-07-17T00:00:00Z</dc:date><text:p>insert</text:p></office:change-info><table:dependencies><table:dependency table:id="m1"/></table:dependencies></table:insertion>
      <table:deletion table:id="d1" table:type="column" table:position="4" table:acceptance-state="rejected"><office:change-info><dc:creator>D</dc:creator><dc:date>2026-07-17T00:00:00Z</dc:date></office:change-info><table:deletions><table:change-deletion table:id="i1"/><table:cell-content-deletion><table:cell-address table:table="0" table:column="4" table:row="1"/><table:change-track-table-cell office:value-type="float" office:value="12.5" table:formula="of:=1+1"><text:p>12.5</text:p></table:change-track-table-cell></table:cell-content-deletion></table:deletions><table:cut-offs><table:insertion-cut-off table:id="i1" table:position="1"/><table:movement-cut-off table:start-position="2" table:end-position="5"/></table:cut-offs></table:deletion>
      <table:movement table:id="m1"><table:source-range-address table:start-table="0" table:start-column="1" table:start-row="2" table:end-table="0" table:end-column="3" table:end-row="4"/><table:target-range-address table:table="0" table:column="5" table:row="6"/><office:change-info><dc:creator>Mover</dc:creator><dc:date>2026-07-17T00:00:00Z</dc:date></office:change-info></table:movement>
      <table:cell-content-change table:id="c1" table:acceptance-state="accepted"><table:cell-address table:table="0" table:column="1" table:row="2"/><office:change-info><dc:creator>C</dc:creator><dc:date>2026-07-17T00:00:00Z</dc:date><text:p>A &amp;&#x20;B</text:p></office:change-info><table:previous><table:change-track-table-cell office:value-type="string" office:string-value="old" table:matrix-covered="false"><text:p>old</text:p></table:change-track-table-cell></table:previous></table:cell-content-change>
    </table:tracked-changes>"#;
    let tracked = parse(xml)
        .expect("test fixture or operation should succeed")
        .expect("test fixture or operation should succeed");
    assert!(tracked.enabled);
    assert_eq!(tracked.changes.len(), 4);
    let Change::Insertion(insertion) = &tracked.changes[0] else {
        panic!("expected insertion")
    };
    assert_eq!(insertion.count.as_str(), "3");
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
    let empty = parse("<table:tracked-changes/>")
        .expect("test fixture or operation should succeed")
        .expect("test fixture or operation should succeed");
    assert!(!empty.enabled);
    assert!(empty.changes.is_empty());
    let info = "<office:change-info><dc:creator>A</dc:creator><dc:date>2026-07-17T00:00:00Z</dc:date></office:change-info>";
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
    ] {
        assert!(parse(&fragment).is_err(), "accepted {fragment}");
    }
    assert!(parse_tracked_changes(
        r#"<table:tracked-changes xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"/>"#
    )
    .is_err());
}

#[test]
fn rejects_wrong_host_and_odf_child_order() {
    let info = "<office:change-info><dc:creator>A</dc:creator><dc:date>2026-07-17T00:00:00Z</dc:date></office:change-info>";
    for fragment in [
        format!(
            r#"<table:tracked-changes><table:movement table:id="m"><office:change-info><dc:creator>A</dc:creator><dc:date>2026-07-17T00:00:00Z</dc:date></office:change-info><table:source-range-address table:table="0" table:column="0" table:row="0"/><table:target-range-address table:table="0" table:column="1" table:row="1"/></table:movement></table:tracked-changes>"#
        ),
        format!(
            r#"<table:tracked-changes><table:cell-content-change table:id="c">{info}<table:cell-address table:table="0" table:column="0" table:row="0"/><table:previous><table:change-track-table-cell/></table:previous></table:cell-content-change></table:tracked-changes>"#
        ),
        format!(
            r#"<table:tracked-changes><table:insertion table:id="i" table:type="row" table:position="0"><office:change-info><dc:date>2026-07-17T00:00:00Z</dc:date><dc:creator>A</dc:creator></office:change-info></table:insertion></table:tracked-changes>"#
        ),
    ] {
        assert!(parse(&fragment).is_err(), "accepted {fragment}");
    }

    let misplaced =
        format!(r#"{PREFIX}<table:table table:name="S"/><table:tracked-changes/>{SUFFIX}"#,);
    assert!(parse_tracked_changes(&misplaced).is_err());
}

#[test]
fn source_map_preserves_exact_spans_and_marks_rich_records() {
    let fragment = r#"<table:tracked-changes table:track-changes="false"><table:insertion table:id="plain" table:type="row" table:position="0"><office:change-info><dc:creator>A</dc:creator><dc:date>2026-07-17T00:00:00Z</dc:date></office:change-info></table:insertion><table:insertion table:id="rich" table:type="row" table:position="1"><office:change-info><dc:creator>B</dc:creator><dc:date>2026-07-17T00:00:00Z</dc:date><text:p><text:span>rich</text:span></text:p></office:change-info></table:insertion></table:tracked-changes>"#;
    let xml = format!("{PREFIX}{fragment}{SUFFIX}");
    let source = inspect_tracked_changes_source(&xml, &Limits::default())
        .expect("test fixture or operation should succeed");
    assert_eq!(source.records.len(), 2);
    assert_eq!(
        &xml[source.records[0].element.whole.start..source.records[0].element.whole.end],
        &fragment[fragment
            .find("<table:insertion")
            .expect("test fixture or operation should succeed")
            ..fragment
                .find("</table:insertion>")
                .expect("test fixture or operation should succeed")
                + "</table:insertion>".len()]
    );
    assert!(source.records[0].regenerable);
    assert!(source.records[1].has_rich_content);
    assert!(!source.records[1].regenerable);
    let owner = source
        .owner
        .expect("test fixture or operation should succeed");
    let rewritten = rewrite_owner_tracking(&xml, &owner, Some(true), &Limits::default())
        .expect("test fixture or operation should succeed");
    assert!(rewritten.contains("table:track-changes=\"true\""));
}

#[test]
fn semantic_limits_ignore_unrelated_spreadsheet_content() {
    let unrelated = (0..32)
        .map(|index| format!(r#"<table:table table:name="S{index}"><table:table-row><table:table-cell><text:p>{index}</text:p></table:table-cell></table:table-row></table:table>"#))
        .collect::<String>();
    let xml = format!("{PREFIX}{unrelated}{SUFFIX}");
    let limits = Limits::default()
        .with_max_nodes(1)
        .with_max_changes(1)
        .with_max_value_bytes(8)
        .with_max_aggregate_bytes(16);
    let source = inspect_tracked_changes_source(&xml, &limits)
        .expect("test fixture or operation should succeed");
    assert!(source.changes.is_none());

    let tracked = format!(
        "{PREFIX}<table:tracked-changes><table:insertion table:id=\"i\" table:type=\"row\" table:position=\"0\"><office:change-info><dc:creator>A</dc:creator><dc:date>2026-07-17T00:00:00Z</dc:date></office:change-info></table:insertion></table:tracked-changes>{SUFFIX}"
    );
    assert!(inspect_tracked_changes_source(&tracked, &limits).is_err());
}

#[test]
fn opaque_owner_pi_and_foreign_children_are_preserved_and_flagged() {
    let fragment = r#"<table:tracked-changes><?producer keep?><ext:future xmlns:ext="urn:future" ext:value="keep"/></table:tracked-changes>"#;
    let xml = format!(
        "{PREFIX}{fragment}<ext:outside xmlns:ext=\"urn:outside\" xmlns:table=\"urn:unrelated\" table:value-type=\"inert\"/>{SUFFIX}"
    );
    let source = inspect_tracked_changes_source(&xml, &Limits::default())
        .expect("test fixture or operation should succeed");
    assert!(
        source
            .changes
            .as_ref()
            .expect("test fixture or operation should succeed")
            .changes
            .is_empty()
    );
    let owner = source
        .owner
        .expect("test fixture or operation should succeed");
    assert!(owner.has_unsupported_content);
    assert_eq!(
        &xml[owner.element.whole.start..owner.element.whole.end],
        fragment
    );
    assert!(parse_tracked_changes(&xml).is_ok());
    let dtd = format!("<!DOCTYPE x [<!ENTITY e 'boom'>]>{xml}");
    assert!(parse_tracked_changes(&dtd).is_err());
}

#[test]
fn owner_local_record_insertion_expands_self_closing_owner() {
    let xml = format!("{PREFIX}<table:tracked-changes/>{SUFFIX}");
    let source = inspect_tracked_changes_source(&xml, &Limits::default())
        .expect("test fixture or operation should succeed");
    let owner = source
        .owner
        .expect("test fixture or operation should succeed");
    let fragment = r#"<table:insertion table:id="i" table:type="row" table:position="0"><office:change-info><dc:creator>A</dc:creator><dc:date>2026-07-17T00:00:00Z</dc:date></office:change-info></table:insertion>"#;
    let rewritten = insert_tracked_change_into_owner(&xml, &owner, fragment, &Limits::default())
        .expect("test fixture or operation should succeed");
    assert_eq!(
        rewritten,
        format!("<table:tracked-changes>{fragment}</table:tracked-changes>")
    );
}

#[test]
fn owner_insertion_expands_self_closing_spreadsheet_with_exact_qname() {
    let xml = r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:t="urn:oasis:names:tc:opendocument:xmlns:table:1.0"><o:body><o:spreadsheet t:structure-protected="false"/></o:body></o:document-content>"#.to_string();
    let source = inspect_tracked_changes_source(&xml, &Limits::default())
        .expect("test fixture or operation should succeed");
    let owner = r#"<t:tracked-changes xmlns:t="urn:oasis:names:tc:opendocument:xmlns:table:1.0"/>"#;
    let rewritten =
        insert_tracked_owner_into_spreadsheet(&xml, &source.spreadsheet, owner, &Limits::default())
            .expect("test fixture or operation should succeed");
    assert_eq!(
        rewritten,
        format!(r#"<o:spreadsheet t:structure-protected="false">{owner}</o:spreadsheet>"#)
    );
}

#[test]
fn error_and_xsd_double_specials_round_trip_with_exact_attribute_choice() {
    let info = "<office:change-info><dc:creator>A</dc:creator><dc:date>2026-07-17T00:00:00Z</dc:date></office:change-info>";
    for (value_type, attributes, expected) in [
        ("error", r##" office:string-value="#DIV/0!""##, "#DIV/0!"),
        ("float", r#" office:value="INF""#, "INF"),
        ("percentage", r#" office:value="-INF""#, "-INF"),
        (
            "currency",
            r#" office:value="NaN" office:currency="USD""#,
            "NaN",
        ),
    ] {
        let owner = format!(
            r#"<table:tracked-changes><table:cell-content-change table:id="c"><table:cell-address table:table="0" table:column="0" table:row="0"/>{info}<table:previous><table:change-track-table-cell office:value-type="{value_type}"{attributes}/></table:previous></table:cell-content-change></table:tracked-changes>"#
        );
        let changes = parse(&owner)
            .expect("test fixture or operation should succeed")
            .expect("test fixture or operation should succeed");
        let Change::CellContent(change) = &changes.changes[0] else {
            panic!("expected cell-content change")
        };
        match (&change.previous.value, value_type) {
            (CellValue::Error(Some(value)), "error") => assert_eq!(value, expected),
            (CellValue::Number(value), "float") => {
                assert!(value.is_infinite() && value.is_sign_positive())
            },
            (CellValue::Percentage(value), "percentage") => {
                assert!(value.is_infinite() && value.is_sign_negative())
            },
            (CellValue::Currency { value, code }, "currency") => {
                assert!(value.is_nan());
                assert_eq!(code, "USD");
            },
            other => panic!("unexpected parsed value {other:?}"),
        }
        let written = changes
            .to_xml_fragment()
            .expect("test fixture or operation should succeed");
        assert!(written.contains(expected));
    }

    let invalid = format!(
        r#"<table:tracked-changes><table:cell-content-change table:id="c"><table:cell-address table:table="0" table:column="0" table:row="0"/>{info}<table:previous><table:change-track-table-cell office:value-type="float" office:value="1" office:date-value="2026-01-01"/></table:previous></table:cell-content-change></table:tracked-changes>"#
    );
    assert!(parse(&invalid).is_err());
}

#[test]
fn odf_reserved_namespaces_are_schema_controlled_but_no_namespace_is_opaque() {
    let info = "<office:change-info><dc:creator>A</dc:creator><dc:date>2026-07-17T00:00:00Z</dc:date></office:change-info>";
    let draw = r#"<table:tracked-changes><draw:frame xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"/></table:tracked-changes>"#;
    assert!(parse(draw).is_err());

    let style_attribute = format!(
        r#"<table:tracked-changes><table:insertion xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" table:id="i" table:type="row" table:position="0" style:name="unexpected">{info}</table:insertion></table:tracked-changes>"#
    );
    assert!(parse(&style_attribute).is_err());

    let unqualified_required = format!(
        r#"<table:tracked-changes><table:insertion id="i" type="row" position="0">{info}</table:insertion></table:tracked-changes>"#
    );
    assert!(parse(&unqualified_required).is_err());

    let no_namespace = r#"<table:tracked-changes><insertion xmlns="" id="future"><change-info><creator>opaque</creator></change-info></insertion></table:tracked-changes>"#;
    let changes = parse(no_namespace)
        .expect("test fixture or operation should succeed")
        .expect("test fixture or operation should succeed");
    assert!(changes.changes.is_empty());
}

#[test]
fn parses_libreoffice_change_tracking_fixtures() {
    let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../..");
    let tracked_path =
        root.join("test-data/libreoffice-core/sc/qa/unit/data/ods/change-tracking.ods");
    let tracked =
        crate::Spreadsheet::open(tracked_path).expect("test fixture or operation should succeed");
    let changes = tracked
        .tracked_changes()
        .expect("test fixture or operation should succeed");
    assert_eq!(
        changes
            .changes()
            .expect("test fixture or operation should succeed")
            .changes
            .len(),
        2
    );
    assert!(
        changes
            .changes()
            .expect("test fixture or operation should succeed")
            .changes
            .iter()
            .all(|change| matches!(change, Change::CellContent(_)))
    );

    let protected_path = root
        .join("test-data/libreoffice-core/sc/qa/extras/testdocuments/RecordChangesProtected.ods");
    let protected =
        crate::Spreadsheet::open(protected_path).expect("test fixture or operation should succeed");
    assert!(
        protected
            .tracked_changes()
            .expect("test fixture or operation should succeed")
            .changes()
            .expect("test fixture or operation should succeed")
            .changes
            .is_empty()
    );
}
