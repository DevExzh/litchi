use litchi_odf::{
    MutableSpreadsheet, Spreadsheet, SpreadsheetCellContentChange, SpreadsheetChangeAcceptance,
    SpreadsheetChangeCutOff, SpreadsheetChangeDimension, SpreadsheetChangeInfo,
    SpreadsheetChangeMetadata, SpreadsheetDeletion, SpreadsheetInsertion, SpreadsheetMovement,
    SpreadsheetNestedDeletion, SpreadsheetTrackedCell, SpreadsheetTrackedCellAddress,
    SpreadsheetTrackedCellValue, SpreadsheetTrackedChange, SpreadsheetTrackedChanges,
    SpreadsheetTrackedRangeAddress,
};
use std::num::NonZeroUsize;

fn metadata(id: &str) -> SpreadsheetChangeMetadata {
    SpreadsheetChangeMetadata {
        id: id.to_string(),
        acceptance: SpreadsheetChangeAcceptance::Pending,
        rejecting_change_id: None,
        info: SpreadsheetChangeInfo {
            creator: Some("Zoë & 李".to_string()),
            date: Some("2026-07-19T08:00:00Z".to_string()),
            comments: vec!["before < after & \"quoted\"".to_string()],
        },
        dependencies: Vec::new(),
        deletions: Vec::new(),
    }
}

fn address(table: i64, column: i64, row: i64) -> SpreadsheetTrackedCellAddress {
    SpreadsheetTrackedCellAddress { table, column, row }
}

fn previous_cell() -> SpreadsheetTrackedCell {
    SpreadsheetTrackedCell {
        address: Some("Sheet1.A1".to_string()),
        style_name: Some("Historical & Style".to_string()),
        matrix_covered: false,
        formula: Some("of:=SUM([.A1:.A2])".to_string()),
        matrix_columns: NonZeroUsize::new(2),
        matrix_rows: NonZeroUsize::new(1),
        value: SpreadsheetTrackedCellValue::Currency {
            value: 12.5,
            code: "CNY".to_string(),
        },
        display_text: "人民币 & <旧值>\n第二行".to_string(),
    }
}

fn complete_change_set() -> SpreadsheetTrackedChanges {
    let insertion = SpreadsheetTrackedChange::Insertion(SpreadsheetInsertion {
        metadata: metadata("ct1"),
        dimension: SpreadsheetChangeDimension::Row,
        position: 1,
        count: NonZeroUsize::new(2).unwrap(),
        table: Some(0),
    });

    let mut deletion_metadata = metadata("ct2");
    deletion_metadata.acceptance = SpreadsheetChangeAcceptance::Accepted;
    deletion_metadata.dependencies.push("ct1".to_string());
    deletion_metadata
        .deletions
        .push(SpreadsheetNestedDeletion::CellContent {
            change_id: Some("ct1".to_string()),
            address: Some(address(0, 0, 0)),
            cell: Some(previous_cell()),
        });
    deletion_metadata
        .deletions
        .push(SpreadsheetNestedDeletion::Change {
            change_id: Some("ct1".to_string()),
        });
    let deletion = SpreadsheetTrackedChange::Deletion(SpreadsheetDeletion {
        metadata: deletion_metadata,
        dimension: SpreadsheetChangeDimension::Column,
        position: 3,
        table: Some(0),
        multi_deletion_spanned: Some(1),
        cut_offs: vec![
            SpreadsheetChangeCutOff::Insertion {
                change_id: "ct1".to_string(),
                position: 1,
            },
            SpreadsheetChangeCutOff::MovementPoint { position: 2 },
            SpreadsheetChangeCutOff::MovementRange { start: 3, end: 5 },
        ],
    });

    let mut movement_metadata = metadata("ct3");
    movement_metadata.acceptance = SpreadsheetChangeAcceptance::Rejected;
    movement_metadata.dependencies.push("ct2".to_string());
    let movement = SpreadsheetTrackedChange::Movement(SpreadsheetMovement {
        metadata: movement_metadata,
        source: SpreadsheetTrackedRangeAddress::Range {
            start: address(0, 0, 0),
            end: address(0, 2, 3),
        },
        target: SpreadsheetTrackedRangeAddress::Cell(address(0, 4, 5)),
    });

    let mut cell_metadata = metadata("ct4");
    cell_metadata.dependencies.push("ct3".to_string());
    let cell = SpreadsheetTrackedChange::CellContent(SpreadsheetCellContentChange {
        metadata: cell_metadata,
        address: address(0, 6, 7),
        previous_change_id: Some("ct1".to_string()),
        previous: previous_cell(),
    });

    SpreadsheetTrackedChanges {
        enabled: true,
        changes: vec![insertion, deletion, movement, cell],
    }
}

#[test]
fn authors_every_variant_and_reopens_the_package() {
    let changes = complete_change_set();
    let xml = changes.to_xml_fragment().unwrap();
    assert!(xml.starts_with("<table:tracked-changes table:track-changes=\"true\">"));
    assert!(xml.contains("<table:insertion"));
    assert!(xml.contains("<table:deletion"));
    assert!(xml.contains("<table:movement"));
    assert!(xml.contains("<table:cell-content-change"));
    assert!(xml.contains("table:style-name=\"Historical &amp; Style\""));
    assert!(xml.contains("Zoë &amp; 李"));
    assert!(xml.contains("人民币 &amp; &lt;旧值&gt;"));

    let mut mutable = MutableSpreadsheet::new();
    mutable.add_sheet("Sheet1").unwrap();
    mutable.set_tracked_changes(changes).unwrap();
    let path = std::env::temp_dir().join(format!(
        "litchi-ods-tracked-authoring-{}-{}.ods",
        std::process::id(),
        std::thread::current().name().unwrap_or("test")
    ));
    mutable.save(&path).unwrap();
    let reopened = Spreadsheet::open(&path).unwrap();
    std::fs::remove_file(&path).unwrap();
    let reopened = reopened.tracked_changes().unwrap();
    assert!(reopened.enabled);
    assert_eq!(reopened.changes.len(), 4);
    let SpreadsheetTrackedChange::CellContent(cell) = &reopened.changes[3] else {
        panic!("expected cell-content change");
    };
    assert_eq!(
        cell.previous.style_name.as_deref(),
        Some("Historical & Style")
    );
    assert_eq!(cell.previous.display_text, "人民币 & <旧值>\n第二行");
}

#[test]
fn mutations_are_atomic_when_references_or_cycles_are_invalid() {
    let mut mutable = MutableSpreadsheet::new();
    mutable.set_tracked_changes(complete_change_set()).unwrap();

    let duplicate = SpreadsheetTrackedChange::Insertion(SpreadsheetInsertion {
        metadata: metadata("ct1"),
        dimension: SpreadsheetChangeDimension::Table,
        position: 0,
        count: NonZeroUsize::new(1).unwrap(),
        table: None,
    });
    assert!(mutable.add_tracked_change(duplicate).is_err());
    assert_eq!(mutable.tracked_changes().unwrap().changes.len(), 4);

    assert!(mutable.remove_tracked_change("ct1").is_err());
    assert_eq!(
        mutable.tracked_changes().unwrap().changes[0].metadata().id,
        "ct1"
    );

    let mut first = metadata("a");
    first.dependencies.push("b".to_string());
    let mut second = metadata("b");
    second.dependencies.push("a".to_string());
    let cyclic = SpreadsheetTrackedChanges {
        enabled: true,
        changes: vec![
            SpreadsheetTrackedChange::Insertion(SpreadsheetInsertion {
                metadata: first,
                dimension: SpreadsheetChangeDimension::Row,
                position: 0,
                count: NonZeroUsize::new(1).unwrap(),
                table: Some(0),
            }),
            SpreadsheetTrackedChange::Insertion(SpreadsheetInsertion {
                metadata: second,
                dimension: SpreadsheetChangeDimension::Row,
                position: 1,
                count: NonZeroUsize::new(1).unwrap(),
                table: Some(0),
            }),
        ],
    };
    assert!(mutable.set_tracked_changes(cyclic).is_err());
    assert_eq!(mutable.tracked_changes().unwrap().changes.len(), 4);

    mutable.clear_tracked_changes();
    assert!(mutable.tracked_changes().is_none());
}

#[test]
fn rejects_invalid_positions_values_and_unknown_references() {
    let mut changes = complete_change_set();
    let SpreadsheetTrackedChange::Insertion(insertion) = &mut changes.changes[0] else {
        unreachable!();
    };
    insertion.position = -1;
    assert!(changes.validate().is_err());

    let mut changes = complete_change_set();
    let SpreadsheetTrackedChange::CellContent(cell) = &mut changes.changes[3] else {
        unreachable!();
    };
    cell.previous.value = SpreadsheetTrackedCellValue::Number(f64::NAN);
    assert!(changes.validate().is_err());

    let mut changes = complete_change_set();
    let SpreadsheetTrackedChange::Deletion(deletion) = &mut changes.changes[1] else {
        unreachable!();
    };
    deletion.metadata.dependencies.push("unknown".to_string());
    assert!(changes.validate().is_err());
}
