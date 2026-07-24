use litchi_rtf::{RtfDocument, RtfWriter, TableCellMergeRole};

fn write(document: &RtfDocument<'_>) -> String {
    let mut bytes = Vec::new();
    RtfWriter::new(&mut bytes).write_document(document).unwrap();
    String::from_utf8(bytes).unwrap()
}

#[test]
fn parses_owned_merge_roles_boundaries_and_round_trips_canonically() {
    let source = r#"{\rtf1\trowd\clmgf\clvmgf\cellx900\clmrg\cellx2100\intbl A\cell B\cell\row\trowd\clmgf\clvmrg\cellx900\clmrg\cellx2100\intbl C\cell D\cell\row}"#;
    let document = RtfDocument::parse(source).unwrap();
    let rows = document.tables()[0].rows();
    assert_eq!(
        rows[0].cells()[0].merge().horizontal,
        Some(TableCellMergeRole::First)
    );
    assert_eq!(
        rows[0].cells()[0].merge().vertical,
        Some(TableCellMergeRole::First)
    );
    assert_eq!(
        rows[1].cells()[0].merge().vertical,
        Some(TableCellMergeRole::Continuation)
    );
    assert_eq!(rows[0].cells()[0].right_boundary(), Some(900));
    assert_eq!(rows[0].cells()[1].right_boundary(), Some(2100));
    let first = write(&document);
    assert!(first.contains("\\clmgf\\clvmgf\\cellx900"));
    assert!(first.contains("\\clmgf\\clvmrg\\cellx900"));
    let second = write(&RtfDocument::parse(&first).unwrap());
    assert_eq!(first, second);
}

#[test]
fn resets_trowd_restores_groups_and_snapshots_each_cellx() {
    let source = r#"{\rtf1\trowd{\clmgf}\cellx1000\cellx2000\intbl A\cell B\cell\row\trowd\clmgf\cellx1100\clmrg\cellx2200\intbl C\cell D\cell\row\trowd\cellx1200\intbl E\cell\row}"#;
    let document = RtfDocument::parse(source).unwrap();
    let rows = document.tables()[0].rows();
    assert_eq!(rows[0].cells()[0].merge().horizontal, None);
    assert_eq!(
        rows[1].cells()[0].merge().horizontal,
        Some(TableCellMergeRole::First)
    );
    assert_eq!(
        rows[1].cells()[1].merge().horizontal,
        Some(TableCellMergeRole::Continuation)
    );
    assert_eq!(rows[2].cells()[0].merge().horizontal, None);
}

#[test]
fn rejects_parameters_duplicates_conflicts_and_invalid_horizontal_ordering() {
    for word in ["clmgf0", "clmrg1", "clvmgf-1", "clvmrg2"] {
        let source = format!("{{\\rtf1\\trowd\\{word}\\cellx1000\\intbl X\\cell\\row}}");
        assert!(RtfDocument::parse(&source).is_err(), "accepted {word}");
    }
    for controls in [
        "clmgf\\clmgf",
        "clmrg\\clmrg",
        "clmgf\\clmrg",
        "clmrg\\clmgf",
        "clvmgf\\clvmgf",
        "clvmrg\\clvmrg",
        "clvmgf\\clvmrg",
        "clvmrg\\clvmgf",
    ] {
        let source = format!("{{\\rtf1\\trowd\\{controls}\\cellx1000\\intbl X\\cell\\row}}");
        assert!(RtfDocument::parse(&source).is_err(), "accepted {controls}");
    }
    for source in [
        r#"{\rtf1\trowd\clmrg\cellx1000\intbl X\cell\row}"#,
        r#"{\rtf1\trowd\clmgf\cellx1000\cellx2000\clmrg\cellx3000\intbl X\cell Y\cell Z\cell\row}"#,
    ] {
        assert!(RtfDocument::parse(source).is_err(), "accepted {source}");
    }
    assert!(RtfDocument::parse(r#"{\rtf1\trowd\clmgf\clvmgf\cellx1000\intbl X\cell\row}"#).is_ok());
}

#[test]
fn validates_vertical_boundaries_and_logical_table_boundaries() {
    let valid = r#"{\rtf1\trowd\clvmgf\cellx1000\cellx2000\intbl A\cell B\cell\row\trowd\clvmrg\cellx1000\cellx2000\intbl C\cell D\cell\row}"#;
    assert!(RtfDocument::parse(valid).is_ok());
    let mismatch = r#"{\rtf1\trowd\clvmgf\cellx1000\cellx2000\intbl A\cell B\cell\row\trowd\clvmrg\cellx1100\cellx2000\intbl C\cell D\cell\row}"#;
    assert!(RtfDocument::parse(mismatch).is_err());
    let split = r#"{\rtf1\trowd\clvmgf\cellx1000\intbl A\cell\row\trowd\clvmrg\cellx1000\intbl B\cell\row\pard\par\trowd\clvmrg\cellx1000\intbl C\cell\row}"#;
    assert!(RtfDocument::parse(split).is_err());
}

#[test]
fn applies_end_defined_nested_merge_metadata() {
    let source = r#"{\rtf1\trowd\cellx5000\intbl\itap2 A\nestcell\intbl\itap2 B\nestcell{\*\nesttableprops\itap2\trowd\clmgf\cellx700\clmrg\cellx1700\nestrow}\intbl\itap1\cell\row}"#;
    let document = RtfDocument::parse(source).unwrap();
    let cells = document.tables()[0].rows()[0].cells()[0].nested_tables()[0]
        .table
        .rows()[0]
        .cells();
    assert_eq!(cells[0].merge().horizontal, Some(TableCellMergeRole::First));
    assert_eq!(
        cells[1].merge().horizontal,
        Some(TableCellMergeRole::Continuation)
    );
    assert_eq!(cells[1].right_boundary(), Some(1700));
    assert_eq!(
        RtfDocument::parse(&write(&document)).unwrap().tables()[0].rows()[0].cells()[0]
            .nested_tables()[0]
            .table
            .rows()[0]
            .cells()[1]
            .merge()
            .horizontal,
        Some(TableCellMergeRole::Continuation)
    );
}

#[test]
fn parses_real_libreoffice_merge_fixtures() {
    let base = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/libreoffice-core/sw/qa/extras");
    let horizontal = RtfDocument::parse(
        &std::fs::read_to_string(base.join("rtfimport/data/tdf117403.rtf")).unwrap(),
    )
    .unwrap();
    let cells = horizontal.tables()[0].rows()[0].cells();
    assert_eq!(cells[0].merge().horizontal, Some(TableCellMergeRole::First));
    assert_eq!(
        cells[1].merge().horizontal,
        Some(TableCellMergeRole::Continuation)
    );
    let vertical = RtfDocument::parse(
        &std::fs::read_to_string(base.join("rtfimport/data/tdf148544.rtf")).unwrap(),
    )
    .unwrap();
    assert!(
        vertical
            .tables()
            .iter()
            .flat_map(|table| table.rows())
            .flat_map(|row| row.cells())
            .any(|cell| cell.merge().vertical == Some(TableCellMergeRole::Continuation))
    );
}

#[test]
fn merge_definitions_still_obey_the_cell_cap() {
    let mut source = String::from("{\\rtf1\\trowd");
    for index in 0..=4096 {
        source.push_str("\\clmgf\\cellx");
        source.push_str(&(index + 1).to_string());
    }
    source.push('}');
    assert!(RtfDocument::parse(&source).is_err());
}
