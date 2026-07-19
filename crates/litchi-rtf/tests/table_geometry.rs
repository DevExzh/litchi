use litchi_rtf::{
    RtfDocument, RtfWriter, TableCellMergeRole, TableIndentUnit, TablePreferredWidthUnit,
    TableRowHeight,
};

fn write(document: &RtfDocument<'_>) -> String {
    let mut bytes = Vec::new();
    RtfWriter::new(&mut bytes).write_document(document).unwrap();
    String::from_utf8(bytes).unwrap()
}

#[test]
fn parses_owned_geometry_widths_height_semantics_and_canonical_round_trip() {
    let source = r#"{\rtf1\trowd\trgaph108\trleft-108\trrh240\trftsWidth2\trwWidth2500\trautofit1\tblind-120\tblindtype1\clftsWidth3\clwWidth900\clmgf\cellx900\clftsWidth2\clwWidth2500\clmrg\cellx2100\intbl A\cell B\cell\row\trowd\trrh-300\trftsWidth1\trautofit0\tblind0\tblindtype2\clftsWidth0\clwWidth0\cellx1200\intbl C\cell\row}"#;
    let document = RtfDocument::parse(source).unwrap();
    let rows = document.tables()[0].rows();
    let geometry = rows[0].geometry();
    assert_eq!(geometry.half_gap_twips(), Some(108));
    assert_eq!(geometry.left_edge_twips(), Some(-108));
    assert_eq!(geometry.height(), TableRowHeight::Minimum(240));
    assert_eq!(
        geometry.preferred_width().unwrap().unit(),
        TablePreferredWidthUnit::Percent
    );
    assert_eq!(geometry.preferred_width().unwrap().value(), Some(2500));
    assert!(geometry.auto_fit());
    assert_eq!(geometry.indent().unwrap().unit(), TableIndentUnit::Twips);
    assert_eq!(geometry.indent().unwrap().value(), -120);
    assert_eq!(
        rows[0].cells()[0].preferred_width().unwrap().unit(),
        TablePreferredWidthUnit::Twips
    );
    assert_eq!(
        rows[0].cells()[1].preferred_width().unwrap().unit(),
        TablePreferredWidthUnit::Percent
    );
    assert_eq!(
        rows[0].cells()[0].merge().horizontal,
        Some(TableCellMergeRole::First)
    );
    assert_eq!(rows[0].cells()[1].right_boundary(), Some(2100));
    assert_eq!(rows[1].geometry().height(), TableRowHeight::Exact(300));
    assert_eq!(
        rows[1].geometry().preferred_width().unwrap().unit(),
        TablePreferredWidthUnit::Auto
    );
    assert!(!rows[1].geometry().auto_fit());
    assert_eq!(
        rows[1].cells()[0].preferred_width().unwrap().unit(),
        TablePreferredWidthUnit::Null
    );
    let first = write(&document);
    assert!(first.contains("\\trgaph108\\trleft-108\\trrh240\\trftsWidth2\\trwWidth2500\\trautofit1\\tblind-120\\tblindtype1"));
    assert!(first.contains("\\clftsWidth3\\clwWidth900\\clmgf\\cellx900"));
    assert!(first.contains("\\trrh-300\\trftsWidth1\\tblind0\\tblindtype2"));
    let second = write(&RtfDocument::parse(&first).unwrap());
    assert_eq!(first, second);
}

#[test]
fn restores_groups_resets_trowd_and_snapshots_width_at_each_cellx() {
    let source = r#"{\rtf1\trowd{\trgaph200\trleft300\trrh400\trftsWidth3\trwWidth1000\trautofit1\tblind100\tblindtype1\clftsWidth3\clwWidth400}\cellx1000\intbl A\cell\row\trowd\trgaph50\clftsWidth3\clwWidth600\cellx1000\clftsWidth2\clwWidth1250\cellx2000\intbl B\cell C\cell\row\trowd\cellx900\intbl D\cell\row}"#;
    let document = RtfDocument::parse(source).unwrap();
    let rows = document.tables()[0].rows();
    assert_eq!(rows[0].geometry().half_gap_twips(), None);
    assert_eq!(rows[0].geometry().preferred_width(), None);
    assert_eq!(rows[0].cells()[0].preferred_width(), None);
    assert_eq!(rows[1].geometry().half_gap_twips(), Some(50));
    assert_eq!(
        rows[1].cells()[0].preferred_width().unwrap().value(),
        Some(600)
    );
    assert_eq!(
        rows[1].cells()[1].preferred_width().unwrap().value(),
        Some(1250)
    );
    assert_eq!(rows[2].geometry(), Default::default());
    assert_eq!(rows[2].cells()[0].preferred_width(), None);
}

#[test]
fn applies_geometry_to_end_defined_nested_rows() {
    let source = r#"{\rtf1\trowd\cellx5000\intbl\itap2 A\nestcell\intbl\itap2 B\nestcell{\*\nesttableprops\itap2\trowd\trgaph40\trleft-20\trrh-180\trftsWidth3\trwWidth1700\trautofit1\tblind-10\tblindtype1\clftsWidth3\clwWidth700\cellx700\clftsWidth2\clwWidth2500\cellx1700\nestrow}\intbl\itap1\cell\row}"#;
    let document = RtfDocument::parse(source).unwrap();
    let nested = &document.tables()[0].rows()[0].cells()[0].nested_tables()[0].table;
    let row = &nested.rows()[0];
    assert_eq!(row.geometry().height(), TableRowHeight::Exact(180));
    assert_eq!(
        row.geometry().preferred_width().unwrap().value(),
        Some(1700)
    );
    assert_eq!(row.cells()[0].preferred_width().unwrap().value(), Some(700));
    assert_eq!(
        row.cells()[1].preferred_width().unwrap().unit(),
        TablePreferredWidthUnit::Percent
    );
    let reparsed = RtfDocument::parse(&write(&document)).unwrap();
    assert_eq!(
        reparsed.tables()[0].rows()[0].cells()[0].nested_tables()[0]
            .table
            .rows()[0]
            .geometry(),
        row.geometry()
    );
}

#[test]
fn rejects_missing_invalid_unpaired_and_out_of_range_controls() {
    let controls = [
        "trgaph",
        "trleft",
        "trrh",
        "trautofit",
        "trautofit2",
        "trftsWidth",
        "trftsWidth-1",
        "trftsWidth4",
        "trwWidth1",
        "trftsWidth2",
        "trftsWidth2\\trwWidth5001",
        "trftsWidth3",
        "trftsWidth3\\trwWidth31681",
        "trftsWidth0\\trwWidth1",
        "trftsWidth1\\trwWidth1",
        "tblind31681",
        "tblindtype",
        "tblindtype4",
        "tblind1\\tblindtype0",
        "tblind1\\tblindtype2",
        "tblind5001\\tblindtype3",
        "trrh31681",
        "trrh-31681",
    ];
    for controls in controls {
        let source = format!("{{\\rtf1\\trowd\\{controls}\\cellx1000\\intbl X\\cell\\row}}");
        assert!(RtfDocument::parse(&source).is_err(), "accepted {controls}");
    }
    let cell_controls = [
        "clftsWidth",
        "clftsWidth4",
        "clwWidth1",
        "clftsWidth2",
        "clftsWidth2\\clwWidth5001",
        "clftsWidth3",
        "clftsWidth3\\clwWidth31681",
        "clftsWidth0\\clwWidth1",
        "clftsWidth1\\clwWidth1",
    ];
    for controls in cell_controls {
        let source = format!("{{\\rtf1\\trowd\\{controls}\\cellx1000\\intbl X\\cell\\row}}");
        assert!(RtfDocument::parse(&source).is_err(), "accepted {controls}");
    }
    let defaulted =
        RtfDocument::parse(r#"{\rtf1\trowd\tblind\cellx1000\intbl X\cell\row}"#).unwrap();
    let indent = defaulted.tables()[0].rows()[0].geometry().indent().unwrap();
    assert_eq!(indent.unit(), TableIndentUnit::Twips);
    assert_eq!(indent.value(), 0);
    assert!(RtfDocument::parse(r#"{\rtf1\trowd\trrh0\trgaph31680\trleft-31680\trftsWidth2\trwWidth5000\tblind-5000\tblindtype3\clftsWidth3\clwWidth31680\cellx31680\intbl X\cell\row}"#).is_ok());
}

#[test]
fn parses_real_libreoffice_geometry_fixture() {
    let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../3rdparty/libreoffice-core/sw/qa/extras/rtfimport/data/tdf148544.rtf");
    let document = RtfDocument::parse(&std::fs::read_to_string(path).unwrap()).unwrap();
    let rows = document
        .tables()
        .iter()
        .flat_map(|table| table.rows())
        .collect::<Vec<_>>();
    assert!(
        rows.iter()
            .any(|row| row.geometry().height() == TableRowHeight::Exact(240))
    );
    assert!(
        rows.iter()
            .any(|row| row.geometry().left_edge_twips() == Some(-284)
                && row.geometry().half_gap_twips() == Some(28))
    );
    assert!(
        rows.iter().any(|row| row
            .geometry()
            .preferred_width()
            .is_some_and(|width| width.unit() == TablePreferredWidthUnit::Twips
                && width.value() == Some(3250)))
    );
    assert!(rows.iter().flat_map(|row| row.cells()).any(|cell| {
        cell.preferred_width().is_some_and(|width| {
            width.unit() == TablePreferredWidthUnit::Twips && width.value() == Some(3216)
        })
    }));
    assert!(
        rows.iter()
            .flat_map(|row| row.cells())
            .any(|cell| cell.merge().vertical == Some(TableCellMergeRole::Continuation))
    );
    let reparsed = RtfDocument::parse(&write(&document)).unwrap();
    assert!(
        reparsed
            .tables()
            .iter()
            .flat_map(|table| table.rows())
            .any(|row| row.geometry().height() == TableRowHeight::Exact(240))
    );
}
