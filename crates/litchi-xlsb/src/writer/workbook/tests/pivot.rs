//! Workbook-writer serialization and round-trip tests.

use super::super::WorkbookWriter;
use crate::raw::{Writer, kind};
use crate::writer::MutableWorksheet;
use litchi_opc::constants::relationship_type as rel;
use litchi_opc::{OpcPackage, PackURI};
use std::io::Cursor;

#[test]
fn pivot_chart_round_trips_with_lossless_view_and_cache_graph() {
    use crate::package::xlsx::{Chart, ChartAnchor};

    let mut begin_view = vec![0u8; 32];
    begin_view[28..32].copy_from_slice(&1u32.to_le_bytes());
    let view_name = "RevenuePivot";
    begin_view.extend_from_slice(&(view_name.len() as u32).to_le_bytes());
    for unit in view_name.encode_utf16() {
        begin_view.extend_from_slice(&unit.to_le_bytes());
    }
    let mut view_bytes = Vec::new();
    {
        let mut writer = Writer::new(&mut view_bytes);
        writer
            .write_record(kind::BEGIN_SX_VIEW, &begin_view)
            .unwrap();
        writer
            .write_record(kind::BEGIN_SX_LOCATION, &[0; 36])
            .unwrap();
        writer.write_record(kind::END_SX_LOCATION, &[]).unwrap();
        writer.write_record(kind::END_SX_VIEW, &[]).unwrap();
    }
    let view = crate::pivot_view::Part::from_bytes(view_bytes.clone()).unwrap();

    let chart = Chart::line_chart(
        "Revenue",
        "Pivot Host!$A$2:$A$3",
        "Pivot Host!$B$2:$B$3",
        ChartAnchor::new(3, 0, 10, 14),
    )
    .unwrap()
    .into_pivot_chart(view_name)
    .unwrap();
    let mut sheet = MutableWorksheet::new("Pivot Host");
    sheet.add_pivot_table_view(view).unwrap();
    sheet.add_chart(chart).unwrap();

    let mut workbook = WorkbookWriter::new();
    let cache_id = workbook
        .add_pivot_cache(&crate::package::pivot::PivotCacheDefinition::default())
        .unwrap();
    assert_eq!(cache_id, 1);
    workbook.add_worksheet(sheet);

    let mut output = Cursor::new(Vec::new());
    workbook.save(&mut output).unwrap();
    let bytes = output.into_inner();
    let package = OpcPackage::from_bytes(&bytes).unwrap();
    let sheet_part = package
        .get_part(&PackURI::new("/xl/worksheets/sheet1.bin").unwrap())
        .unwrap();
    let view_relationship = sheet_part
        .rels()
        .iter()
        .find(|relationship| relationship.reltype() == rel::PIVOT_TABLE)
        .expect("worksheet PivotTable relationship missing");
    let view_part = package
        .get_part(&view_relationship.target_partname().unwrap())
        .unwrap();
    assert_eq!(view_part.blob(), view_bytes);
    let cache_relationship = view_part
        .rels()
        .iter()
        .find(|relationship| relationship.reltype() == rel::PIVOT_CACHE_DEFINITION)
        .expect("PivotTable cache relationship missing");
    assert_eq!(
        cache_relationship.target_partname().unwrap(),
        PackURI::new("/xl/pivotCache/pivotCacheDefinition1.bin").unwrap()
    );

    let reader = crate::Workbook::new(Cursor::new(bytes)).unwrap();
    assert_eq!(reader.pivot_views().len(), 1);
    assert_eq!(reader.pivot_views()[0].name(), view_name);
    assert_eq!(reader.pivot_views()[0].cache_id(), 1);
    let drawing = reader
        .sheet_drawing(0)
        .expect("pivot chart drawing missing");
    let source = drawing.charts[0]
        .chart
        .pivot_source
        .as_ref()
        .expect("pivot source missing");
    assert_eq!(source.name, "'Pivot Host'!RevenuePivot");
}

#[test]
fn pivot_chart_refuses_a_missing_view_binding() {
    use crate::package::xlsx::{Chart, ChartAnchor};

    let chart = Chart::line_chart(
        "Revenue",
        "Host!$A$1:$A$2",
        "Host!$B$1:$B$2",
        ChartAnchor::new(2, 0, 8, 12),
    )
    .unwrap()
    .into_pivot_chart("MissingPivot")
    .unwrap();
    let mut sheet = MutableWorksheet::new("Host");
    sheet.add_chart(chart).unwrap();
    let mut workbook = WorkbookWriter::new();
    workbook.add_worksheet(sheet);

    let error = workbook
        .save(Cursor::new(Vec::new()))
        .expect_err("missing PivotTable binding must fail");
    assert!(error.to_string().contains("missing pivot table"));
}
