use litchi_ole::OleWriter;
use litchi_ole::ppt::ole_object::{
    PowerPointOleColorFollow, PowerPointOleContainerKind, PowerPointOleDimensionPolicy,
    PowerPointOleDrawAspect, PowerPointOleEmbedPreferences, PowerPointOleExternalObject,
    PowerPointOleObjectDefinition, PowerPointOleObjectMetadata, PowerPointOleObjectSubtype,
    PowerPointOleObjectType,
};
use litchi_ole::ppt::ole_storage::{
    PowerPointOleStorage, PowerPointOleStorageCompression, PowerPointOleStorageKind,
};
use litchi_ole::ppt::{Package, PowerPointChartKind, PowerPointOlePackageEditor};
use litchi_ole::xls::{
    XlsChart, XlsChartCacheEntry, XlsChartCachedValue, XlsChartCellReference, XlsChartDataLink,
    XlsChartEditor, XlsChartGroupKind, XlsChartLimits, XlsChartLocation, XlsChartSeries,
    XlsChartType,
};
use std::io::Cursor;
use std::path::PathBuf;

fn fixture(name: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/poi/test-data/slideshow")
        .join(name)
}

fn record(kind: u16, data: &[u8]) -> Vec<u8> {
    let mut out = kind.to_le_bytes().to_vec();
    out.extend((data.len() as u16).to_le_bytes());
    out.extend(data);
    out
}

fn bof(kind: u16) -> Vec<u8> {
    let mut data = Vec::new();
    data.extend(0x0600u16.to_le_bytes());
    data.extend(kind.to_le_bytes());
    data.extend([0; 12]);
    record(0x0809, &data)
}

/// A minimal one-worksheet compound file accepted by `XlsChartEditor`.
fn workbook() -> Vec<u8> {
    let mut globals = bof(5);
    let bound_at = globals.len();
    let mut bound = vec![0; 8];
    bound[6] = 1;
    bound.extend(b"S");
    globals.extend(record(0x0085, &bound));
    globals.extend(record(0x000a, &[]));
    let offset = globals.len() as u32;
    globals[bound_at + 4..bound_at + 8].copy_from_slice(&offset.to_le_bytes());
    globals.extend(bof(0x0010));
    globals.extend(record(0x000a, &[]));
    let mut writer = OleWriter::new();
    writer.create_stream(&["Workbook"], &globals).unwrap();
    let mut out = Cursor::new(Vec::new());
    writer.write_to(&mut out).unwrap();
    out.into_inner()
}

/// A chart-bearing compound file: one line-chart series with a BIFF data
/// link, a cached value, and a title.
fn chart_workbook() -> Vec<u8> {
    let mut chart = XlsChart {
        title: Some("Sales".to_string()),
        ..Default::default()
    };
    chart.series.push(XlsChartSeries {
        category_count: 4,
        value_count: 4,
        links: vec![XlsChartDataLink {
            role: 1,
            source_type: 2,
            unlinked_number_format: false,
            number_format: 0,
            formula_tokens: vec![0x3b, 0, 0, 0, 0, 3, 0, 0, 0, 0, 0],
            references: vec![XlsChartCellReference {
                extern_sheet_index: 0,
                first_row: 0,
                last_row: 3,
                first_column: 0,
                last_column: 0,
            }],
        }],
        ..Default::default()
    });
    chart.cached_values.push(XlsChartCacheEntry {
        cache_index: 0,
        row: 0,
        column: 0,
        value: XlsChartCachedValue::Number(42.0),
    });
    let mut editor = XlsChartEditor::open(workbook(), XlsChartLimits::default()).unwrap();
    editor.add(0, 7, 0, chart).unwrap();
    editor.finish().unwrap()
}

fn chart_object(
    id: u32,
    subtype: PowerPointOleObjectSubtype,
    prog_id: &str,
) -> PowerPointOleExternalObject {
    PowerPointOleExternalObject::Object(PowerPointOleObjectDefinition {
        kind: PowerPointOleContainerKind::Embedded(PowerPointOleEmbedPreferences {
            color_follow: PowerPointOleColorFollow::EntireScheme,
            cannot_lock_server: false,
            dimension_policy: PowerPointOleDimensionPolicy::Send,
            is_word_table: false,
            unused: 0,
        }),
        object: PowerPointOleObjectMetadata {
            draw_aspect: PowerPointOleDrawAspect::Content,
            object_type: PowerPointOleObjectType::Embedded,
            id,
            subtype,
            persist_id: 1, // reassigned by the package editor
            unused: [0; 4],
        },
        menu_name: None,
        program_id: Some(prog_id.to_string()),
        clipboard_name: None,
        metafile: None,
    })
}

fn uncompressed_storage(data: Vec<u8>) -> PowerPointOleStorage {
    PowerPointOleStorage {
        kind: PowerPointOleStorageKind::OleObject,
        compression: PowerPointOleStorageCompression::Uncompressed,
        data,
    }
}

#[test]
fn embedded_native_chart_is_parsed_and_corrupt_chart_degrades() {
    let bytes = std::fs::read(fixture("ppt_with_embeded.ppt")).unwrap();
    let mut package = Package::from_reader(Cursor::new(bytes)).unwrap();
    let presentation = package.presentation().unwrap();
    let mut collection = presentation
        .ole_objects()
        .unwrap()
        .expect("fixture has an external-object list");
    // The unmodified fixture embeds Word tables and Excel sheets, not charts.
    assert!(presentation.charts().unwrap().is_empty());
    drop(presentation);

    collection.id_seed = 8;
    let original = std::fs::read(fixture("ppt_with_embeded.ppt")).unwrap();
    let mut editor = PowerPointOlePackageEditor::open(original, collection).unwrap();
    editor
        .add(
            chart_object(7, PowerPointOleObjectSubtype::ExcelChart, "Excel.Chart.8"),
            uncompressed_storage(chart_workbook()),
        )
        .unwrap();
    editor
        .add(
            chart_object(8, PowerPointOleObjectSubtype::Graph, "MSGraph.Chart.8"),
            uncompressed_storage(b"not a compound file".to_vec()),
        )
        .unwrap();
    let bytes = editor.finish().unwrap();

    let mut package = Package::from_reader(Cursor::new(bytes)).unwrap();
    let presentation = package.presentation().unwrap();
    let inventory = presentation.charts().unwrap();

    assert_eq!(
        inventory.charts.len(),
        1,
        "failures: {:?}",
        inventory.failures
    );
    assert_eq!(inventory.failures.len(), 1);
    assert_eq!(inventory.failures[0].object_id, 8);

    let chart = &inventory.charts[0];
    assert_eq!(chart.object_id, 7);
    assert_eq!(chart.kind, PowerPointChartKind::ExcelChart);
    assert_eq!(chart.program_id.as_deref(), Some("Excel.Chart.8"));
    assert_eq!(chart.charts.len(), 1);
    let entry = &chart.charts[0];
    assert_eq!(
        entry.location,
        XlsChartLocation::Embedded {
            sheet_index: 0,
            object_id: 7
        }
    );
    assert!(matches!(
        entry.chart.chart_type(),
        XlsChartType::Single(XlsChartGroupKind::Line { .. })
    ));
    assert_eq!(entry.chart.title.as_deref(), Some("Sales"));
    assert_eq!(entry.chart.series.len(), 1);
    assert_eq!(entry.chart.series[0].links.len(), 1);
    assert_eq!(entry.chart.series[0].links[0].references.len(), 1);
    assert!(
        entry
            .chart
            .cached_values
            .iter()
            .any(|value| value.value == XlsChartCachedValue::Number(42.0))
    );
}

#[test]
fn prog_id_identifies_chart_when_subtype_is_default() {
    let bytes = std::fs::read(fixture("ppt_with_embeded.ppt")).unwrap();
    let mut package = Package::from_reader(Cursor::new(bytes)).unwrap();
    let presentation = package.presentation().unwrap();
    let mut collection = presentation.ole_objects().unwrap().unwrap();
    drop(presentation);
    collection.id_seed = 7;
    let mut editor = PowerPointOlePackageEditor::open(
        std::fs::read(fixture("ppt_with_embeded.ppt")).unwrap(),
        collection,
    )
    .unwrap();
    editor
        .add(
            chart_object(7, PowerPointOleObjectSubtype::Default, "Excel.Chart.8"),
            uncompressed_storage(chart_workbook()),
        )
        .unwrap();
    let bytes = editor.finish().unwrap();

    let mut package = Package::from_reader(Cursor::new(bytes)).unwrap();
    let presentation = package.presentation().unwrap();
    let inventory = presentation.charts().unwrap();
    assert_eq!(inventory.failures.len(), 0);
    assert_eq!(inventory.charts.len(), 1);
    assert_eq!(inventory.charts[0].kind, PowerPointChartKind::ExcelChart);
}

#[test]
fn presentation_without_charts_yields_empty_inventory() {
    let path = PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../test-data/ole/ppt/empty.ppt");
    let mut package = Package::open(path).unwrap();
    let presentation = package.presentation().unwrap();
    assert!(presentation.charts().unwrap().is_empty());
}
