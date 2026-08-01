use std::io::Cursor;
use std::path::PathBuf;

use litchi_ograph::chart::Kind as GraphKind;
use litchi_ole::OleWriter;
use litchi_ole::ppt::chart::{Chart, Frame, Kind};
use litchi_ole::ppt::ole_object::{
    PowerPointOleColorFollow, PowerPointOleContainerKind, PowerPointOleDimensionPolicy,
    PowerPointOleDrawAspect, PowerPointOleEmbedPreferences, PowerPointOleExternalObject,
    PowerPointOleObjectDefinition, PowerPointOleObjectMetadata, PowerPointOleObjectSubtype,
    PowerPointOleObjectType,
};
use litchi_ole::ppt::ole_storage::{
    PowerPointOleStorage, PowerPointOleStorageCompression, PowerPointOleStorageKind,
};
use litchi_ole::ppt::{Package, PowerPointOlePackageEditor};

const BOF: u16 = 0x0809;
const EOF: u16 = 0x000A;
const UNKNOWN: u16 = 0x7777;

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

fn graph_bof(doc_type: u16) -> [u8; 16] {
    let mut payload = [0; 16];
    payload[0..2].copy_from_slice(&0x0680_u16.to_le_bytes());
    payload[2..4].copy_from_slice(&doc_type.to_le_bytes());
    payload[4..6].copy_from_slice(&0x0DBB_u16.to_le_bytes());
    payload[6..8].copy_from_slice(&0x07CD_u16.to_le_bytes());
    payload[8..12].copy_from_slice(&(0x0000_0009_u32 | (6 << 14)).to_le_bytes());
    payload[12..16].copy_from_slice(&(0x06_u32 | (6 << 8)).to_le_bytes());
    payload
}

fn excel_bof(doc_type: u16) -> [u8; 16] {
    let mut payload = [0; 16];
    payload[0..2].copy_from_slice(&0x0600_u16.to_le_bytes());
    payload[2..4].copy_from_slice(&doc_type.to_le_bytes());
    payload
}

fn graph_workbook() -> Vec<u8> {
    let mut workbook = record(BOF, &graph_bof(0x0005));
    workbook.extend(record(UNKNOWN, &[1, 2, 3]));
    workbook.extend(record(EOF, &[]));
    workbook.extend(record(BOF, &graph_bof(0x8000)));
    workbook.extend(record(UNKNOWN, &[4, 5]));
    workbook.extend(record(EOF, &[]));
    workbook
}

fn excel_workbook() -> Vec<u8> {
    let mut workbook = record(BOF, &excel_bof(0x0005));
    workbook.extend(record(UNKNOWN, &[1, 2, 3]));
    workbook.extend(record(EOF, &[]));
    workbook.extend(record(BOF, &excel_bof(0x0020)));
    workbook.extend(record(UNKNOWN, &[4, 5]));
    workbook.extend(record(EOF, &[]));
    workbook
}

fn compound(workbook: &[u8]) -> Vec<u8> {
    let mut writer = OleWriter::new();
    writer
        .create_stream(&["Workbook"], workbook)
        .expect("create Workbook stream");
    let mut out = Cursor::new(Vec::new());
    writer.write_to(&mut out).expect("write compound file");
    out.into_inner()
}

fn chart_object(
    id: u32,
    subtype: PowerPointOleObjectSubtype,
    program: &str,
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
            persist_id: 1, // Reassigned by the package editor.
            unused: [0; 4],
        },
        menu_name: None,
        program_id: Some(program.to_string()),
        clipboard_name: None,
        metafile: None,
    })
}

fn uncompressed(data: Vec<u8>) -> PowerPointOleStorage {
    PowerPointOleStorage {
        kind: PowerPointOleStorageKind::OleObject,
        compression: PowerPointOleStorageCompression::Uncompressed,
        data,
    }
}

fn editor_with_seed(seed: u32) -> PowerPointOlePackageEditor {
    let original = std::fs::read(fixture("ppt_with_embeded.ppt")).expect("read fixture");
    let mut package = Package::from_reader(Cursor::new(original.clone())).expect("open fixture");
    let presentation = package.presentation().expect("read presentation");
    let mut objects = presentation
        .ole_objects()
        .expect("parse objects")
        .expect("fixture object list");
    assert!(
        presentation
            .charts()
            .expect("enumerate fixture charts")
            .is_empty(),
        "fixture itself must not contain native chart objects"
    );
    objects.id_seed = seed;
    PowerPointOlePackageEditor::open(original, objects).expect("open package editor")
}

#[test]
fn graph_and_excel_objects_use_neutral_typed_views() {
    let mut editor = editor_with_seed(9);
    editor
        .add(
            chart_object(7, PowerPointOleObjectSubtype::Graph, "MSGraph.Chart.8"),
            uncompressed(compound(&graph_workbook())),
        )
        .expect("add Graph object");
    editor
        .add(
            chart_object(8, PowerPointOleObjectSubtype::ExcelChart, "Excel.Chart.8"),
            uncompressed(compound(&excel_workbook())),
        )
        .expect("add Excel object");
    editor
        .add(
            chart_object(9, PowerPointOleObjectSubtype::Graph, "MSGraph.Chart.8"),
            uncompressed(b"not a compound file".to_vec()),
        )
        .expect("add corrupt object");

    let bytes = editor.finish().expect("finish package edit");
    let mut package = Package::from_reader(Cursor::new(bytes)).expect("reopen package");
    let presentation = package.presentation().expect("read presentation");
    let inventory = presentation.charts().expect("enumerate charts");

    assert_eq!(inventory.seen(), 3);
    assert_eq!(inventory.len(), 2);
    assert_eq!(inventory.charts().count(), 2);
    assert_eq!(inventory.failures().count(), 1);
    let failure = inventory.failures().next().expect("corrupt object failure");
    assert_eq!(failure.kind(), Kind::Graph);
    assert_eq!(failure.info().object_id(), 9);

    let graph = inventory.get(0).expect("Graph chart");
    assert_eq!(graph.kind(), Kind::Graph);
    assert_eq!(graph.info().object_id(), 7);
    assert_eq!(graph.info().program(), Some("MSGraph.Chart.8"));
    assert!(graph.info().frame().is_none());
    let Chart::Graph(graph) = graph else {
        panic!("kind and variant must agree");
    };
    assert_eq!(graph.package().topology().stream_count(), 1);
    assert_eq!(graph.book().len(), 1);
    let chart = graph
        .book()
        .charts()
        .next()
        .expect("one chart")
        .expect("validated chart");
    assert_eq!(chart.kind(), GraphKind::Graph);
    assert_eq!(chart.records().count(), 3);
    assert!(chart.offset() > 0);

    let excel = inventory.get(1).expect("Excel chart");
    assert_eq!(excel.kind(), Kind::Excel);
    assert_eq!(excel.info().object_id(), 8);
    let Chart::Excel(excel) = excel else {
        panic!("kind and variant must agree");
    };
    assert_eq!(excel.book().len(), 1);
    let chart = excel
        .book()
        .charts()
        .next()
        .expect("one chart")
        .expect("validated chart");
    assert_eq!(chart.kind(), GraphKind::Excel);
    assert_eq!(chart.records().count(), 3);
    assert!(chart.offset() > 0);

    assert!(
        inventory
            .at(Frame::new(1, 1).expect("valid frame selector"))
            .is_none()
    );
    assert!(Frame::new(0, 1).is_none());
    assert!(Frame::new(1, 0).is_none());
    assert_eq!(inventory.on_slide(1).count(), 0);
}

#[test]
fn program_identifies_chart_when_subtype_is_default() {
    let mut editor = editor_with_seed(7);
    editor
        .add(
            chart_object(7, PowerPointOleObjectSubtype::Default, "Excel.Chart.8"),
            uncompressed(compound(&excel_workbook())),
        )
        .expect("add chart");

    let bytes = editor.finish().expect("finish package edit");
    let mut package = Package::from_reader(Cursor::new(bytes)).expect("reopen package");
    let presentation = package.presentation().expect("read presentation");
    let inventory = presentation.charts().expect("enumerate charts");
    assert_eq!(inventory.failures().count(), 0);
    assert_eq!(inventory.charts().count(), 1);
    assert_eq!(inventory.get(0).map(Chart::kind), Some(Kind::Excel));
}

#[test]
fn presentation_without_charts_yields_empty_inventory() {
    let path = PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../test-data/ole/ppt/empty.ppt");
    let mut package = Package::open(path).expect("open empty fixture");
    let presentation = package.presentation().expect("read presentation");
    assert!(presentation.charts().expect("enumerate charts").is_empty());
}
