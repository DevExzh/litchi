#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

use std::io::Cursor;
use std::path::PathBuf;

use litchi_cfb::OleWriter;
use litchi_core::edit::RefusalKind;
use litchi_ograph::chart::Kind as GraphKind;
use litchi_ppt::chart::{Chart, Frame, Kind};
use litchi_ppt::embedded::object::{
    self, ColorFollow, ContainerKind, Definition, DimensionPolicy, DrawAspect, EmbedPreferences,
    ExternalObject, Metadata, ObjectSubtype, ObjectType,
};
use litchi_ppt::embedded::storage::{Kind as StorageKind, Storage};
use litchi_ppt::{Package, raw::Reader};

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
    out.extend(u16::try_from(data.len()).unwrap().to_le_bytes());
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
    workbook.extend(graph_chart(&[4, 5]));
    workbook
}

fn graph_chart(data: &[u8]) -> Vec<u8> {
    let mut chart = record(BOF, &graph_bof(0x8000));
    chart.extend(record(UNKNOWN, data));
    chart.extend(record(EOF, &[]));
    chart
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

fn chart_object(id: u32, subtype: ObjectSubtype, program: &str) -> ExternalObject {
    ExternalObject::Object(Definition {
        kind: ContainerKind::Embedded(EmbedPreferences {
            color_follow: ColorFollow::EntireScheme,
            cannot_lock_server: false,
            dimension_policy: DimensionPolicy::Send,
            is_word_table: false,
            unused: 0,
        }),
        object: Metadata {
            draw_aspect: DrawAspect::Content,
            object_type: ObjectType::Embedded,
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

fn uncompressed(data: Vec<u8>) -> Storage {
    Storage::uncompressed(StorageKind::OleObject, data).unwrap()
}

/// A package-level object edit over a fixture whose six objects carry ids
/// 1..=6; adding ids above the seed raises it.
fn objects_edit() -> object::Edit {
    let original = std::fs::read(fixture("ppt_with_embeded.ppt")).expect("read fixture");
    let package = Package::from_bytes(&original).expect("open fixture");
    let presentation = package.presentation().expect("read presentation");
    assert!(
        presentation.ole_objects().expect("parse objects").is_some(),
        "fixture must carry an external-object list"
    );
    assert!(
        presentation
            .charts()
            .expect("enumerate fixture charts")
            .is_empty(),
        "fixture itself must not contain native chart objects"
    );
    package
        .edit_objects()
        .expect("open package-level object edit")
}

fn committed_bytes(edit: object::Edit) -> Vec<u8> {
    let commit = edit.commit().expect("commit object edit");
    assert!(commit.diagnostics().changed());
    commit.package().to_bytes().expect("committed bytes")
}

#[test]
fn graph_and_excel_objects_use_neutral_typed_views() {
    let mut editor = objects_edit();
    editor
        .add(
            chart_object(7, ObjectSubtype::Graph, "MSGraph.Chart.8"),
            uncompressed(compound(&graph_workbook())),
        )
        .expect("add Graph object");
    editor
        .add(
            chart_object(8, ObjectSubtype::ExcelChart, "Excel.Chart.8"),
            uncompressed(compound(&excel_workbook())),
        )
        .expect("add Excel object");
    editor
        .add(
            chart_object(9, ObjectSubtype::Graph, "MSGraph.Chart.8"),
            uncompressed(b"not a compound file".to_vec()),
        )
        .expect("add corrupt object");

    let bytes = committed_bytes(editor);
    let mut package = Reader::from_reader(Cursor::new(bytes)).expect("reopen package");
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
    let Chart::Graph(typed_graph) = graph else {
        panic!("kind and variant must agree");
    };
    assert_eq!(typed_graph.package().topology().stream_count(), 1);
    assert_eq!(typed_graph.book().len(), 1);
    let graph_chart = typed_graph
        .book()
        .charts()
        .expect("validated chart streams")
        .next()
        .expect("one chart")
        .expect("validated chart");
    assert_eq!(graph_chart.kind(), GraphKind::Graph);
    assert_eq!(graph_chart.records().expect("validated records").count(), 3);
    assert!(graph_chart.offset() > 0);

    let excel = inventory.get(1).expect("Excel chart");
    assert_eq!(excel.kind(), Kind::Excel);
    assert_eq!(excel.info().object_id(), 8);
    let Chart::Excel(typed_excel) = excel else {
        panic!("kind and variant must agree");
    };
    assert_eq!(typed_excel.book().len(), 1);
    let excel_chart = typed_excel
        .book()
        .charts()
        .expect("validated chart streams")
        .next()
        .expect("one chart")
        .expect("validated chart");
    assert_eq!(excel_chart.kind(), GraphKind::Excel);
    assert_eq!(excel_chart.records().expect("validated records").count(), 3);
    assert!(excel_chart.offset() > 0);

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
    let mut editor = objects_edit();
    editor
        .add(
            chart_object(7, ObjectSubtype::Default, "Excel.Chart.8"),
            uncompressed(compound(&excel_workbook())),
        )
        .expect("add chart");

    let bytes = committed_bytes(editor);
    let mut package = Reader::from_reader(Cursor::new(bytes)).expect("reopen package");
    let presentation = package.presentation().expect("read presentation");
    let inventory = presentation.charts().expect("enumerate charts");
    assert_eq!(inventory.failures().count(), 0);
    assert_eq!(inventory.charts().count(), 1);
    assert_eq!(inventory.get(0).map(Chart::kind), Some(Kind::Excel));
}

#[test]
fn graph_package_replacement_publishes_through_the_object_edit() {
    let mut seed = objects_edit();
    seed.add(
        chart_object(7, ObjectSubtype::Graph, "MSGraph.Chart.8"),
        uncompressed(compound(&graph_workbook())),
    )
    .expect("add Graph object");
    let original = Package::from_bytes(&committed_bytes(seed)).expect("seeded package");
    let original_bytes = original.to_bytes().expect("seeded bytes");

    let presentation = original.presentation().expect("read presentation");
    let inventory = presentation.charts().expect("enumerate charts");
    let graph = match inventory.get(0).expect("Graph chart") {
        Chart::Graph(graph) => graph,
        Chart::Excel(_) => panic!("expected Graph chart"),
    };
    let persist_id = graph.info().persist_id();
    let original_storage = presentation
        .ole_storage(persist_id)
        .expect("read storage")
        .expect("Graph storage");

    // An untouched draft is the exact source package and stages no change.
    let untouched = graph.draft_package().expect("open chart package draft");
    assert!(!untouched.view().expect("draft view").is_dirty());
    let mut edit = original.edit_objects().expect("open object edit");
    graph
        .replace_package(&mut edit, untouched)
        .expect("stage untouched draft");
    let no_op = edit.commit().expect("commit untouched draft");
    assert!(!no_op.diagnostics().changed());
    assert!(no_op.patch().is_noop());
    assert_eq!(no_op.package().to_bytes().unwrap(), original_bytes);

    let replacement = graph_chart(&[9, 8, 7]);
    let mut chart_package = graph.draft_package().expect("open chart package draft");
    chart_package
        .replace_chart(
            litchi_ograph::chart::Stream::open(
                replacement.clone(),
                litchi_ograph::Limits::default(),
            )
            .expect("typed Graph chart stream"),
        )
        .expect("stage chart replacement");
    assert_eq!(
        chart_package.view().expect("draft view").chart().as_bytes(),
        replacement.as_slice()
    );
    let mut edit = original.edit_objects().expect("open object edit");
    graph
        .replace_package(&mut edit, chart_package)
        .expect("stage typed host replacement");
    let commit = edit.commit().expect("commit chart replacement");
    assert!(commit.diagnostics().changed());
    let rewritten = commit.package().to_bytes().expect("rewritten bytes");

    let mut reopened = Reader::from_reader(Cursor::new(rewritten.clone())).expect("reopen package");
    let rewritten_presentation = reopened
        .presentation()
        .expect("read rewritten presentation");
    let rewritten_inventory = rewritten_presentation
        .charts()
        .expect("enumerate rewritten charts");
    let Chart::Graph(rewritten_graph) = rewritten_inventory.get(0).expect("rewritten Graph chart")
    else {
        panic!("expected rewritten Graph chart");
    };
    let workbook = rewritten_graph
        .package()
        .workbook()
        .expect("rewritten Workbook");
    assert_eq!(workbook.chart().as_bytes(), replacement.as_slice());
    assert_eq!(rewritten_graph.info().object_id(), 7);
    assert_eq!(rewritten_graph.info().persist_id(), persist_id);
    // The object list is untouched: only the storage record moved.
    assert_eq!(
        rewritten_presentation.ole_objects().unwrap(),
        presentation.ole_objects().unwrap()
    );
    assert_ne!(
        rewritten_presentation.ole_storage(persist_id).unwrap(),
        Some(original_storage)
    );

    // The patch replays exactly onto its source and its inverse restores it.
    let replayed = original
        .apply_objects(commit.patch())
        .expect("replay object patch");
    assert_eq!(replayed.package().to_bytes().unwrap(), rewritten);
    let restored = commit
        .package()
        .apply_objects(&commit.patch().inverse())
        .expect("invert object patch");
    assert_eq!(restored.package().to_bytes().unwrap(), original_bytes);
    assert!(matches!(
        commit.package().apply_objects(commit.patch()),
        Err(object::Error::PatchConflict)
    ));
}

#[test]
fn graph_package_replacement_refuses_a_missing_host_object() {
    let mut seed = objects_edit();
    seed.add(
        chart_object(7, ObjectSubtype::Graph, "MSGraph.Chart.8"),
        uncompressed(compound(&graph_workbook())),
    )
    .expect("add Graph object");
    let original = Package::from_bytes(&committed_bytes(seed)).expect("seeded package");
    let presentation = original.presentation().expect("read presentation");
    let inventory = presentation.charts().expect("enumerate charts");
    let Chart::Graph(graph) = inventory.get(0).expect("Graph chart") else {
        panic!("expected Graph chart");
    };

    let mut edit = original.edit_objects().expect("open object edit");
    edit.remove(7).expect("stage chart object removal");
    let error = graph
        .replace_package(&mut edit, graph.draft_package().unwrap())
        .expect_err("removed chart object is absent");
    assert!(matches!(
        error,
        object::Error::Refused(ref refusal) if refusal.kind() == RefusalKind::NotFound
    ));
}

#[test]
fn presentation_without_charts_yields_empty_inventory() {
    let path = PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../test-data/ole/ppt/empty.ppt");
    let mut package = Reader::open(path).expect("open empty fixture");
    let presentation = package.presentation().expect("read presentation");
    assert!(presentation.charts().expect("enumerate charts").is_empty());
}
