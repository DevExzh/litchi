use super::codec::decode;
use super::model::{Chart, Frame, Graph, Info, Kind};
use super::package::{classify, collect_frames, program_base};
use crate::embedded::object::ObjectSubtype;
use crate::embedded::storage::{Compression, Kind as StorageKind, Storage};
use crate::shapes::{PictureFrameKind, PictureShape, ShapeEnum};
use litchi_cfb::OleWriter;
use litchi_ograph::Limits;
use litchi_ograph::chart::Book;
use std::io::Cursor;
use std::io::Write;

#[test]
fn subtype_and_program_identify_chart_objects() {
    assert_eq!(classify(ObjectSubtype::Graph, None), Some(Kind::Graph));
    assert_eq!(classify(ObjectSubtype::ExcelChart, None), Some(Kind::Excel));
    for (program, kind) in [
        ("MSGraph.Chart.8", Kind::Graph),
        ("MSGraph.Chart", Kind::Graph),
        ("MSGraph", Kind::Graph),
        ("msgraph.chart.8", Kind::Graph),
        ("Excel.Chart.8", Kind::Excel),
        ("Excel.Chart", Kind::Excel),
        ("EXCEL.CHART.8", Kind::Excel),
    ] {
        assert_eq!(
            classify(ObjectSubtype::Default, Some(program)),
            Some(kind),
            "{program}"
        );
    }
    for program in [
        "Excel.Sheet.8",
        "Excel.SheetMacroEnabled.12",
        "Word.Document.8",
        "PowerPoint.Show.8",
        "Equation.3",
        "Excel.ChartTool.8",
    ] {
        assert_eq!(
            classify(ObjectSubtype::Default, Some(program)),
            None,
            "{program}"
        );
    }
}

#[test]
fn program_base_strips_only_numeric_versions() {
    assert_eq!(program_base("Excel.Chart.8"), "Excel.Chart");
    assert_eq!(program_base("Excel.Chart"), "Excel.Chart");
    assert_eq!(program_base("MSGraph"), "MSGraph");
    assert_eq!(program_base("Excel.Chart."), "Excel.Chart.");
}

fn storage(compression: Compression, declared: u32, data: Vec<u8>) -> Storage {
    match compression {
        Compression::Uncompressed => Storage::uncompressed(StorageKind::OleObject, data),
        Compression::Zlib => Storage::compressed(StorageKind::OleObject, declared, data),
    }
    .unwrap()
}

#[test]
fn payload_decoding_is_move_first_and_bounded() {
    let bytes = b"compound".to_vec();
    let pointer = bytes.as_ptr();
    let decoded = decode(
        storage(Compression::Uncompressed, 0, bytes),
        Limits::default(),
    )
    .expect("uncompressed payload");
    assert_eq!(decoded.as_ptr(), pointer);

    let raw = vec![7u8; 65_536];
    let mut encoder = flate2::write::ZlibEncoder::new(Vec::new(), flate2::Compression::default());
    encoder.write_all(&raw).expect("compress");
    let compressed = encoder.finish().expect("finish compression");
    let decoded = decode(
        storage(Compression::Zlib, raw.len() as u32, compressed),
        Limits::default(),
    )
    .expect("compressed payload");
    assert_eq!(decoded, raw);

    let limits = Limits {
        max_package_bytes: 1024,
        ..Limits::default()
    };
    let bomb = storage(Compression::Zlib, 4096, vec![0x78, 0x9c]);
    assert!(decode(bomb, limits).is_err());
}

#[test]
fn ole_frames_are_semantically_attributed_and_first_wins() {
    let mut chart = PictureShape::new(7);
    chart.set_frame_kind(PictureFrameKind::OleObject);
    chart.set_external_object_id(42);
    let mut nested = PictureShape::new(9);
    nested.set_frame_kind(PictureFrameKind::OleObject);
    nested.set_external_object_id(77);
    let mut group = crate::shapes::shape_enum::GroupShape::new(10);
    group.add_child(ShapeEnum::Picture(nested));
    let mut duplicate = PictureShape::new(11);
    duplicate.set_frame_kind(PictureFrameKind::OleObject);
    duplicate.set_external_object_id(42);
    let shapes = vec![
        ShapeEnum::Picture(chart),
        ShapeEnum::Group(group),
        ShapeEnum::Picture(duplicate),
    ];
    let mut frames = Vec::new();
    collect_frames(&shapes, 3, &[42, 77], &mut frames);
    assert_eq!(
        frames,
        vec![
            (42, Frame::new(3, 7).expect("valid frame")),
            (77, Frame::new(3, 9).expect("valid frame"))
        ]
    );
}

fn graph_workbook(chart: &[u8]) -> Vec<u8> {
    let mut workbook = Vec::new();
    let mut globals = [0u8; 16];
    globals[0..2].copy_from_slice(&0x0680_u16.to_le_bytes());
    globals[2..4].copy_from_slice(&0x0005_u16.to_le_bytes());
    globals[4..6].copy_from_slice(&0x0DBB_u16.to_le_bytes());
    globals[6..8].copy_from_slice(&0x07CD_u16.to_le_bytes());
    globals[8..12].copy_from_slice(&(0x0000_0009_u32 | (6 << 14)).to_le_bytes());
    globals[12..16].copy_from_slice(&(0x06_u32 | (6 << 8)).to_le_bytes());
    push_record(&mut workbook, 0x0809, &globals);
    push_record(&mut workbook, 0x7777, &[1]);
    push_record(&mut workbook, 0x000A, &[]);
    workbook.extend_from_slice(chart);
    workbook
}

fn push_record(output: &mut Vec<u8>, kind: u16, payload: &[u8]) {
    output.extend_from_slice(&kind.to_le_bytes());
    output.extend_from_slice(
        &(u16::try_from(payload.len()).expect("small test record")).to_le_bytes(),
    );
    output.extend_from_slice(payload);
}

#[test]
fn semantic_validation_is_opt_in_without_changing_raw_inventory() {
    let mut chart_bytes = Vec::new();
    let mut chart_bof = [0u8; 16];
    chart_bof[0..2].copy_from_slice(&0x0680_u16.to_le_bytes());
    chart_bof[2..4].copy_from_slice(&0x8000_u16.to_le_bytes());
    chart_bof[4..6].copy_from_slice(&0x0DBB_u16.to_le_bytes());
    chart_bof[6..8].copy_from_slice(&0x07CD_u16.to_le_bytes());
    chart_bof[8..12].copy_from_slice(&(0x0000_0009_u32 | (6 << 14)).to_le_bytes());
    chart_bof[12..16].copy_from_slice(&(0x06_u32 | (6 << 8)).to_le_bytes());
    push_record(&mut chart_bytes, 0x0809, &chart_bof);
    push_record(&mut chart_bytes, 0x7777, &[2, 3]);
    push_record(&mut chart_bytes, 0x000A, &[]);

    let workbook = graph_workbook(&chart_bytes);
    let mut writer = OleWriter::new();
    writer
        .create_stream(&["Workbook"], &workbook)
        .expect("Workbook stream");
    let mut package_bytes = Cursor::new(Vec::new());
    writer.write_to(&mut package_bytes).expect("Graph package");

    let package = litchi_ograph::Package::open(package_bytes.into_inner()).expect("Graph package");
    let book = Book::open(workbook).expect("framed Graph Workbook");
    let chart = Chart::Graph(Graph::new(
        Info::new(1, 1, Some("MSGraph.Chart.8".into()), None),
        Box::new(package),
        book,
        Compression::Uncompressed,
    ));

    assert!(chart.charts().next().expect("raw chart").is_ok());
    let semantic = chart
        .semantic_chart(0)
        .expect("semantic chart validation")
        .expect("one semantic chart");
    assert!(semantic.is_pristine());
    assert_eq!(semantic.unknown().len(), 1);
    assert_eq!(
        semantic.encode().expect("lossless replay").as_bytes(),
        chart_bytes
    );
}

#[test]
fn replacement_storage_preserves_compression_mode() {
    let original = b"original chart package";
    let mut encoder = flate2::write::ZlibEncoder::new(Vec::new(), flate2::Compression::default());
    encoder.write_all(original).expect("compress source");
    let source = storage(
        Compression::Zlib,
        original.len() as u32,
        encoder.finish().expect("finish source compression"),
    );
    let replacement = b"replacement chart package with a new size".to_vec();
    let encoded = super::codec::encode_storage(replacement.clone(), source.compression())
        .expect("encode replacement");

    assert_eq!(encoded.compression(), Compression::Zlib);
    assert_eq!(
        encoded.declared_uncompressed_len(),
        Some(replacement.len() as u32)
    );
    assert_eq!(
        decode(encoded, Limits::default()).expect("decode replacement"),
        replacement
    );
}
