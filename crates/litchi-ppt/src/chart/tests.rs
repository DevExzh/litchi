use super::codec::decode;
use super::model::{Frame, Kind};
use super::package::{classify, collect_frames, program_base};
use crate::embedded::object::ObjectSubtype;
use crate::embedded::storage::{Compression, Kind as StorageKind, Storage};
use crate::shapes::{PictureFrameKind, PictureShape, ShapeEnum};
use litchi_ograph::Limits;
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
