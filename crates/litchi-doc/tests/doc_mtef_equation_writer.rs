//! Equation objects written from raw MTEF through the primary DOC
//! transaction, without the `equation` feature: the object storage and its
//! streams, the field, and the `Equation Native` envelope checks.

use litchi_cfb::OleFile;
use litchi_core::Position;
use litchi_doc::builder::Builder;
use litchi_doc::embedded_object::Info;
use litchi_doc::writer::Picture;
use litchi_doc::{
    EQUATION_3_CLSID, EquationContent, EquationKind, MtefEquationObject, Package, raw::Reader,
};
use litchi_ole_common::object::{
    Edit as ObjectEditor, Limits as ObjectLimits, OpenOptions as ObjectOpenOptions,
    Package as ObjectPackage, Target, Targets,
};
use std::io::Cursor;

const MINIMAL_MTEF_HEX: &str =
    include_str!("../../../test-data/ole/doc/mtef/equation3-minimal.hex");

fn open_object(
    bytes: Vec<u8>,
    targets: Targets,
    limits: ObjectLimits,
) -> Result<ObjectEditor, litchi_ole_common::Error> {
    ObjectPackage::open_with_options(
        bytes,
        ObjectOpenOptions::new().targets(targets).limits(limits),
    )?
    .edit()
}

fn decode_hex_fixture() -> Vec<u8> {
    MINIMAL_MTEF_HEX
        .lines()
        .filter_map(|line| {
            line.split_once('#')
                .map_or(Some(line), |(data, _)| Some(data))
        })
        .flat_map(str::split_whitespace)
        .map(|value| u8::from_str_radix(value, 16).expect("fixture contains hex bytes"))
        .collect()
}

fn preview_png() -> Vec<u8> {
    // 1x1 transparent PNG.
    vec![
        0x89, b'P', b'N', b'G', 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D, b'I', b'H', b'D',
        b'R', 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1F,
        0x15, 0xC4, 0x89, 0x00, 0x00, 0x00, 0x0D, b'I', b'D', b'A', b'T', 0x08, 0xD7, 0x63, 0x60,
        0x60, 0x60, 0x60, 0x00, 0x00, 0x00, 0x05, 0x00, 0x01, 0xE2, 0x26, 0x05, 0x9B, 0x00, 0x00,
        0x00, 0x00, b'I', b'E', b'N', b'D', 0xAE, 0x42, 0x60, 0x82,
    ]
}

fn base_doc() -> Vec<u8> {
    let mut writer = Builder::new();
    writer.add_paragraph("native equation follows").unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.build().unwrap().write_to(&mut output).unwrap();
    output.into_inner()
}

#[test]
fn writes_a_native_equation_object_with_its_class_and_streams() {
    let payload = decode_hex_fixture();
    let equation = MtefEquationObject::from_mtef_payload(payload.clone()).unwrap();
    assert_eq!(equation.mtef_payload(), payload);
    let content = EquationContent::new(equation, Picture::new(preview_png()).unwrap(), 4);
    let source = Package::from_bytes(&base_doc()).unwrap();
    let mut edit = source.edit().unwrap();
    // MTEF 3 is an Equation Editor 3.0 object, never a MathType one.
    assert!(
        edit.insert_equation(Position::new(0), 6, EquationKind::MathType, &content)
            .is_err()
    );
    let storage_id = edit
        .insert_equation(Position::new(0), 6, EquationKind::EquationEditor3, &content)
        .unwrap();
    let bytes = edit.commit().unwrap().package().to_bytes().unwrap();
    let name = format!("_{storage_id}");

    let ole = OleFile::open(
        Cursor::new(bytes.clone()),
        litchi_cfb::OleFileLimits::default(),
    )
    .unwrap();
    let storage = ole
        .list_directory_entries(&["ObjectPool"])
        .unwrap()
        .into_iter()
        .find(|entry| entry.name == name)
        .unwrap();
    assert_eq!(storage.clsid, "0002CE02-0000-0000-C000-000000000046");
    for stream in [
        "Equation Native",
        "\u{1}CompObj",
        "\u{1}Ole",
        "\u{3}ObjInfo",
    ] {
        assert!(ole.exists(&["ObjectPool", &name, stream]), "{stream}");
    }

    let target = Target::new(name.clone(), ["ObjectPool", name.as_str()]).unwrap();
    let objects =
        open_object(bytes.clone(), Targets::one(target), ObjectLimits::default()).unwrap();
    let object = objects.objects().get(&name).unwrap();
    assert_eq!(object.storage().clsid(), Some(storage.clsid.as_str()));
    // The ObjInfo Word writes for an equation: a metafile presentation.
    assert_eq!(
        object.stream(&["\u{3}ObjInfo"]),
        Some(&[0x00, 0x00, 0x03, 0x00, 0x04, 0x00][..])
    );
    assert_eq!(Info::of(object).unwrap().unwrap().clipboard_format, 3);
    let comp_obj = object.stream(&["\u{1}CompObj"]).unwrap();
    assert!(
        comp_obj
            .windows(b"Equation.3\0".len())
            .any(|window| window == b"Equation.3\0")
    );
    let mut nested = OleFile::open(
        Cursor::new(object.compound()),
        litchi_cfb::OleFileLimits::default(),
    )
    .unwrap();
    assert_eq!(
        nested.root_entry().unwrap().clsid,
        "0002CE02-0000-0000-C000-000000000046"
    );
    assert_eq!(
        nested.open_stream(&["Equation Native"]).unwrap(),
        MtefEquationObject::from_mtef_payload(payload)
            .unwrap()
            .equation_native()
    );

    let mut package = Reader::from_reader(Cursor::new(bytes)).unwrap();
    let document = package.document().unwrap();
    assert!(
        document
            .text()
            .unwrap()
            .contains("native\u{13} EMBED Equation.3  \u{14}\u{1}\u{15} equation follows")
    );
}

#[test]
fn rejects_malformed_or_oversized_equation_native_envelopes() {
    for payload in [
        vec![],
        vec![3, 1, 1, 3, 0],
        vec![6, 1, 1, 3, 0, 0],
        vec![5, 1, 1, 3, 0, b'x', 1, 0],
    ] {
        assert!(MtefEquationObject::from_mtef_payload(payload).is_err());
    }

    let valid = MtefEquationObject::from_mtef_payload(decode_hex_fixture()).unwrap();
    let mut trailing = valid.equation_native().to_vec();
    trailing.push(0);
    assert!(MtefEquationObject::from_equation_native(trailing).is_err());
    let mut wrong_format = valid.equation_native().to_vec();
    wrong_format[6..8].copy_from_slice(&0x0100u16.to_le_bytes());
    assert!(MtefEquationObject::from_equation_native(wrong_format).is_err());
    let mut oversized = valid.equation_native().to_vec();
    oversized[8..12].copy_from_slice(&(16 * 1024 * 1024 + 1u32).to_le_bytes());
    assert!(MtefEquationObject::from_equation_native(oversized).is_err());

    assert_ne!(EQUATION_3_CLSID, [0; 16]);
}
