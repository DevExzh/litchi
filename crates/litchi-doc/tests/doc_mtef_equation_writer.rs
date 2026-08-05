use litchi_cfb::OleFile;
use litchi_doc::embedded_object::{Info, Limits};
use litchi_doc::writer::{DocPicture, DocWriter};
use litchi_doc::{
    DocEmbeddedObjectEditor, DocMtefEquationWriteOptions, EQUATION_3_CLSID, MtefEquation, Package,
};
use litchi_ole_common::object::{Editor, Limits as ObjectLimits, Target, Targets};
use std::io::Cursor;

const MINIMAL_MTEF_HEX: &str =
    include_str!("../../../test-data/ole/doc/mtef/equation3-minimal.hex");

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
    let mut writer = DocWriter::new();
    writer.add_paragraph("native equation follows").unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}

#[test]
fn authors_native_equation_object_and_preserves_storage_clsid() {
    let payload = decode_hex_fixture();
    let equation = MtefEquation::from_mtef_payload(payload.clone()).unwrap();
    assert_eq!(equation.mtef_payload(), payload);
    let preview = DocPicture::new(preview_png()).unwrap();
    let mut editor = DocEmbeddedObjectEditor::open(base_doc(), Limits::default()).unwrap();
    let reference = editor
        .add_mtef_equation(DocMtefEquationWriteOptions::new(314_159, equation, preview))
        .unwrap();
    assert_eq!(reference.storage_name, "_314159");
    assert_eq!(
        editor
            .objects()
            .unwrap()
            .iter()
            .map(|object| object.storage_id)
            .collect::<Vec<_>>(),
        vec![reference.storage_id]
    );
    let bytes = editor.finish().unwrap();

    let ole = OleFile::open(Cursor::new(bytes.clone())).unwrap();
    let storage = ole
        .list_directory_entries(&["ObjectPool"])
        .unwrap()
        .into_iter()
        .find(|entry| entry.name == "_314159")
        .unwrap();
    assert_eq!(storage.clsid, "0002CE02-0000-0000-C000-000000000046");
    assert!(ole.exists(&["ObjectPool", "_314159", "Equation Native"]));
    assert!(ole.exists(&["ObjectPool", "_314159", "\u{1}CompObj"]));
    assert!(ole.exists(&["ObjectPool", "_314159", "\u{1}Ole"]));
    assert!(ole.exists(&["ObjectPool", "_314159", "\u{3}ObjInfo"]));

    let target = Target::new("_314159", ["ObjectPool", "_314159"]).unwrap();
    let objects =
        Editor::open(bytes.clone(), Targets::one(target), ObjectLimits::default()).unwrap();
    let object = objects.objects().get("_314159").unwrap();
    assert_eq!(object.path(), ["ObjectPool", "_314159"]);
    assert_eq!(object.storage().clsid(), Some(storage.clsid.as_str()));
    assert_eq!(
        object.stream(&["\u{3}ObjInfo"]),
        Some(&[0x00, 0x82, 0x03, 0x00, 0x00, 0x00][..])
    );
    let descriptor = Info::of(object).unwrap().unwrap();
    assert!(descriptor.recompose_on_resize);
    assert!(descriptor.view_object);
    assert_eq!(descriptor.clipboard_format, 3);
    let mut nested = OleFile::open(Cursor::new(object.compound())).unwrap();
    assert_eq!(
        nested.root_entry().unwrap().clsid,
        "0002CE02-0000-0000-C000-000000000046"
    );
    assert_eq!(
        nested.open_stream(&["Equation Native"]).unwrap(),
        MtefEquation::from_mtef_payload(payload)
            .unwrap()
            .equation_native()
    );

    let reopened = DocEmbeddedObjectEditor::open(bytes.clone(), Limits::default()).unwrap();
    assert_eq!(
        reopened
            .objects()
            .unwrap()
            .iter()
            .map(|object| object.storage_id)
            .collect::<Vec<_>>(),
        vec![reference.storage_id]
    );
    let mut package = Package::from_reader(Cursor::new(bytes)).unwrap();
    let document = package.document().unwrap();
    assert!(document.text().unwrap().contains("native equation follows"));
    #[cfg(feature = "formula")]
    assert!(
        document
            .paragraphs()
            .unwrap()
            .iter()
            .any(|paragraph| paragraph.has_formulas())
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
        assert!(MtefEquation::from_mtef_payload(payload).is_err());
    }

    let valid = MtefEquation::from_mtef_payload(decode_hex_fixture()).unwrap();
    let mut trailing = valid.equation_native().to_vec();
    trailing.push(0);
    assert!(MtefEquation::from_equation_native(trailing).is_err());
    let mut wrong_format = valid.equation_native().to_vec();
    wrong_format[6..8].copy_from_slice(&0x0100u16.to_le_bytes());
    assert!(MtefEquation::from_equation_native(wrong_format).is_err());
    let mut oversized = valid.equation_native().to_vec();
    oversized[8..12].copy_from_slice(&(16 * 1024 * 1024 + 1u32).to_le_bytes());
    assert!(MtefEquation::from_equation_native(oversized).is_err());

    assert_ne!(EQUATION_3_CLSID, [0; 16]);
}
