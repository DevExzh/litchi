use litchi_cfb::{OleFile, OleWriter};
use litchi_xls::ole_object::Limits;
use litchi_xls::{
    XlsFtCmo, XlsFtPictFmla, XlsFtPioGrbit, XlsObjSubrecord, XlsOleObjectEditor, XlsOleObjectRecord,
};
use std::io::Cursor;

fn record(kind: u16, body: &[u8]) -> Vec<u8> {
    let mut value = kind.to_le_bytes().to_vec();
    value.extend_from_slice(&(body.len() as u16).to_le_bytes());
    value.extend_from_slice(body);
    value
}

fn object(id: u16, storage: u32, unknown: u8) -> XlsOleObjectRecord {
    XlsOleObjectRecord {
        subrecords: vec![
            XlsObjSubrecord::Common(XlsFtCmo {
                object_type: 8,
                object_id: id,
                flags: 0,
                reserved: [0; 12],
            }),
            XlsObjSubrecord::ClipboardFormat(vec![2, 0]),
            XlsObjSubrecord::PictureFlags(XlsFtPioGrbit { raw: 0 }),
            XlsObjSubrecord::Unknown {
                kind: 0x7777,
                data: vec![unknown],
            },
            XlsObjSubrecord::PictureFormula(XlsFtPictFmla {
                formula: vec![1, 2, 3],
                storage_position: Some(storage),
                control_buffer_size: Some(0),
            }),
            XlsObjSubrecord::End,
        ],
        text_object: None,
    }
}

fn linked_object(id: u16, storage: u32, unknown: u8) -> XlsOleObjectRecord {
    let mut value = object(id, storage, unknown);
    let flags = value
        .subrecords
        .iter_mut()
        .find_map(|subrecord| match subrecord {
            XlsObjSubrecord::PictureFlags(flags) => Some(flags),
            _ => None,
        })
        .expect("test OLE object should contain FtPioGrbit");
    flags.raw = 0x0002;
    value
}

fn nested_cfb(marker: &[u8]) -> Vec<u8> {
    let mut writer = OleWriter::new();
    writer.create_stream(&["CONTENTS"], marker).unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}

fn workbook_stream(objects: &[XlsOleObjectRecord]) -> Vec<u8> {
    let bof = record(0x0809, &[0; 16]);
    let eof = record(0x000A, &[]);
    let mut bound_body = vec![0; 8];
    bound_body[6] = 1;
    bound_body[7] = b'S';
    let bound = record(0x0085, &bound_body);
    let globals_len = bof.len() + bound.len() + eof.len();
    let mut output = bof;
    let mut bound = bound;
    bound[4..8].copy_from_slice(&(globals_len as u32).to_le_bytes());
    output.extend_from_slice(&bound);
    output.extend_from_slice(&eof);
    output.extend_from_slice(&record(0x0809, &[0; 16]));
    output.extend_from_slice(&record(0x7776, b"unknown-sheet-record"));
    for object in objects {
        output.extend_from_slice(&object.to_record_bytes().unwrap());
    }
    output.extend_from_slice(&eof);
    output
}

fn xls(objects: &[XlsOleObjectRecord], storages: &[u32]) -> Vec<u8> {
    let names = storages.iter().map(|id| ("MBD", *id)).collect::<Vec<_>>();
    xls_named(objects, &names)
}

fn xls_named(objects: &[XlsOleObjectRecord], storages: &[(&str, u32)]) -> Vec<u8> {
    let workbook = workbook_stream(objects);
    let payloads = storages
        .iter()
        .map(|(prefix, id)| ((*prefix, *id), nested_cfb(&id.to_le_bytes())))
        .collect::<Vec<_>>();
    let mut writer = OleWriter::new();
    writer.create_stream(&["Workbook"], &workbook).unwrap();
    for ((prefix, id), payload) in &payloads {
        let name = format!("{prefix}{id:08X}");
        writer.create_storage(&[&name]).unwrap();
        let mut nested = OleFile::open(Cursor::new(payload)).unwrap();
        let contents = nested.open_stream(&["CONTENTS"]).unwrap();
        writer
            .create_stream(&[&name, "CONTENTS"], &contents)
            .unwrap();
    }
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}

#[test]
fn opens_link_storage_from_dde_obj_reference() {
    let bytes = xls_named(&[linked_object(1, 42, 1)], &[("LNK", 42)]);
    let editor = XlsOleObjectEditor::new(bytes, Limits::default()).unwrap();
    assert_eq!(editor.objects(0).unwrap().len(), 1);
}

#[test]
fn removes_shared_storage_only_after_last_reference() {
    let bytes = xls(&[object(1, 42, 1), object(2, 42, 2)], &[42]);
    let mut editor = XlsOleObjectEditor::new(bytes, Limits::default()).unwrap();
    editor.remove(0, 1).unwrap();
    let bytes = editor.finish().unwrap();
    let ole = OleFile::open(Cursor::new(bytes.clone())).unwrap();
    assert!(ole.exists(&["MBD0000002A"]));
    let mut editor = XlsOleObjectEditor::new(bytes, Limits::default()).unwrap();
    editor.remove(0, 2).unwrap();
    let bytes = editor.finish().unwrap();
    let ole = OleFile::open(Cursor::new(bytes)).unwrap();
    assert!(!ole.exists(&["MBD0000002A"]));
}

#[test]
fn add_and_reorder_repairs_sheet_offset_and_preserves_unknown_data() {
    let bytes = xls(&[object(1, 1, 0xA1)], &[1]);
    let mut editor = XlsOleObjectEditor::new(bytes, Limits::default()).unwrap();
    editor
        .add(0, object(2, 2, 0xB2), nested_cfb(b"second"))
        .unwrap();
    editor.reorder(0, &[2, 1]).unwrap();
    let bytes = editor.finish().unwrap();
    let editor = XlsOleObjectEditor::new(bytes.clone(), Limits::default()).unwrap();
    assert_eq!(
        editor
            .objects(0)
            .unwrap()
            .iter()
            .map(XlsOleObjectRecord::object_id)
            .collect::<Vec<_>>(),
        vec![2, 1]
    );
    let mut ole = OleFile::open(Cursor::new(bytes)).unwrap();
    assert!(ole.exists(&["MBD00000002"]));
    let workbook = ole.open_stream(&["Workbook"]).unwrap();
    assert!(
        workbook
            .windows(b"unknown-sheet-record".len())
            .any(|value| value == b"unknown-sheet-record")
    );
    assert!(
        editor.objects(0).unwrap()[0]
            .subrecords
            .iter()
            .any(|value| {
                matches!(value, XlsObjSubrecord::Unknown { kind: 0x7777, data } if data == &[0xB2])
            })
    );
}

#[test]
fn invalid_flags_and_reorder_are_atomic() {
    let mut invalid_object = object(1, 1, 0);
    invalid_object.subrecords[2] = XlsObjSubrecord::PictureFlags(XlsFtPioGrbit { raw: 0x12 });
    assert!(invalid_object.validate().is_err());

    let bytes = xls(&[object(1, 1, 1), object(2, 2, 2)], &[1, 2]);
    let mut editor = XlsOleObjectEditor::new(bytes, Limits::default()).unwrap();
    assert!(editor.reorder(0, &[1, 1]).is_err());
    assert_eq!(
        editor
            .objects(0)
            .unwrap()
            .iter()
            .map(XlsOleObjectRecord::object_id)
            .collect::<Vec<_>>(),
        vec![1, 2]
    );
}
