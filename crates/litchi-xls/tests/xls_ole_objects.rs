use litchi_cfb::{OleFile, OleWriter};
use litchi_xls::{FtCf, FtCmo, FtPictFmla, FtPioGrbit, ObjSubrecord, OleObjectRecord};
use std::io::Cursor;

fn record(kind: u16, body: &[u8]) -> Vec<u8> {
    let mut value = kind.to_le_bytes().to_vec();
    value.extend_from_slice(&(body.len() as u16).to_le_bytes());
    value.extend_from_slice(body);
    value
}

fn object(id: u16, storage: u32, unknown: u8) -> OleObjectRecord {
    OleObjectRecord {
        subrecords: vec![
            ObjSubrecord::Common(FtCmo {
                object_type: 8,
                object_id: id,
                flags: 0,
                reserved: [0; 12],
            }),
            ObjSubrecord::PictureFormat(FtCf {
                format: FtCf::ENHANCED_METAFILE,
            }),
            ObjSubrecord::PictureFlags(FtPioGrbit { raw: 0 }),
            ObjSubrecord::Unknown {
                kind: 0x7777,
                data: vec![unknown],
            },
            ObjSubrecord::PictureFormula(FtPictFmla {
                formula: vec![
                    0x05, 0x00, 0x00, 0x00, 0x00, 0x00, 0x02, 0x00, 0x00, 0x00, 0x00, 0x03, 0x00,
                    0x00,
                ],
                storage_position: Some(storage),
                control_buffer_size: None,
            }),
            ObjSubrecord::End,
        ],
        text_object: None,
    }
}

fn linked_object(id: u16, storage: u32, unknown: u8) -> OleObjectRecord {
    let mut value = object(id, storage, unknown);
    let flags = value
        .subrecords
        .iter_mut()
        .find_map(|subrecord| match subrecord {
            ObjSubrecord::PictureFlags(flags) => Some(flags),
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

/// A complete authored workbook stream (one worksheet `S`) with `records`
/// appended to the worksheet substream before its EOF, so the package-level
/// OLE-object family opens it.
fn authored_workbook_stream(records: &[u8]) -> Vec<u8> {
    let mut builder = litchi_xls::Builder::new();
    builder.add_worksheet("S").unwrap();
    let bytes = builder.build().unwrap().to_bytes().unwrap();
    let mut ole = OleFile::open(Cursor::new(bytes), litchi_cfb::OleFileLimits::default()).unwrap();
    let mut workbook = ole.open_stream(&["Workbook"]).unwrap();
    let eof = workbook.len() - 4;
    assert_eq!(&workbook[eof..], &[0x0A, 0x00, 0x00, 0x00]);
    workbook.splice(eof..eof, records.iter().copied());
    workbook
}

fn xls(objects: &[OleObjectRecord], storages: &[u32]) -> Vec<u8> {
    let names = storages.iter().map(|id| ("MBD", *id)).collect::<Vec<_>>();
    xls_named(objects, &names)
}

fn xls_named(objects: &[OleObjectRecord], storages: &[(&str, u32)]) -> Vec<u8> {
    let mut records = record(0x7776, b"unknown-sheet-record");
    for object in objects {
        records.extend_from_slice(&object.to_record_bytes().unwrap());
    }
    let workbook = authored_workbook_stream(&records);
    let payloads = storages
        .iter()
        .map(|(prefix, id)| ((*prefix, *id), nested_cfb(&id.to_le_bytes())))
        .collect::<Vec<_>>();
    let mut writer = OleWriter::new();
    writer.create_stream(&["Workbook"], &workbook).unwrap();
    for ((prefix, id), payload) in &payloads {
        let name = format!("{prefix}{id:08X}");
        writer.create_storage(&[&name]).unwrap();
        let mut nested =
            OleFile::open(Cursor::new(payload), litchi_cfb::OleFileLimits::default()).unwrap();
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
    let mut editor = litchi_xls::Package::from_bytes(&bytes)
        .unwrap()
        .edit_ole_objects()
        .unwrap();
    assert_eq!(editor.objects(0).unwrap().len(), 1);
}

#[test]
fn removes_shared_storage_only_after_last_reference() {
    let bytes = xls(&[object(1, 42, 1), object(2, 42, 2)], &[42]);
    let mut editor = litchi_xls::Package::from_bytes(&bytes)
        .unwrap()
        .edit_ole_objects()
        .unwrap();
    editor.remove_object(0, 1).unwrap();
    let bytes = editor.commit().unwrap().patch().after().to_vec();
    let ole = OleFile::open(
        Cursor::new(bytes.clone()),
        litchi_cfb::OleFileLimits::default(),
    )
    .unwrap();
    assert!(ole.exists(&["MBD0000002A"]));
    let mut editor = litchi_xls::Package::from_bytes(&bytes)
        .unwrap()
        .edit_ole_objects()
        .unwrap();
    editor.remove_object(0, 2).unwrap();
    let bytes = editor.commit().unwrap().patch().after().to_vec();
    let ole = OleFile::open(Cursor::new(bytes), litchi_cfb::OleFileLimits::default()).unwrap();
    assert!(!ole.exists(&["MBD0000002A"]));
}

#[test]
fn add_and_reorder_repairs_sheet_offset_and_preserves_unknown_data() {
    let bytes = xls(&[object(1, 1, 0xA1)], &[1]);
    let mut editor = litchi_xls::Package::from_bytes(&bytes)
        .unwrap()
        .edit_ole_objects()
        .unwrap();
    editor
        .add_object(0, object(2, 2, 0xB2), nested_cfb(b"second"))
        .unwrap();
    editor.reorder_objects(0, &[2, 1]).unwrap();
    let bytes = editor.commit().unwrap().patch().after().to_vec();
    let mut editor = litchi_xls::Package::from_bytes(&bytes)
        .unwrap()
        .edit_ole_objects()
        .unwrap();
    assert_eq!(
        editor
            .objects(0)
            .unwrap()
            .iter()
            .map(OleObjectRecord::object_id)
            .collect::<Vec<_>>(),
        vec![2, 1]
    );
    let mut ole = OleFile::open(Cursor::new(bytes), litchi_cfb::OleFileLimits::default()).unwrap();
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
                matches!(value, ObjSubrecord::Unknown { kind: 0x7777, data } if data == &[0xB2])
            })
    );
}

#[test]
fn invalid_flags_and_reorder_are_atomic() {
    for raw in [0x0010, 0x0020, 0x0012] {
        let mut invalid_object = object(1, 1, 0);
        invalid_object.subrecords[2] = ObjSubrecord::PictureFlags(FtPioGrbit { raw });
        assert!(invalid_object.validate().is_err());
    }

    let bytes = xls(&[object(1, 1, 1), object(2, 2, 2)], &[1, 2]);
    let mut editor = litchi_xls::Package::from_bytes(&bytes)
        .unwrap()
        .edit_ole_objects()
        .unwrap();
    assert!(editor.reorder_objects(0, &[1, 1]).is_err());
    assert_eq!(
        editor
            .objects(0)
            .unwrap()
            .iter()
            .map(OleObjectRecord::object_id)
            .collect::<Vec<_>>(),
        vec![1, 2]
    );
}
