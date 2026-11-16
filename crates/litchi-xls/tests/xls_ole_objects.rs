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

/// The worksheet records of a compound file's Workbook stream, as record
/// type and payload, in order.
fn worksheet_records(bytes: &[u8]) -> Vec<(u16, Vec<u8>)> {
    let mut ole = OleFile::open(Cursor::new(bytes), litchi_cfb::OleFileLimits::default()).unwrap();
    let workbook = ole.open_stream(&["Workbook"]).unwrap();
    let mut out = Vec::new();
    let mut at = 0;
    let mut substream = 0;
    while at + 4 <= workbook.len() {
        let kind = u16::from_le_bytes([workbook[at], workbook[at + 1]]);
        let length = usize::from(u16::from_le_bytes([workbook[at + 2], workbook[at + 3]]));
        if kind == 0x0809 {
            substream += 1;
        }
        if substream > 1 {
            out.push((kind, workbook[at + 4..at + 4 + length].to_vec()));
        }
        at += 4 + length;
    }
    out
}

/// For each `Obj` with an `FtCmo` identifier, the payload of the record
/// right before it (its `MsoDrawing` shape).
fn shapes_by_object(bytes: &[u8]) -> Vec<(u16, u16, Vec<u8>)> {
    let records = worksheet_records(bytes);
    records
        .iter()
        .enumerate()
        .filter(|(_, (kind, payload))| *kind == 0x005D && payload.len() >= 8)
        .map(|(index, (_, payload))| {
            let id = u16::from_le_bytes([payload[6], payload[7]]);
            let previous = index
                .checked_sub(1)
                .map(|previous| records[previous].clone())
                .unwrap_or_default();
            (id, previous.0, previous.1)
        })
        .collect()
}

#[test]
fn removing_an_excel_embedded_object_keeps_every_other_obj_with_its_own_shape() {
    let bytes = std::fs::read(
        std::path::PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/ole/xls/WithEmbeddedObjects.xls"),
    )
    .unwrap();
    let source = shapes_by_object(&bytes);
    assert_eq!(source.len(), 2);
    for (removed, kept) in [(0, 1), (1, 0)] {
        let mut editor = litchi_xls::Package::from_bytes(&bytes)
            .unwrap()
            .edit_ole_objects()
            .unwrap();
        let storage = editor.objects(0).unwrap()[removed].storage_name().unwrap();
        editor.remove_object(0, source[removed].0).unwrap();
        let after = editor.commit().unwrap().patch().after().to_vec();
        let shapes = shapes_by_object(&after);
        assert_eq!(shapes.len(), 1);
        let (id, kind, shape) = &shapes[0];
        assert_eq!(*id, source[kept].0);
        // The kept Obj still follows its own shape (anchor and client data
        // end the shape container).
        assert_eq!(*kind, 0x00EC);
        let tail = &source[kept].2[source[kept].2.len() - 34..];
        assert!(shape.ends_with(tail));
        let ole = OleFile::open(Cursor::new(after), litchi_cfb::OleFileLimits::default()).unwrap();
        assert!(!ole.exists(&[storage.as_str()]));
    }
}

#[test]
fn removing_an_object_without_a_drawing_shape_is_refused() {
    let bytes = xls(&[object(1, 42, 1)], &[42]);
    let mut editor = litchi_xls::Package::from_bytes(&bytes)
        .unwrap()
        .edit_ole_objects()
        .unwrap();
    assert!(editor.remove_object(0, 1).is_err());
    assert_eq!(editor.objects(0).unwrap().len(), 1);
}

/// One row per order of the picture fields: `c` FtCmo, `f` pictFormat,
/// `g` pictFlags, `p` pictFmla, `u` an unknown subrecord, `e` FtEnd.
#[test]
fn picture_field_order_accepts_and_rejects_the_same_sequences() {
    const FORMAT_ORDER: &str =
        "Invalid record 0x005D: OLE Obj pictFormat is out of MS-XLS field order";
    const FLAGS_ORDER: &str =
        "Invalid record 0x005D: OLE Obj pictFlags is out of MS-XLS field order";
    const INCOMPLETE: &str =
        "Invalid record 0x005D: OLE Obj mandatory picture fields are incomplete";
    const ONE_FLAGS: &str = "Invalid record 0x005D: OLE Obj requires one FtPioGrbit";
    const ONE_FORMAT: &str = "Invalid record 0x005D: OLE Obj requires exactly one FtCf pictFormat";
    const ONE_FORMULA: &str = "Invalid record 0x005D: duplicate FtPictFmla";
    let cases: &[(&str, &str)] = &[
        ("cfgupe", "ok"),
        ("cfgpe", "ok"),
        ("cufugupue", "ok"),
        ("cgfpe", FORMAT_ORDER),
        ("cfpge", FLAGS_ORDER),
        ("cpfge", FORMAT_ORDER),
        ("cgpfe", FORMAT_ORDER),
        ("cpgfe", FLAGS_ORDER),
        ("cfgpfe", FORMAT_ORDER),
        ("cfgpge", FLAGS_ORDER),
        ("ce", INCOMPLETE),
        ("cfe", INCOMPLETE),
        ("cfge", INCOMPLETE),
        ("cge", INCOMPLETE),
        ("cpe", ONE_FLAGS),
        ("cfpe", ONE_FLAGS),
        ("cgpe", ONE_FORMAT),
        ("cffgpe", ONE_FORMAT),
        ("cfggpe", ONE_FLAGS),
        ("cfgppe", ONE_FORMULA),
        ("cfgpupe", ONE_FORMULA),
    ];
    let template = object(1, 1, 0);
    for (order, expected) in cases {
        let subrecords = order
            .chars()
            .map(|field| {
                let index = match field {
                    'c' => 0,
                    'f' => 1,
                    'g' => 2,
                    'u' => 3,
                    'p' => 4,
                    'e' => 5,
                    other => panic!("unknown field {other}"),
                };
                template.subrecords[index].clone()
            })
            .collect();
        let record = OleObjectRecord {
            subrecords,
            text_object: None,
        };
        let outcome = match record.validate() {
            Ok(()) => "ok".to_string(),
            Err(error) => error.to_string(),
        };
        assert_eq!(outcome, *expected, "{order}");
    }
}

#[test]
fn invalid_flags_are_rejected() {
    for raw in [0x0010, 0x0020, 0x0012] {
        let mut invalid_object = object(1, 1, 0);
        invalid_object.subrecords[2] = ObjSubrecord::PictureFlags(FtPioGrbit { raw });
        assert!(invalid_object.validate().is_err());
    }
}

const PNG: &[u8] = &[
    0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
    0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1f, 0x15, 0xc4,
    0x89, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9c, 0x63, 0x60, 0x00, 0x02, 0x00,
    0x00, 0x05, 0x00, 0x01, 0xe9, 0xfa, 0xdc, 0xd8, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44,
    0xae, 0x42, 0x60, 0x82,
];

fn fixture(path: &str) -> Vec<u8> {
    std::fs::read(
        std::path::PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .join("../..")
            .join(path),
    )
    .unwrap()
}

fn draft() -> litchi_xls::EmbeddedObjectDraft {
    litchi_xls::EmbeddedObjectDraft::new(
        litchi_xls::EmbeddedPayload::new(
            nested_cfb(b"later sheet"),
            litchi_xls::ole_object::Limits::default(),
        )
        .unwrap(),
        litchi_xls::sheet_features::edit::Anchor {
            first_row: 40,
            first_column: 1,
            last_row: 44,
            last_column: 3,
            behavior: litchi_xls::AnchorBehavior::MoveWithoutResize,
        },
        PNG.to_vec(),
    )
}

/// `Obj` identifiers are unique within a sheet ([MS-XLS] 2.5.143), so an
/// object added to a later worksheet may reuse an identifier of another
/// sheet, and the result reopens.
#[test]
fn objects_are_added_to_every_worksheet_of_excel_files() {
    for path in [
        "test-data/ole/xls/WithEmbeddedObjects.xls",
        "test-data/microsoft-office/excel/excel-pivot-controls.xls",
        "test-data/microsoft-office/excel/excel-visual-analytics.xls",
    ] {
        let package = litchi_xls::Package::from_bytes(&fixture(path)).unwrap();
        let worksheets = package.ole_objects().unwrap().worksheet_count();
        assert!(worksheets > 1, "{path}");
        for worksheet in 0..worksheets {
            let mut editor = package.edit_ole_objects().unwrap();
            let before = editor.objects(worksheet).unwrap().len();
            let id = editor
                .add_embedded_payload(worksheet, draft())
                .unwrap_or_else(|error| panic!("{path} sheet {worksheet}: {error}"));
            let after = editor.commit().unwrap().into_package();
            let snapshot = after.ole_objects().unwrap();
            let objects = snapshot.objects(worksheet).unwrap();
            assert_eq!(objects.len(), before + 1, "{path} sheet {worksheet}");
            assert_eq!(objects[before].object_id(), id);
        }
    }
}

/// Two `Obj` records, each after its own drawing shape, name one storage:
/// the storage stays until the last of them is removed.
#[test]
fn a_shared_storage_stays_until_its_last_object_is_removed() {
    let bytes = fixture("test-data/ole/xls/WithEmbeddedObjects.xls");
    let snapshot = litchi_xls::Package::from_bytes(&bytes)
        .unwrap()
        .ole_objects()
        .unwrap();
    let objects = snapshot.objects(0).unwrap();
    let (first, second) = (objects[0].object_id(), objects[1].object_id());
    let kept = objects[0].storage_name().unwrap();
    let dropped = objects[1].storage_name().unwrap();
    let kept_position = objects[0].storage_position().unwrap();
    let dropped_position = objects[1].storage_position().unwrap();
    // Point the second Obj at the first storage.
    let mut ole = OleFile::open(
        Cursor::new(bytes.clone()),
        litchi_cfb::OleFileLimits::default(),
    )
    .unwrap();
    let mut writer = OleWriter::new();
    for path in ole.list_streams() {
        let parts: Vec<&str> = path.iter().map(String::as_str).collect();
        if parts[0] == dropped {
            continue;
        }
        let mut data = ole.open_stream(&parts).unwrap();
        if parts == ["Workbook"] {
            let needle = dropped_position.to_le_bytes();
            let at = data.windows(4).position(|window| window == needle).unwrap();
            data[at..at + 4].copy_from_slice(&kept_position.to_le_bytes());
        }
        for depth in 1..parts.len() {
            let _ = writer.create_storage(&parts[..depth]);
        }
        writer.create_stream(&parts, &data).unwrap();
    }
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    let shared = litchi_xls::Package::from_bytes(&output.into_inner()).unwrap();
    let names: Vec<_> = shared
        .ole_objects()
        .unwrap()
        .objects(0)
        .unwrap()
        .iter()
        .map(|object| object.storage_name().unwrap())
        .collect();
    assert_eq!(names, [kept.clone(), kept.clone()]);

    let mut editor = shared.edit_ole_objects().unwrap();
    editor.remove_object(0, first).unwrap();
    let one = editor.commit().unwrap().into_package();
    let bytes = one.to_bytes().unwrap();
    let ole = OleFile::open(Cursor::new(bytes), litchi_cfb::OleFileLimits::default()).unwrap();
    assert!(ole.exists(&[kept.as_str()]));
    assert_eq!(shapes_by_object(&one.to_bytes().unwrap()).len(), 1);

    let mut editor = one.edit_ole_objects().unwrap();
    editor.remove_object(0, second).unwrap();
    let none = editor.commit().unwrap().into_package();
    let ole = OleFile::open(
        Cursor::new(none.to_bytes().unwrap()),
        litchi_cfb::OleFileLimits::default(),
    )
    .unwrap();
    assert!(!ole.exists(&[kept.as_str()]));
}
