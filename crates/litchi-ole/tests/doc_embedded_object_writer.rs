use litchi_ole::doc::writer::DocWriter;
use litchi_ole::doc::{DocEmbeddedObjectEditor, DocEmbeddedObjectWriteOptions, Package};
use litchi_ole::{LegacyOfficeObjectLimits, OleFile, OleWriter};
use std::io::Cursor;
use std::fs;
use std::path::PathBuf;

fn base_doc() -> Vec<u8> {
    let mut writer = DocWriter::new();
    writer.add_paragraph("before embedded objects").unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}

fn object_cfb(marker: &[u8]) -> Vec<u8> {
    let mut writer = OleWriter::new();
    writer.create_stream(&["CONTENTS"], marker).unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}

fn inert_picf(marker: u32) -> Vec<u8> {
    let mut value = 12u32.to_le_bytes().to_vec();
    value.extend_from_slice(&marker.to_le_bytes());
    value.extend_from_slice(&[0; 4]);
    value
}

fn options(id: u32) -> DocEmbeddedObjectWriteOptions {
    DocEmbeddedObjectWriteOptions::new(id, object_cfb(&id.to_le_bytes()), inert_picf(id))
}

#[test]
fn managed_objects_add_reorder_remove_and_reopen_transactionally() {
    let original = base_doc();
    let mut editor = DocEmbeddedObjectEditor::open(original, LegacyOfficeObjectLimits::default()).unwrap();
    editor.add(options(11)).unwrap();
    editor.add(options(22)).unwrap();
    assert_eq!(editor.objects().unwrap().iter().map(|value| value.storage_id).collect::<Vec<_>>(), vec![11, 22]);
    editor.reorder(&[22, 11]).unwrap();
    assert_eq!(editor.objects().unwrap().iter().map(|value| value.storage_id).collect::<Vec<_>>(), vec![22, 11]);
    let bytes = editor.finish().unwrap();

    let mut reopened = DocEmbeddedObjectEditor::open(bytes, LegacyOfficeObjectLimits::default()).unwrap();
    assert_eq!(reopened.objects().unwrap().iter().map(|value| value.storage_id).collect::<Vec<_>>(), vec![22, 11]);
    reopened.remove(22).unwrap();
    let bytes = reopened.finish().unwrap();
    let ole = OleFile::open(Cursor::new(bytes.clone())).unwrap();
    assert!(!ole.exists(&["ObjectPool", "_22"]));
    assert!(ole.exists(&["ObjectPool", "_11"]));
    let mut package = Package::from_reader(Cursor::new(bytes)).unwrap();
    let document = package.document().unwrap();
    assert!(document.text().unwrap().contains("before embedded objects"));
}

#[test]
fn malformed_add_and_reorder_leave_editor_state_unchanged() {
    let mut editor = DocEmbeddedObjectEditor::open(base_doc(), LegacyOfficeObjectLimits::default()).unwrap();
    editor.add(options(1)).unwrap();
    editor.add(options(2)).unwrap();
    let before = editor.objects().unwrap();
    let mut invalid = options(3);
    invalid.picture_data[0..4].copy_from_slice(&999u32.to_le_bytes());
    assert!(editor.add(invalid).is_err());
    assert!(editor.reorder(&[1, 1]).is_err());
    assert_eq!(editor.objects().unwrap(), before);
}

#[test]
fn producer_fixtures_append_only_when_their_layout_is_supported() {
    let root = PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../..");
    let fixtures = [
        root.join("test-data/poi/test-data/document/word_with_embeded.doc"),
        root.join("test-data/libreoffice-core/embeddedobj/qa/cppunit/data/insert-file-config.doc"),
    ];
    let mut supported = 0usize;
    let mut rejected = 0usize;
    for (ordinal, path) in fixtures.into_iter().enumerate() {
        let original = fs::read(&path).unwrap();
        match DocEmbeddedObjectEditor::open(original.clone(), LegacyOfficeObjectLimits::default()) {
            Ok(mut editor) => {
                let id = 2_000_000 + ordinal as u32;
                editor.add(options(id)).unwrap();
                let output = editor.finish().unwrap();
                let reopened = DocEmbeddedObjectEditor::open(output.clone(), LegacyOfficeObjectLimits::default()).unwrap();
                assert!(reopened.objects().unwrap().iter().any(|value| value.storage_id == id));
                let mut package = Package::from_reader(Cursor::new(output)).unwrap();
                assert!(!package.document().unwrap().text().unwrap().is_empty());
                supported += 1;
            }
            Err(_) => {
                // Construction owns no external state and therefore cannot mutate the fixture.
                assert_eq!(fs::read(&path).unwrap(), original);
                rejected += 1;
            }
        }
    }
    assert_eq!(supported + rejected, 2);
}
