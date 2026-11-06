//! Explicit child parsing and publication from native Excel hosts; no activation.
#![cfg(all(feature = "docx", feature = "xls", feature = "xlsx", feature = "xlsb"))]
use litchi::{File, FileFormat};
fn native(name: &str) -> Vec<u8> {
    std::fs::read(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/microsoft-office")
            .join(name),
    )
    .unwrap()
}
fn word_payload(bytes: &[u8], format: FileFormat) -> Vec<u8> {
    match format {
        FileFormat::Xls => {
            use litchi::ole::object::{OpenOptions, Package, Target, Targets};
            let host = litchi::xls::Package::from_bytes(bytes).unwrap();
            let inventory = host.ole_objects().unwrap();
            let name = inventory.objects(0).unwrap()[0].storage_name().unwrap();
            let package = Package::open_with_options(
                bytes.to_vec(),
                OpenOptions::new()
                    .targets(Targets::new([Target::new(&name, [name.clone()]).unwrap()]).unwrap()),
            )
            .unwrap();
            package
                .objects()
                .get(&name)
                .unwrap()
                .stream(&["Package"])
                .unwrap()
                .to_vec()
        },
        FileFormat::Xlsx => {
            use litchi::xlsx::ole_objects::{OleObjectTarget, Snapshot};
            let package = litchi::xlsx::Package::from_bytes(bytes)
                .unwrap()
                .into_plain_opc();
            let snapshot = Snapshot::load(
                &package,
                &litchi::opc::PackURI::new("/xl/worksheets/sheet1.xml").unwrap(),
            )
            .unwrap();
            let Some(OleObjectTarget::Internal(payload)) =
                &snapshot.objects().unwrap().objects[0].target
            else {
                panic!("native internal package")
            };
            payload.data.clone()
        },
        FileFormat::Xlsb => {
            let p = litchi::xlsb::Package::from_bytes(bytes).unwrap();
            let wb = p.workbook().unwrap();
            let entries = wb.embedded().unwrap();
            assert_eq!(entries.len(), 1);
            let litchi::ooxml_common::embedded::Target::Internal(payload) = entries[0].target()
            else {
                panic!("native internal package")
            };
            payload.bytes().to_vec()
        },
        _ => unreachable!(),
    }
}
fn assert_form(child: &File, first: &str) {
    assert_eq!(child.format(), Some(FileFormat::Docx));
    assert_eq!(
        child
            .get("/body/p[1]")
            .unwrap()
            .view()
            .text()
            .unwrap()
            .as_deref(),
        Some(first)
    );
    let text = child.text().unwrap();
    for marker in [
        "Fixture ID: word-legacy-form-010",
        "TEXT_FIELD_010: FORM_VALUE_010",
        "CHECKBOX_FIELD_010:",
        "DROPDOWN_FIELD_010:",
        "END_FORM_010",
    ] {
        assert!(text.contains(marker), "missing {marker}");
    }
    let paragraphs = child
        .query("p")
        .unwrap()
        .collect::<litchi::Result<Vec<_>>>()
        .unwrap();
    assert_eq!(paragraphs.len(), 6);
    for p in paragraphs {
        assert_eq!(child.get(p.path().to_path()).unwrap().path(), p.path());
    }
}
#[test]
fn native_excel_embedded_word_payloads_open_with_identical_form_semantics() {
    for (ext, format) in [
        ("xls", FileFormat::Xls),
        ("xlsx", FileFormat::Xlsx),
        ("xlsb", FileFormat::Xlsb),
    ] {
        let bytes = native(&format!("excel/excel-links-objects.{ext}"));
        let payload = word_payload(&bytes, format);
        assert!(payload == native("word/word-legacy-form-unprotected.docx"));
        let child = litchi::open_bytes(payload.clone()).unwrap();
        assert_form(&child, "WORD_LEGACY_FORM_010");
        assert!(child.to_bytes().unwrap() == payload);
    }
}
#[test]
fn edited_native_word_child_republishes_into_xlsb_without_refreshing_parent_caches() {
    let bytes = native("excel/excel-links-objects.xlsb");
    let payload = word_payload(&bytes, FileFormat::Xlsb);
    let child = litchi::open_bytes(payload.clone()).unwrap();
    let changed = child
        .set(
            "/body/p[1]",
            litchi::props! {"text"=>"EMBEDDED_FORM_REVIEWED"},
        )
        .unwrap();
    assert_form(&changed, "EMBEDDED_FORM_REVIEWED");
    let host = litchi::xlsb::Package::from_bytes(&bytes).unwrap();
    let mut edit = host.edit_opc().unwrap();
    let part = litchi::opc::PackURI::new("/xl/embeddings/Microsoft_Word_Document.docx").unwrap();
    let changed = changed.to_bytes().unwrap();
    edit.update(|opc| {
        opc.get_part_mut(&part)?.set_blob(changed.clone());
        Ok(())
    })
    .unwrap();
    let commit = edit.commit().unwrap();
    let output = tempfile::NamedTempFile::with_suffix(".xlsb").unwrap();
    commit.package().save(output.path()).unwrap();
    let after = std::fs::read(output.path()).unwrap();
    let reopened = litchi::open_bytes(word_payload(&after, FileFormat::Xlsb)).unwrap();
    assert_form(&reopened, "EMBEDDED_FORM_REVIEWED");
    let parent = litchi::open_bytes(after.clone()).unwrap();
    assert_eq!(parent.format(), Some(FileFormat::Xlsb));
    assert_eq!(
        parent
            .get("Sheet1!B2")
            .unwrap()
            .view()
            .text()
            .unwrap()
            .as_deref(),
        Some("49.75")
    );
    assert_eq!(
        parent
            .get("Sheet1!B3")
            .unwrap()
            .view()
            .text()
            .unwrap()
            .as_deref(),
        Some("14")
    );
    let inverse = litchi::xlsb::Package::from_bytes(&after)
        .unwrap()
        .apply_opc(&commit.patch().inverse())
        .unwrap();
    assert!(word_payload(&inverse.package().to_bytes().unwrap(), FileFormat::Xlsb) == payload);
    assert!(host.to_bytes().unwrap() == bytes);
    assert!(child.to_bytes().unwrap() == payload);
}
