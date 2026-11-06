//! Native Excel OLE metadata and inert package payloads.
use litchi_xlsx::{
    Package,
    ole_objects::{OleObjectRelationshipKind, OleObjectTarget, Snapshot},
};
use sha2::{Digest, Sha256};
fn native(name: &str) -> Vec<u8> {
    std::fs::read(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/microsoft-office")
            .join(name),
    )
    .unwrap()
}
fn hash(bytes: &[u8]) -> String {
    Sha256::digest(bytes)
        .iter()
        .map(|b| format!("{b:02x}"))
        .collect()
}
fn objects(p: &Package) -> Snapshot {
    Snapshot::load(
        &p.clone().into_plain_opc(),
        &litchi_opc::PackURI::new("/xl/worksheets/sheet1.xml").unwrap(),
    )
    .unwrap()
}
#[test]
fn native_word_object_keeps_package_preview_and_anchor_identity() {
    let bytes = native("excel/excel-links-objects.xlsx");
    let p = Package::from_bytes(&bytes).unwrap();
    let snapshot = objects(&p);
    let graph = snapshot.objects().unwrap();
    assert_eq!(graph.objects.len(), 1);
    let o = &graph.objects[0];
    assert_eq!(o.program_id.as_deref(), Some("Document"));
    assert_eq!(o.shape_id, 1027);
    assert_eq!(o.relationship_id, "rId3");
    assert_eq!(o.relationship_kind, OleObjectRelationshipKind::Package);
    let Some(OleObjectTarget::Internal(payload)) = &o.target else {
        panic!("internal Word package")
    };
    assert_eq!(
        payload.part_name,
        "/xl/embeddings/Microsoft_Word_Document.docx"
    );
    assert_eq!(
        payload.content_type,
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    );
    assert_eq!(payload.data.len(), 13759);
    assert_eq!(
        hash(&payload.data),
        "a4b37e93f7cf0c93213e419cb42796e361785f4354d50a585dcd5b60ed910778"
    );
    assert!(payload.data == native("word/word-legacy-form-unprotected.docx"));
    let properties = o.properties.as_ref().unwrap();
    assert_eq!(properties.preview_relationship_id, "rId4");
    let preview = properties.preview.as_ref().unwrap();
    assert_eq!(preview.part_name, "/xl/media/image2.emf");
    assert_eq!(preview.data.len(), 69384);
    assert_eq!(
        hash(&preview.data),
        "828d607c3857f92f8339115446c50beef2346cd3bdf796a22b941ca3c84aed37"
    );
    let anchor = &properties.anchor;
    assert_eq!(
        (anchor.move_with_cells, anchor.size_with_cells),
        (Some(true), None)
    );
    assert_eq!(
        (
            anchor.from.column,
            anchor.from.column_offset,
            anchor.from.row,
            anchor.from.row_offset
        ),
        (0, 0, 11, 0)
    );
    assert_eq!(
        (
            anchor.to.column,
            anchor.to.column_offset,
            anchor.to.row,
            anchor.to.row_offset
        ),
        (2, 635000, 20, 114300)
    );
    assert_eq!(properties.default_size, Some(false));
    assert!(o.link.is_none());
    assert!(
        p.edit_ole_objects("Sheet1")
            .unwrap()
            .commit()
            .unwrap()
            .package()
            .to_bytes()
            .unwrap()
            == bytes
    );
    assert!(p.to_bytes().unwrap() == bytes);
}
#[test]
fn native_object_metadata_edit_changes_only_worksheet_xml_and_inverts() {
    let bytes = native("excel/excel-links-objects.xlsx");
    let p = Package::from_bytes(&bytes).unwrap();
    let before = objects(&p);
    let mut expected = before.objects().unwrap().clone();
    expected.objects[0].properties.as_mut().unwrap().alt_text =
        Some("Embedded form reviewed".into());
    let mut edit = p.edit_ole_objects("Sheet1").unwrap();
    assert!(
        edit.edit_object(1027, |o| {
            o.properties.as_mut().unwrap().alt_text = Some("Embedded form reviewed".into());
            Ok(())
        })
        .unwrap()
    );
    assert!(
        edit.edit_object(1027, |o| {
            o.shape_id = 999;
            Ok(())
        })
        .is_err()
    );
    assert_eq!(edit.objects().unwrap().unwrap(), &expected);
    let commit = edit.commit().unwrap();
    let output = tempfile::NamedTempFile::with_suffix(".xlsx").unwrap();
    commit.package().save(output.path()).unwrap();
    let after = std::fs::read(output.path()).unwrap();
    let reopened = Package::from_bytes(&after).unwrap();
    assert_eq!(objects(&reopened).objects().unwrap(), &expected);
    let original = members(&bytes);
    let changed = members(&after);
    assert_eq!(
        original.keys().collect::<Vec<_>>(),
        changed.keys().collect::<Vec<_>>()
    );
    let differences = original
        .iter()
        .filter(|(name, data)| changed[*name] != **data)
        .map(|(name, _)| name.as_str())
        .collect::<Vec<_>>();
    assert_eq!(differences, ["xl/worksheets/sheet1.xml"]);
    let worksheet = std::str::from_utf8(&changed["xl/worksheets/sheet1.xml"]).unwrap();
    let attribute = " altText=\"Embedded form reviewed\"";
    assert_eq!(worksheet.matches(attribute).count(), 1);
    assert!(
        worksheet.replacen(attribute, "", 1).as_bytes() == original["xl/worksheets/sheet1.xml"]
    );

    let inverse = reopened
        .apply_ole_objects(&commit.patch().inverse())
        .unwrap();
    let restored = members(&inverse.package().to_bytes().unwrap());
    assert_eq!(
        original.keys().collect::<Vec<_>>(),
        restored.keys().collect::<Vec<_>>()
    );
    for (name, data) in original {
        assert!(restored[&name] == data, "inverse changed {name}");
    }
    assert!(p.to_bytes().unwrap() == bytes);
}

fn members(bytes: &[u8]) -> std::collections::BTreeMap<String, Vec<u8>> {
    let reader = litchi_opc::phys_pkg::OwnedPhysPkgReader::from_bytes(
        bytes.to_vec(),
        litchi_opc::SourceLimits::default(),
    )
    .unwrap();
    reader
        .member_names()
        .unwrap()
        .into_iter()
        .map(|name| {
            let bytes = reader.read_member(&name).unwrap();
            (name, bytes)
        })
        .collect()
}
