//! Native Excel embedded package discovery stays inert.
use litchi_ooxml_common::embedded::{Kind, SourceTarget, Target};
use litchi_xlsb::{Package, SourceBackedWorkbook};
use sha2::{Digest, Sha256};
fn native(name: &str) -> Vec<u8> {
    std::fs::read(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/microsoft-office")
            .join(name),
    )
    .unwrap()
}
#[test]
fn native_embedded_word_catalog_agrees_across_eager_and_source_readers() {
    let bytes = native("excel/excel-links-objects.xlsb");
    let p = Package::from_bytes(&bytes).unwrap();
    let wb = p.workbook().unwrap();
    let entries = wb.embedded().unwrap();
    assert_eq!(entries.len(), 1);
    let e = &entries[0];
    assert_eq!(e.source().as_str(), "/xl/worksheets/sheet1.bin");
    assert_eq!(e.id(), "rId4");
    assert_eq!(e.kind(), Kind::Package);
    let Target::Internal(payload) = e.target() else {
        panic!("native internal Word payload")
    };
    assert_eq!(
        payload.part().as_str(),
        "/xl/embeddings/Microsoft_Word_Document.docx"
    );
    assert_eq!(payload.bytes().len(), 13759);
    assert_eq!(
        Sha256::digest(payload.bytes())
            .iter()
            .map(|b| format!("{b:02x}"))
            .collect::<String>(),
        "a4b37e93f7cf0c93213e419cb42796e361785f4354d50a585dcd5b60ed910778"
    );
    assert!(payload.bytes() == native("word/word-legacy-form-unprotected.docx"));
    let source = SourceBackedWorkbook::from_reader(std::io::Cursor::new(&bytes)).unwrap();
    let indexed = source.embedded().unwrap();
    assert_eq!(indexed.len(), 1);
    assert_eq!(indexed[0].source(), e.source());
    assert_eq!(indexed[0].id(), e.id());
    assert_eq!(indexed[0].kind(), e.kind());
    let SourceTarget::Internal(lazy) = indexed[0].target() else {
        panic!("native internal Word payload")
    };
    assert_eq!(lazy.part(), payload.part());
    assert_eq!(lazy.content_type(), payload.content_type());
    assert_eq!(lazy.declared_uncompressed_size().unwrap(), 13759);
    assert!(lazy.data().unwrap().as_bytes() == payload.bytes());
    let mut streamed = Vec::new();
    assert_eq!(lazy.stream_to(&mut streamed).unwrap(), 13759);
    assert_eq!(streamed, payload.bytes());
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

#[test]
fn native_word_payload_replacement_preserves_every_other_package_member() {
    let bytes = native("excel/excel-links-objects.xlsb");
    let p = Package::from_bytes(&bytes).unwrap();
    let replacement = native("word/word-text-structure.docx");
    let name = litchi_opc::PackURI::new("/xl/embeddings/Microsoft_Word_Document.docx").unwrap();
    let mut edit = p.edit_opc().unwrap();
    edit.update(|opc| {
        opc.get_part_mut(&name)?.set_blob(replacement.clone());
        Ok(())
    })
    .unwrap();
    let commit = edit.commit().unwrap();
    let output = tempfile::NamedTempFile::with_suffix(".xlsb").unwrap();
    commit.package().save(output.path()).unwrap();
    let after = std::fs::read(output.path()).unwrap();
    let reopened = Package::from_bytes(&after).unwrap();
    let wb = reopened.workbook().unwrap();
    let entries = wb.embedded().unwrap();
    assert_eq!(entries.len(), 1);
    assert_eq!(entries[0].source().as_str(), "/xl/worksheets/sheet1.bin");
    assert_eq!(entries[0].id(), "rId4");
    let Target::Internal(payload) = entries[0].target() else {
        panic!("internal Word package")
    };
    assert!(payload.bytes() == replacement);
    let before = members(&bytes);
    let after = members(&after);
    assert_eq!(
        before.keys().collect::<Vec<_>>(),
        after.keys().collect::<Vec<_>>()
    );
    let differences = before
        .iter()
        .filter(|(name, data)| after[*name] != **data)
        .map(|(name, _)| name.as_str())
        .collect::<Vec<_>>();
    assert_eq!(differences, ["xl/embeddings/Microsoft_Word_Document.docx"]);
    let restored = reopened
        .apply_opc(&commit.patch().inverse())
        .unwrap()
        .into_package();
    let restored_members = members(&restored.to_bytes().unwrap());
    assert_eq!(
        before.keys().collect::<Vec<_>>(),
        restored_members.keys().collect::<Vec<_>>()
    );
    let inverse_differences = before
        .iter()
        .filter(|(name, data)| restored_members[*name] != **data)
        .map(|(name, _)| name.as_str())
        .collect::<Vec<_>>();
    assert!(
        inverse_differences.is_empty(),
        "inverse changed {inverse_differences:?}"
    );
    assert!(p.to_bytes().unwrap() == bytes);
}
