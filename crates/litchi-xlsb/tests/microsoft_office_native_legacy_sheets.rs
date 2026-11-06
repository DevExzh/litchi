//! Native macro-sheet metadata stays inert during ordinary worksheet edits.
fn native() -> Vec<u8> {
    std::fs::read(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/microsoft-office/excel/excel-legacy-sheets.xlsb"),
    )
    .unwrap()
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

use litchi_xlsb::{Package, SourceBackedWorkbook, package::PackageError, sheet::Kind};
#[test]
fn native_macro_catalog_and_index_presence_have_explicit_typed_boundaries() {
    let bytes = native();
    let p = Package::from_bytes(&bytes).unwrap();
    let wb = p.workbook().unwrap();
    let source = SourceBackedWorkbook::from_reader(std::io::Cursor::new(&bytes)).unwrap();
    assert_eq!(source.sheet_names().unwrap(), ["Macro_016", "Sheet1"]);
    assert_eq!(source.worksheet_names().unwrap(), ["Sheet1"]);
    assert_eq!(
        wb.sheets()
            .map(|s| s.name().to_string())
            .collect::<Vec<_>>(),
        source.sheet_names().unwrap()
    );
    assert_eq!(wb.sheet("Macro_016").unwrap().unwrap().kind(), Kind::Macro);
    assert_eq!(wb.active_catalog_position(), Some(0));
    assert_eq!(wb.active_worksheet_index(), None);
    assert_eq!(source.active_catalog_position().unwrap(), Some(0));
    assert_eq!(source.active_worksheet_index().unwrap(), None);
    let m = source.sheet("Macro_016").unwrap().unwrap();
    assert_eq!(m.kind().unwrap(), Kind::Macro);
    assert_eq!(m.workbook_position().unwrap(), 0);
    assert!(source.worksheet_by_name("Macro_016").unwrap().is_none());
    assert!(matches!(
        m.materialize(),
        Err(PackageError::UnsupportedFeature(
            "source-backed XLSB non-worksheet materialization is not supported"
        ))
    ));
    assert!(matches!(
        m.indexed_values(),
        Err(PackageError::UnsupportedFeature(
            "binary-index value lookup requires an XLSB worksheet"
        ))
    ));
    let parts = members(&bytes);
    assert_eq!(parts["xl/macrosheets/sheet1.bin"].len(), 490);
    assert_eq!(parts["xl/macrosheets/binaryIndex1.bin"].len(), 61);
    assert_eq!(
        hash(&parts["xl/macrosheets/sheet1.bin"]),
        "0b5c619c585d7a34f7c88604b7c1f4b32cf6169633a387d9e13f8473c08c8dea"
    );
    assert_eq!(
        hash(&parts["xl/macrosheets/binaryIndex1.bin"]),
        "1217929ac6f593077f93a1a5126c2fa9ddbacf651ed73b4a8f12da87c730dd23"
    );
    let index = litchi_xlsb::raw::Records::new(
        &parts["xl/macrosheets/binaryIndex1.bin"],
        litchi_xlsb::raw::Limits::DEFAULT,
    )
    .map(|r| {
        let r = r.unwrap();
        (r.kind().get(), r.payload().len())
    })
    .collect::<Vec<_>>();
    assert_eq!(index, [(0x2a, 24), (0x28, 30), (0x115, 0)]);
    assert!(!parts.keys().any(|name| name.contains("vbaProject")));
    assert!(p.to_bytes().unwrap() == bytes);
}
#[test]
fn native_ordinary_worksheet_edit_preserves_macro_body_and_binary_index() {
    let bytes = native();
    let p = Package::from_bytes(&bytes).unwrap();
    let mut edit = p.edit_cell_values("Sheet1").unwrap();
    edit.set_scalar(
        litchi_xlsb::cell_values::Reference::new(3, 3).unwrap(),
        litchi_xlsb::cell_values::Scalar::Text("Reviewed ordinary worksheet".into()),
    )
    .unwrap();
    let commit = edit.commit().unwrap();
    let output = tempfile::NamedTempFile::with_suffix(".xlsb").unwrap();
    commit.package().save(output.path()).unwrap();
    let after = std::fs::read(output.path()).unwrap();
    let reopened = Package::from_bytes(&after).unwrap();
    let original = members(&bytes);
    let changed = members(&after);
    assert_eq!(
        original.keys().collect::<Vec<_>>(),
        changed.keys().collect::<Vec<_>>()
    );
    let diffs = original
        .iter()
        .filter(|(n, b)| changed[*n] != **b)
        .map(|(n, _)| n.as_str())
        .collect::<Vec<_>>();
    assert_eq!(
        diffs,
        ["xl/worksheets/binaryIndex1.bin", "xl/worksheets/sheet1.bin"]
    );
    for (name, part) in &original {
        if name.starts_with("xl/macrosheets/") {
            assert!(changed[name] == *part, "changed {name}");
        }
    }
    let source = SourceBackedWorkbook::from_reader(std::io::Cursor::new(&after)).unwrap();
    assert_eq!(
        source.sheet("Macro_016").unwrap().unwrap().kind().unwrap(),
        Kind::Macro
    );
    assert_eq!(
        source
            .worksheet_by_name("Sheet1")
            .unwrap()
            .unwrap()
            .indexed_values()
            .unwrap()
            .unwrap()
            .cached_value(3, 3)
            .unwrap(),
        Some(litchi_core::sheet::CellValue::String(
            "Reviewed ordinary worksheet".into()
        ))
    );
    let restored = reopened
        .apply_cell_values(&commit.patch().inverse())
        .unwrap();
    assert!(restored.package().to_bytes().unwrap() == bytes);
}

fn hash(bytes: &[u8]) -> String {
    use sha2::{Digest, Sha256};
    Sha256::digest(bytes)
        .iter()
        .map(|b| format!("{b:02x}"))
        .collect()
}
