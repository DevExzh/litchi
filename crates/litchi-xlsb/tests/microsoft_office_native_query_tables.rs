//! Native query-table bindings and saved imported values; no refresh or I/O.
use litchi_core::sheet::{Cell as _, CellValue, WorkbookTrait};
use litchi_xlsb::{Package, SourceBackedWorkbook, package::table};
use sha2::{Digest, Sha256};
fn native() -> Vec<u8> {
    std::fs::read(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/microsoft-office/excel/excel-connections.xlsb"),
    )
    .unwrap()
}
fn members(bytes: &[u8]) -> std::collections::BTreeMap<String, Vec<u8>> {
    let p = litchi_opc::phys_pkg::OwnedPhysPkgReader::from_bytes(
        bytes.to_vec(),
        litchi_opc::SourceLimits::default(),
    )
    .unwrap();
    p.member_names()
        .unwrap()
        .into_iter()
        .map(|n| {
            let b = p.read_member(&n).unwrap();
            (n, b)
        })
        .collect()
}
fn check_rows(p: &Package) {
    let wb = p.workbook().unwrap();
    assert_eq!(wb.worksheet_names(), ["native-query-010", "Sheet1"]);
    for (sheet, first) in [(0, 1), (1, 8)] {
        let data = wb.worksheet(sheet).unwrap();
        for (i, (category, amount)) in [("ALPHA_010", 12.5), ("BETA_010", 20.), ("ALPHA_010", 7.5)]
            .into_iter()
            .enumerate()
        {
            let row = first + i as u32;
            assert_eq!(
                data.get_cell(row, 0).unwrap().value(),
                &CellValue::Int((i + 1) as i64)
            );
            assert_eq!(
                data.get_cell(row, 1).unwrap().value(),
                &CellValue::String(category.into())
            );
            assert_eq!(
                data.get_cell(row, 2).unwrap().value(),
                &if i == 1 {
                    CellValue::Int(20)
                } else {
                    CellValue::Float(amount)
                }
            );
        }
    }
}
#[test]
fn native_query_table_metadata_and_relationship_owners_match_saved_imports() {
    let bytes = native();
    let p = Package::from_bytes(&bytes).unwrap();
    check_rows(&p);
    let wb = p.workbook().unwrap();
    assert_eq!(wb.structured_tables().len(), 1);
    let (owner, t) = &wb.structured_tables()[0];
    assert_eq!(*owner, 0);
    assert_eq!(t.id, 1);
    assert_eq!(t.name.as_deref(), Some("native_query_010"));
    assert_eq!(t.display_name, t.name);
    assert_eq!(t.table_type, table::Type::QueryTable);
    assert_eq!(
        t.range,
        table::Range {
            first_row: 0,
            last_row: 3,
            first_column: 0,
            last_column: 2
        }
    );
    assert_eq!((t.header_row_count, t.totals_row_count), (1, 0));
    // This ListObject leaves dwConnID zero. The QSI connection binding remains
    // in the opaque query-table part, not in this typed table field.
    assert_eq!(t.connection_id, None);
    assert_eq!(
        t.columns
            .iter()
            .map(|c| (c.id, c.caption.as_deref(), c.query_table_field_id))
            .collect::<Vec<_>>(),
        [
            (1, Some("ID"), 1),
            (2, Some("Category"), 2),
            (3, Some("Amount"), 3)
        ]
    );
    let source = SourceBackedWorkbook::from_reader(std::io::Cursor::new(&bytes)).unwrap();
    assert_eq!(
        source.structured_tables().unwrap().as_ref(),
        wb.structured_tables()
    );
    let opc = p.opc_package();
    for (owner, id, target, len, hash) in [
        (
            "/xl/tables/table1.bin",
            "rId1",
            "/xl/queryTables/queryTable1.bin",
            201,
            "4b4c7075bdc13467abde3fe7f4d067593051027ebc0615d7e76e35daa5c65ce6",
        ),
        (
            "/xl/worksheets/sheet2.bin",
            "rId2",
            "/xl/queryTables/queryTable2.bin",
            77,
            "f0eb8bc05607e979f4b3cbe3ed09cbe587ad205cd0d5ff6edad3aa0578c81ea5",
        ),
    ] {
        let owner = opc
            .get_part(&litchi_opc::PackURI::new(owner).unwrap())
            .unwrap();
        let rel = owner.rels().get(id).unwrap();
        assert!(!rel.is_external());
        assert_eq!(
            rel.reltype(),
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/queryTable"
        );
        assert_eq!(rel.target_partname().unwrap().as_str(), target);
        let part = opc.get_part(&rel.target_partname().unwrap()).unwrap();
        assert_eq!(part.content_type(), "application/vnd.ms-excel.queryTable");
        assert_eq!(part.blob().len(), len);
        assert_eq!(
            Sha256::digest(part.blob())
                .iter()
                .map(|b| format!("{b:02x}"))
                .collect::<String>(),
            hash
        );
    }
    assert!(p.to_bytes().unwrap() == bytes);
}
#[test]
fn native_ordinary_cell_edit_keeps_query_parts_bindings_and_import_caches() {
    let bytes = native();
    let p = Package::from_bytes(&bytes).unwrap();
    let before = members(&bytes);
    let tables = p.workbook().unwrap().structured_tables().to_vec();
    let mut edit = p.edit_cell_values("Sheet1").unwrap();
    edit.set_scalar(
        litchi_xlsb::cell_values::Reference::new(2, 1).unwrap(),
        litchi_xlsb::cell_values::Scalar::Number(400.),
    )
    .unwrap();
    let commit = edit.commit().unwrap();
    let output = tempfile::NamedTempFile::with_suffix(".xlsb").unwrap();
    commit.package().save(output.path()).unwrap();
    let after_bytes = std::fs::read(output.path()).unwrap();
    let after = members(&after_bytes);
    let reopened = Package::from_bytes(&after_bytes).unwrap();
    check_rows(&reopened);
    assert_eq!(reopened.workbook().unwrap().structured_tables(), tables);
    assert_eq!(
        reopened
            .workbook()
            .unwrap()
            .worksheet(1)
            .unwrap()
            .get_cell(2, 1)
            .unwrap()
            .value(),
        &CellValue::Float(400.)
    );
    assert_eq!(
        before.keys().collect::<Vec<_>>(),
        after.keys().collect::<Vec<_>>()
    );
    let changed = before
        .iter()
        .filter(|(n, b)| after[*n] != **b)
        .map(|(n, _)| n.as_str())
        .collect::<Vec<_>>();
    assert_eq!(
        changed,
        ["xl/worksheets/binaryIndex2.bin", "xl/worksheets/sheet2.bin"]
    );
    let restored = reopened
        .apply_cell_values(&commit.patch().inverse())
        .unwrap();
    assert!(restored.package().to_bytes().unwrap() == bytes);
    assert!(p.to_bytes().unwrap() == bytes);
}
