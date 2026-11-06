//! Native Excel table filters and saved sort metadata remain inert.

use litchi_core::sheet::{Cell as _, CellValue};
use litchi_xlsb::{
    Package, SourceBackedWorkbook,
    builder::{Builder, Worksheet},
    package::table,
    raw::{Limits, Records},
};
use std::{collections::BTreeMap, path::PathBuf};

fn native(name: &str) -> Vec<u8> {
    std::fs::read(
        PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/microsoft-office/excel")
            .join(name),
    )
    .unwrap()
}

fn members(bytes: &[u8]) -> BTreeMap<String, Vec<u8>> {
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

fn filter_records(bytes: &[u8]) -> Vec<(u16, Vec<u8>)> {
    Records::new(bytes, Limits::DEFAULT)
        .map(|r| r.unwrap())
        .filter(|r| matches!(r.kind().get(), 0xA1..=0xAF | 530..=533))
        .map(|r| (r.kind().get(), r.payload().to_vec()))
        .collect()
}

#[test]
fn native_table_model_publication_retains_filter_and_sort_records() {
    for name in [
        "excel-pivot-controls.xlsb",
        "excel-visual-analytics.xlsb",
        "excel-values-formulas-chart.xlsb",
        "excel-connections.xlsb",
    ] {
        let source = members(&native(name));
        let data = &source["xl/tables/table1.bin"];
        let table = table::parse_table_part(data, Limits::DEFAULT).unwrap();
        let expected_filter = table.auto_filter.clone();
        let expected_sort = table.sort_state.clone();
        let mut builder = Builder::new();
        let mut sheet = Worksheet::new("Native table metadata");
        sheet.add_table(table).unwrap();
        builder.add_worksheet(sheet);
        let after = builder.build().unwrap().to_bytes().unwrap();
        let output = members(&after);
        assert_eq!(
            filter_records(&output["xl/tables/table1.bin"]),
            filter_records(data),
            "{name}"
        );
        let parsed =
            table::parse_table_part(&output["xl/tables/table1.bin"], Limits::DEFAULT).unwrap();
        assert_eq!(parsed.auto_filter, expected_filter);
        assert_eq!(parsed.sort_state, expected_sort);
    }
}

fn assert_filtered_table(table: &table::Table) {
    use table::{Range, SortOn};
    assert_eq!(table.display_name.as_deref(), Some("Table1"));
    assert_eq!(table.columns[1].caption.as_deref(), Some("Region"));
    assert_eq!(table.columns[3].caption.as_deref(), Some("Revenue"));
    let filter = table.auto_filter.as_ref().unwrap();
    assert_eq!(
        filter.range(),
        Range {
            first_row: 0,
            last_row: 8,
            first_column: 0,
            last_column: 5
        }
    );
    assert!(!filter.has_unmodeled_records());
    assert_eq!(filter.columns().len(), 1);
    let column = &filter.columns()[0];
    assert_eq!(column.index(), 1);
    assert!(!column.hide_button() && !column.button_in_next_column());
    let values = column.values().unwrap();
    assert!(!values.include_blank());
    assert_eq!(values.values(), ["North"]);
    let sort = table.sort_state.as_ref().unwrap();
    assert!(!sort.by_columns() && !sort.case_sensitive() && !sort.alternate_method());
    assert!(!sort.has_unmodeled_records());
    assert_eq!(
        sort.range(),
        Range {
            first_row: 1,
            last_row: 7,
            first_column: 0,
            last_column: 5
        }
    );
    assert_eq!(sort.conditions().len(), 1);
    let condition = &sort.conditions()[0];
    assert!(condition.descending());
    assert_eq!(condition.sort_on(), SortOn::Value);
    assert_eq!(condition.custom_list(), None);
    // Excel saved a broader condition range than the state range. Keep the
    // literal metadata; this test does not infer an effective sort region.
    assert_eq!(
        condition.range(),
        Range {
            first_row: 0,
            last_row: 8,
            first_column: 3,
            last_column: 3
        }
    );
}

#[test]
fn native_table_filter_and_sort_owners_agree_across_eager_and_source_readers() {
    for name in ["excel-pivot-controls.xlsb", "excel-visual-analytics.xlsb"] {
        let bytes = native(name);
        let package = Package::from_bytes(&bytes).unwrap();
        let workbook = package.workbook().unwrap();
        assert_eq!(workbook.structured_tables().len(), 1);
        let (owner, table) = &workbook.structured_tables()[0];
        assert_eq!(*owner, 1);
        assert_filtered_table(table);
        // The filter belongs to the table, not to the worksheet range API.
        let catalog_position = if name == "excel-pivot-controls.xlsb" {
            2
        } else {
            1
        };
        assert!(
            package
                .edit_auto_filter()
                .unwrap()
                .auto_filter(catalog_position)
                .unwrap()
                .is_none()
        );
        let source = SourceBackedWorkbook::from_reader(std::io::Cursor::new(&bytes)).unwrap();
        let tables = source.structured_tables().unwrap();
        assert_eq!(tables.as_ref(), workbook.structured_tables());
    }
}

#[test]
fn native_cell_edit_preserves_table_filter_sort_and_unrelated_package_members() {
    for name in ["excel-pivot-controls.xlsb", "excel-visual-analytics.xlsb"] {
        let bytes = native(name);
        let package = Package::from_bytes(&bytes).unwrap();
        let mut edit = package.edit_cell_values("Sheet1").unwrap();
        edit.set_scalar(
            litchi_xlsb::cell_values::Reference::new(1, 3).unwrap(),
            litchi_xlsb::cell_values::Scalar::Number(321.5),
        )
        .unwrap();
        let (updated, patch) = edit.commit().unwrap().into_parts();
        let after = updated.to_bytes().unwrap();
        let before_members = members(&bytes);
        let after_members = members(&after);
        assert_eq!(
            before_members.keys().collect::<Vec<_>>(),
            after_members.keys().collect::<Vec<_>>()
        );
        let changed = before_members
            .iter()
            .filter(|(name, data)| **data != after_members[*name])
            .map(|(name, _)| name.as_str())
            .collect::<Vec<_>>();
        assert_eq!(
            changed,
            ["xl/worksheets/binaryIndex2.bin", "xl/worksheets/sheet2.bin"]
        );
        let output = tempfile::NamedTempFile::with_suffix(".xlsb").unwrap();
        updated.save(output.path()).unwrap();
        let reopened = Package::open(output.path()).unwrap();
        let workbook = reopened.workbook().unwrap();
        assert_filtered_table(&workbook.structured_tables()[0].1);
        assert_eq!(
            workbook
                .worksheet(1)
                .unwrap()
                .get_cell(1, 3)
                .unwrap()
                .value(),
            &CellValue::Float(321.5)
        );
        let source = SourceBackedWorkbook::from_reader(std::io::Cursor::new(&after)).unwrap();
        let indexed = source
            .sheet("Sheet1")
            .unwrap()
            .unwrap()
            .indexed_values()
            .unwrap()
            .unwrap();
        assert_eq!(
            indexed.cached_value(1, 3).unwrap(),
            Some(CellValue::Float(321.5))
        );
        let restored = reopened
            .apply_cell_values(&patch.inverse())
            .unwrap()
            .into_package();
        assert!(restored.to_bytes().unwrap() == bytes);
        assert!(package.to_bytes().unwrap() == bytes);
    }
}

#[test]
fn native_table_range_change_requires_explicit_filter_sort_rebinding() {
    let data = members(&native("excel-pivot-controls.xlsb"));
    for keep_filter in [true, false] {
        let mut table =
            table::parse_table_part(&data["xl/tables/table1.bin"], Limits::DEFAULT).unwrap();
        if !keep_filter {
            table.auto_filter = None;
        }
        table.range.first_row += 1;
        table.range.last_row += 1;
        let mut builder = Builder::new();
        let mut sheet = Worksheet::new("Moved table");
        sheet.add_table(table).unwrap();
        builder.add_worksheet(sheet);
        assert!(matches!(
            builder.build(),
            Err(litchi_xlsb::package::PackageError::Unrecognized {
                typ: "table filter/sort metadata",
                ..
            })
        ));
    }
}
