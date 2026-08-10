//! Failure-atomic planning for fresh XLSB XML Maps authoring.

use crate::package::error::{Error, Result};
use crate::writer::MutableWorksheet;
use std::collections::HashSet;

const MAX_ROW: u32 = 1_048_575;
const MAX_COLUMN: u32 = 16_383;

pub(super) struct XmlMapsWritePlan {
    pub(super) map_info_xml: Option<Vec<u8>>,
    pub(super) worksheets: Vec<WorksheetXmlMapsWritePlan>,
}

pub(super) struct WorksheetXmlMapsWritePlan {
    pub(super) table_parts: Vec<Vec<u8>>,
    pub(super) single_cells: Option<Vec<u8>>,
}

fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidFormula(format!("Custom XML Maps: {}", message.into()))
}

pub(super) fn stage(
    catalog: Option<&crate::xml_maps::XmlMapInfo>,
    connections: Option<&crate::package::connections::Connections>,
    worksheets: &[MutableWorksheet],
) -> Result<XmlMapsWritePlan> {
    let has_bindings = worksheets
        .iter()
        .any(|sheet| !sheet.mapped_tables.is_empty() || !sheet.single_cell_mappings.is_empty());
    if has_bindings && catalog.is_none() {
        return Err(invalid("worksheet bindings require a MapInfo catalog"));
    }

    let map_info_xml = catalog
        .map(|value| -> Result<Vec<u8>> {
            crate::xml_maps::validate_catalog(value, crate::xml_maps::XmlMapLimits::DEFAULT)?;
            Ok(crate::xml_maps::serialize_xml_map_info(
                value,
                crate::xml_maps::XmlMapConformance::Transitional,
            )?)
        })
        .transpose()?;
    let map_ids = catalog
        .map(|value| value.maps.iter().map(|map| map.id).collect::<HashSet<_>>())
        .unwrap_or_default();
    if let Some(catalog) = catalog {
        for map in &catalog.maps {
            if let Some(connection_id) = map
                .data_binding
                .as_ref()
                .and_then(|binding| binding.connection_id)
                && connections.is_none_or(|values| values.by_id(connection_id).is_none())
            {
                return Err(invalid(format!(
                    "map ID {} references unknown inert connection ID {connection_id}",
                    map.id
                )));
            }
        }
    }
    let mut list_ids = HashSet::new();
    for sheet in worksheets {
        for (table_index, table) in sheet.tables.iter().enumerate() {
            if !(1..=0xFFFF_FFFE).contains(&table.id) || !list_ids.insert(table.id) {
                return Err(invalid(format!(
                    "structured table ID {} is zero, reserved, or duplicated",
                    table.id
                )));
            }
            validate_range(&table.range)?;
            if table.header_row_count > 1 || table.totals_row_count > 1 {
                return Err(invalid(format!(
                    "structured table {} header and totals row counts must be Boolean",
                    table.id
                )));
            }
            let connection_id = table.connection_id.unwrap_or(0);
            if connection_id != 0 {
                if table.table_type != crate::package::table::Type::Xml {
                    return Err(invalid(format!(
                        "non-XML structured table {} cannot reference connection ID {connection_id}",
                        table.id
                    )));
                }
                if connections.is_none_or(|values| values.by_id(connection_id).is_none()) {
                    return Err(invalid(format!(
                        "XML structured table {} references unknown inert connection ID {connection_id}",
                        table.id
                    )));
                }
            }
            let height = table.range.last_row - table.range.first_row + 1;
            let reserved_rows = table
                .header_row_count
                .checked_add(table.totals_row_count)
                .ok_or_else(|| invalid("structured table row-count sum overflows"))?;
            if height <= reserved_rows {
                return Err(invalid(format!(
                    "structured table {} range height must exceed header plus totals rows",
                    table.id
                )));
            }
            if sheet.xml_table_overlaps_auto_filter(&table.range) {
                return Err(invalid(format!(
                    "structured table {} overlaps the worksheet AutoFilter",
                    table.id
                )));
            }
            if let Some(other) = sheet.tables[..table_index]
                .iter()
                .find(|other| ranges_overlap(&table.range, &other.range))
            {
                return Err(invalid(format!(
                    "structured tables {} and {} overlap",
                    other.id, table.id
                )));
            }
        }
    }

    let limits = crate::xml_maps::Limits::DEFAULT;
    let mut plans = Vec::new();
    plans
        .try_reserve_exact(worksheets.len())
        .map_err(|source| Error::Allocation {
            resource: "XML Maps worksheet plans",
            source,
        })?;
    for sheet in worksheets {
        if let Some(catalog) = catalog {
            crate::xml_maps::validate_binding_map_ids(
                catalog,
                &sheet.mapped_tables,
                &sheet.single_cell_mappings,
            )?;
        }
        let mut mapped_ids = HashSet::new();
        let mut table_parts = Vec::new();
        table_parts
            .try_reserve_exact(sheet.tables.len())
            .map_err(|source| Error::Allocation {
                resource: "XML-mapped table parts",
                source,
            })?;
        for table in &sheet.tables {
            let base = crate::package::table::write::write_table_part(table)?;
            let mapping = sheet
                .mapped_tables
                .iter()
                .find(|mapping| mapping.table_id() == table.id);
            if let Some(mapping) = mapping {
                mapped_ids.insert(table.id);
                if table.table_type != crate::package::table::Type::Xml || table.single_cell {
                    return Err(invalid(format!(
                        "mapped table {} must be a normal LTXML table",
                        table.id
                    )));
                }
                validate_column_dependencies(table, mapping, &map_ids)?;
                table_parts.push(crate::xml_maps::apply_table_bindings(
                    &base, mapping, limits,
                )?);
            } else {
                table_parts.push(base);
            }
        }
        if mapped_ids.len() != sheet.mapped_tables.len() {
            let missing = sheet
                .mapped_tables
                .iter()
                .find(|mapping| !mapped_ids.contains(&mapping.table_id()))
                .map(|mapping| mapping.table_id())
                .unwrap_or(0);
            return Err(invalid(format!(
                "mapped table ID {missing} does not exist on worksheet {:?}",
                sheet.name()
            )));
        }

        let mut cells = HashSet::new();
        for binding in &sheet.single_cell_mappings {
            if !(1..=0xFFFF_FFFE).contains(&binding.table_id())
                || !list_ids.insert(binding.table_id())
            {
                return Err(invalid(format!(
                    "single-cell list ID {} is reserved or collides with another workbook table ID",
                    binding.table_id()
                )));
            }
            let cell = binding.cell();
            if cell.row() > MAX_ROW || cell.column() > MAX_COLUMN {
                return Err(invalid("single-cell mapping is outside worksheet bounds"));
            }
            if !cells.insert(cell) {
                return Err(invalid(
                    "multiple XML mappings target the same worksheet cell",
                ));
            }
            if sheet.tables.iter().any(|table| {
                cell.row() >= table.range.first_row
                    && cell.row() <= table.range.last_row
                    && cell.column() >= table.range.first_column
                    && cell.column() <= table.range.last_column
            }) {
                return Err(invalid(format!(
                    "single-cell mapping {} overlaps a structured table",
                    binding.table_id()
                )));
            }
            if sheet.xml_mapping_overlaps_auto_filter(cell.row(), cell.column()) {
                return Err(invalid(format!(
                    "single-cell mapping {} overlaps the worksheet AutoFilter",
                    binding.table_id()
                )));
            }
        }
        let single_cells = if sheet.single_cell_mappings.is_empty() {
            None
        } else {
            Some(crate::xml_maps::serialize_single_cells(
                &sheet.single_cell_mappings,
                limits,
            )?)
        };
        plans.push(WorksheetXmlMapsWritePlan {
            table_parts,
            single_cells,
        });
    }
    Ok(XmlMapsWritePlan {
        map_info_xml,
        worksheets: plans,
    })
}

fn validate_range(range: &crate::package::table::Range) -> Result<()> {
    if range.first_row > range.last_row
        || range.first_column > range.last_column
        || range.last_row > MAX_ROW
        || range.last_column > MAX_COLUMN
    {
        return Err(invalid("structured table has an invalid worksheet range"));
    }
    Ok(())
}

fn ranges_overlap(
    left: &crate::package::table::Range,
    right: &crate::package::table::Range,
) -> bool {
    left.first_row <= right.last_row
        && left.last_row >= right.first_row
        && left.first_column <= right.last_column
        && left.last_column >= right.first_column
}

fn validate_column_dependencies(
    table: &crate::package::table::Table,
    mapping: &crate::xml_maps::MappedTable,
    map_ids: &HashSet<u32>,
) -> Result<()> {
    for binding in mapping.columns() {
        if !table
            .columns
            .iter()
            .any(|column| column.id == binding.column_id())
        {
            return Err(invalid(format!(
                "mapped table {} references unknown column ID {}",
                table.id,
                binding.column_id()
            )));
        }
        if !map_ids.contains(&binding.map_id()) {
            return Err(invalid(format!(
                "mapped table {} column {} references unknown map ID {}",
                table.id,
                binding.column_id(),
                binding.map_id()
            )));
        }
    }
    Ok(())
}

#[cfg(test)]
#[allow(
    clippy::cast_possible_truncation,
    reason = "XML-map fixtures derive bounded wire counts from short literal vectors"
)]
mod tests {
    use super::*;
    use crate::package::table::{Column, Range, Table, Type};
    use crate::writer::WorkbookWriter;
    use litchi_opc::constants::relationship_type as rel;
    use litchi_opc::{OpcPackage, PackURI};
    use std::io::Cursor;

    fn catalog() -> crate::xml_maps::XmlMapInfo {
        crate::xml_maps::XmlMapInfo {
            selection_namespaces: "xmlns:r='urn:test'".to_string(),
            schemas: vec![crate::xml_maps::XmlSchema {
                id: "schema-1".to_string(),
                schema_reference: None,
                namespace: Some("urn:test".to_string()),
                payload_xml: None,
            }],
            maps: vec![crate::xml_maps::XmlMap {
                id: 7,
                name: "Values".to_string(),
                root_element: "root".to_string(),
                schema_id: "schema-1".to_string(),
                show_import_export_validation_errors: false,
                auto_fit: false,
                append: false,
                preserve_sort_auto_filter_layout: true,
                preserve_format: true,
                data_binding: None,
            }],
        }
    }

    #[test]
    fn writer_reopens_map_info_table_binding_and_single_cell_part() {
        let datatype = crate::xml_maps::XmlDataType::new(1).unwrap();
        let table_binding = crate::xml_maps::ColumnBinding::new(
            1,
            7,
            datatype,
            crate::xml_maps::XPath::new("/r/items/item").unwrap(),
            false,
        )
        .unwrap();
        let mapped_table = crate::xml_maps::MappedTable::new(2, vec![table_binding]).unwrap();
        let single = crate::xml_maps::SingleCellBinding::new(
            3,
            1,
            crate::xml_maps::CellReference::new(4, 1).unwrap(),
            7,
            datatype,
            crate::xml_maps::XPath::new("/r/value").unwrap(),
        )
        .unwrap();

        let mut sheet = MutableWorksheet::new("Mapped");
        sheet
            .add_table(Table {
                id: 2,
                display_name: Some("MappedValues".to_string()),
                range: Range {
                    first_row: 0,
                    last_row: 2,
                    first_column: 0,
                    last_column: 0,
                },
                table_type: Type::Xml,
                header_row_count: 1,
                columns: vec![Column {
                    id: 1,
                    name: Some("Value".to_string()),
                    ..Column::default()
                }],
                ..Table::default()
            })
            .unwrap();
        sheet.set_mapped_table(mapped_table.clone()).unwrap();
        sheet.set_single_cell_mapping(single.clone()).unwrap();

        let mut workbook = WorkbookWriter::new();
        workbook.set_xml_maps(catalog()).unwrap();
        workbook.add_worksheet(sheet);
        let mut bytes = Cursor::new(Vec::new());
        workbook.save(&mut bytes).unwrap();

        let package = OpcPackage::from_bytes(&bytes.into_inner()).unwrap();
        let workbook_part = package
            .get_part(&PackURI::new("/xl/workbook.bin").unwrap())
            .unwrap();
        let maps_rel = workbook_part
            .rels()
            .iter()
            .find(|relationship| {
                relationship.reltype() == litchi_ooxml_common::spreadsheet_xml_maps::REL
            })
            .unwrap();
        let maps_part = package
            .get_part(&maps_rel.target_partname().unwrap())
            .unwrap();
        assert_eq!(
            crate::xml_maps::parse_xml_map_info(maps_part.blob()).unwrap(),
            catalog()
        );

        let sheet_part = package
            .get_part(&PackURI::new("/xl/worksheets/sheet1.bin").unwrap())
            .unwrap();
        let table_rel = sheet_part
            .rels()
            .iter()
            .find(|relationship| relationship.reltype() == rel::TABLE)
            .unwrap();
        let table_part = package
            .get_part(&table_rel.target_partname().unwrap())
            .unwrap();
        assert_eq!(
            crate::xml_maps::parse_table_bindings(
                table_part.blob(),
                crate::xml_maps::Limits::DEFAULT,
            )
            .unwrap()
            .value(),
            &mapped_table
        );

        let single_rel = sheet_part
            .rels()
            .iter()
            .find(|relationship| {
                relationship.reltype()
                    == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableSingleCells"
            })
            .unwrap();
        let single_part = package
            .get_part(&single_rel.target_partname().unwrap())
            .unwrap();
        assert_eq!(
            crate::xml_maps::parse_single_cells(
                single_part.blob(),
                crate::xml_maps::Limits::DEFAULT,
            )
            .unwrap()
            .value(),
            &[single]
        );
    }

    #[test]
    fn reserved_single_cell_list_id_fails_before_worksheet_mutation() {
        let mut sheet = MutableWorksheet::new("Mapped");
        let mapping = crate::xml_maps::SingleCellBinding::new(
            u32::MAX,
            1,
            crate::xml_maps::CellReference::new(0, 0).unwrap(),
            7,
            crate::xml_maps::XmlDataType::new(1).unwrap(),
            crate::xml_maps::XPath::new("/r/value").unwrap(),
        );
        assert!(mapping.is_err());
        assert!(sheet.single_cell_mappings().is_empty());
        assert!(!sheet.clear_single_cell_mappings());
    }

    fn xml_table(id: u32, table_type: Type, column_id: u32, range: Range) -> Table {
        Table {
            id,
            display_name: Some(format!("Mapped{id}")),
            range,
            table_type,
            header_row_count: 1,
            columns: vec![Column {
                id: column_id,
                name: Some("Value".to_string()),
                ..Column::default()
            }],
            ..Table::default()
        }
    }

    fn mapped_table(table_id: u32, column_id: u32, map_id: u32) -> crate::xml_maps::MappedTable {
        crate::xml_maps::MappedTable::new(
            table_id,
            vec![
                crate::xml_maps::ColumnBinding::new(
                    column_id,
                    map_id,
                    crate::xml_maps::XmlDataType::new(1).unwrap(),
                    crate::xml_maps::XPath::new("/r/items/item").unwrap(),
                    false,
                )
                .unwrap(),
            ],
        )
        .unwrap()
    }

    fn single_cell(table_id: u32, row: u32, column: u32) -> crate::xml_maps::SingleCellBinding {
        crate::xml_maps::SingleCellBinding::new(
            table_id,
            1,
            crate::xml_maps::CellReference::new(row, column).unwrap(),
            7,
            crate::xml_maps::XmlDataType::new(1).unwrap(),
            crate::xml_maps::XPath::new("/r/value").unwrap(),
        )
        .unwrap()
    }

    fn authored(sheet: MutableWorksheet) -> WorkbookWriter {
        let mut workbook = WorkbookWriter::new();
        workbook.set_xml_maps(catalog()).unwrap();
        workbook.add_worksheet(sheet);
        workbook
    }

    fn assert_save_failure_is_atomic(
        mut workbook: WorkbookWriter,
        expected: &[(usize, usize, usize)],
    ) {
        let before_catalog = workbook.xml_maps().cloned();
        let mut output = Cursor::new(Vec::new());
        assert!(workbook.save(&mut output).is_err());
        assert!(output.get_ref().is_empty());
        assert_eq!(workbook.xml_maps(), before_catalog.as_ref());
        assert_eq!(workbook.worksheet_count(), expected.len());
        for (index, &(tables, mapped, single)) in expected.iter().enumerate() {
            let sheet = workbook.get_worksheet_mut(index).unwrap();
            assert_eq!(sheet.tables().len(), tables);
            assert_eq!(sheet.mapped_tables().len(), mapped);
            assert_eq!(sheet.single_cell_mappings().len(), single);
        }
    }

    #[test]
    fn dependency_and_range_preflight_failures_are_atomic() {
        let range = Range {
            first_row: 0,
            last_row: 2,
            first_column: 0,
            last_column: 0,
        };
        let mut cases = Vec::new();

        let mut unknown_map = MutableWorksheet::new("Unknown Map");
        unknown_map
            .add_table(xml_table(2, Type::Xml, 1, range))
            .unwrap();
        unknown_map.set_mapped_table(mapped_table(2, 1, 8)).unwrap();
        cases.push(("unknown map", authored(unknown_map), vec![(1, 1, 0)]));

        let mut missing_table = MutableWorksheet::new("Missing Table");
        missing_table
            .set_mapped_table(mapped_table(9, 1, 7))
            .unwrap();
        cases.push(("missing table", authored(missing_table), vec![(0, 1, 0)]));

        let mut non_xml = MutableWorksheet::new("Non XML");
        non_xml
            .add_table(xml_table(2, Type::Range, 1, range))
            .unwrap();
        non_xml.set_mapped_table(mapped_table(2, 1, 7)).unwrap();
        cases.push(("non-XML table", authored(non_xml), vec![(1, 1, 0)]));

        let mut missing_column = MutableWorksheet::new("Missing Column");
        missing_column
            .add_table(xml_table(2, Type::Xml, 1, range))
            .unwrap();
        missing_column
            .set_mapped_table(mapped_table(2, 2, 7))
            .unwrap();
        cases.push(("missing column", authored(missing_column), vec![(1, 1, 0)]));

        let mut duplicate = WorkbookWriter::new();
        duplicate.set_xml_maps(catalog()).unwrap();
        for name in ["Duplicate One", "Duplicate Two"] {
            let mut sheet = MutableWorksheet::new(name);
            sheet.add_table(xml_table(2, Type::Xml, 1, range)).unwrap();
            duplicate.add_worksheet(sheet);
        }
        cases.push(("duplicate list ID", duplicate, vec![(1, 0, 0), (1, 0, 0)]));

        let mut invalid_range = MutableWorksheet::new("Invalid Range");
        invalid_range.tables.push(xml_table(
            2,
            Type::Xml,
            1,
            Range {
                first_row: 3,
                last_row: 2,
                first_column: 0,
                last_column: 0,
            },
        ));
        cases.push(("invalid range", authored(invalid_range), vec![(1, 0, 0)]));

        let mut table_overlap = MutableWorksheet::new("Table Overlap");
        table_overlap
            .add_table(xml_table(2, Type::Xml, 1, range))
            .unwrap();
        table_overlap
            .set_single_cell_mapping(single_cell(3, 1, 0))
            .unwrap();
        cases.push(("table overlap", authored(table_overlap), vec![(1, 0, 1)]));

        let mut filter_overlap = MutableWorksheet::new("Filter Overlap");
        filter_overlap.set_auto_filter(0, 2, 0, 2);
        filter_overlap
            .set_single_cell_mapping(single_cell(3, 1, 1))
            .unwrap();
        cases.push((
            "AutoFilter overlap",
            authored(filter_overlap),
            vec![(0, 0, 1)],
        ));

        let mut overlapping_tables = MutableWorksheet::new("Overlapping Tables");
        overlapping_tables
            .add_table(xml_table(2, Type::Xml, 1, range))
            .unwrap();
        overlapping_tables
            .add_table(xml_table(
                4,
                Type::Xml,
                1,
                Range {
                    first_row: 2,
                    last_row: 4,
                    first_column: 0,
                    last_column: 0,
                },
            ))
            .unwrap();
        cases.push((
            "overlapping ordinary tables",
            authored(overlapping_tables),
            vec![(2, 0, 0)],
        ));

        let mut table_filter_overlap = MutableWorksheet::new("Table Filter Overlap");
        table_filter_overlap.set_auto_filter(0, 3, 0, 1);
        table_filter_overlap
            .add_table(xml_table(2, Type::Xml, 1, range))
            .unwrap();
        cases.push((
            "ordinary table AutoFilter overlap",
            authored(table_filter_overlap),
            vec![(1, 0, 0)],
        ));

        let mut no_data_rows = MutableWorksheet::new("No Data Rows");
        no_data_rows
            .add_table(xml_table(
                2,
                Type::Xml,
                1,
                Range {
                    first_row: 0,
                    last_row: 0,
                    first_column: 0,
                    last_column: 0,
                },
            ))
            .unwrap();
        cases.push((
            "no table data rows",
            authored(no_data_rows),
            vec![(1, 0, 0)],
        ));

        for (name, header_row_count, totals_row_count) in [
            ("non-Boolean header rows", 2, 0),
            ("non-Boolean totals rows", 0, 2),
        ] {
            let mut invalid_boolean = MutableWorksheet::new(name);
            let mut table = xml_table(2, Type::Xml, 1, range);
            table.header_row_count = header_row_count;
            table.totals_row_count = totals_row_count;
            invalid_boolean.add_table(table).unwrap();
            cases.push((name, authored(invalid_boolean), vec![(1, 0, 0)]));
        }

        let mut non_xml_connection = MutableWorksheet::new("Non XML Connection");
        let mut table = xml_table(2, Type::Range, 1, range);
        table.connection_id = Some(91);
        non_xml_connection.add_table(table).unwrap();
        cases.push((
            "non-XML table connection",
            authored(non_xml_connection),
            vec![(1, 0, 0)],
        ));

        let mut missing_table_connection = MutableWorksheet::new("Missing Table Connection");
        let mut table = xml_table(2, Type::Xml, 1, range);
        table.connection_id = Some(91);
        missing_table_connection.add_table(table).unwrap();
        cases.push((
            "missing XML table connection",
            authored(missing_table_connection),
            vec![(1, 0, 0)],
        ));

        let mut missing_connection_catalog = catalog();
        missing_connection_catalog.maps[0].data_binding = Some(crate::xml_maps::DataBinding {
            data_binding_name: None,
            file_binding: Some(true),
            connection_id: Some(91),
            file_binding_name: None,
            load_mode: 0,
            payload_xml: None,
        });
        let mut missing_connection = WorkbookWriter::new();
        missing_connection
            .set_xml_maps(missing_connection_catalog)
            .unwrap();
        missing_connection.add_worksheet(MutableWorksheet::new("Missing Connection"));
        cases.push((
            "missing inert connection",
            missing_connection,
            vec![(0, 0, 0)],
        ));

        for (name, workbook, expected) in cases {
            assert_save_failure_is_atomic(workbook, &expected);
            assert!(!name.is_empty());
        }
    }

    #[test]
    fn oversized_catalog_setter_is_atomic() {
        let mut oversized = catalog();
        oversized.selection_namespaces =
            "x".repeat(litchi_ooxml_common::spreadsheet_xml_maps::MAX_STRING_BYTES + 1);
        let mut workbook = WorkbookWriter::new();
        assert!(workbook.set_xml_maps(oversized).is_err());
        assert!(workbook.xml_maps().is_none());
    }

    #[test]
    fn existing_inert_connection_satisfies_data_binding_dependency() {
        let mut value = catalog();
        value.maps[0].data_binding = Some(crate::xml_maps::DataBinding {
            data_binding_name: None,
            file_binding: Some(true),
            connection_id: Some(91),
            file_binding_name: None,
            load_mode: 0,
            payload_xml: None,
        });
        let mut workbook = WorkbookWriter::new();
        workbook.set_xml_maps(value).unwrap();
        workbook
            .set_connections(crate::package::connections::Connections {
                connections: vec![crate::package::connections::Connection {
                    connection_id: 91,
                    name: "Inert XML source".to_string(),
                    ..Default::default()
                }],
            })
            .unwrap();
        workbook.add_worksheet(MutableWorksheet::new("Mapped"));
        let mut output = Cursor::new(Vec::new());
        workbook.save(&mut output).unwrap();
        assert!(!output.get_ref().is_empty());
    }

    #[test]
    fn table_connection_rules_accept_xml_id_and_non_xml_zero() {
        let mut workbook = WorkbookWriter::new();
        workbook.set_xml_maps(catalog()).unwrap();
        workbook
            .set_connections(crate::package::connections::Connections {
                connections: vec![crate::package::connections::Connection {
                    connection_id: 91,
                    name: "Inert table source".to_string(),
                    ..Default::default()
                }],
            })
            .unwrap();
        let mut sheet = MutableWorksheet::new("Connections");
        let mut xml = xml_table(
            2,
            Type::Xml,
            1,
            Range {
                first_row: 0,
                last_row: 2,
                first_column: 0,
                last_column: 0,
            },
        );
        xml.connection_id = Some(91);
        sheet.add_table(xml).unwrap();
        let mut range = xml_table(
            4,
            Type::Range,
            1,
            Range {
                first_row: 0,
                last_row: 2,
                first_column: 1,
                last_column: 1,
            },
        );
        range.connection_id = Some(0);
        sheet.add_table(range).unwrap();
        workbook.add_worksheet(sheet);
        let mut output = Cursor::new(Vec::new());
        workbook.save(&mut output).unwrap();
        assert!(!output.get_ref().is_empty());
    }

    #[test]
    fn multi_sheet_relationship_ids_and_part_names_do_not_collide() {
        let range = Range {
            first_row: 0,
            last_row: 2,
            first_column: 0,
            last_column: 0,
        };
        let mut workbook = WorkbookWriter::new();
        workbook.set_xml_maps(catalog()).unwrap();
        for (ordinal, name) in ["First", "Second"].into_iter().enumerate() {
            let table_id = 2 + ordinal as u32 * 2;
            let mut sheet = MutableWorksheet::new(name);
            sheet
                .add_table(xml_table(table_id, Type::Xml, 1, range))
                .unwrap();
            sheet
                .set_mapped_table(mapped_table(table_id, 1, 7))
                .unwrap();
            sheet
                .set_single_cell_mapping(single_cell(table_id + 1, 4, 1))
                .unwrap();
            workbook.add_worksheet(sheet);
        }
        let mut output = Cursor::new(Vec::new());
        workbook.save(&mut output).unwrap();
        let package = OpcPackage::from_bytes(&output.into_inner()).unwrap();
        let mut targets = HashSet::new();
        for sheet_index in 1..=2 {
            let sheet = package
                .get_part(&PackURI::new(format!("/xl/worksheets/sheet{sheet_index}.bin")).unwrap())
                .unwrap();
            let relationships = sheet.rels().iter().collect::<Vec<_>>();
            let relationship_ids = relationships
                .iter()
                .map(|relationship| relationship.r_id())
                .collect::<HashSet<_>>();
            assert_eq!(relationship_ids.len(), relationships.len());
            for relationship in relationships.into_iter().filter(|relationship| {
                relationship.reltype() == rel::TABLE
                    || relationship.reltype()
                        == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableSingleCells"
            }) {
                assert!(targets.insert(relationship.target_partname().unwrap()));
            }
        }
        assert_eq!(targets.len(), 4);
    }
}
