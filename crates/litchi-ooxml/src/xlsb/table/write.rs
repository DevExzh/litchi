//! Serializer for the XLSB Table (ListObject) stream (MS-XLSB 2.1.7.51).
//!
//! This is the exact inverse of `parse.rs`: record order, payload layouts,
//! nullable-string and DXFId encodings all mirror the reader so authored
//! tables round-trip through `parse_table_part` and, at the package level,
//! through `Workbook::structured_tables`.

use crate::xlsb::error::{Error, Result};
use crate::xlsb::table::model::{Column, Formula, StyleInfo, Table};
use litchi_xlsb::raw::Writer;
use litchi_xlsb::raw::kind as rt;

/// `DXFId` value meaning no differential formatting (MS-XLSB 2.5.38).
const NO_DXF: u32 = u32::MAX;
/// `dwConnID` value meaning no external connection (MS-XLSB 2.4.100).
const NO_CONNECTION: u32 = 0;

// `BrtBeginList` flags word (MS-XLSB 2.4.100).
const LIST_SHOWN_TOTAL_ROW: u32 = 1 << 0;
const LIST_SINGLE_CELL: u32 = 1 << 1;
const LIST_INSERT_ROW_VISIBLE: u32 = 1 << 2;
const LIST_INSERT_ROW_INSERTED_CELLS: u32 = 1 << 3;
const LIST_PUBLISHED: u32 = 1 << 4;

// `BrtTableStyleClient` flags word (MS-XLSB 2.4.847).
const STYLE_FIRST_COLUMN: u16 = 1 << 0;
const STYLE_LAST_COLUMN: u16 = 1 << 1;
const STYLE_ROW_STRIPES: u16 = 1 << 2;
const STYLE_COLUMN_STRIPES: u16 = 1 << 3;

// `BrtListCCFmla`/`BrtListTrFmla` flags byte (MS-XLSB 2.4.706).
const FORMULA_ARRAY: u8 = 1 << 1;

/// Leading blank field of `BrtList14` (MS-XLSB 2.4.705).
const FRT_BLANK_LEN: usize = 4;

fn malformed(context: &str, detail: impl Into<String>) -> Error {
    Error::Unrecognized {
        typ: context.to_string(),
        val: detail.into(),
    }
}

fn write_wide_string(data: &mut Vec<u8>, value: &str) {
    data.extend_from_slice(&(value.encode_utf16().count() as u32).to_le_bytes());
    for unit in value.encode_utf16() {
        data.extend_from_slice(&unit.to_le_bytes());
    }
}

fn write_nullable_wide_string(data: &mut Vec<u8>, value: &Option<String>) {
    match value {
        Some(value) => write_wide_string(data, value),
        None => data.extend_from_slice(&u32::MAX.to_le_bytes()),
    }
}

fn write_dxf_id(data: &mut Vec<u8>, id: Option<u32>) {
    data.extend_from_slice(&id.unwrap_or(NO_DXF).to_le_bytes());
}

fn write_blob(data: &mut Vec<u8>, blob: &[u8]) {
    data.extend_from_slice(&(blob.len() as u32).to_le_bytes());
    data.extend_from_slice(blob);
}

/// `BrtBeginList` payload (MS-XLSB 2.4.100).
fn list_payload(table: &Table) -> Vec<u8> {
    let mut data = Vec::with_capacity(64);
    data.extend_from_slice(&table.range.first_row.to_le_bytes());
    data.extend_from_slice(&table.range.last_row.to_le_bytes());
    data.extend_from_slice(&table.range.first_column.to_le_bytes());
    data.extend_from_slice(&table.range.last_column.to_le_bytes());
    data.extend_from_slice(&(table.table_type as u32).to_le_bytes());
    data.extend_from_slice(&table.id.to_le_bytes());
    data.extend_from_slice(&table.header_row_count.to_le_bytes());
    data.extend_from_slice(&table.totals_row_count.to_le_bytes());
    let mut flags = 0u32;
    if table.totals_row_shown {
        flags |= LIST_SHOWN_TOTAL_ROW;
    }
    if table.single_cell {
        flags |= LIST_SINGLE_CELL;
    }
    if table.insert_row_visible {
        flags |= LIST_INSERT_ROW_VISIBLE;
    }
    if table.insert_row_inserted_cells {
        flags |= LIST_INSERT_ROW_INSERTED_CELLS;
    }
    if table.published {
        flags |= LIST_PUBLISHED;
    }
    data.extend_from_slice(&flags.to_le_bytes());
    write_dxf_id(&mut data, table.header_dxf_id);
    write_dxf_id(&mut data, table.data_dxf_id);
    write_dxf_id(&mut data, table.totals_dxf_id);
    write_dxf_id(&mut data, table.border_dxf_id);
    write_dxf_id(&mut data, table.header_border_dxf_id);
    write_dxf_id(&mut data, table.totals_border_dxf_id);
    data.extend_from_slice(&table.connection_id.unwrap_or(NO_CONNECTION).to_le_bytes());
    write_nullable_wide_string(&mut data, &table.name);
    write_nullable_wide_string(&mut data, &table.display_name);
    write_nullable_wide_string(&mut data, &table.comment);
    write_nullable_wide_string(&mut data, &table.header_style);
    write_nullable_wide_string(&mut data, &table.data_style);
    write_nullable_wide_string(&mut data, &table.totals_style);
    data
}

/// `BrtBeginListCol` payload (MS-XLSB 2.4.101).
fn column_payload(column: &Column) -> Vec<u8> {
    let mut data = Vec::with_capacity(48);
    data.extend_from_slice(&column.id.to_le_bytes());
    data.extend_from_slice(&(column.totals_row_function as u32).to_le_bytes());
    write_dxf_id(&mut data, column.header_dxf_id);
    write_dxf_id(&mut data, column.insert_row_dxf_id);
    write_dxf_id(&mut data, column.totals_dxf_id);
    data.extend_from_slice(&column.query_table_field_id.to_le_bytes());
    write_nullable_wide_string(&mut data, &column.name);
    write_nullable_wide_string(&mut data, &column.caption);
    write_nullable_wide_string(&mut data, &column.totals_row_label);
    write_nullable_wide_string(&mut data, &column.header_style);
    write_nullable_wide_string(&mut data, &column.insert_row_style);
    write_nullable_wide_string(&mut data, &column.totals_style);
    data
}

/// `BrtListCCFmla`/`BrtListTrFmla` payload (MS-XLSB 2.4.706, 2.4.708).
fn formula_payload(formula: &Formula) -> Vec<u8> {
    let mut data = Vec::with_capacity(formula.tokens.len() + formula.extra.len() + 9);
    data.push(if formula.array { FORMULA_ARRAY } else { 0 });
    write_blob(&mut data, &formula.tokens);
    write_blob(&mut data, &formula.extra);
    data
}

/// `BrtTableStyleClient` payload (MS-XLSB 2.4.847).
fn style_client_payload(style: &StyleInfo) -> Vec<u8> {
    let mut flags = 0u16;
    if style.show_first_column {
        flags |= STYLE_FIRST_COLUMN;
    }
    if style.show_last_column {
        flags |= STYLE_LAST_COLUMN;
    }
    if style.show_row_stripes {
        flags |= STYLE_ROW_STRIPES;
    }
    if style.show_column_stripes {
        flags |= STYLE_COLUMN_STRIPES;
    }
    let mut data = Vec::with_capacity(16);
    data.extend_from_slice(&flags.to_le_bytes());
    write_nullable_wide_string(&mut data, &style.name);
    data
}

/// Serialize one table into its complete table-part stream.
pub(crate) fn write_table_part(table: &Table) -> Result<Vec<u8>> {
    let mut data = Vec::with_capacity(256);
    let mut writer = Writer::new(&mut data);
    writer.write_record(rt::BEGIN_LIST, &list_payload(table))?;

    if !table.columns.is_empty() {
        let declared = u32::try_from(table.columns.len())
            .map_err(|_| malformed("BrtBeginListCols", "column count overflow"))?;
        writer.write_record(rt::BEGIN_LIST_COLS, &declared.to_le_bytes())?;
        for column in &table.columns {
            writer.write_record(rt::BEGIN_LIST_COL, &column_payload(column))?;
            if let Some(formula) = &column.calculated_column_formula {
                writer.write_record(rt::LIST_CC_FMLA, &formula_payload(formula))?;
            }
            if let Some(formula) = &column.totals_row_formula {
                writer.write_record(rt::LIST_TR_FMLA, &formula_payload(formula))?;
            }
            writer.write_record(rt::END_LIST_COL, &[])?;
        }
        writer.write_record(rt::END_LIST_COLS, &[])?;
    }

    if let Some(style) = &table.style_info {
        writer.write_record(rt::TABLE_STYLE_CLIENT, &style_client_payload(style))?;
    }

    if table.alternate_text.is_some() || table.alternate_text_summary.is_some() {
        let mut payload = vec![0u8; FRT_BLANK_LEN];
        write_nullable_wide_string(&mut payload, &table.alternate_text);
        write_nullable_wide_string(&mut payload, &table.alternate_text_summary);
        writer.write_record(rt::LIST14, &payload)?;
    }

    writer.write_record(rt::END_LIST, &[])?;
    Ok(data)
}

/// Serialize a worksheet's `BrtBeginListParts` collection (MS-XLSB 2.4.103).
pub(crate) fn write_list_parts<W: std::io::Write>(
    writer: &mut Writer<W>,
    rel_ids: &[String],
) -> Result<()> {
    let declared = u32::try_from(rel_ids.len())
        .map_err(|_| malformed("BrtBeginListParts", "table count overflow"))?;
    writer.write_record(rt::BEGIN_LIST_PARTS, &declared.to_le_bytes())?;
    for rel_id in rel_ids {
        let mut payload = Vec::with_capacity(rel_id.len() * 2 + 4);
        write_wide_string(&mut payload, rel_id);
        writer.write_record(rt::LIST_PART, &payload)?;
    }
    writer.write_record(rt::END_LIST_PARTS, &[])?;
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::xlsb::table::model::{Range, TotalsRowFunction, Type};
    use crate::xlsb::table::parse_table_part;

    fn sample_table() -> Table {
        Table {
            id: 7,
            name: Some("SalesTable".to_string()),
            display_name: Some("SalesTable".to_string()),
            comment: Some("Quarterly sales".to_string()),
            range: Range {
                first_row: 0,
                last_row: 9,
                first_column: 1,
                last_column: 3,
            },
            table_type: Type::Range,
            header_row_count: 1,
            totals_row_count: 1,
            totals_row_shown: true,
            published: true,
            columns: vec![
                Column {
                    id: 1,
                    name: Some("Region".to_string()),
                    totals_row_function: TotalsRowFunction::None,
                    ..Column::default()
                },
                Column {
                    id: 2,
                    name: Some("Amount".to_string()),
                    caption: Some("Total".to_string()),
                    totals_row_function: TotalsRowFunction::Sum,
                    totals_row_label: Some("Grand Total".to_string()),
                    ..Column::default()
                },
                Column {
                    id: 3,
                    name: Some("Ratio".to_string()),
                    calculated_column_formula: Some(Formula {
                        array: true,
                        tokens: vec![0x1E, 0x02],
                        extra: Vec::new(),
                    }),
                    ..Column::default()
                },
            ],
            style_info: Some(StyleInfo {
                name: Some("TableStyleMedium2".to_string()),
                show_first_column: false,
                show_last_column: false,
                show_row_stripes: true,
                show_column_stripes: false,
            }),
            alternate_text: Some("Sales by region".to_string()),
            alternate_text_summary: Some("Summary".to_string()),
            ..Table::default()
        }
    }

    #[test]
    fn serialized_table_round_trips_through_the_reader() {
        let table = sample_table();
        let bytes = write_table_part(&table).unwrap();
        let parsed = parse_table_part(&bytes).unwrap();
        assert_eq!(parsed, table);
    }

    #[test]
    fn serialized_minimal_table_round_trips() {
        let table = Table {
            id: 1,
            display_name: Some("T".to_string()),
            range: Range {
                first_row: 0,
                last_row: 0,
                first_column: 0,
                last_column: 0,
            },
            single_cell: true,
            ..Table::default()
        };
        let bytes = write_table_part(&table).unwrap();
        let parsed = parse_table_part(&bytes).unwrap();
        assert_eq!(parsed, table);
    }

    #[test]
    fn serialized_list_parts_round_trip_through_rel_id_reader() {
        // Wrap in a worksheet-shaped stream: sheet begin, list parts, sheet end.
        let mut stream = Vec::new();
        let mut writer = Writer::new(&mut stream);
        writer.write_record(rt::BEGIN_SHEET, &[]).unwrap();
        writer
            .write_record(rt::BEGIN_LIST_PARTS, &2u32.to_le_bytes())
            .unwrap();
        for rel_id in ["rId5", "rId9"] {
            let mut payload = Vec::new();
            write_wide_string(&mut payload, rel_id);
            writer.write_record(rt::LIST_PART, &payload).unwrap();
        }
        writer.write_record(rt::END_LIST_PARTS, &[]).unwrap();
        writer.write_record(rt::END_SHEET, &[]).unwrap();
        let rel_ids = crate::xlsb::table::parse_table_part_rel_ids(&stream).unwrap();
        assert_eq!(rel_ids, ["rId5", "rId9"]);
    }
}
