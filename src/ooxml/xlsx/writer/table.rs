//! Table XML serialization for XLSX.

use crate::common::xml::escape::escape_xml;
use crate::ooxml::xlsx::table::{Table, TableColumn, TableFormula, TableStyleInfo};
use crate::sheet::Result as SheetResult;
use std::fmt::Write as FmtWrite;

/// Serialize a table to XML.
pub fn serialize_table(table: &Table) -> SheetResult<String> {
    let mut xml = String::with_capacity(2048);

    // Table root element with namespace
    write!(
        xml,
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="{}" name="{}" displayName="{}" ref="{}""#,
        table.id,
        escape_xml(&table.name),
        escape_xml(&table.display_name),
        escape_xml(&table.ref_range)
    )
    .map_err(|e| format!("XML write error: {}", e))?;

    // Optional attributes
    if let Some(ref comment) = table.comment {
        write!(xml, r#" comment="{}""#, escape_xml(comment))
            .map_err(|e| format!("XML write error: {}", e))?;
    }
    if let Some(table_type) = table.table_type {
        write!(xml, r#" tableType="{}""#, table_type.as_str())
            .map_err(|e| format!("XML write error: {}", e))?;
    }
    if let Some(count) = table.header_row_count {
        write!(xml, r#" headerRowCount="{}""#, count)
            .map_err(|e| format!("XML write error: {}", e))?;
    }
    if let Some(count) = table.totals_row_count {
        write!(xml, r#" totalsRowCount="{}""#, count)
            .map_err(|e| format!("XML write error: {}", e))?;
    }
    if let Some(shown) = table.totals_row_shown {
        write!(xml, r#" totalsRowShown="{}""#, if shown { 1 } else { 0 })
            .map_err(|e| format!("XML write error: {}", e))?;
    }

    xml.push('>');

    // Auto-filter
    if let Some(ref auto_filter_range) = table.auto_filter_range {
        write!(
            xml,
            r#"<autoFilter ref="{}"/>"#,
            escape_xml(auto_filter_range)
        )
        .map_err(|e| format!("XML write error: {}", e))?;
    }

    // Sort state
    if let Some(ref sort_state) = table.sort_state {
        serialize_sort_state(&mut xml, sort_state)?;
    }

    // Table columns
    if !table.columns.is_empty() {
        write!(xml, r#"<tableColumns count="{}">"#, table.columns.len())
            .map_err(|e| format!("XML write error: {}", e))?;

        for column in &table.columns {
            serialize_table_column(&mut xml, column)?;
        }

        xml.push_str("</tableColumns>");
    }

    // Table style info
    if let Some(ref style_info) = table.style_info {
        serialize_table_style_info(&mut xml, style_info)?;
    }

    xml.push_str("</table>");
    Ok(xml)
}

fn serialize_sort_state(
    xml: &mut String,
    sort_state: &crate::ooxml::xlsx::sort::SortState,
) -> SheetResult<()> {
    write!(
        xml,
        r#"<sortState ref="{}""#,
        escape_xml(&sort_state.ref_range)
    )
    .map_err(|e| format!("XML write error: {}", e))?;

    if let Some(v) = sort_state.column_sort {
        write!(xml, r#" columnSort="{}""#, if v { 1 } else { 0 })
            .map_err(|e| format!("XML write error: {}", e))?;
    }
    if let Some(v) = sort_state.case_sensitive {
        write!(xml, r#" caseSensitive="{}""#, if v { 1 } else { 0 })
            .map_err(|e| format!("XML write error: {}", e))?;
    }
    if let Some(method) = sort_state.sort_method {
        write!(xml, r#" sortMethod="{}""#, method.as_str())
            .map_err(|e| format!("XML write error: {}", e))?;
    }

    if sort_state.conditions.is_empty() {
        xml.push_str("/>");
    } else {
        xml.push('>');
        for condition in &sort_state.conditions {
            write!(
                xml,
                r#"<sortCondition ref="{}""#,
                escape_xml(&condition.ref_range)
            )
            .map_err(|e| format!("XML write error: {}", e))?;

            if let Some(v) = condition.descending {
                write!(xml, r#" descending="{}""#, if v { 1 } else { 0 })
                    .map_err(|e| format!("XML write error: {}", e))?;
            }
            if let Some(sort_by) = condition.sort_by {
                write!(xml, r#" sortBy="{}""#, sort_by.as_str())
                    .map_err(|e| format!("XML write error: {}", e))?;
            }

            xml.push_str("/>");
        }
        xml.push_str("</sortState>");
    }

    Ok(())
}

fn serialize_table_column(xml: &mut String, column: &TableColumn) -> SheetResult<()> {
    write!(
        xml,
        r#"<tableColumn id="{}" name="{}""#,
        column.id,
        escape_xml(&column.name)
    )
    .map_err(|e| format!("XML write error: {}", e))?;

    if let Some(ref unique_name) = column.unique_name {
        write!(xml, r#" uniqueName="{}""#, escape_xml(unique_name))
            .map_err(|e| format!("XML write error: {}", e))?;
    }
    if let Some(func) = column.totals_row_function {
        write!(xml, r#" totalsRowFunction="{}""#, func.as_str())
            .map_err(|e| format!("XML write error: {}", e))?;
    }
    if let Some(ref label) = column.totals_row_label {
        write!(xml, r#" totalsRowLabel="{}""#, escape_xml(label))
            .map_err(|e| format!("XML write error: {}", e))?;
    }

    // Check if we have nested elements
    let has_nested =
        column.calculated_column_formula.is_some() || column.totals_row_formula.is_some();

    if has_nested {
        xml.push('>');

        if let Some(ref formula) = column.calculated_column_formula {
            serialize_table_formula(xml, "calculatedColumnFormula", formula)?;
        }
        if let Some(ref formula) = column.totals_row_formula {
            serialize_table_formula(xml, "totalsRowFormula", formula)?;
        }

        xml.push_str("</tableColumn>");
    } else {
        xml.push_str("/>");
    }

    Ok(())
}

fn serialize_table_formula(xml: &mut String, tag: &str, formula: &TableFormula) -> SheetResult<()> {
    write!(xml, "<{}", tag).map_err(|e| format!("XML write error: {}", e))?;

    if let Some(array) = formula.array {
        write!(xml, r#" array="{}""#, if array { 1 } else { 0 })
            .map_err(|e| format!("XML write error: {}", e))?;
    }

    xml.push('>');
    xml.push_str(&escape_xml(&formula.text));
    write!(xml, "</{}>", tag).map_err(|e| format!("XML write error: {}", e))?;

    Ok(())
}

fn serialize_table_style_info(xml: &mut String, style_info: &TableStyleInfo) -> SheetResult<()> {
    xml.push_str("<tableStyleInfo");

    if let Some(ref name) = style_info.name {
        write!(xml, r#" name="{}""#, escape_xml(name))
            .map_err(|e| format!("XML write error: {}", e))?;
    }
    if let Some(v) = style_info.show_first_column {
        write!(xml, r#" showFirstColumn="{}""#, if v { 1 } else { 0 })
            .map_err(|e| format!("XML write error: {}", e))?;
    }
    if let Some(v) = style_info.show_last_column {
        write!(xml, r#" showLastColumn="{}""#, if v { 1 } else { 0 })
            .map_err(|e| format!("XML write error: {}", e))?;
    }
    if let Some(v) = style_info.show_row_stripes {
        write!(xml, r#" showRowStripes="{}""#, if v { 1 } else { 0 })
            .map_err(|e| format!("XML write error: {}", e))?;
    }
    if let Some(v) = style_info.show_column_stripes {
        write!(xml, r#" showColumnStripes="{}""#, if v { 1 } else { 0 })
            .map_err(|e| format!("XML write error: {}", e))?;
    }

    xml.push_str("/>");
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::ooxml::xlsx::sort::{SortCondition, SortState};
    use crate::ooxml::xlsx::table::{TableColumn, TableFormula, TableStyleInfo, TotalsRowFunction};

    fn create_test_table() -> Table {
        let mut table = Table::new(1u32, "TestTable", "A1:D5");
        table.display_name = "Test Table".to_string();
        table.comment = Some("Test comment".to_string());
        table.header_row_count = Some(1);
        table.totals_row_count = Some(1);
        table.totals_row_shown = Some(true);
        table.columns = vec![
            TableColumn::new(1u32, "Column A"),
            TableColumn::new(2u32, "Column B"),
        ];
        table
    }

    #[test]
    fn test_serialize_table_basic() {
        let table = create_test_table();
        let xml = serialize_table(&table).unwrap();

        assert!(xml.contains(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><table"#));
        assert!(xml.contains(r#"id="1""#));
        assert!(xml.contains(r#"name="TestTable""#));
        assert!(xml.contains(r#"displayName="Test Table""#));
        assert!(xml.contains(r#"ref="A1:D5""#));
        assert!(xml.contains(r#"comment="Test comment""#));
        assert!(xml.contains(r#"headerRowCount="1""#));
        assert!(xml.contains(r#"totalsRowCount="1""#));
        assert!(xml.contains(r#"totalsRowShown="1""#));
        assert!(xml.contains("</table>"));
    }

    #[test]
    fn test_serialize_table_with_columns() {
        let table = create_test_table();
        let xml = serialize_table(&table).unwrap();

        assert!(xml.contains(r#"<tableColumns count="2">"#));
        assert!(xml.contains(r#"<tableColumn id="1" name="Column A"/>"#));
        assert!(xml.contains(r#"<tableColumn id="2" name="Column B"/>"#));
        assert!(xml.contains("</tableColumns>"));
    }

    #[test]
    fn test_serialize_table_with_auto_filter() {
        let mut table = create_test_table();
        table.auto_filter_range = Some("A1:D5".to_string());
        let xml = serialize_table(&table).unwrap();

        assert!(xml.contains(r#"<autoFilter ref="A1:D5"/>"#));
    }

    #[test]
    fn test_serialize_table_with_style_info() {
        let mut table = create_test_table();
        let mut style_info = TableStyleInfo::new();
        style_info.name = Some("TableStyleMedium2".to_string());
        style_info.show_first_column = Some(true);
        style_info.show_last_column = Some(false);
        style_info.show_row_stripes = Some(true);
        style_info.show_column_stripes = Some(false);
        table.style_info = Some(style_info);

        let xml = serialize_table(&table).unwrap();

        assert!(xml.contains(r#"<tableStyleInfo name="TableStyleMedium2" showFirstColumn="1" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>"#));
    }

    #[test]
    fn test_serialize_table_column_with_totals_function() {
        let mut table = create_test_table();
        let mut col = TableColumn::new(1u32, "Sales");
        col.totals_row_function = Some(TotalsRowFunction::Sum);
        table.columns = vec![col];

        let xml = serialize_table(&table).unwrap();
        assert!(xml.contains(r#"totalsRowFunction="sum""#));
    }

    #[test]
    fn test_serialize_table_column_with_formula() {
        let mut table = create_test_table();
        let mut col = TableColumn::new(1u32, "Calculated");
        col.calculated_column_formula = Some(TableFormula {
            array: Some(false),
            text: "=[@Price]*[@Qty]".to_string(),
        });
        table.columns = vec![col];

        let xml = serialize_table(&table).unwrap();
        // The formula is serialized with array="0" attribute when array is Some(false)
        assert!(
            xml.contains("<calculatedColumnFormula"),
            "Expected <calculatedColumnFormula> in XML: {}",
            xml
        );
        assert!(
            xml.contains("=[@Price]*[@Qty]"),
            "Expected formula text in XML: {}",
            xml
        );
    }

    #[test]
    fn test_serialize_table_with_sort_state() {
        let mut table = create_test_table();
        let sort_condition = SortCondition {
            ref_range: "A2:A10".to_string(),
            descending: Some(true),
            sort_by: None,
            custom_list: None,
            dxf_id: None,
            icon_set: None,
            icon_id: None,
        };
        table.sort_state = Some(SortState {
            ref_range: "A2:D10".to_string(),
            column_sort: Some(true),
            case_sensitive: Some(false),
            sort_method: None,
            conditions: vec![sort_condition],
        });

        let xml = serialize_table(&table).unwrap();
        assert!(xml.contains(r#"<sortState ref="A2:D10" columnSort="1" caseSensitive="0">"#));
        assert!(xml.contains(r#"<sortCondition ref="A2:A10" descending="1"/>"#));
        assert!(xml.contains("</sortState>"));
    }

    #[test]
    fn test_serialize_table_escapes_xml() {
        let mut table = create_test_table();
        table.name = "Table<>&\"'".to_string();
        table.display_name = "Test <Table>".to_string();

        let xml = serialize_table(&table).unwrap();
        assert!(xml.contains("Table&lt;&gt;&amp;")); // XML escaped
    }
}
