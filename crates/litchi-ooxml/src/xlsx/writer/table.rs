//! Table XML serialization for XLSX.

use crate::xlsx::Cell;
use crate::xlsx::table::{Table, TableColumn, TableFormula, TableStyleInfo};
use litchi_core::sheet::Result as SheetResult;
use litchi_core::xml::escape::escape_xml;
use std::collections::HashSet;
use std::fmt::Write as FmtWrite;

fn validate_table(table: &Table) -> SheetResult<()> {
    if table.id == 0 {
        return Err("table ID must be positive".into());
    }
    if table.name.is_empty() || table.display_name.is_empty() {
        return Err("table name and display name cannot be empty".into());
    }
    if table.columns.is_empty() {
        return Err("table must contain at least one column".into());
    }
    let width = table_range_width(&table.ref_range)?;
    if usize::try_from(width) != Ok(table.columns.len()) {
        return Err(format!(
            "table range spans {width} columns, but {} columns were provided",
            table.columns.len()
        )
        .into());
    }
    if let Some(auto_filter_range) = &table.auto_filter_range {
        validate_nested_range(&table.ref_range, auto_filter_range, "auto-filter")?;
    }
    if let Some(sort_state) = &table.sort_state {
        let auto_filter_range = table
            .auto_filter_range
            .as_deref()
            .ok_or("table sort state requires an auto-filter")?;
        validate_nested_range(auto_filter_range, &sort_state.ref_range, "sort state")?;
        for condition in &sort_state.conditions {
            validate_nested_range(
                &sort_state.ref_range,
                &condition.ref_range,
                "sort condition",
            )?;
        }
    }
    let mut ids = HashSet::with_capacity(table.columns.len());
    let mut names = HashSet::with_capacity(table.columns.len());
    for column in &table.columns {
        if column.id == 0 || !ids.insert(column.id) {
            return Err(format!("invalid or duplicate table column ID {}", column.id).into());
        }
        if column.name.is_empty() || !names.insert(column.name.to_ascii_lowercase()) {
            return Err(format!("empty or duplicate table column name '{}'", column.name).into());
        }
    }
    Ok(())
}

fn table_range_width(range: &str) -> SheetResult<u32> {
    let ((first_column, _), (second_column, _)) = table_range_bounds(range)?;
    Ok(second_column - first_column + 1)
}

fn validate_nested_range(container: &str, nested: &str, description: &str) -> SheetResult<()> {
    let ((first_column, first_row), (last_column, last_row)) = table_range_bounds(container)?;
    let ((nested_first_column, nested_first_row), (nested_last_column, nested_last_row)) =
        table_range_bounds(nested)?;
    if nested_first_column < first_column
        || nested_first_row < first_row
        || nested_last_column > last_column
        || nested_last_row > last_row
    {
        return Err(format!("{description} range '{nested}' is outside '{container}'").into());
    }
    Ok(())
}

fn table_range_bounds(range: &str) -> SheetResult<((u32, u32), (u32, u32))> {
    let mut references = range.split(':');
    let first = references.next().ok_or("empty table range")?;
    let second = references.next().unwrap_or(first);
    if references.next().is_some() {
        return Err(format!("invalid table range '{range}'").into());
    }
    let (first_column, first_row) = Cell::reference_to_coords(&first.replace('$', ""))?;
    let (second_column, second_row) = Cell::reference_to_coords(&second.replace('$', ""))?;
    if second_column < first_column || second_row < first_row {
        return Err(format!("table range '{range}' is descending").into());
    }
    Ok(((first_column, first_row), (second_column, second_row)))
}

/// Serialize a table to XML.
pub fn serialize_table(table: &Table) -> SheetResult<String> {
    validate_table(table)?;
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
    if let Some(published) = table.published {
        write!(xml, r#" published="{}""#, if published { 1 } else { 0 })
            .map_err(|e| format!("XML write error: {}", e))?;
    }

    xml.push('>');

    // Auto-filter
    if let Some(ref auto_filter_range) = table.auto_filter_range {
        if let Some(ref sort_state) = table.sort_state {
            write!(
                xml,
                r#"<autoFilter ref="{}">"#,
                escape_xml(auto_filter_range)
            )
            .map_err(|e| format!("XML write error: {}", e))?;
            serialize_sort_state(&mut xml, sort_state)?;
            xml.push_str("</autoFilter>");
        } else {
            write!(
                xml,
                r#"<autoFilter ref="{}"/>"#,
                escape_xml(auto_filter_range)
            )
            .map_err(|e| format!("XML write error: {}", e))?;
        }
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
    sort_state: &crate::xlsx::sort::SortState,
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
            condition.validate()?;
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
            if let Some(custom_list) = &condition.custom_list {
                write!(xml, r#" customList="{}""#, escape_xml(custom_list))
                    .map_err(|e| format!("XML write error: {}", e))?;
            }
            if let Some(dxf_id) = condition.dxf_id {
                write!(xml, r#" dxfId="{}""#, dxf_id)
                    .map_err(|e| format!("XML write error: {}", e))?;
            }
            if let Some(icon_set) = condition.icon_set {
                write!(xml, r#" iconSet="{}""#, icon_set.as_str())
                    .map_err(|e| format!("XML write error: {}", e))?;
            }
            if let Some(icon_id) = condition.icon_id {
                write!(xml, r#" iconId="{}""#, icon_id)
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
    use crate::xlsx::sort::{SortCondition, SortState};
    use crate::xlsx::table::{TableColumn, TableFormula, TableStyleInfo, TotalsRowFunction};

    fn create_test_table() -> Table {
        let mut table = Table::new(1u32, "TestTable", "A1:B5");
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
        assert!(xml.contains(r#"ref="A1:B5""#));
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
        table.auto_filter_range = Some("A1:B5".to_string());
        let xml = serialize_table(&table).unwrap();

        assert!(xml.contains(r#"<autoFilter ref="A1:B5"/>"#));
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
        table.ref_range = "A1:A5".to_string();

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
        table.ref_range = "A1:A5".to_string();

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
            ref_range: "A2:A5".to_string(),
            descending: Some(true),
            sort_by: None,
            custom_list: None,
            dxf_id: None,
            icon_set: None,
            icon_id: None,
        };
        table.sort_state = Some(SortState {
            ref_range: "A2:B5".to_string(),
            column_sort: Some(true),
            case_sensitive: Some(false),
            sort_method: None,
            conditions: vec![sort_condition],
        });
        table.auto_filter_range = Some("A1:B5".to_string());

        let xml = serialize_table(&table).unwrap();
        assert!(xml.contains(r#"<sortState ref="A2:B5" columnSort="1" caseSensitive="0">"#));
        assert!(xml.contains(r#"<sortCondition ref="A2:A5" descending="1"/>"#));
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

    #[test]
    fn serialized_table_round_trips_through_parser() {
        use crate::xlsx::conditional_formatting::IconSet;
        use crate::xlsx::sort::SortBy;
        use crate::xlsx::table::parse_table_xml;

        let mut table = create_test_table();
        table.published = Some(true);
        table.auto_filter_range = Some("A1:B4".to_string());
        table.sort_state = Some(SortState {
            ref_range: "A2:B4".to_string(),
            column_sort: Some(false),
            case_sensitive: Some(true),
            sort_method: None,
            conditions: vec![SortCondition {
                ref_range: "B2:B4".to_string(),
                descending: Some(true),
                sort_by: Some(SortBy::Icon),
                custom_list: Some("High,Low".to_string()),
                dxf_id: None,
                icon_set: Some(IconSet::ThreeArrows),
                icon_id: Some(2),
            }],
        });
        table.columns[0].calculated_column_formula = Some(TableFormula {
            array: Some(false),
            text: "=[@[Column B]]*2".to_string(),
        });

        let xml = serialize_table(&table).unwrap();
        let parsed = parse_table_xml(&xml).unwrap().unwrap();

        assert_eq!(parsed.id, table.id);
        assert_eq!(parsed.display_name, table.display_name);
        assert_eq!(parsed.ref_range, table.ref_range);
        assert_eq!(parsed.published, Some(true));
        assert_eq!(parsed.auto_filter_range.as_deref(), Some("A1:B4"));
        assert_eq!(parsed.columns.len(), 2);
        assert_eq!(
            parsed.columns[0]
                .calculated_column_formula
                .as_ref()
                .map(|formula| formula.text.as_str()),
            Some("=[@[Column B]]*2")
        );
        let condition = &parsed.sort_state.unwrap().conditions[0];
        assert_eq!(condition.sort_by, Some(SortBy::Icon));
        assert_eq!(condition.custom_list.as_deref(), Some("High,Low"));
        assert_eq!(condition.dxf_id, None);
        assert_eq!(condition.icon_set, Some(IconSet::ThreeArrows));
        assert_eq!(condition.icon_id, Some(2));
    }

    #[test]
    fn rejects_invalid_table_models() {
        let mut table = create_test_table();
        table.columns[1].id = 1;
        assert!(serialize_table(&table).is_err());

        let mut table = create_test_table();
        table.auto_filter_range = Some("A1:C5".to_string());
        assert!(serialize_table(&table).is_err());

        let mut table = create_test_table();
        table.sort_state = Some(SortState {
            ref_range: "A2:B5".to_string(),
            column_sort: None,
            case_sensitive: None,
            sort_method: None,
            conditions: vec![SortCondition::new("C2:C5")],
        });
        table.auto_filter_range = Some("A1:B5".to_string());
        assert!(serialize_table(&table).is_err());

        let mut table = create_test_table();
        table.sort_state = Some(SortState::new("A2:B5"));
        assert!(serialize_table(&table).is_err());
    }
}
