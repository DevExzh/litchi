//! `SpreadsheetML` table-part serialization.

use std::fmt::Write as FmtWrite;

use crate::error::{Result, allocation, invalid};
use crate::sort::SortState;

use super::model::{Table, TableColumn, TableFormula, TableStyleInfo, validate_table};
use super::{MAX_OUTPUT_BYTES, limit, xml_error};

/// Write a validated table part using the transitional `SpreadsheetML` namespace.
pub fn write_table_xml(table: &Table) -> Result<Vec<u8>> {
    validate_table(table)?;
    let mut output = BoundedXml::new();
    output.literal(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main""#,
    )?;
    output.attribute_display("id", table.id)?;
    output.attribute("name", &table.name)?;
    output.attribute("displayName", &table.display_name)?;
    output.attribute("ref", &table.ref_range)?;
    if let Some(value) = table.comment.as_deref() {
        output.attribute("comment", value)?;
    }
    if let Some(value) = table.table_type {
        output.attribute("tableType", value.as_str())?;
    }
    if let Some(value) = table.header_row_count {
        output.attribute_display("headerRowCount", value)?;
    }
    if let Some(value) = table.totals_row_count {
        output.attribute_display("totalsRowCount", value)?;
    }
    if let Some(value) = table.totals_row_shown {
        output.attribute_bool("totalsRowShown", value)?;
    }
    if let Some(value) = table.published {
        output.attribute_bool("published", value)?;
    }
    output.literal(">")?;

    if let Some(auto_filter_range) = table.auto_filter_range.as_deref() {
        output.attribute_element_start("autoFilter", &[("ref", auto_filter_range)])?;
        if let Some(sort_state) = table.sort_state.as_ref() {
            output.literal(">")?;
            write_sort_state(&mut output, sort_state)?;
            output.literal("</autoFilter>")?;
        } else {
            output.close_empty_element()?;
        }
    }

    output.literal("<tableColumns")?;
    output.attribute_display("count", table.columns.len())?;
    output.literal(">")?;
    for column in &table.columns {
        write_table_column(&mut output, column)?;
    }
    output.literal("</tableColumns>")?;

    if let Some(style) = table.style_info.as_ref() {
        write_style_info(&mut output, style)?;
    }
    output.literal("</table>")?;
    Ok(output.finish())
}

/// Serialize a table part as UTF-8 text for the legacy host writer API.
pub fn serialize_table(table: &Table) -> Result<String> {
    String::from_utf8(write_table_xml(table)?).map_err(|error| xml_error(error.to_string()))
}

fn write_sort_state(output: &mut BoundedXml, sort_state: &SortState) -> Result<()> {
    output.literal("<sortState")?;
    output.attribute("ref", &sort_state.ref_range)?;
    if let Some(value) = sort_state.column_sort {
        output.attribute_bool("columnSort", value)?;
    }
    if let Some(value) = sort_state.case_sensitive {
        output.attribute_bool("caseSensitive", value)?;
    }
    if let Some(value) = sort_state.sort_method {
        output.attribute("sortMethod", value.as_str())?;
    }
    if sort_state.conditions.is_empty() {
        output.close_empty_element()?;
        return Ok(());
    }
    output.literal(">")?;
    for condition in &sort_state.conditions {
        condition
            .validate()
            .map_err(|error| invalid(error.to_string()))?;
        output.literal("<sortCondition")?;
        output.attribute("ref", &condition.ref_range)?;
        if let Some(value) = condition.descending {
            output.attribute_bool("descending", value)?;
        }
        if let Some(value) = condition.sort_by {
            output.attribute("sortBy", value.as_str())?;
        }
        if let Some(value) = condition.custom_list.as_deref() {
            output.attribute("customList", value)?;
        }
        if let Some(value) = condition.dxf_id {
            output.attribute_display("dxfId", value)?;
        }
        if let Some(value) = condition.icon_set {
            output.attribute("iconSet", value.as_str())?;
        }
        if let Some(value) = condition.icon_id {
            output.attribute_display("iconId", value)?;
        }
        output.close_empty_element()?;
    }
    output.literal("</sortState>")
}

fn write_table_column(output: &mut BoundedXml, column: &TableColumn) -> Result<()> {
    output.literal("<tableColumn")?;
    output.attribute_display("id", column.id)?;
    output.attribute("name", &column.name)?;
    if let Some(value) = column.unique_name.as_deref() {
        output.attribute("uniqueName", value)?;
    }
    if let Some(value) = column.totals_row_function {
        output.attribute("totalsRowFunction", value.as_str())?;
    }
    if let Some(value) = column.totals_row_label.as_deref() {
        output.attribute("totalsRowLabel", value)?;
    }
    if column.calculated_column_formula.is_none() && column.totals_row_formula.is_none() {
        return output.close_empty_element();
    }
    output.literal(">")?;
    if let Some(value) = column.calculated_column_formula.as_ref() {
        write_formula(output, "calculatedColumnFormula", value)?;
    }
    if let Some(value) = column.totals_row_formula.as_ref() {
        write_formula(output, "totalsRowFormula", value)?;
    }
    output.literal("</tableColumn>")
}

fn write_formula(output: &mut BoundedXml, name: &str, formula: &TableFormula) -> Result<()> {
    output.literal("<")?;
    output.literal(name)?;
    if let Some(value) = formula.array {
        output.attribute_bool("array", value)?;
    }
    output.literal(">")?;
    output.escaped(&formula.text)?;
    output.literal("</")?;
    output.literal(name)?;
    output.literal(">")
}

fn write_style_info(output: &mut BoundedXml, style: &TableStyleInfo) -> Result<()> {
    output.literal("<tableStyleInfo")?;
    if let Some(value) = style.name.as_deref() {
        output.attribute("name", value)?;
    }
    if let Some(value) = style.show_first_column {
        output.attribute_bool("showFirstColumn", value)?;
    }
    if let Some(value) = style.show_last_column {
        output.attribute_bool("showLastColumn", value)?;
    }
    if let Some(value) = style.show_row_stripes {
        output.attribute_bool("showRowStripes", value)?;
    }
    if let Some(value) = style.show_column_stripes {
        output.attribute_bool("showColumnStripes", value)?;
    }
    output.close_empty_element()
}

struct BoundedXml {
    bytes: Vec<u8>,
}

impl BoundedXml {
    fn new() -> Self {
        Self {
            bytes: Vec::with_capacity(2048),
        }
    }

    fn append(&mut self, bytes: &[u8]) -> Result<()> {
        let length = self
            .bytes
            .len()
            .checked_add(bytes.len())
            .ok_or_else(|| limit("serialized table XML bytes"))?;
        if length > MAX_OUTPUT_BYTES {
            return Err(limit("serialized table XML bytes"));
        }
        self.bytes
            .try_reserve(bytes.len())
            .map_err(|source| allocation("serialized table XML bytes", source))?;
        self.bytes.extend_from_slice(bytes);
        Ok(())
    }

    fn literal(&mut self, value: &str) -> Result<()> {
        self.append(value.as_bytes())
    }

    fn escaped(&mut self, value: &str) -> Result<()> {
        for character in value.chars() {
            match character {
                '&' => self.literal("&amp;")?,
                '<' => self.literal("&lt;")?,
                '>' => self.literal("&gt;")?,
                '"' => self.literal("&quot;")?,
                '\'' => self.literal("&apos;")?,
                _ => {
                    let mut bytes = [0; 4];
                    self.append(character.encode_utf8(&mut bytes).as_bytes())?;
                },
            }
        }
        Ok(())
    }

    fn attribute(&mut self, name: &str, value: &str) -> Result<()> {
        self.literal(" ")?;
        self.literal(name)?;
        self.literal("=\"")?;
        self.escaped(value)?;
        self.literal("\"")
    }

    fn attribute_display(&mut self, name: &str, value: impl std::fmt::Display) -> Result<()> {
        let mut text = String::new();
        write!(&mut text, "{value}")
            .map_err(|_source| invalid("failed to format table XML value"))?;
        self.attribute(name, &text)
    }

    fn attribute_bool(&mut self, name: &str, value: bool) -> Result<()> {
        self.attribute(name, if value { "1" } else { "0" })
    }

    fn attribute_element_start(&mut self, name: &str, attributes: &[(&str, &str)]) -> Result<()> {
        self.literal("<")?;
        self.literal(name)?;
        for (attribute, value) in attributes {
            self.attribute(attribute, value)?;
        }
        Ok(())
    }

    fn close_empty_element(&mut self) -> Result<()> {
        self.literal("/>")
    }

    fn finish(self) -> Vec<u8> {
        self.bytes
    }
}
