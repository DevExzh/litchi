use litchi_core::sheet::Result as SheetResult;
use litchi_core::xml::escape_xml;
use std::fmt::Write as FmtWrite;

use super::cache::{Definition, Field as CacheField, Item, Records};
use super::fields::{AxisField, AxisItem, DataField, Field, FieldItem, PageField};
use super::filters::Filter;
use super::styles::{Location, Style};

const XML_HEADER: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#;
const SPREADSHEET_NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

pub struct TableDefinition<'a> {
    pub name: &'a str,
    pub cache_id: u32,
    pub location: &'a Location,
    pub pivot_fields: &'a [Field],
    pub row_fields: &'a [AxisField],
    pub col_fields: &'a [AxisField],
    pub page_fields: &'a [PageField],
    pub data_fields: &'a [DataField],
    pub row_items: &'a [AxisItem],
    pub col_items: &'a [AxisItem],
    pub filters: &'a [Filter],
    pub style: Option<&'a Style>,
}

pub fn write_pivot_table(def: &TableDefinition<'_>) -> SheetResult<String> {
    let name = def.name;
    let cache_id = def.cache_id;
    let location = def.location;
    let pivot_fields = def.pivot_fields;
    let row_fields = def.row_fields;
    let col_fields = def.col_fields;
    let page_fields = def.page_fields;
    let data_fields = def.data_fields;
    let row_items = def.row_items;
    let col_items = def.col_items;
    let filters = def.filters;
    let style = def.style;
    let mut xml = String::with_capacity(8192);

    xml.push_str(XML_HEADER);
    write!(
        &mut xml,
        r#"<pivotTableDefinition xmlns="{}" name="{}" cacheId="{}" dataOnRows="0" dataCaption="Values" updatedVersion="3" minRefreshableVersion="3" showCalcMbrs="0" useAutoFormatting="1" itemPrintTitles="1" createdVersion="3" indent="0" compact="0" compactData="0" gridDropZones="1">"#,
        SPREADSHEET_NS,
        escape_xml(name),
        cache_id
    )?;

    write_location(&mut xml, location)?;

    if !pivot_fields.is_empty() {
        write_pivot_fields(&mut xml, pivot_fields)?;
    }

    if !row_fields.is_empty() {
        write!(&mut xml, r#"<rowFields count="{}">"#, row_fields.len())?;
        for field in row_fields {
            write!(&mut xml, r#"<field x="{}"/>"#, field.x)?;
        }
        xml.push_str("</rowFields>");
    }

    if !row_items.is_empty() {
        write_row_col_items(&mut xml, "rowItems", row_items)?;
    }

    if !col_fields.is_empty() {
        write!(&mut xml, r#"<colFields count="{}">"#, col_fields.len())?;
        for field in col_fields {
            write!(&mut xml, r#"<field x="{}"/>"#, field.x)?;
        }
        xml.push_str("</colFields>");
    }

    if !col_items.is_empty() {
        write_row_col_items(&mut xml, "colItems", col_items)?;
    }

    if !page_fields.is_empty() {
        write_page_fields(&mut xml, page_fields)?;
    }

    if !data_fields.is_empty() {
        write_data_fields(&mut xml, data_fields)?;
    }

    if !filters.is_empty() {
        write_filters(&mut xml, filters)?;
    }

    if let Some(style) = style {
        write_pivot_table_style(&mut xml, style)?;
    }

    xml.push_str("</pivotTableDefinition>");

    Ok(xml)
}

fn write_location(xml: &mut String, location: &Location) -> SheetResult<()> {
    write!(
        xml,
        r#"<location ref="{}" firstHeaderRow="{}" firstDataRow="{}" firstDataCol="{}""#,
        escape_xml(&location.reference),
        location.first_header_row,
        location.first_data_row,
        location.first_data_col
    )?;

    if let Some(row_page_count) = location.row_page_count {
        write!(xml, r#" rowPageCount="{}""#, row_page_count)?;
    }
    if let Some(col_page_count) = location.col_page_count {
        write!(xml, r#" colPageCount="{}""#, col_page_count)?;
    }

    xml.push_str("/>");
    Ok(())
}

fn write_pivot_fields(xml: &mut String, fields: &[Field]) -> SheetResult<()> {
    write!(xml, r#"<pivotFields count="{}">"#, fields.len())?;

    for field in fields {
        xml.push_str("<pivotField");

        if let Some(name) = &field.name {
            write!(xml, r#" name="{}""#, escape_xml(name))?;
        }
        if let Some(axis) = &field.axis {
            write!(xml, r#" axis="{}""#, axis.as_str())?;
        }
        if let Some(data_field) = field.data_field {
            write!(
                xml,
                r#" dataField="{}""#,
                if data_field { "1" } else { "0" }
            )?;
        }

        if !field.show_drop_downs {
            xml.push_str(r#" showDropDowns="0""#);
        }
        if !field.compact {
            xml.push_str(r#" compact="0""#);
        }
        if !field.outline {
            xml.push_str(r#" outline="0""#);
        }
        if !field.subtotal_top {
            xml.push_str(r#" subtotalTop="0""#);
        }
        if !field.drag_to_row {
            xml.push_str(r#" dragToRow="0""#);
        }
        if !field.drag_to_col {
            xml.push_str(r#" dragToCol="0""#);
        }
        if !field.drag_to_page {
            xml.push_str(r#" dragToPage="0""#);
        }
        if !field.drag_to_data {
            xml.push_str(r#" dragToData="0""#);
        }
        if !field.drag_off {
            xml.push_str(r#" dragOff="0""#);
        }
        if !field.show_all {
            xml.push_str(r#" showAll="0""#);
        }
        if !field.top_auto_show {
            xml.push_str(r#" topAutoShow="0""#);
        }
        if field.item_page_count != 10 {
            write!(xml, r#" itemPageCount="{}""#, field.item_page_count)?;
        }
        if field.sort_type.as_str() != "manual" {
            write!(xml, r#" sortType="{}""#, field.sort_type.as_str())?;
        }
        if !field.default_subtotal {
            xml.push_str(r#" defaultSubtotal="0""#);
        }

        if let Some(sum_subtotal) = field.sum_subtotal {
            write!(
                xml,
                r#" sumSubtotal="{}""#,
                if sum_subtotal { "1" } else { "0" }
            )?;
        }
        if let Some(count_a_subtotal) = field.count_a_subtotal {
            write!(
                xml,
                r#" countASubtotal="{}""#,
                if count_a_subtotal { "1" } else { "0" }
            )?;
        }
        if let Some(avg_subtotal) = field.avg_subtotal {
            write!(
                xml,
                r#" avgSubtotal="{}""#,
                if avg_subtotal { "1" } else { "0" }
            )?;
        }
        if let Some(max_subtotal) = field.max_subtotal {
            write!(
                xml,
                r#" maxSubtotal="{}""#,
                if max_subtotal { "1" } else { "0" }
            )?;
        }
        if let Some(min_subtotal) = field.min_subtotal {
            write!(
                xml,
                r#" minSubtotal="{}""#,
                if min_subtotal { "1" } else { "0" }
            )?;
        }

        if field.items.is_empty() {
            xml.push_str("/>");
        } else {
            xml.push('>');
            write_field_items(xml, &field.items)?;
            xml.push_str("</pivotField>");
        }
    }

    xml.push_str("</pivotFields>");
    Ok(())
}

fn write_field_items(xml: &mut String, items: &[FieldItem]) -> SheetResult<()> {
    write!(xml, r#"<items count="{}">"#, items.len())?;

    for item in items {
        xml.push_str("<item");

        if let Some(name) = &item.name {
            write!(xml, r#" n="{}""#, escape_xml(name))?;
        }

        let item_type_str = item.item_type.as_str();
        if item_type_str != "data" {
            write!(xml, r#" t="{}""#, item_type_str)?;
        }

        if let Some(hidden) = item.hidden {
            write!(xml, r#" h="{}""#, if hidden { "1" } else { "0" })?;
        }
        if let Some(selected) = item.selected {
            write!(xml, r#" s="{}""#, if selected { "1" } else { "0" })?;
        }
        if !item.show_detail {
            xml.push_str(r#" sd="0""#);
        }

        xml.push_str("/>");
    }

    xml.push_str("</items>");
    Ok(())
}

fn write_row_col_items(xml: &mut String, tag_name: &str, items: &[AxisItem]) -> SheetResult<()> {
    write!(xml, r#"<{} count="{}">"#, tag_name, items.len())?;

    for item in items {
        xml.push_str("<i");

        let item_type_str = item.item_type.as_str();
        if item_type_str != "data" {
            write!(xml, r#" t="{}""#, item_type_str)?;
        }
        if item.r != 0 {
            write!(xml, r#" r="{}""#, item.r)?;
        }
        if item.i != 0 {
            write!(xml, r#" i="{}""#, item.i)?;
        }

        if item.x.is_empty() {
            xml.push_str("/>");
        } else {
            xml.push('>');
            for &x_val in &item.x {
                write!(xml, r#"<x v="{}"/>"#, x_val)?;
            }
            xml.push_str("</i>");
        }
    }

    write!(xml, "</{}>", tag_name)?;
    Ok(())
}

fn write_page_fields(xml: &mut String, page_fields: &[PageField]) -> SheetResult<()> {
    write!(xml, r#"<pageFields count="{}">"#, page_fields.len())?;

    for field in page_fields {
        write!(xml, r#"<pageField fld="{}""#, field.fld)?;

        if let Some(item) = field.item {
            write!(xml, r#" item="{}""#, item)?;
        }
        if let Some(hier) = field.hier {
            write!(xml, r#" hier="{}""#, hier)?;
        }
        if let Some(name) = &field.name {
            write!(xml, r#" name="{}""#, escape_xml(name))?;
        }
        if let Some(cap) = &field.cap {
            write!(xml, r#" cap="{}""#, escape_xml(cap))?;
        }

        xml.push_str("/>");
    }

    xml.push_str("</pageFields>");
    Ok(())
}

fn write_data_fields(xml: &mut String, data_fields: &[DataField]) -> SheetResult<()> {
    write!(xml, r#"<dataFields count="{}">"#, data_fields.len())?;

    for field in data_fields {
        write!(xml, r#"<dataField"#)?;

        if let Some(name) = &field.name {
            write!(xml, r#" name="{}""#, escape_xml(name))?;
        }

        write!(xml, r#" fld="{}""#, field.fld)?;

        let subtotal_str = field.subtotal.as_str();
        if subtotal_str != "sum" {
            write!(xml, r#" subtotal="{}""#, subtotal_str)?;
        }

        if field.show_data_as != "normal" {
            write!(xml, r#" showDataAs="{}""#, escape_xml(&field.show_data_as))?;
        }
        if field.base_field != -1 {
            write!(xml, r#" baseField="{}""#, field.base_field)?;
        }
        if field.base_item != 1048832 {
            write!(xml, r#" baseItem="{}""#, field.base_item)?;
        }
        if let Some(num_fmt_id) = field.num_fmt_id {
            write!(xml, r#" numFmtId="{}""#, num_fmt_id)?;
        }

        xml.push_str("/>");
    }

    xml.push_str("</dataFields>");
    Ok(())
}

fn write_filters(xml: &mut String, filters: &[Filter]) -> SheetResult<()> {
    write!(xml, r#"<filters count="{}">"#, filters.len())?;

    for filter in filters {
        write!(
            xml,
            r#"<filter fld="{}" type="{}" id="{}""#,
            filter.fld,
            escape_xml(&filter.filter_type),
            filter.id
        )?;

        if let Some(mp_fld) = filter.mp_fld {
            write!(xml, r#" mpFld="{}""#, mp_fld)?;
        }
        if let Some(eval_order) = filter.eval_order {
            write!(xml, r#" evalOrder="{}""#, eval_order)?;
        }
        if let Some(name) = &filter.name {
            write!(xml, r#" name="{}""#, escape_xml(name))?;
        }
        if let Some(desc) = &filter.description {
            write!(xml, r#" description="{}""#, escape_xml(desc))?;
        }
        if let Some(sv1) = &filter.string_value1 {
            write!(xml, r#" stringValue1="{}""#, escape_xml(sv1))?;
        }
        if let Some(sv2) = &filter.string_value2 {
            write!(xml, r#" stringValue2="{}""#, escape_xml(sv2))?;
        }

        xml.push_str("/>");
    }

    xml.push_str("</filters>");
    Ok(())
}

fn write_pivot_table_style(xml: &mut String, style: &Style) -> SheetResult<()> {
    xml.push_str("<pivotTableStyleInfo");

    if let Some(name) = &style.name {
        write!(xml, r#" name="{}""#, escape_xml(name))?;
    }
    if let Some(show_row_headers) = style.show_row_headers {
        write!(
            xml,
            r#" showRowHeaders="{}""#,
            if show_row_headers { "1" } else { "0" }
        )?;
    }
    if let Some(show_col_headers) = style.show_col_headers {
        write!(
            xml,
            r#" showColHeaders="{}""#,
            if show_col_headers { "1" } else { "0" }
        )?;
    }
    if let Some(show_row_stripes) = style.show_row_stripes {
        write!(
            xml,
            r#" showRowStripes="{}""#,
            if show_row_stripes { "1" } else { "0" }
        )?;
    }
    if let Some(show_col_stripes) = style.show_col_stripes {
        write!(
            xml,
            r#" showColStripes="{}""#,
            if show_col_stripes { "1" } else { "0" }
        )?;
    }
    if let Some(show_last_column) = style.show_last_column {
        write!(
            xml,
            r#" showLastColumn="{}""#,
            if show_last_column { "1" } else { "0" }
        )?;
    }

    xml.push_str("/>");
    Ok(())
}

fn write_bool_attribute(xml: &mut String, name: &str, value: bool) -> SheetResult<()> {
    write!(xml, r#" {name}="{}""#, if value { "1" } else { "0" })?;
    Ok(())
}

fn validate_pivot_cache_definition(cache_def: &Definition) -> SheetResult<()> {
    if cache_def.id.as_deref() == Some("") {
        return Err("pivot cache records relationship ID cannot be empty".into());
    }
    if cache_def.source_relationship_id.as_deref() == Some("") {
        return Err("pivot cache source relationship ID cannot be empty".into());
    }
    if !matches!(
        cache_def.source_type.as_str(),
        "worksheet" | "external" | "consolidation" | "scenario"
    ) {
        return Err(format!(
            "invalid pivot cache source type '{}'",
            cache_def.source_type
        )
        .into());
    }
    let has_worksheet_source = cache_def.source_worksheet.is_some()
        || cache_def.source_ref.is_some()
        || cache_def.source_name.is_some()
        || cache_def.source_relationship_id.is_some();
    if cache_def.source_type == "worksheet" && !has_worksheet_source {
        return Err("worksheet pivot cache is missing worksheetSource metadata".into());
    }
    if cache_def.source_type != "worksheet" && has_worksheet_source {
        return Err("worksheetSource metadata requires a worksheet pivot-cache source".into());
    }
    if cache_def.cache_fields.is_empty() {
        return Err("pivot cache must contain at least one cache field".into());
    }
    if cache_def
        .refreshed_date
        .is_some_and(|value| !value.is_finite())
    {
        return Err("pivot cache refreshedDate must be finite".into());
    }
    for field in &cache_def.cache_fields {
        if field.name.is_empty() {
            return Err("pivot cache field name cannot be empty".into());
        }
        for item in &field.shared_items {
            match item {
                Item::Index(_) => {
                    return Err("shared-item indexes are only valid in pivot cache records".into());
                },
                Item::Number(value) if !value.is_finite() => {
                    return Err("pivot cache shared number must be finite".into());
                },
                _ => {},
            }
        }
    }
    Ok(())
}

pub fn write_pivot_cache_definition(cache_def: &Definition) -> SheetResult<String> {
    validate_pivot_cache_definition(cache_def)?;
    let mut xml = String::with_capacity(4096);

    xml.push_str(XML_HEADER);
    write!(
        xml,
        r#"<pivotCacheDefinition xmlns="{}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships""#,
        SPREADSHEET_NS
    )?;

    if let Some(id) = &cache_def.id {
        write!(xml, r#" r:id="{}""#, escape_xml(id))?;
    }

    if cache_def.invalid {
        xml.push_str(r#" invalid="1""#);
    }
    if !cache_def.save_data {
        xml.push_str(r#" saveData="0""#);
    }
    if cache_def.refresh_on_load {
        xml.push_str(r#" refreshOnLoad="1""#);
    }
    if let Some(value) = cache_def.optimize_memory {
        write_bool_attribute(&mut xml, "optimizeMemory", value)?;
    }
    if !cache_def.enable_refresh {
        xml.push_str(r#" enableRefresh="0""#);
    }
    if let Some(value) = &cache_def.refreshed_by {
        write!(xml, r#" refreshedBy="{}""#, escape_xml(value))?;
    }
    if let Some(value) = cache_def.refreshed_date {
        write!(xml, r#" refreshedDate="{}""#, value)?;
    }
    if let Some(value) = &cache_def.refreshed_date_iso {
        write!(xml, r#" refreshedDateIso="{}""#, escape_xml(value))?;
    }
    if cache_def.background_query {
        xml.push_str(r#" backgroundQuery="1""#);
    }
    if let Some(value) = cache_def.missing_items_limit {
        write!(xml, r#" missingItemsLimit="{}""#, value)?;
    }

    write!(
        xml,
        r#" createdVersion="{}" refreshedVersion="{}" minRefreshableVersion="{}""#,
        cache_def.created_version, cache_def.refreshed_version, cache_def.min_refreshable_version
    )?;
    if let Some(value) = cache_def.record_count {
        write!(xml, r#" recordCount="{}""#, value)?;
    }
    if let Some(value) = cache_def.upgrade_on_refresh {
        write_bool_attribute(&mut xml, "upgradeOnRefresh", value)?;
    }
    if let Some(value) = cache_def.tuples_cache {
        write_bool_attribute(&mut xml, "tupleCache", value)?;
    }
    if let Some(value) = cache_def.supports_subquery {
        write_bool_attribute(&mut xml, "supportSubquery", value)?;
    }
    if let Some(value) = cache_def.supports_advanced_drill {
        write_bool_attribute(&mut xml, "supportAdvancedDrill", value)?;
    }
    xml.push('>');

    write!(xml, r#"<cacheSource type="{}""#, cache_def.source_type)?;
    if let Some(connection_id) = cache_def.source_connection_id {
        write!(xml, r#" connectionId="{}""#, connection_id)?;
    }
    if cache_def.source_type == "worksheet" {
        xml.push('>');
        xml.push_str("<worksheetSource");

        if let Some(sheet) = &cache_def.source_worksheet {
            write!(xml, r#" sheet="{}""#, escape_xml(sheet))?;
        }
        if let Some(ref_str) = &cache_def.source_ref {
            write!(xml, r#" ref="{}""#, escape_xml(ref_str))?;
        }
        if let Some(name) = &cache_def.source_name {
            write!(xml, r#" name="{}""#, escape_xml(name))?;
        }
        if let Some(relationship_id) = &cache_def.source_relationship_id {
            write!(xml, r#" r:id="{}""#, escape_xml(relationship_id))?;
        }

        xml.push_str("/>");
        xml.push_str("</cacheSource>");
    } else {
        xml.push_str("/>");
    }

    write_cache_fields(&mut xml, &cache_def.cache_fields)?;

    xml.push_str("</pivotCacheDefinition>");

    Ok(xml)
}

fn write_cache_fields(xml: &mut String, fields: &[CacheField]) -> SheetResult<()> {
    write!(xml, r#"<cacheFields count="{}">"#, fields.len())?;

    for field in fields {
        write!(xml, r#"<cacheField name="{}""#, escape_xml(&field.name))?;

        if let Some(num_fmt_id) = field.num_fmt_id {
            write!(xml, r#" numFmtId="{}""#, num_fmt_id)?;
        }
        if !field.database_field {
            xml.push_str(r#" databaseField="0""#);
        }
        if let Some(caption) = &field.caption {
            write!(xml, r#" caption="{}""#, escape_xml(caption))?;
        }
        if let Some(property_name) = &field.property_name {
            write!(xml, r#" propertyName="{}""#, escape_xml(property_name))?;
        }
        if let Some(server_field) = field.server_field {
            write_bool_attribute(xml, "serverField", server_field)?;
        }
        if !field.unique_list {
            xml.push_str(r#" uniqueList="0""#);
        }
        if let Some(formula) = &field.formula {
            write!(xml, r#" formula="{}""#, escape_xml(formula))?;
        }
        if let Some(sql_type) = field.sql_type {
            write!(xml, r#" sqlType="{}""#, sql_type)?;
        }
        if let Some(hierarchy) = field.hierarchy {
            write!(xml, r#" hierarchy="{}""#, hierarchy)?;
        }
        if let Some(level) = field.level {
            write!(xml, r#" level="{}""#, level)?;
        }
        if let Some(mapping_count) = field.mapping_count {
            write!(xml, r#" mappingCount="{}""#, mapping_count)?;
        }
        if let Some(member_property_field) = field.member_property_field {
            write_bool_attribute(xml, "memberPropertyField", member_property_field)?;
        }

        if field.shared_items.is_empty() {
            xml.push_str("/>");
        } else {
            xml.push('>');
            write_shared_items(xml, &field.shared_items)?;
            xml.push_str("</cacheField>");
        }
    }

    xml.push_str("</cacheFields>");
    Ok(())
}

fn write_shared_items(xml: &mut String, items: &[Item]) -> SheetResult<()> {
    write!(xml, r#"<sharedItems count="{}">"#, items.len())?;

    for item in items {
        match item {
            Item::Missing => xml.push_str("<m/>"),
            Item::Index(_) => {
                return Err("shared-item indexes are only valid in pivot cache records".into());
            },
            Item::Number(n) => write!(xml, r#"<n v="{}"/>"#, n)?,
            Item::Boolean(b) => write!(xml, r#"<b v="{}"/>"#, if *b { "1" } else { "0" })?,
            Item::Error(e) => write!(xml, r#"<e v="{}"/>"#, escape_xml(e))?,
            Item::String(s) => write!(xml, r#"<s v="{}"/>"#, escape_xml(s))?,
            Item::DateTime(d) => write!(xml, r#"<d v="{}"/>"#, escape_xml(d))?,
        }
    }

    xml.push_str("</sharedItems>");
    Ok(())
}

pub fn write_pivot_cache_records(records: &Records) -> SheetResult<String> {
    for record in &records.records {
        if record
            .values
            .iter()
            .any(|value| matches!(value, Item::Number(number) if !number.is_finite()))
        {
            return Err("pivot cache record number must be finite".into());
        }
    }
    let mut xml = String::with_capacity(4096);

    xml.push_str(XML_HEADER);
    write!(
        xml,
        r#"<pivotCacheRecords xmlns="{}" count="{}">"#,
        SPREADSHEET_NS,
        records.records.len()
    )?;

    for record in &records.records {
        xml.push_str("<r>");
        for value in &record.values {
            match value {
                Item::Missing => xml.push_str("<m/>"),
                Item::Index(index) => write!(xml, r#"<x v="{}"/>"#, index)?,
                Item::Number(n) => write!(xml, r#"<n v="{}"/>"#, n)?,
                Item::Boolean(b) => write!(xml, r#"<b v="{}"/>"#, if *b { "1" } else { "0" })?,
                Item::Error(e) => write!(xml, r#"<e v="{}"/>"#, escape_xml(e))?,
                Item::String(s) => write!(xml, r#"<s v="{}"/>"#, escape_xml(s))?,
                Item::DateTime(d) => write!(xml, r#"<d v="{}"/>"#, escape_xml(d))?,
            }
        }
        xml.push_str("</r>");
    }

    xml.push_str("</pivotCacheRecords>");

    Ok(xml)
}

#[cfg(test)]
mod tests {
    use crate::pivot::cache::CacheRecord;
    use crate::pivot::reader::{read_pivot_cache_definition, read_pivot_cache_records};

    use super::*;

    fn assert_compact_xml(xml: &str) {
        assert!(
            !xml.bytes()
                .any(|byte| matches!(byte, b'\n' | b'\r' | b'\t'))
        );
        assert!(!xml.contains("> <"));
    }

    #[test]
    fn every_public_pivot_writer_emits_compact_xml() {
        let location = Location {
            reference: "A1:B2".to_owned(),
            ..Default::default()
        };
        let table = TableDefinition {
            name: "Pivot",
            cache_id: 1,
            location: &location,
            pivot_fields: &[],
            row_fields: &[],
            col_fields: &[],
            page_fields: &[],
            data_fields: &[],
            row_items: &[],
            col_items: &[],
            filters: &[],
            style: None,
        };
        assert_compact_xml(&write_pivot_table(&table).unwrap());

        let mut definition = Definition {
            source_ref: Some("A1:B2".to_owned()),
            ..Default::default()
        };
        definition.cache_fields.push(CacheField {
            name: "Value".to_owned(),
            ..Default::default()
        });
        assert_compact_xml(&write_pivot_cache_definition(&definition).unwrap());
        assert_compact_xml(&write_pivot_cache_records(&Records::default()).unwrap());
    }

    #[test]
    fn writes_complete_pivot_cache_definition_round_trip() {
        let mut definition = Definition {
            id: Some("records & data".to_string()),
            optimize_memory: Some(true),
            enable_refresh: false,
            refreshed_by: Some("A & B".to_string()),
            refreshed_date: Some(42.5),
            refreshed_date_iso: Some("2026-07-14T00:00:00Z".to_string()),
            missing_items_limit: Some(12),
            record_count: Some(2),
            upgrade_on_refresh: Some(true),
            tuples_cache: Some(false),
            supports_subquery: Some(true),
            supports_advanced_drill: Some(false),
            source_connection_id: Some(8),
            source_worksheet: Some("Data & More".to_string()),
            source_ref: Some("$A$1:$B$3".to_string()),
            source_relationship_id: Some("source-sheet".to_string()),
            ..Default::default()
        };
        definition.cache_fields.push(CacheField {
            name: "Region & Area".to_string(),
            num_fmt_id: Some(4),
            database_field: false,
            caption: Some("Area".to_string()),
            property_name: Some("Property".to_string()),
            server_field: Some(false),
            unique_list: false,
            level: Some(3),
            formula: Some("a < b".to_string()),
            sql_type: Some(-1),
            hierarchy: Some(2),
            mapping_count: Some(1),
            member_property_field: Some(true),
            shared_items: vec![Item::String("North & West".to_string()), Item::Number(2.5)],
        });

        let xml = write_pivot_cache_definition(&definition).unwrap();
        assert_compact_xml(&xml);
        let parsed = read_pivot_cache_definition(&xml).unwrap().unwrap();

        assert_eq!(parsed.id.as_deref(), Some("records & data"));
        assert_eq!(
            parsed.source_relationship_id.as_deref(),
            Some("source-sheet")
        );
        assert_eq!(parsed.refreshed_by.as_deref(), Some("A & B"));
        assert_eq!(parsed.cache_fields[0].formula.as_deref(), Some("a < b"));
        assert_eq!(parsed.cache_fields[0].mapping_count, Some(1));
        assert_eq!(parsed.cache_fields[0].member_property_field, Some(true));
    }

    #[test]
    fn writes_pivot_cache_record_indexes_round_trip() {
        let records = Records {
            records: vec![CacheRecord {
                values: vec![
                    Item::Index(3),
                    Item::Missing,
                    Item::String("A & B".to_string()),
                ],
            }],
        };

        let xml = write_pivot_cache_records(&records).unwrap();
        assert_compact_xml(&xml);
        let parsed = read_pivot_cache_records(&xml).unwrap().unwrap();

        assert!(matches!(parsed.records[0].values[0], Item::Index(3)));
        assert!(matches!(
            &parsed.records[0].values[2],
            Item::String(value) if value == "A & B"
        ));
    }

    #[test]
    fn rejects_invalid_pivot_cache_output() {
        assert!(write_pivot_cache_definition(&Definition::default()).is_err());

        let mut definition = Definition {
            source_ref: Some("A1".to_string()),
            ..Default::default()
        };
        definition.cache_fields.push(CacheField {
            name: "One".to_string(),
            shared_items: vec![Item::Index(0)],
            ..Default::default()
        });
        assert!(write_pivot_cache_definition(&definition).is_err());

        let records = Records {
            records: vec![CacheRecord {
                values: vec![Item::Number(f64::NAN)],
            }],
        };
        assert!(write_pivot_cache_records(&records).is_err());
    }
}
