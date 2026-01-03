use crate::ooxml::opc::constants::relationship_type as rt;
use crate::ooxml::opc::{OpcPackage, PackURI};
use crate::ooxml::pivot::{
    PivotAxis, PivotDataField, PivotFieldRole, PivotTable, PivotValueFunction,
};
use crate::ooxml::xlsx::parsers::workbook_parser;
use crate::sheet::Result as SheetResult;

pub fn read_pivot_tables(package: &OpcPackage) -> SheetResult<Vec<PivotTable>> {
    let workbook_uri = PackURI::new("/xl/workbook.xml")?;
    let workbook_part = package.get_part(&workbook_uri)?;
    let workbook_xml = std::str::from_utf8(workbook_part.blob())?;

    let (worksheets, _, _) = workbook_parser::parse_workbook_xml(workbook_xml)?;

    if worksheets.is_empty() {
        return Ok(Vec::new());
    }

    let workbook_rels = workbook_part.rels();
    let mut tables = Vec::new();

    for ws_info in worksheets {
        let rel = match workbook_rels.get(ws_info.relationship_id.as_str()) {
            Some(r) => r,
            None => continue,
        };

        let sheet_uri = rel.target_partname()?;
        let sheet_part = package.get_part(&sheet_uri)?;
        let sheet_rels = sheet_part.rels();

        for rel in sheet_rels.iter() {
            if rel.reltype() != rt::PIVOT_TABLE {
                continue;
            }

            let table_uri = rel.target_partname()?;
            let table_part = package.get_part(&table_uri)?;
            let xml = std::str::from_utf8(table_part.blob())?;

            if let Some(table) = parse_pivot_table_definition(xml, &ws_info.name)? {
                tables.push(table);
            }
        }
    }

    Ok(tables)
}

fn parse_pivot_table_definition(xml: &str, sheet_name: &str) -> SheetResult<Option<PivotTable>> {
    let (name, cache_id) = match extract_pivot_table_root_attrs(xml) {
        Some(v) => v,
        None => return Ok(None),
    };

    let location_ref = extract_location_ref(xml).unwrap_or_default();
    let pivot_field_names = parse_pivot_field_names(xml);

    let row_field_indexes = parse_axis_field_indexes(xml, "rowFields", "field", "x");
    let column_field_indexes = parse_axis_field_indexes(xml, "colFields", "field", "x");
    let filter_field_indexes = parse_axis_field_indexes(xml, "pageFields", "pageField", "fld");
    let data_fields = parse_data_fields(xml, &pivot_field_names);

    let row_fields = build_roles(&row_field_indexes, PivotAxis::Row, &pivot_field_names);
    let column_fields = build_roles(&column_field_indexes, PivotAxis::Column, &pivot_field_names);
    let filter_fields = build_roles(&filter_field_indexes, PivotAxis::Filter, &pivot_field_names);

    Ok(Some(PivotTable {
        name,
        source_sheet: None,
        source_ref: None,
        field_names: pivot_field_names.clone(),
        sheet_name: sheet_name.to_string(),
        cache_id,
        location_ref,
        row_fields,
        column_fields,
        filter_fields,
        data_fields,
    }))
}

fn extract_pivot_table_root_attrs(xml: &str) -> Option<(String, u32)> {
    let start = xml.find("<pivotTableDefinition")?;
    let after = &xml[start..];
    let gt_rel = after.find('>')?;
    let tag = &xml[start..start + gt_rel + 1];

    let name = extract_attr(tag, "name")?;
    let cache_id = extract_attr(tag, "cacheId")
        .and_then(|s| s.parse::<u32>().ok())
        .unwrap_or(0);

    Some((name, cache_id))
}

fn extract_location_ref(xml: &str) -> Option<String> {
    let start = xml.find("<location ")?;
    let after = &xml[start..];
    let end_rel = after.find("/>")?;
    let tag = &xml[start..start + end_rel + 2];
    extract_attr(tag, "ref")
}

fn parse_pivot_field_names(xml: &str) -> Vec<String> {
    let mut result = Vec::new();

    let start = match xml.find("<pivotFields") {
        Some(s) => s,
        None => return result,
    };

    let end_rel = match xml[start..].find("</pivotFields>") {
        Some(e) => e,
        None => return result,
    };

    let section = &xml[start..start + end_rel];
    let mut pos = 0;
    let marker = "<pivotField";

    while let Some(rel) = section[pos..].find(marker) {
        let s = pos + rel;
        let after = &section[s..];
        let gt_rel = match after.find('>') {
            Some(g) => g,
            None => break,
        };
        let tag = &section[s..s + gt_rel + 1];
        let name = extract_attr(tag, "name").unwrap_or_else(|| format!("Field{}", result.len()));
        result.push(name);
        pos = s + gt_rel + 1;
    }

    result
}

fn parse_axis_field_indexes(
    xml: &str,
    container_tag: &str,
    field_tag: &str,
    index_attr: &str,
) -> Vec<u32> {
    let mut result = Vec::new();
    let container_open = format!("<{}", container_tag);

    let start = match xml.find(&container_open) {
        Some(s) => s,
        None => return result,
    };

    let close = format!("</{}>", container_tag);
    let end_rel = match xml[start..].find(&close) {
        Some(e) => e,
        None => return result,
    };

    let section = &xml[start..start + end_rel];
    let mut pos = 0;
    let field_open = format!("<{}", field_tag);

    while let Some(rel) = section[pos..].find(&field_open) {
        let s = pos + rel;
        let after = &section[s..];
        let end_rel = match after.find("/>") {
            Some(e) => e,
            None => break,
        };
        let tag = &section[s..s + end_rel + 2];

        if let Some(idx_str) = extract_attr(tag, index_attr)
            && let Ok(idx) = idx_str.parse::<u32>()
        {
            result.push(idx);
        }

        pos = s + end_rel + 2;
    }

    result
}

fn parse_data_fields(xml: &str, pivot_field_names: &[String]) -> Vec<PivotDataField> {
    let mut result = Vec::new();

    let start = match xml.find("<dataFields") {
        Some(s) => s,
        None => return result,
    };

    let end_rel = match xml[start..].find("</dataFields>") {
        Some(e) => e,
        None => return result,
    };

    let section = &xml[start..start + end_rel];
    let mut pos = 0;
    let marker = "<dataField ";

    while let Some(rel) = section[pos..].find(marker) {
        let s = pos + rel;
        let after = &section[s..];
        let end_rel = match after.find("/>") {
            Some(e) => e,
            None => break,
        };
        let tag = &section[s..s + end_rel + 2];

        let field_index = extract_attr(tag, "fld").and_then(|s| s.parse::<u32>().ok());
        if let Some(idx) = field_index {
            let field_name = pivot_field_names
                .get(idx as usize)
                .cloned()
                .unwrap_or_else(|| format!("Field{}", idx));
            let subtotal = extract_attr(tag, "subtotal");
            let func = map_subtotal_to_function(subtotal.as_deref());
            let display_name = extract_attr(tag, "name");

            result.push(PivotDataField {
                field_name,
                function: func,
                display_name,
            });
        }

        pos = s + end_rel + 2;
    }

    result
}

fn build_roles(
    indexes: &[u32],
    axis: PivotAxis,
    pivot_field_names: &[String],
) -> Vec<PivotFieldRole> {
    let mut roles = Vec::new();

    for (position, idx) in indexes.iter().enumerate() {
        let name = pivot_field_names
            .get(*idx as usize)
            .cloned()
            .unwrap_or_else(|| format!("Field{}", idx));
        roles.push(PivotFieldRole {
            field_name: name,
            axis,
            position: position as u32,
        });
    }

    roles
}

fn map_subtotal_to_function(subtotal: Option<&str>) -> PivotValueFunction {
    match subtotal.map(|s| s.to_ascii_lowercase()) {
        None => PivotValueFunction::Sum,
        Some(ref s) if s == "sum" => PivotValueFunction::Sum,
        Some(ref s) if s == "count" || s == "counta" => PivotValueFunction::Count,
        Some(ref s) if s == "average" || s == "avg" => PivotValueFunction::Average,
        Some(ref s) if s == "min" => PivotValueFunction::Min,
        Some(ref s) if s == "max" => PivotValueFunction::Max,
        Some(_) => PivotValueFunction::Custom,
    }
}

fn extract_attr(tag: &str, attr: &str) -> Option<String> {
    let pattern = format!("{}=\"", attr);
    let start = tag.find(&pattern)? + pattern.len();
    let rest = &tag[start..];
    let end = rest.find('"')?;
    Some(rest[..end].to_string())
}
