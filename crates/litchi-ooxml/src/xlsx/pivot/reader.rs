use crate::pivot::{PivotAxis, PivotDataField, PivotFieldRole, PivotTable, PivotValueFunction};
use crate::xlsx::parsers::workbook_parser;
use litchi_core::sheet::Result as SheetResult;
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{OpcPackage, PackURI};

use super::cache::{PivotCacheDefinition, PivotCacheField, SharedItem};

pub fn read_pivot_tables(package: &OpcPackage) -> SheetResult<Vec<PivotTable>> {
    let workbook_part = package.main_document_part()?;
    let workbook_xml = std::str::from_utf8(workbook_part.blob())?;

    let (worksheets, _, _) = workbook_parser::parse_workbook_xml(workbook_xml)?;

    if worksheets.is_empty() {
        return Ok(Vec::new());
    }

    let workbook_rels = workbook_part.rels();
    let mut tables = Vec::new();

    for ws_info in worksheets {
        let rel = workbook_rels
            .get(ws_info.relationship_id.as_str())
            .ok_or_else(|| {
                format!(
                    "worksheet '{}' references missing relationship '{}'",
                    ws_info.name, ws_info.relationship_id
                )
            })?;
        if !matches!(rel.reltype(), rt::WORKSHEET | rt::STRICT_WORKSHEET) {
            return Err(format!(
                "worksheet '{}' relationship has invalid type '{}'",
                ws_info.name,
                rel.reltype()
            )
            .into());
        }
        if rel.is_external() {
            return Err(format!(
                "worksheet '{}' relationship cannot be external",
                ws_info.name
            )
            .into());
        }

        let sheet_uri = rel.target_partname()?;
        let sheet_part = package.get_part(&sheet_uri)?;
        require_content_type(&sheet_uri, sheet_part.content_type(), ct::SML_WORKSHEET)?;
        let sheet_rels = sheet_part.rels();

        for rel in sheet_rels.iter() {
            if !matches!(rel.reltype(), rt::PIVOT_TABLE | rt::STRICT_PIVOT_TABLE) {
                continue;
            }
            if rel.is_external() {
                return Err(format!(
                    "worksheet '{}' pivot-table relationship cannot be external",
                    ws_info.name
                )
                .into());
            }

            let table_uri = rel.target_partname()?;
            let table_part = package.get_part(&table_uri)?;
            require_content_type(&table_uri, table_part.content_type(), ct::SML_PIVOT_TABLE)?;
            let xml = std::str::from_utf8(table_part.blob())?;

            if let Some(table) = parse_pivot_table_definition(xml, &ws_info.name)? {
                tables.push(table);
            }
        }
    }

    Ok(tables)
}

fn require_content_type(uri: &PackURI, actual: &str, expected: &str) -> SheetResult<()> {
    if actual != expected {
        return Err(
            format!("part '{uri}' has content type '{actual}', expected '{expected}'").into(),
        );
    }
    Ok(())
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

pub fn read_pivot_table_definition(xml: &str) -> SheetResult<Option<PivotTable>> {
    parse_pivot_table_definition(xml, "")
}

pub fn read_pivot_cache_definition(xml: &str) -> SheetResult<Option<PivotCacheDefinition>> {
    let mut cache_def = PivotCacheDefinition::default();

    let root_start = match xml.find("<pivotCacheDefinition") {
        Some(s) => s,
        None => return Ok(None),
    };
    let root_after = &xml[root_start..];
    let root_end = match root_after.find('>') {
        Some(e) => e,
        None => return Ok(None),
    };
    let root_tag = &xml[root_start..root_start + root_end + 1];

    if let Some(id) = extract_attr(root_tag, "id") {
        cache_def.id = Some(id);
    }
    if let Some(val) = extract_attr(root_tag, "invalid") {
        cache_def.invalid = val == "1" || val.to_lowercase() == "true";
    }
    if let Some(val) = extract_attr(root_tag, "saveData") {
        cache_def.save_data = val == "1" || val.to_lowercase() == "true";
    }
    if let Some(val) = extract_attr(root_tag, "refreshOnLoad") {
        cache_def.refresh_on_load = val == "1" || val.to_lowercase() == "true";
    }
    if let Some(val) = extract_attr(root_tag, "backgroundQuery") {
        cache_def.background_query = val == "1" || val.to_lowercase() == "true";
    }

    if let Some(ws_source_start) = xml.find("<worksheetSource") {
        let ws_source_after = &xml[ws_source_start..];
        if let Some(ws_source_end) = ws_source_after.find("/>") {
            let ws_source_tag = &xml[ws_source_start..ws_source_start + ws_source_end + 2];
            cache_def.source_worksheet = extract_attr(ws_source_tag, "sheet");
            cache_def.source_ref = extract_attr(ws_source_tag, "ref");
            cache_def.source_name = extract_attr(ws_source_tag, "name");
        }
    }

    cache_def.cache_fields = parse_cache_fields(xml);

    Ok(Some(cache_def))
}

fn parse_cache_fields(xml: &str) -> Vec<PivotCacheField> {
    let mut fields = Vec::new();

    let start = match xml.find("<cacheFields") {
        Some(s) => s,
        None => return fields,
    };

    let end_rel = match xml[start..].find("</cacheFields>") {
        Some(e) => e,
        None => return fields,
    };

    let section = &xml[start..start + end_rel];
    let mut pos = 0;

    while let Some(rel) = section[pos..].find("<cacheField") {
        let field_start = pos + rel;
        let field_after = &section[field_start..];

        let field_end = match field_after.find("</cacheField>") {
            Some(e) => field_start + e + 13,
            None => match field_after.find("/>") {
                Some(e) => field_start + e + 2,
                None => break,
            },
        };

        let field_xml = &section[field_start..field_end];
        if let Some(field) = parse_cache_field(field_xml) {
            fields.push(field);
        }

        pos = field_end;
    }

    fields
}

fn parse_cache_field(xml: &str) -> Option<PivotCacheField> {
    let tag_end = xml.find('>')?;
    let tag = &xml[..tag_end + 1];

    let name = extract_attr(tag, "name")?;
    let database_field = extract_attr(tag, "databaseField")
        .map(|val| val == "1" || val.to_lowercase() == "true")
        .unwrap_or(true);
    let caption = extract_attr(tag, "caption");
    let num_fmt_id = extract_attr(tag, "numFmtId").and_then(|val| val.parse().ok());
    let shared_items = parse_shared_items(xml);

    Some(PivotCacheField {
        name,
        database_field,
        caption,
        num_fmt_id,
        shared_items,
        ..Default::default()
    })
}

fn parse_shared_items(xml: &str) -> Vec<SharedItem> {
    let mut items = Vec::new();

    let start = match xml.find("<sharedItems") {
        Some(s) => s,
        None => return items,
    };

    let end = match xml[start..].find("</sharedItems>") {
        Some(e) => start + e,
        None => return items,
    };

    let section = &xml[start..end];

    let mut pos = 0;
    while pos < section.len() {
        if let Some(m_pos) = section[pos..].find("<m") {
            items.push(SharedItem::Missing);
            pos += m_pos + 1;
        } else if let Some(n_pos) = section[pos..].find("<n ") {
            let n_start = pos + n_pos;
            if let Some(n_end) = section[n_start..].find("/>") {
                let n_tag = &section[n_start..n_start + n_end + 2];
                if let Some(v_str) = extract_attr(n_tag, "v")
                    && let Ok(v) = v_str.parse::<f64>()
                {
                    items.push(SharedItem::Number(v));
                }
                pos = n_start + n_end + 2;
            } else {
                break;
            }
        } else if let Some(s_pos) = section[pos..].find("<s ") {
            let s_start = pos + s_pos;
            if let Some(s_end) = section[s_start..].find("/>") {
                let s_tag = &section[s_start..s_start + s_end + 2];
                if let Some(v) = extract_attr(s_tag, "v") {
                    items.push(SharedItem::String(v));
                }
                pos = s_start + s_end + 2;
            } else {
                break;
            }
        } else if let Some(b_pos) = section[pos..].find("<b ") {
            let b_start = pos + b_pos;
            if let Some(b_end) = section[b_start..].find("/>") {
                let b_tag = &section[b_start..b_start + b_end + 2];
                if let Some(v_str) = extract_attr(b_tag, "v") {
                    let v = v_str == "1" || v_str.to_lowercase() == "true";
                    items.push(SharedItem::Boolean(v));
                }
                pos = b_start + b_end + 2;
            } else {
                break;
            }
        } else if let Some(e_pos) = section[pos..].find("<e ") {
            let e_start = pos + e_pos;
            if let Some(e_end) = section[e_start..].find("/>") {
                let e_tag = &section[e_start..e_start + e_end + 2];
                if let Some(v) = extract_attr(e_tag, "v") {
                    items.push(SharedItem::Error(v));
                }
                pos = e_start + e_end + 2;
            } else {
                break;
            }
        } else if let Some(d_pos) = section[pos..].find("<d ") {
            let d_start = pos + d_pos;
            if let Some(d_end) = section[d_start..].find("/>") {
                let d_tag = &section[d_start..d_start + d_end + 2];
                if let Some(v) = extract_attr(d_tag, "v") {
                    items.push(SharedItem::DateTime(v));
                }
                pos = d_start + d_end + 2;
            } else {
                break;
            }
        } else {
            break;
        }
    }

    items
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
    let marker = "<pivotField ";

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

#[cfg(test)]
mod tests {
    use litchi_opc::Part;
    use litchi_opc::part::BlobPart;

    use super::*;

    fn package_with_pivot_table() -> (OpcPackage, PackURI) {
        let mut package = OpcPackage::new();
        let workbook_uri = PackURI::new("/custom/book.xml").unwrap();
        let worksheet_uri = PackURI::new("/custom/sheets/data.xml").unwrap();
        let table_uri = PackURI::new("/custom/pivots/table.xml").unwrap();
        let mut workbook_part = BlobPart::new(
            workbook_uri,
            ct::SML_SHEET_MAIN.to_string(),
            br#"<workbook xmlns="http://purl.oclc.org/ooxml/spreadsheetml/main"
                    xmlns:r="http://purl.oclc.org/ooxml/officeDocument/relationships">
                    <sheets><sheet name="Pivot" sheetId="1" r:id="rId1"/></sheets>
                </workbook>"#
                .to_vec(),
        );
        workbook_part.relate_to("sheets/data.xml", rt::STRICT_WORKSHEET);
        let mut worksheet_part = BlobPart::new(
            worksheet_uri.clone(),
            ct::SML_WORKSHEET.to_string(),
            Vec::new(),
        );
        worksheet_part.relate_to("../pivots/table.xml", rt::STRICT_PIVOT_TABLE);
        package.relate_to("custom/book.xml", rt::STRICT_OFFICE_DOCUMENT);
        package.add_part(Box::new(workbook_part));
        package.add_part(Box::new(worksheet_part));
        package.add_part(Box::new(BlobPart::new(
            table_uri,
            ct::SML_PIVOT_TABLE.to_string(),
            br#"<pivotTableDefinition xmlns="http://purl.oclc.org/ooxml/spreadsheetml/main"
                    name="PivotOne" cacheId="7" dataCaption="Values">
                    <location ref="A1:C5" firstHeaderRow="1" firstDataRow="2" firstDataCol="1"/>
                    <pivotFields count="1"><pivotField name="Region"/></pivotFields>
                    <rowFields count="1"><field x="0"/></rowFields>
                </pivotTableDefinition>"#
                .to_vec(),
        )));
        (package, worksheet_uri)
    }

    #[test]
    fn resolves_strict_custom_pivot_table_parts() {
        let (package, _) = package_with_pivot_table();
        let tables = read_pivot_tables(&package).unwrap();

        assert_eq!(tables.len(), 1);
        assert_eq!(tables[0].name, "PivotOne");
        assert_eq!(tables[0].sheet_name, "Pivot");
        assert_eq!(tables[0].row_fields[0].field_name, "Region");
    }

    #[test]
    fn rejects_external_and_wrong_content_type_pivot_parts() {
        let (mut package, worksheet_uri) = package_with_pivot_table();
        package
            .get_part_mut(&worksheet_uri)
            .unwrap()
            .rels_mut()
            .add_relationship(
                rt::STRICT_PIVOT_TABLE.to_string(),
                "https://example.com/pivot.xml".to_string(),
                "rId1".to_string(),
                true,
            );
        assert!(read_pivot_tables(&package).is_err());

        let (mut package, _) = package_with_pivot_table();
        package.add_part(Box::new(BlobPart::new(
            PackURI::new("/custom/pivots/table.xml").unwrap(),
            ct::SML_WORKSHEET.to_string(),
            Vec::new(),
        )));
        assert!(read_pivot_tables(&package).is_err());
    }
}
