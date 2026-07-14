use crate::pivot::{PivotAxis, PivotDataField, PivotFieldRole, PivotTable, PivotValueFunction};
use crate::xlsx::parsers::workbook_parser;
use litchi_core::sheet::Result as SheetResult;
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{OpcPackage, PackURI};
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::ResolveResult;
use quick_xml::reader::NsReader;

use super::cache::{PivotCacheDefinition, PivotCacheField, SharedItem};
use crate::common::xml::unqualified_attribute_value;
use crate::xlsx::Cell;
use crate::xlsx::namespace::is_spreadsheetml_name;

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

            let table = parse_pivot_table_definition(xml, &ws_info.name)?.ok_or_else(|| {
                format!("pivot-table part '{table_uri}' has no pivotTableDefinition root")
            })?;
            tables.push(table);
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

#[derive(Clone, Copy, PartialEq, Eq)]
enum TableContext {
    Root,
    Location,
    PivotFields,
    PivotField,
    RowFields,
    ColFields,
    PageFields,
    DataFields,
    Other,
}

struct RawDataField {
    field_index: u32,
    function: PivotValueFunction,
    display_name: Option<String>,
}

struct PivotTableParser {
    name: String,
    cache_id: u32,
    sheet_name: String,
    location_ref: String,
    field_names: Vec<String>,
    row_indexes: Vec<u32>,
    column_indexes: Vec<u32>,
    row_field_count: usize,
    column_field_count: usize,
    filter_indexes: Vec<u32>,
    data_fields: Vec<RawDataField>,
    expected_pivot_fields: Option<u32>,
    expected_row_fields: Option<u32>,
    expected_col_fields: Option<u32>,
    expected_page_fields: Option<u32>,
    expected_data_fields: Option<u32>,
    saw_location: bool,
    saw_pivot_fields: bool,
    saw_row_fields: bool,
    saw_col_fields: bool,
    saw_page_fields: bool,
    saw_data_fields: bool,
}

impl PivotTableParser {
    fn from_root(
        element: &BytesStart<'_>,
        decoder: Decoder,
        sheet_name: &str,
    ) -> SheetResult<Self> {
        let name = required_string(element, b"name", decoder, "pivot-table name")?;
        if name.is_empty() {
            return Err("pivot-table name cannot be empty".into());
        }
        let cache_id = required_u32(element, b"cacheId", decoder, "pivot-table cache ID")?;
        required_string(element, b"dataCaption", decoder, "pivot-table data caption")?;
        Ok(Self {
            name,
            cache_id,
            sheet_name: sheet_name.to_string(),
            location_ref: String::new(),
            field_names: Vec::new(),
            row_indexes: Vec::new(),
            column_indexes: Vec::new(),
            row_field_count: 0,
            column_field_count: 0,
            filter_indexes: Vec::new(),
            data_fields: Vec::new(),
            expected_pivot_fields: None,
            expected_row_fields: None,
            expected_col_fields: None,
            expected_page_fields: None,
            expected_data_fields: None,
            saw_location: false,
            saw_pivot_fields: false,
            saw_row_fields: false,
            saw_col_fields: false,
            saw_page_fields: false,
            saw_data_fields: false,
        })
    }

    fn parse(xml: &str, sheet_name: &str) -> SheetResult<Option<PivotTable>> {
        let mut reader = NsReader::from_reader(xml.as_bytes());
        let mut parser: Option<Self> = None;
        let mut stack = Vec::new();
        let mut closed_root = false;

        loop {
            let decoder = reader.decoder();
            let event = reader.read_event()?.into_owned();
            let resolver = reader.resolver().clone();
            let (namespace, event) = resolver.resolve_event(event);
            match event {
                Event::Start(element) if stack.is_empty() => {
                    if closed_root {
                        return Err("pivot table contains multiple root elements".into());
                    }
                    if !is_spreadsheetml_name(&namespace, element.name(), b"pivotTableDefinition") {
                        return Ok(None);
                    }
                    parser = Some(Self::from_root(&element, decoder, sheet_name)?);
                    stack.push(TableContext::Root);
                },
                Event::Empty(element) if stack.is_empty() => {
                    if closed_root {
                        return Err("pivot table contains multiple root elements".into());
                    }
                    if !is_spreadsheetml_name(&namespace, element.name(), b"pivotTableDefinition") {
                        return Ok(None);
                    }
                    let root = Self::from_root(&element, decoder, sheet_name)?;
                    root.finish(TableContext::Root)?;
                    return Ok(Some(root.build()?));
                },
                Event::Start(element) => {
                    let parent = *stack.last().ok_or("pivot table is missing its root")?;
                    let context = parser
                        .as_mut()
                        .ok_or("pivot table parser is not initialized")?
                        .start(parent, &namespace, &element, decoder)?;
                    stack.push(context);
                },
                Event::Empty(element) => {
                    let parent = *stack.last().ok_or("pivot table is missing its root")?;
                    let parser = parser
                        .as_mut()
                        .ok_or("pivot table parser is not initialized")?;
                    let context = parser.start(parent, &namespace, &element, decoder)?;
                    parser.finish(context)?;
                },
                Event::End(element) => {
                    let context = stack
                        .pop()
                        .ok_or("pivot table has a closing element outside its root")?;
                    parser
                        .as_mut()
                        .ok_or("pivot table parser is not initialized")?
                        .finish(context)?;
                    if context == TableContext::Root {
                        if !is_spreadsheetml_name(
                            &namespace,
                            element.name(),
                            b"pivotTableDefinition",
                        ) {
                            return Err("pivot table has an invalid root closing element".into());
                        }
                        closed_root = true;
                    }
                },
                Event::Eof if parser.is_none() => return Ok(None),
                Event::Eof if !closed_root || !stack.is_empty() => {
                    return Err("pivot table has a missing or unterminated root".into());
                },
                Event::Eof => break,
                _ => {},
            }
        }

        parser.map(Self::build).transpose()
    }

    fn start(
        &mut self,
        parent: TableContext,
        namespace: &ResolveResult<'_>,
        element: &BytesStart<'_>,
        decoder: Decoder,
    ) -> SheetResult<TableContext> {
        if parent == TableContext::Root
            && is_spreadsheetml_name(namespace, element.name(), b"location")
        {
            mark_once(&mut self.saw_location, "pivot-table location")?;
            self.location_ref =
                required_string(element, b"ref", decoder, "pivot-table location reference")?;
            validate_cell_range(&self.location_ref, "pivot-table location")?;
            required_u32(
                element,
                b"firstHeaderRow",
                decoder,
                "pivot first header row",
            )?;
            required_u32(element, b"firstDataRow", decoder, "pivot first data row")?;
            required_u32(element, b"firstDataCol", decoder, "pivot first data column")?;
            return Ok(TableContext::Location);
        }
        if parent == TableContext::Root
            && is_spreadsheetml_name(namespace, element.name(), b"pivotFields")
        {
            mark_once(&mut self.saw_pivot_fields, "pivotFields")?;
            self.expected_pivot_fields =
                optional_u32(element, b"count", decoder, "pivotFields count")?;
            return Ok(TableContext::PivotFields);
        }
        if parent == TableContext::PivotFields
            && is_spreadsheetml_name(namespace, element.name(), b"pivotField")
        {
            let name = unqualified_attribute_value(element, b"name", decoder)?
                .unwrap_or_else(|| format!("Field{}", self.field_names.len()));
            self.field_names.push(name);
            return Ok(TableContext::PivotField);
        }
        if parent == TableContext::Root
            && is_spreadsheetml_name(namespace, element.name(), b"rowFields")
        {
            mark_once(&mut self.saw_row_fields, "rowFields")?;
            self.expected_row_fields = optional_u32(element, b"count", decoder, "rowFields count")?;
            return Ok(TableContext::RowFields);
        }
        if parent == TableContext::Root
            && is_spreadsheetml_name(namespace, element.name(), b"colFields")
        {
            mark_once(&mut self.saw_col_fields, "colFields")?;
            self.expected_col_fields = optional_u32(element, b"count", decoder, "colFields count")?;
            return Ok(TableContext::ColFields);
        }
        if matches!(parent, TableContext::RowFields | TableContext::ColFields)
            && is_spreadsheetml_name(namespace, element.name(), b"field")
        {
            let index = required_i32(element, b"x", decoder, "pivot axis field index")?;
            if parent == TableContext::RowFields {
                self.row_field_count += 1;
            } else {
                self.column_field_count += 1;
            }
            if index >= 0 {
                let index = u32::try_from(index).map_err(|_| "pivot field index overflows")?;
                if parent == TableContext::RowFields {
                    self.row_indexes.push(index);
                } else {
                    self.column_indexes.push(index);
                }
            } else if index != -2 {
                return Err(format!("invalid pivot axis field index '{index}'").into());
            }
            return Ok(TableContext::Other);
        }
        if parent == TableContext::Root
            && is_spreadsheetml_name(namespace, element.name(), b"pageFields")
        {
            mark_once(&mut self.saw_page_fields, "pageFields")?;
            self.expected_page_fields =
                optional_u32(element, b"count", decoder, "pageFields count")?;
            return Ok(TableContext::PageFields);
        }
        if parent == TableContext::PageFields
            && is_spreadsheetml_name(namespace, element.name(), b"pageField")
        {
            let index = required_i32(element, b"fld", decoder, "pivot page field index")?;
            if index < 0 {
                return Err(format!("invalid pivot page field index '{index}'").into());
            }
            self.filter_indexes
                .push(u32::try_from(index).map_err(|_| "pivot page field index overflows")?);
            return Ok(TableContext::Other);
        }
        if parent == TableContext::Root
            && is_spreadsheetml_name(namespace, element.name(), b"dataFields")
        {
            mark_once(&mut self.saw_data_fields, "dataFields")?;
            self.expected_data_fields =
                optional_u32(element, b"count", decoder, "dataFields count")?;
            return Ok(TableContext::DataFields);
        }
        if parent == TableContext::DataFields
            && is_spreadsheetml_name(namespace, element.name(), b"dataField")
        {
            let field_index = required_u32(element, b"fld", decoder, "pivot data field index")?;
            let subtotal = unqualified_attribute_value(element, b"subtotal", decoder)?;
            let function = parse_subtotal(subtotal.as_deref())?;
            self.data_fields.push(RawDataField {
                field_index,
                function,
                display_name: unqualified_attribute_value(element, b"name", decoder)?,
            });
            return Ok(TableContext::Other);
        }
        Ok(TableContext::Other)
    }

    fn finish(&self, context: TableContext) -> SheetResult<()> {
        match context {
            TableContext::Root if !self.saw_location => {
                Err("pivot table is missing its required location".into())
            },
            TableContext::PivotFields => validate_count(
                self.expected_pivot_fields,
                self.field_names.len(),
                "pivotFields",
            ),
            TableContext::RowFields => {
                validate_count(self.expected_row_fields, self.row_field_count, "rowFields")
            },
            TableContext::ColFields => validate_count(
                self.expected_col_fields,
                self.column_field_count,
                "colFields",
            ),
            TableContext::PageFields => validate_count(
                self.expected_page_fields,
                self.filter_indexes.len(),
                "pageFields",
            ),
            TableContext::DataFields => validate_count(
                self.expected_data_fields,
                self.data_fields.len(),
                "dataFields",
            ),
            _ => Ok(()),
        }
    }

    fn build(self) -> SheetResult<PivotTable> {
        if !self.field_names.is_empty() {
            for index in self
                .row_indexes
                .iter()
                .chain(&self.column_indexes)
                .chain(&self.filter_indexes)
                .chain(self.data_fields.iter().map(|field| &field.field_index))
            {
                let index_out_of_range = match usize::try_from(*index) {
                    Ok(index) => index >= self.field_names.len(),
                    Err(_) => true,
                };
                if index_out_of_range {
                    return Err(format!(
                        "pivot field index {index} exceeds {} pivot fields",
                        self.field_names.len()
                    )
                    .into());
                }
            }
        }
        let row_fields = build_roles(&self.row_indexes, PivotAxis::Row, &self.field_names);
        let column_fields = build_roles(&self.column_indexes, PivotAxis::Column, &self.field_names);
        let filter_fields = build_roles(&self.filter_indexes, PivotAxis::Filter, &self.field_names);
        let data_fields = self
            .data_fields
            .into_iter()
            .map(|field| PivotDataField {
                field_name: self
                    .field_names
                    .get(field.field_index as usize)
                    .cloned()
                    .unwrap_or_else(|| format!("Field{}", field.field_index)),
                function: field.function,
                display_name: field.display_name,
            })
            .collect();
        Ok(PivotTable {
            name: self.name,
            source_sheet: None,
            source_ref: None,
            field_names: self.field_names,
            sheet_name: self.sheet_name,
            cache_id: self.cache_id,
            location_ref: self.location_ref,
            row_fields,
            column_fields,
            filter_fields,
            data_fields,
        })
    }
}

fn parse_pivot_table_definition(xml: &str, sheet_name: &str) -> SheetResult<Option<PivotTable>> {
    PivotTableParser::parse(xml, sheet_name)
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

fn parse_subtotal(subtotal: Option<&str>) -> SheetResult<PivotValueFunction> {
    match subtotal.unwrap_or("sum") {
        "sum" => Ok(PivotValueFunction::Sum),
        "count" | "countNums" => Ok(PivotValueFunction::Count),
        "average" => Ok(PivotValueFunction::Average),
        "min" => Ok(PivotValueFunction::Min),
        "max" => Ok(PivotValueFunction::Max),
        "product" | "stdDev" | "stdDevp" | "var" | "varp" => Ok(PivotValueFunction::Custom),
        value => Err(format!("invalid pivot data subtotal '{value}'").into()),
    }
}

fn required_string(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    description: &str,
) -> SheetResult<String> {
    unqualified_attribute_value(element, name, decoder)?
        .ok_or_else(|| format!("missing {description} attribute").into())
}

fn optional_u32(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    description: &str,
) -> SheetResult<Option<u32>> {
    unqualified_attribute_value(element, name, decoder)?
        .map(|value| {
            value
                .parse::<u32>()
                .map_err(|_| format!("invalid {description} '{value}'").into())
        })
        .transpose()
}

fn required_u32(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    description: &str,
) -> SheetResult<u32> {
    optional_u32(element, name, decoder, description)?
        .ok_or_else(|| format!("missing {description} attribute").into())
}

fn required_i32(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    description: &str,
) -> SheetResult<i32> {
    let value = required_string(element, name, decoder, description)?;
    value
        .parse::<i32>()
        .map_err(|_| format!("invalid {description} '{value}'").into())
}

fn mark_once(seen: &mut bool, description: &str) -> SheetResult<()> {
    if *seen {
        return Err(format!("duplicate {description} element").into());
    }
    *seen = true;
    Ok(())
}

fn validate_count(expected: Option<u32>, actual: usize, description: &str) -> SheetResult<()> {
    if let Some(expected) = expected
        && usize::try_from(expected) != Ok(actual)
    {
        return Err(
            format!("{description} count is {expected}, but {actual} elements were found").into(),
        );
    }
    Ok(())
}

fn validate_cell_range(range: &str, description: &str) -> SheetResult<()> {
    let mut references = range.split(':');
    let first = references
        .next()
        .ok_or_else(|| format!("empty {description}"))?;
    Cell::reference_to_coords(&first.replace('$', ""))?;
    if let Some(second) = references.next() {
        Cell::reference_to_coords(&second.replace('$', ""))?;
    }
    if references.next().is_some() {
        return Err(format!("invalid {description} range '{range}'").into());
    }
    Ok(())
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

    #[test]
    fn parses_prefixed_pivot_table_definition() {
        let xml = r#"<p:pivotTableDefinition
                xmlns:p="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                xmlns:f="urn:foreign" name="Sales &amp; Margin" cacheId="4" dataCaption="Values">
                <f:location ref="XFE1"/>
                <p:location ref="$A$1:$C$5" firstHeaderRow="1" firstDataRow="2" firstDataCol="1"/>
                <p:pivotFields count="2"><p:pivotField name="Region"/><p:pivotField/></p:pivotFields>
                <p:rowFields count="2"><p:field x="-2"/><p:field x="0"/></p:rowFields>
                <p:colFields count="1"><p:field x="1"/></p:colFields>
                <p:pageFields count="1"><p:pageField fld="0"/></p:pageFields>
                <p:dataFields count="1"><p:dataField fld="1" subtotal="average" name="Average Margin"/></p:dataFields>
            </p:pivotTableDefinition>"#;
        let table = read_pivot_table_definition(xml).unwrap().unwrap();

        assert_eq!(table.name, "Sales & Margin");
        assert_eq!(table.location_ref, "$A$1:$C$5");
        assert_eq!(table.field_names, ["Region", "Field1"]);
        assert_eq!(table.row_fields.len(), 1);
        assert_eq!(table.column_fields[0].field_name, "Field1");
        assert_eq!(table.filter_fields[0].field_name, "Region");
        assert_eq!(table.data_fields[0].function, PivotValueFunction::Average);
    }

    #[test]
    fn rejects_malformed_pivot_table_definitions() {
        const S: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        let root = |body: &str, attributes: &str| {
            format!(
                r#"<pivotTableDefinition xmlns="{S}" name="P" cacheId="1" dataCaption="V" {attributes}>{body}</pivotTableDefinition>"#
            )
        };
        let location =
            r#"<location ref="A1:B2" firstHeaderRow="1" firstDataRow="1" firstDataCol="1"/>"#;
        let invalid = [
            format!(
                r#"<pivotTableDefinition xmlns="{S}" name="P" cacheId="1">{location}</pivotTableDefinition>"#
            ),
            root("", ""),
            root(
                r#"<location ref="XFE1" firstHeaderRow="1" firstDataRow="1" firstDataCol="1"/>"#,
                "",
            ),
            root(
                &format!(r#"{location}<pivotFields count="2"><pivotField/></pivotFields>"#),
                "",
            ),
            root(
                &format!(
                    r#"{location}<location ref="A1" firstHeaderRow="1" firstDataRow="1" firstDataCol="1"/>"#
                ),
                "",
            ),
            root(
                &format!(
                    r#"{location}<dataFields><dataField fld="0" subtotal="median"/></dataFields>"#
                ),
                "",
            ),
            root(
                &format!(
                    r#"{location}<pivotFields><pivotField/></pivotFields><rowFields><field x="1"/></rowFields>"#
                ),
                "",
            ),
        ];
        for xml in invalid {
            assert!(read_pivot_table_definition(&xml).is_err(), "accepted {xml}");
        }
        assert!(
            read_pivot_table_definition(r#"<pivotTableDefinition xmlns="urn:foreign"/>"#)
                .unwrap()
                .is_none()
        );
    }
}
