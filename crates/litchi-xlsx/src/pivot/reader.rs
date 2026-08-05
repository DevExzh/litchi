use std::collections::HashMap;

use crate::raw::parse_catalog;
use litchi_core::sheet::Result as SheetResult;
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{OpcPackage, PackURI};
use super::{
    PivotAxis, PivotDataField, PivotFieldRole, PivotTable, PivotValueFunction,
};
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{NamespaceResolver, ResolveResult};
use quick_xml::reader::NsReader;

use super::cache::{
    CacheRecord, Definition, Field, Records, Item,
};
use crate::raw::namespace::{is_spreadsheetml_name, relationship_attribute_value};
use litchi_ooxml_common::xml::unqualified_attribute_value;
use litchi_sheet::Rect;

pub fn read_pivot_tables(package: &OpcPackage) -> SheetResult<Vec<PivotTable>> {
    let workbook_part = package.main_document_part()?;
    let workbook_xml = workbook_part.blob();

    let workbook = parse_catalog(workbook_xml)?;

    let workbook_rels = workbook_part.rels();
    let mut worksheet_uris = HashMap::with_capacity(workbook.sheets.len());
    let mut worksheet_names_by_uri = HashMap::with_capacity(workbook.sheets.len());
    for worksheet in &workbook.sheets {
        let rel = workbook_rels
            .get(&worksheet.relationship_id)
            .ok_or_else(|| {
                format!(
                    "worksheet '{}' references missing relationship '{}'",
                    worksheet.name, worksheet.relationship_id
                )
            })?;
        // Non-worksheet sheets (chartsheets, dialog sheets, macro sheets)
        // cannot host pivot tables and are skipped rather than rejected.
        if !matches!(rel.reltype(), rt::WORKSHEET | rt::STRICT_WORKSHEET) {
            continue;
        }
        if rel.is_external() {
            return Err(format!(
                "worksheet '{}' relationship cannot be external",
                worksheet.name
            )
            .into());
        }
        let worksheet_uri = rel.target_partname()?;
        let worksheet_part = package.get_part(&worksheet_uri)?;
        require_content_type(
            &worksheet_uri,
            worksheet_part.content_type(),
            ct::SML_WORKSHEET,
        )?;
        if worksheet_names_by_uri
            .insert(worksheet_uri.clone(), worksheet.name.clone())
            .is_some()
        {
            return Err(format!("multiple workbook sheets target part '{worksheet_uri}'").into());
        }
        worksheet_uris.insert(worksheet.relationship_id.clone(), worksheet_uri);
    }
    let (pivot_caches, pivot_cache_ids_by_uri) = resolve_workbook_pivot_caches(
        package,
        workbook_rels,
        &workbook.pivot_caches,
        &worksheet_names_by_uri,
    )?;
    if workbook.sheets.is_empty() {
        return Ok(Vec::new());
    }
    let mut tables = Vec::new();

    for ws_info in workbook.sheets {
        // Sheets skipped above (chartsheets and other non-worksheet kinds)
        // have no resolved worksheet URI and are ignored here as well.
        let Some(sheet_uri) = worksheet_uris.get(&ws_info.relationship_id) else {
            continue;
        };
        let sheet_uri = sheet_uri.clone();
        let sheet_part = package.get_part(&sheet_uri)?;
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
            let cache_uri = resolve_pivot_table_cache_uri(table_part)?;
            let expected_cache_id = pivot_cache_ids_by_uri.get(&cache_uri).ok_or_else(|| {
                format!(
                    "pivot-table part '{table_uri}' references cache definition '{cache_uri}' that is not listed by the workbook"
                )
            })?;
            let cache = pivot_caches
                .get(expected_cache_id)
                .ok_or("resolved pivot cache is missing")?;
            let bytes = litchi_ooxml_common::mce::process_part(table_part)?;
            let xml = std::str::from_utf8(bytes.as_ref())?;

            let mut table = parse_pivot_table_definition_with_cache(
                xml,
                &ws_info.name,
                Some(&cache.definition.cache_fields),
            )?
            .ok_or_else(|| {
                format!("pivot-table part '{table_uri}' has no pivotTableDefinition root")
            })?;
            if table.cache_id != *expected_cache_id {
                return Err(format!(
                    "pivot-table part '{table_uri}' declares cache ID {}, but its relationship targets workbook cache ID {expected_cache_id}",
                    table.cache_id
                )
                .into());
            }
            table.source_sheet = cache.definition.source_worksheet.clone();
            table.source_ref = cache.definition.source_ref.clone();
            tables.push(table);
        }
    }

    Ok(tables)
}

struct ResolvedPivotCache {
    definition: Definition,
}

fn resolve_workbook_pivot_caches(
    package: &OpcPackage,
    workbook_rels: &litchi_opc::Relationships,
    cache_references: &[crate::raw::PivotCache],
    worksheet_names_by_uri: &HashMap<PackURI, String>,
) -> SheetResult<(HashMap<u32, ResolvedPivotCache>, HashMap<PackURI, u32>)> {
    let mut caches = HashMap::with_capacity(cache_references.len());
    let mut ids_by_uri = HashMap::with_capacity(cache_references.len());
    for cache_reference in cache_references {
        let rel = workbook_rels
            .get(&cache_reference.relationship_id)
            .ok_or_else(|| {
                format!(
                    "workbook pivot cache {} references missing relationship '{}'",
                    cache_reference.cache_id, cache_reference.relationship_id
                )
            })?;
        if !matches!(
            rel.reltype(),
            rt::PIVOT_CACHE_DEFINITION | rt::STRICT_PIVOT_CACHE_DEFINITION
        ) {
            return Err(format!(
                "workbook pivot cache {} relationship has invalid type '{}'",
                cache_reference.cache_id,
                rel.reltype()
            )
            .into());
        }
        if rel.is_external() {
            return Err(format!(
                "workbook pivot cache {} relationship cannot be external",
                cache_reference.cache_id
            )
            .into());
        }
        let cache_uri = rel.target_partname()?;
        let cache_part = package.get_part(&cache_uri)?;
        require_content_type(
            &cache_uri,
            cache_part.content_type(),
            ct::SML_PIVOT_CACHE_DEFINITION,
        )?;
        let bytes = litchi_ooxml_common::mce::process_part(cache_part)?;
        let xml = std::str::from_utf8(bytes.as_ref())?;
        let mut definition = read_pivot_cache_definition(xml)?.ok_or_else(|| {
            format!("pivot-cache part '{cache_uri}' has no pivotCacheDefinition root")
        })?;
        validate_pivot_cache_relationships(
            package,
            cache_part,
            &cache_uri,
            &mut definition,
            worksheet_names_by_uri,
        )?;
        if ids_by_uri
            .insert(cache_uri.clone(), cache_reference.cache_id)
            .is_some()
        {
            return Err(
                format!("multiple workbook pivot cache IDs target part '{cache_uri}'").into(),
            );
        }
        caches.insert(cache_reference.cache_id, ResolvedPivotCache { definition });
    }
    Ok((caches, ids_by_uri))
}

fn validate_pivot_cache_relationships(
    package: &OpcPackage,
    cache_part: &dyn litchi_opc::Part,
    cache_uri: &PackURI,
    definition: &mut Definition,
    worksheet_names_by_uri: &HashMap<PackURI, String>,
) -> SheetResult<()> {
    if let Some(relationship_id) = definition.id.as_deref() {
        let rel = cache_part.rels().get(relationship_id).ok_or_else(|| {
            format!(
                "pivot-cache part '{cache_uri}' references missing records relationship '{relationship_id}'"
            )
        })?;
        if !matches!(
            rel.reltype(),
            rt::PIVOT_CACHE_RECORDS | rt::STRICT_PIVOT_CACHE_RECORDS
        ) {
            return Err(format!(
                "pivot-cache part '{cache_uri}' records relationship has invalid type '{}'",
                rel.reltype()
            )
            .into());
        }
        if rel.is_external() {
            return Err(format!(
                "pivot-cache part '{cache_uri}' records relationship cannot be external"
            )
            .into());
        }
        let records_uri = rel.target_partname()?;
        let records_part = package.get_part(&records_uri)?;
        require_content_type(
            &records_uri,
            records_part.content_type(),
            ct::SML_PIVOT_CACHE_RECORDS,
        )?;
        let bytes = litchi_ooxml_common::mce::process_part(records_part)?;
        let records_xml = std::str::from_utf8(bytes.as_ref())?;
        validate_pivot_cache_records(
            records_xml,
            &definition.cache_fields,
            definition.record_count,
        )?;
    }

    if let Some(relationship_id) = definition.source_relationship_id.as_deref() {
        let rel = cache_part.rels().get(relationship_id).ok_or_else(|| {
            format!(
                "pivot-cache part '{cache_uri}' references missing source worksheet relationship '{relationship_id}'"
            )
        })?;
        if !matches!(rel.reltype(), rt::WORKSHEET | rt::STRICT_WORKSHEET) {
            return Err(format!(
                "pivot-cache part '{cache_uri}' source worksheet relationship has invalid type '{}'",
                rel.reltype()
            )
            .into());
        }
        if rel.is_external() {
            return Err(format!(
                "pivot-cache part '{cache_uri}' source worksheet relationship cannot be external"
            )
            .into());
        }
        let worksheet_uri = rel.target_partname()?;
        let worksheet_part = package.get_part(&worksheet_uri)?;
        require_content_type(
            &worksheet_uri,
            worksheet_part.content_type(),
            ct::SML_WORKSHEET,
        )?;
        let workbook_name = worksheet_names_by_uri.get(&worksheet_uri).ok_or_else(|| {
            format!(
                "pivot-cache part '{cache_uri}' source worksheet '{worksheet_uri}' is not listed by the workbook"
            )
        })?;
        if let Some(source_name) = definition.source_worksheet.as_deref() {
            if source_name != workbook_name {
                return Err(format!(
                    "pivot-cache part '{cache_uri}' names source worksheet '{source_name}', but its relationship targets '{workbook_name}'"
                )
                .into());
            }
        } else {
            definition.source_worksheet = Some(workbook_name.clone());
        }
    }
    Ok(())
}

fn resolve_pivot_table_cache_uri(table_part: &dyn litchi_opc::Part) -> SheetResult<PackURI> {
    let mut matching = table_part.rels().iter().filter(|rel| {
        matches!(
            rel.reltype(),
            rt::PIVOT_CACHE_DEFINITION | rt::STRICT_PIVOT_CACHE_DEFINITION
        )
    });
    let rel = matching
        .next()
        .ok_or("pivot-table part is missing its cache-definition relationship")?;
    if matching.next().is_some() {
        return Err("pivot-table part has multiple cache-definition relationships".into());
    }
    if rel.is_external() {
        return Err("pivot-table cache-definition relationship cannot be external".into());
    }
    rel.target_partname().map_err(Into::into)
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

    fn parse(
        xml: &str,
        sheet_name: &str,
        cache_fields: Option<&[Field]>,
    ) -> SheetResult<Option<PivotTable>> {
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

        parser
            .map(|mut parser| {
                if let Some(cache_fields) = cache_fields {
                    parser.apply_cache_fields(cache_fields)?;
                }
                parser.build()
            })
            .transpose()
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
            TableContext::Root if !self.saw_pivot_fields => {
                Err("pivot table is missing its required pivotFields".into())
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

    fn apply_cache_fields(&mut self, cache_fields: &[Field]) -> SheetResult<()> {
        if self.field_names.len() != cache_fields.len() {
            return Err(format!(
                "pivot table has {} pivot fields, but its cache defines {} fields",
                self.field_names.len(),
                cache_fields.len()
            )
            .into());
        }
        self.field_names = cache_fields
            .iter()
            .map(|field| field.name.clone())
            .collect();
        Ok(())
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
    parse_pivot_table_definition_with_cache(xml, sheet_name, None)
}

fn parse_pivot_table_definition_with_cache(
    xml: &str,
    sheet_name: &str,
    cache_fields: Option<&[Field]>,
) -> SheetResult<Option<PivotTable>> {
    PivotTableParser::parse(xml, sheet_name, cache_fields)
}

pub fn read_pivot_table_definition(xml: &str) -> SheetResult<Option<PivotTable>> {
    parse_pivot_table_definition(xml, "")
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum CacheContext {
    Root,
    CacheSource,
    WorksheetSource,
    CacheFields,
    CacheField,
    Items,
    Item,
    Other,
}

struct PivotCacheParser {
    cache: Definition,
    pending_field: Option<Field>,
    expected_fields: Option<u32>,
    expected_shared_items: Option<u32>,
    saw_cache_source: bool,
    saw_worksheet_source: bool,
    saw_cache_fields: bool,
    field_saw_shared_items: bool,
}

impl PivotCacheParser {
    fn from_root(
        element: &BytesStart<'_>,
        decoder: Decoder,
        resolver: &NamespaceResolver,
    ) -> SheetResult<Self> {
        let mut cache = Definition {
            background_query: false,
            created_version: 0,
            refreshed_version: 0,
            min_refreshable_version: 0,
            ..Default::default()
        };
        cache.id = relationship_attribute_value(element, b"id", decoder, resolver)?;
        if cache.id.as_deref() == Some("") {
            return Err("pivot cache records relationship ID cannot be empty".into());
        }
        cache.invalid =
            optional_bool(element, b"invalid", decoder, "pivot cache invalid")?.unwrap_or(false);
        cache.save_data =
            optional_bool(element, b"saveData", decoder, "pivot cache saveData")?.unwrap_or(true);
        cache.refresh_on_load = optional_bool(
            element,
            b"refreshOnLoad",
            decoder,
            "pivot cache refreshOnLoad",
        )?
        .unwrap_or(false);
        cache.optimize_memory = optional_bool(
            element,
            b"optimizeMemory",
            decoder,
            "pivot cache optimizeMemory",
        )?;
        cache.enable_refresh = optional_bool(
            element,
            b"enableRefresh",
            decoder,
            "pivot cache enableRefresh",
        )?
        .unwrap_or(true);
        cache.refreshed_by = unqualified_attribute_value(element, b"refreshedBy", decoder)?;
        cache.refreshed_date = optional_f64(
            element,
            b"refreshedDate",
            decoder,
            "pivot cache refreshedDate",
        )?;
        cache.refreshed_date_iso =
            unqualified_attribute_value(element, b"refreshedDateIso", decoder)?;
        cache.background_query = optional_bool(
            element,
            b"backgroundQuery",
            decoder,
            "pivot cache backgroundQuery",
        )?
        .unwrap_or(false);
        cache.missing_items_limit = optional_u32(
            element,
            b"missingItemsLimit",
            decoder,
            "pivot cache missingItemsLimit",
        )?;
        cache.created_version = optional_u8(
            element,
            b"createdVersion",
            decoder,
            "pivot cache createdVersion",
        )?
        .unwrap_or(0);
        cache.refreshed_version = optional_u8(
            element,
            b"refreshedVersion",
            decoder,
            "pivot cache refreshedVersion",
        )?
        .unwrap_or(0);
        cache.min_refreshable_version = optional_u8(
            element,
            b"minRefreshableVersion",
            decoder,
            "pivot cache minRefreshableVersion",
        )?
        .unwrap_or(0);
        cache.record_count =
            optional_u32(element, b"recordCount", decoder, "pivot cache recordCount")?;
        cache.upgrade_on_refresh = optional_bool(
            element,
            b"upgradeOnRefresh",
            decoder,
            "pivot cache upgradeOnRefresh",
        )?;
        cache.tuples_cache =
            optional_bool(element, b"tupleCache", decoder, "pivot cache tupleCache")?;
        cache.supports_subquery = optional_bool(
            element,
            b"supportSubquery",
            decoder,
            "pivot cache supportSubquery",
        )?;
        cache.supports_advanced_drill = optional_bool(
            element,
            b"supportAdvancedDrill",
            decoder,
            "pivot cache supportAdvancedDrill",
        )?;
        Ok(Self {
            cache,
            pending_field: None,
            expected_fields: None,
            expected_shared_items: None,
            saw_cache_source: false,
            saw_worksheet_source: false,
            saw_cache_fields: false,
            field_saw_shared_items: false,
        })
    }

    fn parse(xml: &str) -> SheetResult<Option<Definition>> {
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
                        return Err("pivot cache contains multiple root elements".into());
                    }
                    if !is_spreadsheetml_name(&namespace, element.name(), b"pivotCacheDefinition") {
                        return Ok(None);
                    }
                    parser = Some(Self::from_root(&element, decoder, &resolver)?);
                    stack.push(CacheContext::Root);
                },
                Event::Empty(element) if stack.is_empty() => {
                    if !is_spreadsheetml_name(&namespace, element.name(), b"pivotCacheDefinition") {
                        return Ok(None);
                    }
                    let mut parser = Self::from_root(&element, decoder, &resolver)?;
                    parser.finish(CacheContext::Root)?;
                    return Ok(Some(parser.cache));
                },
                Event::Start(element) => {
                    let parent = *stack.last().ok_or("pivot cache is missing its root")?;
                    let context = parser
                        .as_mut()
                        .ok_or("pivot cache parser is not initialized")?
                        .start(parent, &namespace, &element, decoder, &resolver)?;
                    stack.push(context);
                },
                Event::Empty(element) => {
                    let parent = *stack.last().ok_or("pivot cache is missing its root")?;
                    let parser = parser
                        .as_mut()
                        .ok_or("pivot cache parser is not initialized")?;
                    let context = parser.start(parent, &namespace, &element, decoder, &resolver)?;
                    parser.finish(context)?;
                },
                Event::End(element) => {
                    let context = stack
                        .pop()
                        .ok_or("pivot cache has a closing element outside its root")?;
                    parser
                        .as_mut()
                        .ok_or("pivot cache parser is not initialized")?
                        .finish(context)?;
                    if context == CacheContext::Root {
                        if !is_spreadsheetml_name(
                            &namespace,
                            element.name(),
                            b"pivotCacheDefinition",
                        ) {
                            return Err("pivot cache has an invalid root closing element".into());
                        }
                        closed_root = true;
                    }
                },
                Event::Eof if parser.is_none() => return Ok(None),
                Event::Eof if !closed_root || !stack.is_empty() => {
                    return Err("pivot cache has a missing or unterminated root".into());
                },
                Event::Eof => break,
                _ => {},
            }
        }
        Ok(parser.map(|parser| parser.cache))
    }

    fn start(
        &mut self,
        parent: CacheContext,
        namespace: &ResolveResult<'_>,
        element: &BytesStart<'_>,
        decoder: Decoder,
        resolver: &NamespaceResolver,
    ) -> SheetResult<CacheContext> {
        if parent == CacheContext::Root
            && is_spreadsheetml_name(namespace, element.name(), b"cacheSource")
        {
            mark_once(&mut self.saw_cache_source, "pivot cacheSource")?;
            if self.saw_cache_fields {
                return Err("pivot cacheSource must precede cacheFields".into());
            }
            let source_type =
                required_string(element, b"type", decoder, "pivot cache source type")?;
            if !matches!(
                source_type.as_str(),
                "worksheet" | "external" | "consolidation" | "scenario"
            ) {
                return Err(format!("invalid pivot cache source type '{source_type}'").into());
            }
            self.cache.source_type = source_type;
            self.cache.source_connection_id = optional_u32(
                element,
                b"connectionId",
                decoder,
                "pivot cache source connection ID",
            )?;
            return Ok(CacheContext::CacheSource);
        }
        if parent == CacheContext::CacheSource
            && is_spreadsheetml_name(namespace, element.name(), b"worksheetSource")
        {
            mark_once(
                &mut self.saw_worksheet_source,
                "pivot cache worksheetSource",
            )?;
            if self.cache.source_type != "worksheet" {
                return Err("worksheetSource requires a worksheet pivot-cache source".into());
            }
            self.cache.source_worksheet = unqualified_attribute_value(element, b"sheet", decoder)?;
            self.cache.source_ref = unqualified_attribute_value(element, b"ref", decoder)?;
            if let Some(reference) = self.cache.source_ref.as_deref() {
                validate_cell_range(reference, "pivot cache source")?;
            }
            self.cache.source_name = unqualified_attribute_value(element, b"name", decoder)?;
            self.cache.source_relationship_id =
                relationship_attribute_value(element, b"id", decoder, resolver)?;
            if self.cache.source_relationship_id.as_deref() == Some("") {
                return Err("pivot cache source relationship ID cannot be empty".into());
            }
            return Ok(CacheContext::WorksheetSource);
        }
        if parent == CacheContext::Root
            && is_spreadsheetml_name(namespace, element.name(), b"cacheFields")
        {
            mark_once(&mut self.saw_cache_fields, "pivot cacheFields")?;
            if !self.saw_cache_source {
                return Err("pivot cacheFields must follow cacheSource".into());
            }
            self.expected_fields =
                optional_u32(element, b"count", decoder, "pivot cacheFields count")?;
            return Ok(CacheContext::CacheFields);
        }
        if parent == CacheContext::CacheFields
            && is_spreadsheetml_name(namespace, element.name(), b"cacheField")
        {
            if self.pending_field.is_some() {
                return Err("nested pivot cache field".into());
            }
            self.pending_field = Some(parse_cache_field_element(element, decoder)?);
            self.field_saw_shared_items = false;
            return Ok(CacheContext::CacheField);
        }
        if parent == CacheContext::CacheField
            && is_spreadsheetml_name(namespace, element.name(), b"sharedItems")
        {
            mark_once(&mut self.field_saw_shared_items, "pivot cache sharedItems")?;
            self.expected_shared_items =
                optional_u32(element, b"count", decoder, "pivot sharedItems count")?;
            return Ok(CacheContext::Items);
        }
        if parent == CacheContext::Items {
            let item = parse_shared_item(namespace, element, decoder)?;
            if let Some(item) = item {
                self.pending_field
                    .as_mut()
                    .ok_or("pivot shared item outside a cache field")?
                    .shared_items
                    .push(item);
                return Ok(CacheContext::Item);
            }
        }
        Ok(CacheContext::Other)
    }

    fn finish(&mut self, context: CacheContext) -> SheetResult<()> {
        match context {
            CacheContext::Root if !self.saw_cache_source => {
                Err("pivot cache is missing cacheSource".into())
            },
            CacheContext::Root if !self.saw_cache_fields => {
                Err("pivot cache is missing cacheFields".into())
            },
            CacheContext::Items => validate_count(
                self.expected_shared_items,
                self.pending_field
                    .as_ref()
                    .ok_or("missing pending pivot cache field")?
                    .shared_items
                    .len(),
                "pivot sharedItems",
            ),
            CacheContext::CacheField => {
                self.cache.cache_fields.push(
                    self.pending_field
                        .take()
                        .ok_or("missing pending pivot cache field")?,
                );
                self.expected_shared_items = None;
                Ok(())
            },
            CacheContext::CacheFields => validate_count(
                self.expected_fields,
                self.cache.cache_fields.len(),
                "pivot cacheFields",
            ),
            _ => Ok(()),
        }
    }
}

pub fn read_pivot_cache_definition(xml: &str) -> SheetResult<Option<Definition>> {
    PivotCacheParser::parse(xml)
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum CacheRecordsContext {
    Root,
    Record,
    Item,
    Other,
}

struct RecordsParser {
    records: Records,
    pending_record: Option<CacheRecord>,
    expected_records: u32,
    actual_records: usize,
    pending_value_count: usize,
    expected_field_count: Option<usize>,
    shared_item_counts: Vec<usize>,
    retain_records: bool,
}

impl RecordsParser {
    fn from_root(
        element: &BytesStart<'_>,
        decoder: Decoder,
        cache_fields: Option<&[Field]>,
        definition_record_count: Option<u32>,
        retain_records: bool,
    ) -> SheetResult<Self> {
        let expected_records =
            required_u32(element, b"count", decoder, "pivot cache records count")?;
        if let Some(definition_record_count) = definition_record_count
            && definition_record_count != expected_records
        {
            return Err(format!(
                "pivot cache definition declares {definition_record_count} records, but the records part declares {expected_records}"
            )
            .into());
        }
        Ok(Self {
            records: Records::default(),
            pending_record: None,
            expected_records,
            actual_records: 0,
            pending_value_count: 0,
            expected_field_count: cache_fields.map(<[Field]>::len),
            shared_item_counts: cache_fields
                .map(|fields| {
                    fields
                        .iter()
                        .map(|field| field.shared_items.len())
                        .collect()
                })
                .unwrap_or_default(),
            retain_records,
        })
    }

    fn parse(
        xml: &str,
        cache_fields: Option<&[Field]>,
        definition_record_count: Option<u32>,
        retain_records: bool,
    ) -> SheetResult<Option<Records>> {
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
                        return Err("pivot cache records contain multiple root elements".into());
                    }
                    if !is_spreadsheetml_name(&namespace, element.name(), b"pivotCacheRecords") {
                        return Ok(None);
                    }
                    parser = Some(Self::from_root(
                        &element,
                        decoder,
                        cache_fields,
                        definition_record_count,
                        retain_records,
                    )?);
                    stack.push(CacheRecordsContext::Root);
                },
                Event::Empty(element) if stack.is_empty() => {
                    if !is_spreadsheetml_name(&namespace, element.name(), b"pivotCacheRecords") {
                        return Ok(None);
                    }
                    let parser = Self::from_root(
                        &element,
                        decoder,
                        cache_fields,
                        definition_record_count,
                        retain_records,
                    )?;
                    parser.validate_count()?;
                    return Ok(Some(parser.records));
                },
                Event::Start(element) => {
                    let parent = *stack
                        .last()
                        .ok_or("pivot cache records are missing their root")?;
                    let context = parser
                        .as_mut()
                        .ok_or("pivot cache records parser is not initialized")?
                        .start(parent, &namespace, &element, decoder)?;
                    stack.push(context);
                },
                Event::Empty(element) => {
                    let parent = *stack
                        .last()
                        .ok_or("pivot cache records are missing their root")?;
                    let parser = parser
                        .as_mut()
                        .ok_or("pivot cache records parser is not initialized")?;
                    let context = parser.start(parent, &namespace, &element, decoder)?;
                    parser.finish(context)?;
                },
                Event::End(element) => {
                    let context = stack
                        .pop()
                        .ok_or("pivot cache records have a closing element outside their root")?;
                    parser
                        .as_mut()
                        .ok_or("pivot cache records parser is not initialized")?
                        .finish(context)?;
                    if context == CacheRecordsContext::Root {
                        if !is_spreadsheetml_name(&namespace, element.name(), b"pivotCacheRecords")
                        {
                            return Err(
                                "pivot cache records have an invalid root closing element".into()
                            );
                        }
                        closed_root = true;
                    }
                },
                Event::Eof if parser.is_none() => return Ok(None),
                Event::Eof if !closed_root || !stack.is_empty() => {
                    return Err("pivot cache records have a missing or unterminated root".into());
                },
                Event::Eof => break,
                _ => {},
            }
        }
        Ok(parser.map(|parser| parser.records))
    }

    fn start(
        &mut self,
        parent: CacheRecordsContext,
        namespace: &ResolveResult<'_>,
        element: &BytesStart<'_>,
        decoder: Decoder,
    ) -> SheetResult<CacheRecordsContext> {
        if parent == CacheRecordsContext::Root
            && is_spreadsheetml_name(namespace, element.name(), b"r")
        {
            if self.pending_record.is_some() {
                return Err("nested pivot cache record".into());
            }
            self.pending_record = Some(CacheRecord::default());
            self.pending_value_count = 0;
            return Ok(CacheRecordsContext::Record);
        }
        if parent == CacheRecordsContext::Record
            && let Some(value) = parse_record_value(namespace, element, decoder)?
        {
            if let Some(expected) = self.expected_field_count
                && self.pending_value_count >= expected
            {
                return Err(
                    format!("pivot cache record has more than {expected} field values").into(),
                );
            }
            if let Item::Index(index) = &value
                && self.expected_field_count.is_some()
            {
                let shared_item_count = self.shared_item_counts[self.pending_value_count];
                if usize::try_from(*index).map_or(true, |index| index >= shared_item_count) {
                    return Err(format!(
                        "pivot cache shared-item index {index} exceeds the {shared_item_count} items for field {}",
                        self.pending_value_count
                    )
                    .into());
                }
            }
            self.pending_value_count += 1;
            if self.retain_records {
                self.pending_record
                    .as_mut()
                    .ok_or("pivot cache value outside a record")?
                    .values
                    .push(value);
            }
            return Ok(CacheRecordsContext::Item);
        }
        Ok(CacheRecordsContext::Other)
    }

    fn finish(&mut self, context: CacheRecordsContext) -> SheetResult<()> {
        match context {
            CacheRecordsContext::Record => {
                if let Some(expected) = self.expected_field_count
                    && self.pending_value_count != expected
                {
                    return Err(format!(
                        "pivot cache record has {} values, expected {expected}",
                        self.pending_value_count
                    )
                    .into());
                }
                let record = self
                    .pending_record
                    .take()
                    .ok_or("missing pending pivot cache record")?;
                self.actual_records += 1;
                if self.retain_records {
                    self.records.records.push(record);
                }
                Ok(())
            },
            CacheRecordsContext::Root => self.validate_count(),
            _ => Ok(()),
        }
    }

    fn validate_count(&self) -> SheetResult<()> {
        validate_count(
            Some(self.expected_records),
            self.actual_records,
            "pivot cache records",
        )
    }
}

pub fn read_pivot_cache_records(xml: &str) -> SheetResult<Option<Records>> {
    RecordsParser::parse(xml, None, None, true)
}

fn validate_pivot_cache_records(
    xml: &str,
    cache_fields: &[Field],
    definition_record_count: Option<u32>,
) -> SheetResult<()> {
    RecordsParser::parse(xml, Some(cache_fields), definition_record_count, false)?
        .ok_or("pivot-cache records part has no pivotCacheRecords root")?;
    Ok(())
}

fn parse_cache_field_element(
    element: &BytesStart<'_>,
    decoder: Decoder,
) -> SheetResult<Field> {
    let mut field = Field {
        name: required_string(element, b"name", decoder, "pivot cache field name")?,
        ..Default::default()
    };
    field.caption = unqualified_attribute_value(element, b"caption", decoder)?;
    field.property_name = unqualified_attribute_value(element, b"propertyName", decoder)?;
    field.server_field =
        optional_bool(element, b"serverField", decoder, "pivot cache serverField")?;
    field.unique_list =
        optional_bool(element, b"uniqueList", decoder, "pivot cache uniqueList")?.unwrap_or(true);
    field.num_fmt_id = optional_u32(
        element,
        b"numFmtId",
        decoder,
        "pivot cache number format ID",
    )?;
    field.formula = unqualified_attribute_value(element, b"formula", decoder)?;
    field.sql_type = optional_i32(element, b"sqlType", decoder, "pivot cache SQL type")?;
    field.hierarchy = optional_i32(element, b"hierarchy", decoder, "pivot cache hierarchy")?;
    field.level = optional_u32(element, b"level", decoder, "pivot cache level")?;
    field.mapping_count = optional_u32(
        element,
        b"mappingCount",
        decoder,
        "pivot cache mapping count",
    )?;
    field.database_field = optional_bool(
        element,
        b"databaseField",
        decoder,
        "pivot cache databaseField",
    )?
    .unwrap_or(true);
    field.member_property_field = optional_bool(
        element,
        b"memberPropertyField",
        decoder,
        "pivot cache memberPropertyField",
    )?;
    Ok(field)
}

fn parse_shared_item(
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    decoder: Decoder,
) -> SheetResult<Option<Item>> {
    if is_spreadsheetml_name(namespace, element.name(), b"m") {
        return Ok(Some(Item::Missing));
    }
    if is_spreadsheetml_name(namespace, element.name(), b"n") {
        let value = required_f64(element, b"v", decoder, "pivot shared number")?;
        return Ok(Some(Item::Number(value)));
    }
    if is_spreadsheetml_name(namespace, element.name(), b"b") {
        let value = required_bool(element, b"v", decoder, "pivot shared boolean")?;
        return Ok(Some(Item::Boolean(value)));
    }
    if is_spreadsheetml_name(namespace, element.name(), b"e") {
        return Ok(Some(Item::Error(required_string(
            element,
            b"v",
            decoder,
            "pivot shared error",
        )?)));
    }
    if is_spreadsheetml_name(namespace, element.name(), b"s") {
        return Ok(Some(Item::String(required_string(
            element,
            b"v",
            decoder,
            "pivot shared string",
        )?)));
    }
    if is_spreadsheetml_name(namespace, element.name(), b"d") {
        return Ok(Some(Item::DateTime(required_string(
            element,
            b"v",
            decoder,
            "pivot shared date-time",
        )?)));
    }
    Ok(None)
}

fn parse_record_value(
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    decoder: Decoder,
) -> SheetResult<Option<Item>> {
    if is_spreadsheetml_name(namespace, element.name(), b"x") {
        return Ok(Some(Item::Index(required_u32(
            element,
            b"v",
            decoder,
            "pivot cache shared-item index",
        )?)));
    }
    parse_shared_item(namespace, element, decoder)
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

fn optional_i32(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    description: &str,
) -> SheetResult<Option<i32>> {
    unqualified_attribute_value(element, name, decoder)?
        .map(|value| {
            value
                .parse::<i32>()
                .map_err(|_| format!("invalid {description} '{value}'").into())
        })
        .transpose()
}

fn optional_u8(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    description: &str,
) -> SheetResult<Option<u8>> {
    unqualified_attribute_value(element, name, decoder)?
        .map(|value| {
            value
                .parse::<u8>()
                .map_err(|_| format!("invalid {description} '{value}'").into())
        })
        .transpose()
}

fn optional_bool(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    description: &str,
) -> SheetResult<Option<bool>> {
    unqualified_attribute_value(element, name, decoder)?
        .map(|value| match value.as_str() {
            "1" | "true" => Ok(true),
            "0" | "false" => Ok(false),
            _ => Err(format!("invalid {description} '{value}'").into()),
        })
        .transpose()
}

fn required_bool(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    description: &str,
) -> SheetResult<bool> {
    optional_bool(element, name, decoder, description)?
        .ok_or_else(|| format!("missing {description} attribute").into())
}

fn optional_f64(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    description: &str,
) -> SheetResult<Option<f64>> {
    unqualified_attribute_value(element, name, decoder)?
        .map(|value| {
            let parsed = value
                .parse::<f64>()
                .map_err(|_| format!("invalid {description} '{value}'"))?;
            if !parsed.is_finite() {
                return Err(format!("{description} must be finite").into());
            }
            Ok(parsed)
        })
        .transpose()
}

fn required_f64(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    description: &str,
) -> SheetResult<f64> {
    optional_f64(element, name, decoder, description)?
        .ok_or_else(|| format!("missing {description} attribute").into())
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
    Rect::from_a1(&first.replace('$', ""))?;
    if let Some(second) = references.next() {
        Rect::from_a1(&second.replace('$', ""))?;
    }
    if references.next().is_some() {
        return Err(format!("invalid {description} range '{range}'").into());
    }
    Ok(())
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
        let source_uri = PackURI::new("/custom/sheets/source.xml").unwrap();
        let table_uri = PackURI::new("/custom/pivots/table.xml").unwrap();
        let cache_uri = PackURI::new("/custom/cache/cache.xml").unwrap();
        let records_uri = PackURI::new("/custom/cache/records.xml").unwrap();
        let mut workbook_part = BlobPart::new(
            workbook_uri,
            ct::SML_SHEET_MAIN.to_string(),
            br#"<workbook xmlns="http://purl.oclc.org/ooxml/spreadsheetml/main"
                    xmlns:r="http://purl.oclc.org/ooxml/officeDocument/relationships">
                    <sheets><sheet name="Pivot" sheetId="1" r:id="rId1"/>
                        <sheet name="Source" sheetId="2" r:id="rId2"/></sheets>
                    <pivotCaches><pivotCache cacheId="7" r:id="rId3"/></pivotCaches>
                </workbook>"#
                .to_vec(),
        );
        workbook_part.relate_to("sheets/data.xml", rt::STRICT_WORKSHEET);
        workbook_part.relate_to("sheets/source.xml", rt::STRICT_WORKSHEET);
        workbook_part.relate_to("cache/cache.xml", rt::STRICT_PIVOT_CACHE_DEFINITION);
        let mut worksheet_part = BlobPart::new(
            worksheet_uri.clone(),
            ct::SML_WORKSHEET.to_string(),
            Vec::new(),
        );
        worksheet_part.relate_to("../pivots/table.xml", rt::STRICT_PIVOT_TABLE);
        let mut cache_part = BlobPart::new(
            cache_uri,
            ct::SML_PIVOT_CACHE_DEFINITION.to_string(),
            br#"<pivotCacheDefinition xmlns="http://purl.oclc.org/ooxml/spreadsheetml/main"
                    xmlns:r="http://purl.oclc.org/ooxml/officeDocument/relationships"
                    r:id="rId1" recordCount="2">
                    <cacheSource type="worksheet"><worksheetSource ref="$A$1:$B$3" r:id="rId2"/></cacheSource>
                    <cacheFields count="1"><cacheField name="Cache Region"/></cacheFields>
                </pivotCacheDefinition>"#
                .to_vec(),
        );
        cache_part.relate_to("records.xml", rt::STRICT_PIVOT_CACHE_RECORDS);
        cache_part.relate_to("../sheets/source.xml", rt::STRICT_WORKSHEET);
        let mut table_part = BlobPart::new(
            table_uri,
            ct::SML_PIVOT_TABLE.to_string(),
            br#"<pivotTableDefinition xmlns="http://purl.oclc.org/ooxml/spreadsheetml/main"
                    name="PivotOne" cacheId="7" dataCaption="Values">
                    <location ref="A1:C5" firstHeaderRow="1" firstDataRow="2" firstDataCol="1"/>
                    <pivotFields count="1"><pivotField/></pivotFields>
                    <rowFields count="1"><field x="0"/></rowFields>
                </pivotTableDefinition>"#
                .to_vec(),
        );
        table_part.relate_to("../cache/cache.xml", rt::STRICT_PIVOT_CACHE_DEFINITION);
        package.relate_to("custom/book.xml", rt::STRICT_OFFICE_DOCUMENT);
        package.add_part(Box::new(workbook_part));
        package.add_part(Box::new(worksheet_part));
        package.add_part(Box::new(BlobPart::new(
            source_uri,
            ct::SML_WORKSHEET.to_string(),
            Vec::new(),
        )));
        package.add_part(Box::new(cache_part));
        package.add_part(Box::new(BlobPart::new(
            records_uri,
            ct::SML_PIVOT_CACHE_RECORDS.to_string(),
            br#"<pivotCacheRecords xmlns="http://purl.oclc.org/ooxml/spreadsheetml/main" count="2">
                    <r><s v="North"/></r><r><s v="South"/></r>
                </pivotCacheRecords>"#
                .to_vec(),
        )));
        package.add_part(Box::new(table_part));
        (package, worksheet_uri)
    }

    #[test]
    fn resolves_strict_custom_pivot_table_parts() {
        let (package, _) = package_with_pivot_table();
        let tables = read_pivot_tables(&package).unwrap();

        assert_eq!(tables.len(), 1);
        assert_eq!(tables[0].name, "PivotOne");
        assert_eq!(tables[0].sheet_name, "Pivot");
        assert_eq!(tables[0].source_sheet.as_deref(), Some("Source"));
        assert_eq!(tables[0].source_ref.as_deref(), Some("$A$1:$B$3"));
        assert_eq!(tables[0].field_names, ["Cache Region"]);
        assert_eq!(tables[0].row_fields[0].field_name, "Cache Region");
    }

    #[test]
    fn tolerates_chartsheet_entries_in_sheet_walk() {
        let (mut package, _) = package_with_pivot_table();
        let workbook_uri = PackURI::new("/custom/book.xml").unwrap();
        let workbook_part = package.get_part_mut(&workbook_uri).unwrap();
        let updated = std::str::from_utf8(workbook_part.blob()).unwrap().replace(
            "</sheets>",
            r#"<sheet name="Chart1" sheetId="3" r:id="rId4"/></sheets>"#,
        );
        workbook_part.set_blob(updated.into_bytes());
        workbook_part.relate_to(
            "chartsheets/chart1.xml",
            "http://purl.oclc.org/ooxml/officeDocument/relationships/chartsheet",
        );
        package.add_part(Box::new(BlobPart::new(
            PackURI::new("/custom/chartsheets/chart1.xml").unwrap(),
            "application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml"
                .to_string(),
            Vec::new(),
        )));

        let tables = read_pivot_tables(&package).unwrap();
        assert_eq!(tables.len(), 1);
        assert_eq!(tables[0].name, "PivotOne");
    }

    #[test]
    fn poi_fixture_with_chartsheet_reads_pivot_tables() {
        const POI_CHARTSHEET: &[u8] = include_bytes!(
            "../../../../test-data/poi/test-data/spreadsheet/WithChartSheet.xlsx"
        );
        let package = OpcPackage::from_bytes(POI_CHARTSHEET).unwrap();
        let tables = read_pivot_tables(&package).unwrap();
        assert_eq!(tables.len(), 5);
        assert!(
            tables
                .iter()
                .any(|table| table.name == "PivotTable2" && table.sheet_name == "Sheet2")
        );
    }

    #[test]
    fn rejects_external_and_wrong_content_type_pivot_parts() {
        let (mut package, worksheet_uri) = package_with_pivot_table();
        let relationships = package.get_part_mut(&worksheet_uri).unwrap().rels_mut();
        relationships.remove("rId1").unwrap();
        relationships.add_relationship(
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
    fn rejects_invalid_pivot_cache_relationship_graphs() {
        let workbook_uri = PackURI::new("/custom/book.xml").unwrap();
        let table_uri = PackURI::new("/custom/pivots/table.xml").unwrap();
        let cache_uri = PackURI::new("/custom/cache/cache.xml").unwrap();
        let records_uri = PackURI::new("/custom/cache/records.xml").unwrap();

        let (mut package, _) = package_with_pivot_table();
        let relationships = package.get_part_mut(&workbook_uri).unwrap().rels_mut();
        relationships.remove("rId3").unwrap();
        relationships.add_relationship(
            rt::STRICT_PIVOT_CACHE_DEFINITION.to_string(),
            "https://example.com/cache.xml".to_string(),
            "rId3".to_string(),
            true,
        );
        assert!(read_pivot_tables(&package).is_err());

        let (mut package, _) = package_with_pivot_table();
        package
            .get_part_mut(&table_uri)
            .unwrap()
            .rels_mut()
            .add_relationship(
                rt::STRICT_PIVOT_CACHE_DEFINITION.to_string(),
                "../cache/duplicate.xml".to_string(),
                "duplicate".to_string(),
                false,
            );
        assert!(read_pivot_tables(&package).is_err());

        let (mut package, _) = package_with_pivot_table();
        let table_part = package.get_part_mut(&table_uri).unwrap();
        let changed = std::str::from_utf8(table_part.blob())
            .unwrap()
            .replace("cacheId=\"7\"", "cacheId=\"8\"");
        table_part.set_blob(changed.into_bytes());
        assert!(read_pivot_tables(&package).is_err());

        let (mut package, _) = package_with_pivot_table();
        package.add_part(Box::new(BlobPart::new(
            records_uri.clone(),
            ct::SML_WORKSHEET.to_string(),
            Vec::new(),
        )));
        assert!(read_pivot_tables(&package).is_err());

        let (mut package, _) = package_with_pivot_table();
        package.add_part(Box::new(BlobPart::new(
            records_uri,
            ct::SML_PIVOT_CACHE_RECORDS.to_string(),
            br#"<pivotCacheRecords xmlns="http://purl.oclc.org/ooxml/spreadsheetml/main" count="2">
                    <r><x v="0"/></r><r><s v="South"/></r>
                </pivotCacheRecords>"#
                .to_vec(),
        )));
        assert!(read_pivot_tables(&package).is_err());

        let (mut package, _) = package_with_pivot_table();
        let relationships = package.get_part_mut(&cache_uri).unwrap().rels_mut();
        relationships.remove("rId2").unwrap();
        relationships.add_relationship(
            rt::STRICT_WORKSHEET.to_string(),
            "https://example.com/source.xml".to_string(),
            "rId2".to_string(),
            true,
        );
        assert!(read_pivot_tables(&package).is_err());

        let (mut package, _) = package_with_pivot_table();
        let workbook_part = package.get_part_mut(&workbook_uri).unwrap();
        let changed = std::str::from_utf8(workbook_part.blob()).unwrap().replace(
            r#"<pivotCaches><pivotCache cacheId="7" r:id="rId3"/></pivotCaches>"#,
            "",
        );
        workbook_part.set_blob(changed.into_bytes());
        assert!(read_pivot_tables(&package).is_err());

        let (mut package, _) = package_with_pivot_table();
        package.add_part(Box::new(BlobPart::new(
            cache_uri,
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

    #[test]
    fn parses_prefixed_pivot_cache_definition_and_shared_items() {
        let xml = r##"<p:pivotCacheDefinition
                xmlns:p="http://purl.oclc.org/ooxml/spreadsheetml/main"
                xmlns:r="http://purl.oclc.org/ooxml/officeDocument/relationships"
                xmlns:f="urn:foreign" r:id="records" invalid="true" saveData="0"
                refreshOnLoad="1" optimizeMemory="false" enableRefresh="0"
                refreshedBy="Alice &amp; Bob" refreshedDate="42.5" backgroundQuery="true"
                missingItemsLimit="10" createdVersion="7" recordCount="6"
                upgradeOnRefresh="1" tupleCache="0" supportSubquery="true">
                <f:cacheSource type="worksheet"><p:worksheetSource ref="XFE1"/></f:cacheSource>
                <p:cacheSource type="worksheet" connectionId="8"><p:worksheetSource
                    sheet="Data &amp; More" ref="$A$1:$B$4" r:id="source-sheet"/></p:cacheSource>
                <p:cacheFields count="2">
                    <p:cacheField name="Region" caption="Area" databaseField="false"
                            uniqueList="0" numFmtId="4" formula="x" sqlType="-1"
                            hierarchy="2" level="3" mappingCount="4" memberPropertyField="true">
                        <p:sharedItems count="6"><p:m/><p:n v="2.5"/><p:b v="true"/>
                            <p:e v="#N/A"/><p:s v="North &amp; West"/><p:d v="2026-07-14T00:00:00Z"/>
                        </p:sharedItems>
                    </p:cacheField>
                    <p:cacheField name="Sales"/>
                </p:cacheFields>
            </p:pivotCacheDefinition>"##;
        let cache = read_pivot_cache_definition(xml).unwrap().unwrap();

        assert_eq!(cache.id.as_deref(), Some("records"));
        assert!(cache.invalid);
        assert!(!cache.save_data);
        assert_eq!(cache.refreshed_by.as_deref(), Some("Alice & Bob"));
        assert_eq!(cache.source_worksheet.as_deref(), Some("Data & More"));
        assert_eq!(cache.source_ref.as_deref(), Some("$A$1:$B$4"));
        assert_eq!(cache.source_connection_id, Some(8));
        assert_eq!(
            cache.source_relationship_id.as_deref(),
            Some("source-sheet")
        );
        assert_eq!(cache.cache_fields.len(), 2);
        let field = &cache.cache_fields[0];
        assert!(!field.database_field);
        assert_eq!(field.sql_type, Some(-1));
        assert_eq!(field.mapping_count, Some(4));
        assert_eq!(field.member_property_field, Some(true));
        assert_eq!(field.shared_items.len(), 6);
        assert!(matches!(field.shared_items[0], Item::Missing));
        assert!(matches!(field.shared_items[1], Item::Number(2.5)));
        assert!(matches!(field.shared_items[2], Item::Boolean(true)));
        assert!(
            matches!(&field.shared_items[4], Item::String(value) if value == "North & West")
        );
    }

    #[test]
    fn rejects_malformed_pivot_cache_definitions() {
        const S: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        let valid_source = r#"<cacheSource type="worksheet"><worksheetSource sheet="Data" ref="A1:B2"/></cacheSource>"#;
        let invalid = [
            format!(r#"<pivotCacheDefinition xmlns="{S}"><cacheFields/></pivotCacheDefinition>"#),
            format!(
                r#"<pivotCacheDefinition xmlns="{S}" invalid="yes">{valid_source}<cacheFields/></pivotCacheDefinition>"#
            ),
            format!(
                r#"<pivotCacheDefinition xmlns="{S}"><cacheSource type="bad"/><cacheFields/></pivotCacheDefinition>"#
            ),
            format!(
                r#"<pivotCacheDefinition xmlns="{S}">{valid_source}<cacheFields count="2"><cacheField name="One"/></cacheFields></pivotCacheDefinition>"#
            ),
            format!(
                r#"<pivotCacheDefinition xmlns="{S}">{valid_source}<cacheSource type="worksheet"/><cacheFields/></pivotCacheDefinition>"#
            ),
            format!(
                r#"<pivotCacheDefinition xmlns="{S}"><cacheFields/>{valid_source}</pivotCacheDefinition>"#
            ),
            format!(
                r#"<pivotCacheDefinition xmlns="{S}">{valid_source}<cacheFields><cacheField/></cacheFields></pivotCacheDefinition>"#
            ),
            format!(
                r#"<pivotCacheDefinition xmlns="{S}">{valid_source}<cacheFields><cacheField name="One"><sharedItems><n v="NaN"/></sharedItems></cacheField></cacheFields></pivotCacheDefinition>"#
            ),
            format!(
                r#"<pivotCacheDefinition xmlns="{S}">{valid_source}<cacheFields><cacheField name="One"><sharedItems count="1"/></cacheField></cacheFields></pivotCacheDefinition>"#
            ),
            format!(
                r#"<pivotCacheDefinition xmlns="{S}" xmlns:r="{R}" r:id="">{valid_source}<cacheFields/></pivotCacheDefinition>"#
            ),
            format!(
                r#"<pivotCacheDefinition xmlns="{S}" xmlns:r="{R}"><cacheSource type="worksheet"><worksheetSource r:id=""/></cacheSource><cacheFields/></pivotCacheDefinition>"#
            ),
        ];
        for xml in invalid {
            assert!(read_pivot_cache_definition(&xml).is_err(), "accepted {xml}");
        }
        assert!(
            read_pivot_cache_definition(r#"<pivotCacheDefinition xmlns="urn:foreign"/>"#)
                .unwrap()
                .is_none()
        );
    }

    #[test]
    fn parses_prefixed_pivot_cache_records() {
        let xml = r##"<p:pivotCacheRecords
                xmlns:p="http://purl.oclc.org/ooxml/spreadsheetml/main"
                xmlns:f="urn:foreign" count="2">
                <f:r><p:n v="99"/></f:r>
                <p:r><p:x v="3"/><p:m/><p:n v="2.5"/><p:b v="false"/>
                    <p:e v="#N/A"/><p:s v="North &amp; West"/><p:d v="2026-07-14T00:00:00Z"/></p:r>
                <p:r/>
            </p:pivotCacheRecords>"##;
        let records = read_pivot_cache_records(xml).unwrap().unwrap();

        assert_eq!(records.records.len(), 2);
        let values = &records.records[0].values;
        assert_eq!(values.len(), 7);
        assert!(matches!(values[0], Item::Index(3)));
        assert!(matches!(values[1], Item::Missing));
        assert!(matches!(values[2], Item::Number(2.5)));
        assert!(matches!(values[3], Item::Boolean(false)));
        assert!(matches!(&values[5], Item::String(value) if value == "North & West"));
        assert!(records.records[1].values.is_empty());
    }

    #[test]
    fn rejects_malformed_pivot_cache_records() {
        const S: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        for xml in [
            format!(r#"<pivotCacheRecords xmlns="{S}"/>"#),
            format!(r#"<pivotCacheRecords xmlns="{S}" count="2"><r/></pivotCacheRecords>"#),
            format!(r#"<pivotCacheRecords xmlns="{S}" count="1"><r><x/></r></pivotCacheRecords>"#),
            format!(
                r#"<pivotCacheRecords xmlns="{S}" count="1"><r><b v="yes"/></r></pivotCacheRecords>"#
            ),
            format!(
                r#"<pivotCacheRecords xmlns="{S}" count="1"><r><n v="NaN"/></r></pivotCacheRecords>"#
            ),
        ] {
            assert!(read_pivot_cache_records(&xml).is_err(), "accepted {xml}");
        }
        assert!(
            read_pivot_cache_records(r#"<pivotCacheRecords xmlns="urn:foreign" count="0"/>"#)
                .unwrap()
                .is_none()
        );
    }
}
