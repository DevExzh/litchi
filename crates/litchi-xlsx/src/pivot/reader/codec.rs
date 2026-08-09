//! `SpreadsheetML` XML codecs for pivot tables, cache definitions, and records.

use super::super::cache::{CacheRecord, Definition, Field, Item, Records};
use super::super::{PivotAxis, PivotDataField, PivotFieldRole, PivotTable, PivotValueFunction};
use super::model::{
    CacheContext, CacheRecordsContext, PivotCacheParser, PivotTableParser, RawDataField,
    RecordsParser, TableContext,
};
use super::validation::{
    mark_once, optional_bool, optional_f64, optional_i32, optional_u8, optional_u32, required_bool,
    required_f64, required_i32, required_string, required_u32, validate_cell_range, validate_count,
};
use crate::raw::namespace::{is_spreadsheetml_name, relationship_attribute_value};
use litchi_core::sheet::Result as SheetResult;
use litchi_ooxml_common::xml::unqualified_attribute_value;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{NamespaceResolver, ResolveResult};
use quick_xml::reader::NsReader;

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

pub(super) fn parse_pivot_table_definition_with_cache(
    xml: &str,
    sheet_name: &str,
    cache_fields: Option<&[Field]>,
) -> SheetResult<Option<PivotTable>> {
    PivotTableParser::parse(xml, sheet_name, cache_fields)
}

pub fn read_pivot_table_definition(xml: &str) -> SheetResult<Option<PivotTable>> {
    parse_pivot_table_definition(xml, "")
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

pub(super) fn validate_pivot_cache_records(
    xml: &str,
    cache_fields: &[Field],
    definition_record_count: Option<u32>,
) -> SheetResult<()> {
    RecordsParser::parse(xml, Some(cache_fields), definition_record_count, false)?
        .ok_or("pivot-cache records part has no pivotCacheRecords root")?;
    Ok(())
}

fn parse_cache_field_element(element: &BytesStart<'_>, decoder: Decoder) -> SheetResult<Field> {
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
            .unwrap_or_else(|| format!("Field{idx}"));
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
