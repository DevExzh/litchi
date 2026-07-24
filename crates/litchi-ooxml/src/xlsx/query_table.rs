//! Typed SpreadsheetML query-table metadata.
//!
//! Query-table parts describe how already-imported worksheet data is refreshed.
//! This module deliberately treats the referenced connection as an inert numeric
//! identifier: it never opens a URL, runs a command, or evaluates a formula.

use crate::common::mce::process_ooxml;
use crate::error::{OoxmlError, Result};
use quick_xml::events::Event;
use quick_xml::name::ResolveResult;
use quick_xml::reader::NsReader;
use std::collections::HashSet;
use litchi_opc::part::BlobPart;
use litchi_opc::{OpcPackage, PackURI};

pub const QUERY_TABLE_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.queryTable+xml";
pub const QUERY_TABLE_RELATIONSHIP_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/queryTable";
pub const STRICT_QUERY_TABLE_RELATIONSHIP_TYPE: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/queryTable";

const TRANSITIONAL: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const STRICT: &str = "http://purl.oclc.org/ooxml/spreadsheetml/main";
const MAX_PART_BYTES: usize = 8 * 1024 * 1024;
const MAX_TEXT_BYTES: usize = 1024 * 1024;
const MAX_NODES: usize = 20_000;
const MAX_DEPTH: usize = 64;
const MAX_FIELDS: usize = 16_384;
const MAX_DELETED_FIELDS: usize = 16_384;
const MAX_EXTENSIONS: usize = 1_024;
const MAX_SORT_CONDITIONS: usize = 64;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum QueryTableConformance {
    Transitional,
    Strict,
}

impl QueryTableConformance {
    fn namespace(self) -> &'static str {
        match self {
            Self::Transitional => TRANSITIONAL,
            Self::Strict => STRICT,
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum QueryTableGrowShrinkType {
    InsertDelete,
    InsertClear,
    OverwriteClear,
}

impl QueryTableGrowShrinkType {
    fn parse(value: &str) -> Result<Self> {
        match value {
            "insertDelete" => Ok(Self::InsertDelete),
            "insertClear" => Ok(Self::InsertClear),
            "overwriteClear" => Ok(Self::OverwriteClear),
            _ => Err(invalid(format!("invalid growShrinkType '{value}'"))),
        }
    }

    fn as_str(self) -> &'static str {
        match self {
            Self::InsertDelete => "insertDelete",
            Self::InsertClear => "insertClear",
            Self::OverwriteClear => "overwriteClear",
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum QueryTableSortMethod {
    Stroke,
    PinYin,
    None,
}

impl QueryTableSortMethod {
    fn parse(value: &str) -> Result<Self> {
        match value {
            "stroke" => Ok(Self::Stroke),
            "pinYin" => Ok(Self::PinYin),
            "none" => Ok(Self::None),
            _ => Err(invalid(format!("invalid sortMethod '{value}'"))),
        }
    }

    fn as_str(self) -> &'static str {
        match self {
            Self::Stroke => "stroke",
            Self::PinYin => "pinYin",
            Self::None => "none",
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum QueryTableSortBy {
    Value,
    CellColor,
    FontColor,
    Icon,
}

impl QueryTableSortBy {
    fn parse(value: &str) -> Result<Self> {
        match value {
            "value" => Ok(Self::Value),
            "cellColor" => Ok(Self::CellColor),
            "fontColor" => Ok(Self::FontColor),
            "icon" => Ok(Self::Icon),
            _ => Err(invalid(format!("invalid sortBy '{value}'"))),
        }
    }

    fn as_str(self) -> &'static str {
        match self {
            Self::Value => "value",
            Self::CellColor => "cellColor",
            Self::FontColor => "fontColor",
            Self::Icon => "icon",
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum QueryTableIconSet {
    ThreeArrows,
    ThreeArrowsGray,
    ThreeFlags,
    ThreeTrafficLights1,
    ThreeTrafficLights2,
    ThreeSigns,
    ThreeSymbols,
    ThreeSymbols2,
    FourArrows,
    FourArrowsGray,
    FourRedToBlack,
    FourRating,
    FourTrafficLights,
    FiveArrows,
    FiveArrowsGray,
    FiveRating,
    FiveQuarters,
}

impl QueryTableIconSet {
    fn parse(value: &str) -> Result<Self> {
        match value {
            "3Arrows" => Ok(Self::ThreeArrows),
            "3ArrowsGray" => Ok(Self::ThreeArrowsGray),
            "3Flags" => Ok(Self::ThreeFlags),
            "3TrafficLights1" => Ok(Self::ThreeTrafficLights1),
            "3TrafficLights2" => Ok(Self::ThreeTrafficLights2),
            "3Signs" => Ok(Self::ThreeSigns),
            "3Symbols" => Ok(Self::ThreeSymbols),
            "3Symbols2" => Ok(Self::ThreeSymbols2),
            "4Arrows" => Ok(Self::FourArrows),
            "4ArrowsGray" => Ok(Self::FourArrowsGray),
            "4RedToBlack" => Ok(Self::FourRedToBlack),
            "4Rating" => Ok(Self::FourRating),
            "4TrafficLights" => Ok(Self::FourTrafficLights),
            "5Arrows" => Ok(Self::FiveArrows),
            "5ArrowsGray" => Ok(Self::FiveArrowsGray),
            "5Rating" => Ok(Self::FiveRating),
            "5Quarters" => Ok(Self::FiveQuarters),
            _ => Err(invalid(format!("invalid iconSet '{value}'"))),
        }
    }

    fn as_str(self) -> &'static str {
        match self {
            Self::ThreeArrows => "3Arrows",
            Self::ThreeArrowsGray => "3ArrowsGray",
            Self::ThreeFlags => "3Flags",
            Self::ThreeTrafficLights1 => "3TrafficLights1",
            Self::ThreeTrafficLights2 => "3TrafficLights2",
            Self::ThreeSigns => "3Signs",
            Self::ThreeSymbols => "3Symbols",
            Self::ThreeSymbols2 => "3Symbols2",
            Self::FourArrows => "4Arrows",
            Self::FourArrowsGray => "4ArrowsGray",
            Self::FourRedToBlack => "4RedToBlack",
            Self::FourRating => "4Rating",
            Self::FourTrafficLights => "4TrafficLights",
            Self::FiveArrows => "5Arrows",
            Self::FiveArrowsGray => "5ArrowsGray",
            Self::FiveRating => "5Rating",
            Self::FiveQuarters => "5Quarters",
        }
    }

    fn cardinality(self) -> u32 {
        match self {
            Self::ThreeArrows
            | Self::ThreeArrowsGray
            | Self::ThreeFlags
            | Self::ThreeTrafficLights1
            | Self::ThreeTrafficLights2
            | Self::ThreeSigns
            | Self::ThreeSymbols
            | Self::ThreeSymbols2 => 3,
            Self::FourArrows
            | Self::FourArrowsGray
            | Self::FourRedToBlack
            | Self::FourRating
            | Self::FourTrafficLights => 4,
            _ => 5,
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct QueryTableExtensionAttribute {
    qualified_name: String,
    value: String,
}

impl QueryTableExtensionAttribute {
    pub fn qualified_name(&self) -> &str {
        &self.qualified_name
    }

    pub fn value(&self) -> &str {
        &self.value
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct QueryTableExtensionList {
    namespaces: Vec<QueryTableExtensionAttribute>,
    attributes: Vec<QueryTableExtensionAttribute>,
    extensions: Vec<QueryTableExtension>,
}

impl QueryTableExtensionList {
    pub fn extension_uris(&self) -> impl Iterator<Item = &str> {
        self.extensions
            .iter()
            .map(|extension| extension.uri.as_str())
    }

    pub fn len(&self) -> usize {
        self.extensions.len()
    }

    pub fn is_empty(&self) -> bool {
        self.extensions.is_empty()
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct QueryTableExtension {
    uri: String,
    namespaces: Vec<QueryTableExtensionAttribute>,
    attributes: Vec<QueryTableExtensionAttribute>,
    content: Vec<XmlNode>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct QueryTableSortCondition {
    reference: String,
    descending: Option<bool>,
    sort_by: Option<QueryTableSortBy>,
    custom_list: Option<String>,
    differential_format_id: Option<u32>,
    icon_set: Option<QueryTableIconSet>,
    icon_id: Option<u32>,
    extension_attributes: Vec<QueryTableExtensionAttribute>,
}

impl QueryTableSortCondition {
    pub fn reference(&self) -> &str {
        &self.reference
    }
    pub fn descending(&self) -> Option<bool> {
        self.descending
    }
    pub fn sort_by(&self) -> Option<QueryTableSortBy> {
        self.sort_by
    }
    pub fn custom_list(&self) -> Option<&str> {
        self.custom_list.as_deref()
    }
    pub fn differential_format_id(&self) -> Option<u32> {
        self.differential_format_id
    }
    pub fn icon_set(&self) -> Option<QueryTableIconSet> {
        self.icon_set
    }
    pub fn icon_id(&self) -> Option<u32> {
        self.icon_id
    }
    pub fn extension_attributes(&self) -> &[QueryTableExtensionAttribute] {
        &self.extension_attributes
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct QueryTableSortState {
    reference: String,
    column_sort: Option<bool>,
    case_sensitive: Option<bool>,
    sort_method: Option<QueryTableSortMethod>,
    conditions: Vec<QueryTableSortCondition>,
    extension_attributes: Vec<QueryTableExtensionAttribute>,
}

impl QueryTableSortState {
    pub fn reference(&self) -> &str {
        &self.reference
    }
    pub fn column_sort(&self) -> Option<bool> {
        self.column_sort
    }
    pub fn case_sensitive(&self) -> Option<bool> {
        self.case_sensitive
    }
    pub fn sort_method(&self) -> Option<QueryTableSortMethod> {
        self.sort_method
    }
    pub fn conditions(&self) -> &[QueryTableSortCondition] {
        &self.conditions
    }
    pub fn extension_attributes(&self) -> &[QueryTableExtensionAttribute] {
        &self.extension_attributes
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct QueryTableField {
    id: u32,
    name: Option<String>,
    data_bound: Option<bool>,
    row_numbers: Option<bool>,
    fill_formulas: Option<bool>,
    clipped: Option<bool>,
    table_column_id: Option<u32>,
    extension_list: Option<QueryTableExtensionList>,
    extension_attributes: Vec<QueryTableExtensionAttribute>,
}

impl QueryTableField {
    pub fn new(id: u32) -> Self {
        Self {
            id,
            name: None,
            data_bound: None,
            row_numbers: None,
            fill_formulas: None,
            clipped: None,
            table_column_id: None,
            extension_list: None,
            extension_attributes: Vec::new(),
        }
    }
    pub fn set_name(&mut self, value: Option<String>) { self.name = value; }
    pub fn set_data_bound(&mut self, value: Option<bool>) { self.data_bound = value; }
    pub fn set_row_numbers(&mut self, value: Option<bool>) { self.row_numbers = value; }
    pub fn set_fill_formulas(&mut self, value: Option<bool>) { self.fill_formulas = value; }
    pub fn set_clipped(&mut self, value: Option<bool>) { self.clipped = value; }
    pub fn set_table_column_id(&mut self, value: Option<u32>) { self.table_column_id = value; }
    pub fn id(&self) -> u32 {
        self.id
    }
    pub fn name(&self) -> Option<&str> {
        self.name.as_deref()
    }
    pub fn data_bound(&self) -> Option<bool> {
        self.data_bound
    }
    pub fn row_numbers(&self) -> Option<bool> {
        self.row_numbers
    }
    pub fn fill_formulas(&self) -> Option<bool> {
        self.fill_formulas
    }
    pub fn clipped(&self) -> Option<bool> {
        self.clipped
    }
    pub fn table_column_id(&self) -> Option<u32> {
        self.table_column_id
    }
    pub fn extension_list(&self) -> Option<&QueryTableExtensionList> {
        self.extension_list.as_ref()
    }
    pub fn extension_attributes(&self) -> &[QueryTableExtensionAttribute] {
        &self.extension_attributes
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct QueryTableRefresh {
    preserve_sort_filter_layout: Option<bool>,
    field_id_wrapped: Option<bool>,
    headers_in_last_refresh: Option<bool>,
    minimum_version: Option<u8>,
    next_id: Option<u32>,
    unbound_columns_left: Option<u32>,
    unbound_columns_right: Option<u32>,
    declared_field_count: Option<u32>,
    fields: Vec<QueryTableField>,
    declared_deleted_field_count: Option<u32>,
    deleted_fields: Option<Vec<String>>,
    sort_state: Option<QueryTableSortState>,
    extension_list: Option<QueryTableExtensionList>,
    extension_attributes: Vec<QueryTableExtensionAttribute>,
}

impl QueryTableRefresh {
    pub fn new() -> Self {
        Self {
            preserve_sort_filter_layout: None,
            field_id_wrapped: None,
            headers_in_last_refresh: None,
            minimum_version: None,
            next_id: None,
            unbound_columns_left: None,
            unbound_columns_right: None,
            declared_field_count: None,
            fields: Vec::new(),
            declared_deleted_field_count: None,
            deleted_fields: None,
            sort_state: None,
            extension_list: None,
            extension_attributes: Vec::new(),
        }
    }
    pub fn fields_mut(&mut self) -> &mut Vec<QueryTableField> { &mut self.fields }
    pub fn add_field(&mut self, field: QueryTableField) { self.fields.push(field); }
    pub fn deleted_fields_mut(&mut self) -> &mut Option<Vec<String>> { &mut self.deleted_fields }
    pub fn add_deleted_field(&mut self, name: String) {
        self.deleted_fields.get_or_insert_with(Vec::new).push(name);
    }
    pub fn set_preserve_sort_filter_layout(&mut self, value: Option<bool>) { self.preserve_sort_filter_layout = value; }
    pub fn set_field_id_wrapped(&mut self, value: Option<bool>) { self.field_id_wrapped = value; }
    pub fn set_headers_in_last_refresh(&mut self, value: Option<bool>) { self.headers_in_last_refresh = value; }
    pub fn set_minimum_version(&mut self, value: Option<u8>) { self.minimum_version = value; }
    pub fn set_next_id(&mut self, value: Option<u32>) { self.next_id = value; }
    pub fn set_unbound_columns_left(&mut self, value: Option<u32>) { self.unbound_columns_left = value; }
    pub fn set_unbound_columns_right(&mut self, value: Option<u32>) { self.unbound_columns_right = value; }
    pub fn preserve_sort_filter_layout(&self) -> Option<bool> {
        self.preserve_sort_filter_layout
    }
    pub fn field_id_wrapped(&self) -> Option<bool> {
        self.field_id_wrapped
    }
    pub fn headers_in_last_refresh(&self) -> Option<bool> {
        self.headers_in_last_refresh
    }
    pub fn minimum_version(&self) -> Option<u8> {
        self.minimum_version
    }
    pub fn next_id(&self) -> Option<u32> {
        self.next_id
    }
    pub fn unbound_columns_left(&self) -> Option<u32> {
        self.unbound_columns_left
    }
    pub fn unbound_columns_right(&self) -> Option<u32> {
        self.unbound_columns_right
    }
    pub fn declared_field_count(&self) -> Option<u32> {
        self.declared_field_count
    }
    pub fn fields(&self) -> &[QueryTableField] {
        &self.fields
    }
    pub fn declared_deleted_field_count(&self) -> Option<u32> {
        self.declared_deleted_field_count
    }
    pub fn deleted_fields(&self) -> Option<&[String]> {
        self.deleted_fields.as_deref()
    }
    pub fn sort_state(&self) -> Option<&QueryTableSortState> {
        self.sort_state.as_ref()
    }
    pub fn extension_list(&self) -> Option<&QueryTableExtensionList> {
        self.extension_list.as_ref()
    }
    pub fn extension_attributes(&self) -> &[QueryTableExtensionAttribute] {
        &self.extension_attributes
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct QueryTable {
    name: String,
    headers: Option<bool>,
    row_numbers: Option<bool>,
    disable_refresh: Option<bool>,
    background_refresh: Option<bool>,
    first_background_refresh: Option<bool>,
    refresh_on_load: Option<bool>,
    grow_shrink_type: Option<QueryTableGrowShrinkType>,
    fill_formulas: Option<bool>,
    remove_data_on_save: Option<bool>,
    disable_edit: Option<bool>,
    preserve_formatting: Option<bool>,
    adjust_column_width: Option<bool>,
    intermediate: Option<bool>,
    connection_id: u32,
    auto_format_id: Option<u32>,
    apply_number_formats: Option<bool>,
    apply_border_formats: Option<bool>,
    apply_font_formats: Option<bool>,
    apply_pattern_formats: Option<bool>,
    apply_alignment_formats: Option<bool>,
    apply_width_height_formats: Option<bool>,
    refresh: Option<QueryTableRefresh>,
    extension_list: Option<QueryTableExtensionList>,
    namespaces: Vec<QueryTableExtensionAttribute>,
    extension_attributes: Vec<QueryTableExtensionAttribute>,
}

impl QueryTable {
    pub fn new(name: impl Into<String>, connection_id: u32) -> Self {
        Self {
            name: name.into(),
            headers: None,
            row_numbers: None,
            disable_refresh: None,
            background_refresh: None,
            first_background_refresh: None,
            refresh_on_load: None,
            grow_shrink_type: None,
            fill_formulas: None,
            remove_data_on_save: None,
            disable_edit: None,
            preserve_formatting: None,
            adjust_column_width: None,
            intermediate: None,
            connection_id,
            auto_format_id: None,
            apply_number_formats: None,
            apply_border_formats: None,
            apply_font_formats: None,
            apply_pattern_formats: None,
            apply_alignment_formats: None,
            apply_width_height_formats: None,
            refresh: None,
            extension_list: None,
            namespaces: Vec::new(),
            extension_attributes: Vec::new(),
        }
    }
    pub fn set_name(&mut self, value: String) { self.name = value; }
    pub fn set_connection_id(&mut self, value: u32) { self.connection_id = value; }
    pub fn set_headers(&mut self, value: Option<bool>) { self.headers = value; }
    pub fn set_row_numbers(&mut self, value: Option<bool>) { self.row_numbers = value; }
    pub fn set_disable_refresh(&mut self, value: Option<bool>) { self.disable_refresh = value; }
    pub fn set_background_refresh(&mut self, value: Option<bool>) { self.background_refresh = value; }
    pub fn set_first_background_refresh(&mut self, value: Option<bool>) { self.first_background_refresh = value; }
    pub fn set_refresh_on_load(&mut self, value: Option<bool>) { self.refresh_on_load = value; }
    pub fn set_grow_shrink_type(&mut self, value: Option<QueryTableGrowShrinkType>) { self.grow_shrink_type = value; }
    pub fn set_fill_formulas(&mut self, value: Option<bool>) { self.fill_formulas = value; }
    pub fn set_remove_data_on_save(&mut self, value: Option<bool>) { self.remove_data_on_save = value; }
    pub fn set_disable_edit(&mut self, value: Option<bool>) { self.disable_edit = value; }
    pub fn set_preserve_formatting(&mut self, value: Option<bool>) { self.preserve_formatting = value; }
    pub fn set_adjust_column_width(&mut self, value: Option<bool>) { self.adjust_column_width = value; }
    pub fn set_intermediate(&mut self, value: Option<bool>) { self.intermediate = value; }
    pub fn set_refresh(&mut self, value: Option<QueryTableRefresh>) { self.refresh = value; }
    pub fn refresh_mut(&mut self) -> Option<&mut QueryTableRefresh> { self.refresh.as_mut() }
    pub fn name(&self) -> &str {
        &self.name
    }
    pub fn headers(&self) -> Option<bool> {
        self.headers
    }
    pub fn row_numbers(&self) -> Option<bool> {
        self.row_numbers
    }
    pub fn disable_refresh(&self) -> Option<bool> {
        self.disable_refresh
    }
    pub fn background_refresh(&self) -> Option<bool> {
        self.background_refresh
    }
    pub fn first_background_refresh(&self) -> Option<bool> {
        self.first_background_refresh
    }
    pub fn refresh_on_load(&self) -> Option<bool> {
        self.refresh_on_load
    }
    pub fn grow_shrink_type(&self) -> Option<QueryTableGrowShrinkType> {
        self.grow_shrink_type
    }
    pub fn fill_formulas(&self) -> Option<bool> {
        self.fill_formulas
    }
    pub fn remove_data_on_save(&self) -> Option<bool> {
        self.remove_data_on_save
    }
    pub fn disable_edit(&self) -> Option<bool> {
        self.disable_edit
    }
    pub fn preserve_formatting(&self) -> Option<bool> {
        self.preserve_formatting
    }
    pub fn adjust_column_width(&self) -> Option<bool> {
        self.adjust_column_width
    }
    pub fn intermediate(&self) -> Option<bool> {
        self.intermediate
    }
    pub fn connection_id(&self) -> u32 {
        self.connection_id
    }
    pub fn auto_format_id(&self) -> Option<u32> {
        self.auto_format_id
    }
    pub fn apply_number_formats(&self) -> Option<bool> {
        self.apply_number_formats
    }
    pub fn apply_border_formats(&self) -> Option<bool> {
        self.apply_border_formats
    }
    pub fn apply_font_formats(&self) -> Option<bool> {
        self.apply_font_formats
    }
    pub fn apply_pattern_formats(&self) -> Option<bool> {
        self.apply_pattern_formats
    }
    pub fn apply_alignment_formats(&self) -> Option<bool> {
        self.apply_alignment_formats
    }
    pub fn apply_width_height_formats(&self) -> Option<bool> {
        self.apply_width_height_formats
    }
    pub fn refresh(&self) -> Option<&QueryTableRefresh> {
        self.refresh.as_ref()
    }
    pub fn extension_list(&self) -> Option<&QueryTableExtensionList> {
        self.extension_list.as_ref()
    }
    pub fn extension_attributes(&self) -> &[QueryTableExtensionAttribute] {
        &self.extension_attributes
    }

    pub fn to_xml(&self, conformance: QueryTableConformance) -> Result<Vec<u8>> {
        write_query_table(self, conformance)
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct WorksheetQueryTable {
    relationship_id: String,
    part_name: String,
    query_table: QueryTable,
}

impl WorksheetQueryTable {
    pub(crate) fn new(relationship_id: String, part_name: String, query_table: QueryTable) -> Self {
        Self {
            relationship_id,
            part_name,
            query_table,
        }
    }

    pub fn relationship_id(&self) -> &str {
        &self.relationship_id
    }
    pub fn part_name(&self) -> &str {
        &self.part_name
    }
    pub fn query_table(&self) -> &QueryTable {
        &self.query_table
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct XmlAttribute {
    qualified_name: String,
    value: String,
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct XmlNode {
    qualified_name: String,
    local_name: String,
    namespace: String,
    attributes: Vec<XmlAttribute>,
    children: Vec<XmlNode>,
    text: String,
}

pub fn is_query_table_relationship_type(value: &str) -> bool {
    matches!(
        value,
        QUERY_TABLE_RELATIONSHIP_TYPE | STRICT_QUERY_TABLE_RELATIONSHIP_TYPE
    )
}

pub fn parse_query_table(xml: &[u8]) -> Result<QueryTable> {
    if xml.len() > MAX_PART_BYTES {
        return Err(invalid("query-table part is too large"));
    }
    let processed = process_ooxml(xml)?;
    if processed.len() > MAX_PART_BYTES {
        return Err(invalid("MCE-expanded query-table part is too large"));
    }
    let root = parse_xml_tree(processed.as_ref())?;
    parse_query_table_node(&root)
}

pub fn write_query_table(
    value: &QueryTable,
    conformance: QueryTableConformance,
) -> Result<Vec<u8>> {
    let mut output = Vec::new();
    output.extend_from_slice(br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
    output.extend_from_slice(b"<queryTable xmlns=\"");
    escape_attribute(&mut output, conformance.namespace());
    output.push(b'\"');
    write_namespaces(&mut output, &value.namespaces, true)?;
    write_extra_attributes(&mut output, &value.extension_attributes)?;
    write_string_attribute(&mut output, "name", &value.name);
    write_bool_attribute(&mut output, "headers", value.headers);
    write_bool_attribute(&mut output, "rowNumbers", value.row_numbers);
    write_bool_attribute(&mut output, "disableRefresh", value.disable_refresh);
    write_bool_attribute(&mut output, "backgroundRefresh", value.background_refresh);
    write_bool_attribute(
        &mut output,
        "firstBackgroundRefresh",
        value.first_background_refresh,
    );
    write_bool_attribute(&mut output, "refreshOnLoad", value.refresh_on_load);
    if let Some(kind) = value.grow_shrink_type {
        write_string_attribute(&mut output, "growShrinkType", kind.as_str());
    }
    write_bool_attribute(&mut output, "fillFormulas", value.fill_formulas);
    write_bool_attribute(&mut output, "removeDataOnSave", value.remove_data_on_save);
    write_bool_attribute(&mut output, "disableEdit", value.disable_edit);
    write_bool_attribute(&mut output, "preserveFormatting", value.preserve_formatting);
    write_bool_attribute(&mut output, "adjustColumnWidth", value.adjust_column_width);
    write_bool_attribute(&mut output, "intermediate", value.intermediate);
    write_u32_attribute(&mut output, "connectionId", Some(value.connection_id));
    write_u32_attribute(&mut output, "autoFormatId", value.auto_format_id);
    write_bool_attribute(
        &mut output,
        "applyNumberFormats",
        value.apply_number_formats,
    );
    write_bool_attribute(
        &mut output,
        "applyBorderFormats",
        value.apply_border_formats,
    );
    write_bool_attribute(&mut output, "applyFontFormats", value.apply_font_formats);
    write_bool_attribute(
        &mut output,
        "applyPatternFormats",
        value.apply_pattern_formats,
    );
    write_bool_attribute(
        &mut output,
        "applyAlignmentFormats",
        value.apply_alignment_formats,
    );
    write_bool_attribute(
        &mut output,
        "applyWidthHeightFormats",
        value.apply_width_height_formats,
    );
    if value.refresh.is_none() && value.extension_list.is_none() {
        output.extend_from_slice(b"/>");
        return Ok(output);
    }
    output.push(b'>');
    if let Some(refresh) = &value.refresh {
        write_refresh(&mut output, refresh)?;
    }
    if let Some(extension_list) = &value.extension_list {
        write_extension_list(&mut output, extension_list)?;
    }
    output.extend_from_slice(b"</queryTable>");
    Ok(output)
}

fn parse_xml_tree(xml: &[u8]) -> Result<XmlNode> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut stack: Vec<XmlNode> = Vec::new();
    let mut root = None;
    let mut node_count = 0usize;
    loop {
        let decoder = reader.decoder();
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(xml_error)?;
        match event {
            Event::Start(start) => {
                node_count = node_count
                    .checked_add(1)
                    .ok_or_else(|| invalid("too many XML nodes"))?;
                if node_count > MAX_NODES || stack.len() >= MAX_DEPTH {
                    return Err(invalid("query-table XML resource limit exceeded"));
                }
                stack.push(make_node(&namespace, &start, decoder)?);
            },
            Event::Empty(start) => {
                node_count = node_count
                    .checked_add(1)
                    .ok_or_else(|| invalid("too many XML nodes"))?;
                if node_count > MAX_NODES || stack.len() >= MAX_DEPTH {
                    return Err(invalid("query-table XML resource limit exceeded"));
                }
                attach_node(
                    make_node(&namespace, &start, decoder)?,
                    &mut stack,
                    &mut root,
                )?;
            },
            Event::End(_) => {
                let node = stack
                    .pop()
                    .ok_or_else(|| invalid("unexpected XML end tag"))?;
                attach_node(node, &mut stack, &mut root)?;
            },
            Event::Text(text) => {
                let decoded = text.decode().map_err(xml_error)?;
                let decoded = quick_xml::escape::unescape(&decoded).map_err(xml_error)?;
                append_text(&mut stack, &decoded)?;
            },
            Event::CData(text) => {
                let decoded = text.decode().map_err(xml_error)?;
                append_text(&mut stack, &decoded)?;
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("DTD and processing instructions are rejected"));
            },
            Event::Eof => break,
            Event::Decl(_) | Event::Comment(_) | Event::GeneralRef(_) => {},
        }
        buffer.clear();
    }
    if !stack.is_empty() {
        return Err(invalid("unterminated query-table XML"));
    }
    root.ok_or_else(|| invalid("query-table part has no root element"))
}

fn make_node(
    namespace: &ResolveResult<'_>,
    start: &quick_xml::events::BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
) -> Result<XmlNode> {
    let qualified_name = std::str::from_utf8(start.name().as_ref())
        .map_err(xml_error)?
        .to_string();
    let local_name = std::str::from_utf8(start.local_name().as_ref())
        .map_err(xml_error)?
        .to_string();
    let namespace = match namespace {
        ResolveResult::Bound(value) => std::str::from_utf8(value.as_ref())
            .map_err(xml_error)?
            .to_string(),
        ResolveResult::Unbound => String::new(),
        ResolveResult::Unknown(prefix) => {
            return Err(invalid(format!(
                "unbound XML namespace prefix '{}'",
                String::from_utf8_lossy(prefix)
            )));
        },
    };
    let mut attributes = Vec::new();
    let mut names = HashSet::new();
    for attribute in start.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        let name = std::str::from_utf8(attribute.key.as_ref())
            .map_err(xml_error)?
            .to_string();
        if !names.insert(name.clone()) {
            return Err(invalid("duplicate XML attribute"));
        }
        let value = attribute
            .decoded_and_normalized_value(quick_xml::XmlVersion::Implicit1_0, decoder)
            .map_err(xml_error)?
            .into_owned();
        bounded_text(&value)?;
        attributes.push(XmlAttribute {
            qualified_name: name,
            value,
        });
    }
    Ok(XmlNode {
        qualified_name,
        local_name,
        namespace,
        attributes,
        children: Vec::new(),
        text: String::new(),
    })
}

fn attach_node(node: XmlNode, stack: &mut [XmlNode], root: &mut Option<XmlNode>) -> Result<()> {
    if let Some(parent) = stack.last_mut() {
        parent.children.push(node);
    } else if root.replace(node).is_some() {
        return Err(invalid("query-table part has multiple root elements"));
    }
    Ok(())
}

fn append_text(stack: &mut [XmlNode], value: &str) -> Result<()> {
    if value.len() > MAX_TEXT_BYTES {
        return Err(invalid("query-table text is too large"));
    }
    if let Some(node) = stack.last_mut() {
        if node.text.len().saturating_add(value.len()) > MAX_TEXT_BYTES {
            return Err(invalid("query-table element text is too large"));
        }
        node.text.push_str(value);
    } else if !value.trim().is_empty() {
        return Err(invalid("text outside query-table root"));
    }
    Ok(())
}

fn parse_query_table_node(node: &XmlNode) -> Result<QueryTable> {
    require_element(node, "queryTable")?;
    require_whitespace(node)?;
    let known = [
        "name",
        "headers",
        "rowNumbers",
        "disableRefresh",
        "backgroundRefresh",
        "firstBackgroundRefresh",
        "refreshOnLoad",
        "growShrinkType",
        "fillFormulas",
        "removeDataOnSave",
        "disableEdit",
        "preserveFormatting",
        "adjustColumnWidth",
        "intermediate",
        "connectionId",
        "autoFormatId",
        "applyNumberFormats",
        "applyBorderFormats",
        "applyFontFormats",
        "applyPatternFormats",
        "applyAlignmentFormats",
        "applyWidthHeightFormats",
    ];
    let parsed = ParsedAttributes::new(node, &known)?;
    let name = parsed.required_string("name")?;
    let connection_id = parsed.required_u32("connectionId")?;
    let grow_shrink_type = parsed
        .optional_string("growShrinkType")?
        .map(|value| QueryTableGrowShrinkType::parse(&value))
        .transpose()?;
    let mut refresh = None;
    let mut extension_list = None;
    let mut phase = 0u8;
    for child in &node.children {
        match child.local_name.as_str() {
            "queryTableRefresh" if is_core(child) => {
                if phase != 0 || refresh.is_some() {
                    return Err(invalid("duplicate or out-of-order queryTableRefresh"));
                }
                refresh = Some(parse_refresh(child)?);
                phase = 1;
            },
            "extLst" if is_core(child) => {
                if phase > 1 || extension_list.is_some() {
                    return Err(invalid("duplicate or out-of-order queryTable extLst"));
                }
                extension_list = Some(parse_extension_list(child)?);
                phase = 2;
            },
            _ => {
                return Err(invalid(format!(
                    "unexpected queryTable child '{}'",
                    child.qualified_name
                )));
            },
        }
    }
    Ok(QueryTable {
        name,
        headers: parsed.optional_bool("headers")?,
        row_numbers: parsed.optional_bool("rowNumbers")?,
        disable_refresh: parsed.optional_bool("disableRefresh")?,
        background_refresh: parsed.optional_bool("backgroundRefresh")?,
        first_background_refresh: parsed.optional_bool("firstBackgroundRefresh")?,
        refresh_on_load: parsed.optional_bool("refreshOnLoad")?,
        grow_shrink_type,
        fill_formulas: parsed.optional_bool("fillFormulas")?,
        remove_data_on_save: parsed.optional_bool("removeDataOnSave")?,
        disable_edit: parsed.optional_bool("disableEdit")?,
        preserve_formatting: parsed.optional_bool("preserveFormatting")?,
        adjust_column_width: parsed.optional_bool("adjustColumnWidth")?,
        intermediate: parsed.optional_bool("intermediate")?,
        connection_id,
        auto_format_id: parsed.optional_u32("autoFormatId")?,
        apply_number_formats: parsed.optional_bool("applyNumberFormats")?,
        apply_border_formats: parsed.optional_bool("applyBorderFormats")?,
        apply_font_formats: parsed.optional_bool("applyFontFormats")?,
        apply_pattern_formats: parsed.optional_bool("applyPatternFormats")?,
        apply_alignment_formats: parsed.optional_bool("applyAlignmentFormats")?,
        apply_width_height_formats: parsed.optional_bool("applyWidthHeightFormats")?,
        refresh,
        extension_list,
        namespaces: namespace_attributes(node),
        extension_attributes: parsed.extension_attributes,
    })
}

fn parse_refresh(node: &XmlNode) -> Result<QueryTableRefresh> {
    require_whitespace(node)?;
    let known = [
        "preserveSortFilterLayout",
        "fieldIdWrapped",
        "headersInLastRefresh",
        "minimumVersion",
        "nextId",
        "unboundColumnsLeft",
        "unboundColumnsRight",
    ];
    let parsed = ParsedAttributes::new(node, &known)?;
    let mut fields = None;
    let mut deleted_fields = None;
    let mut sort_state = None;
    let mut extension_list = None;
    let mut phase = 0u8;
    for child in &node.children {
        match child.local_name.as_str() {
            "queryTableFields" if is_core(child) => {
                if phase != 0 || fields.is_some() {
                    return Err(invalid("duplicate or out-of-order queryTableFields"));
                }
                fields = Some(parse_fields(child)?);
                phase = 1;
            },
            "queryTableDeletedFields" if is_core(child) => {
                if phase > 1 || deleted_fields.is_some() {
                    return Err(invalid("duplicate or out-of-order queryTableDeletedFields"));
                }
                deleted_fields = Some(parse_deleted_fields(child)?);
                phase = 2;
            },
            "sortState" if is_core(child) => {
                if phase > 2 || sort_state.is_some() {
                    return Err(invalid("duplicate or out-of-order query-table sortState"));
                }
                sort_state = Some(parse_sort_state(child)?);
                phase = 3;
            },
            "extLst" if is_core(child) => {
                if phase > 3 || extension_list.is_some() {
                    return Err(invalid(
                        "duplicate or out-of-order queryTableRefresh extLst",
                    ));
                }
                extension_list = Some(parse_extension_list(child)?);
                phase = 4;
            },
            _ => {
                return Err(invalid(format!(
                    "unexpected queryTableRefresh child '{}'",
                    child.qualified_name
                )));
            },
        }
    }
    let (declared_field_count, fields) =
        fields.ok_or_else(|| invalid("queryTableRefresh requires queryTableFields"))?;
    let (declared_deleted_field_count, deleted_fields) = match deleted_fields {
        Some((count, values)) => (count, Some(values)),
        None => (None, None),
    };
    Ok(QueryTableRefresh {
        preserve_sort_filter_layout: parsed.optional_bool("preserveSortFilterLayout")?,
        field_id_wrapped: parsed.optional_bool("fieldIdWrapped")?,
        headers_in_last_refresh: parsed.optional_bool("headersInLastRefresh")?,
        minimum_version: parsed.optional_u8("minimumVersion")?,
        next_id: parsed.optional_u32("nextId")?,
        unbound_columns_left: parsed.optional_u32("unboundColumnsLeft")?,
        unbound_columns_right: parsed.optional_u32("unboundColumnsRight")?,
        declared_field_count,
        fields,
        declared_deleted_field_count,
        deleted_fields,
        sort_state,
        extension_list,
        extension_attributes: parsed.extension_attributes,
    })
}

fn parse_fields(node: &XmlNode) -> Result<(Option<u32>, Vec<QueryTableField>)> {
    require_whitespace(node)?;
    let parsed = ParsedAttributes::new(node, &["count"])?;
    if node.children.len() > MAX_FIELDS {
        return Err(invalid("too many query-table fields"));
    }
    let mut fields = Vec::with_capacity(node.children.len());
    let mut ids = HashSet::with_capacity(node.children.len());
    for child in &node.children {
        require_element(child, "queryTableField")?;
        let field = parse_field(child)?;
        if !ids.insert(field.id) {
            return Err(invalid(format!(
                "duplicate query-table field ID {}",
                field.id
            )));
        }
        fields.push(field);
    }
    let count = parsed.optional_u32("count")?;
    if count.is_some_and(|value| value as usize != fields.len()) {
        return Err(invalid("queryTableFields count does not match children"));
    }
    Ok((count, fields))
}

fn parse_field(node: &XmlNode) -> Result<QueryTableField> {
    require_whitespace(node)?;
    let parsed = ParsedAttributes::new(
        node,
        &[
            "id",
            "name",
            "dataBound",
            "rowNumbers",
            "fillFormulas",
            "clipped",
            "tableColumnId",
        ],
    )?;
    let mut extension_list = None;
    for child in &node.children {
        require_element(child, "extLst")?;
        if extension_list.is_some() {
            return Err(invalid("queryTableField has multiple extLst children"));
        }
        extension_list = Some(parse_extension_list(child)?);
    }
    Ok(QueryTableField {
        id: parsed.required_u32("id")?,
        name: parsed.optional_string("name")?,
        data_bound: parsed.optional_bool("dataBound")?,
        row_numbers: parsed.optional_bool("rowNumbers")?,
        fill_formulas: parsed.optional_bool("fillFormulas")?,
        clipped: parsed.optional_bool("clipped")?,
        table_column_id: parsed.optional_u32("tableColumnId")?,
        extension_list,
        extension_attributes: parsed.extension_attributes,
    })
}

fn parse_deleted_fields(node: &XmlNode) -> Result<(Option<u32>, Vec<String>)> {
    require_whitespace(node)?;
    let parsed = ParsedAttributes::new(node, &["count"])?;
    if node.children.is_empty() {
        return Err(invalid(
            "queryTableDeletedFields requires at least one deletedField",
        ));
    }
    if node.children.len() > MAX_DELETED_FIELDS {
        return Err(invalid("too many deleted query-table fields"));
    }
    let mut values = Vec::with_capacity(node.children.len());
    for child in &node.children {
        require_element(child, "deletedField")?;
        require_leaf(child)?;
        let attrs = ParsedAttributes::new(child, &["name"])?;
        if !attrs.extension_attributes.is_empty() {
            return Err(invalid("deletedField does not allow extension attributes"));
        }
        values.push(attrs.required_string("name")?);
    }
    let count = parsed.optional_u32("count")?;
    if count.is_some_and(|value| value as usize != values.len()) {
        return Err(invalid(
            "queryTableDeletedFields count does not match children",
        ));
    }
    Ok((count, values))
}

fn parse_sort_state(node: &XmlNode) -> Result<QueryTableSortState> {
    require_whitespace(node)?;
    let parsed =
        ParsedAttributes::new(node, &["ref", "columnSort", "caseSensitive", "sortMethod"])?;
    let reference = parsed.required_string("ref")?;
    parse_range(&reference)?;
    if node.children.len() > MAX_SORT_CONDITIONS {
        return Err(invalid("too many query-table sort conditions"));
    }
    let mut conditions = Vec::with_capacity(node.children.len());
    for child in &node.children {
        require_element(child, "sortCondition")?;
        require_leaf(child)?;
        let attrs = ParsedAttributes::new(
            child,
            &[
                "ref",
                "descending",
                "sortBy",
                "customList",
                "dxfId",
                "iconSet",
                "iconId",
            ],
        )?;
        let reference = attrs.required_string("ref")?;
        parse_range(&reference)?;
        let sort_by = attrs
            .optional_string("sortBy")?
            .map(|value| QueryTableSortBy::parse(&value))
            .transpose()?;
        let icon_set = attrs
            .optional_string("iconSet")?
            .map(|value| QueryTableIconSet::parse(&value))
            .transpose()?;
        let differential_format_id = attrs.optional_u32("dxfId")?;
        let icon_id = attrs.optional_u32("iconId")?;
        match sort_by.unwrap_or(QueryTableSortBy::Value) {
            QueryTableSortBy::CellColor | QueryTableSortBy::FontColor
                if differential_format_id.is_none() =>
            {
                return Err(invalid("color sort requires dxfId"));
            },
            QueryTableSortBy::Icon if icon_set.is_none() => {
                return Err(invalid("icon sort requires iconSet"));
            },
            QueryTableSortBy::Icon
                if icon_id.is_some_and(|value| value >= icon_set.unwrap().cardinality()) =>
            {
                return Err(invalid("sort iconId exceeds icon-set cardinality"));
            },
            _ => {},
        }
        conditions.push(QueryTableSortCondition {
            reference,
            descending: attrs.optional_bool("descending")?,
            sort_by,
            custom_list: attrs.optional_string("customList")?,
            differential_format_id,
            icon_set,
            icon_id,
            extension_attributes: attrs.extension_attributes,
        });
    }
    Ok(QueryTableSortState {
        reference,
        column_sort: parsed.optional_bool("columnSort")?,
        case_sensitive: parsed.optional_bool("caseSensitive")?,
        sort_method: parsed
            .optional_string("sortMethod")?
            .map(|value| QueryTableSortMethod::parse(&value))
            .transpose()?,
        conditions,
        extension_attributes: parsed.extension_attributes,
    })
}

fn parse_extension_list(node: &XmlNode) -> Result<QueryTableExtensionList> {
    require_whitespace(node)?;
    let parsed = ParsedAttributes::new(node, &[])?;
    if node.children.len() > MAX_EXTENSIONS {
        return Err(invalid("too many query-table extensions"));
    }
    let mut extensions = Vec::with_capacity(node.children.len());
    for child in &node.children {
        require_element(child, "ext")?;
        require_whitespace(child)?;
        let attrs = ParsedAttributes::new(child, &["uri"])?;
        if child.children.len() > 1 {
            return Err(invalid(
                "query-table extension has more than one payload element",
            ));
        }
        extensions.push(QueryTableExtension {
            uri: attrs.required_string("uri")?,
            namespaces: namespace_attributes(child),
            attributes: attrs.extension_attributes,
            content: child.children.clone(),
        });
    }
    Ok(QueryTableExtensionList {
        namespaces: namespace_attributes(node),
        attributes: parsed.extension_attributes,
        extensions,
    })
}

struct ParsedAttributes {
    values: Vec<(String, String)>,
    extension_attributes: Vec<QueryTableExtensionAttribute>,
}

impl ParsedAttributes {
    fn new(node: &XmlNode, known: &[&str]) -> Result<Self> {
        let mut values = Vec::new();
        let mut extension_attributes = Vec::new();
        for attribute in &node.attributes {
            if is_namespace_attribute(&attribute.qualified_name) {
                continue;
            }
            if attribute.qualified_name.contains(':') {
                extension_attributes.push(QueryTableExtensionAttribute {
                    qualified_name: attribute.qualified_name.clone(),
                    value: attribute.value.clone(),
                });
            } else if known.contains(&attribute.qualified_name.as_str()) {
                values.push((attribute.qualified_name.clone(), attribute.value.clone()));
            } else {
                return Err(invalid(format!(
                    "unexpected '{}' attribute on {}",
                    attribute.qualified_name, node.qualified_name
                )));
            }
        }
        Ok(Self {
            values,
            extension_attributes,
        })
    }

    fn optional_string(&self, name: &str) -> Result<Option<String>> {
        let value = self
            .values
            .iter()
            .find(|(key, _)| key == name)
            .map(|(_, value)| value.clone());
        if let Some(value) = &value {
            bounded_text(value)?;
        }
        Ok(value)
    }

    fn required_string(&self, name: &str) -> Result<String> {
        self.optional_string(name)?
            .ok_or_else(|| invalid(format!("missing '{name}' attribute")))
    }

    fn optional_u32(&self, name: &str) -> Result<Option<u32>> {
        self.optional_string(name)?
            .map(|value| {
                value
                    .parse::<u32>()
                    .map_err(|_| invalid(format!("invalid unsigned integer '{value}'")))
            })
            .transpose()
    }

    fn required_u32(&self, name: &str) -> Result<u32> {
        self.optional_u32(name)?
            .ok_or_else(|| invalid(format!("missing '{name}' attribute")))
    }

    fn optional_u8(&self, name: &str) -> Result<Option<u8>> {
        self.optional_string(name)?
            .map(|value| {
                value
                    .parse::<u8>()
                    .map_err(|_| invalid(format!("invalid byte '{value}'")))
            })
            .transpose()
    }

    fn optional_bool(&self, name: &str) -> Result<Option<bool>> {
        self.optional_string(name)?
            .map(|value| match value.as_str() {
                "1" | "true" => Ok(true),
                "0" | "false" => Ok(false),
                _ => Err(invalid(format!("invalid boolean '{value}'"))),
            })
            .transpose()
    }
}

fn write_refresh(output: &mut Vec<u8>, value: &QueryTableRefresh) -> Result<()> {
    output.extend_from_slice(b"<queryTableRefresh");
    write_extra_attributes(output, &value.extension_attributes)?;
    write_bool_attribute(
        output,
        "preserveSortFilterLayout",
        value.preserve_sort_filter_layout,
    );
    write_bool_attribute(output, "fieldIdWrapped", value.field_id_wrapped);
    write_bool_attribute(
        output,
        "headersInLastRefresh",
        value.headers_in_last_refresh,
    );
    if let Some(minimum_version) = value.minimum_version {
        write_string_attribute(output, "minimumVersion", &minimum_version.to_string());
    }
    write_u32_attribute(output, "nextId", value.next_id);
    write_u32_attribute(output, "unboundColumnsLeft", value.unbound_columns_left);
    write_u32_attribute(output, "unboundColumnsRight", value.unbound_columns_right);
    output.push(b'>');
    output.extend_from_slice(b"<queryTableFields");
    write_u32_attribute(output, "count", value.declared_field_count);
    if value.fields.is_empty() {
        output.extend_from_slice(b"/>");
    } else {
        output.push(b'>');
        for field in &value.fields {
            write_field(output, field)?;
        }
        output.extend_from_slice(b"</queryTableFields>");
    }
    if let Some(deleted_fields) = &value.deleted_fields {
        output.extend_from_slice(b"<queryTableDeletedFields");
        write_u32_attribute(output, "count", value.declared_deleted_field_count);
        output.push(b'>');
        for name in deleted_fields {
            output.extend_from_slice(b"<deletedField");
            write_string_attribute(output, "name", name);
            output.extend_from_slice(b"/>");
        }
        output.extend_from_slice(b"</queryTableDeletedFields>");
    }
    if let Some(sort_state) = &value.sort_state {
        write_sort_state(output, sort_state)?;
    }
    if let Some(extension_list) = &value.extension_list {
        write_extension_list(output, extension_list)?;
    }
    output.extend_from_slice(b"</queryTableRefresh>");
    Ok(())
}

fn write_field(output: &mut Vec<u8>, value: &QueryTableField) -> Result<()> {
    output.extend_from_slice(b"<queryTableField");
    write_extra_attributes(output, &value.extension_attributes)?;
    write_u32_attribute(output, "id", Some(value.id));
    if let Some(name) = &value.name {
        write_string_attribute(output, "name", name);
    }
    write_bool_attribute(output, "dataBound", value.data_bound);
    write_bool_attribute(output, "rowNumbers", value.row_numbers);
    write_bool_attribute(output, "fillFormulas", value.fill_formulas);
    write_bool_attribute(output, "clipped", value.clipped);
    write_u32_attribute(output, "tableColumnId", value.table_column_id);
    if let Some(extension_list) = &value.extension_list {
        output.push(b'>');
        write_extension_list(output, extension_list)?;
        output.extend_from_slice(b"</queryTableField>");
    } else {
        output.extend_from_slice(b"/>");
    }
    Ok(())
}

fn write_sort_state(output: &mut Vec<u8>, value: &QueryTableSortState) -> Result<()> {
    output.extend_from_slice(b"<sortState");
    write_extra_attributes(output, &value.extension_attributes)?;
    write_string_attribute(output, "ref", &value.reference);
    write_bool_attribute(output, "columnSort", value.column_sort);
    write_bool_attribute(output, "caseSensitive", value.case_sensitive);
    if let Some(method) = value.sort_method {
        write_string_attribute(output, "sortMethod", method.as_str());
    }
    if value.conditions.is_empty() {
        output.extend_from_slice(b"/>");
        return Ok(());
    }
    output.push(b'>');
    for condition in &value.conditions {
        output.extend_from_slice(b"<sortCondition");
        write_extra_attributes(output, &condition.extension_attributes)?;
        write_string_attribute(output, "ref", &condition.reference);
        write_bool_attribute(output, "descending", condition.descending);
        if let Some(sort_by) = condition.sort_by {
            write_string_attribute(output, "sortBy", sort_by.as_str());
        }
        if let Some(custom_list) = &condition.custom_list {
            write_string_attribute(output, "customList", custom_list);
        }
        write_u32_attribute(output, "dxfId", condition.differential_format_id);
        if let Some(icon_set) = condition.icon_set {
            write_string_attribute(output, "iconSet", icon_set.as_str());
        }
        write_u32_attribute(output, "iconId", condition.icon_id);
        output.extend_from_slice(b"/>");
    }
    output.extend_from_slice(b"</sortState>");
    Ok(())
}

fn write_extension_list(output: &mut Vec<u8>, value: &QueryTableExtensionList) -> Result<()> {
    output.extend_from_slice(b"<extLst");
    write_namespaces(output, &value.namespaces, false)?;
    write_extra_attributes(output, &value.attributes)?;
    if value.extensions.is_empty() {
        output.extend_from_slice(b"/>");
        return Ok(());
    }
    output.push(b'>');
    for extension in &value.extensions {
        output.extend_from_slice(b"<ext");
        write_namespaces(output, &extension.namespaces, false)?;
        write_extra_attributes(output, &extension.attributes)?;
        write_string_attribute(output, "uri", &extension.uri);
        if extension.content.is_empty() {
            output.extend_from_slice(b"/>");
        } else {
            output.push(b'>');
            for child in &extension.content {
                write_xml_node(output, child)?;
            }
            output.extend_from_slice(b"</ext>");
        }
    }
    output.extend_from_slice(b"</extLst>");
    Ok(())
}

fn write_xml_node(output: &mut Vec<u8>, node: &XmlNode) -> Result<()> {
    output.push(b'<');
    output.extend_from_slice(node.qualified_name.as_bytes());
    for attribute in &node.attributes {
        output.push(b' ');
        output.extend_from_slice(attribute.qualified_name.as_bytes());
        output.extend_from_slice(b"=\"");
        escape_attribute(output, &attribute.value);
        output.push(b'\"');
    }
    if node.children.is_empty() && node.text.is_empty() {
        output.extend_from_slice(b"/>");
        return Ok(());
    }
    output.push(b'>');
    escape_text(output, &node.text);
    for child in &node.children {
        write_xml_node(output, child)?;
    }
    output.extend_from_slice(b"</");
    output.extend_from_slice(node.qualified_name.as_bytes());
    output.push(b'>');
    Ok(())
}

fn write_namespaces(
    output: &mut Vec<u8>,
    namespaces: &[QueryTableExtensionAttribute],
    skip_default: bool,
) -> Result<()> {
    for namespace in namespaces {
        if skip_default && namespace.qualified_name == "xmlns" {
            continue;
        }
        if namespace.qualified_name == "xmlns"
            && matches!(namespace.value.as_str(), TRANSITIONAL | STRICT)
        {
            continue;
        }
        write_named_attribute(output, &namespace.qualified_name, &namespace.value)?;
    }
    Ok(())
}

fn write_extra_attributes(
    output: &mut Vec<u8>,
    attributes: &[QueryTableExtensionAttribute],
) -> Result<()> {
    for attribute in attributes {
        write_named_attribute(output, &attribute.qualified_name, &attribute.value)?;
    }
    Ok(())
}

fn write_named_attribute(output: &mut Vec<u8>, name: &str, value: &str) -> Result<()> {
    if name.is_empty()
        || name.bytes().any(|value| {
            value.is_ascii_whitespace() || matches!(value, b'<' | b'>' | b'\'' | b'\"' | b'=')
        })
    {
        return Err(invalid("invalid preserved XML attribute name"));
    }
    output.push(b' ');
    output.extend_from_slice(name.as_bytes());
    output.extend_from_slice(b"=\"");
    escape_attribute(output, value);
    output.push(b'\"');
    Ok(())
}

fn write_string_attribute(output: &mut Vec<u8>, name: &str, value: &str) {
    output.push(b' ');
    output.extend_from_slice(name.as_bytes());
    output.extend_from_slice(b"=\"");
    escape_attribute(output, value);
    output.push(b'\"');
}

fn write_bool_attribute(output: &mut Vec<u8>, name: &str, value: Option<bool>) {
    if let Some(value) = value {
        write_string_attribute(output, name, if value { "1" } else { "0" });
    }
}

fn write_u32_attribute(output: &mut Vec<u8>, name: &str, value: Option<u32>) {
    if let Some(value) = value {
        write_string_attribute(output, name, &value.to_string());
    }
}

fn escape_attribute(output: &mut Vec<u8>, value: &str) {
    for character in value.chars() {
        match character {
            '&' => output.extend_from_slice(b"&amp;"),
            '<' => output.extend_from_slice(b"&lt;"),
            '"' => output.extend_from_slice(b"&quot;"),
            '\t' => output.extend_from_slice(b"&#x9;"),
            '\n' => output.extend_from_slice(b"&#xA;"),
            '\r' => output.extend_from_slice(b"&#xD;"),
            _ => {
                let mut buffer = [0; 4];
                output.extend_from_slice(character.encode_utf8(&mut buffer).as_bytes());
            },
        }
    }
}

fn escape_text(output: &mut Vec<u8>, value: &str) {
    for character in value.chars() {
        match character {
            '&' => output.extend_from_slice(b"&amp;"),
            '<' => output.extend_from_slice(b"&lt;"),
            '>' => output.extend_from_slice(b"&gt;"),
            _ => {
                let mut buffer = [0; 4];
                output.extend_from_slice(character.encode_utf8(&mut buffer).as_bytes());
            },
        }
    }
}

fn namespace_attributes(node: &XmlNode) -> Vec<QueryTableExtensionAttribute> {
    node.attributes
        .iter()
        .filter(|attribute| {
            is_namespace_attribute(&attribute.qualified_name)
                && !(attribute.qualified_name == "xmlns"
                    && matches!(attribute.value.as_str(), TRANSITIONAL | STRICT))
        })
        .map(|attribute| QueryTableExtensionAttribute {
            qualified_name: attribute.qualified_name.clone(),
            value: attribute.value.clone(),
        })
        .collect()
}

fn is_namespace_attribute(name: &str) -> bool {
    name == "xmlns" || name.starts_with("xmlns:")
}
fn is_core(node: &XmlNode) -> bool {
    matches!(node.namespace.as_str(), TRANSITIONAL | STRICT)
}

fn require_element(node: &XmlNode, local_name: &str) -> Result<()> {
    if is_core(node) && node.local_name == local_name {
        Ok(())
    } else {
        Err(invalid(format!(
            "expected '{local_name}', found '{}'",
            node.qualified_name
        )))
    }
}

fn require_whitespace(node: &XmlNode) -> Result<()> {
    if node.text.trim().is_empty() {
        Ok(())
    } else {
        Err(invalid(format!(
            "unexpected text in '{}'",
            node.qualified_name
        )))
    }
}

fn require_leaf(node: &XmlNode) -> Result<()> {
    require_whitespace(node)?;
    if node.children.is_empty() {
        Ok(())
    } else {
        Err(invalid(format!("'{}' must be empty", node.qualified_name)))
    }
}

fn bounded_text(value: &str) -> Result<()> {
    if value.len() <= MAX_TEXT_BYTES {
        Ok(())
    } else {
        Err(invalid("query-table string is too large"))
    }
}

fn parse_range(value: &str) -> Result<()> {
    let mut parts = value.split(':');
    let first = parse_cell(parts.next().unwrap_or(""))?;
    let last = parts.next().map(parse_cell).transpose()?.unwrap_or(first);
    if parts.next().is_some() || first.0 > last.0 || first.1 > last.1 {
        return Err(invalid(format!("invalid sort range '{value}'")));
    }
    Ok(())
}

fn parse_cell(value: &str) -> Result<(u32, u32)> {
    let bytes = value.as_bytes();
    let mut index = 0usize;
    if bytes.get(index) == Some(&b'$') {
        index += 1;
    }
    let column_start = index;
    while bytes.get(index).is_some_and(u8::is_ascii_alphabetic) {
        index += 1;
    }
    if index == column_start {
        return Err(invalid(format!("invalid cell reference '{value}'")));
    }
    let mut column = 0u32;
    for byte in &bytes[column_start..index] {
        column = column
            .checked_mul(26)
            .and_then(|current| {
                current.checked_add(u32::from(byte.to_ascii_uppercase() - b'A' + 1))
            })
            .ok_or_else(|| invalid("cell column overflow"))?;
    }
    if bytes.get(index) == Some(&b'$') {
        index += 1;
    }
    let row = std::str::from_utf8(&bytes[index..])
        .ok()
        .and_then(|value| value.parse::<u32>().ok())
        .ok_or_else(|| invalid(format!("invalid cell reference '{value}'")))?;
    if !(1..=16_384).contains(&column) || !(1..=1_048_576).contains(&row) {
        return Err(invalid(format!("cell reference '{value}' is out of range")));
    }
    Ok((column, row))
}

fn xml_error(error: impl std::fmt::Display) -> OoxmlError {
    OoxmlError::Xml(error.to_string())
}
fn invalid(message: impl Into<String>) -> OoxmlError {
    OoxmlError::InvalidFormat(message.into())
}

/// Load and validate every inert query-table part owned by a worksheet.
pub fn load_worksheet_query_tables(
    package: &OpcPackage,
    worksheet_part: &PackURI,
) -> Result<Vec<WorksheetQueryTable>> {
    let worksheet = package.get_part(worksheet_part)?;
    let mut result = Vec::new();
    let mut seen_parts = HashSet::new();
    for relationship in worksheet
        .rels()
        .iter()
        .filter(|relationship| is_query_table_relationship_type(relationship.reltype()))
    {
        if relationship.is_external() {
            return Err(invalid("query-table relationship cannot be external"));
        }
        let part_name = relationship.target_partname()?;
        if !seen_parts.insert(part_name.clone()) {
            return Err(invalid("worksheet has duplicate query-table targets"));
        }
        let part = package.get_part(&part_name)?;
        if part.content_type() != QUERY_TABLE_CONTENT_TYPE {
            return Err(invalid(format!(
                "query-table part '{}' has invalid content type",
                part_name
            )));
        }
        if part.rels().iter().next().is_some() {
            return Err(invalid("query-table parts must not have relationships"));
        }
        result.push(WorksheetQueryTable::new(
            relationship.r_id().to_string(),
            part_name.to_string(),
            parse_query_table(part.blob())?,
        ));
    }
    result.sort_by(|left, right| left.relationship_id.cmp(&right.relationship_id));
    Ok(result)
}

pub fn find_worksheet_query_table(
    package: &OpcPackage,
    worksheet_part: &PackURI,
    relationship_id: &str,
) -> Result<Option<WorksheetQueryTable>> {
    Ok(load_worksheet_query_tables(package, worksheet_part)?
        .into_iter()
        .find(|item| item.relationship_id == relationship_id))
}

/// Add an inert query-table part. The referenced connection must already exist.
pub fn add_worksheet_query_table(
    package: &mut OpcPackage,
    worksheet_part: &PackURI,
    query_table: QueryTable,
    conformance: QueryTableConformance,
) -> Result<WorksheetQueryTable> {
    validate_query_table_connection(package, query_table.connection_id)?;
    let xml = query_table.to_xml(conformance)?;
    parse_query_table(&xml)?;
    let part_name = next_query_table_part_name(package)?;
    let relationship_id = next_query_table_relationship_id(package, worksheet_part)?;
    let target = part_name.relative_ref(worksheet_part.base_uri());
    let relationship_type = match conformance {
        QueryTableConformance::Transitional => QUERY_TABLE_RELATIONSHIP_TYPE,
        QueryTableConformance::Strict => STRICT_QUERY_TABLE_RELATIONSHIP_TYPE,
    };
    package.try_add_part(Box::new(BlobPart::new(
        part_name.clone(),
        QUERY_TABLE_CONTENT_TYPE.into(),
        xml,
    )))?;
    package.get_part_mut(worksheet_part)?.rels_mut().add_relationship(
        relationship_type.into(),
        target,
        relationship_id.clone(),
        false,
    );
    let _ = package.clear_digital_signatures();
    Ok(WorksheetQueryTable::new(
        relationship_id,
        part_name.to_string(),
        query_table,
    ))
}

pub fn replace_worksheet_query_table(
    package: &mut OpcPackage,
    worksheet_part: &PackURI,
    relationship_id: &str,
    query_table: QueryTable,
    conformance: QueryTableConformance,
) -> Result<()> {
    validate_query_table_connection(package, query_table.connection_id)?;
    let existing = find_worksheet_query_table(package, worksheet_part, relationship_id)?
        .ok_or_else(|| invalid("query-table relationship was not found"))?;
    let xml = query_table.to_xml(conformance)?;
    parse_query_table(&xml)?;
    let part_name = PackURI::new(existing.part_name()).map_err(invalid)?;
    package.add_part(Box::new(BlobPart::new(
        part_name,
        QUERY_TABLE_CONTENT_TYPE.into(),
        xml,
    )));
    let _ = package.clear_digital_signatures();
    Ok(())
}

pub fn update_worksheet_query_table(
    package: &mut OpcPackage,
    worksheet_part: &PackURI,
    relationship_id: &str,
    query_table: QueryTable,
    conformance: QueryTableConformance,
) -> Result<()> {
    replace_worksheet_query_table(
        package,
        worksheet_part,
        relationship_id,
        query_table,
        conformance,
    )
}

pub fn remove_worksheet_query_table(
    package: &mut OpcPackage,
    worksheet_part: &PackURI,
    relationship_id: &str,
) -> Result<bool> {
    let Some(existing) = find_worksheet_query_table(package, worksheet_part, relationship_id)? else {
        return Ok(false);
    };
    let part_name = PackURI::new(existing.part_name()).map_err(invalid)?;
    package
        .get_part_mut(worksheet_part)?
        .rels_mut()
        .remove(relationship_id);
    if !package_part_is_referenced(package, &part_name) {
        package.remove_part(&part_name);
    }
    let _ = package.clear_digital_signatures();
    Ok(true)
}

/// Reassign deterministic query-table relationship IDs in caller-specified order.
pub fn reorder_worksheet_query_tables(
    package: &mut OpcPackage,
    worksheet_part: &PackURI,
    ordered_relationship_ids: &[String],
) -> Result<Vec<WorksheetQueryTable>> {
    let existing = load_worksheet_query_tables(package, worksheet_part)?;
    if existing.len() != ordered_relationship_ids.len() {
        return Err(invalid("query-table reorder must contain every relationship"));
    }
    let existing_ids = existing
        .iter()
        .map(|item| item.relationship_id.clone())
        .collect::<HashSet<_>>();
    let ordered_ids = ordered_relationship_ids.iter().cloned().collect::<HashSet<_>>();
    if existing_ids != ordered_ids || ordered_ids.len() != ordered_relationship_ids.len() {
        return Err(invalid("query-table reorder is not a permutation"));
    }
    let mut ordered = Vec::with_capacity(existing.len());
    for id in ordered_relationship_ids {
        ordered.push(
            existing
                .iter()
                .find(|item| &item.relationship_id == id)
                .expect("permutation was validated")
                .clone(),
        );
    }
    let relationship_type = package
        .get_part(worksheet_part)?
        .rels()
        .iter()
        .find(|relationship| is_query_table_relationship_type(relationship.reltype()))
        .map(|relationship| relationship.reltype().to_string())
        .unwrap_or_else(|| QUERY_TABLE_RELATIONSHIP_TYPE.into());
    let worksheet = package.get_part_mut(worksheet_part)?;
    for item in &existing {
        worksheet.rels_mut().remove(&item.relationship_id);
    }
    let mut result = Vec::with_capacity(ordered.len());
    for (offset, item) in ordered.into_iter().enumerate() {
        let id = format!("rIdQueryTable{}", offset + 1);
        let part_name = PackURI::new(item.part_name()).map_err(invalid)?;
        worksheet.rels_mut().add_relationship(
            relationship_type.clone(),
            part_name.relative_ref(worksheet_part.base_uri()),
            id.clone(),
            false,
        );
        result.push(WorksheetQueryTable::new(id, item.part_name, item.query_table));
    }
    let _ = package.clear_digital_signatures();
    Ok(result)
}

fn validate_query_table_connection(package: &OpcPackage, connection_id: u32) -> Result<()> {
    let connections = crate::xlsx::connections::load_from_package(package)
        .map_err(|_| invalid("connections graph is invalid"))?
        .ok_or_else(|| invalid("query table requires a connections part"))?;
    if connections
        .connections
        .iter()
        .any(|connection| connection.id == connection_id)
    {
        Ok(())
    } else {
        Err(invalid(format!(
            "query table references missing connection ID {connection_id}"
        )))
    }
}

fn next_query_table_part_name(package: &OpcPackage) -> Result<PackURI> {
    for suffix in 1..=65_537u32 {
        let candidate = PackURI::new(&format!("/xl/queryTables/queryTable{suffix}.xml"))
            .map_err(invalid)?;
        if package.get_part(&candidate).is_err() {
            return Ok(candidate);
        }
    }
    Err(invalid("no free query-table part name"))
}

fn next_query_table_relationship_id(
    package: &OpcPackage,
    worksheet_part: &PackURI,
) -> Result<String> {
    let relationships = package.get_part(worksheet_part)?.rels();
    for suffix in 1..=65_537u32 {
        let candidate = format!("rIdQueryTable{suffix}");
        if relationships.get(&candidate).is_none() {
            return Ok(candidate);
        }
    }
    Err(invalid("no free query-table relationship ID"))
}

fn package_part_is_referenced(package: &OpcPackage, target: &PackURI) -> bool {
    package.iter_parts().any(|part| {
        part.rels().iter().any(|relationship| {
            !relationship.is_external()
                && relationship.target_partname().is_ok_and(|name| name == *target)
        })
    }) || package.rels().iter().any(|relationship| {
        !relationship.is_external()
            && relationship.target_partname().is_ok_and(|name| name == *target)
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parses_and_writes_complete_strict_query_table() {
        let xml = br#"<queryTable xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="A&amp;B" headers="1" growShrinkType="insertClear" connectionId="7"><queryTableRefresh preserveSortFilterLayout="1" minimumVersion="3" nextId="9"><queryTableFields count="2"><queryTableField id="1" name="One" tableColumnId="4"/><queryTableField id="2" dataBound="0"><extLst><ext uri="field"><z:data xmlns:z="urn:test" value="x"/></ext></extLst></queryTableField></queryTableFields><queryTableDeletedFields count="1"><deletedField name="Old"/></queryTableDeletedFields><sortState ref="A1:B9" caseSensitive="1" sortMethod="pinYin"><sortCondition ref="B2:B9" descending="1" sortBy="icon" iconSet="3Arrows" iconId="2"/></sortState><extLst><ext uri="refresh"/></extLst></queryTableRefresh><extLst><ext uri="root"/></extLst></queryTable>"#;
        let value = parse_query_table(xml).unwrap();
        assert_eq!(value.name(), "A&B");
        assert_eq!(
            value.grow_shrink_type(),
            Some(QueryTableGrowShrinkType::InsertClear)
        );
        let refresh = value.refresh().unwrap();
        assert_eq!(refresh.fields().len(), 2);
        assert_eq!(refresh.deleted_fields().unwrap(), &["Old"]);
        assert_eq!(
            refresh.sort_state().unwrap().conditions()[0].icon_id(),
            Some(2)
        );
        let strict = value.to_xml(QueryTableConformance::Strict).unwrap();
        assert!(std::str::from_utf8(&strict).unwrap().contains(STRICT));
        assert_eq!(parse_query_table(&strict).unwrap(), value);
    }

    #[test]
    fn processes_mce_fallback_and_preserves_extension_payload() {
        let xml = br#"<s:queryTable xmlns:s="http://purl.oclc.org/ooxml/spreadsheetml/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:q="urn:future" mc:Ignorable="q" name="MCE" connectionId="4"><mc:AlternateContent><mc:Choice Requires="q"><q:future/></mc:Choice><mc:Fallback><s:queryTableRefresh><s:queryTableFields count="1"><s:queryTableField id="3"/></s:queryTableFields></s:queryTableRefresh></mc:Fallback></mc:AlternateContent><s:extLst><s:ext uri="u"><q:payload value="safe"/></s:ext></s:extLst></s:queryTable>"#;
        let value = parse_query_table(xml).unwrap();
        assert_eq!(value.refresh().unwrap().fields()[0].id(), 3);
        assert_eq!(
            value
                .extension_list()
                .unwrap()
                .extension_uris()
                .collect::<Vec<_>>(),
            vec!["u"]
        );
        let round_trip = value.to_xml(QueryTableConformance::Strict).unwrap();
        assert!(
            !std::str::from_utf8(&round_trip)
                .unwrap()
                .contains("q:payload")
        );
        assert_eq!(parse_query_table(&round_trip).unwrap(), value);
    }

    #[test]
    fn rejects_malformed_and_resource_abuse() {
        let invalid_documents = [
            r#"<queryTable xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="x"/>"#,
            r#"<queryTable xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="x" connectionId="1" headers="maybe"/>"#,
            r#"<queryTable xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="x" connectionId="1" growShrinkType="append"/>"#,
            r#"<queryTable xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="x" connectionId="1"><queryTableRefresh/></queryTable>"#,
            r#"<queryTable xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="x" connectionId="1"><queryTableRefresh><queryTableFields count="2"><queryTableField id="1"/></queryTableFields></queryTableRefresh></queryTable>"#,
            r#"<queryTable xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="x" connectionId="1"><queryTableRefresh><queryTableFields><queryTableField id="1"/><queryTableField id="1"/></queryTableFields></queryTableRefresh></queryTable>"#,
            r#"<queryTable xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="x" connectionId="1"><queryTableRefresh><queryTableFields/><queryTableDeletedFields/></queryTableRefresh></queryTable>"#,
            r#"<queryTable xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="x" connectionId="1"><queryTableRefresh><sortState ref="A1"/><queryTableFields/></queryTableRefresh></queryTable>"#,
            r#"<queryTable xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="x" connectionId="1"><queryTableRefresh><queryTableFields/><sortState ref="B2:A1"/></queryTableRefresh></queryTable>"#,
            r#"<queryTable xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="x" connectionId="1"><queryTableRefresh><queryTableFields/><sortState ref="A1"><sortCondition ref="A1" sortBy="icon"/></sortState></queryTableRefresh></queryTable>"#,
            r#"<!DOCTYPE x><queryTable xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="x" connectionId="1"/>"#,
            r#"<?bad x?><queryTable xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="x" connectionId="1"/>"#,
        ];
        for xml in invalid_documents {
            assert!(parse_query_table(xml.as_bytes()).is_err(), "{xml}");
        }
        assert!(parse_query_table(&vec![b' '; MAX_PART_BYTES + 1]).is_err());
    }

    #[test]
    fn loads_real_libreoffice_query_table_parts_through_workbook() {
        let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../..");
        let workbook = crate::xlsx::Workbook::open(
            root.join("test-data/poi/test-data/spreadsheet/StructuredRefs-lots-with-lookups.xlsx"),
        )
        .unwrap();
        let tables = workbook.query_tables_on_sheet("Query").unwrap();
        assert_eq!(tables.len(), 1);
        assert_eq!(tables[0].query_table().connection_id(), 2);
        assert_eq!(tables[0].query_table().name(), "Query from RDS");
    }
}
