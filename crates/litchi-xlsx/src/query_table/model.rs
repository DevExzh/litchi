//! Semantic SpreadsheetML query-table model.
//!
//! The surrounding module supplies the contextual owner name. Historical
//! `QueryTable*` names remain aliases in `mod.rs` for source compatibility.

use super::codec::write_query_table;
use crate::error::{Error, Result};

pub(crate) const TRANSITIONAL: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
pub(crate) const STRICT: &str = "http://purl.oclc.org/ooxml/spreadsheetml/main";

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Conformance {
    Transitional,
    Strict,
}

impl Conformance {
    pub(crate) fn namespace(self) -> &'static str {
        match self {
            Self::Transitional => TRANSITIONAL,
            Self::Strict => STRICT,
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum GrowShrinkType {
    InsertDelete,
    InsertClear,
    OverwriteClear,
}

impl GrowShrinkType {
    pub(crate) fn parse(value: &str) -> Result<Self> {
        match value {
            "insertDelete" => Ok(Self::InsertDelete),
            "insertClear" => Ok(Self::InsertClear),
            "overwriteClear" => Ok(Self::OverwriteClear),
            _ => Err(invalid(format!("invalid growShrinkType '{value}'"))),
        }
    }

    pub(crate) fn as_str(self) -> &'static str {
        match self {
            Self::InsertDelete => "insertDelete",
            Self::InsertClear => "insertClear",
            Self::OverwriteClear => "overwriteClear",
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SortMethod {
    Stroke,
    PinYin,
    None,
}

impl SortMethod {
    pub(crate) fn parse(value: &str) -> Result<Self> {
        match value {
            "stroke" => Ok(Self::Stroke),
            "pinYin" => Ok(Self::PinYin),
            "none" => Ok(Self::None),
            _ => Err(invalid(format!("invalid sortMethod '{value}'"))),
        }
    }

    pub(crate) fn as_str(self) -> &'static str {
        match self {
            Self::Stroke => "stroke",
            Self::PinYin => "pinYin",
            Self::None => "none",
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SortBy {
    Value,
    CellColor,
    FontColor,
    Icon,
}

impl SortBy {
    pub(crate) fn parse(value: &str) -> Result<Self> {
        match value {
            "value" => Ok(Self::Value),
            "cellColor" => Ok(Self::CellColor),
            "fontColor" => Ok(Self::FontColor),
            "icon" => Ok(Self::Icon),
            _ => Err(invalid(format!("invalid sortBy '{value}'"))),
        }
    }

    pub(crate) fn as_str(self) -> &'static str {
        match self {
            Self::Value => "value",
            Self::CellColor => "cellColor",
            Self::FontColor => "fontColor",
            Self::Icon => "icon",
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum IconSet {
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

impl IconSet {
    pub(crate) fn parse(value: &str) -> Result<Self> {
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

    pub(crate) fn as_str(self) -> &'static str {
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

    pub(crate) fn cardinality(self) -> u32 {
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
pub struct ExtensionAttribute {
    pub(crate) qualified_name: String,
    pub(crate) value: String,
}

impl ExtensionAttribute {
    pub fn qualified_name(&self) -> &str {
        &self.qualified_name
    }

    pub fn value(&self) -> &str {
        &self.value
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ExtensionList {
    pub(crate) namespaces: Vec<ExtensionAttribute>,
    pub(crate) attributes: Vec<ExtensionAttribute>,
    pub(crate) extensions: Vec<Extension>,
}

impl ExtensionList {
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
pub(crate) struct XmlAttribute {
    pub(crate) qualified_name: String,
    pub(crate) value: String,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct XmlNode {
    pub(crate) qualified_name: String,
    pub(crate) local_name: String,
    pub(crate) namespace: String,
    pub(crate) attributes: Vec<XmlAttribute>,
    pub(crate) children: Vec<XmlNode>,
    pub(crate) text: String,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct Extension {
    pub(crate) uri: String,
    pub(crate) namespaces: Vec<ExtensionAttribute>,
    pub(crate) attributes: Vec<ExtensionAttribute>,
    pub(crate) content: Vec<XmlNode>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SortCondition {
    pub(crate) reference: String,
    pub(crate) descending: Option<bool>,
    pub(crate) sort_by: Option<SortBy>,
    pub(crate) custom_list: Option<String>,
    pub(crate) differential_format_id: Option<u32>,
    pub(crate) icon_set: Option<IconSet>,
    pub(crate) icon_id: Option<u32>,
    pub(crate) extension_attributes: Vec<ExtensionAttribute>,
}

impl SortCondition {
    pub fn reference(&self) -> &str {
        &self.reference
    }
    pub fn descending(&self) -> Option<bool> {
        self.descending
    }
    pub fn sort_by(&self) -> Option<SortBy> {
        self.sort_by
    }
    pub fn custom_list(&self) -> Option<&str> {
        self.custom_list.as_deref()
    }
    pub fn differential_format_id(&self) -> Option<u32> {
        self.differential_format_id
    }
    pub fn icon_set(&self) -> Option<IconSet> {
        self.icon_set
    }
    pub fn icon_id(&self) -> Option<u32> {
        self.icon_id
    }
    pub fn extension_attributes(&self) -> &[ExtensionAttribute] {
        &self.extension_attributes
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SortState {
    pub(crate) reference: String,
    pub(crate) column_sort: Option<bool>,
    pub(crate) case_sensitive: Option<bool>,
    pub(crate) sort_method: Option<SortMethod>,
    pub(crate) conditions: Vec<SortCondition>,
    pub(crate) extension_attributes: Vec<ExtensionAttribute>,
}

impl SortState {
    pub fn reference(&self) -> &str {
        &self.reference
    }
    pub fn column_sort(&self) -> Option<bool> {
        self.column_sort
    }
    pub fn case_sensitive(&self) -> Option<bool> {
        self.case_sensitive
    }
    pub fn sort_method(&self) -> Option<SortMethod> {
        self.sort_method
    }
    pub fn conditions(&self) -> &[SortCondition] {
        &self.conditions
    }
    pub fn extension_attributes(&self) -> &[ExtensionAttribute] {
        &self.extension_attributes
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Field {
    pub(crate) id: u32,
    pub(crate) name: Option<String>,
    pub(crate) data_bound: Option<bool>,
    pub(crate) row_numbers: Option<bool>,
    pub(crate) fill_formulas: Option<bool>,
    pub(crate) clipped: Option<bool>,
    pub(crate) table_column_id: Option<u32>,
    pub(crate) extension_list: Option<ExtensionList>,
    pub(crate) extension_attributes: Vec<ExtensionAttribute>,
}

impl Field {
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
    pub fn set_name(&mut self, value: Option<String>) {
        self.name = value;
    }
    pub fn set_data_bound(&mut self, value: Option<bool>) {
        self.data_bound = value;
    }
    pub fn set_row_numbers(&mut self, value: Option<bool>) {
        self.row_numbers = value;
    }
    pub fn set_fill_formulas(&mut self, value: Option<bool>) {
        self.fill_formulas = value;
    }
    pub fn set_clipped(&mut self, value: Option<bool>) {
        self.clipped = value;
    }
    pub fn set_table_column_id(&mut self, value: Option<u32>) {
        self.table_column_id = value;
    }
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
    pub fn extension_list(&self) -> Option<&ExtensionList> {
        self.extension_list.as_ref()
    }
    pub fn extension_attributes(&self) -> &[ExtensionAttribute] {
        &self.extension_attributes
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Refresh {
    pub(crate) preserve_sort_filter_layout: Option<bool>,
    pub(crate) field_id_wrapped: Option<bool>,
    pub(crate) headers_in_last_refresh: Option<bool>,
    pub(crate) minimum_version: Option<u8>,
    pub(crate) next_id: Option<u32>,
    pub(crate) unbound_columns_left: Option<u32>,
    pub(crate) unbound_columns_right: Option<u32>,
    pub(crate) declared_field_count: Option<u32>,
    pub(crate) fields: Vec<Field>,
    pub(crate) declared_deleted_field_count: Option<u32>,
    pub(crate) deleted_fields: Option<Vec<String>>,
    pub(crate) sort_state: Option<SortState>,
    pub(crate) extension_list: Option<ExtensionList>,
    pub(crate) extension_attributes: Vec<ExtensionAttribute>,
}

impl Default for Refresh {
    fn default() -> Self {
        Self::new()
    }
}

impl Refresh {
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
    pub fn fields_mut(&mut self) -> &mut Vec<Field> {
        &mut self.fields
    }
    pub fn add_field(&mut self, field: Field) {
        self.fields.push(field);
    }
    pub fn deleted_fields_mut(&mut self) -> &mut Option<Vec<String>> {
        &mut self.deleted_fields
    }
    pub fn add_deleted_field(&mut self, name: String) {
        self.deleted_fields.get_or_insert_with(Vec::new).push(name);
    }
    pub fn set_preserve_sort_filter_layout(&mut self, value: Option<bool>) {
        self.preserve_sort_filter_layout = value;
    }
    pub fn set_field_id_wrapped(&mut self, value: Option<bool>) {
        self.field_id_wrapped = value;
    }
    pub fn set_headers_in_last_refresh(&mut self, value: Option<bool>) {
        self.headers_in_last_refresh = value;
    }
    pub fn set_minimum_version(&mut self, value: Option<u8>) {
        self.minimum_version = value;
    }
    pub fn set_next_id(&mut self, value: Option<u32>) {
        self.next_id = value;
    }
    pub fn set_unbound_columns_left(&mut self, value: Option<u32>) {
        self.unbound_columns_left = value;
    }
    pub fn set_unbound_columns_right(&mut self, value: Option<u32>) {
        self.unbound_columns_right = value;
    }
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
    pub fn fields(&self) -> &[Field] {
        &self.fields
    }
    pub fn declared_deleted_field_count(&self) -> Option<u32> {
        self.declared_deleted_field_count
    }
    pub fn deleted_fields(&self) -> Option<&[String]> {
        self.deleted_fields.as_deref()
    }
    pub fn sort_state(&self) -> Option<&SortState> {
        self.sort_state.as_ref()
    }
    pub fn extension_list(&self) -> Option<&ExtensionList> {
        self.extension_list.as_ref()
    }
    pub fn extension_attributes(&self) -> &[ExtensionAttribute] {
        &self.extension_attributes
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Table {
    pub(crate) name: String,
    pub(crate) headers: Option<bool>,
    pub(crate) row_numbers: Option<bool>,
    pub(crate) disable_refresh: Option<bool>,
    pub(crate) background_refresh: Option<bool>,
    pub(crate) first_background_refresh: Option<bool>,
    pub(crate) refresh_on_load: Option<bool>,
    pub(crate) grow_shrink_type: Option<GrowShrinkType>,
    pub(crate) fill_formulas: Option<bool>,
    pub(crate) remove_data_on_save: Option<bool>,
    pub(crate) disable_edit: Option<bool>,
    pub(crate) preserve_formatting: Option<bool>,
    pub(crate) adjust_column_width: Option<bool>,
    pub(crate) intermediate: Option<bool>,
    pub(crate) connection_id: u32,
    pub(crate) auto_format_id: Option<u32>,
    pub(crate) apply_number_formats: Option<bool>,
    pub(crate) apply_border_formats: Option<bool>,
    pub(crate) apply_font_formats: Option<bool>,
    pub(crate) apply_pattern_formats: Option<bool>,
    pub(crate) apply_alignment_formats: Option<bool>,
    pub(crate) apply_width_height_formats: Option<bool>,
    pub(crate) refresh: Option<Refresh>,
    pub(crate) extension_list: Option<ExtensionList>,
    pub(crate) namespaces: Vec<ExtensionAttribute>,
    pub(crate) extension_attributes: Vec<ExtensionAttribute>,
}

impl Table {
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
    pub fn set_name(&mut self, value: String) {
        self.name = value;
    }
    pub fn set_connection_id(&mut self, value: u32) {
        self.connection_id = value;
    }
    pub fn set_headers(&mut self, value: Option<bool>) {
        self.headers = value;
    }
    pub fn set_row_numbers(&mut self, value: Option<bool>) {
        self.row_numbers = value;
    }
    pub fn set_disable_refresh(&mut self, value: Option<bool>) {
        self.disable_refresh = value;
    }
    pub fn set_background_refresh(&mut self, value: Option<bool>) {
        self.background_refresh = value;
    }
    pub fn set_first_background_refresh(&mut self, value: Option<bool>) {
        self.first_background_refresh = value;
    }
    pub fn set_refresh_on_load(&mut self, value: Option<bool>) {
        self.refresh_on_load = value;
    }
    pub fn set_grow_shrink_type(&mut self, value: Option<GrowShrinkType>) {
        self.grow_shrink_type = value;
    }
    pub fn set_fill_formulas(&mut self, value: Option<bool>) {
        self.fill_formulas = value;
    }
    pub fn set_remove_data_on_save(&mut self, value: Option<bool>) {
        self.remove_data_on_save = value;
    }
    pub fn set_disable_edit(&mut self, value: Option<bool>) {
        self.disable_edit = value;
    }
    pub fn set_preserve_formatting(&mut self, value: Option<bool>) {
        self.preserve_formatting = value;
    }
    pub fn set_adjust_column_width(&mut self, value: Option<bool>) {
        self.adjust_column_width = value;
    }
    pub fn set_intermediate(&mut self, value: Option<bool>) {
        self.intermediate = value;
    }
    pub fn set_refresh(&mut self, value: Option<Refresh>) {
        self.refresh = value;
    }
    pub fn refresh_mut(&mut self) -> Option<&mut Refresh> {
        self.refresh.as_mut()
    }
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
    pub fn grow_shrink_type(&self) -> Option<GrowShrinkType> {
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
    pub fn refresh(&self) -> Option<&Refresh> {
        self.refresh.as_ref()
    }
    pub fn extension_list(&self) -> Option<&ExtensionList> {
        self.extension_list.as_ref()
    }
    pub fn extension_attributes(&self) -> &[ExtensionAttribute] {
        &self.extension_attributes
    }

    pub fn to_xml(&self, conformance: Conformance) -> Result<Vec<u8>> {
        write_query_table(self, conformance)
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct WorksheetTable {
    pub(crate) relationship_id: String,
    pub(crate) part_name: String,
    pub(crate) query_table: Table,
}

impl WorksheetTable {
    pub fn new(relationship_id: String, part_name: String, query_table: Table) -> Self {
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
    pub fn query_table(&self) -> &Table {
        &self.query_table
    }
}

pub(crate) fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}
