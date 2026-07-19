//! Inert reader for Microsoft ChartEx (`cx:chartSpace`) parts.

use super::chart_ex_style::{ChartColorStyleDocument, ChartStyleDocument, discover_chart_styles};
use crate::error::{OoxmlError, Result};
use litchi_opc::OpcPackage;
use litchi_opc::part::Part;
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use std::collections::HashSet;

pub const CHART_EX_CONTENT_TYPE: &str = "application/vnd.ms-office.chartex+xml";

const CX: &str = "http://schemas.microsoft.com/office/drawing/2014/chartex";
const A: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const A_STRICT: &str = "http://purl.oclc.org/ooxml/drawingml/main";
const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const R_STRICT: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships";
const PACKAGE_REL: [&str; 2] = [
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package",
    "http://purl.oclc.org/ooxml/officeDocument/relationships/package",
];
const OLE_REL: [&str; 2] = [
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject",
    "http://purl.oclc.org/ooxml/officeDocument/relationships/oleObject",
];
const IMAGE_REL: [&str; 2] = [
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
    "http://purl.oclc.org/ooxml/officeDocument/relationships/image",
];
const OLE_CONTENT_TYPE: &str = "application/vnd.openxmlformats-officedocument.oleObject";
const WORKBOOK_CONTENT_TYPES: [&str; 3] = [
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    "application/vnd.ms-excel.sheet.macroEnabled.12",
    "application/vnd.ms-excel.sheet.binary.macroEnabled.12",
];

const MAX_XML_BYTES: usize = 64 * 1024 * 1024;
const MAX_NODES: usize = 250_000;
const MAX_DEPTH: usize = 128;
const MAX_ATTRIBUTES: usize = 64;
const MAX_STRING_BYTES: usize = 8 * 1024 * 1024;
const MAX_DATA_SETS: usize = 65_536;
const MAX_FEATURES: usize = 256;
const MAX_LEVELS_PER_DIMENSION: usize = 4096;
const MAX_POINTS_PER_LEVEL: u32 = 1_000_000;
const MAX_FORMULA_BYTES: usize = 32 * 1024;
const MAX_SERIES: usize = 65_536;
const MAX_AXES: usize = 4_096;
const MAX_AXIS_REFS_PER_SERIES: usize = 64;
const MAX_SUBTOTALS: usize = 100_000;
const MAX_CULTURE_NAME_LEN: usize = 64;
const MAX_ATTRIBUTION_LEN: usize = 4_096;
const MAX_GEO_STRING_LEN: usize = 8_192;
const MAX_GEO_POLYGON_DATA_LEN: usize = 1024 * 1024;
const MAX_GEO_RESULTS: usize = 65_536;
const MAX_GEO_CACHE_ENTRIES: usize = 1_024;
const MAX_GEO_BINARY_BYTES: usize = 1024 * 1024;
const MAX_SERIES_POINTS: usize = 100_000;
const MAX_DATA_LABELS: usize = 100_000;
const MAX_LABEL_TEXT_BYTES: usize = 32 * 1024;
const MAX_FORMAT_OVERRIDES: usize = 65_536;
const MAX_PRINT_TEXT_BYTES: usize = 32 * 1024;

/// A validated ChartEx part wrapper.
pub struct ChartExPart<'a> {
    part: &'a dyn Part,
}

/// Typed metadata from the bounded ChartEx container and data-index core.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChartExInfo {
    pub version: String,
    pub features: Vec<String>,
    pub fallback_image_relationship_id: Option<String>,
    pub data_sets: Vec<ChartExDataSet>,
    pub series: Vec<ChartExSeriesDataReference>,
    pub axes: Vec<ChartExAxis>,
    pub has_plot_surface: bool,
    pub chart: ChartExChart,
    pub plot_area: ChartExPlotArea,
    pub chart_space_formatting: ChartExChartSpaceFormatting,
    pub external_data: Option<ChartExExternalData>,
    pub has_title: bool,
    pub has_legend: bool,
}

#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct ChartExPlotArea {
    pub region: ChartExPlotAreaRegion,
    pub shape_properties: Option<ChartExDrawingPayload>,
    pub has_extension_list: bool,
}

#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct ChartExPlotAreaRegion {
    pub plot_surface: Option<ChartExPlotSurface>,
    pub has_extension_list: bool,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChartExPlotSurface {
    pub shape_properties: Option<ChartExDrawingPayload>,
    pub has_extension_list: bool,
}

#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct ChartExChartSpaceFormatting {
    pub shape_properties: Option<ChartExDrawingPayload>,
    pub text_properties: Option<ChartExDrawingPayload>,
    pub color_mapping_override: Option<ChartExDrawingPayload>,
    pub format_overrides: Option<Vec<ChartExFormatOverride>>,
    pub print_settings: Option<ChartExPrintSettings>,
    pub has_extension_list: bool,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChartExFormatOverride {
    pub index: u32,
    pub shape_properties: Option<ChartExDrawingPayload>,
    pub has_extension_list: bool,
}

#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct ChartExPrintSettings {
    pub header_footer: Option<ChartExHeaderFooter>,
    pub page_margins: Option<ChartExPageMargins>,
    pub page_setup: Option<ChartExPageSetup>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChartExHeaderFooter {
    pub align_with_margins: bool,
    pub different_odd_even: bool,
    pub different_first: bool,
    pub odd_header: Option<String>,
    pub odd_footer: Option<String>,
    pub even_header: Option<String>,
    pub even_footer: Option<String>,
    pub first_header: Option<String>,
    pub first_footer: Option<String>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChartExPageMargins {
    pub left: String,
    pub right: String,
    pub top: String,
    pub bottom: String,
    pub header: String,
    pub footer: String,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ChartExPageOrientation {
    Default,
    Portrait,
    Landscape,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChartExPageSetup {
    pub paper_size: u32,
    pub first_page_number: u32,
    pub orientation: ChartExPageOrientation,
    pub black_and_white: bool,
    pub draft: bool,
    pub use_first_page_number: bool,
    pub horizontal_dpi: i32,
    pub vertical_dpi: i32,
    pub copies: u32,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChartExDataSet {
    pub id: u32,
    pub string_dimensions: usize,
    pub numeric_dimensions: usize,
    pub dimensions: Vec<ChartExDimension>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ChartExStringDimensionType {
    Category,
    ColorString,
    EntityId,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ChartExNumericDimensionType {
    Value,
    X,
    Y,
    Size,
    ColorValue,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ChartExFormulaDirection {
    Column,
    Row,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChartExFormula {
    pub expression: String,
    pub direction: ChartExFormulaDirection,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ChartExDimension {
    String {
        kind: ChartExStringDimensionType,
        formula: Option<ChartExFormula>,
        name_formula: Option<ChartExFormula>,
        levels: Vec<ChartExStringLevel>,
    },
    Numeric {
        kind: ChartExNumericDimensionType,
        formula: Option<ChartExFormula>,
        name_formula: Option<ChartExFormula>,
        levels: Vec<ChartExNumericLevel>,
    },
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChartExStringLevel {
    pub point_count: u32,
    pub name: Option<String>,
    pub points: Vec<ChartExStringPoint>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChartExNumericLevel {
    pub point_count: u32,
    pub name: Option<String>,
    pub format_code: Option<String>,
    pub points: Vec<ChartExNumericPoint>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChartExStringPoint {
    pub index: u32,
    pub value: String,
}

/// Numeric values retain their XML Schema double lexical form (`INF`, `-INF`, and `NaN` included).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChartExNumericPoint {
    pub index: u32,
    pub value: String,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ChartExSeriesLayout {
    BoxWhisker,
    ClusteredColumn,
    Funnel,
    ParetoLine,
    RegionMap,
    Sunburst,
    Treemap,
    Waterfall,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChartExSeriesDataReference {
    pub layout: ChartExSeriesLayout,
    pub text: Option<ChartExText>,
    pub shape_properties: Option<ChartExDrawingPayload>,
    pub value_colors: Option<ChartExValueColors>,
    pub value_color_positions: Option<ChartExValueColorPositions>,
    pub data_points: Vec<ChartExDataPoint>,
    pub data_labels: Option<ChartExDataLabels>,
    pub data_id: Option<u32>,
    pub hidden: bool,
    pub owner_index: Option<u32>,
    pub unique_id: Option<String>,
    pub format_index: Option<u32>,
    pub layout_properties: Option<ChartExLayoutProperties>,
    pub axis_ids: Vec<u32>,
}

/// A DrawingML subtree retained by the document's lossless source XML.
/// Only its bounded, namespace-checked outer payload is exposed here.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChartExDrawingPayload {
    pub child_elements: usize,
    pub attributes: usize,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ChartExText {
    Data {
        formula: Option<ChartExFormula>,
        value: Option<String>,
    },
    Rich(ChartExDrawingPayload),
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ChartExColorKind {
    ScRgb,
    Srgb,
    Hsl,
    System,
    Scheme,
    Preset,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChartExSolidColor {
    pub kind: ChartExColorKind,
    pub value: Option<String>,
    pub modifier_count: usize,
}

#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct ChartExValueColors {
    pub minimum: Option<ChartExSolidColor>,
    pub middle: Option<ChartExSolidColor>,
    pub maximum: Option<ChartExSolidColor>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ChartExColorPosition {
    Extreme,
    Number(String),
    Percent(String),
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChartExValueColorPositions {
    pub count: u8,
    pub minimum: Option<ChartExColorPosition>,
    pub middle: Option<ChartExColorPosition>,
    pub maximum: Option<ChartExColorPosition>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChartExDataPoint {
    pub index: u32,
    pub shape_properties: Option<ChartExDrawingPayload>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ChartExDataLabelPosition {
    BestFit,
    Bottom,
    Center,
    InsideBase,
    InsideEnd,
    Left,
    OutsideEnd,
    Right,
    Top,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChartExNumberFormat {
    pub format_code: String,
    pub source_linked: Option<bool>,
}

#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct ChartExDataLabelVisibility {
    pub series_name: Option<bool>,
    pub category_name: Option<bool>,
    pub value: Option<bool>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChartExDataLabel {
    pub index: u32,
    pub position: Option<ChartExDataLabelPosition>,
    pub number_format: Option<ChartExNumberFormat>,
    pub shape_properties: Option<ChartExDrawingPayload>,
    pub text_properties: Option<ChartExDrawingPayload>,
    pub visibility: Option<ChartExDataLabelVisibility>,
    pub separator: Option<String>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChartExDataLabels {
    pub position: Option<ChartExDataLabelPosition>,
    pub number_format: Option<ChartExNumberFormat>,
    pub shape_properties: Option<ChartExDrawingPayload>,
    pub text_properties: Option<ChartExDrawingPayload>,
    pub visibility: Option<ChartExDataLabelVisibility>,
    pub separator: Option<String>,
    pub labels: Vec<ChartExDataLabel>,
    pub hidden_indices: Vec<u32>,
}

#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct ChartExChart {
    pub title: Option<ChartExChartTitle>,
    pub legend: Option<ChartExLegend>,
    pub has_extension_list: bool,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ChartExSidePosition {
    Left,
    Right,
    Top,
    Bottom,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ChartExPositionAlignment {
    Minimum,
    Center,
    Maximum,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChartExOffset {
    pub top: String,
    pub left: String,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChartExChartTitle {
    pub position: ChartExSidePosition,
    pub alignment: ChartExPositionAlignment,
    pub overlay: bool,
    pub text: Option<ChartExText>,
    pub shape_properties: Option<ChartExDrawingPayload>,
    pub text_properties: Option<ChartExDrawingPayload>,
    pub offset: Option<ChartExOffset>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChartExLegend {
    pub position: ChartExSidePosition,
    pub alignment: ChartExPositionAlignment,
    pub overlay: bool,
    pub shape_properties: Option<ChartExDrawingPayload>,
    pub text_properties: Option<ChartExDrawingPayload>,
    pub offset: Option<ChartExOffset>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ChartExParentLabelLayout {
    None,
    Banner,
    Overlapping,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ChartExRegionLabelLayout {
    None,
    BestFitOnly,
    ShowAll,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ChartExClosedSide {
    Left,
    Right,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ChartExQuartileMethod {
    Inclusive,
    Exclusive,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ChartExGeoProjection {
    Mercator,
    Miller,
    Robinson,
    Albers,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ChartExGeoMappingLevel {
    DataOnly,
    PostalCode,
    County,
    State,
    CountryRegion,
    CountryRegionList,
    World,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ChartExDoubleOrAutomatic {
    Automatic,
    Number(String),
}

#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct ChartExElementVisibility {
    pub connector_lines: Option<bool>,
    pub mean_line: Option<bool>,
    pub mean_marker: Option<bool>,
    pub nonoutliers: Option<bool>,
    pub outliers: Option<bool>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ChartExBinningChoice {
    Size(String),
    Count(u32),
}

#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct ChartExBinning {
    pub choice: Option<ChartExBinningChoice>,
    pub interval_closed: Option<ChartExClosedSide>,
    pub underflow: Option<ChartExDoubleOrAutomatic>,
    pub overflow: Option<ChartExDoubleOrAutomatic>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChartExGeography {
    pub projection: Option<ChartExGeoProjection>,
    pub viewed_region: Option<ChartExGeoMappingLevel>,
    pub culture_language: String,
    pub culture_region: String,
    pub attribution: String,
    pub has_cache: bool,
    pub cache: Option<ChartExGeoCache>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChartExGeoCache {
    pub provider: String,
    pub entries: Vec<ChartExGeoCacheEntry>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ChartExGeoCacheEntry {
    Binary {
        encoded_characters: usize,
        decoded_bytes: usize,
    },
    Clear(ChartExGeoClear),
}

#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct ChartExGeoClear {
    pub location_query_results: Option<Vec<ChartExGeoLocationQueryResult>>,
    pub data_entity_query_results: Option<Vec<ChartExGeoDataEntityQueryResult>>,
    pub data_point_to_entity_query_results: Option<Vec<ChartExGeoDataPointToEntityQueryResult>>,
    pub child_entities_query_results: Option<Vec<ChartExGeoChildEntitiesQueryResult>>,
    pub parent_entities_query_results: Option<Vec<ChartExGeoParentEntitiesQueryResult>>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ChartExGeoEntityType {
    Address,
    AdminDistrict,
    AdminDistrict2,
    AdminDistrict3,
    Continent,
    CountryRegion,
    Locality,
    Ocean,
    Planet,
    PostalCode,
    Region,
    Unsupported,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChartExGeoLocationQuery {
    pub country_region: Option<String>,
    pub admin_district1: Option<String>,
    pub admin_district2: Option<String>,
    pub postal_code: Option<String>,
    pub entity_type: ChartExGeoEntityType,
}

#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct ChartExGeoAddress {
    pub address1: Option<String>,
    pub country_region: Option<String>,
    pub admin_district1: Option<String>,
    pub admin_district2: Option<String>,
    pub postal_code: Option<String>,
    pub locality: Option<String>,
    pub iso_country_code: Option<String>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChartExGeoLocation {
    pub latitude: Option<String>,
    pub longitude: Option<String>,
    pub entity_name: String,
    pub entity_type: ChartExGeoEntityType,
    pub address: Option<ChartExGeoAddress>,
}

#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct ChartExGeoLocationQueryResult {
    pub query: Option<ChartExGeoLocationQuery>,
    pub location: Option<ChartExGeoLocation>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChartExGeoPolygon {
    pub polygon_id: String,
    pub num_points: String,
    pub pca_rings: String,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChartExGeoData {
    pub entity_name: String,
    pub entity_id: String,
    pub east: String,
    pub west: String,
    pub north: String,
    pub south: String,
    pub polygons: Option<Vec<ChartExGeoPolygon>>,
    pub copyrights: Option<Vec<String>>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChartExGeoDataEntityQuery {
    pub entity_type: ChartExGeoEntityType,
    pub entity_id: String,
}

#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct ChartExGeoDataEntityQueryResult {
    pub query: Option<ChartExGeoDataEntityQuery>,
    pub data: Option<ChartExGeoData>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChartExGeoDataPointQuery {
    pub entity_type: ChartExGeoEntityType,
    pub latitude: String,
    pub longitude: String,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChartExGeoDataPointToEntityQuery {
    pub entity_type: ChartExGeoEntityType,
    pub entity_id: String,
}

#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct ChartExGeoDataPointToEntityQueryResult {
    pub point_query: Option<ChartExGeoDataPointQuery>,
    pub entity_query: Option<ChartExGeoDataPointToEntityQuery>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChartExGeoChildEntitiesQuery {
    pub entity_id: String,
    pub child_types: Option<Vec<ChartExGeoEntityType>>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChartExGeoHierarchyEntity {
    pub entity_name: String,
    pub entity_id: String,
    pub entity_type: ChartExGeoEntityType,
}

#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct ChartExGeoChildEntitiesQueryResult {
    pub query: Option<ChartExGeoChildEntitiesQuery>,
    pub children: Option<Vec<ChartExGeoHierarchyEntity>>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChartExGeoEntity {
    pub entity_name: String,
    pub entity_type: ChartExGeoEntityType,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChartExGeoParentEntitiesQueryResult {
    pub entity_id: String,
    pub entity: Option<ChartExGeoEntity>,
    pub parent_entity_id: Option<String>,
}

#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct ChartExLayoutProperties {
    pub parent_label: Option<ChartExParentLabelLayout>,
    pub region_label: Option<ChartExRegionLabelLayout>,
    pub visibility: Option<ChartExElementVisibility>,
    pub aggregation: bool,
    pub binning: Option<ChartExBinning>,
    pub geography: Option<ChartExGeography>,
    pub quartile_method: Option<ChartExQuartileMethod>,
    pub subtotals: Vec<u32>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ChartExAxisScaling {
    Category {
        gap_width: Option<ChartExDoubleOrAutomatic>,
    },
    Value {
        minimum: Option<ChartExDoubleOrAutomatic>,
        maximum: Option<ChartExDoubleOrAutomatic>,
        major_unit: Option<ChartExDoubleOrAutomatic>,
        minor_unit: Option<ChartExDoubleOrAutomatic>,
    },
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChartExAxisTitle {
    pub text: Option<ChartExText>,
    pub shape_properties: Option<ChartExDrawingPayload>,
    pub text_properties: Option<ChartExDrawingPayload>,
    pub offset: Option<ChartExOffset>,
    pub has_extension_list: bool,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ChartExAxisUnit {
    Hundreds,
    Thousands,
    TenThousands,
    HundredThousands,
    Millions,
    TenMillions,
    HundredMillions,
    Billions,
    Trillions,
    Percentage,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChartExAxisUnitsLabel {
    pub text: Option<ChartExText>,
    pub shape_properties: Option<ChartExDrawingPayload>,
    pub text_properties: Option<ChartExDrawingPayload>,
    pub has_extension_list: bool,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChartExAxisUnits {
    pub unit: Option<ChartExAxisUnit>,
    pub label: Option<ChartExAxisUnitsLabel>,
    pub has_extension_list: bool,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChartExGridlines {
    pub shape_properties: Option<ChartExDrawingPayload>,
    pub has_extension_list: bool,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ChartExTickMarkType {
    Inside,
    Outside,
    Cross,
    None,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChartExTickMarks {
    pub kind: Option<ChartExTickMarkType>,
    pub has_extension_list: bool,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChartExTickLabels {
    pub has_extension_list: bool,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChartExAxis {
    pub id: u32,
    pub hidden: bool,
    pub scaling: ChartExAxisScaling,
    pub title: Option<ChartExAxisTitle>,
    pub units: Option<ChartExAxisUnits>,
    pub major_gridlines: Option<ChartExGridlines>,
    pub minor_gridlines: Option<ChartExGridlines>,
    pub major_tick_marks: Option<ChartExTickMarks>,
    pub minor_tick_marks: Option<ChartExTickMarks>,
    pub tick_labels: Option<ChartExTickLabels>,
    pub number_format: Option<ChartExNumberFormat>,
    pub shape_properties: Option<ChartExDrawingPayload>,
    pub text_properties: Option<ChartExDrawingPayload>,
    pub has_extension_list: bool,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChartExExternalData {
    pub relationship_id: String,
    pub auto_update: bool,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ChartExExternalDataTarget {
    EmbeddedPackage {
        part_name: String,
        content_type: String,
    },
    OleObject {
        part_name: String,
    },
}

/// A parsed inert document. Unsupported subtrees remain byte-for-byte preserved.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChartExDocument {
    info: ChartExInfo,
    xml: Vec<u8>,
    external_data_target: Option<ChartExExternalDataTarget>,
    fallback_image_part_name: Option<String>,
    chart_style: Option<ChartStyleDocument>,
    chart_color_style: Option<ChartColorStyleDocument>,
}

impl<'a> ChartExPart<'a> {
    pub fn from_part(part: &'a dyn Part) -> Result<Self> {
        if part.content_type() != CHART_EX_CONTENT_TYPE {
            return invalid("ChartEx part has the wrong content type");
        }
        Ok(Self { part })
    }

    pub fn parse(&self) -> Result<ChartExDocument> {
        parse_document(self.part.blob())
    }

    /// Parse and validate referenced package resources without opening their bytes.
    pub fn parse_in_package(&self, package: &OpcPackage) -> Result<ChartExDocument> {
        let mut document = self.parse()?;
        if let Some(external) = &document.info.external_data {
            if external.auto_update {
                return invalid("auto-updating ChartEx external data is rejected");
            }
            document.external_data_target = Some(validate_external_data(
                package,
                self.part,
                &external.relationship_id,
            )?);
        }
        if let Some(id) = &document.info.fallback_image_relationship_id {
            document.fallback_image_part_name =
                Some(validate_fallback_image(package, self.part, id)?);
        }
        let (chart_style, chart_color_style) = discover_chart_styles(package, self.part)?;
        document.chart_style = chart_style;
        document.chart_color_style = chart_color_style;
        Ok(document)
    }

    pub fn part(&self) -> &'a dyn Part {
        self.part
    }
}

impl ChartExDocument {
    pub fn info(&self) -> &ChartExInfo {
        &self.info
    }

    pub fn external_data_target(&self) -> Option<&ChartExExternalDataTarget> {
        self.external_data_target.as_ref()
    }

    pub fn fallback_image_part_name(&self) -> Option<&str> {
        self.fallback_image_part_name.as_deref()
    }

    pub fn chart_style(&self) -> Option<&ChartStyleDocument> {
        self.chart_style.as_ref()
    }

    pub fn chart_color_style(&self) -> Option<&ChartColorStyleDocument> {
        self.chart_color_style.as_ref()
    }

    /// Return the validated source XML unchanged.
    pub fn to_xml(&self) -> Vec<u8> {
        self.xml.clone()
    }
}

#[derive(Debug)]
struct Attribute {
    namespace: String,
    name: String,
    value: String,
}

#[derive(Default)]
struct Scan {
    root_depth: Option<usize>,
    root_closed: bool,
    root_rank: Option<u8>,
    root_children: HashSet<String>,
    chart_data_depth: Option<usize>,
    chart_depth: Option<usize>,
    data_depth: Option<usize>,
    current_data: Option<ChartExDataSet>,
    data_ids: HashSet<u32>,
    leaf_depth: Option<usize>,
    info: Option<ChartExInfo>,
}

fn parse_document(xml: &[u8]) -> Result<ChartExDocument> {
    if xml.len() > MAX_XML_BYTES {
        return limit("ChartEx XML bytes");
    }
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut scan = Scan::default();
    let mut depth = 0usize;
    let mut nodes = 0usize;
    let mut strings = 0usize;
    loop {
        let event = reader.read_event().map_err(xml_error)?;
        match event {
            Event::Start(element) => {
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid_error("ChartEx depth overflow"))?;
                inspect_start(&reader, &element, depth, false, &mut scan, &mut strings)?;
            },
            Event::Empty(element) => {
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid_error("ChartEx depth overflow"))?;
                inspect_start(&reader, &element, depth, true, &mut scan, &mut strings)?;
                inspect_end(&reader, element.name(), depth, &mut scan)?;
                depth -= 1;
            },
            Event::End(element) => {
                inspect_end(&reader, element.name(), depth, &mut scan)?;
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid_error("unbalanced ChartEx XML"))?;
            },
            Event::Text(value) => add_strings(&mut strings, value.as_ref().len())?,
            Event::CData(value) => add_strings(&mut strings, value.as_ref().len())?,
            Event::GeneralRef(value) => add_strings(&mut strings, value.as_ref().len())?,
            Event::DocType(_) | Event::PI(_) => {
                return invalid("DTD and processing instructions are rejected in ChartEx XML");
            },
            Event::Eof => break,
            _ => {},
        }
        nodes += 1;
        if nodes > MAX_NODES || depth > MAX_DEPTH {
            return limit("ChartEx XML structure");
        }
    }
    if depth != 0 || !scan.root_closed || scan.root_depth.is_none() {
        return invalid("missing or unterminated cx:chartSpace root");
    }
    let mut info = scan
        .info
        .ok_or_else(|| invalid_error("missing ChartEx metadata"))?;
    if info.data_sets.is_empty() {
        return invalid("ChartEx chartData requires at least one data set");
    }
    let (data_sets, series, axes, has_plot_surface, chart, plot_area, chart_space_formatting) =
        parse_data_graph(xml, &info.version, &info.features)?;
    info.data_sets = data_sets;
    info.series = series;
    info.axes = axes;
    info.has_plot_surface = has_plot_surface;
    info.chart = chart;
    info.plot_area = plot_area;
    info.chart_space_formatting = chart_space_formatting;
    Ok(ChartExDocument {
        info,
        xml: xml.to_vec(),
        external_data_target: None,
        fallback_image_part_name: None,
        chart_style: None,
        chart_color_style: None,
    })
}

fn inspect_start(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    depth: usize,
    empty: bool,
    scan: &mut Scan,
    strings: &mut usize,
) -> Result<()> {
    let namespace = resolved(reader.resolver().resolve_element(element.name()).0)?;
    let local_name = element.local_name();
    let local = std::str::from_utf8(local_name.as_ref()).map_err(xml_error)?;
    add_strings(strings, namespace.len() + local.len())?;
    let attributes = attributes(reader, element, strings)?;

    if depth == 1 {
        if scan.root_depth.is_some() || namespace != CX || local != "chartSpace" || empty {
            return invalid("ChartEx XML must have one non-empty cx:chartSpace root");
        }
        let version = optional(&attributes, "", "version")
            .unwrap_or("0.0")
            .to_owned();
        if version.len() > 64 {
            return limit("ChartEx version bytes");
        }
        let features = parse_features(optional(&attributes, "", "featureList").unwrap_or(""))?;
        let fallback = optional(&attributes, "", "fallbackImg")
            .filter(|value| !value.is_empty())
            .map(str::to_owned);
        if let Some(id) = &fallback {
            validate_id(id)?;
        }
        reject_unknown(
            &attributes,
            &[("", "version"), ("", "featureList"), ("", "fallbackImg")],
            "chartSpace",
        )?;
        scan.info = Some(ChartExInfo {
            version,
            features,
            fallback_image_relationship_id: fallback,
            data_sets: Vec::new(),
            series: Vec::new(),
            axes: Vec::new(),
            has_plot_surface: false,
            chart: ChartExChart::default(),
            plot_area: ChartExPlotArea::default(),
            chart_space_formatting: ChartExChartSpaceFormatting::default(),
            external_data: None,
            has_title: false,
            has_legend: false,
        });
        scan.root_depth = Some(depth);
        return Ok(());
    }
    if scan.leaf_depth.is_some_and(|value| depth > value) {
        return invalid("ChartEx leaf element contains child elements");
    }
    let root_depth = scan
        .root_depth
        .ok_or_else(|| invalid_error("content before ChartEx root"))?;
    if depth == root_depth + 1 {
        let rank = root_child_rank(&namespace, local)
            .ok_or_else(|| invalid_error(format!("unsupported direct ChartEx child '{local}'")))?;
        if scan.root_rank.is_some_and(|previous| rank < previous)
            || !scan.root_children.insert(local.to_owned())
        {
            return invalid("ChartEx root children are duplicated or out of schema order");
        }
        scan.root_rank = Some(rank);
        match local {
            "chartData" => {
                if empty || rank != 0 {
                    return invalid("ChartEx chartData must be non-empty and first");
                }
                scan.chart_data_depth = Some(depth);
            },
            "chart" => {
                if empty || !scan.root_children.contains("chartData") {
                    return invalid("ChartEx chart must follow chartData");
                }
                scan.chart_depth = Some(depth);
            },
            _ => {},
        }
        return Ok(());
    }
    if scan
        .chart_data_depth
        .is_some_and(|value| depth == value + 1)
    {
        if namespace != CX {
            return invalid("foreign direct content in ChartEx chartData");
        }
        match local {
            "externalData" => {
                if scan
                    .info
                    .as_ref()
                    .is_some_and(|value| value.external_data.is_some())
                    || !empty
                {
                    return invalid("ChartEx externalData must be a unique leaf");
                }
                let id = required_any(&attributes, &[R, R_STRICT], "id")?.to_owned();
                validate_id(&id)?;
                let auto_update =
                    parse_bool(optional(&attributes, CX, "autoUpdate").unwrap_or("0"))?;
                scan.info.as_mut().expect("root initialized").external_data =
                    Some(ChartExExternalData {
                        relationship_id: id,
                        auto_update,
                    });
            },
            "data" => {
                if empty || scan.current_data.is_some() {
                    return invalid("ChartEx data must be a non-empty direct chartData child");
                }
                let id = required(&attributes, "", "id")?
                    .parse::<u32>()
                    .map_err(|_| invalid_error("invalid ChartEx data ID"))?;
                if !scan.data_ids.insert(id) || scan.data_ids.len() > MAX_DATA_SETS {
                    return invalid("ChartEx data IDs are duplicate or excessive");
                }
                scan.data_depth = Some(depth);
                scan.current_data = Some(ChartExDataSet {
                    id,
                    string_dimensions: 0,
                    numeric_dimensions: 0,
                    dimensions: Vec::new(),
                });
            },
            "extLst" => {},
            _ => return invalid("invalid direct ChartEx chartData child"),
        }
        return Ok(());
    }
    if scan.data_depth.is_some_and(|value| depth == value + 1) && namespace == CX {
        match local {
            "strDim" | "numDim" => {
                if empty || required(&attributes, "", "type")?.len() > 64 {
                    return invalid("ChartEx dimension requires a bounded type and content");
                }
                let data = scan.current_data.as_mut().expect("data depth initialized");
                if local == "strDim" {
                    data.string_dimensions += 1;
                } else {
                    data.numeric_dimensions += 1;
                }
            },
            "extLst" => {},
            _ => return invalid("invalid direct ChartEx data child"),
        }
        return Ok(());
    }
    if scan.chart_depth.is_some_and(|value| depth == value + 1) && namespace == CX {
        match local {
            "title" => scan.info.as_mut().expect("root initialized").has_title = true,
            "legend" => scan.info.as_mut().expect("root initialized").has_legend = true,
            "plotArea" | "extLst" => {},
            _ => return invalid("invalid direct ChartEx chart child"),
        }
    }
    Ok(())
}

fn inspect_end(
    reader: &NsReader<&[u8]>,
    name: quick_xml::name::QName<'_>,
    depth: usize,
    scan: &mut Scan,
) -> Result<()> {
    let namespace = resolved(reader.resolver().resolve_element(name).0)?;
    let local_name = name.local_name();
    let local = std::str::from_utf8(local_name.as_ref()).map_err(xml_error)?;
    if scan.data_depth == Some(depth) && namespace == CX && local == "data" {
        let data = scan.current_data.take().expect("data depth initialized");
        if data.string_dimensions + data.numeric_dimensions == 0 {
            return invalid("ChartEx data requires at least one dimension");
        }
        scan.info
            .as_mut()
            .expect("root initialized")
            .data_sets
            .push(data);
        scan.data_depth = None;
    } else if scan.chart_data_depth == Some(depth) && namespace == CX && local == "chartData" {
        scan.chart_data_depth = None;
    } else if scan.chart_depth == Some(depth) && namespace == CX && local == "chart" {
        scan.chart_depth = None;
    } else if scan.root_depth == Some(depth) && namespace == CX && local == "chartSpace" {
        if !scan.root_children.contains("chartData") || !scan.root_children.contains("chart") {
            return invalid("ChartEx root requires chartData followed by chart");
        }
        scan.root_closed = true;
    }
    if scan.leaf_depth == Some(depth) {
        scan.leaf_depth = None;
    }
    Ok(())
}

fn root_child_rank(namespace: &str, local: &str) -> Option<u8> {
    if namespace != CX {
        return None;
    }
    match local {
        "chartData" => Some(0),
        "chart" => Some(1),
        "spPr" => Some(2),
        "txPr" => Some(3),
        "clrMapOvr" => Some(4),
        "fmtOvrs" => Some(5),
        "printSettings" => Some(6),
        "extLst" => Some(7),
        _ => None,
    }
}

#[derive(Debug)]
struct MiniNode {
    namespace: String,
    name: String,
    attributes: Vec<Attribute>,
    children: Vec<MiniNode>,
    text: String,
}

type ParsedDataGraph = (
    Vec<ChartExDataSet>,
    Vec<ChartExSeriesDataReference>,
    Vec<ChartExAxis>,
    bool,
    ChartExChart,
    ChartExPlotArea,
    ChartExChartSpaceFormatting,
);

fn parse_data_graph(xml: &[u8], version: &str, features: &[String]) -> Result<ParsedDataGraph> {
    let root = parse_mini_tree(xml)?;
    let chart_space_formatting = parse_chart_space_formatting(&root)?;
    let chart_data = one_child(&root, CX, "chartData")?
        .ok_or_else(|| invalid_error("missing ChartEx chartData"))?;
    let chart =
        one_child(&root, CX, "chart")?.ok_or_else(|| invalid_error("missing ChartEx chart"))?;
    let chart_info = parse_chart(chart, offset_feature_allowed(version, features))?;
    let data_sets = parse_chart_data(chart_data)?;
    let plot_area = one_child(chart, CX, "plotArea")?
        .ok_or_else(|| invalid_error("ChartEx chart is missing plotArea"))?;
    reject_unknown(&plot_area.attributes, &[], "plotArea")?;
    if !plot_area.text.trim().is_empty() {
        return invalid("unexpected text in ChartEx plotArea");
    }
    let mut plot_rank = 0u8;
    let mut region = None;
    let mut axes = Vec::new();
    let mut plot_shape_properties = None;
    let mut plot_has_extension_list = false;
    let mut singleton_seen = HashSet::new();
    for child in &plot_area.children {
        if child.namespace != CX {
            return invalid("foreign direct child in ChartEx plotArea");
        }
        let current = match child.name.as_str() {
            "plotAreaRegion" => 0,
            "axis" => 1,
            "spPr" => 2,
            "extLst" => 3,
            _ => return invalid("invalid direct child in ChartEx plotArea"),
        };
        if current < plot_rank {
            return invalid("ChartEx plotArea children are out of order");
        }
        plot_rank = current;
        match child.name.as_str() {
            "plotAreaRegion" if region.is_none() => region = Some(child),
            "axis" => {
                if axes.len() >= MAX_AXES {
                    return limit("ChartEx axis count");
                }
                axes.push(parse_axis(
                    child,
                    offset_feature_allowed(version, features),
                )?);
            },
            "spPr" if singleton_seen.insert(child.name.as_str()) => {
                plot_shape_properties = Some(parse_drawing_payload(child, "plotArea spPr")?);
            },
            "extLst" if singleton_seen.insert(child.name.as_str()) => {
                plot_has_extension_list = true
            },
            _ => return invalid("duplicate ChartEx plotArea child"),
        }
    }
    let region =
        region.ok_or_else(|| invalid_error("ChartEx plotArea is missing plotAreaRegion"))?;
    reject_unknown(&region.attributes, &[], "plotAreaRegion")?;
    if !region.text.trim().is_empty() {
        return invalid("unexpected text in ChartEx plotAreaRegion");
    }
    let mut series = Vec::new();
    let mut region_rank = 0u8;
    let mut plot_surface = None;
    let mut region_ext_seen = false;
    for child in &region.children {
        if child.namespace != CX {
            return invalid("foreign direct child in ChartEx plotAreaRegion");
        }
        let current = match child.name.as_str() {
            "plotSurface" => 0,
            "series" => 1,
            "extLst" => 2,
            _ => return invalid("invalid direct child in ChartEx plotAreaRegion"),
        };
        if current < region_rank {
            return invalid("ChartEx plotAreaRegion children are out of order");
        }
        region_rank = current;
        match child.name.as_str() {
            "plotSurface" if plot_surface.is_none() => {
                plot_surface = Some(parse_plot_surface(child)?);
            },
            "series" => {
                if series.len() >= MAX_SERIES {
                    return limit("ChartEx series count");
                }
                series.push(parse_series(child)?);
            },
            "extLst" if !region_ext_seen => region_ext_seen = true,
            _ => return invalid("duplicate ChartEx plotAreaRegion child"),
        }
    }
    let ids = data_sets
        .iter()
        .map(|value| value.id)
        .collect::<HashSet<_>>();
    let mut axis_ids = HashSet::new();
    for axis in &axes {
        if !axis_ids.insert(axis.id) {
            return invalid("duplicate ChartEx axis ID");
        }
    }
    let mut unique_ids = HashSet::new();
    for (index, value) in series.iter().enumerate() {
        if value.data_id.is_some_and(|id| !ids.contains(&id)) {
            return invalid("ChartEx series dataId does not resolve to chartData");
        }
        if let Some(data_id) = value.data_id {
            let data = data_sets
                .iter()
                .find(|data| data.id == data_id)
                .ok_or_else(|| {
                    invalid_error("ChartEx series dataId does not resolve to chartData")
                })?;
            validate_series_point_references(value, data)?;
        }
        if value
            .owner_index
            .is_some_and(|owner| owner as usize >= series.len() || owner as usize == index)
        {
            return invalid("ChartEx series ownerIdx is out of range or self-referential");
        }
        if let Some(id) = &value.unique_id
            && !unique_ids.insert(id.as_str())
        {
            return invalid("duplicate ChartEx series uniqueId");
        }
        let mut references = HashSet::new();
        for axis_id in &value.axis_ids {
            if !references.insert(*axis_id) {
                return invalid("duplicate ChartEx series axisId");
            }
            if !axis_ids.contains(axis_id) {
                return invalid("ChartEx series axisId does not resolve to plotArea axis");
            }
        }
    }
    let has_plot_surface = plot_surface.is_some();
    let plot_area_info = ChartExPlotArea {
        region: ChartExPlotAreaRegion {
            plot_surface,
            has_extension_list: region_ext_seen,
        },
        shape_properties: plot_shape_properties,
        has_extension_list: plot_has_extension_list,
    };
    Ok((
        data_sets,
        series,
        axes,
        has_plot_surface,
        chart_info,
        plot_area_info,
        chart_space_formatting,
    ))
}

fn parse_chart_space_formatting(root: &MiniNode) -> Result<ChartExChartSpaceFormatting> {
    let shape_properties = one_child(root, CX, "spPr")?
        .map(|node| parse_drawing_payload(node, "chartSpace spPr"))
        .transpose()?;
    let text_properties = one_child(root, CX, "txPr")?
        .map(|node| parse_drawing_payload(node, "chartSpace txPr"))
        .transpose()?;
    let color_mapping_override = one_child(root, CX, "clrMapOvr")?
        .map(|node| parse_drawing_payload(node, "chartSpace clrMapOvr"))
        .transpose()?;
    let format_overrides = one_child(root, CX, "fmtOvrs")?
        .map(parse_format_overrides)
        .transpose()?;
    let print_settings = one_child(root, CX, "printSettings")?
        .map(parse_print_settings)
        .transpose()?;
    Ok(ChartExChartSpaceFormatting {
        shape_properties,
        text_properties,
        color_mapping_override,
        format_overrides,
        print_settings,
        has_extension_list: one_child(root, CX, "extLst")?.is_some(),
    })
}

fn parse_format_overrides(node: &MiniNode) -> Result<Vec<ChartExFormatOverride>> {
    reject_unknown(&node.attributes, &[], "fmtOvrs")?;
    if !node.text.trim().is_empty() {
        return invalid("unexpected text in ChartEx fmtOvrs");
    }
    let mut values = Vec::new();
    let mut indices = HashSet::new();
    for child in &node.children {
        if child.namespace != CX || child.name != "fmtOvr" {
            return invalid("invalid direct child in ChartEx fmtOvrs");
        }
        if values.len() >= MAX_FORMAT_OVERRIDES {
            return limit("ChartEx format override count");
        }
        let value = parse_format_override(child)?;
        if !indices.insert(value.index) {
            return invalid("duplicate ChartEx format override index");
        }
        values.push(value);
    }
    Ok(values)
}

fn parse_format_override(node: &MiniNode) -> Result<ChartExFormatOverride> {
    reject_unknown(&node.attributes, &[("", "idx")], "fmtOvr")?;
    if !node.text.trim().is_empty() {
        return invalid("unexpected text in ChartEx fmtOvr");
    }
    let index = parse_u32(
        required(&node.attributes, "", "idx")?,
        "format override index",
    )?;
    let mut shape_properties = None;
    let mut has_extension_list = false;
    for child in &node.children {
        if child.namespace != CX {
            return invalid("foreign direct child in ChartEx fmtOvr");
        }
        match child.name.as_str() {
            "spPr" if shape_properties.is_none() && !has_extension_list => {
                shape_properties = Some(parse_drawing_payload(child, "fmtOvr spPr")?);
            },
            "extLst" if !has_extension_list => has_extension_list = true,
            _ => {
                return invalid("ChartEx fmtOvr children are invalid, duplicated, or out of order");
            },
        }
    }
    Ok(ChartExFormatOverride {
        index,
        shape_properties,
        has_extension_list,
    })
}

fn parse_print_settings(node: &MiniNode) -> Result<ChartExPrintSettings> {
    reject_unknown(&node.attributes, &[], "printSettings")?;
    if !node.text.trim().is_empty() {
        return invalid("unexpected text in ChartEx printSettings");
    }
    let mut result = ChartExPrintSettings::default();
    let mut rank = 0u8;
    let mut seen = HashSet::new();
    for child in &node.children {
        if child.namespace != CX {
            return invalid("foreign direct child in ChartEx printSettings");
        }
        let current = match child.name.as_str() {
            "headerFooter" => 0,
            "pageMargins" => 1,
            "pageSetup" => 2,
            _ => return invalid("invalid direct child in ChartEx printSettings"),
        };
        if current < rank || !seen.insert(child.name.as_str()) {
            return invalid("ChartEx printSettings children are out of order or duplicated");
        }
        rank = current;
        match child.name.as_str() {
            "headerFooter" => result.header_footer = Some(parse_header_footer(child)?),
            "pageMargins" => result.page_margins = Some(parse_page_margins(child)?),
            "pageSetup" => result.page_setup = Some(parse_page_setup(child)?),
            _ => unreachable!(),
        }
    }
    Ok(result)
}

fn parse_header_footer(node: &MiniNode) -> Result<ChartExHeaderFooter> {
    reject_unknown(
        &node.attributes,
        &[
            ("", "alignWithMargins"),
            ("", "differentOddEven"),
            ("", "differentFirst"),
        ],
        "headerFooter",
    )?;
    if !node.text.trim().is_empty() {
        return invalid("unexpected text in ChartEx headerFooter");
    }
    let mut texts: [Option<String>; 6] = Default::default();
    let mut rank = 0usize;
    for child in &node.children {
        if child.namespace != CX {
            return invalid("foreign direct child in ChartEx headerFooter");
        }
        let current = match child.name.as_str() {
            "oddHeader" => 0,
            "oddFooter" => 1,
            "evenHeader" => 2,
            "evenFooter" => 3,
            "firstHeader" => 4,
            "firstFooter" => 5,
            _ => return invalid("invalid direct child in ChartEx headerFooter"),
        };
        if current < rank || texts[current].is_some() {
            return invalid("ChartEx headerFooter children are out of order or duplicated");
        }
        rank = current;
        texts[current] = Some(parse_print_text(child)?);
    }
    Ok(ChartExHeaderFooter {
        align_with_margins: optional(&node.attributes, "", "alignWithMargins")
            .map(parse_bool)
            .transpose()?
            .unwrap_or(true),
        different_odd_even: optional(&node.attributes, "", "differentOddEven")
            .map(parse_bool)
            .transpose()?
            .unwrap_or(false),
        different_first: optional(&node.attributes, "", "differentFirst")
            .map(parse_bool)
            .transpose()?
            .unwrap_or(false),
        odd_header: texts[0].take(),
        odd_footer: texts[1].take(),
        even_header: texts[2].take(),
        even_footer: texts[3].take(),
        first_header: texts[4].take(),
        first_footer: texts[5].take(),
    })
}

fn parse_print_text(node: &MiniNode) -> Result<String> {
    reject_unknown(&node.attributes, &[], "header/footer text")?;
    if !node.children.is_empty() {
        return invalid("ChartEx header/footer text must have simple content");
    }
    if node.text.len() > MAX_PRINT_TEXT_BYTES {
        return limit("ChartEx header/footer text bytes");
    }
    Ok(node.text.clone())
}

fn parse_page_margins(node: &MiniNode) -> Result<ChartExPageMargins> {
    reject_unknown(
        &node.attributes,
        &[
            ("", "l"),
            ("", "r"),
            ("", "t"),
            ("", "b"),
            ("", "header"),
            ("", "footer"),
        ],
        "pageMargins",
    )?;
    require_empty_content(node, "pageMargins")?;
    let value = |name| -> Result<String> {
        let value = required(&node.attributes, "", name)?;
        if !valid_xml_double(value) {
            return invalid("invalid ChartEx page margin");
        }
        Ok(value.to_owned())
    };
    Ok(ChartExPageMargins {
        left: value("l")?,
        right: value("r")?,
        top: value("t")?,
        bottom: value("b")?,
        header: value("header")?,
        footer: value("footer")?,
    })
}

fn parse_page_setup(node: &MiniNode) -> Result<ChartExPageSetup> {
    reject_unknown(
        &node.attributes,
        &[
            ("", "paperSize"),
            ("", "firstPageNumber"),
            ("", "orientation"),
            ("", "blackAndWhite"),
            ("", "draft"),
            ("", "useFirstPageNumber"),
            ("", "horizontalDpi"),
            ("", "verticalDpi"),
            ("", "copies"),
        ],
        "pageSetup",
    )?;
    require_empty_content(node, "pageSetup")?;
    Ok(ChartExPageSetup {
        paper_size: optional(&node.attributes, "", "paperSize")
            .map(|value| parse_u32(value, "pageSetup paperSize"))
            .transpose()?
            .unwrap_or(1),
        first_page_number: optional(&node.attributes, "", "firstPageNumber")
            .map(|value| parse_u32(value, "pageSetup firstPageNumber"))
            .transpose()?
            .unwrap_or(1),
        orientation: parse_page_orientation(
            optional(&node.attributes, "", "orientation").unwrap_or("default"),
        )?,
        black_and_white: optional(&node.attributes, "", "blackAndWhite")
            .map(parse_bool)
            .transpose()?
            .unwrap_or(false),
        draft: optional(&node.attributes, "", "draft")
            .map(parse_bool)
            .transpose()?
            .unwrap_or(false),
        use_first_page_number: optional(&node.attributes, "", "useFirstPageNumber")
            .map(parse_bool)
            .transpose()?
            .unwrap_or(false),
        horizontal_dpi: optional(&node.attributes, "", "horizontalDpi")
            .map(|value| parse_i32(value, "pageSetup horizontalDpi"))
            .transpose()?
            .unwrap_or(600),
        vertical_dpi: optional(&node.attributes, "", "verticalDpi")
            .map(|value| parse_i32(value, "pageSetup verticalDpi"))
            .transpose()?
            .unwrap_or(600),
        copies: optional(&node.attributes, "", "copies")
            .map(|value| parse_u32(value, "pageSetup copies"))
            .transpose()?
            .unwrap_or(1),
    })
}

fn parse_page_orientation(value: &str) -> Result<ChartExPageOrientation> {
    match value {
        "default" => Ok(ChartExPageOrientation::Default),
        "portrait" => Ok(ChartExPageOrientation::Portrait),
        "landscape" => Ok(ChartExPageOrientation::Landscape),
        _ => invalid("invalid ChartEx page orientation"),
    }
}

fn parse_chart(node: &MiniNode, offset_allowed: bool) -> Result<ChartExChart> {
    reject_unknown(&node.attributes, &[], "chart")?;
    if !node.text.trim().is_empty() {
        return invalid("unexpected text in ChartEx chart");
    }
    let mut title = None;
    let mut legend = None;
    let mut plot_area_seen = false;
    let mut has_extension_list = false;
    let mut rank = 0u8;
    let mut seen = HashSet::new();
    for child in &node.children {
        if child.namespace != CX {
            return invalid("foreign direct child in ChartEx chart");
        }
        let current = match child.name.as_str() {
            "title" => 0,
            "plotArea" => 1,
            "legend" => 2,
            "extLst" => 3,
            _ => return invalid("invalid direct child in ChartEx chart"),
        };
        if current < rank || !seen.insert(child.name.as_str()) {
            return invalid("ChartEx chart children are out of order or duplicated");
        }
        rank = current;
        match child.name.as_str() {
            "title" => title = Some(parse_chart_title(child, offset_allowed)?),
            "plotArea" => plot_area_seen = true,
            "legend" => legend = Some(parse_legend(child, offset_allowed)?),
            "extLst" => has_extension_list = true,
            _ => unreachable!(),
        }
    }
    if !plot_area_seen {
        return invalid("ChartEx chart requires plotArea");
    }
    Ok(ChartExChart {
        title,
        legend,
        has_extension_list,
    })
}

fn parse_chart_title(node: &MiniNode, offset_allowed: bool) -> Result<ChartExChartTitle> {
    reject_unknown(
        &node.attributes,
        &[("", "pos"), ("", "align"), ("", "overlay")],
        "chart title",
    )?;
    if !node.text.trim().is_empty() {
        return invalid("unexpected text in ChartEx chart title");
    }
    let position = parse_side_position(optional(&node.attributes, "", "pos").unwrap_or("t"))?;
    let alignment =
        parse_position_alignment(optional(&node.attributes, "", "align").unwrap_or("ctr"))?;
    let overlay = optional(&node.attributes, "", "overlay")
        .map(parse_bool)
        .transpose()?
        .unwrap_or(false);
    let mut text = None;
    let mut shape_properties = None;
    let mut text_properties = None;
    let mut offset = None;
    let mut rank = 0u8;
    let mut seen = HashSet::new();
    for child in &node.children {
        if child.namespace != CX {
            return invalid("foreign direct child in ChartEx chart title");
        }
        let current = match child.name.as_str() {
            "tx" => 0,
            "spPr" => 1,
            "txPr" => 2,
            "offset" => 3,
            "extLst" => 4,
            _ => return invalid("invalid direct child in ChartEx chart title"),
        };
        if current < rank || !seen.insert(child.name.as_str()) {
            return invalid("ChartEx chart title children are out of order or duplicated");
        }
        rank = current;
        match child.name.as_str() {
            "tx" => text = Some(parse_shared_text(child, "chart title tx")?),
            "spPr" => shape_properties = Some(parse_drawing_payload(child, "chart title spPr")?),
            "txPr" => text_properties = Some(parse_drawing_payload(child, "chart title txPr")?),
            "offset" => {
                if !offset_allowed {
                    return invalid(
                        "ChartEx chart title offset requires version 1.0 or feature mp",
                    );
                }
                offset = Some(parse_offset(child)?);
            },
            "extLst" => {},
            _ => unreachable!(),
        }
    }
    Ok(ChartExChartTitle {
        position,
        alignment,
        overlay,
        text,
        shape_properties,
        text_properties,
        offset,
    })
}

fn parse_legend(node: &MiniNode, offset_allowed: bool) -> Result<ChartExLegend> {
    reject_unknown(
        &node.attributes,
        &[("", "pos"), ("", "align"), ("", "overlay")],
        "legend",
    )?;
    if !node.text.trim().is_empty() {
        return invalid("unexpected text in ChartEx legend");
    }
    let position = parse_side_position(optional(&node.attributes, "", "pos").unwrap_or("r"))?;
    let alignment =
        parse_position_alignment(optional(&node.attributes, "", "align").unwrap_or("ctr"))?;
    let overlay = optional(&node.attributes, "", "overlay")
        .map(parse_bool)
        .transpose()?
        .unwrap_or(false);
    let mut shape_properties = None;
    let mut text_properties = None;
    let mut offset = None;
    let mut rank = 0u8;
    let mut seen = HashSet::new();
    for child in &node.children {
        if child.namespace != CX {
            return invalid("foreign direct child in ChartEx legend");
        }
        let current = match child.name.as_str() {
            "spPr" => 0,
            "txPr" => 1,
            "offset" => 2,
            "extLst" => 3,
            _ => return invalid("invalid direct child in ChartEx legend"),
        };
        if current < rank || !seen.insert(child.name.as_str()) {
            return invalid("ChartEx legend children are out of order or duplicated");
        }
        rank = current;
        match child.name.as_str() {
            "spPr" => shape_properties = Some(parse_drawing_payload(child, "legend spPr")?),
            "txPr" => text_properties = Some(parse_drawing_payload(child, "legend txPr")?),
            "offset" => {
                if !offset_allowed {
                    return invalid("ChartEx legend offset requires version 1.0 or feature mp");
                }
                offset = Some(parse_offset(child)?);
            },
            "extLst" => {},
            _ => unreachable!(),
        }
    }
    Ok(ChartExLegend {
        position,
        alignment,
        overlay,
        shape_properties,
        text_properties,
        offset,
    })
}

fn parse_offset(node: &MiniNode) -> Result<ChartExOffset> {
    reject_unknown(&node.attributes, &[("", "top"), ("", "left")], "offset")?;
    require_empty_content(node, "offset")?;
    let top = required(&node.attributes, "", "top")?;
    let left = required(&node.attributes, "", "left")?;
    if !valid_xml_double(top) || !valid_xml_double(left) {
        return invalid("invalid ChartEx offset coordinate");
    }
    Ok(ChartExOffset {
        top: top.to_owned(),
        left: left.to_owned(),
    })
}

fn parse_side_position(value: &str) -> Result<ChartExSidePosition> {
    match value {
        "l" => Ok(ChartExSidePosition::Left),
        "r" => Ok(ChartExSidePosition::Right),
        "t" => Ok(ChartExSidePosition::Top),
        "b" => Ok(ChartExSidePosition::Bottom),
        _ => invalid("invalid ChartEx side position"),
    }
}

fn parse_position_alignment(value: &str) -> Result<ChartExPositionAlignment> {
    match value {
        "min" => Ok(ChartExPositionAlignment::Minimum),
        "ctr" => Ok(ChartExPositionAlignment::Center),
        "max" => Ok(ChartExPositionAlignment::Maximum),
        _ => invalid("invalid ChartEx position alignment"),
    }
}

fn offset_feature_allowed(version: &str, features: &[String]) -> bool {
    version
        .split('.')
        .next()
        .and_then(|value| value.parse::<u32>().ok())
        .is_some_and(|major| major >= 1)
        || features.iter().any(|feature| feature == "mp")
}

fn validate_series_point_references(
    series: &ChartExSeriesDataReference,
    data: &ChartExDataSet,
) -> Result<()> {
    let bound = data
        .dimensions
        .iter()
        .flat_map(|dimension| match dimension {
            ChartExDimension::String { levels, .. } => levels
                .iter()
                .map(|level| level.point_count)
                .collect::<Vec<_>>(),
            ChartExDimension::Numeric { levels, .. } => levels
                .iter()
                .map(|level| level.point_count)
                .collect::<Vec<_>>(),
        })
        .max();
    let Some(bound) = bound else {
        return Ok(());
    };
    if series.data_points.iter().any(|point| point.index >= bound) {
        return invalid("ChartEx dataPt index does not resolve to cached series data");
    }
    if let Some(labels) = &series.data_labels {
        if labels.labels.iter().any(|label| label.index >= bound)
            || labels.hidden_indices.iter().any(|index| *index >= bound)
        {
            return invalid("ChartEx data label index does not resolve to cached series data");
        }
    }
    Ok(())
}

fn parse_mini_tree(xml: &[u8]) -> Result<MiniNode> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut stack = Vec::<MiniNode>::new();
    let mut root = None;
    let mut nodes = 0usize;
    let mut strings = 0usize;
    loop {
        let event = reader.read_event().map_err(xml_error)?;
        match event {
            Event::Start(element) => {
                nodes += 1;
                if nodes > MAX_NODES || stack.len() >= MAX_DEPTH {
                    return limit("ChartEx semantic XML structure");
                }
                let namespace = resolved(reader.resolver().resolve_element(element.name()).0)?;
                let local_name = element.local_name();
                let name = std::str::from_utf8(local_name.as_ref())
                    .map_err(xml_error)?
                    .to_owned();
                add_strings(&mut strings, namespace.len() + name.len())?;
                let node = MiniNode {
                    namespace,
                    name,
                    attributes: attributes(&reader, &element, &mut strings)?,
                    children: Vec::new(),
                    text: String::new(),
                };
                stack.push(node);
            },
            Event::Empty(element) => {
                nodes += 1;
                if nodes > MAX_NODES || stack.len() >= MAX_DEPTH {
                    return limit("ChartEx semantic XML structure");
                }
                let namespace = resolved(reader.resolver().resolve_element(element.name()).0)?;
                let local_name = element.local_name();
                let name = std::str::from_utf8(local_name.as_ref())
                    .map_err(xml_error)?
                    .to_owned();
                add_strings(&mut strings, namespace.len() + name.len())?;
                let node = MiniNode {
                    namespace,
                    name,
                    attributes: attributes(&reader, &element, &mut strings)?,
                    children: Vec::new(),
                    text: String::new(),
                };
                attach_mini(node, &mut stack, &mut root)?;
            },
            Event::End(_) => {
                let node = stack
                    .pop()
                    .ok_or_else(|| invalid_error("unexpected ChartEx closing element"))?;
                attach_mini(node, &mut stack, &mut root)?;
            },
            Event::Text(value) => {
                let decoded = value.decode().map_err(xml_error)?;
                let decoded = quick_xml::escape::unescape(&decoded).map_err(xml_error)?;
                add_strings(&mut strings, decoded.len())?;
                if let Some(node) = stack.last_mut() {
                    node.text.push_str(&decoded);
                } else if !decoded.trim().is_empty() {
                    return invalid("text outside ChartEx root");
                }
            },
            Event::CData(value) => {
                let decoded = value.decode().map_err(xml_error)?;
                add_strings(&mut strings, decoded.len())?;
                if let Some(node) = stack.last_mut() {
                    node.text.push_str(&decoded);
                }
            },
            Event::GeneralRef(reference) => {
                let name = reference.decode().map_err(xml_error)?;
                let value = match name.as_ref() {
                    "amp" => "&",
                    "lt" => "<",
                    "gt" => ">",
                    "apos" => "'",
                    "quot" => "\"",
                    _ => return invalid("custom entity in ChartEx data is rejected"),
                };
                add_strings(&mut strings, value.len())?;
                if let Some(node) = stack.last_mut() {
                    node.text.push_str(value);
                }
            },
            Event::DocType(_) | Event::PI(_) => {
                return invalid("DTD and processing instructions are rejected in ChartEx data");
            },
            Event::Eof => break,
            _ => {},
        }
    }
    if !stack.is_empty() {
        return invalid("unterminated ChartEx semantic XML");
    }
    root.ok_or_else(|| invalid_error("missing ChartEx semantic root"))
}

fn attach_mini(node: MiniNode, stack: &mut [MiniNode], root: &mut Option<MiniNode>) -> Result<()> {
    if let Some(parent) = stack.last_mut() {
        parent.children.push(node);
    } else if root.replace(node).is_some() {
        return invalid("multiple ChartEx XML roots");
    }
    Ok(())
}

fn parse_chart_data(node: &MiniNode) -> Result<Vec<ChartExDataSet>> {
    let mut rank = 0u8;
    let mut data_sets = Vec::new();
    let mut ids = HashSet::new();
    for child in &node.children {
        if child.namespace != CX {
            return invalid("foreign direct child in ChartEx chartData");
        }
        let current = match child.name.as_str() {
            "externalData" => 0,
            "data" => 1,
            "extLst" => 2,
            _ => return invalid("invalid direct child in ChartEx chartData"),
        };
        if current < rank || (current != 1 && current == rank && current != 0) {
            return invalid("ChartEx chartData children are out of order or duplicated");
        }
        rank = current;
        if child.name == "data" {
            let value = parse_data_set(child)?;
            if !ids.insert(value.id) {
                return invalid("duplicate ChartEx data ID");
            }
            data_sets.push(value);
        }
    }
    if data_sets.is_empty() {
        return invalid("ChartEx chartData requires data");
    }
    Ok(data_sets)
}

fn parse_data_set(node: &MiniNode) -> Result<ChartExDataSet> {
    reject_unknown(&node.attributes, &[("", "id")], "data")?;
    let id = parse_u32(required(&node.attributes, "", "id")?, "data id")?;
    let mut dimensions = Vec::new();
    let mut ext_seen = false;
    for child in &node.children {
        if child.namespace != CX {
            return invalid("foreign direct child in ChartEx data");
        }
        match child.name.as_str() {
            "strDim" if !ext_seen => dimensions.push(parse_string_dimension(child)?),
            "numDim" if !ext_seen => dimensions.push(parse_numeric_dimension(child)?),
            "extLst" if !ext_seen => ext_seen = true,
            _ => return invalid("ChartEx data dimensions are invalid or out of order"),
        }
    }
    if dimensions.is_empty() {
        return invalid("ChartEx data requires at least one dimension");
    }
    let string_dimensions = dimensions
        .iter()
        .filter(|value| matches!(value, ChartExDimension::String { .. }))
        .count();
    let numeric_dimensions = dimensions.len() - string_dimensions;
    Ok(ChartExDataSet {
        id,
        string_dimensions,
        numeric_dimensions,
        dimensions,
    })
}

fn parse_string_dimension(node: &MiniNode) -> Result<ChartExDimension> {
    reject_unknown(&node.attributes, &[("", "type")], "strDim")?;
    let kind = match required(&node.attributes, "", "type")? {
        "cat" => ChartExStringDimensionType::Category,
        "colorStr" => ChartExStringDimensionType::ColorString,
        "entityId" => ChartExStringDimensionType::EntityId,
        _ => return invalid("invalid ChartEx string dimension type"),
    };
    let (formula, name_formula, level_nodes) = dimension_children(node)?;
    let levels = level_nodes
        .into_iter()
        .map(parse_string_level)
        .collect::<Result<Vec<_>>>()?;
    Ok(ChartExDimension::String {
        kind,
        formula,
        name_formula,
        levels,
    })
}

fn parse_numeric_dimension(node: &MiniNode) -> Result<ChartExDimension> {
    reject_unknown(&node.attributes, &[("", "type")], "numDim")?;
    let kind = match required(&node.attributes, "", "type")? {
        "val" => ChartExNumericDimensionType::Value,
        "x" => ChartExNumericDimensionType::X,
        "y" => ChartExNumericDimensionType::Y,
        "size" => ChartExNumericDimensionType::Size,
        "colorVal" => ChartExNumericDimensionType::ColorValue,
        _ => return invalid("invalid ChartEx numeric dimension type"),
    };
    let (formula, name_formula, level_nodes) = dimension_children(node)?;
    let levels = level_nodes
        .into_iter()
        .map(parse_numeric_level)
        .collect::<Result<Vec<_>>>()?;
    Ok(ChartExDimension::Numeric {
        kind,
        formula,
        name_formula,
        levels,
    })
}

fn dimension_children<'a>(
    node: &'a MiniNode,
) -> Result<(
    Option<ChartExFormula>,
    Option<ChartExFormula>,
    Vec<&'a MiniNode>,
)> {
    let mut formula = None;
    let mut name_formula = None;
    let mut levels = Vec::new();
    let mut rank = 0u8;
    for child in &node.children {
        if child.namespace != CX {
            return invalid("foreign direct child in ChartEx dimension");
        }
        let current = match child.name.as_str() {
            "f" => 0,
            "nf" => 1,
            "lvl" => 2,
            _ => return invalid("invalid ChartEx dimension child"),
        };
        if current < rank {
            return invalid("ChartEx dimension children are out of order");
        }
        rank = current;
        match child.name.as_str() {
            "f" if formula.is_none() && levels.is_empty() => formula = Some(parse_formula(child)?),
            "nf" if formula.is_some() && name_formula.is_none() && levels.is_empty() => {
                name_formula = Some(parse_formula(child)?)
            },
            "lvl" => levels.push(child),
            _ => return invalid("ChartEx dimension formula/literal choice is invalid"),
        }
    }
    if formula.is_none() && levels.is_empty() {
        return invalid("ChartEx dimension requires a formula or literal levels");
    }
    if levels.len() > MAX_LEVELS_PER_DIMENSION {
        return limit("ChartEx dimension levels");
    }
    Ok((formula, name_formula, levels))
}

fn parse_formula(node: &MiniNode) -> Result<ChartExFormula> {
    if !node.children.is_empty() {
        return invalid("ChartEx formula must have simple content");
    }
    reject_unknown(&node.attributes, &[("", "dir")], "formula")?;
    if node.text.is_empty() || node.text.len() > MAX_FORMULA_BYTES {
        return invalid("ChartEx formula is empty or excessive");
    }
    let direction = match optional(&node.attributes, "", "dir").unwrap_or("col") {
        "col" => ChartExFormulaDirection::Column,
        "row" => ChartExFormulaDirection::Row,
        _ => return invalid("invalid ChartEx formula direction"),
    };
    Ok(ChartExFormula {
        expression: node.text.clone(),
        direction,
    })
}

fn parse_string_level(node: &MiniNode) -> Result<ChartExStringLevel> {
    reject_unknown(
        &node.attributes,
        &[("", "ptCount"), ("", "name")],
        "string level",
    )?;
    let point_count = level_count(node)?;
    let mut indices = HashSet::new();
    let mut points = Vec::new();
    for point in &node.children {
        if point.namespace != CX || point.name != "pt" || !point.children.is_empty() {
            return invalid("invalid ChartEx string level point");
        }
        reject_unknown(&point.attributes, &[("", "idx")], "string point")?;
        let index = parse_u32(
            required(&point.attributes, "", "idx")?,
            "string point index",
        )?;
        if index >= point_count || !indices.insert(index) {
            return invalid("ChartEx string point index is duplicate or outside ptCount");
        }
        points.push(ChartExStringPoint {
            index,
            value: point.text.clone(),
        });
    }
    Ok(ChartExStringLevel {
        point_count,
        name: bounded_optional(node, "name", 1024)?,
        points,
    })
}

fn parse_numeric_level(node: &MiniNode) -> Result<ChartExNumericLevel> {
    reject_unknown(
        &node.attributes,
        &[("", "ptCount"), ("", "formatCode"), ("", "name")],
        "numeric level",
    )?;
    let point_count = level_count(node)?;
    let mut indices = HashSet::new();
    let mut points = Vec::new();
    for point in &node.children {
        if point.namespace != CX || point.name != "pt" || !point.children.is_empty() {
            return invalid("invalid ChartEx numeric level point");
        }
        reject_unknown(&point.attributes, &[("", "idx")], "numeric point")?;
        let index = parse_u32(
            required(&point.attributes, "", "idx")?,
            "numeric point index",
        )?;
        let value = point.text.trim();
        if index >= point_count || !indices.insert(index) || !valid_xml_double(value) {
            return invalid("invalid ChartEx numeric point");
        }
        points.push(ChartExNumericPoint {
            index,
            value: value.to_owned(),
        });
    }
    Ok(ChartExNumericLevel {
        point_count,
        name: bounded_optional(node, "name", 1024)?,
        format_code: bounded_optional(node, "formatCode", 255)?,
        points,
    })
}

fn level_count(node: &MiniNode) -> Result<u32> {
    let value = parse_u32(required(&node.attributes, "", "ptCount")?, "level ptCount")?;
    if value > MAX_POINTS_PER_LEVEL || node.children.len() > value as usize {
        return limit("ChartEx level point count");
    }
    Ok(value)
}

fn parse_plot_surface(node: &MiniNode) -> Result<ChartExPlotSurface> {
    reject_unknown(&node.attributes, &[], "plotSurface")?;
    if !node.text.trim().is_empty() {
        return invalid("unexpected text in ChartEx plotSurface");
    }
    let mut rank = 0u8;
    let mut seen = HashSet::new();
    let mut shape_properties = None;
    let mut has_extension_list = false;
    for child in &node.children {
        if child.namespace != CX {
            return invalid("foreign direct child in ChartEx plotSurface");
        }
        let current = match child.name.as_str() {
            "spPr" => 0,
            "extLst" => 1,
            _ => return invalid("invalid direct child in ChartEx plotSurface"),
        };
        if current < rank || !seen.insert(child.name.as_str()) {
            return invalid("ChartEx plotSurface children are out of order or duplicated");
        }
        rank = current;
        match child.name.as_str() {
            "spPr" => shape_properties = Some(parse_drawing_payload(child, "plotSurface spPr")?),
            "extLst" => has_extension_list = true,
            _ => unreachable!(),
        }
    }
    Ok(ChartExPlotSurface {
        shape_properties,
        has_extension_list,
    })
}

fn parse_axis(node: &MiniNode, offset_allowed: bool) -> Result<ChartExAxis> {
    reject_unknown(&node.attributes, &[("", "id"), ("", "hidden")], "axis")?;
    if !node.text.trim().is_empty() {
        return invalid("unexpected text in ChartEx axis");
    }
    let id = parse_u32(required(&node.attributes, "", "id")?, "axis id")?;
    let hidden = optional(&node.attributes, "", "hidden")
        .map(parse_bool)
        .transpose()?
        .unwrap_or(false);
    let mut scaling = None;
    let mut title = None;
    let mut units = None;
    let mut major_gridlines = None;
    let mut minor_gridlines = None;
    let mut major_tick_marks = None;
    let mut minor_tick_marks = None;
    let mut tick_labels = None;
    let mut number_format = None;
    let mut shape_properties = None;
    let mut text_properties = None;
    let mut has_extension_list = false;
    let mut rank = 0u8;
    let mut seen = HashSet::new();
    for child in &node.children {
        if child.namespace != CX {
            return invalid("foreign direct child in ChartEx axis");
        }
        let current = match child.name.as_str() {
            "catScaling" | "valScaling" => 0,
            "title" => 1,
            "units" => 2,
            "majorGridlines" => 3,
            "minorGridlines" => 4,
            "majorTickMarks" => 5,
            "minorTickMarks" => 6,
            "tickLabels" => 7,
            "numFmt" => 8,
            "spPr" => 9,
            "txPr" => 10,
            "extLst" => 11,
            _ => return invalid("invalid direct child in ChartEx axis"),
        };
        if current < rank {
            return invalid("ChartEx axis children are out of order");
        }
        rank = current;
        if current == 0 {
            if scaling.is_some() {
                return invalid("ChartEx axis requires exactly one scaling choice");
            }
            scaling = Some(if child.name == "catScaling" {
                parse_category_scaling(child)?
            } else {
                parse_value_scaling(child)?
            });
        } else if !seen.insert(child.name.as_str()) {
            return invalid("duplicate ChartEx axis child");
        } else {
            match child.name.as_str() {
                "title" => title = Some(parse_axis_title(child, offset_allowed)?),
                "units" => units = Some(parse_axis_units(child)?),
                "majorGridlines" => {
                    major_gridlines = Some(parse_gridlines(child, "majorGridlines")?)
                },
                "minorGridlines" => {
                    minor_gridlines = Some(parse_gridlines(child, "minorGridlines")?)
                },
                "majorTickMarks" => {
                    major_tick_marks = Some(parse_tick_marks(child, "majorTickMarks")?)
                },
                "minorTickMarks" => {
                    minor_tick_marks = Some(parse_tick_marks(child, "minorTickMarks")?)
                },
                "tickLabels" => tick_labels = Some(parse_tick_labels(child)?),
                "numFmt" => number_format = Some(parse_number_format(child)?),
                "spPr" => shape_properties = Some(parse_drawing_payload(child, "axis spPr")?),
                "txPr" => text_properties = Some(parse_drawing_payload(child, "axis txPr")?),
                "extLst" => has_extension_list = true,
                _ => unreachable!(),
            }
        }
    }
    Ok(ChartExAxis {
        id,
        hidden,
        scaling: scaling.ok_or_else(|| invalid_error("ChartEx axis is missing scaling"))?,
        title,
        units,
        major_gridlines,
        minor_gridlines,
        major_tick_marks,
        minor_tick_marks,
        tick_labels,
        number_format,
        shape_properties,
        text_properties,
        has_extension_list,
    })
}

fn parse_axis_title(node: &MiniNode, offset_allowed: bool) -> Result<ChartExAxisTitle> {
    reject_unknown(&node.attributes, &[], "axis title")?;
    if !node.text.trim().is_empty() {
        return invalid("unexpected text in ChartEx axis title");
    }
    let mut text = None;
    let mut shape_properties = None;
    let mut text_properties = None;
    let mut offset = None;
    let mut has_extension_list = false;
    let mut rank = 0u8;
    let mut seen = HashSet::new();
    for child in &node.children {
        if child.namespace != CX {
            return invalid("foreign direct child in ChartEx axis title");
        }
        let current = match child.name.as_str() {
            "tx" => 0,
            "spPr" => 1,
            "txPr" => 2,
            "offset" => 3,
            "extLst" => 4,
            _ => return invalid("invalid direct child in ChartEx axis title"),
        };
        if current < rank || !seen.insert(child.name.as_str()) {
            return invalid("ChartEx axis title children are out of order or duplicated");
        }
        rank = current;
        match child.name.as_str() {
            "tx" => text = Some(parse_shared_text(child, "axis title tx")?),
            "spPr" => shape_properties = Some(parse_drawing_payload(child, "axis title spPr")?),
            "txPr" => text_properties = Some(parse_drawing_payload(child, "axis title txPr")?),
            "offset" => {
                if !offset_allowed {
                    return invalid("ChartEx axis title offset requires version 1.0 or feature mp");
                }
                offset = Some(parse_offset(child)?);
            },
            "extLst" => has_extension_list = true,
            _ => unreachable!(),
        }
    }
    Ok(ChartExAxisTitle {
        text,
        shape_properties,
        text_properties,
        offset,
        has_extension_list,
    })
}

fn parse_axis_units(node: &MiniNode) -> Result<ChartExAxisUnits> {
    reject_unknown(&node.attributes, &[("", "unit")], "axis units")?;
    if !node.text.trim().is_empty() {
        return invalid("unexpected text in ChartEx axis units");
    }
    let unit = optional(&node.attributes, "", "unit")
        .map(parse_axis_unit)
        .transpose()?;
    let mut label = None;
    let mut has_extension_list = false;
    let mut rank = 0u8;
    let mut seen = HashSet::new();
    for child in &node.children {
        if child.namespace != CX {
            return invalid("foreign direct child in ChartEx axis units");
        }
        let current = match child.name.as_str() {
            "unitsLabel" => 0,
            "extLst" => 1,
            _ => return invalid("invalid direct child in ChartEx axis units"),
        };
        if current < rank || !seen.insert(child.name.as_str()) {
            return invalid("ChartEx axis units children are out of order or duplicated");
        }
        rank = current;
        match child.name.as_str() {
            "unitsLabel" => label = Some(parse_axis_units_label(child)?),
            "extLst" => has_extension_list = true,
            _ => unreachable!(),
        }
    }
    Ok(ChartExAxisUnits {
        unit,
        label,
        has_extension_list,
    })
}

fn parse_axis_unit(value: &str) -> Result<ChartExAxisUnit> {
    match value {
        "hundreds" => Ok(ChartExAxisUnit::Hundreds),
        "thousands" => Ok(ChartExAxisUnit::Thousands),
        "tenThousands" => Ok(ChartExAxisUnit::TenThousands),
        "hundredThousands" => Ok(ChartExAxisUnit::HundredThousands),
        "millions" => Ok(ChartExAxisUnit::Millions),
        "tenMillions" => Ok(ChartExAxisUnit::TenMillions),
        "hundredMillions" => Ok(ChartExAxisUnit::HundredMillions),
        "billions" => Ok(ChartExAxisUnit::Billions),
        "trillions" => Ok(ChartExAxisUnit::Trillions),
        "percentage" => Ok(ChartExAxisUnit::Percentage),
        _ => invalid("invalid ChartEx axis display unit"),
    }
}

fn parse_axis_units_label(node: &MiniNode) -> Result<ChartExAxisUnitsLabel> {
    reject_unknown(&node.attributes, &[], "axis units label")?;
    if !node.text.trim().is_empty() {
        return invalid("unexpected text in ChartEx axis units label");
    }
    let mut text = None;
    let mut shape_properties = None;
    let mut text_properties = None;
    let mut has_extension_list = false;
    let mut rank = 0u8;
    let mut seen = HashSet::new();
    for child in &node.children {
        if child.namespace != CX {
            return invalid("foreign direct child in ChartEx axis units label");
        }
        let current = match child.name.as_str() {
            "tx" => 0,
            "spPr" => 1,
            "txPr" => 2,
            "extLst" => 3,
            _ => return invalid("invalid direct child in ChartEx axis units label"),
        };
        if current < rank || !seen.insert(child.name.as_str()) {
            return invalid("ChartEx axis units label children are out of order or duplicated");
        }
        rank = current;
        match child.name.as_str() {
            "tx" => text = Some(parse_shared_text(child, "axis units label tx")?),
            "spPr" => {
                shape_properties = Some(parse_drawing_payload(child, "axis units label spPr")?)
            },
            "txPr" => {
                text_properties = Some(parse_drawing_payload(child, "axis units label txPr")?)
            },
            "extLst" => has_extension_list = true,
            _ => unreachable!(),
        }
    }
    Ok(ChartExAxisUnitsLabel {
        text,
        shape_properties,
        text_properties,
        has_extension_list,
    })
}

fn parse_gridlines(node: &MiniNode, label: &str) -> Result<ChartExGridlines> {
    reject_unknown(&node.attributes, &[], label)?;
    if !node.text.trim().is_empty() {
        return invalid(format!("unexpected text in ChartEx {label}"));
    }
    let mut shape_properties = None;
    let mut has_extension_list = false;
    let mut rank = 0u8;
    let mut seen = HashSet::new();
    for child in &node.children {
        if child.namespace != CX {
            return invalid(format!("foreign direct child in ChartEx {label}"));
        }
        let current = match child.name.as_str() {
            "spPr" => 0,
            "extLst" => 1,
            _ => return invalid(format!("invalid direct child in ChartEx {label}")),
        };
        if current < rank || !seen.insert(child.name.as_str()) {
            return invalid(format!(
                "ChartEx {label} children are out of order or duplicated"
            ));
        }
        rank = current;
        match child.name.as_str() {
            "spPr" => shape_properties = Some(parse_drawing_payload(child, label)?),
            "extLst" => has_extension_list = true,
            _ => unreachable!(),
        }
    }
    Ok(ChartExGridlines {
        shape_properties,
        has_extension_list,
    })
}

fn parse_tick_marks(node: &MiniNode, label: &str) -> Result<ChartExTickMarks> {
    reject_unknown(&node.attributes, &[("", "type")], label)?;
    if !node.text.trim().is_empty() {
        return invalid(format!("unexpected text in ChartEx {label}"));
    }
    let kind = optional(&node.attributes, "", "type")
        .map(parse_tick_mark_type)
        .transpose()?;
    let has_extension_list = parse_extension_only(node, label)?;
    Ok(ChartExTickMarks {
        kind,
        has_extension_list,
    })
}

fn parse_tick_mark_type(value: &str) -> Result<ChartExTickMarkType> {
    match value {
        "in" => Ok(ChartExTickMarkType::Inside),
        "out" => Ok(ChartExTickMarkType::Outside),
        "cross" => Ok(ChartExTickMarkType::Cross),
        "none" => Ok(ChartExTickMarkType::None),
        _ => invalid("invalid ChartEx tick mark type"),
    }
}

fn parse_tick_labels(node: &MiniNode) -> Result<ChartExTickLabels> {
    reject_unknown(&node.attributes, &[], "tickLabels")?;
    if !node.text.trim().is_empty() {
        return invalid("unexpected text in ChartEx tickLabels");
    }
    Ok(ChartExTickLabels {
        has_extension_list: parse_extension_only(node, "tickLabels")?,
    })
}

fn parse_extension_only(node: &MiniNode, label: &str) -> Result<bool> {
    let mut has_extension_list = false;
    for child in &node.children {
        if child.namespace != CX || child.name != "extLst" || has_extension_list {
            return invalid(format!(
                "invalid or duplicate direct child in ChartEx {label}"
            ));
        }
        has_extension_list = true;
    }
    Ok(has_extension_list)
}

fn parse_category_scaling(node: &MiniNode) -> Result<ChartExAxisScaling> {
    reject_unknown(&node.attributes, &[("", "gapWidth")], "catScaling")?;
    require_empty_content(node, "catScaling")?;
    let gap_width = optional(&node.attributes, "", "gapWidth")
        .map(|value| parse_nonnegative_or_auto(value, "category gapWidth"))
        .transpose()?;
    Ok(ChartExAxisScaling::Category { gap_width })
}

fn parse_value_scaling(node: &MiniNode) -> Result<ChartExAxisScaling> {
    reject_unknown(
        &node.attributes,
        &[
            ("", "max"),
            ("", "min"),
            ("", "majorUnit"),
            ("", "minorUnit"),
        ],
        "valScaling",
    )?;
    require_empty_content(node, "valScaling")?;
    let maximum = optional(&node.attributes, "", "max")
        .map(|value| parse_double_or_auto(value, "value axis maximum"))
        .transpose()?;
    let minimum = optional(&node.attributes, "", "min")
        .map(|value| parse_double_or_auto(value, "value axis minimum"))
        .transpose()?;
    let major_unit = optional(&node.attributes, "", "majorUnit")
        .map(|value| parse_positive_or_auto(value, "value axis majorUnit"))
        .transpose()?;
    let minor_unit = optional(&node.attributes, "", "minorUnit")
        .map(|value| parse_positive_or_auto(value, "value axis minorUnit"))
        .transpose()?;
    Ok(ChartExAxisScaling::Value {
        minimum,
        maximum,
        major_unit,
        minor_unit,
    })
}

fn parse_layout_properties(node: &MiniNode) -> Result<ChartExLayoutProperties> {
    reject_unknown(&node.attributes, &[], "layoutPr")?;
    if !node.text.trim().is_empty() {
        return invalid("unexpected text in ChartEx layoutPr");
    }
    let mut result = ChartExLayoutProperties::default();
    let mut rank = 0u8;
    let mut seen = HashSet::new();
    let mut aggregation_choice = false;
    for child in &node.children {
        if child.namespace != CX {
            return invalid("foreign direct child in ChartEx layoutPr");
        }
        let current = match child.name.as_str() {
            "parentLabelLayout" => 0,
            "regionLabelLayout" => 1,
            "visibility" => 2,
            "aggregation" | "binning" => 3,
            "geography" => 4,
            "statistics" => 5,
            "subtotals" => 6,
            "extLst" => 7,
            _ => return invalid("invalid direct child in ChartEx layoutPr"),
        };
        if current < rank || !seen.insert(child.name.as_str()) {
            return invalid("ChartEx layoutPr children are out of order or duplicated");
        }
        rank = current;
        match child.name.as_str() {
            "parentLabelLayout" => result.parent_label = Some(parse_parent_label(child)?),
            "regionLabelLayout" => result.region_label = Some(parse_region_label(child)?),
            "visibility" => result.visibility = Some(parse_visibility(child)?),
            "aggregation" => {
                if aggregation_choice {
                    return invalid(
                        "ChartEx layoutPr aggregation and binning are mutually exclusive",
                    );
                }
                require_empty_element(child, "aggregation")?;
                aggregation_choice = true;
                result.aggregation = true;
            },
            "binning" => {
                if aggregation_choice {
                    return invalid(
                        "ChartEx layoutPr aggregation and binning are mutually exclusive",
                    );
                }
                aggregation_choice = true;
                result.binning = Some(parse_binning(child)?);
            },
            "geography" => result.geography = Some(parse_geography(child)?),
            "statistics" => result.quartile_method = parse_statistics(child)?,
            "subtotals" => result.subtotals = parse_subtotals(child)?,
            "extLst" => {},
            _ => unreachable!(),
        }
    }
    Ok(result)
}

fn parse_parent_label(node: &MiniNode) -> Result<ChartExParentLabelLayout> {
    reject_unknown(&node.attributes, &[("", "val")], "parentLabelLayout")?;
    require_empty_content(node, "parentLabelLayout")?;
    match required(&node.attributes, "", "val")? {
        "none" => Ok(ChartExParentLabelLayout::None),
        "banner" => Ok(ChartExParentLabelLayout::Banner),
        "overlapping" => Ok(ChartExParentLabelLayout::Overlapping),
        _ => invalid("invalid ChartEx parentLabelLayout value"),
    }
}

fn parse_region_label(node: &MiniNode) -> Result<ChartExRegionLabelLayout> {
    reject_unknown(&node.attributes, &[("", "val")], "regionLabelLayout")?;
    require_empty_content(node, "regionLabelLayout")?;
    match required(&node.attributes, "", "val")? {
        "none" => Ok(ChartExRegionLabelLayout::None),
        "bestFitOnly" => Ok(ChartExRegionLabelLayout::BestFitOnly),
        "showAll" => Ok(ChartExRegionLabelLayout::ShowAll),
        _ => invalid("invalid ChartEx regionLabelLayout value"),
    }
}

fn parse_visibility(node: &MiniNode) -> Result<ChartExElementVisibility> {
    let allowed = &[
        ("", "connectorLines"),
        ("", "meanLine"),
        ("", "meanMarker"),
        ("", "nonoutliers"),
        ("", "outliers"),
    ];
    reject_unknown(&node.attributes, allowed, "visibility")?;
    require_empty_content(node, "visibility")?;
    Ok(ChartExElementVisibility {
        connector_lines: optional(&node.attributes, "", "connectorLines")
            .map(parse_bool)
            .transpose()?,
        mean_line: optional(&node.attributes, "", "meanLine")
            .map(parse_bool)
            .transpose()?,
        mean_marker: optional(&node.attributes, "", "meanMarker")
            .map(parse_bool)
            .transpose()?,
        nonoutliers: optional(&node.attributes, "", "nonoutliers")
            .map(parse_bool)
            .transpose()?,
        outliers: optional(&node.attributes, "", "outliers")
            .map(parse_bool)
            .transpose()?,
    })
}

fn parse_binning(node: &MiniNode) -> Result<ChartExBinning> {
    reject_unknown(
        &node.attributes,
        &[("", "intervalClosed"), ("", "underflow"), ("", "overflow")],
        "binning",
    )?;
    if !node.text.trim().is_empty() {
        return invalid("unexpected text in ChartEx binning");
    }
    let interval_closed = optional(&node.attributes, "", "intervalClosed")
        .map(|value| match value {
            "l" => Ok(ChartExClosedSide::Left),
            "r" => Ok(ChartExClosedSide::Right),
            _ => invalid("invalid ChartEx binning intervalClosed"),
        })
        .transpose()?;
    let underflow = optional(&node.attributes, "", "underflow")
        .map(|value| parse_double_or_auto(value, "binning underflow"))
        .transpose()?;
    let overflow = optional(&node.attributes, "", "overflow")
        .map(|value| parse_double_or_auto(value, "binning overflow"))
        .transpose()?;
    let mut choice = None;
    for child in &node.children {
        if child.namespace != CX || !matches!(child.name.as_str(), "binSize" | "binCount") {
            return invalid("invalid direct child in ChartEx binning");
        }
        if choice.is_some() || !child.attributes.is_empty() || !child.children.is_empty() {
            return invalid("ChartEx binning permits at most one simple-content choice");
        }
        let value = child.text.trim();
        choice = Some(if child.name == "binSize" {
            if !valid_xml_double(value) {
                return invalid("invalid ChartEx binSize");
            }
            ChartExBinningChoice::Size(value.to_owned())
        } else {
            ChartExBinningChoice::Count(parse_u32(value, "binCount")?)
        });
    }
    Ok(ChartExBinning {
        choice,
        interval_closed,
        underflow,
        overflow,
    })
}

fn parse_geography(node: &MiniNode) -> Result<ChartExGeography> {
    let allowed = &[
        ("", "projectionType"),
        ("", "viewedRegionType"),
        ("", "cultureLanguage"),
        ("", "cultureRegion"),
        ("", "attribution"),
    ];
    reject_unknown(&node.attributes, allowed, "geography")?;
    if !node.text.trim().is_empty() {
        return invalid("unexpected text in ChartEx geography");
    }
    let projection = optional(&node.attributes, "", "projectionType")
        .map(|value| match value {
            "mercator" => Ok(ChartExGeoProjection::Mercator),
            "miller" => Ok(ChartExGeoProjection::Miller),
            "robinson" => Ok(ChartExGeoProjection::Robinson),
            "albers" => Ok(ChartExGeoProjection::Albers),
            _ => invalid("invalid ChartEx geography projectionType"),
        })
        .transpose()?;
    let viewed_region = optional(&node.attributes, "", "viewedRegionType")
        .map(|value| match value {
            "dataOnly" => Ok(ChartExGeoMappingLevel::DataOnly),
            "postalCode" => Ok(ChartExGeoMappingLevel::PostalCode),
            "county" => Ok(ChartExGeoMappingLevel::County),
            "state" => Ok(ChartExGeoMappingLevel::State),
            "countryRegion" => Ok(ChartExGeoMappingLevel::CountryRegion),
            "countryRegionList" => Ok(ChartExGeoMappingLevel::CountryRegionList),
            "world" => Ok(ChartExGeoMappingLevel::World),
            _ => invalid("invalid ChartEx geography viewedRegionType"),
        })
        .transpose()?;
    let culture_language = bounded_required(node, "cultureLanguage", MAX_CULTURE_NAME_LEN)?;
    let culture_region = bounded_required(node, "cultureRegion", MAX_CULTURE_NAME_LEN)?;
    let attribution = bounded_required(node, "attribution", MAX_ATTRIBUTION_LEN)?;
    let mut cache = None;
    for child in &node.children {
        if child.namespace != CX || child.name != "geoCache" || cache.is_some() {
            return invalid("invalid or duplicate ChartEx geography child");
        }
        cache = Some(parse_geo_cache(child)?);
    }
    let has_cache = cache.is_some();
    Ok(ChartExGeography {
        projection,
        viewed_region,
        culture_language,
        culture_region,
        attribution,
        has_cache,
        cache,
    })
}

fn parse_geo_cache(node: &MiniNode) -> Result<ChartExGeoCache> {
    reject_unknown(&node.attributes, &[("", "provider")], "geoCache")?;
    if !node.text.trim().is_empty() {
        return invalid("unexpected text in ChartEx geoCache");
    }
    let provider = geo_required_string(node, "provider", MAX_GEO_STRING_LEN)?;
    if node.children.is_empty() {
        return invalid("ChartEx geoCache requires binary or clear content");
    }
    if node.children.len() > MAX_GEO_CACHE_ENTRIES {
        return limit("ChartEx geography cache entries");
    }
    let mut entries = Vec::with_capacity(node.children.len());
    for child in &node.children {
        if child.namespace != CX {
            return invalid("foreign direct child in ChartEx geoCache");
        }
        entries.push(match child.name.as_str() {
            "binary" => {
                reject_unknown(&child.attributes, &[], "geography binary")?;
                if !child.children.is_empty() {
                    return invalid("ChartEx geography binary contains elements");
                }
                let (encoded_characters, decoded_bytes) = validate_geo_base64(&child.text)?;
                ChartExGeoCacheEntry::Binary {
                    encoded_characters,
                    decoded_bytes,
                }
            },
            "clear" => ChartExGeoCacheEntry::Clear(parse_geo_clear(child)?),
            _ => return invalid("invalid direct child in ChartEx geoCache"),
        });
    }
    Ok(ChartExGeoCache { provider, entries })
}

fn parse_geo_clear(node: &MiniNode) -> Result<ChartExGeoClear> {
    reject_geo_container(node, "geography clear cache")?;
    let mut result = ChartExGeoClear::default();
    let mut rank = 0u8;
    let mut seen = HashSet::new();
    for child in &node.children {
        let current = geo_ordered_child(
            child,
            &[
                "geoLocationQueryResults",
                "geoDataEntityQueryResults",
                "geoDataPointToEntityQueryResults",
                "geoChildEntitiesQueryResults",
                "geoParentEntitiesQueryResults",
            ],
        )?;
        if current < rank || !seen.insert(child.name.as_str()) {
            return invalid(
                "ChartEx clear geography cache children are out of order or duplicated",
            );
        }
        rank = current;
        match child.name.as_str() {
            "geoLocationQueryResults" => {
                result.location_query_results = Some(parse_geo_collection(
                    child,
                    "geoLocationQueryResult",
                    parse_geo_location_query_result,
                )?)
            },
            "geoDataEntityQueryResults" => {
                result.data_entity_query_results = Some(parse_geo_collection(
                    child,
                    "geoDataEntityQueryResult",
                    parse_geo_data_entity_query_result,
                )?)
            },
            "geoDataPointToEntityQueryResults" => {
                result.data_point_to_entity_query_results = Some(parse_geo_collection(
                    child,
                    "geoDataPointToEntityQueryResult",
                    parse_geo_data_point_to_entity_query_result,
                )?)
            },
            "geoChildEntitiesQueryResults" => {
                result.child_entities_query_results = Some(parse_geo_collection(
                    child,
                    "geoChildEntitiesQueryResult",
                    parse_geo_child_entities_query_result,
                )?)
            },
            "geoParentEntitiesQueryResults" => {
                result.parent_entities_query_results = Some(parse_geo_collection(
                    child,
                    "geoParentEntitiesQueryResult",
                    parse_geo_parent_entities_query_result,
                )?)
            },
            _ => unreachable!(),
        }
    }
    Ok(result)
}

fn parse_geo_collection<T>(
    node: &MiniNode,
    item_name: &str,
    parser: fn(&MiniNode) -> Result<T>,
) -> Result<Vec<T>> {
    reject_geo_container(node, &node.name)?;
    if node.children.len() > MAX_GEO_RESULTS {
        return limit("ChartEx geography query results");
    }
    node.children
        .iter()
        .map(|child| {
            if child.namespace != CX || child.name != item_name {
                return invalid(format!("invalid direct child in ChartEx {}", node.name));
            }
            parser(child)
        })
        .collect()
}

fn parse_geo_location_query_result(node: &MiniNode) -> Result<ChartExGeoLocationQueryResult> {
    reject_geo_container(node, "geoLocationQueryResult")?;
    let mut result = ChartExGeoLocationQueryResult::default();
    for child in geo_unique_ordered(node, &["geoLocationQuery", "geoLocations"])? {
        if child.name == "geoLocationQuery" {
            result.query = Some(parse_geo_location_query(child)?);
        } else {
            reject_geo_container(child, "geoLocations")?;
            if child.children.len() > 1 {
                return invalid("geoLocations permits at most one geoLocation");
            }
            result.location = child
                .children
                .first()
                .map(|value| {
                    if value.namespace != CX || value.name != "geoLocation" {
                        return invalid("invalid direct child in geoLocations");
                    }
                    parse_geo_location(value)
                })
                .transpose()?;
        }
    }
    Ok(result)
}

fn parse_geo_location_query(node: &MiniNode) -> Result<ChartExGeoLocationQuery> {
    let allowed = &[
        ("", "countryRegion"),
        ("", "adminDistrict1"),
        ("", "adminDistrict2"),
        ("", "postalCode"),
        ("", "entityType"),
    ];
    reject_unknown(&node.attributes, allowed, "geoLocationQuery")?;
    require_empty_content(node, "geoLocationQuery")?;
    Ok(ChartExGeoLocationQuery {
        country_region: geo_optional_string(node, "countryRegion", MAX_GEO_STRING_LEN)?,
        admin_district1: geo_optional_string(node, "adminDistrict1", MAX_GEO_STRING_LEN)?,
        admin_district2: geo_optional_string(node, "adminDistrict2", MAX_GEO_STRING_LEN)?,
        postal_code: geo_optional_string(node, "postalCode", MAX_GEO_STRING_LEN)?,
        entity_type: parse_geo_entity_type(required(&node.attributes, "", "entityType")?)?,
    })
}

fn parse_geo_location(node: &MiniNode) -> Result<ChartExGeoLocation> {
    reject_unknown(
        &node.attributes,
        &[
            ("", "latitude"),
            ("", "longitude"),
            ("", "entityName"),
            ("", "entityType"),
        ],
        "geoLocation",
    )?;
    if !node.text.trim().is_empty() {
        return invalid("unexpected text in ChartEx geoLocation");
    }
    let address = match node.children.as_slice() {
        [] => None,
        [child] if child.namespace == CX && child.name == "address" => {
            Some(parse_geo_address(child)?)
        },
        _ => return invalid("geoLocation permits at most one ordered address"),
    };
    Ok(ChartExGeoLocation {
        latitude: geo_optional_double(node, "latitude")?,
        longitude: geo_optional_double(node, "longitude")?,
        entity_name: geo_required_string(node, "entityName", MAX_GEO_STRING_LEN)?,
        entity_type: parse_geo_entity_type(required(&node.attributes, "", "entityType")?)?,
        address,
    })
}

fn parse_geo_address(node: &MiniNode) -> Result<ChartExGeoAddress> {
    let allowed = &[
        ("", "address1"),
        ("", "countryRegion"),
        ("", "adminDistrict1"),
        ("", "adminDistrict2"),
        ("", "postalCode"),
        ("", "locality"),
        ("", "isoCountryCode"),
    ];
    reject_unknown(&node.attributes, allowed, "geography address")?;
    require_empty_content(node, "geography address")?;
    Ok(ChartExGeoAddress {
        address1: geo_optional_string(node, "address1", MAX_GEO_STRING_LEN)?,
        country_region: geo_optional_string(node, "countryRegion", MAX_GEO_STRING_LEN)?,
        admin_district1: geo_optional_string(node, "adminDistrict1", MAX_GEO_STRING_LEN)?,
        admin_district2: geo_optional_string(node, "adminDistrict2", MAX_GEO_STRING_LEN)?,
        postal_code: geo_optional_string(node, "postalCode", MAX_GEO_STRING_LEN)?,
        locality: geo_optional_string(node, "locality", MAX_GEO_STRING_LEN)?,
        iso_country_code: geo_optional_string(node, "isoCountryCode", MAX_GEO_STRING_LEN)?,
    })
}

fn parse_geo_data_entity_query_result(node: &MiniNode) -> Result<ChartExGeoDataEntityQueryResult> {
    reject_geo_container(node, "geoDataEntityQueryResult")?;
    let mut result = ChartExGeoDataEntityQueryResult::default();
    for child in geo_unique_ordered(node, &["geoDataEntityQuery", "geoData"])? {
        if child.name == "geoDataEntityQuery" {
            result.query = Some(parse_geo_data_entity_query(child)?);
        } else {
            result.data = Some(parse_geo_data(child)?);
        }
    }
    Ok(result)
}

fn parse_geo_data_entity_query(node: &MiniNode) -> Result<ChartExGeoDataEntityQuery> {
    reject_unknown(
        &node.attributes,
        &[("", "entityType"), ("", "entityId")],
        "geoDataEntityQuery",
    )?;
    require_empty_content(node, "geoDataEntityQuery")?;
    Ok(ChartExGeoDataEntityQuery {
        entity_type: parse_geo_entity_type(required(&node.attributes, "", "entityType")?)?,
        entity_id: geo_required_string(node, "entityId", MAX_GEO_STRING_LEN)?,
    })
}

fn parse_geo_data(node: &MiniNode) -> Result<ChartExGeoData> {
    reject_unknown(
        &node.attributes,
        &[
            ("", "entityName"),
            ("", "entityId"),
            ("", "east"),
            ("", "west"),
            ("", "north"),
            ("", "south"),
        ],
        "geoData",
    )?;
    if !node.text.trim().is_empty() {
        return invalid("unexpected text in ChartEx geoData");
    }
    let mut polygons = None;
    let mut copyrights = None;
    for child in geo_unique_ordered(node, &["geoPolygons", "copyrights"])? {
        if child.name == "geoPolygons" {
            polygons = Some(parse_geo_collection(
                child,
                "geoPolygon",
                parse_geo_polygon,
            )?);
        } else {
            copyrights = Some(parse_geo_copyrights(child)?);
        }
    }
    Ok(ChartExGeoData {
        entity_name: geo_required_string(node, "entityName", MAX_GEO_STRING_LEN)?,
        entity_id: geo_required_string(node, "entityId", MAX_GEO_STRING_LEN)?,
        east: geo_required_double(node, "east")?,
        west: geo_required_double(node, "west")?,
        north: geo_required_double(node, "north")?,
        south: geo_required_double(node, "south")?,
        polygons,
        copyrights,
    })
}

fn parse_geo_polygon(node: &MiniNode) -> Result<ChartExGeoPolygon> {
    reject_unknown(
        &node.attributes,
        &[("", "polygonId"), ("", "numPoints"), ("", "pcaRings")],
        "geoPolygon",
    )?;
    require_empty_content(node, "geoPolygon")?;
    let num_points = geo_required_string(node, "numPoints", 128)?;
    validate_xsd_integer(&num_points, "geoPolygon numPoints")?;
    Ok(ChartExGeoPolygon {
        polygon_id: geo_required_string(node, "polygonId", MAX_GEO_STRING_LEN)?,
        num_points,
        pca_rings: geo_required_string(node, "pcaRings", MAX_GEO_POLYGON_DATA_LEN)?,
    })
}

fn parse_geo_copyrights(node: &MiniNode) -> Result<Vec<String>> {
    reject_geo_container(node, "copyrights")?;
    if node.children.len() > MAX_GEO_RESULTS {
        return limit("ChartEx geography copyrights");
    }
    node.children
        .iter()
        .map(|child| {
            if child.namespace != CX
                || child.name != "copyright"
                || !child.attributes.is_empty()
                || !child.children.is_empty()
            {
                return invalid("invalid direct child in ChartEx copyrights");
            }
            if child.text.len() > MAX_GEO_STRING_LEN {
                return limit("ChartEx geography copyright");
            }
            Ok(child.text.clone())
        })
        .collect()
}

fn parse_geo_data_point_to_entity_query_result(
    node: &MiniNode,
) -> Result<ChartExGeoDataPointToEntityQueryResult> {
    reject_geo_container(node, "geoDataPointToEntityQueryResult")?;
    let mut result = ChartExGeoDataPointToEntityQueryResult::default();
    for child in geo_unique_ordered(node, &["geoDataPointQuery", "geoDataPointToEntityQuery"])? {
        if child.name == "geoDataPointQuery" {
            result.point_query = Some(parse_geo_data_point_query(child)?);
        } else {
            result.entity_query = Some(parse_geo_data_point_to_entity_query(child)?);
        }
    }
    Ok(result)
}

fn parse_geo_data_point_query(node: &MiniNode) -> Result<ChartExGeoDataPointQuery> {
    reject_unknown(
        &node.attributes,
        &[("", "entityType"), ("", "latitude"), ("", "longitude")],
        "geoDataPointQuery",
    )?;
    require_empty_content(node, "geoDataPointQuery")?;
    Ok(ChartExGeoDataPointQuery {
        entity_type: parse_geo_entity_type(required(&node.attributes, "", "entityType")?)?,
        latitude: geo_required_double(node, "latitude")?,
        longitude: geo_required_double(node, "longitude")?,
    })
}

fn parse_geo_data_point_to_entity_query(
    node: &MiniNode,
) -> Result<ChartExGeoDataPointToEntityQuery> {
    reject_unknown(
        &node.attributes,
        &[("", "entityType"), ("", "entityId")],
        "geoDataPointToEntityQuery",
    )?;
    require_empty_content(node, "geoDataPointToEntityQuery")?;
    Ok(ChartExGeoDataPointToEntityQuery {
        entity_type: parse_geo_entity_type(required(&node.attributes, "", "entityType")?)?,
        entity_id: geo_required_string(node, "entityId", MAX_GEO_STRING_LEN)?,
    })
}

fn parse_geo_child_entities_query_result(
    node: &MiniNode,
) -> Result<ChartExGeoChildEntitiesQueryResult> {
    reject_geo_container(node, "geoChildEntitiesQueryResult")?;
    let mut result = ChartExGeoChildEntitiesQueryResult::default();
    for child in geo_unique_ordered(node, &["geoChildEntitiesQuery", "geoChildEntities"])? {
        if child.name == "geoChildEntitiesQuery" {
            result.query = Some(parse_geo_child_entities_query(child)?);
        } else {
            result.children = Some(parse_geo_collection(
                child,
                "geoHierarchyEntity",
                parse_geo_hierarchy_entity,
            )?);
        }
    }
    Ok(result)
}

fn parse_geo_child_entities_query(node: &MiniNode) -> Result<ChartExGeoChildEntitiesQuery> {
    reject_unknown(
        &node.attributes,
        &[("", "entityId")],
        "geoChildEntitiesQuery",
    )?;
    if !node.text.trim().is_empty() {
        return invalid("unexpected text in geoChildEntitiesQuery");
    }
    let child_types = match node.children.as_slice() {
        [] => None,
        [child] if child.namespace == CX && child.name == "geoChildTypes" => {
            reject_geo_container(child, "geoChildTypes")?;
            if child.children.len() > MAX_GEO_RESULTS {
                return limit("ChartEx geography child types");
            }
            Some(
                child
                    .children
                    .iter()
                    .map(|value| {
                        if value.namespace != CX
                            || value.name != "entityType"
                            || !value.attributes.is_empty()
                            || !value.children.is_empty()
                        {
                            return invalid("invalid direct child in geoChildTypes");
                        }
                        parse_geo_entity_type(value.text.trim())
                    })
                    .collect::<Result<Vec<_>>>()?,
            )
        },
        _ => return invalid("geoChildEntitiesQuery permits at most one geoChildTypes"),
    };
    Ok(ChartExGeoChildEntitiesQuery {
        entity_id: geo_required_string(node, "entityId", MAX_GEO_STRING_LEN)?,
        child_types,
    })
}

fn parse_geo_hierarchy_entity(node: &MiniNode) -> Result<ChartExGeoHierarchyEntity> {
    reject_unknown(
        &node.attributes,
        &[("", "entityName"), ("", "entityId"), ("", "entityType")],
        "geoHierarchyEntity",
    )?;
    require_empty_content(node, "geoHierarchyEntity")?;
    Ok(ChartExGeoHierarchyEntity {
        entity_name: geo_required_string(node, "entityName", MAX_GEO_STRING_LEN)?,
        entity_id: geo_required_string(node, "entityId", MAX_GEO_STRING_LEN)?,
        entity_type: parse_geo_entity_type(required(&node.attributes, "", "entityType")?)?,
    })
}

fn parse_geo_parent_entities_query_result(
    node: &MiniNode,
) -> Result<ChartExGeoParentEntitiesQueryResult> {
    reject_geo_container(node, "geoParentEntitiesQueryResult")?;
    let children = &node.children;
    if children.is_empty()
        || children[0].namespace != CX
        || children[0].name != "geoParentEntitiesQuery"
    {
        return invalid("geoParentEntitiesQueryResult requires geoParentEntitiesQuery first");
    }
    reject_unknown(
        &children[0].attributes,
        &[("", "entityId")],
        "geoParentEntitiesQuery",
    )?;
    require_empty_content(&children[0], "geoParentEntitiesQuery")?;
    let entity_id = geo_required_string(&children[0], "entityId", MAX_GEO_STRING_LEN)?;
    let mut entity = None;
    let mut parent_entity_id = None;
    let mut rank = 0u8;
    for child in children.iter().skip(1) {
        let current = geo_ordered_child(child, &["geoEntity", "geoParentEntity"])?;
        if current < rank {
            return invalid("invalid geoParentEntitiesQueryResult order");
        }
        rank = current;
        if child.name == "geoEntity" {
            if entity.is_some() {
                return invalid("duplicate geoEntity");
            }
            reject_unknown(
                &child.attributes,
                &[("", "entityName"), ("", "entityType")],
                "geoEntity",
            )?;
            require_empty_content(child, "geoEntity")?;
            entity = Some(ChartExGeoEntity {
                entity_name: geo_required_string(child, "entityName", MAX_GEO_STRING_LEN)?,
                entity_type: parse_geo_entity_type(required(&child.attributes, "", "entityType")?)?,
            });
        } else {
            if parent_entity_id.is_some() {
                return invalid("duplicate geoParentEntity");
            }
            reject_unknown(&child.attributes, &[("", "entityId")], "geoParentEntity")?;
            require_empty_content(child, "geoParentEntity")?;
            parent_entity_id = Some(geo_required_string(child, "entityId", MAX_GEO_STRING_LEN)?);
        }
    }
    Ok(ChartExGeoParentEntitiesQueryResult {
        entity_id,
        entity,
        parent_entity_id,
    })
}

fn reject_geo_container(node: &MiniNode, label: &str) -> Result<()> {
    reject_unknown(&node.attributes, &[], label)?;
    if !node.text.trim().is_empty() {
        return invalid(format!("unexpected text in ChartEx {label}"));
    }
    Ok(())
}

fn geo_ordered_child(child: &MiniNode, names: &[&str]) -> Result<u8> {
    if child.namespace != CX {
        return invalid("foreign child in ChartEx geography cache");
    }
    names
        .iter()
        .position(|name| *name == child.name)
        .map(|value| value as u8)
        .ok_or_else(|| invalid_error(format!("invalid geography cache child '{}'", child.name)))
}

fn geo_unique_ordered<'a>(node: &'a MiniNode, names: &[&str]) -> Result<Vec<&'a MiniNode>> {
    let mut rank = 0u8;
    let mut seen = HashSet::new();
    for child in &node.children {
        let current = geo_ordered_child(child, names)?;
        if current < rank || !seen.insert(child.name.as_str()) {
            return invalid(format!("invalid {} order or cardinality", node.name));
        }
        rank = current;
    }
    Ok(node.children.iter().collect())
}

fn parse_geo_entity_type(value: &str) -> Result<ChartExGeoEntityType> {
    match value {
        "Address" => Ok(ChartExGeoEntityType::Address),
        "AdminDistrict" => Ok(ChartExGeoEntityType::AdminDistrict),
        "AdminDistrict2" => Ok(ChartExGeoEntityType::AdminDistrict2),
        "AdminDistrict3" => Ok(ChartExGeoEntityType::AdminDistrict3),
        "Continent" => Ok(ChartExGeoEntityType::Continent),
        "CountryRegion" => Ok(ChartExGeoEntityType::CountryRegion),
        "Locality" => Ok(ChartExGeoEntityType::Locality),
        "Ocean" => Ok(ChartExGeoEntityType::Ocean),
        "Planet" => Ok(ChartExGeoEntityType::Planet),
        "PostalCode" => Ok(ChartExGeoEntityType::PostalCode),
        "Region" => Ok(ChartExGeoEntityType::Region),
        "Unsupported" => Ok(ChartExGeoEntityType::Unsupported),
        _ => invalid("invalid ChartEx geography entity type"),
    }
}

fn geo_required_string(node: &MiniNode, name: &str, maximum: usize) -> Result<String> {
    let value = required(&node.attributes, "", name)?;
    if value.len() > maximum {
        return limit("ChartEx geography string");
    }
    Ok(value.to_owned())
}

fn geo_optional_string(node: &MiniNode, name: &str, maximum: usize) -> Result<Option<String>> {
    optional(&node.attributes, "", name)
        .map(|value| {
            if value.len() > maximum {
                return limit("ChartEx geography string");
            }
            Ok(value.to_owned())
        })
        .transpose()
}

fn geo_required_double(node: &MiniNode, name: &str) -> Result<String> {
    let value = required(&node.attributes, "", name)?;
    if !valid_xml_double(value) {
        return invalid(format!("invalid ChartEx geography {name}"));
    }
    Ok(value.to_owned())
}

fn geo_optional_double(node: &MiniNode, name: &str) -> Result<Option<String>> {
    optional(&node.attributes, "", name)
        .map(|value| {
            if !valid_xml_double(value) {
                return invalid(format!("invalid ChartEx geography {name}"));
            }
            Ok(value.to_owned())
        })
        .transpose()
}

fn validate_xsd_integer(value: &str, label: &str) -> Result<()> {
    let digits = value.strip_prefix(['+', '-']).unwrap_or(value);
    if digits.is_empty() || !digits.bytes().all(|byte| byte.is_ascii_digit()) {
        return invalid(format!("invalid ChartEx {label}"));
    }
    Ok(())
}

fn validate_geo_base64(value: &str) -> Result<(usize, usize)> {
    let mut encoded = 0usize;
    let mut padding = 0usize;
    let mut saw_padding = false;
    for byte in value.bytes() {
        if matches!(byte, b' ' | b'\t' | b'\r' | b'\n') {
            continue;
        }
        encoded += 1;
        if byte == b'=' {
            saw_padding = true;
            padding += 1;
            if padding > 2 {
                return invalid("invalid ChartEx geography base64 padding");
            }
        } else if saw_padding || !(byte.is_ascii_alphanumeric() || matches!(byte, b'+' | b'/')) {
            return invalid("invalid ChartEx geography base64 data");
        }
    }
    if encoded % 4 != 0 {
        return invalid("invalid ChartEx geography base64 length");
    }
    let decoded = encoded
        .checked_div(4)
        .and_then(|value| value.checked_mul(3))
        .and_then(|value| value.checked_sub(padding))
        .ok_or_else(|| invalid_error("ChartEx geography base64 size overflow"))?;
    if decoded > MAX_GEO_BINARY_BYTES {
        return limit("ChartEx geography binary data");
    }
    Ok((encoded, decoded))
}

fn parse_statistics(node: &MiniNode) -> Result<Option<ChartExQuartileMethod>> {
    reject_unknown(&node.attributes, &[("", "quartileMethod")], "statistics")?;
    require_empty_content(node, "statistics")?;
    optional(&node.attributes, "", "quartileMethod")
        .map(|value| match value {
            "inclusive" => Ok(ChartExQuartileMethod::Inclusive),
            "exclusive" => Ok(ChartExQuartileMethod::Exclusive),
            _ => invalid("invalid ChartEx statistics quartileMethod"),
        })
        .transpose()
}

fn parse_subtotals(node: &MiniNode) -> Result<Vec<u32>> {
    reject_unknown(&node.attributes, &[], "subtotals")?;
    if !node.text.trim().is_empty() {
        return invalid("unexpected text in ChartEx subtotals");
    }
    let mut values = Vec::new();
    let mut unique = HashSet::new();
    for child in &node.children {
        if values.len() >= MAX_SUBTOTALS {
            return limit("ChartEx subtotal count");
        }
        if child.namespace != CX
            || child.name != "idx"
            || !child.attributes.is_empty()
            || !child.children.is_empty()
        {
            return invalid("invalid ChartEx subtotal index");
        }
        let value = parse_u32(child.text.trim(), "subtotal index")?;
        if !unique.insert(value) {
            return invalid("duplicate ChartEx subtotal index");
        }
        values.push(value);
    }
    Ok(values)
}

fn parse_double_or_auto(value: &str, label: &str) -> Result<ChartExDoubleOrAutomatic> {
    if value == "auto" {
        return Ok(ChartExDoubleOrAutomatic::Automatic);
    }
    if !valid_xml_double(value) {
        return invalid(format!("invalid ChartEx {label}"));
    }
    Ok(ChartExDoubleOrAutomatic::Number(value.to_owned()))
}

fn parse_nonnegative_or_auto(value: &str, label: &str) -> Result<ChartExDoubleOrAutomatic> {
    if value == "auto" {
        return Ok(ChartExDoubleOrAutomatic::Automatic);
    }
    let number = value
        .parse::<f64>()
        .map_err(|_| invalid_error(format!("invalid ChartEx {label}")))?;
    if number.is_nan() || number < 0.0 {
        return invalid(format!("invalid ChartEx {label}"));
    }
    Ok(ChartExDoubleOrAutomatic::Number(value.to_owned()))
}

fn parse_positive_or_auto(value: &str, label: &str) -> Result<ChartExDoubleOrAutomatic> {
    if value == "auto" {
        return Ok(ChartExDoubleOrAutomatic::Automatic);
    }
    let number = value
        .parse::<f64>()
        .map_err(|_| invalid_error(format!("invalid ChartEx {label}")))?;
    if number.is_nan() || number <= 0.0 {
        return invalid(format!("invalid ChartEx {label}"));
    }
    Ok(ChartExDoubleOrAutomatic::Number(value.to_owned()))
}

fn require_empty_element(node: &MiniNode, label: &str) -> Result<()> {
    reject_unknown(&node.attributes, &[], label)?;
    require_empty_content(node, label)
}

fn require_empty_content(node: &MiniNode, label: &str) -> Result<()> {
    if !node.children.is_empty() || !node.text.trim().is_empty() {
        invalid(format!("ChartEx {label} must be empty"))
    } else {
        Ok(())
    }
}

fn bounded_required(node: &MiniNode, name: &str, max: usize) -> Result<String> {
    let value = required(&node.attributes, "", name)?;
    if value.len() > max {
        return limit("ChartEx attribute string");
    }
    Ok(value.to_owned())
}

fn parse_drawing_payload(node: &MiniNode, label: &str) -> Result<ChartExDrawingPayload> {
    if !node.text.trim().is_empty() {
        return invalid(format!("unexpected text in ChartEx {label}"));
    }
    if node
        .children
        .iter()
        .any(|child| !matches!(child.namespace.as_str(), A | A_STRICT))
    {
        return invalid(format!("foreign direct child in ChartEx {label}"));
    }
    Ok(ChartExDrawingPayload {
        child_elements: node.children.len(),
        attributes: node.attributes.len(),
    })
}

fn parse_shared_text(node: &MiniNode, label: &str) -> Result<ChartExText> {
    reject_unknown(&node.attributes, &[], label)?;
    if !node.text.trim().is_empty() || node.children.len() != 1 {
        return invalid(format!("ChartEx {label} requires exactly one text choice"));
    }
    let child = &node.children[0];
    if child.namespace != CX {
        return invalid(format!("foreign ChartEx {label} choice"));
    }
    match child.name.as_str() {
        "txData" => parse_text_data(child),
        "rich" => Ok(ChartExText::Rich(parse_drawing_payload(
            child,
            "rich text",
        )?)),
        _ => invalid(format!("invalid ChartEx {label} choice")),
    }
}

fn parse_text_data(node: &MiniNode) -> Result<ChartExText> {
    reject_unknown(&node.attributes, &[], "txData")?;
    if !node.text.trim().is_empty() || node.children.is_empty() || node.children.len() > 2 {
        return invalid("invalid ChartEx txData choice");
    }
    let mut formula = None;
    let mut value = None;
    for (index, child) in node.children.iter().enumerate() {
        if child.namespace != CX || !matches!(child.name.as_str(), "f" | "v") {
            return invalid("invalid direct child in ChartEx txData");
        }
        match child.name.as_str() {
            "f" if index == 0 && formula.is_none() => formula = Some(parse_formula(child)?),
            "v" if value.is_none() && child.children.is_empty() && child.attributes.is_empty() => {
                if child.text.len() > MAX_LABEL_TEXT_BYTES {
                    return limit("ChartEx text value bytes");
                }
                value = Some(child.text.clone());
            },
            _ => return invalid("ChartEx txData children are out of order or duplicated"),
        }
    }
    if formula.is_none() && value.is_none() {
        return invalid("ChartEx txData is empty");
    }
    Ok(ChartExText::Data { formula, value })
}

fn parse_value_colors(node: &MiniNode) -> Result<ChartExValueColors> {
    reject_unknown(&node.attributes, &[], "valueColors")?;
    if !node.text.trim().is_empty() {
        return invalid("unexpected text in ChartEx valueColors");
    }
    let mut result = ChartExValueColors::default();
    let mut rank = 0u8;
    let mut seen = HashSet::new();
    for child in &node.children {
        if child.namespace != CX {
            return invalid("foreign direct child in ChartEx valueColors");
        }
        let current = match child.name.as_str() {
            "minColor" => 0,
            "midColor" => 1,
            "maxColor" => 2,
            _ => return invalid("invalid direct child in ChartEx valueColors"),
        };
        if current < rank || !seen.insert(child.name.as_str()) {
            return invalid("ChartEx valueColors children are out of order or duplicated");
        }
        rank = current;
        let color = parse_solid_color(child)?;
        match child.name.as_str() {
            "minColor" => result.minimum = Some(color),
            "midColor" => result.middle = Some(color),
            "maxColor" => result.maximum = Some(color),
            _ => unreachable!(),
        }
    }
    Ok(result)
}

fn parse_solid_color(node: &MiniNode) -> Result<ChartExSolidColor> {
    reject_unknown(&node.attributes, &[], "solid color")?;
    if !node.text.trim().is_empty() || node.children.len() != 1 {
        return invalid("ChartEx solid color requires exactly one DrawingML color choice");
    }
    let color = &node.children[0];
    if !matches!(color.namespace.as_str(), A | A_STRICT) {
        return invalid("ChartEx solid color choice has the wrong namespace");
    }
    let kind = match color.name.as_str() {
        "scrgbClr" => ChartExColorKind::ScRgb,
        "srgbClr" => ChartExColorKind::Srgb,
        "hslClr" => ChartExColorKind::Hsl,
        "sysClr" => ChartExColorKind::System,
        "schemeClr" => ChartExColorKind::Scheme,
        "prstClr" => ChartExColorKind::Preset,
        _ => return invalid("invalid ChartEx DrawingML color choice"),
    };
    if !color.text.trim().is_empty()
        || color
            .children
            .iter()
            .any(|child| !matches!(child.namespace.as_str(), A | A_STRICT))
    {
        return invalid("invalid direct payload in ChartEx DrawingML color");
    }
    let value = optional(&color.attributes, "", "val").map(str::to_owned);
    match kind {
        ChartExColorKind::Srgb => {
            let value = value
                .as_deref()
                .ok_or_else(|| invalid_error("missing ChartEx sRGB color value"))?;
            if value.len() != 6 || !value.bytes().all(|byte| byte.is_ascii_hexdigit()) {
                return invalid("invalid ChartEx sRGB color value");
            }
        },
        ChartExColorKind::System | ChartExColorKind::Scheme | ChartExColorKind::Preset => {
            if value.as_deref().is_none_or(|value| {
                value.is_empty()
                    || value.len() > 128
                    || value.bytes().any(|byte| byte.is_ascii_whitespace())
            }) {
                return invalid("invalid ChartEx DrawingML color token");
            }
        },
        ChartExColorKind::ScRgb | ChartExColorKind::Hsl => {},
    }
    Ok(ChartExSolidColor {
        kind,
        value,
        modifier_count: color.children.len(),
    })
}

fn parse_value_color_positions(node: &MiniNode) -> Result<ChartExValueColorPositions> {
    reject_unknown(&node.attributes, &[("", "count")], "valueColorPositions")?;
    if !node.text.trim().is_empty() {
        return invalid("unexpected text in ChartEx valueColorPositions");
    }
    let count = match optional(&node.attributes, "", "count").unwrap_or("2") {
        "2" => 2,
        "3" => 3,
        _ => return invalid("invalid ChartEx valueColorPositions count"),
    };
    let mut minimum = None;
    let mut middle = None;
    let mut maximum = None;
    let mut rank = 0u8;
    let mut seen = HashSet::new();
    for child in &node.children {
        if child.namespace != CX {
            return invalid("foreign direct child in ChartEx valueColorPositions");
        }
        let current = match child.name.as_str() {
            "min" => 0,
            "mid" => 1,
            "max" => 2,
            _ => return invalid("invalid direct child in ChartEx valueColorPositions"),
        };
        if current < rank || !seen.insert(child.name.as_str()) {
            return invalid("ChartEx valueColorPositions children are out of order or duplicated");
        }
        rank = current;
        let value = parse_color_position(child, child.name != "mid")?;
        match child.name.as_str() {
            "min" => minimum = Some(value),
            "mid" => middle = Some(value),
            "max" => maximum = Some(value),
            _ => unreachable!(),
        }
    }
    Ok(ChartExValueColorPositions {
        count,
        minimum,
        middle,
        maximum,
    })
}

fn parse_color_position(node: &MiniNode, allow_extreme: bool) -> Result<ChartExColorPosition> {
    reject_unknown(&node.attributes, &[], "color position")?;
    if !node.text.trim().is_empty() || node.children.len() != 1 {
        return invalid("ChartEx color position requires exactly one choice");
    }
    let child = &node.children[0];
    if child.namespace != CX {
        return invalid("foreign ChartEx color position choice");
    }
    match child.name.as_str() {
        "extreme" if allow_extreme => {
            require_empty_element(child, "extreme color position")?;
            Ok(ChartExColorPosition::Extreme)
        },
        "number" => {
            let value = parse_position_value(child, "number color position")?;
            if !valid_xml_double(&value) {
                return invalid("invalid ChartEx number color position");
            }
            Ok(ChartExColorPosition::Number(value))
        },
        "percent" => {
            let value = parse_position_value(child, "percent color position")?;
            let number = value
                .parse::<f64>()
                .map_err(|_| invalid_error("invalid ChartEx percent color position"))?;
            if !number.is_finite() || !(0.0..=100.0).contains(&number) {
                return invalid("invalid ChartEx percent color position");
            }
            Ok(ChartExColorPosition::Percent(value))
        },
        _ => invalid("invalid ChartEx color position choice"),
    }
}

fn parse_position_value(node: &MiniNode, label: &str) -> Result<String> {
    reject_unknown(&node.attributes, &[("", "val")], label)?;
    require_empty_content(node, label)?;
    Ok(required(&node.attributes, "", "val")?.to_owned())
}

fn parse_data_point(node: &MiniNode) -> Result<ChartExDataPoint> {
    reject_unknown(&node.attributes, &[("", "idx")], "dataPt")?;
    if !node.text.trim().is_empty() {
        return invalid("unexpected text in ChartEx dataPt");
    }
    let index = parse_u32(required(&node.attributes, "", "idx")?, "data point index")?;
    let mut shape_properties = None;
    let mut ext_seen = false;
    for child in &node.children {
        if child.namespace != CX {
            return invalid("foreign direct child in ChartEx dataPt");
        }
        match child.name.as_str() {
            "spPr" if shape_properties.is_none() && !ext_seen => {
                shape_properties = Some(parse_drawing_payload(child, "data point spPr")?)
            },
            "extLst" if !ext_seen => ext_seen = true,
            _ => {
                return invalid("ChartEx dataPt children are invalid, duplicated, or out of order");
            },
        }
    }
    Ok(ChartExDataPoint {
        index,
        shape_properties,
    })
}

fn parse_data_labels(node: &MiniNode) -> Result<ChartExDataLabels> {
    reject_unknown(&node.attributes, &[("", "pos")], "dataLabels")?;
    if !node.text.trim().is_empty() {
        return invalid("unexpected text in ChartEx dataLabels");
    }
    let position = optional(&node.attributes, "", "pos")
        .map(parse_label_position)
        .transpose()?;
    let mut number_format = None;
    let mut shape_properties = None;
    let mut text_properties = None;
    let mut visibility = None;
    let mut separator = None;
    let mut labels = Vec::new();
    let mut hidden_indices = Vec::new();
    let mut rank = 0u8;
    let mut singleton_seen = HashSet::new();
    let mut label_indices = HashSet::new();
    let mut hidden_set = HashSet::new();
    for child in &node.children {
        if child.namespace != CX {
            return invalid("foreign direct child in ChartEx dataLabels");
        }
        let current = match child.name.as_str() {
            "numFmt" => 0,
            "spPr" => 1,
            "txPr" => 2,
            "visibility" => 3,
            "separator" => 4,
            "dataLabel" => 5,
            "dataLabelHidden" => 6,
            "extLst" => 7,
            _ => return invalid("invalid direct child in ChartEx dataLabels"),
        };
        if current < rank {
            return invalid("ChartEx dataLabels children are out of order");
        }
        rank = current;
        if !matches!(child.name.as_str(), "dataLabel" | "dataLabelHidden")
            && !singleton_seen.insert(child.name.as_str())
        {
            return invalid("duplicate ChartEx dataLabels child");
        }
        match child.name.as_str() {
            "numFmt" => number_format = Some(parse_number_format(child)?),
            "spPr" => shape_properties = Some(parse_drawing_payload(child, "dataLabels spPr")?),
            "txPr" => text_properties = Some(parse_drawing_payload(child, "dataLabels txPr")?),
            "visibility" => visibility = Some(parse_label_visibility(child)?),
            "separator" => separator = Some(parse_separator(child)?),
            "dataLabel" => {
                if labels.len() >= MAX_DATA_LABELS {
                    return limit("ChartEx data label count");
                }
                let label = parse_data_label(child)?;
                if !label_indices.insert(label.index) || hidden_set.contains(&label.index) {
                    return invalid("duplicate or conflicting ChartEx data label index");
                }
                labels.push(label);
            },
            "dataLabelHidden" => {
                if hidden_indices.len() >= MAX_DATA_LABELS {
                    return limit("ChartEx hidden data label count");
                }
                let index = parse_hidden_label(child)?;
                if !hidden_set.insert(index) || label_indices.contains(&index) {
                    return invalid("duplicate or conflicting ChartEx data label index");
                }
                hidden_indices.push(index);
            },
            "extLst" => {},
            _ => unreachable!(),
        }
    }
    Ok(ChartExDataLabels {
        position,
        number_format,
        shape_properties,
        text_properties,
        visibility,
        separator,
        labels,
        hidden_indices,
    })
}

fn parse_data_label(node: &MiniNode) -> Result<ChartExDataLabel> {
    reject_unknown(&node.attributes, &[("", "idx"), ("", "pos")], "dataLabel")?;
    if !node.text.trim().is_empty() {
        return invalid("unexpected text in ChartEx dataLabel");
    }
    let index = parse_u32(required(&node.attributes, "", "idx")?, "data label index")?;
    let position = optional(&node.attributes, "", "pos")
        .map(parse_label_position)
        .transpose()?;
    let mut number_format = None;
    let mut shape_properties = None;
    let mut text_properties = None;
    let mut visibility = None;
    let mut separator = None;
    let mut rank = 0u8;
    let mut seen = HashSet::new();
    for child in &node.children {
        if child.namespace != CX {
            return invalid("foreign direct child in ChartEx dataLabel");
        }
        let current = match child.name.as_str() {
            "numFmt" => 0,
            "spPr" => 1,
            "txPr" => 2,
            "visibility" => 3,
            "separator" => 4,
            "extLst" => 5,
            _ => return invalid("invalid direct child in ChartEx dataLabel"),
        };
        if current < rank || !seen.insert(child.name.as_str()) {
            return invalid("ChartEx dataLabel children are out of order or duplicated");
        }
        rank = current;
        match child.name.as_str() {
            "numFmt" => number_format = Some(parse_number_format(child)?),
            "spPr" => shape_properties = Some(parse_drawing_payload(child, "dataLabel spPr")?),
            "txPr" => text_properties = Some(parse_drawing_payload(child, "dataLabel txPr")?),
            "visibility" => visibility = Some(parse_label_visibility(child)?),
            "separator" => separator = Some(parse_separator(child)?),
            "extLst" => {},
            _ => unreachable!(),
        }
    }
    Ok(ChartExDataLabel {
        index,
        position,
        number_format,
        shape_properties,
        text_properties,
        visibility,
        separator,
    })
}

fn parse_hidden_label(node: &MiniNode) -> Result<u32> {
    reject_unknown(&node.attributes, &[("", "idx")], "dataLabelHidden")?;
    require_empty_content(node, "dataLabelHidden")?;
    parse_u32(
        required(&node.attributes, "", "idx")?,
        "hidden data label index",
    )
}

fn parse_number_format(node: &MiniNode) -> Result<ChartExNumberFormat> {
    reject_unknown(
        &node.attributes,
        &[("", "formatCode"), ("", "sourceLinked")],
        "numFmt",
    )?;
    require_empty_content(node, "numFmt")?;
    let format_code = bounded_required(node, "formatCode", 255)?;
    let source_linked = optional(&node.attributes, "", "sourceLinked")
        .map(parse_bool)
        .transpose()?;
    Ok(ChartExNumberFormat {
        format_code,
        source_linked,
    })
}

fn parse_label_visibility(node: &MiniNode) -> Result<ChartExDataLabelVisibility> {
    reject_unknown(
        &node.attributes,
        &[("", "seriesName"), ("", "categoryName"), ("", "value")],
        "data label visibility",
    )?;
    require_empty_content(node, "data label visibility")?;
    Ok(ChartExDataLabelVisibility {
        series_name: optional(&node.attributes, "", "seriesName")
            .map(parse_bool)
            .transpose()?,
        category_name: optional(&node.attributes, "", "categoryName")
            .map(parse_bool)
            .transpose()?,
        value: optional(&node.attributes, "", "value")
            .map(parse_bool)
            .transpose()?,
    })
}

fn parse_separator(node: &MiniNode) -> Result<String> {
    reject_unknown(&node.attributes, &[], "data label separator")?;
    if !node.children.is_empty() {
        return invalid("ChartEx data label separator must have simple content");
    }
    if node.text.len() > MAX_LABEL_TEXT_BYTES {
        return limit("ChartEx data label separator bytes");
    }
    Ok(node.text.clone())
}

fn parse_label_position(value: &str) -> Result<ChartExDataLabelPosition> {
    match value {
        "bestFit" => Ok(ChartExDataLabelPosition::BestFit),
        "b" => Ok(ChartExDataLabelPosition::Bottom),
        "ctr" => Ok(ChartExDataLabelPosition::Center),
        "inBase" => Ok(ChartExDataLabelPosition::InsideBase),
        "inEnd" => Ok(ChartExDataLabelPosition::InsideEnd),
        "l" => Ok(ChartExDataLabelPosition::Left),
        "outEnd" => Ok(ChartExDataLabelPosition::OutsideEnd),
        "r" => Ok(ChartExDataLabelPosition::Right),
        "t" => Ok(ChartExDataLabelPosition::Top),
        _ => invalid("invalid ChartEx data label position"),
    }
}

fn parse_series(node: &MiniNode) -> Result<ChartExSeriesDataReference> {
    reject_unknown(
        &node.attributes,
        &[
            ("", "layoutId"),
            ("", "hidden"),
            ("", "ownerIdx"),
            ("", "uniqueId"),
            ("", "formatIdx"),
        ],
        "series",
    )?;
    let layout = match required(&node.attributes, "", "layoutId")? {
        "boxWhisker" => ChartExSeriesLayout::BoxWhisker,
        "clusteredColumn" => ChartExSeriesLayout::ClusteredColumn,
        "funnel" => ChartExSeriesLayout::Funnel,
        "paretoLine" => ChartExSeriesLayout::ParetoLine,
        "regionMap" => ChartExSeriesLayout::RegionMap,
        "sunburst" => ChartExSeriesLayout::Sunburst,
        "treemap" => ChartExSeriesLayout::Treemap,
        "waterfall" => ChartExSeriesLayout::Waterfall,
        _ => return invalid("invalid ChartEx series layoutId"),
    };
    let mut rank = 0u8;
    let mut text = None;
    let mut shape_properties = None;
    let mut value_colors = None;
    let mut value_color_positions = None;
    let mut data_points = Vec::new();
    let mut data_point_indices = HashSet::new();
    let mut data_labels = None;
    let mut data_id = None;
    let mut layout_properties = None;
    let mut axis_ids = Vec::new();
    let mut singleton_seen = HashSet::new();
    for child in &node.children {
        let current = series_child_rank(child)
            .ok_or_else(|| invalid_error("invalid direct ChartEx series child"))?;
        if current < rank {
            return invalid("ChartEx series children are out of order");
        }
        rank = current;
        if !matches!(child.name.as_str(), "dataPt" | "axisId")
            && !singleton_seen.insert(child.name.as_str())
        {
            return invalid("duplicate ChartEx series child");
        }
        if child.namespace == CX && child.name == "tx" {
            text = Some(parse_shared_text(child, "series tx")?);
        } else if child.namespace == CX && child.name == "spPr" {
            shape_properties = Some(parse_drawing_payload(child, "series spPr")?);
        } else if child.namespace == CX && child.name == "valueColors" {
            value_colors = Some(parse_value_colors(child)?);
        } else if child.namespace == CX && child.name == "valueColorPositions" {
            value_color_positions = Some(parse_value_color_positions(child)?);
        } else if child.namespace == CX && child.name == "dataPt" {
            if data_points.len() >= MAX_SERIES_POINTS {
                return limit("ChartEx series data point count");
            }
            let point = parse_data_point(child)?;
            if !data_point_indices.insert(point.index) {
                return invalid("duplicate ChartEx series data point index");
            }
            data_points.push(point);
        } else if child.namespace == CX && child.name == "dataLabels" {
            data_labels = Some(parse_data_labels(child)?);
        } else if child.namespace == CX && child.name == "dataId" {
            if data_id.is_some() || !child.children.is_empty() || !child.text.trim().is_empty() {
                return invalid("ChartEx series dataId must be a unique leaf");
            }
            reject_unknown(&child.attributes, &[("", "val")], "series dataId")?;
            data_id = Some(parse_u32(
                required(&child.attributes, "", "val")?,
                "series dataId",
            )?);
        } else if child.namespace == CX && child.name == "layoutPr" {
            if layout_properties.is_some() {
                return invalid("duplicate ChartEx series layoutPr");
            }
            layout_properties = Some(parse_layout_properties(child)?);
        } else if child.namespace == CX && child.name == "axisId" {
            if axis_ids.len() >= MAX_AXIS_REFS_PER_SERIES {
                return limit("ChartEx series axis reference count");
            }
            reject_unknown(&child.attributes, &[], "series axisId")?;
            if !child.children.is_empty() {
                return invalid("ChartEx series axisId must have simple content");
            }
            axis_ids.push(parse_u32(child.text.trim(), "series axisId")?);
        }
    }
    let unique_id = bounded_optional(node, "uniqueId", 1024)?;
    Ok(ChartExSeriesDataReference {
        layout,
        text,
        shape_properties,
        value_colors,
        value_color_positions,
        data_points,
        data_labels,
        data_id,
        hidden: optional(&node.attributes, "", "hidden")
            .map(parse_bool)
            .transpose()?
            .unwrap_or(false),
        owner_index: optional(&node.attributes, "", "ownerIdx")
            .map(|value| parse_u32(value, "series ownerIdx"))
            .transpose()?,
        unique_id,
        format_index: optional(&node.attributes, "", "formatIdx")
            .map(|value| parse_u32(value, "series formatIdx"))
            .transpose()?,
        layout_properties,
        axis_ids,
    })
}

fn series_child_rank(node: &MiniNode) -> Option<u8> {
    if node.namespace == CX {
        match node.name.as_str() {
            "tx" => Some(0),
            "spPr" => Some(1),
            "valueColors" => Some(2),
            "valueColorPositions" => Some(3),
            "dataPt" => Some(4),
            "dataLabels" => Some(5),
            "dataId" => Some(6),
            "layoutPr" => Some(7),
            "axisId" => Some(8),
            "extLst" => Some(9),
            _ => None,
        }
    } else {
        None
    }
}

fn one_child<'a>(node: &'a MiniNode, namespace: &str, name: &str) -> Result<Option<&'a MiniNode>> {
    let mut values = node
        .children
        .iter()
        .filter(|value| value.namespace == namespace && value.name == name);
    let value = values.next();
    if values.next().is_some() {
        invalid(format!("duplicate ChartEx {name}"))
    } else {
        Ok(value)
    }
}

fn parse_u32(value: &str, label: &str) -> Result<u32> {
    value
        .parse()
        .map_err(|_| invalid_error(format!("invalid ChartEx {label}")))
}
fn parse_i32(value: &str, label: &str) -> Result<i32> {
    value
        .parse()
        .map_err(|_| invalid_error(format!("invalid ChartEx {label}")))
}
fn bounded_optional(node: &MiniNode, name: &str, max: usize) -> Result<Option<String>> {
    optional(&node.attributes, "", name)
        .map(|value| {
            if value.len() <= max {
                Ok(value.to_owned())
            } else {
                limit("ChartEx attribute string")
            }
        })
        .transpose()
}
fn valid_xml_double(value: &str) -> bool {
    matches!(value, "INF" | "-INF" | "NaN") || (!value.is_empty() && value.parse::<f64>().is_ok())
}

fn attributes(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    strings: &mut usize,
) -> Result<Vec<Attribute>> {
    let mut values = Vec::new();
    for item in element.attributes().with_checks(true) {
        let item = item.map_err(xml_error)?;
        if item.key.as_ref() == b"xmlns" || item.key.as_ref().starts_with(b"xmlns:") {
            continue;
        }
        if values.len() >= MAX_ATTRIBUTES {
            return limit("ChartEx element attributes");
        }
        let (namespace, local) = reader.resolver().resolve_attribute(item.key);
        let namespace = resolved(namespace)?;
        let name = std::str::from_utf8(local.as_ref())
            .map_err(xml_error)?
            .to_owned();
        let value = item
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
            .map_err(xml_error)?
            .into_owned();
        add_strings(strings, namespace.len() + name.len() + value.len())?;
        if values
            .iter()
            .any(|existing: &Attribute| existing.namespace == namespace && existing.name == name)
        {
            return invalid("duplicate expanded ChartEx attribute");
        }
        values.push(Attribute {
            namespace,
            name,
            value,
        });
    }
    Ok(values)
}

fn validate_external_data(
    package: &OpcPackage,
    part: &dyn Part,
    id: &str,
) -> Result<ChartExExternalDataTarget> {
    let relationship = internal_relationship(part, id)?;
    reject_target(relationship.target_ref())?;
    let target = relationship.target_partname().map_err(OoxmlError::Opc)?;
    if !target.as_str().starts_with("/ppt/embeddings/") || target.as_str().ends_with('/') {
        return invalid("ChartEx external data escapes /ppt/embeddings/");
    }
    let target_part = package
        .get_part(&target)
        .map_err(|_| invalid_error("ChartEx external data target is missing"))?;
    if PACKAGE_REL.contains(&relationship.reltype()) {
        if !WORKBOOK_CONTENT_TYPES.contains(&target_part.content_type()) {
            return invalid("ChartEx package relationship targets a non-workbook part");
        }
        Ok(ChartExExternalDataTarget::EmbeddedPackage {
            part_name: target.as_str().to_owned(),
            content_type: target_part.content_type().to_owned(),
        })
    } else if OLE_REL.contains(&relationship.reltype()) {
        if target_part.content_type() != OLE_CONTENT_TYPE {
            return invalid("ChartEx OLE relationship has mismatched content type");
        }
        Ok(ChartExExternalDataTarget::OleObject {
            part_name: target.as_str().to_owned(),
        })
    } else {
        invalid("ChartEx externalData relationship has the wrong type")
    }
}

fn validate_fallback_image(package: &OpcPackage, part: &dyn Part, id: &str) -> Result<String> {
    let relationship = internal_relationship(part, id)?;
    if !IMAGE_REL.contains(&relationship.reltype()) {
        return invalid("ChartEx fallback image relationship has the wrong type");
    }
    reject_target(relationship.target_ref())?;
    let target = relationship.target_partname().map_err(OoxmlError::Opc)?;
    if !target.as_str().starts_with("/ppt/media/") || target.as_str().ends_with('/') {
        return invalid("ChartEx fallback image escapes /ppt/media/");
    }
    let target_part = package
        .get_part(&target)
        .map_err(|_| invalid_error("ChartEx fallback image target is missing"))?;
    if !target_part.content_type().starts_with("image/") {
        return invalid("ChartEx fallback image has a non-image content type");
    }
    Ok(target.as_str().to_owned())
}

fn internal_relationship<'a>(part: &'a dyn Part, id: &str) -> Result<&'a litchi_opc::Relationship> {
    validate_id(id)?;
    let relationship = part
        .rels()
        .get(id)
        .ok_or_else(|| invalid_error(format!("missing ChartEx relationship '{id}'")))?;
    if relationship.is_external() {
        return invalid("external ChartEx relationships are not loaded");
    }
    Ok(relationship)
}

fn reject_target(value: &str) -> Result<()> {
    let lower = value.to_ascii_lowercase();
    if value.is_empty()
        || value.contains(['?', '#', '\\'])
        || lower.contains("%2e")
        || lower.contains("%2f")
        || lower.contains("%5c")
    {
        return invalid("ambiguous or encoded ChartEx relationship target");
    }
    Ok(())
}

fn parse_features(value: &str) -> Result<Vec<String>> {
    let features = value
        .split_ascii_whitespace()
        .map(str::to_owned)
        .collect::<Vec<_>>();
    if features.len() > MAX_FEATURES || features.iter().any(|value| value.len() > 128) {
        return limit("ChartEx feature list");
    }
    Ok(features)
}

fn parse_bool(value: &str) -> Result<bool> {
    match value {
        "0" | "false" => Ok(false),
        "1" | "true" => Ok(true),
        _ => invalid("invalid ChartEx boolean"),
    }
}

fn validate_id(value: &str) -> Result<()> {
    let mut bytes = value.bytes();
    let Some(first) = bytes.next() else {
        return invalid("ChartEx relationship ID is empty");
    };
    if value.len() > 255
        || !(first.is_ascii_alphabetic() || first == b'_')
        || !bytes.all(|byte| byte.is_ascii_alphanumeric() || matches!(byte, b'_' | b'-' | b'.'))
    {
        return invalid("invalid ChartEx relationship ID");
    }
    Ok(())
}

fn optional<'a>(attributes: &'a [Attribute], namespace: &str, name: &str) -> Option<&'a str> {
    attributes
        .iter()
        .find(|value| value.namespace == namespace && value.name == name)
        .map(|value| value.value.as_str())
}
fn required<'a>(attributes: &'a [Attribute], namespace: &str, name: &str) -> Result<&'a str> {
    optional(attributes, namespace, name)
        .filter(|value| !value.is_empty())
        .ok_or_else(|| invalid_error(format!("missing ChartEx attribute '{name}'")))
}
fn required_any<'a>(
    attributes: &'a [Attribute],
    namespaces: &[&str],
    name: &str,
) -> Result<&'a str> {
    let mut values = attributes
        .iter()
        .filter(|value| namespaces.contains(&value.namespace.as_str()) && value.name == name);
    let value = values
        .next()
        .ok_or_else(|| invalid_error(format!("missing ChartEx relationship attribute '{name}'")))?;
    if values.next().is_some() {
        return invalid("duplicate ChartEx relationship attribute aliases");
    }
    Ok(&value.value)
}
fn reject_unknown(attributes: &[Attribute], allowed: &[(&str, &str)], element: &str) -> Result<()> {
    if attributes
        .iter()
        .any(|value| !allowed.contains(&(value.namespace.as_str(), value.name.as_str())))
    {
        return invalid(format!("unexpected attribute on ChartEx {element}"));
    }
    Ok(())
}
fn resolved(value: ResolveResult<'_>) -> Result<String> {
    match value {
        ResolveResult::Bound(Namespace(value)) => Ok(std::str::from_utf8(value.as_ref())
            .map_err(xml_error)?
            .to_owned()),
        ResolveResult::Unbound => Ok(String::new()),
        ResolveResult::Unknown(prefix) => invalid(format!(
            "unbound ChartEx prefix '{}'",
            String::from_utf8_lossy(prefix.as_ref())
        )),
    }
}
fn add_strings(total: &mut usize, amount: usize) -> Result<()> {
    *total = total
        .checked_add(amount)
        .ok_or_else(|| invalid_error("ChartEx string size overflow"))?;
    if *total > MAX_STRING_BYTES {
        limit("ChartEx string bytes")
    } else {
        Ok(())
    }
}
fn xml_error(error: impl std::fmt::Display) -> OoxmlError {
    OoxmlError::InvalidFormat(format!("invalid ChartEx XML: {error}"))
}
fn invalid_error(message: impl Into<String>) -> OoxmlError {
    OoxmlError::InvalidFormat(message.into())
}
fn invalid<T>(message: impl Into<String>) -> Result<T> {
    Err(invalid_error(message))
}
fn limit<T>(name: &str) -> Result<T> {
    invalid(format!("{name} exceeds resource limit"))
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_opc::PackURI;
    use litchi_opc::part::BlobPart;

    fn producer_xml(auto_update: &str) -> String {
        format!(
            r#"<?xml version="1.0"?><cx:chartSpace xmlns:cx="{CX}" xmlns:a="{A}" xmlns:r="{R}" version="1.0" featureList="sunburst" fallbackImg="rIdImage"><cx:chartData><cx:externalData r:id="rIdWorkbook" cx:autoUpdate="{auto_update}"/><cx:data id="0"><cx:strDim type="cat"><cx:f dir="row">Sheet1!$A$2:$A$3</cx:f><cx:lvl ptCount="2" name="Category"><cx:pt idx="0">A</cx:pt><cx:pt idx="1">B</cx:pt></cx:lvl></cx:strDim><cx:numDim type="size"><cx:lvl ptCount="2" formatCode="General"><cx:pt idx="0">3</cx:pt><cx:pt idx="1">4</cx:pt></cx:lvl></cx:numDim></cx:data></cx:chartData><cx:chart><cx:title pos="t" align="ctr" overlay="1"><cx:tx><cx:txData><cx:v>Quarterly</cx:v></cx:txData></cx:tx><cx:spPr><a:noFill/></cx:spPr><cx:txPr><a:bodyPr/><a:lstStyle/><a:p/></cx:txPr><cx:offset top="0.1" left="-0.2"/></cx:title><cx:plotArea><cx:plotAreaRegion><cx:plotSurface><cx:spPr><a:solidFill><a:schemeClr val="accent2"/></a:solidFill></cx:spPr><cx:extLst/></cx:plotSurface><cx:series layoutId="sunburst" uniqueId="series-0"><cx:tx><cx:txData><cx:f>Sheet1!$B$1</cx:f><cx:v>Revenue</cx:v></cx:txData></cx:tx><cx:spPr><a:solidFill><a:srgbClr val="4472C4"/></a:solidFill></cx:spPr><cx:valueColors><cx:minColor><a:srgbClr val="0000FF"/></cx:minColor><cx:midColor><a:schemeClr val="accent1"/></cx:midColor><cx:maxColor><a:srgbClr val="FF0000"/></cx:maxColor></cx:valueColors><cx:valueColorPositions count="3"><cx:min><cx:extreme/></cx:min><cx:mid><cx:percent val="50"/></cx:mid><cx:max><cx:number val="10"/></cx:max></cx:valueColorPositions><cx:dataPt idx="1"><cx:spPr><a:ln/></cx:spPr></cx:dataPt><cx:dataLabels pos="bestFit"><cx:numFmt formatCode="0.0" sourceLinked="0"/><cx:spPr/><cx:txPr><a:bodyPr/><a:lstStyle/><a:p/></cx:txPr><cx:visibility seriesName="1" categoryName="true" value="0"/><cx:separator> | </cx:separator><cx:dataLabel idx="0" pos="t"><cx:spPr/><cx:visibility value="1"/><cx:separator>:</cx:separator></cx:dataLabel><cx:dataLabelHidden idx="1"/></cx:dataLabels><cx:dataId val="0"/><cx:layoutPr><cx:parentLabelLayout val="banner"/><cx:regionLabelLayout val="bestFitOnly"/><cx:visibility connectorLines="1" meanLine="false" outliers="true"/><cx:binning intervalClosed="l" underflow="auto" overflow="10"><cx:binCount>4</cx:binCount></cx:binning><cx:geography projectionType="mercator" viewedRegionType="countryRegion" cultureLanguage="en-US" cultureRegion="US" attribution="Map data"><cx:geoCache provider="Microsoft"><cx:binary>AQID</cx:binary><cx:clear><cx:geoLocationQueryResults><cx:geoLocationQueryResult><cx:geoLocationQuery countryRegion="US" entityType="CountryRegion"/><cx:geoLocations><cx:geoLocation latitude="47.6" longitude="-122.3" entityName="United States" entityType="CountryRegion"><cx:address countryRegion="United States" isoCountryCode="US"/></cx:geoLocation></cx:geoLocations></cx:geoLocationQueryResult></cx:geoLocationQueryResults><cx:geoDataEntityQueryResults><cx:geoDataEntityQueryResult><cx:geoDataEntityQuery entityType="CountryRegion" entityId="US"/><cx:geoData entityName="United States" entityId="US" east="-66" west="-125" north="49" south="24"><cx:geoPolygons><cx:geoPolygon polygonId="p1" numPoints="4" pcaRings="0,0 1,0 1,1 0,0"/></cx:geoPolygons><cx:copyrights><cx:copyright>Map provider</cx:copyright></cx:copyrights></cx:geoData></cx:geoDataEntityQueryResult></cx:geoDataEntityQueryResults><cx:geoDataPointToEntityQueryResults><cx:geoDataPointToEntityQueryResult><cx:geoDataPointQuery entityType="CountryRegion" latitude="47.6" longitude="-122.3"/><cx:geoDataPointToEntityQuery entityType="CountryRegion" entityId="US"/></cx:geoDataPointToEntityQueryResult></cx:geoDataPointToEntityQueryResults><cx:geoChildEntitiesQueryResults><cx:geoChildEntitiesQueryResult><cx:geoChildEntitiesQuery entityId="US"><cx:geoChildTypes><cx:entityType>AdminDistrict</cx:entityType></cx:geoChildTypes></cx:geoChildEntitiesQuery><cx:geoChildEntities><cx:geoHierarchyEntity entityName="Washington" entityId="WA" entityType="AdminDistrict"/></cx:geoChildEntities></cx:geoChildEntitiesQueryResult></cx:geoChildEntitiesQueryResults><cx:geoParentEntitiesQueryResults><cx:geoParentEntitiesQueryResult><cx:geoParentEntitiesQuery entityId="WA"/><cx:geoEntity entityName="Washington" entityType="AdminDistrict"/><cx:geoParentEntity entityId="US"/></cx:geoParentEntitiesQueryResult></cx:geoParentEntitiesQueryResults></cx:clear></cx:geoCache></cx:geography><cx:statistics quartileMethod="inclusive"/><cx:subtotals><cx:idx>1</cx:idx><cx:idx>3</cx:idx></cx:subtotals></cx:layoutPr><cx:axisId>7</cx:axisId><cx:axisId>8</cx:axisId></cx:series><cx:extLst/></cx:plotAreaRegion><cx:axis id="7"><cx:catScaling gapWidth="0.5"/></cx:axis><cx:axis id="8" hidden="0"><cx:valScaling min="auto" max="10" majorUnit="2" minorUnit="auto"/><cx:title><cx:tx><cx:txData><cx:v>Value Axis</cx:v></cx:txData></cx:tx><cx:spPr><a:noFill/></cx:spPr><cx:txPr><a:bodyPr/><a:lstStyle/><a:p/></cx:txPr><cx:offset top="0.25" left="0.5"/><cx:extLst/></cx:title><cx:units unit="millions"><cx:unitsLabel><cx:tx><cx:txData><cx:v>Millions</cx:v></cx:txData></cx:tx><cx:spPr><a:noFill/></cx:spPr><cx:txPr><a:bodyPr/><a:lstStyle/><a:p/></cx:txPr><cx:extLst/></cx:unitsLabel><cx:extLst/></cx:units><cx:majorGridlines><cx:spPr><a:ln/></cx:spPr><cx:extLst/></cx:majorGridlines><cx:minorGridlines/><cx:majorTickMarks type="in"><cx:extLst/></cx:majorTickMarks><cx:minorTickMarks type="none"/><cx:tickLabels><cx:extLst/></cx:tickLabels><cx:numFmt formatCode="0.00" sourceLinked="true"/><cx:spPr><a:ln w="12700"/></cx:spPr><cx:txPr><a:bodyPr/><a:lstStyle/><a:p/></cx:txPr><cx:extLst/></cx:axis><cx:spPr><a:noFill/></cx:spPr><cx:extLst/></cx:plotArea><cx:legend pos="r" align="max" overlay="false"><cx:spPr><a:noFill/></cx:spPr><cx:txPr><a:bodyPr/><a:lstStyle/><a:p/></cx:txPr><cx:offset top="INF" left="0"/></cx:legend><cx:extLst/></cx:chart><cx:spPr><a:noFill/></cx:spPr><cx:txPr><a:bodyPr/><a:lstStyle/><a:p/></cx:txPr><cx:clrMapOvr bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/><cx:fmtOvrs><cx:fmtOvr idx="2"><cx:spPr><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill></cx:spPr><cx:extLst/></cx:fmtOvr></cx:fmtOvrs><cx:printSettings><cx:headerFooter alignWithMargins="0" differentOddEven="1" differentFirst="true"><cx:oddHeader>Quarterly Report</cx:oddHeader><cx:oddFooter>Confidential</cx:oddFooter><cx:evenHeader>Even Page</cx:evenHeader><cx:evenFooter>2</cx:evenFooter><cx:firstHeader>Cover</cx:firstHeader><cx:firstFooter>First</cx:firstFooter></cx:headerFooter><cx:pageMargins l="0.7" r="0.7" t="0.75" b="0.75" header="0.3" footer="0.3"/><cx:pageSetup paperSize="9" firstPageNumber="2" orientation="landscape" blackAndWhite="1" draft="false" useFirstPageNumber="true" horizontalDpi="600" verticalDpi="300" copies="2"/></cx:printSettings><cx:extLst/></cx:chartSpace>"#
        )
    }

    fn chart_part(xml: String) -> BlobPart {
        BlobPart::new(
            PackURI::new("/ppt/charts/chartEx1.xml").unwrap(),
            CHART_EX_CONTENT_TYPE.into(),
            xml.into_bytes(),
        )
    }

    #[test]
    fn parses_producer_core_and_preserves_round_trip() {
        let part = chart_part(producer_xml("0"));
        let document = ChartExPart::from_part(&part).unwrap().parse().unwrap();
        assert_eq!(document.info().version, "1.0");
        assert_eq!(document.info().features, ["sunburst"]);
        let data = &document.info().data_sets[0];
        assert_eq!(
            (data.id, data.string_dimensions, data.numeric_dimensions),
            (0, 1, 1)
        );
        assert!(
            matches!(&data.dimensions[0], ChartExDimension::String { kind: ChartExStringDimensionType::Category, formula: Some(ChartExFormula { direction: ChartExFormulaDirection::Row, .. }), levels, .. } if levels[0].points[0].value == "A")
        );
        assert!(
            matches!(&data.dimensions[1], ChartExDimension::Numeric { kind: ChartExNumericDimensionType::Size, levels, .. } if levels[0].points[0].value == "3")
        );
        let series = &document.info().series[0];
        assert_eq!(series.layout, ChartExSeriesLayout::Sunburst);
        assert!(
            matches!(&series.text, Some(ChartExText::Data { formula: Some(ChartExFormula { direction: ChartExFormulaDirection::Column, .. }), value: Some(value) }) if value == "Revenue")
        );
        assert_eq!(series.shape_properties.as_ref().unwrap().child_elements, 1);
        assert_eq!(
            series
                .value_colors
                .as_ref()
                .unwrap()
                .middle
                .as_ref()
                .unwrap()
                .kind,
            ChartExColorKind::Scheme
        );
        assert!(
            matches!(series.value_color_positions.as_ref().unwrap().middle, Some(ChartExColorPosition::Percent(ref value)) if value == "50")
        );
        assert_eq!(series.data_points[0].index, 1);
        let labels = series.data_labels.as_ref().unwrap();
        assert_eq!(labels.position, Some(ChartExDataLabelPosition::BestFit));
        assert_eq!(labels.number_format.as_ref().unwrap().format_code, "0.0");
        assert_eq!(
            labels.visibility.as_ref().unwrap().category_name,
            Some(true)
        );
        assert_eq!(labels.labels[0].index, 0);
        assert_eq!(labels.hidden_indices, [1]);
        assert_eq!(series.data_id, Some(0));
        assert_eq!(series.axis_ids, [7, 8]);
        let layout = series.layout_properties.as_ref().unwrap();
        assert_eq!(layout.parent_label, Some(ChartExParentLabelLayout::Banner));
        assert_eq!(
            layout.region_label,
            Some(ChartExRegionLabelLayout::BestFitOnly)
        );
        assert_eq!(
            layout.visibility.as_ref().unwrap().connector_lines,
            Some(true)
        );
        assert!(matches!(
            layout.binning.as_ref().unwrap().choice,
            Some(ChartExBinningChoice::Count(4))
        ));
        assert!(matches!(
            layout.geography.as_ref().unwrap(),
            ChartExGeography {
                projection: Some(ChartExGeoProjection::Mercator),
                viewed_region: Some(ChartExGeoMappingLevel::CountryRegion),
                has_cache: true,
                ..
            }
        ));
        let cache = layout.geography.as_ref().unwrap().cache.as_ref().unwrap();
        assert_eq!(cache.provider, "Microsoft");
        assert!(matches!(
            cache.entries[0],
            ChartExGeoCacheEntry::Binary {
                encoded_characters: 4,
                decoded_bytes: 3
            }
        ));
        let ChartExGeoCacheEntry::Clear(clear) = &cache.entries[1] else {
            panic!("expected clear geography cache")
        };
        assert_eq!(
            clear.location_query_results.as_ref().unwrap()[0]
                .location
                .as_ref()
                .unwrap()
                .entity_name,
            "United States"
        );
        assert_eq!(
            clear.data_entity_query_results.as_ref().unwrap()[0]
                .data
                .as_ref()
                .unwrap()
                .polygons
                .as_ref()
                .unwrap()[0]
                .num_points,
            "4"
        );
        assert_eq!(
            clear.data_point_to_entity_query_results.as_ref().unwrap()[0]
                .entity_query
                .as_ref()
                .unwrap()
                .entity_id,
            "US"
        );
        assert_eq!(
            clear.child_entities_query_results.as_ref().unwrap()[0]
                .children
                .as_ref()
                .unwrap()[0]
                .entity_id,
            "WA"
        );
        assert_eq!(
            clear.parent_entities_query_results.as_ref().unwrap()[0]
                .parent_entity_id
                .as_deref(),
            Some("US")
        );
        assert_eq!(
            layout.quartile_method,
            Some(ChartExQuartileMethod::Inclusive)
        );
        assert_eq!(layout.subtotals, [1, 3]);
        assert!(document.info().has_plot_surface);
        assert_eq!(document.info().axes.len(), 2);
        assert!(matches!(
            document.info().axes[0].scaling,
            ChartExAxisScaling::Category { .. }
        ));
        assert!(matches!(
            document.info().axes[1].scaling,
            ChartExAxisScaling::Value { .. }
        ));
        let value_axis = &document.info().axes[1];
        assert!(
            matches!(&value_axis.title.as_ref().unwrap().text, Some(ChartExText::Data { value: Some(value), .. }) if value == "Value Axis")
        );
        assert_eq!(
            value_axis
                .title
                .as_ref()
                .unwrap()
                .offset
                .as_ref()
                .unwrap()
                .left,
            "0.5"
        );
        assert!(value_axis.title.as_ref().unwrap().has_extension_list);
        let units = value_axis.units.as_ref().unwrap();
        assert_eq!(units.unit, Some(ChartExAxisUnit::Millions));
        assert!(
            matches!(&units.label.as_ref().unwrap().text, Some(ChartExText::Data { value: Some(value), .. }) if value == "Millions")
        );
        assert!(units.has_extension_list && units.label.as_ref().unwrap().has_extension_list);
        assert_eq!(
            value_axis
                .major_gridlines
                .as_ref()
                .unwrap()
                .shape_properties
                .as_ref()
                .unwrap()
                .child_elements,
            1
        );
        assert!(value_axis.minor_gridlines.is_some());
        assert_eq!(
            value_axis.major_tick_marks.as_ref().unwrap().kind,
            Some(ChartExTickMarkType::Inside)
        );
        assert_eq!(
            value_axis.minor_tick_marks.as_ref().unwrap().kind,
            Some(ChartExTickMarkType::None)
        );
        assert!(value_axis.tick_labels.as_ref().unwrap().has_extension_list);
        assert_eq!(
            value_axis.number_format.as_ref().unwrap(),
            &ChartExNumberFormat {
                format_code: "0.00".into(),
                source_linked: Some(true)
            }
        );
        assert!(
            value_axis.shape_properties.is_some()
                && value_axis.text_properties.is_some()
                && value_axis.has_extension_list
        );
        let plot_area = &document.info().plot_area;
        assert!(plot_area.shape_properties.is_some() && plot_area.has_extension_list);
        let surface = plot_area.region.plot_surface.as_ref().unwrap();
        assert_eq!(surface.shape_properties.as_ref().unwrap().child_elements, 1);
        assert!(surface.has_extension_list && plot_area.region.has_extension_list);
        let chart_space = &document.info().chart_space_formatting;
        assert!(chart_space.shape_properties.is_some() && chart_space.text_properties.is_some());
        assert_eq!(
            chart_space
                .color_mapping_override
                .as_ref()
                .unwrap()
                .attributes,
            12
        );
        let overrides = chart_space.format_overrides.as_ref().unwrap();
        assert_eq!(overrides.len(), 1);
        assert_eq!(overrides[0].index, 2);
        assert!(overrides[0].shape_properties.is_some() && overrides[0].has_extension_list);
        let print = chart_space.print_settings.as_ref().unwrap();
        let header_footer = print.header_footer.as_ref().unwrap();
        assert_eq!(
            (
                header_footer.align_with_margins,
                header_footer.different_odd_even,
                header_footer.different_first
            ),
            (false, true, true)
        );
        assert_eq!(
            header_footer.odd_header.as_deref(),
            Some("Quarterly Report")
        );
        assert_eq!(print.page_margins.as_ref().unwrap().left, "0.7");
        let page_setup = print.page_setup.as_ref().unwrap();
        assert_eq!(
            (
                page_setup.paper_size,
                page_setup.first_page_number,
                page_setup.orientation
            ),
            (9, 2, ChartExPageOrientation::Landscape)
        );
        assert_eq!(
            (
                page_setup.black_and_white,
                page_setup.draft,
                page_setup.use_first_page_number
            ),
            (true, false, true)
        );
        assert_eq!(
            (
                page_setup.horizontal_dpi,
                page_setup.vertical_dpi,
                page_setup.copies
            ),
            (600, 300, 2)
        );
        assert!(chart_space.has_extension_list);
        let title = document.info().chart.title.as_ref().unwrap();
        assert_eq!(
            (title.position, title.alignment, title.overlay),
            (
                ChartExSidePosition::Top,
                ChartExPositionAlignment::Center,
                true
            )
        );
        assert!(
            matches!(&title.text, Some(ChartExText::Data { formula: None, value: Some(value) }) if value == "Quarterly")
        );
        assert_eq!(
            title.offset.as_ref().unwrap(),
            &ChartExOffset {
                top: "0.1".into(),
                left: "-0.2".into()
            }
        );
        let legend = document.info().chart.legend.as_ref().unwrap();
        assert_eq!(
            (legend.position, legend.alignment, legend.overlay),
            (
                ChartExSidePosition::Right,
                ChartExPositionAlignment::Maximum,
                false
            )
        );
        assert_eq!(legend.offset.as_ref().unwrap().top, "INF");
        assert!(document.info().chart.has_extension_list);
        assert!(document.info().has_title && document.info().has_legend);
        assert_eq!(document.to_xml(), part.blob());
        let reparsed = chart_part(String::from_utf8(document.to_xml()).unwrap());
        assert_eq!(
            ChartExPart::from_part(&reparsed)
                .unwrap()
                .parse()
                .unwrap()
                .info(),
            document.info()
        );
    }

    #[test]
    fn validates_inert_package_relationships_without_opening_targets() {
        let mut chart = chart_part(producer_xml("0"));
        chart.rels_mut().add_relationship(
            PACKAGE_REL[0].into(),
            "../embeddings/data.xlsx".into(),
            "rIdWorkbook".into(),
            false,
        );
        chart.rels_mut().add_relationship(
            IMAGE_REL[1].into(),
            "../media/fallback.png".into(),
            "rIdImage".into(),
            false,
        );
        let mut package = OpcPackage::new();
        package.add_part(Box::new(BlobPart::new(
            PackURI::new("/ppt/embeddings/data.xlsx").unwrap(),
            WORKBOOK_CONTENT_TYPES[0].into(),
            b"not opened".to_vec(),
        )));
        package.add_part(Box::new(BlobPart::new(
            PackURI::new("/ppt/media/fallback.png").unwrap(),
            "image/png".into(),
            b"not opened".to_vec(),
        )));
        let document = ChartExPart::from_part(&chart)
            .unwrap()
            .parse_in_package(&package)
            .unwrap();
        assert!(
            matches!(document.external_data_target(), Some(ChartExExternalDataTarget::EmbeddedPackage { part_name, .. }) if part_name == "/ppt/embeddings/data.xlsx")
        );
        assert_eq!(
            document.fallback_image_part_name(),
            Some("/ppt/media/fallback.png")
        );
    }

    #[test]
    fn rejects_hostile_schema_and_resource_cases() {
        let cases = [
            producer_xml("0").replace(CX, "urn:vendor:chartex"),
            producer_xml("0").replace(
                "<cx:data id=\"0\">",
                "<cx:data id=\"0\"></cx:data><cx:data id=\"0\">",
            ),
            producer_xml("0").replace("<cx:chartData>", "<cx:chart/><cx:chartData>"),
            producer_xml("0").replace("<cx:strDim type=\"cat\">", "<cx:strDim>"),
            format!(
                "<!DOCTYPE x [<!ENTITY e SYSTEM 'file:///etc/passwd'>]>{}",
                producer_xml("0")
            ),
        ];
        for xml in cases {
            assert!(
                ChartExPart::from_part(&chart_part(xml))
                    .unwrap()
                    .parse()
                    .is_err()
            );
        }
        let oversized = chart_part(" ".repeat(MAX_XML_BYTES + 1));
        assert!(ChartExPart::from_part(&oversized).unwrap().parse().is_err());
    }

    #[test]
    fn rejects_invalid_dimension_choices_points_and_series_references() {
        let base = producer_xml("0");
        let cases = [
            base.replace("type=\"cat\"", "type=\"vendor\""),
            base.replace("type=\"size\"", "type=\"z\""),
            base.replace("<cx:f dir=\"row\">", "<cx:nf dir=\"row\">")
                .replace("</cx:f>", "</cx:nf>"),
            base.replace("</cx:strDim>", "<cx:f>late</cx:f></cx:strDim>"),
            base.replace(
                "ptCount=\"1\" name=\"Category\"",
                "ptCount=\"1\" name=\"Category\"",
            )
            .replace(
                "</cx:lvl></cx:strDim>",
                "<cx:pt idx=\"0\">B</cx:pt></cx:lvl></cx:strDim>",
            ),
            base.replace("<cx:pt idx=\"0\">3</cx:pt>", "<cx:pt idx=\"9\">3</cx:pt>"),
            base.replace(">3</cx:pt>", ">not-a-number</cx:pt>"),
            base.replace("<cx:dataId val=\"0\"/>", "<cx:dataId val=\"99\"/>"),
            base.replace(
                "<cx:dataId val=\"0\"/>",
                "<cx:dataId val=\"0\"/><cx:dataId val=\"0\"/>",
            ),
            base.replace(
                "<cx:dataId val=\"0\"/>",
                "<cx:layoutPr/><cx:dataId val=\"0\"/>",
            ),
        ];
        for xml in cases {
            assert!(
                ChartExPart::from_part(&chart_part(xml))
                    .unwrap()
                    .parse()
                    .is_err()
            );
        }
    }

    #[test]
    fn rejects_invalid_layout_axes_scaling_and_plot_surface_grammar() {
        let base = producer_xml("0");
        let cases = [
            base.replace("<cx:binning intervalClosed=\"l\"", "<cx:aggregation/><cx:binning intervalClosed=\"l\""),
            base.replace("<cx:binCount>4</cx:binCount>", "<cx:binSize>1</cx:binSize><cx:binCount>4</cx:binCount>"),
            base.replace("intervalClosed=\"l\"", "intervalClosed=\"center\""),
            base.replace("quartileMethod=\"inclusive\"", "quartileMethod=\"median\""),
            base.replace("<cx:idx>3</cx:idx>", "<cx:idx>1</cx:idx>"),
            base.replace(" cultureRegion=\"US\"", ""),
            base.replace("<cx:axis id=\"8\" hidden=\"0\">", "<cx:axis id=\"7\" hidden=\"0\">"),
            base.replace("<cx:axisId>8</cx:axisId>", "<cx:axisId>99</cx:axisId>"),
            base.replace("<cx:axisId>8</cx:axisId>", "<cx:axisId>7</cx:axisId>"),
            base.replace("majorUnit=\"2\"", "majorUnit=\"0\""),
            base.replace("gapWidth=\"0.5\"", "gapWidth=\"-1\""),
            base.replace("<cx:catScaling gapWidth=\"0.5\"/>", "<cx:catScaling gapWidth=\"0.5\"/><cx:valScaling/>"),
            base.replace(r#"<cx:plotSurface><cx:spPr><a:solidFill><a:schemeClr val="accent2"/></a:solidFill></cx:spPr><cx:extLst/></cx:plotSurface>"#, "").replace("<cx:series layoutId", "<cx:series layoutId").replace("</cx:series><cx:extLst/></cx:plotAreaRegion>", r#"</cx:series><cx:plotSurface><cx:spPr><a:solidFill><a:schemeClr val="accent2"/></a:solidFill></cx:spPr><cx:extLst/></cx:plotSurface></cx:plotAreaRegion>"#),
            base.replace(r#"<cx:plotSurface><cx:spPr><a:solidFill><a:schemeClr val="accent2"/></a:solidFill></cx:spPr><cx:extLst/></cx:plotSurface>"#, r#"<cx:plotSurface><cx:spPr><a:solidFill><a:schemeClr val="accent2"/></a:solidFill></cx:spPr><cx:extLst/></cx:plotSurface><cx:plotSurface><cx:spPr><a:solidFill><a:schemeClr val="accent2"/></a:solidFill></cx:spPr><cx:extLst/></cx:plotSurface>"#),
            base.replace("<cx:parentLabelLayout val=\"banner\"/>", "").replace("<cx:regionLabelLayout val=\"bestFitOnly\"/>", "<cx:regionLabelLayout val=\"bestFitOnly\"/><cx:parentLabelLayout val=\"banner\"/>"),
        ];
        for xml in cases {
            assert!(
                ChartExPart::from_part(&chart_part(xml))
                    .unwrap()
                    .parse()
                    .is_err()
            );
        }
    }

    #[test]
    fn rejects_invalid_axis_title_units_gridlines_ticks_and_formatting() {
        let base = producer_xml("0");
        let cases = [
            base.replace(
                "<cx:title><cx:tx><cx:txData><cx:v>Value Axis",
                "<cx:title vendor=\"1\"><cx:tx><cx:txData><cx:v>Value Axis",
            ),
            base.replace(
                "<cx:title><cx:tx><cx:txData><cx:v>Value Axis",
                "<cx:units/><cx:title><cx:tx><cx:txData><cx:v>Value Axis",
            ),
            base.replace(
                "</cx:title><cx:units unit=\"millions\">",
                "</cx:title><cx:title/><cx:units unit=\"millions\">",
            ),
            base.replace(
                "<cx:v>Value Axis</cx:v></cx:txData></cx:tx><cx:spPr>",
                "<cx:v>Value Axis</cx:v></cx:txData></cx:tx><cx:tx/><cx:spPr>",
            ),
            base.replace("unit=\"millions\"", "unit=\"vendor\""),
            base.replace(
                "</cx:unitsLabel><cx:extLst/></cx:units>",
                "</cx:unitsLabel><cx:unitsLabel/><cx:extLst/></cx:units>",
            ),
            base.replace(
                "<cx:units unit=\"millions\"><cx:unitsLabel>",
                "<cx:units unit=\"millions\"><cx:extLst/><cx:unitsLabel>",
            ),
            base.replace(
                "<cx:v>Millions</cx:v></cx:txData></cx:tx><cx:spPr>",
                "<cx:v>Millions</cx:v></cx:txData></cx:tx><cx:tx/><cx:spPr>",
            ),
            base.replace("<cx:majorGridlines>", "<cx:majorGridlines vendor=\"1\">"),
            base.replace(
                "<cx:majorGridlines><cx:spPr><a:ln/></cx:spPr>",
                "<cx:majorGridlines><cx:spPr/><cx:spPr><a:ln/></cx:spPr>",
            ),
            base.replace(
                "<cx:majorGridlines>",
                "<cx:minorGridlines/><cx:majorGridlines>",
            ),
            base.replace("type=\"in\"", "type=\"vendor\""),
            base.replace(
                "<cx:majorTickMarks type=\"in\"><cx:extLst/>",
                "<cx:majorTickMarks type=\"in\"><a:extLst/>",
            ),
            base.replace("<cx:tickLabels>", "<cx:tickLabels vendor=\"1\">"),
            base.replace(
                "<cx:tickLabels><cx:extLst/></cx:tickLabels>",
                "<cx:tickLabels><cx:vendor/></cx:tickLabels>",
            ),
            base.replace(
                "<cx:numFmt formatCode=\"0.00\" sourceLinked=\"true\"/>",
                "<cx:numFmt sourceLinked=\"true\"/>",
            ),
            base.replace(
                "<cx:spPr><a:ln w=\"12700\"/></cx:spPr><cx:txPr>",
                "<cx:spPr><cx:ln w=\"12700\"/></cx:spPr><cx:txPr>",
            ),
            base.replace(
                "</cx:txPr><cx:extLst/></cx:axis>",
                "</cx:txPr><cx:txPr/><cx:extLst/></cx:axis>",
            ),
            base.replace(
                "<cx:extLst/></cx:axis>",
                "<cx:extLst/><cx:extLst/></cx:axis>",
            ),
            base.replace(
                "<cx:offset top=\"0.25\" left=\"0.5\"/>",
                "<cx:offset top=\"0.25\"/>",
            ),
            base.replace(" version=\"1.0\"", " version=\"0.0\"")
                .replace("<cx:offset top=\"0.1\" left=\"-0.2\"/>", "")
                .replace("<cx:offset top=\"INF\" left=\"0\"/>", ""),
        ];
        for xml in cases {
            assert!(
                ChartExPart::from_part(&chart_part(xml))
                    .unwrap()
                    .parse()
                    .is_err()
            );
        }
    }

    #[test]
    fn rejects_hostile_geography_cache_grammar_and_bounds() {
        let base = producer_xml("0");
        let cases = [
            base.replace("<cx:geoCache provider=\"Microsoft\">", "<cx:geoCache>"),
            base.replace("<cx:binary>AQID</cx:binary><cx:clear>", ""),
            base.replace("<cx:binary>AQID</cx:binary>", "<cx:binary>AQI!</cx:binary>"),
            base.replace("<cx:geoLocationQueryResults>", "<cx:geoDataEntityQueryResults/><cx:geoLocationQueryResults>"),
            base.replace("<cx:geoDataEntityQueryResults>", "<cx:geoLocationQueryResults/><cx:geoDataEntityQueryResults>"),
            base.replacen("entityType=\"CountryRegion\"", "entityType=\"countryRegion\"", 1),
            base.replace("<cx:geoParentEntitiesQuery entityId=\"WA\"/>", ""),
            base.replace("<cx:geoLocations><cx:geoLocation", "<cx:geoLocations><cx:geoLocation entityName=\"duplicate\" entityType=\"Region\"/><cx:geoLocation"),
            base.replace("<cx:geoCache provider=\"Microsoft\">", "<cx:geoCache xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" provider=\"Microsoft\" r:id=\"rIdMap\">"),
            base.replace("<cx:geoPolygon polygonId=\"p1\" numPoints=\"4\"", "<cx:geoPolygon polygonId=\"p1\" numPoints=\"four\""),
        ];
        for xml in cases {
            assert!(
                ChartExPart::from_part(&chart_part(xml))
                    .unwrap()
                    .parse()
                    .is_err()
            );
        }
        let oversized = "A".repeat(((MAX_GEO_BINARY_BYTES + 1 + 2) / 3) * 4);
        assert!(
            ChartExPart::from_part(&chart_part(base.replace("AQID", &oversized)))
                .unwrap()
                .parse()
                .is_err()
        );
    }

    #[test]
    fn applies_chart_print_defaults() {
        let xml = producer_xml("0")
            .replace(" alignWithMargins=\"0\" differentOddEven=\"1\" differentFirst=\"true\"", "")
            .replace("<cx:pageSetup paperSize=\"9\" firstPageNumber=\"2\" orientation=\"landscape\" blackAndWhite=\"1\" draft=\"false\" useFirstPageNumber=\"true\" horizontalDpi=\"600\" verticalDpi=\"300\" copies=\"2\"/>", "<cx:pageSetup/>");
        let document = ChartExPart::from_part(&chart_part(xml))
            .unwrap()
            .parse()
            .unwrap();
        let print = document
            .info()
            .chart_space_formatting
            .print_settings
            .as_ref()
            .unwrap();
        let header_footer = print.header_footer.as_ref().unwrap();
        assert_eq!(
            (
                header_footer.align_with_margins,
                header_footer.different_odd_even,
                header_footer.different_first
            ),
            (true, false, false)
        );
        let setup = print.page_setup.as_ref().unwrap();
        assert_eq!(
            (setup.paper_size, setup.first_page_number, setup.orientation),
            (1, 1, ChartExPageOrientation::Default)
        );
        assert_eq!(
            (
                setup.black_and_white,
                setup.draft,
                setup.use_first_page_number
            ),
            (false, false, false)
        );
        assert_eq!(
            (setup.horizontal_dpi, setup.vertical_dpi, setup.copies),
            (600, 600, 1)
        );
    }

    #[test]
    fn rejects_invalid_plot_and_chart_space_formatting_and_print_settings() {
        let base = producer_xml("0");
        let cases = [
            base.replace("<cx:plotSurface>", "<cx:plotSurface vendor=\"1\">"),
            base.replace("<cx:plotSurface><cx:spPr><a:solidFill><a:schemeClr val=\"accent2\"/></a:solidFill>", "<cx:plotSurface><cx:spPr><cx:solidFill><a:schemeClr val=\"accent2\"/></cx:solidFill>"),
            base.replace("<cx:plotSurface><cx:spPr>", "<cx:plotSurface><cx:extLst/><cx:spPr>"),
            base.replace("</cx:axis><cx:spPr><a:noFill/></cx:spPr>", "</cx:axis><cx:spPr/><cx:spPr><a:noFill/></cx:spPr>"),
            base.replace("</cx:series><cx:extLst/></cx:plotAreaRegion>", "</cx:series><cx:extLst/><cx:series layoutId=\"funnel\"/></cx:plotAreaRegion>"),
            base.replace("</cx:chart><cx:spPr><a:noFill/></cx:spPr><cx:txPr>", "</cx:chart><cx:txPr/><cx:spPr><a:noFill/></cx:spPr><cx:txPr>"),
            base.replace("</cx:chart><cx:spPr><a:noFill/></cx:spPr>", "</cx:chart><cx:spPr><cx:noFill/></cx:spPr>"),
            base.replace("folHlink=\"folHlink\"/>", "folHlink=\"folHlink\"><cx:vendor/></cx:clrMapOvr>"),
            base.replace("<cx:fmtOvrs>", "<cx:fmtOvrs vendor=\"1\">"),
            base.replace("<cx:fmtOvr idx=\"2\">", "<cx:fmtOvr>"),
            base.replace("</cx:fmtOvr></cx:fmtOvrs>", "</cx:fmtOvr><cx:fmtOvr idx=\"2\"/></cx:fmtOvrs>"),
            base.replace("<cx:fmtOvr idx=\"2\"><cx:spPr><a:solidFill><a:srgbClr val=\"00FF00\"/></a:solidFill>", "<cx:fmtOvr idx=\"2\"><cx:spPr><cx:solidFill><a:srgbClr val=\"00FF00\"/></cx:solidFill>"),
            base.replace("<cx:fmtOvr idx=\"2\"><cx:spPr>", "<cx:fmtOvr idx=\"2\"><cx:extLst/><cx:spPr>"),
            base.replace("<cx:printSettings>", "<cx:printSettings vendor=\"1\">"),
            base.replace("</cx:headerFooter><cx:pageMargins", "</cx:headerFooter><cx:pageSetup/><cx:pageMargins"),
            base.replace("</cx:headerFooter><cx:pageMargins", "</cx:headerFooter><cx:headerFooter/><cx:pageMargins"),
            base.replace("differentFirst=\"true\"", "differentFirst=\"yes\""),
            base.replace("<cx:oddHeader>Quarterly Report</cx:oddHeader><cx:oddFooter>", "<cx:oddFooter/><cx:oddHeader>Quarterly Report</cx:oddHeader><cx:oddFooter>"),
            base.replace("<cx:oddHeader>Quarterly Report</cx:oddHeader>", "<cx:oddHeader><cx:v>Quarterly Report</cx:v></cx:oddHeader>"),
            base.replace(" l=\"0.7\"", ""),
            base.replace("l=\"0.7\"", "l=\"invalid\""),
            base.replace("orientation=\"landscape\"", "orientation=\"sideways\""),
            base.replace("blackAndWhite=\"1\"", "blackAndWhite=\"yes\""),
            base.replace("paperSize=\"9\"", "paperSize=\"-1\""),
            base.replace("horizontalDpi=\"600\"", "horizontalDpi=\"999999999999\""),
            base.replace("copies=\"2\"/>", "copies=\"2\"><cx:vendor/></cx:pageSetup>"),
        ];
        for (index, xml) in cases.into_iter().enumerate() {
            assert!(
                ChartExPart::from_part(&chart_part(xml))
                    .unwrap()
                    .parse()
                    .is_err(),
                "hostile chart-space case {index}"
            );
        }
        let oversized = base.replace("Quarterly Report", &"x".repeat(MAX_PRINT_TEXT_BYTES + 1));
        assert!(
            ChartExPart::from_part(&chart_part(oversized))
                .unwrap()
                .parse()
                .is_err()
        );

        let many = (0..=MAX_FORMAT_OVERRIDES)
            .map(|index| format!("<cx:fmtOvr idx=\"{index}\"/>"))
            .collect::<String>();
        let excessive = base.replace("<cx:fmtOvr idx=\"2\"><cx:spPr><a:solidFill><a:srgbClr val=\"00FF00\"/></a:solidFill></cx:spPr><cx:extLst/></cx:fmtOvr>", &many);
        assert!(
            ChartExPart::from_part(&chart_part(excessive))
                .unwrap()
                .parse()
                .is_err()
        );
    }

    #[test]
    fn parses_strict_rich_series_text_as_bounded_inert_drawingml() {
        let xml = producer_xml("0").replace(A, A_STRICT).replace(
            "<cx:txData><cx:f>Sheet1!$B$1</cx:f><cx:v>Revenue</cx:v></cx:txData>",
            "<cx:rich><a:bodyPr/><a:lstStyle/><a:p/></cx:rich>",
        );
        let document = ChartExPart::from_part(&chart_part(xml))
            .unwrap()
            .parse()
            .unwrap();
        assert!(matches!(
            document.info().series[0].text,
            Some(ChartExText::Rich(ChartExDrawingPayload {
                child_elements: 3,
                ..
            }))
        ));
    }

    #[test]
    fn rejects_invalid_series_formatting_colors_points_and_labels() {
        let base = producer_xml("0");
        let cases = [
            base.replace("</cx:txData></cx:tx>", "</cx:txData><cx:rich/></cx:tx>"),
            base.replace("<cx:f>Sheet1!$B$1</cx:f><cx:v>Revenue</cx:v>", "<cx:v>Revenue</cx:v><cx:f>Sheet1!$B$1</cx:f>"),
            base.replacen("<cx:spPr>", "<a:spPr>", 1).replacen("</cx:spPr>", "</a:spPr>", 1),
            base.replacen("<a:solidFill>", "<cx:solidFill>", 1).replacen("</a:solidFill>", "</cx:solidFill>", 1),
            base.replace("val=\"0000FF\"", "val=\"00GGFF\""),
            base.replace("<cx:minColor>", "<cx:maxColor>").replace("</cx:minColor>", "</cx:maxColor>"),
            base.replace("count=\"3\"", "count=\"4\""),
            base.replace("<cx:mid><cx:percent val=\"50\"/></cx:mid>", "<cx:mid><cx:extreme/></cx:mid>"),
            base.replace("<cx:percent val=\"50\"/>", "<cx:percent val=\"101\"/>"),
            base.replace("<cx:number val=\"10\"/>", "<cx:number val=\"not-number\"/>"),
            base.replace("</cx:dataPt>", "</cx:dataPt><cx:dataPt idx=\"1\"/>"),
            base.replace("<cx:dataPt idx=\"1\">", "<cx:dataPt idx=\"99\">"),
            base.replace("<cx:dataLabels pos=\"bestFit\">", "<cx:dataLabels pos=\"vendor\">"),
            base.replace("<cx:numFmt formatCode=\"0.0\" sourceLinked=\"0\"/>", "<cx:numFmt sourceLinked=\"0\"/>"),
            base.replace("categoryName=\"true\"", "categoryName=\"yes\""),
            base.replace("<cx:dataLabelHidden idx=\"1\"/>", "<cx:dataLabelHidden idx=\"0\"/>"),
            base.replace("<cx:dataLabel idx=\"0\"", "<cx:dataLabel idx=\"99\""),
            base.replace("<cx:dataLabelHidden idx=\"1\"/>", "<cx:dataLabelHidden idx=\"99\"/>"),
            base.replace("<cx:visibility seriesName=\"1\" categoryName=\"true\" value=\"0\"/><cx:separator> | </cx:separator>", "<cx:separator> | </cx:separator><cx:visibility seriesName=\"1\" categoryName=\"true\" value=\"0\"/>"),
            base.replace("</cx:dataLabels><cx:dataId", "</cx:dataLabels><cx:dataLabels/><cx:dataId"),
            base.replace("<cx:dataLabel idx=\"0\" pos=\"t\"><cx:spPr/>", "<cx:dataLabel idx=\"0\" pos=\"t\"><cx:spPr/><cx:spPr/>"),
        ];
        for xml in cases {
            assert!(
                ChartExPart::from_part(&chart_part(xml))
                    .unwrap()
                    .parse()
                    .is_err()
            );
        }
        let oversized = base.replace(" | ", &"x".repeat(MAX_LABEL_TEXT_BYTES + 1));
        assert!(
            ChartExPart::from_part(&chart_part(oversized))
                .unwrap()
                .parse()
                .is_err()
        );
    }

    #[test]
    fn applies_title_legend_defaults_and_accepts_mp_feature_offsets() {
        let xml = producer_xml("0")
            .replace(
                " version=\"1.0\" featureList=\"sunburst\"",
                " version=\"0.0\" featureList=\"mp\"",
            )
            .replace(" pos=\"t\" align=\"ctr\" overlay=\"1\"", "")
            .replace(" pos=\"r\" align=\"max\" overlay=\"false\"", "");
        let document = ChartExPart::from_part(&chart_part(xml))
            .unwrap()
            .parse()
            .unwrap();
        let title = document.info().chart.title.as_ref().unwrap();
        assert_eq!(
            (title.position, title.alignment, title.overlay),
            (
                ChartExSidePosition::Top,
                ChartExPositionAlignment::Center,
                false
            )
        );
        let legend = document.info().chart.legend.as_ref().unwrap();
        assert_eq!(
            (legend.position, legend.alignment, legend.overlay),
            (
                ChartExSidePosition::Right,
                ChartExPositionAlignment::Center,
                false
            )
        );
    }

    #[test]
    fn rejects_invalid_chart_title_legend_and_offset_grammar() {
        let base = producer_xml("0");
        let cases = [
            base.replace("</cx:title><cx:plotArea>", "</cx:title><cx:title/><cx:plotArea>"),
            base.replace("<cx:plotArea>", "<cx:legend/><cx:plotArea>"),
            base.replace("<cx:extLst/></cx:chart>", "<cx:extLst/><cx:extLst/></cx:chart>"),
            base.replace("<cx:chart>", "<cx:chart vendor=\"1\">"),
            base.replace("<cx:title pos=\"t\"", "<cx:title pos=\"vendor\""),
            base.replace("align=\"ctr\" overlay=\"1\"", "align=\"middle\" overlay=\"1\""),
            base.replace("overlay=\"1\"", "overlay=\"yes\""),
            base.replace("<cx:tx><cx:txData><cx:v>Quarterly</cx:v></cx:txData></cx:tx><cx:spPr>", "<cx:spPr/><cx:tx><cx:txData><cx:v>Quarterly</cx:v></cx:txData></cx:tx><cx:spPr>"),
            base.replace("<cx:tx><cx:txData><cx:v>Quarterly</cx:v></cx:txData></cx:tx>", "<cx:tx/>"),
            base.replacen("<a:noFill/>", "<cx:noFill/>", 1),
            base.replace("<cx:offset top=\"0.1\" left=\"-0.2\"/>", "<cx:offset left=\"-0.2\"/>"),
            base.replace("top=\"0.1\" left=\"-0.2\"", "top=\"bad\" left=\"-0.2\""),
            base.replace("<cx:offset top=\"0.1\" left=\"-0.2\"/>", "<cx:offset top=\"0.1\" left=\"-0.2\"><cx:x/></cx:offset>"),
            base.replace(" version=\"1.0\"", " version=\"0.0\""),
            base.replace("<cx:legend pos=\"r\"", "<cx:legend pos=\"side\""),
            base.replace("align=\"max\" overlay=\"false\"", "align=\"middle\" overlay=\"false\""),
            base.replace("overlay=\"false\"", "overlay=\"off\""),
            base.replace("<cx:txPr><a:bodyPr/><a:lstStyle/><a:p/></cx:txPr><cx:offset top=\"INF\"", "<cx:offset top=\"INF\" left=\"0\"/><cx:txPr><a:bodyPr/><a:lstStyle/><a:p/></cx:txPr><cx:offset top=\"INF\""),
            base.replace("<cx:legend pos=\"r\" align=\"max\" overlay=\"false\"><cx:spPr>", "<cx:legend pos=\"r\" align=\"max\" overlay=\"false\"><cx:spPr/><cx:spPr>"),
            base.replace("<cx:plotArea>", "<cx:vendorPlotArea>").replace("</cx:plotArea>", "</cx:vendorPlotArea>"),
        ];
        for xml in cases {
            assert!(
                ChartExPart::from_part(&chart_part(xml))
                    .unwrap()
                    .parse()
                    .is_err()
            );
        }
    }

    #[test]
    fn rejects_auto_update_external_missing_wrong_type_and_ambiguous_targets() {
        for (auto, rel_type, target, external) in [
            ("1", PACKAGE_REL[0], "../embeddings/data.xlsx", false),
            ("0", IMAGE_REL[0], "../embeddings/data.xlsx", false),
            ("0", PACKAGE_REL[0], "https://example.test/data.xlsx", true),
            ("0", PACKAGE_REL[0], "../embeddings/data.xlsx#x", false),
        ] {
            let mut chart = chart_part(producer_xml(auto).replace(" fallbackImg=\"rIdImage\"", ""));
            chart.rels_mut().add_relationship(
                rel_type.into(),
                target.into(),
                "rIdWorkbook".into(),
                external,
            );
            let mut package = OpcPackage::new();
            package.add_part(Box::new(BlobPart::new(
                PackURI::new("/ppt/embeddings/data.xlsx").unwrap(),
                WORKBOOK_CONTENT_TYPES[0].into(),
                Vec::new(),
            )));
            assert!(
                ChartExPart::from_part(&chart)
                    .unwrap()
                    .parse_in_package(&package)
                    .is_err()
            );
        }
    }
}
