//! Contextual values for the bounded ChartEx container and data-index graph.

use super::super::style::{ColorDocument, Document as StyleDocument};

/// Typed metadata from the bounded ChartEx container and data-index core.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Info {
    pub version: String,
    pub features: Vec<String>,
    pub fallback_image_relationship_id: Option<String>,
    pub data_sets: Vec<DataSet>,
    pub series: Vec<SeriesDataReference>,
    pub axes: Vec<Axis>,
    pub has_plot_surface: bool,
    pub chart: Chart,
    pub plot_area: PlotArea,
    pub chart_space_formatting: ChartSpaceFormatting,
    pub external_data: Option<ExternalData>,
    pub has_title: bool,
    pub has_legend: bool,
}

#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct PlotArea {
    pub region: PlotAreaRegion,
    pub shape_properties: Option<DrawingPayload>,
    pub has_extension_list: bool,
}

#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct PlotAreaRegion {
    pub plot_surface: Option<PlotSurface>,
    pub has_extension_list: bool,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PlotSurface {
    pub shape_properties: Option<DrawingPayload>,
    pub has_extension_list: bool,
}

#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct ChartSpaceFormatting {
    pub shape_properties: Option<DrawingPayload>,
    pub text_properties: Option<DrawingPayload>,
    pub color_mapping_override: Option<DrawingPayload>,
    pub format_overrides: Option<Vec<FormatOverride>>,
    pub print_settings: Option<PrintSettings>,
    pub has_extension_list: bool,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct FormatOverride {
    pub index: u32,
    pub shape_properties: Option<DrawingPayload>,
    pub has_extension_list: bool,
}

#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct PrintSettings {
    pub header_footer: Option<HeaderFooter>,
    pub page_margins: Option<PageMargins>,
    pub page_setup: Option<PageSetup>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct HeaderFooter {
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
pub struct PageMargins {
    pub left: String,
    pub right: String,
    pub top: String,
    pub bottom: String,
    pub header: String,
    pub footer: String,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PageOrientation {
    Default,
    Portrait,
    Landscape,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PageSetup {
    pub paper_size: u32,
    pub first_page_number: u32,
    pub orientation: PageOrientation,
    pub black_and_white: bool,
    pub draft: bool,
    pub use_first_page_number: bool,
    pub horizontal_dpi: i32,
    pub vertical_dpi: i32,
    pub copies: u32,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DataSet {
    pub id: u32,
    pub string_dimensions: usize,
    pub numeric_dimensions: usize,
    pub dimensions: Vec<Dimension>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum StringDimensionType {
    Category,
    ColorString,
    EntityId,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum NumericDimensionType {
    Value,
    X,
    Y,
    Size,
    ColorValue,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FormulaDirection {
    Column,
    Row,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Formula {
    pub expression: String,
    pub direction: FormulaDirection,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum Dimension {
    String {
        kind: StringDimensionType,
        formula: Option<Formula>,
        name_formula: Option<Formula>,
        levels: Vec<StringLevel>,
    },
    Numeric {
        kind: NumericDimensionType,
        formula: Option<Formula>,
        name_formula: Option<Formula>,
        levels: Vec<NumericLevel>,
    },
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct StringLevel {
    pub point_count: u32,
    pub name: Option<String>,
    pub points: Vec<StringPoint>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct NumericLevel {
    pub point_count: u32,
    pub name: Option<String>,
    pub format_code: Option<String>,
    pub points: Vec<NumericPoint>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct StringPoint {
    pub index: u32,
    pub value: String,
}

/// Numeric values retain their XML Schema double lexical form (`INF`, `-INF`, and `NaN` included).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct NumericPoint {
    pub index: u32,
    pub value: String,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SeriesLayout {
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
pub struct SeriesDataReference {
    pub layout: SeriesLayout,
    pub text: Option<Text>,
    pub shape_properties: Option<DrawingPayload>,
    pub value_colors: Option<ValueColors>,
    pub value_color_positions: Option<ValueColorPositions>,
    pub data_points: Vec<DataPoint>,
    pub data_labels: Option<DataLabels>,
    pub data_id: Option<u32>,
    pub hidden: bool,
    pub owner_index: Option<u32>,
    pub unique_id: Option<String>,
    pub format_index: Option<u32>,
    pub layout_properties: Option<LayoutProperties>,
    pub axis_ids: Vec<u32>,
}

/// A DrawingML subtree retained by the document's lossless source XML.
/// Only its bounded, namespace-checked outer payload is exposed here.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DrawingPayload {
    pub child_elements: usize,
    pub attributes: usize,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum Text {
    Data {
        formula: Option<Formula>,
        value: Option<String>,
    },
    Rich(DrawingPayload),
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ColorKind {
    ScRgb,
    Srgb,
    Hsl,
    System,
    Scheme,
    Preset,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SolidColor {
    pub kind: ColorKind,
    pub value: Option<String>,
    pub modifier_count: usize,
}

#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct ValueColors {
    pub minimum: Option<SolidColor>,
    pub middle: Option<SolidColor>,
    pub maximum: Option<SolidColor>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ColorPosition {
    Extreme,
    Number(String),
    Percent(String),
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ValueColorPositions {
    pub count: u8,
    pub minimum: Option<ColorPosition>,
    pub middle: Option<ColorPosition>,
    pub maximum: Option<ColorPosition>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DataPoint {
    pub index: u32,
    pub shape_properties: Option<DrawingPayload>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DataLabelPosition {
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
pub struct NumberFormat {
    pub format_code: String,
    pub source_linked: Option<bool>,
}

#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct DataLabelVisibility {
    pub series_name: Option<bool>,
    pub category_name: Option<bool>,
    pub value: Option<bool>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DataLabel {
    pub index: u32,
    pub position: Option<DataLabelPosition>,
    pub number_format: Option<NumberFormat>,
    pub shape_properties: Option<DrawingPayload>,
    pub text_properties: Option<DrawingPayload>,
    pub visibility: Option<DataLabelVisibility>,
    pub separator: Option<String>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DataLabels {
    pub position: Option<DataLabelPosition>,
    pub number_format: Option<NumberFormat>,
    pub shape_properties: Option<DrawingPayload>,
    pub text_properties: Option<DrawingPayload>,
    pub visibility: Option<DataLabelVisibility>,
    pub separator: Option<String>,
    pub labels: Vec<DataLabel>,
    pub hidden_indices: Vec<u32>,
}

#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct Chart {
    pub title: Option<ChartTitle>,
    pub legend: Option<Legend>,
    pub has_extension_list: bool,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SidePosition {
    Left,
    Right,
    Top,
    Bottom,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PositionAlignment {
    Minimum,
    Center,
    Maximum,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Offset {
    pub top: String,
    pub left: String,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChartTitle {
    pub position: SidePosition,
    pub alignment: PositionAlignment,
    pub overlay: bool,
    pub text: Option<Text>,
    pub shape_properties: Option<DrawingPayload>,
    pub text_properties: Option<DrawingPayload>,
    pub offset: Option<Offset>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Legend {
    pub position: SidePosition,
    pub alignment: PositionAlignment,
    pub overlay: bool,
    pub shape_properties: Option<DrawingPayload>,
    pub text_properties: Option<DrawingPayload>,
    pub offset: Option<Offset>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ParentLabelLayout {
    None,
    Banner,
    Overlapping,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum RegionLabelLayout {
    None,
    BestFitOnly,
    ShowAll,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ClosedSide {
    Left,
    Right,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum QuartileMethod {
    Inclusive,
    Exclusive,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum GeoProjection {
    Mercator,
    Miller,
    Robinson,
    Albers,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum GeoMappingLevel {
    DataOnly,
    PostalCode,
    County,
    State,
    CountryRegion,
    CountryRegionList,
    World,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum DoubleOrAutomatic {
    Automatic,
    Number(String),
}

#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct ElementVisibility {
    pub connector_lines: Option<bool>,
    pub mean_line: Option<bool>,
    pub mean_marker: Option<bool>,
    pub nonoutliers: Option<bool>,
    pub outliers: Option<bool>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum BinningChoice {
    Size(String),
    Count(u32),
}

#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct Binning {
    pub choice: Option<BinningChoice>,
    pub interval_closed: Option<ClosedSide>,
    pub underflow: Option<DoubleOrAutomatic>,
    pub overflow: Option<DoubleOrAutomatic>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Geography {
    pub projection: Option<GeoProjection>,
    pub viewed_region: Option<GeoMappingLevel>,
    pub culture_language: String,
    pub culture_region: String,
    pub attribution: String,
    pub has_cache: bool,
    pub cache: Option<GeoCache>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct GeoCache {
    pub provider: String,
    pub entries: Vec<GeoCacheEntry>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum GeoCacheEntry {
    Binary {
        encoded_characters: usize,
        decoded_bytes: usize,
    },
    Clear(GeoClear),
}

#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct GeoClear {
    pub location_query_results: Option<Vec<GeoLocationQueryResult>>,
    pub data_entity_query_results: Option<Vec<GeoDataEntityQueryResult>>,
    pub data_point_to_entity_query_results: Option<Vec<GeoDataPointToEntityQueryResult>>,
    pub child_entities_query_results: Option<Vec<GeoChildEntitiesQueryResult>>,
    pub parent_entities_query_results: Option<Vec<GeoParentEntitiesQueryResult>>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum GeoEntityType {
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
pub struct GeoLocationQuery {
    pub country_region: Option<String>,
    pub admin_district1: Option<String>,
    pub admin_district2: Option<String>,
    pub postal_code: Option<String>,
    pub entity_type: GeoEntityType,
}

#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct GeoAddress {
    pub address1: Option<String>,
    pub country_region: Option<String>,
    pub admin_district1: Option<String>,
    pub admin_district2: Option<String>,
    pub postal_code: Option<String>,
    pub locality: Option<String>,
    pub iso_country_code: Option<String>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct GeoLocation {
    pub latitude: Option<String>,
    pub longitude: Option<String>,
    pub entity_name: String,
    pub entity_type: GeoEntityType,
    pub address: Option<GeoAddress>,
}

#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct GeoLocationQueryResult {
    pub query: Option<GeoLocationQuery>,
    pub location: Option<GeoLocation>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct GeoPolygon {
    pub polygon_id: String,
    pub num_points: String,
    pub pca_rings: String,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct GeoData {
    pub entity_name: String,
    pub entity_id: String,
    pub east: String,
    pub west: String,
    pub north: String,
    pub south: String,
    pub polygons: Option<Vec<GeoPolygon>>,
    pub copyrights: Option<Vec<String>>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct GeoDataEntityQuery {
    pub entity_type: GeoEntityType,
    pub entity_id: String,
}

#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct GeoDataEntityQueryResult {
    pub query: Option<GeoDataEntityQuery>,
    pub data: Option<GeoData>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct GeoDataPointQuery {
    pub entity_type: GeoEntityType,
    pub latitude: String,
    pub longitude: String,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct GeoDataPointToEntityQuery {
    pub entity_type: GeoEntityType,
    pub entity_id: String,
}

#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct GeoDataPointToEntityQueryResult {
    pub point_query: Option<GeoDataPointQuery>,
    pub entity_query: Option<GeoDataPointToEntityQuery>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct GeoChildEntitiesQuery {
    pub entity_id: String,
    pub child_types: Option<Vec<GeoEntityType>>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct GeoHierarchyEntity {
    pub entity_name: String,
    pub entity_id: String,
    pub entity_type: GeoEntityType,
}

#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct GeoChildEntitiesQueryResult {
    pub query: Option<GeoChildEntitiesQuery>,
    pub children: Option<Vec<GeoHierarchyEntity>>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct GeoEntity {
    pub entity_name: String,
    pub entity_type: GeoEntityType,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct GeoParentEntitiesQueryResult {
    pub entity_id: String,
    pub entity: Option<GeoEntity>,
    pub parent_entity_id: Option<String>,
}

#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct LayoutProperties {
    pub parent_label: Option<ParentLabelLayout>,
    pub region_label: Option<RegionLabelLayout>,
    pub visibility: Option<ElementVisibility>,
    pub aggregation: bool,
    pub binning: Option<Binning>,
    pub geography: Option<Geography>,
    pub quartile_method: Option<QuartileMethod>,
    pub subtotals: Vec<u32>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum AxisScaling {
    Category {
        gap_width: Option<DoubleOrAutomatic>,
    },
    Value {
        minimum: Option<DoubleOrAutomatic>,
        maximum: Option<DoubleOrAutomatic>,
        major_unit: Option<DoubleOrAutomatic>,
        minor_unit: Option<DoubleOrAutomatic>,
    },
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct AxisTitle {
    pub text: Option<Text>,
    pub shape_properties: Option<DrawingPayload>,
    pub text_properties: Option<DrawingPayload>,
    pub offset: Option<Offset>,
    pub has_extension_list: bool,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum AxisUnit {
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
pub struct AxisUnitsLabel {
    pub text: Option<Text>,
    pub shape_properties: Option<DrawingPayload>,
    pub text_properties: Option<DrawingPayload>,
    pub has_extension_list: bool,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct AxisUnits {
    pub unit: Option<AxisUnit>,
    pub label: Option<AxisUnitsLabel>,
    pub has_extension_list: bool,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Gridlines {
    pub shape_properties: Option<DrawingPayload>,
    pub has_extension_list: bool,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TickMarkType {
    Inside,
    Outside,
    Cross,
    None,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TickMarks {
    pub kind: Option<TickMarkType>,
    pub has_extension_list: bool,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TickLabels {
    pub has_extension_list: bool,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Axis {
    pub id: u32,
    pub hidden: bool,
    pub scaling: AxisScaling,
    pub title: Option<AxisTitle>,
    pub units: Option<AxisUnits>,
    pub major_gridlines: Option<Gridlines>,
    pub minor_gridlines: Option<Gridlines>,
    pub major_tick_marks: Option<TickMarks>,
    pub minor_tick_marks: Option<TickMarks>,
    pub tick_labels: Option<TickLabels>,
    pub number_format: Option<NumberFormat>,
    pub shape_properties: Option<DrawingPayload>,
    pub text_properties: Option<DrawingPayload>,
    pub has_extension_list: bool,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ExternalData {
    pub relationship_id: String,
    pub auto_update: bool,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ExternalDataTarget {
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
pub struct Document {
    pub(crate) info: Info,
    pub(crate) xml: Vec<u8>,
    pub(crate) external_data_target: Option<ExternalDataTarget>,
    pub(crate) fallback_image_part_name: Option<String>,
    pub(crate) chart_style: Option<StyleDocument>,
    pub(crate) chart_color_style: Option<ColorDocument>,
}
