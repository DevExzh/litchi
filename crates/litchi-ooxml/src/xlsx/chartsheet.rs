//! Typed SpreadsheetML chartsheets and their inert workbook/drawing/chart package graph.

use crate::error::{OoxmlError, Result};
use crate::xlsx::printer_settings::{
    MAX_SETTINGS_BYTES, PRINTER_CT, PRINTER_REL, PrinterSettingsResource, STRICT_PRINTER_REL,
    is_printer_relationship, validate_printer_settings_uri, validate_settings_bytes,
};
use litchi_ooxml_common::{MceCapabilities, MceLimits, process_markup_compatibility};
use litchi_opc::constants::relationship_type as rt;
use litchi_opc::{BlobPart, OpcPackage, PackURI, Part};
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use std::collections::{BTreeMap, BTreeSet, HashSet};

const SML: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const STRICT_SML: &str = "http://purl.oclc.org/ooxml/spreadsheetml/main";
const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const STRICT_REL: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships";
const XDR: &str = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
const STRICT_XDR: &str = "http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing";
const CHART: &str = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const STRICT_CHART: &str = "http://purl.oclc.org/ooxml/drawingml/chart";
const DRAWING_MAIN: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const STRICT_DRAWING_MAIN: &str = "http://purl.oclc.org/ooxml/drawingml/main";
const CHART_EX: &str = "http://schemas.microsoft.com/office/drawing/2014/chartex";
const CHART_EX_CHOICE: &str = "http://schemas.microsoft.com/office/drawing/2015/9/8/chartex";
const CHART_STYLE: &str = "http://schemas.microsoft.com/office/drawing/2012/chartStyle";
const CHARTSHEET_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet";
const STRICT_CHARTSHEET_REL: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/chartsheet";
const IMAGE_REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
const STRICT_IMAGE_REL: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships/image";
const CHART_EX_REL: &str = "http://schemas.microsoft.com/office/2014/relationships/chartEx";
const CHART_STYLE_REL: &str = "http://schemas.microsoft.com/office/2011/relationships/chartStyle";
const CHART_COLOR_STYLE_REL: &str =
    "http://schemas.microsoft.com/office/2011/relationships/chartColorStyle";
const CHART_USER_SHAPES_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartUserShapes";
const STRICT_CHART_USER_SHAPES_REL: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/chartUserShapes";
const THEME_OVERRIDE_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/themeOverride";
const STRICT_THEME_OVERRIDE_REL: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/themeOverride";
const VML_DRAWING_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing";
const STRICT_VML_DRAWING_REL: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/vmlDrawing";
const CHARTSHEET_CT: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml";
const DRAWING_CT: &str = "application/vnd.openxmlformats-officedocument.drawing+xml";
const CHART_CT: &str = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";
const CHART_EX_CT: &str = "application/vnd.ms-office.chartex+xml";
const CHART_STYLE_CT: &str = "application/vnd.ms-office.chartstyle+xml";
const CHART_COLOR_STYLE_CT: &str = "application/vnd.ms-office.chartcolorstyle+xml";
const CHART_USER_SHAPES_CT: &str =
    "application/vnd.openxmlformats-officedocument.drawingml.chartshapes+xml";
const THEME_OVERRIDE_CT: &str = "application/vnd.openxmlformats-officedocument.themeOverride+xml";
const VML_DRAWING_CT: &str = "application/vnd.openxmlformats-officedocument.vmlDrawing";
const MAX_XML_BYTES: usize = 32 * 1024 * 1024;
const MAX_DRAWING_BYTES: usize = 16 * 1024 * 1024;
const MAX_CHART_BYTES: usize = 32 * 1024 * 1024;
const MAX_BACKGROUND_IMAGE_BYTES: usize = 32 * 1024 * 1024;
const MAX_VML_DRAWING_BYTES: usize = 16 * 1024 * 1024;
const MAX_TOTAL_RESOURCE_BYTES: usize = 128 * 1024 * 1024;
const MAX_NODES: usize = 500_000;
const MAX_DEPTH: usize = 256;
const MAX_STRING_BYTES: usize = 4 * 1024 * 1024;
const MAX_VIEWS: usize = 256;
const MAX_CUSTOM_VIEWS: usize = 1024;
const MAX_CHARTS: usize = 256;
const MAX_CHART_EX_BYTES: usize = 16 * 1024 * 1024;
const MAX_CHART_STYLE_BYTES: usize = 4 * 1024 * 1024;
const MAX_CHART_COLOR_STYLE_BYTES: usize = 4 * 1024 * 1024;
const MAX_CHART_STYLE_PARTS: usize = 16;
const MAX_CHART_USER_SHAPES_BYTES: usize = 8 * 1024 * 1024;
const MAX_CHART_USER_SHAPE_IMAGES: usize = 256;
const MAX_CHART_USER_SHAPE_IMAGE_BYTES: usize = 32 * 1024 * 1024;
const MAX_CHART_DIRECT_IMAGES: usize = 256;
const MAX_CHART_THEME_OVERRIDE_BYTES: usize = 4 * 1024 * 1024;
const MAX_CHART_THEME_IMAGES: usize = 256;
const MAX_CHART_EMBEDDED_PACKAGE_BYTES: usize = 64 * 1024 * 1024;
const MAX_WEB_PUBLISH_ITEMS: usize = 4096;
const MAX_WEB_PUBLISH_STRING_BYTES: usize = 64 * 1024;
const MAX_EXTENSIONS: usize = 1024;
const MAX_EXTENSION_URI_BYTES: usize = 1024;
const MAX_EXTENSION_PAYLOAD_BYTES: usize = 1024 * 1024;
const MAX_RETAINED_EXTENSION_BYTES: usize = 4 * 1024 * 1024;
const MAX_EXTENSION_RELATIONSHIPS: usize = 1024;
const MAX_EXTENSION_RELATIONSHIP_STRING_BYTES: usize = 64 * 1024;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ChartSheetConformance {
    Transitional,
    Strict,
}

impl ChartSheetConformance {
    fn sml(self) -> &'static str {
        match self {
            Self::Transitional => SML,
            Self::Strict => STRICT_SML,
        }
    }
    fn rel(self) -> &'static str {
        match self {
            Self::Transitional => REL,
            Self::Strict => STRICT_REL,
        }
    }
    fn xdr(self) -> &'static str {
        match self {
            Self::Transitional => XDR,
            Self::Strict => STRICT_XDR,
        }
    }
    fn chart(self) -> &'static str {
        match self {
            Self::Transitional => CHART,
            Self::Strict => STRICT_CHART,
        }
    }
    fn chartsheet_rel(self) -> &'static str {
        match self {
            Self::Transitional => CHARTSHEET_REL,
            Self::Strict => STRICT_CHARTSHEET_REL,
        }
    }
    fn drawing_rel(self) -> &'static str {
        match self {
            Self::Transitional => rt::DRAWING,
            Self::Strict => rt::STRICT_DRAWING,
        }
    }
    fn chart_rel(self) -> &'static str {
        match self {
            Self::Transitional => rt::CHART,
            Self::Strict => rt::STRICT_CHART,
        }
    }
    fn image_rel(self) -> &'static str {
        match self {
            Self::Transitional => IMAGE_REL,
            Self::Strict => STRICT_IMAGE_REL,
        }
    }
    fn chart_user_shapes_rel(self) -> &'static str {
        match self {
            Self::Transitional => CHART_USER_SHAPES_REL,
            Self::Strict => STRICT_CHART_USER_SHAPES_REL,
        }
    }
    fn theme_override_rel(self) -> &'static str {
        match self {
            Self::Transitional => THEME_OVERRIDE_REL,
            Self::Strict => STRICT_THEME_OVERRIDE_REL,
        }
    }
    fn package_rel(self) -> &'static str {
        match self {
            Self::Transitional => rt::PACKAGE,
            Self::Strict => rt::STRICT_PACKAGE,
        }
    }
    fn vml_drawing_rel(self) -> &'static str {
        match self {
            Self::Transitional => VML_DRAWING_REL,
            Self::Strict => STRICT_VML_DRAWING_REL,
        }
    }
    fn printer_rel(self) -> &'static str {
        match self {
            Self::Transitional => PRINTER_REL,
            Self::Strict => STRICT_PRINTER_REL,
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ChartSheetState {
    Visible,
    Hidden,
    VeryHidden,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PageOrientation {
    Default,
    Portrait,
    Landscape,
}

#[derive(Debug, Clone, PartialEq)]
pub struct ChartSheetColor {
    pub automatic: Option<bool>,
    pub indexed: Option<u32>,
    pub rgb: Option<String>,
    pub theme: Option<u32>,
    pub tint: Option<f64>,
}

#[derive(Debug, Clone, PartialEq)]
pub struct ChartSheetProperties {
    pub published: Option<bool>,
    pub code_name: Option<String>,
    pub tab_color: Option<ChartSheetColor>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChartSheetView {
    pub tab_selected: Option<bool>,
    pub zoom_scale: Option<u32>,
    pub workbook_view_id: u32,
    pub zoom_to_fit: Option<bool>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChartSheetProtection {
    pub password_hash: Option<String>,
    pub content: Option<bool>,
    pub objects: Option<bool>,
}

/// One saved chartsheet view from `CT_CustomChartsheetView`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChartSheetCustomView {
    /// Braced UUID lexical form required by SpreadsheetML `ST_Guid`.
    pub guid: String,
    pub scale: Option<u32>,
    pub state: Option<ChartSheetState>,
    pub zoom_to_fit: Option<bool>,
}

#[derive(Debug, Clone, Copy, PartialEq)]
pub struct ChartSheetMargins {
    pub left: f64,
    pub right: f64,
    pub top: f64,
    pub bottom: f64,
    pub header: f64,
    pub footer: f64,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChartSheetPageSetup {
    pub paper_size: Option<u32>,
    pub first_page_number: Option<u32>,
    pub orientation: Option<PageOrientation>,
    pub use_printer_defaults: Option<bool>,
    pub black_and_white: Option<bool>,
    pub draft: Option<bool>,
    pub use_first_page_number: Option<bool>,
    pub horizontal_dpi: Option<u32>,
    pub vertical_dpi: Option<u32>,
    pub copies: Option<u32>,
    /// Inert relationship reference to a binary Printer Settings part.
    pub printer_settings_relationship_id: Option<String>,
}

#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct ChartSheetHeaderFooter {
    pub different_odd_even: Option<bool>,
    pub different_first: Option<bool>,
    pub scale_with_document: Option<bool>,
    pub align_with_margins: Option<bool>,
    pub odd_header: Option<String>,
    pub odd_footer: Option<String>,
    pub even_header: Option<String>,
    pub even_footer: Option<String>,
    pub first_header: Option<String>,
    pub first_footer: Option<String>,
}

/// Schema-complete `ST_WebSourceType` values for inert web-publishing metadata.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ChartSheetWebSourceType {
    Sheet,
    PrintArea,
    AutoFilter,
    Range,
    Chart,
    PivotTable,
    Query,
    Label,
}

impl ChartSheetWebSourceType {
    pub fn as_str(self) -> &'static str {
        match self {
            Self::Sheet => "sheet",
            Self::PrintArea => "printArea",
            Self::AutoFilter => "autoFilter",
            Self::Range => "range",
            Self::Chart => "chart",
            Self::PivotTable => "pivotTable",
            Self::Query => "query",
            Self::Label => "label",
        }
    }
    fn parse(value: &str) -> Result<Self> {
        match value {
            "sheet" => Ok(Self::Sheet),
            "printArea" => Ok(Self::PrintArea),
            "autoFilter" => Ok(Self::AutoFilter),
            "range" => Ok(Self::Range),
            "chart" => Ok(Self::Chart),
            "pivotTable" => Ok(Self::PivotTable),
            "query" => Ok(Self::Query),
            "label" => Ok(Self::Label),
            _ => Err(invalid(format!("invalid web publish sourceType '{value}'"))),
        }
    }
}

/// One `webPublishItem`; all paths, names, and references are opaque inert strings.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChartSheetWebPublishItem {
    pub id: u32,
    pub div_id: String,
    pub source_type: ChartSheetWebSourceType,
    pub source_ref: Option<String>,
    pub source_object: Option<String>,
    pub destination_file: String,
    pub title: Option<String>,
    /// `None` preserves the schema default; no publishing action is ever performed.
    pub auto_republish: Option<bool>,
}

/// A present `webPublishItems` collection is schema-required to be non-empty.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChartSheetWebPublishItems {
    /// Preserves explicit `count`; when present it must equal `items.len()`.
    pub count: Option<u32>,
    pub items: Vec<ChartSheetWebPublishItem>,
}

#[derive(Debug, Clone, PartialEq)]
pub struct ChartSheet {
    pub properties: Option<ChartSheetProperties>,
    pub views: Vec<ChartSheetView>,
    pub protection: Option<ChartSheetProtection>,
    /// `None` preserves absence; a present collection must be non-empty.
    pub custom_views: Option<Vec<ChartSheetCustomView>>,
    pub margins: Option<ChartSheetMargins>,
    pub page_setup: Option<ChartSheetPageSetup>,
    pub header_footer: Option<ChartSheetHeaderFooter>,
    pub drawing_relationship_id: String,
    pub legacy_drawing_relationship_id: Option<String>,
    pub legacy_header_footer_drawing_relationship_id: Option<String>,
    /// Relationship for the optional tiled chartsheet background image.
    pub background_picture_relationship_id: Option<String>,
    /// Inert web-publishing metadata. Destinations and sources are never resolved or accessed.
    pub web_publish_items: Option<ChartSheetWebPublishItems>,
    /// Inert, canonicalized wildcard extension markup. Payload semantics are never interpreted.
    pub extension_list: Option<ChartSheetExtensionList>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChartSheetExtension {
    pub uri: String,
    /// Canonical, namespace-aware XML for the single wildcard child of `ext`.
    pub payload_xml: Vec<u8>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChartSheetExtensionList {
    pub extensions: Vec<ChartSheetExtension>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ChartSheetExtensionRelationshipTarget {
    Internal { part_name: String },
    External { target: String },
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChartSheetExtensionRelationship {
    pub relationship_id: String,
    pub relationship_type: String,
    pub target: ChartSheetExtensionRelationshipTarget,
}

/// Supported inert image media types for chartsheet backgrounds.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ChartSheetImageContentType {
    Png,
    Jpeg,
    Gif,
    Bmp,
    Tiff,
    Emf,
    Wmf,
}

impl ChartSheetImageContentType {
    pub fn as_str(self) -> &'static str {
        match self {
            Self::Png => "image/png",
            Self::Jpeg => "image/jpeg",
            Self::Gif => "image/gif",
            Self::Bmp => "image/bmp",
            Self::Tiff => "image/tiff",
            Self::Emf => "image/x-emf",
            Self::Wmf => "image/x-wmf",
        }
    }
    fn parse(value: &str) -> Result<Self> {
        match value {
            "image/png" => Ok(Self::Png),
            "image/jpeg" => Ok(Self::Jpeg),
            "image/gif" => Ok(Self::Gif),
            "image/bmp" => Ok(Self::Bmp),
            "image/tiff" => Ok(Self::Tiff),
            "image/x-emf" => Ok(Self::Emf),
            "image/x-wmf" => Ok(Self::Wmf),
            _ => Err(invalid(format!(
                "unsupported chartsheet background image content type '{value}'"
            ))),
        }
    }
}

/// Opaque internal package resource referenced by `chartsheet/picture`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChartSheetBackgroundPicture {
    pub relationship_id: String,
    pub part_name: String,
    pub content_type: ChartSheetImageContentType,
    /// Preserved without decoding, rendering, metadata inspection, or external fetches.
    pub data: Vec<u8>,
}

/// Chartsheet `pageSetup` relationship paired with a generic inert DEVMODE resource.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChartSheetPrinterSettings {
    pub relationship_id: String,
    pub resource: PrinterSettingsResource,
}

/// Opaque internal VML drawing bytes; VML semantics are intentionally not interpreted.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChartSheetVmlDrawingResource {
    pub relationship_id: String,
    pub part_name: String,
    pub content_type: String,
    /// Preserved without shape rendering, script, hyperlink, external fetch, or macro execution.
    pub data: Vec<u8>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChartSheetChartResource {
    pub relationship_id: String,
    pub part_name: String,
    pub content_type: String,
    /// Preserved without chart evaluation, external-data loading, or macro execution.
    pub data: Vec<u8>,
    pub kind: ChartSheetChartResourceKind,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChartSheetChartCompanionResource {
    pub relationship_id: String,
    pub part_name: String,
    pub content_type: String,
    /// Preserved without rendering, theme application, formula evaluation, or external access.
    pub data: Vec<u8>,
}

/// Supported inert image payload types for a chartEx chart-user-shapes drawing.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ChartSheetChartImageContentType {
    Bmp,
    Gif,
    Png,
    Tif,
    Tiff,
    Icon,
    Pcx,
    Jpeg,
    Jp2,
    Emf,
    Wmf,
    Svg,
}

impl ChartSheetChartImageContentType {
    pub fn as_str(self) -> &'static str {
        match self {
            Self::Bmp => "image/bmp",
            Self::Gif => "image/gif",
            Self::Png => "image/png",
            Self::Tif => "image/tif",
            Self::Tiff => "image/tiff",
            Self::Icon => "image/x-icon",
            Self::Pcx => "image/x-pcx",
            Self::Jpeg => "image/jpeg",
            Self::Jp2 => "image/jp2",
            Self::Emf => "image/x-emf",
            Self::Wmf => "image/x-wmf",
            Self::Svg => "image/svg+xml",
        }
    }
    fn parse(value: &str) -> Result<Self> {
        match value {
            "image/bmp" => Ok(Self::Bmp),
            "image/gif" => Ok(Self::Gif),
            "image/png" => Ok(Self::Png),
            "image/tif" => Ok(Self::Tif),
            "image/tiff" => Ok(Self::Tiff),
            "image/x-icon" => Ok(Self::Icon),
            "image/x-pcx" => Ok(Self::Pcx),
            "image/jpeg" => Ok(Self::Jpeg),
            "image/jp2" => Ok(Self::Jp2),
            "image/x-emf" => Ok(Self::Emf),
            "image/x-wmf" => Ok(Self::Wmf),
            "image/svg+xml" => Ok(Self::Svg),
            _ => Err(invalid(format!(
                "unsupported chart user-shape image content type '{value}'"
            ))),
        }
    }
    fn validates_part_name(self, value: &str) -> bool {
        let lower = value.to_ascii_lowercase();
        match self {
            Self::Bmp => lower.ends_with(".bmp"),
            Self::Gif => lower.ends_with(".gif"),
            Self::Png => lower.ends_with(".png"),
            Self::Tif => lower.ends_with(".tif"),
            Self::Tiff => lower.ends_with(".tiff"),
            Self::Icon => lower.ends_with(".ico"),
            Self::Pcx => lower.ends_with(".pcx"),
            Self::Jpeg => lower.ends_with(".jpg") || lower.ends_with(".jpeg"),
            Self::Jp2 => lower.ends_with(".jp2"),
            Self::Emf => lower.ends_with(".emf"),
            Self::Wmf => lower.ends_with(".wmf"),
            Self::Svg => lower.ends_with(".svg"),
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChartSheetChartImageResource {
    pub relationship_id: String,
    pub part_name: String,
    pub content_type: ChartSheetChartImageContentType,
    /// Preserved without decoding, rendering, metadata inspection, or external fetches.
    pub data: Vec<u8>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChartSheetChartUserShapesResource {
    pub relationship_id: String,
    pub part_name: String,
    pub content_type: String,
    /// Preserved without interpreting shapes, actions, hyperlinks, or embedded markup.
    pub data: Vec<u8>,
    pub images: Vec<ChartSheetChartImageResource>,
}

pub type ChartSheetChartUserShapeImageContentType = ChartSheetChartImageContentType;
pub type ChartSheetChartUserShapeImageResource = ChartSheetChartImageResource;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ChartSheetChartEmbeddedPackageContentType {
    Docx,
    Dotx,
    Potx,
    Ppsx,
    Pptx,
    Sldx,
    Thmx,
    Xlsx,
    Xltx,
}
impl ChartSheetChartEmbeddedPackageContentType {
    pub fn as_str(self) -> &'static str {
        match self {
            Self::Docx => "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            Self::Dotx => "application/vnd.openxmlformats-officedocument.wordprocessingml.template",
            Self::Potx => "application/vnd.openxmlformats-officedocument.presentationml.template",
            Self::Ppsx => "application/vnd.openxmlformats-officedocument.presentationml.slideshow",
            Self::Pptx => {
                "application/vnd.openxmlformats-officedocument.presentationml.presentation"
            },
            Self::Sldx => "application/vnd.openxmlformats-officedocument.presentationml.slide",
            Self::Thmx => "application/vnd.ms-officetheme",
            Self::Xlsx => "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            Self::Xltx => "application/vnd.openxmlformats-officedocument.spreadsheetml.template",
        }
    }
    fn parse(value: &str) -> Result<Self> {
        match value {
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document" => {
                Ok(Self::Docx)
            },
            "application/vnd.openxmlformats-officedocument.wordprocessingml.template" => {
                Ok(Self::Dotx)
            },
            "application/vnd.openxmlformats-officedocument.presentationml.template" => {
                Ok(Self::Potx)
            },
            "application/vnd.openxmlformats-officedocument.presentationml.slideshow" => {
                Ok(Self::Ppsx)
            },
            "application/vnd.openxmlformats-officedocument.presentationml.presentation" => {
                Ok(Self::Pptx)
            },
            "application/vnd.openxmlformats-officedocument.presentationml.slide" => Ok(Self::Sldx),
            "application/vnd.ms-officetheme" => Ok(Self::Thmx),
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" => Ok(Self::Xlsx),
            "application/vnd.openxmlformats-officedocument.spreadsheetml.template" => {
                Ok(Self::Xltx)
            },
            _ => Err(invalid(format!(
                "unsupported or active chartEx embedded-package content type '{value}'"
            ))),
        }
    }
    fn validates_part_name(self, value: &str) -> bool {
        let suffix = match self {
            Self::Docx => ".docx",
            Self::Dotx => ".dotx",
            Self::Potx => ".potx",
            Self::Ppsx => ".ppsx",
            Self::Pptx => ".pptx",
            Self::Sldx => ".sldx",
            Self::Thmx => ".thmx",
            Self::Xlsx => ".xlsx",
            Self::Xltx => ".xltx",
        };
        value.to_ascii_lowercase().ends_with(suffix)
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChartSheetChartThemeOverrideResource {
    pub relationship_id: String,
    pub part_name: String,
    pub content_type: String,
    /// Preserved without theme application, image rendering, or external access.
    pub data: Vec<u8>,
    pub images: Vec<ChartSheetChartImageResource>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChartSheetChartEmbeddedPackageResource {
    pub relationship_id: String,
    pub part_name: String,
    pub content_type: ChartSheetChartEmbeddedPackageContentType,
    /// Preserved without opening, parsing, activation, macro execution, or external access.
    pub data: Vec<u8>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ChartSheetChartOutboundResource {
    Image(ChartSheetChartImageResource),
    ThemeOverride(ChartSheetChartThemeOverrideResource),
    EmbeddedPackage(ChartSheetChartEmbeddedPackageResource),
}
impl ChartSheetChartOutboundResource {
    fn relationship_id(&self) -> &str {
        match self {
            Self::Image(value) => &value.relationship_id,
            Self::ThemeOverride(value) => &value.relationship_id,
            Self::EmbeddedPackage(value) => &value.relationship_id,
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ChartSheetChartResourceKind {
    Classic,
    Extended {
        styles: Vec<ChartSheetChartCompanionResource>,
        color_styles: Vec<ChartSheetChartCompanionResource>,
        user_shapes: Option<ChartSheetChartUserShapesResource>,
        outbound_resources: Vec<ChartSheetChartOutboundResource>,
    },
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChartSheetDrawingResource {
    pub part_name: String,
    pub content_type: String,
    /// Preserved without rendering or interpreting drawing actions.
    pub data: Vec<u8>,
    pub charts: Vec<ChartSheetChartResource>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChartSheetEntry {
    pub name: String,
    pub sheet_id: u32,
    pub state: ChartSheetState,
    pub workbook_relationship_id: String,
    pub part_name: String,
}

#[derive(Debug, Clone, PartialEq)]
pub struct ChartSheetPackage {
    pub entry: ChartSheetEntry,
    pub chartsheet: ChartSheet,
    pub drawing: ChartSheetDrawingResource,
    pub legacy_drawing: Option<ChartSheetVmlDrawingResource>,
    pub legacy_header_footer_drawing: Option<ChartSheetVmlDrawingResource>,
    pub background_picture: Option<ChartSheetBackgroundPicture>,
    pub printer_settings: Option<ChartSheetPrinterSettings>,
    /// Metadata for otherwise-unknown relationships referenced by extension markup.
    pub extension_relationships: Vec<ChartSheetExtensionRelationship>,
}

#[derive(Clone)]
struct Attribute {
    namespace: String,
    name: String,
    value: String,
}
#[derive(Clone)]
struct Node {
    namespace: String,
    name: String,
    attributes: Vec<Attribute>,
    children: Vec<Node>,
    text: String,
    content: Vec<NodeContent>,
}
#[derive(Clone)]
enum NodeContent {
    Text(String),
    Child(usize),
}

/// Parses the selected, bounded core of a complete Chartsheet part.
pub fn parse_chartsheet(xml: &[u8]) -> Result<(ChartSheetConformance, ChartSheet)> {
    let root = parse_document(xml, MAX_XML_BYTES)?;
    let conformance = root_conformance(&root, "chartsheet")?;
    whitespace(&root)?;
    no_attributes(&root, &[])?;
    validate_root_order(&root)?;
    let properties = one_child(&root, conformance.sml(), "sheetPr")?
        .map(parse_properties)
        .transpose()?;
    let views = parse_views(required_child(&root, conformance.sml(), "sheetViews")?)?;
    let protection = one_child(&root, conformance.sml(), "sheetProtection")?
        .map(parse_protection)
        .transpose()?;
    let custom_views = one_child(&root, conformance.sml(), "customSheetViews")?
        .map(|node| parse_custom_views(node, conformance))
        .transpose()?;
    let margins = one_child(&root, conformance.sml(), "pageMargins")?
        .map(parse_margins)
        .transpose()?;
    let page_setup = one_child(&root, conformance.sml(), "pageSetup")?
        .map(|node| parse_page_setup(node, conformance))
        .transpose()?;
    let header_footer = one_child(&root, conformance.sml(), "headerFooter")?
        .map(parse_header_footer)
        .transpose()?;
    let drawing = required_child(&root, conformance.sml(), "drawing")?;
    leaf(drawing, "chartsheet drawing")?;
    let drawing_relationship_id = required(drawing, conformance.rel(), "id")?.to_owned();
    no_attributes(drawing, &[(conformance.rel(), "id")])?;
    let legacy_drawing_relationship_id =
        parse_relationship_leaf(&root, conformance, "legacyDrawing")?;
    let legacy_header_footer_drawing_relationship_id =
        parse_relationship_leaf(&root, conformance, "legacyDrawingHF")?;
    let background_picture_relationship_id = one_child(&root, conformance.sml(), "picture")?
        .map(|picture| {
            leaf(picture, "chartsheet picture")?;
            no_attributes(picture, &[(conformance.rel(), "id")])?;
            Ok::<_, OoxmlError>(required(picture, conformance.rel(), "id")?.to_owned())
        })
        .transpose()?;
    let web_publish_items = one_child(&root, conformance.sml(), "webPublishItems")?
        .map(|node| parse_web_publish_items(node, conformance))
        .transpose()?;
    let extension_list = one_child(&root, conformance.sml(), "extLst")?
        .map(|node| parse_extension_list(node, conformance))
        .transpose()?;
    let value = ChartSheet {
        properties,
        views,
        protection,
        custom_views,
        margins,
        page_setup,
        header_footer,
        drawing_relationship_id,
        legacy_drawing_relationship_id,
        legacy_header_footer_drawing_relationship_id,
        background_picture_relationship_id,
        web_publish_items,
        extension_list,
    };
    validate_chartsheet(&value)?;
    Ok((conformance, value))
}

fn validate_root_order(root: &Node) -> Result<()> {
    let mut last = 0u8;
    for child in &root.children {
        if child.namespace != root.namespace {
            return Err(invalid(format!(
                "unsupported chartsheet child namespace for '{}'",
                child.name
            )));
        }
        let order = match child.name.as_str() {
            "sheetPr" => 1,
            "sheetViews" => 2,
            "sheetProtection" => 3,
            "customSheetViews" => 4,
            "pageMargins" => 5,
            "pageSetup" => 6,
            "headerFooter" => 7,
            "drawing" => 8,
            "legacyDrawing" => 9,
            "legacyDrawingHF" => 10,
            "picture" => 11,
            "webPublishItems" => 12,
            "extLst" => 13,
            name => return Err(invalid(format!("unsupported chartsheet child '{name}'"))),
        };
        if order <= last {
            return Err(invalid(
                "chartsheet children are duplicated or out of schema order",
            ));
        }
        last = order;
    }
    Ok(())
}

fn parse_relationship_leaf(
    root: &Node,
    conformance: ChartSheetConformance,
    name: &str,
) -> Result<Option<String>> {
    one_child(root, conformance.sml(), name)?
        .map(|node| {
            leaf(node, name)?;
            no_attributes(node, &[(conformance.rel(), "id")])?;
            Ok(required(node, conformance.rel(), "id")?.to_owned())
        })
        .transpose()
}

fn parse_web_publish_items(
    node: &Node,
    conformance: ChartSheetConformance,
) -> Result<ChartSheetWebPublishItems> {
    whitespace(node)?;
    no_attributes(node, &[("", "count")])?;
    if node.children.is_empty() {
        return Err(invalid(
            "webPublishItems requires at least one webPublishItem",
        ));
    }
    if node.children.len() > MAX_WEB_PUBLISH_ITEMS {
        return Err(limit("web publish item count"));
    }
    let count = u32_optional(node, "count")?;
    let mut items = Vec::with_capacity(node.children.len());
    for child in &node.children {
        if child.namespace != conformance.sml() || child.name != "webPublishItem" {
            return Err(invalid("webPublishItems contains an unsupported child"));
        }
        leaf(child, "web publish item")?;
        whitespace(child)?;
        no_attributes(
            child,
            &[
                ("", "id"),
                ("", "divId"),
                ("", "sourceType"),
                ("", "sourceRef"),
                ("", "sourceObject"),
                ("", "destinationFile"),
                ("", "title"),
                ("", "autoRepublish"),
            ],
        )?;
        items.push(ChartSheetWebPublishItem {
            id: required(child, "", "id")?
                .parse()
                .map_err(|_| invalid("invalid web publish item id"))?,
            div_id: required(child, "", "divId")?.to_owned(),
            source_type: ChartSheetWebSourceType::parse(required(child, "", "sourceType")?)?,
            source_ref: optional(child, "", "sourceRef").map(str::to_owned),
            source_object: optional(child, "", "sourceObject").map(str::to_owned),
            destination_file: required(child, "", "destinationFile")?.to_owned(),
            title: optional(child, "", "title").map(str::to_owned),
            auto_republish: bool_optional(child, "autoRepublish")?,
        });
    }
    let value = ChartSheetWebPublishItems { count, items };
    validate_web_publish_items(&value)?;
    Ok(value)
}

fn parse_extension_list(
    node: &Node,
    conformance: ChartSheetConformance,
) -> Result<ChartSheetExtensionList> {
    whitespace(node)?;
    no_attributes(node, &[])?;
    if node.children.is_empty() {
        return Err(invalid("extLst requires at least one ext"));
    }
    if node.children.len() > MAX_EXTENSIONS {
        return Err(limit("extension count"));
    }
    let mut extensions = Vec::with_capacity(node.children.len());
    let mut total = 0usize;
    for child in &node.children {
        if child.namespace != conformance.sml() || child.name != "ext" {
            return Err(invalid("extLst contains an unsupported child"));
        }
        whitespace(child)?;
        no_attributes(child, &[("", "uri")])?;
        let uri = required(child, "", "uri")?.to_owned();
        validate_extension_uri(&uri)?;
        if child.children.len() != 1 {
            return Err(invalid("ext requires exactly one wildcard child"));
        }
        let payload_xml = canonical_extension_payload_node(&child.children[0], conformance)?;
        total = total
            .checked_add(payload_xml.len())
            .ok_or_else(|| limit("retained extension bytes"))?;
        if total > MAX_RETAINED_EXTENSION_BYTES {
            return Err(limit("retained extension bytes"));
        }
        extensions.push(ChartSheetExtension { uri, payload_xml });
    }
    Ok(ChartSheetExtensionList { extensions })
}

fn parse_properties(node: &Node) -> Result<ChartSheetProperties> {
    whitespace(node)?;
    no_attributes(node, &[("", "published"), ("", "codeName")])?;
    let published = optional(node, "", "published")
        .map(|v| parse_bool(v, "published"))
        .transpose()?;
    let code_name = optional(node, "", "codeName").map(str::to_owned);
    let tab_color = one_child_any_core(node, "tabColor")?
        .map(parse_color)
        .transpose()?;
    if node.children.len() > usize::from(tab_color.is_some()) {
        return Err(invalid("sheetPr contains unsupported children"));
    }
    Ok(ChartSheetProperties {
        published,
        code_name,
        tab_color,
    })
}

fn parse_color(node: &Node) -> Result<ChartSheetColor> {
    leaf(node, "tab color")?;
    no_attributes(
        node,
        &[
            ("", "auto"),
            ("", "indexed"),
            ("", "rgb"),
            ("", "theme"),
            ("", "tint"),
        ],
    )?;
    Ok(ChartSheetColor {
        automatic: bool_optional(node, "auto")?,
        indexed: u32_optional(node, "indexed")?,
        rgb: optional(node, "", "rgb").map(str::to_owned),
        theme: u32_optional(node, "theme")?,
        tint: optional(node, "", "tint")
            .map(|v| v.parse().map_err(|_| invalid("invalid tab color tint")))
            .transpose()?,
    })
}

fn parse_views(node: &Node) -> Result<Vec<ChartSheetView>> {
    whitespace(node)?;
    no_attributes(node, &[])?;
    if node.children.is_empty() {
        return Err(invalid("sheetViews requires at least one sheetView"));
    }
    if node.children.len() > MAX_VIEWS {
        return Err(limit("view count"));
    }
    let mut views = Vec::new();
    for child in &node.children {
        if child.name != "sheetView" || !is_core(&child.namespace) {
            return Err(invalid("sheetViews contains an unsupported child"));
        }
        leaf(child, "chartsheet view")?;
        no_attributes(
            child,
            &[
                ("", "tabSelected"),
                ("", "zoomScale"),
                ("", "workbookViewId"),
                ("", "zoomToFit"),
            ],
        )?;
        views.push(ChartSheetView {
            tab_selected: bool_optional(child, "tabSelected")?,
            zoom_scale: u32_optional(child, "zoomScale")?,
            workbook_view_id: required(child, "", "workbookViewId")?
                .parse()
                .map_err(|_| invalid("invalid workbookViewId"))?,
            zoom_to_fit: bool_optional(child, "zoomToFit")?,
        });
    }
    Ok(views)
}

fn parse_protection(node: &Node) -> Result<ChartSheetProtection> {
    leaf(node, "chartsheet protection")?;
    no_attributes(node, &[("", "password"), ("", "content"), ("", "objects")])?;
    Ok(ChartSheetProtection {
        password_hash: optional(node, "", "password").map(str::to_owned),
        content: bool_optional(node, "content")?,
        objects: bool_optional(node, "objects")?,
    })
}

fn parse_custom_views(
    node: &Node,
    conformance: ChartSheetConformance,
) -> Result<Vec<ChartSheetCustomView>> {
    whitespace(node)?;
    no_attributes(node, &[])?;
    if node.children.is_empty() {
        return Err(invalid(
            "customSheetViews requires at least one customSheetView",
        ));
    }
    if node.children.len() > MAX_CUSTOM_VIEWS {
        return Err(limit("custom view count"));
    }
    let mut values = Vec::with_capacity(node.children.len());
    for child in &node.children {
        if child.namespace != conformance.sml() || child.name != "customSheetView" {
            return Err(invalid("customSheetViews contains an unsupported child"));
        }
        leaf(child, "custom chartsheet view")?;
        no_attributes(
            child,
            &[
                ("", "guid"),
                ("", "scale"),
                ("", "state"),
                ("", "zoomToFit"),
            ],
        )?;
        values.push(ChartSheetCustomView {
            guid: required(child, "", "guid")?.to_owned(),
            scale: u32_optional(child, "scale")?,
            state: optional(child, "", "state").map(parse_state).transpose()?,
            zoom_to_fit: bool_optional(child, "zoomToFit")?,
        });
    }
    Ok(values)
}

fn parse_margins(node: &Node) -> Result<ChartSheetMargins> {
    leaf(node, "chartsheet margins")?;
    no_attributes(
        node,
        &[
            ("", "left"),
            ("", "right"),
            ("", "top"),
            ("", "bottom"),
            ("", "header"),
            ("", "footer"),
        ],
    )?;
    let number = |name| {
        required(node, "", name)?
            .parse()
            .map_err(|_| invalid(format!("invalid {name} page margin")))
    };
    Ok(ChartSheetMargins {
        left: number("left")?,
        right: number("right")?,
        top: number("top")?,
        bottom: number("bottom")?,
        header: number("header")?,
        footer: number("footer")?,
    })
}

fn parse_page_setup(
    node: &Node,
    conformance: ChartSheetConformance,
) -> Result<ChartSheetPageSetup> {
    leaf(node, "chartsheet page setup")?;
    no_attributes(
        node,
        &[
            ("", "paperSize"),
            ("", "firstPageNumber"),
            ("", "orientation"),
            ("", "usePrinterDefaults"),
            ("", "blackAndWhite"),
            ("", "draft"),
            ("", "useFirstPageNumber"),
            ("", "horizontalDpi"),
            ("", "verticalDpi"),
            ("", "copies"),
            (conformance.rel(), "id"),
        ],
    )?;
    Ok(ChartSheetPageSetup {
        paper_size: u32_optional(node, "paperSize")?,
        first_page_number: u32_optional(node, "firstPageNumber")?,
        orientation: optional(node, "", "orientation")
            .map(parse_orientation)
            .transpose()?,
        use_printer_defaults: bool_optional(node, "usePrinterDefaults")?,
        black_and_white: bool_optional(node, "blackAndWhite")?,
        draft: bool_optional(node, "draft")?,
        use_first_page_number: bool_optional(node, "useFirstPageNumber")?,
        horizontal_dpi: u32_optional(node, "horizontalDpi")?,
        vertical_dpi: u32_optional(node, "verticalDpi")?,
        copies: u32_optional(node, "copies")?,
        printer_settings_relationship_id: optional(node, conformance.rel(), "id")
            .map(str::to_owned),
    })
}

fn parse_header_footer(node: &Node) -> Result<ChartSheetHeaderFooter> {
    whitespace(node)?;
    no_attributes(
        node,
        &[
            ("", "differentOddEven"),
            ("", "differentFirst"),
            ("", "scaleWithDoc"),
            ("", "alignWithMargins"),
        ],
    )?;
    let mut value = ChartSheetHeaderFooter {
        different_odd_even: bool_optional(node, "differentOddEven")?,
        different_first: bool_optional(node, "differentFirst")?,
        scale_with_document: bool_optional(node, "scaleWithDoc")?,
        align_with_margins: bool_optional(node, "alignWithMargins")?,
        ..Default::default()
    };
    let mut last = 0u8;
    for child in &node.children {
        if !is_core(&child.namespace) {
            return Err(invalid("headerFooter has a foreign child"));
        }
        leaf(child, "header/footer text")?;
        no_attributes(child, &[])?;
        let (order, target) = match child.name.as_str() {
            "oddHeader" => (1, &mut value.odd_header),
            "oddFooter" => (2, &mut value.odd_footer),
            "evenHeader" => (3, &mut value.even_header),
            "evenFooter" => (4, &mut value.even_footer),
            "firstHeader" => (5, &mut value.first_header),
            "firstFooter" => (6, &mut value.first_footer),
            _ => return Err(invalid("unsupported headerFooter child")),
        };
        if order <= last {
            return Err(invalid(
                "headerFooter children are duplicated or out of schema order",
            ));
        }
        last = order;
        *target = Some(child.text.clone());
    }
    Ok(value)
}

/// Deterministically serializes one complete Chartsheet part.
pub fn write_chartsheet(value: &ChartSheet, conformance: ChartSheetConformance) -> Result<Vec<u8>> {
    validate_chartsheet(value)?;
    let mut out = Vec::new();
    out.extend_from_slice(b"<x:chartsheet xmlns:x=\"");
    escape(&mut out, conformance.sml());
    out.extend_from_slice(b"\" xmlns:r=\"");
    escape(&mut out, conformance.rel());
    out.extend_from_slice(b"\">");
    if let Some(properties) = &value.properties {
        out.extend_from_slice(b"<x:sheetPr");
        bool_attr_opt(&mut out, "published", properties.published);
        attr_opt(&mut out, "codeName", properties.code_name.as_deref());
        if let Some(color) = &properties.tab_color {
            out.push(b'>');
            out.extend_from_slice(b"<x:tabColor");
            bool_attr_opt(&mut out, "auto", color.automatic);
            u32_attr_opt(&mut out, "indexed", color.indexed);
            attr_opt(&mut out, "rgb", color.rgb.as_deref());
            u32_attr_opt(&mut out, "theme", color.theme);
            if let Some(v) = color.tint {
                attr(&mut out, "tint", &v.to_string());
            }
            out.extend_from_slice(b"/></x:sheetPr>");
        } else {
            out.extend_from_slice(b"/>");
        }
    }
    out.extend_from_slice(b"<x:sheetViews>");
    for view in &value.views {
        out.extend_from_slice(b"<x:sheetView");
        bool_attr_opt(&mut out, "tabSelected", view.tab_selected);
        u32_attr_opt(&mut out, "zoomScale", view.zoom_scale);
        attr(
            &mut out,
            "workbookViewId",
            &view.workbook_view_id.to_string(),
        );
        bool_attr_opt(&mut out, "zoomToFit", view.zoom_to_fit);
        out.extend_from_slice(b"/>");
    }
    out.extend_from_slice(b"</x:sheetViews>");
    if let Some(protection) = &value.protection {
        out.extend_from_slice(b"<x:sheetProtection");
        attr_opt(&mut out, "password", protection.password_hash.as_deref());
        bool_attr_opt(&mut out, "content", protection.content);
        bool_attr_opt(&mut out, "objects", protection.objects);
        out.extend_from_slice(b"/>");
    }
    if let Some(custom_views) = &value.custom_views {
        out.extend_from_slice(b"<x:customSheetViews>");
        for view in custom_views {
            out.extend_from_slice(b"<x:customSheetView");
            attr(&mut out, "guid", &view.guid);
            u32_attr_opt(&mut out, "scale", view.scale);
            if let Some(state) = view.state {
                attr(
                    &mut out,
                    "state",
                    match state {
                        ChartSheetState::Visible => "visible",
                        ChartSheetState::Hidden => "hidden",
                        ChartSheetState::VeryHidden => "veryHidden",
                    },
                );
            }
            bool_attr_opt(&mut out, "zoomToFit", view.zoom_to_fit);
            out.extend_from_slice(b"/>");
        }
        out.extend_from_slice(b"</x:customSheetViews>");
    }
    if let Some(m) = value.margins {
        out.extend_from_slice(b"<x:pageMargins");
        for (name, value) in [
            ("left", m.left),
            ("right", m.right),
            ("top", m.top),
            ("bottom", m.bottom),
            ("header", m.header),
            ("footer", m.footer),
        ] {
            attr(&mut out, name, &value.to_string());
        }
        out.extend_from_slice(b"/>");
    }
    if let Some(setup) = &value.page_setup {
        out.extend_from_slice(b"<x:pageSetup");
        u32_attr_opt(&mut out, "paperSize", setup.paper_size);
        u32_attr_opt(&mut out, "firstPageNumber", setup.first_page_number);
        if let Some(v) = setup.orientation {
            attr(
                &mut out,
                "orientation",
                match v {
                    PageOrientation::Default => "default",
                    PageOrientation::Portrait => "portrait",
                    PageOrientation::Landscape => "landscape",
                },
            );
        }
        bool_attr_opt(&mut out, "usePrinterDefaults", setup.use_printer_defaults);
        bool_attr_opt(&mut out, "blackAndWhite", setup.black_and_white);
        bool_attr_opt(&mut out, "draft", setup.draft);
        bool_attr_opt(&mut out, "useFirstPageNumber", setup.use_first_page_number);
        u32_attr_opt(&mut out, "horizontalDpi", setup.horizontal_dpi);
        u32_attr_opt(&mut out, "verticalDpi", setup.vertical_dpi);
        u32_attr_opt(&mut out, "copies", setup.copies);
        attr_opt(
            &mut out,
            "r:id",
            setup.printer_settings_relationship_id.as_deref(),
        );
        out.extend_from_slice(b"/>");
    }
    if let Some(hf) = &value.header_footer {
        out.extend_from_slice(b"<x:headerFooter");
        bool_attr_opt(&mut out, "differentOddEven", hf.different_odd_even);
        bool_attr_opt(&mut out, "differentFirst", hf.different_first);
        bool_attr_opt(&mut out, "scaleWithDoc", hf.scale_with_document);
        bool_attr_opt(&mut out, "alignWithMargins", hf.align_with_margins);
        let children = [
            ("oddHeader", &hf.odd_header),
            ("oddFooter", &hf.odd_footer),
            ("evenHeader", &hf.even_header),
            ("evenFooter", &hf.even_footer),
            ("firstHeader", &hf.first_header),
            ("firstFooter", &hf.first_footer),
        ];
        if children.iter().all(|(_, value)| value.is_none()) {
            out.extend_from_slice(b"/>");
        } else {
            out.push(b'>');
            for (name, value) in children {
                if let Some(value) = value {
                    out.extend_from_slice(b"<x:");
                    out.extend_from_slice(name.as_bytes());
                    out.push(b'>');
                    escape_text(&mut out, value);
                    out.extend_from_slice(b"</x:");
                    out.extend_from_slice(name.as_bytes());
                    out.push(b'>');
                }
            }
            out.extend_from_slice(b"</x:headerFooter>");
        }
    }
    out.extend_from_slice(b"<x:drawing");
    attr(&mut out, "r:id", &value.drawing_relationship_id);
    out.extend_from_slice(b"/>");
    if let Some(id) = &value.legacy_drawing_relationship_id {
        write_relationship_leaf(&mut out, "legacyDrawing", id);
    }
    if let Some(id) = &value.legacy_header_footer_drawing_relationship_id {
        write_relationship_leaf(&mut out, "legacyDrawingHF", id);
    }
    if let Some(id) = &value.background_picture_relationship_id {
        out.extend_from_slice(b"<x:picture");
        attr(&mut out, "r:id", id);
        out.extend_from_slice(b"/>");
    }
    if let Some(collection) = &value.web_publish_items {
        out.extend_from_slice(b"<x:webPublishItems");
        u32_attr_opt(&mut out, "count", collection.count);
        out.push(b'>');
        for item in &collection.items {
            out.extend_from_slice(b"<x:webPublishItem");
            attr(&mut out, "id", &item.id.to_string());
            attr(&mut out, "divId", &item.div_id);
            attr(&mut out, "sourceType", item.source_type.as_str());
            attr_opt(&mut out, "sourceRef", item.source_ref.as_deref());
            attr_opt(&mut out, "sourceObject", item.source_object.as_deref());
            attr(&mut out, "destinationFile", &item.destination_file);
            attr_opt(&mut out, "title", item.title.as_deref());
            bool_attr_opt(&mut out, "autoRepublish", item.auto_republish);
            out.extend_from_slice(b"/>");
        }
        out.extend_from_slice(b"</x:webPublishItems>");
    }
    if let Some(list) = &value.extension_list {
        out.extend_from_slice(b"<x:extLst>");
        for extension in &list.extensions {
            out.extend_from_slice(b"<x:ext");
            attr(&mut out, "uri", &extension.uri);
            out.push(b'>');
            let payload = canonical_extension_payload(&extension.payload_xml, conformance)?;
            out.extend_from_slice(&payload);
            out.extend_from_slice(b"</x:ext>");
        }
        out.extend_from_slice(b"</x:extLst>");
    }
    out.extend_from_slice(b"</x:chartsheet>");
    if out.len() > MAX_XML_BYTES {
        return Err(limit("serialized XML bytes"));
    }
    Ok(out)
}

fn write_relationship_leaf(out: &mut Vec<u8>, name: &str, id: &str) {
    out.extend_from_slice(b"<x:");
    out.extend_from_slice(name.as_bytes());
    attr(out, "r:id", id);
    out.extend_from_slice(b"/>");
}

/// Loads one workbook-referenced chartsheet and validates its bounded leaf graph.
pub fn load_chartsheet(
    package: &OpcPackage,
    workbook_name: &PackURI,
    workbook_relationship_id: &str,
) -> Result<ChartSheetPackage> {
    if package.rels().iter().any(|rel| {
        matches!(rel.reltype(), CHARTSHEET_REL | STRICT_CHARTSHEET_REL)
            || is_printer_relationship(rel.reltype())
    }) {
        return Err(invalid(
            "package root cannot source a chartsheet or Printer Settings relationship",
        ));
    }
    let workbook = package.get_part(workbook_name)?;
    require_workbook(workbook)?;
    let workbook_root = parse_document(workbook.blob(), MAX_XML_BYTES)?;
    let conformance = root_conformance(&workbook_root, "workbook")?;
    let relationship = internal_relationship(
        workbook,
        workbook_relationship_id,
        conformance.chartsheet_rel(),
    )?;
    let chartsheet_name = relationship.target_partname()?;
    if !chartsheet_name.as_str().starts_with("/xl/chartsheets/") {
        return Err(invalid("chartsheet target is outside /xl/chartsheets"));
    }
    let entry = workbook_entry(
        &workbook_root,
        conformance,
        workbook_relationship_id,
        chartsheet_name.to_string(),
    )?;
    let chartsheet_part = package.get_part(&chartsheet_name)?;
    require_content_type(chartsheet_part, CHARTSHEET_CT, "chartsheet")?;
    let (part_conformance, chartsheet) = parse_chartsheet(chartsheet_part.blob())?;
    if part_conformance != conformance {
        return Err(invalid("workbook and chartsheet conformance differ"));
    }
    let drawing_rel = internal_relationship(
        chartsheet_part,
        &chartsheet.drawing_relationship_id,
        conformance.drawing_rel(),
    )?;
    let drawing_name = drawing_rel.target_partname()?;
    if !drawing_name.as_str().starts_with("/xl/drawings/") {
        return Err(invalid("chartsheet drawing is outside /xl/drawings"));
    }
    let legacy_drawing = chartsheet
        .legacy_drawing_relationship_id
        .as_deref()
        .map(|id| load_vml_resource(package, chartsheet_part, id, conformance))
        .transpose()?;
    let legacy_header_footer_drawing = chartsheet
        .legacy_header_footer_drawing_relationship_id
        .as_deref()
        .map(|id| load_vml_resource(package, chartsheet_part, id, conformance))
        .transpose()?;
    let background_picture = if let Some(id) = &chartsheet.background_picture_relationship_id {
        let rel = internal_relationship(chartsheet_part, id, conformance.image_rel())?;
        let name = rel.target_partname()?;
        if !name.as_str().starts_with("/xl/media/") {
            return Err(invalid("chartsheet background image is outside /xl/media"));
        }
        let part = package.get_part(&name)?;
        let content_type = ChartSheetImageContentType::parse(part.content_type())?;
        if part.blob().len() > MAX_BACKGROUND_IMAGE_BYTES {
            return Err(limit("background image bytes"));
        }
        if !part.rels().is_empty() {
            return Err(invalid(
                "chartsheet background image must be a relationship-free leaf",
            ));
        }
        Some(ChartSheetBackgroundPicture {
            relationship_id: id.clone(),
            part_name: name.to_string(),
            content_type,
            data: part.blob().to_vec(),
        })
    } else {
        None
    };
    let printer_settings = if let Some(id) = chartsheet
        .page_setup
        .as_ref()
        .and_then(|setup| setup.printer_settings_relationship_id.as_ref())
    {
        let rel = internal_relationship(chartsheet_part, id, conformance.printer_rel())?;
        let name = rel.target_partname()?;
        validate_printer_settings_uri(&name)?;
        let part = package.get_part(&name)?;
        require_content_type(part, PRINTER_CT, "Printer Settings")?;
        validate_settings_bytes(part.blob())?;
        if !part.rels().is_empty() {
            return Err(invalid(
                "chartsheet Printer Settings must be a relationship-free leaf",
            ));
        }
        Some(ChartSheetPrinterSettings {
            relationship_id: id.clone(),
            resource: PrinterSettingsResource {
                part_name: name.to_string(),
                data: part.blob().to_vec(),
            },
        })
    } else {
        None
    };
    let known_relationships = known_chartsheet_relationship_ids(&chartsheet);
    let extension_ids = extension_relationship_ids(&chartsheet, conformance)?;
    let mut extension_relationships = Vec::new();
    for id in extension_ids.difference(&known_relationships) {
        if extension_relationships.len() >= MAX_EXTENSION_RELATIONSHIPS {
            return Err(limit("extension relationship count"));
        }
        let relationship = chartsheet_part
            .rels()
            .get(id)
            .ok_or_else(|| invalid(format!("missing extension relationship '{id}'")))?;
        validate_extension_relationship_string(relationship.reltype(), "type")?;
        let target = if relationship.is_external() {
            validate_extension_relationship_string(relationship.target_ref(), "target")?;
            ChartSheetExtensionRelationshipTarget::External {
                target: relationship.target_ref().to_owned(),
            }
        } else {
            let name = relationship.target_partname()?;
            ChartSheetExtensionRelationshipTarget::Internal {
                part_name: name.to_string(),
            }
        };
        extension_relationships.push(ChartSheetExtensionRelationship {
            relationship_id: id.clone(),
            relationship_type: relationship.reltype().to_owned(),
            target,
        });
    }
    let expected_relationships = known_relationships.len() + extension_relationships.len();
    if chartsheet_part.rels().iter().count() != expected_relationships {
        return Err(invalid(
            "bounded chartsheet has unsupported or unreferenced relationships",
        ));
    }
    let drawing_part = package.get_part(&drawing_name)?;
    require_content_type(drawing_part, DRAWING_CT, "drawing")?;
    if drawing_part.blob().len() > MAX_DRAWING_BYTES {
        return Err(limit("drawing bytes"));
    }
    let chart_references = drawing_chart_references(drawing_part.blob(), conformance)?;
    if drawing_part.rels().iter().count() != chart_references.len() {
        return Err(invalid(
            "bounded chartsheet drawing has unsupported or unreferenced relationships",
        ));
    }
    let mut charts = Vec::new();
    let mut total = drawing_part.blob().len();
    for vml in [&legacy_drawing, &legacy_header_footer_drawing]
        .into_iter()
        .flatten()
    {
        add_resource(
            &mut total,
            vml.data.len(),
            MAX_VML_DRAWING_BYTES,
            "VML drawing bytes",
        )?;
    }
    if let Some(picture) = &background_picture {
        add_resource(
            &mut total,
            picture.data.len(),
            MAX_BACKGROUND_IMAGE_BYTES,
            "background image bytes",
        )?;
    }
    if let Some(settings) = &printer_settings {
        add_resource(
            &mut total,
            settings.resource.data.len(),
            MAX_SETTINGS_BYTES,
            "Printer Settings bytes",
        )?;
    }
    for reference in chart_references {
        charts.push(load_chart_resource(
            package,
            drawing_part,
            &reference,
            conformance,
            &mut total,
        )?);
    }
    Ok(ChartSheetPackage {
        entry,
        chartsheet,
        drawing: ChartSheetDrawingResource {
            part_name: drawing_name.to_string(),
            content_type: drawing_part.content_type().to_owned(),
            data: drawing_part.blob().to_vec(),
            charts,
        },
        legacy_drawing,
        legacy_header_footer_drawing,
        background_picture,
        printer_settings,
        extension_relationships,
    })
}

fn load_vml_resource(
    package: &OpcPackage,
    chartsheet_part: &dyn Part,
    id: &str,
    conformance: ChartSheetConformance,
) -> Result<ChartSheetVmlDrawingResource> {
    let rel = internal_relationship(chartsheet_part, id, conformance.vml_drawing_rel())?;
    let name = rel.target_partname()?;
    if !name.as_str().starts_with("/xl/drawings/") || !name.as_str().ends_with(".vml") {
        return Err(invalid(
            "chartsheet VML drawing target is outside /xl/drawings or lacks .vml suffix",
        ));
    }
    let part = package.get_part(&name)?;
    require_content_type(part, VML_DRAWING_CT, "VML drawing")?;
    if part.blob().len() > MAX_VML_DRAWING_BYTES {
        return Err(limit("VML drawing bytes"));
    }
    if !part.rels().is_empty() {
        return Err(invalid(
            "chartsheet VML drawing must be a relationship-free leaf",
        ));
    }
    Ok(ChartSheetVmlDrawingResource {
        relationship_id: id.to_owned(),
        part_name: name.to_string(),
        content_type: part.content_type().to_owned(),
        data: part.blob().to_vec(),
    })
}

fn load_chart_resource(
    package: &OpcPackage,
    drawing_part: &dyn Part,
    reference: &DrawingChartReference,
    conformance: ChartSheetConformance,
    total: &mut usize,
) -> Result<ChartSheetChartResource> {
    let relationship_type = match reference.kind {
        DrawingChartKind::Classic => conformance.chart_rel(),
        DrawingChartKind::Extended => CHART_EX_REL,
    };
    let relationship =
        internal_relationship(drawing_part, &reference.relationship_id, relationship_type)?;
    let name = relationship.target_partname()?;
    if !name.as_str().starts_with("/xl/charts/") || !name.as_str().ends_with(".xml") {
        return Err(invalid(
            "chart target is outside /xl/charts or lacks .xml suffix",
        ));
    }
    let part = package.get_part(&name)?;
    let kind = match reference.kind {
        DrawingChartKind::Classic => {
            require_content_type(part, CHART_CT, "chart")?;
            validate_chart_xml(part.blob(), conformance)?;
            if !part.rels().is_empty() {
                return Err(invalid(
                    "bounded classic chart must be a relationship-free leaf",
                ));
            }
            add_resource(total, part.blob().len(), MAX_CHART_BYTES, "chart bytes")?;
            ChartSheetChartResourceKind::Classic
        },
        DrawingChartKind::Extended => {
            require_content_type(part, CHART_EX_CT, "chartEx")?;
            let references = validate_chart_ex_relationships(part.blob(), conformance)?;
            add_resource(
                total,
                part.blob().len(),
                MAX_CHART_EX_BYTES,
                "chartEx bytes",
            )?;
            let (mut styles, mut color_styles, mut user_shapes, mut outbound_resources) =
                (Vec::new(), Vec::new(), None, Vec::new());
            let (mut theme_seen, mut package_seen) = (false, false);
            for relationship in part.rels().iter() {
                if relationship.reltype() == conformance.chart_user_shapes_rel() {
                    if user_shapes.is_some() {
                        return Err(invalid(
                            "chartEx has multiple chartUserShapes relationships",
                        ));
                    }
                    if relationship.is_external() {
                        return Err(invalid("external chartUserShapes relationship is rejected"));
                    }
                    user_shapes = Some(load_chart_user_shapes_resource(
                        package,
                        relationship,
                        conformance,
                        total,
                    )?);
                    continue;
                }
                if relationship.reltype() == conformance.image_rel() {
                    if outbound_resources
                        .iter()
                        .filter(|value| matches!(value, ChartSheetChartOutboundResource::Image(_)))
                        .count()
                        >= MAX_CHART_DIRECT_IMAGES
                    {
                        return Err(limit("chartEx direct image count"));
                    }
                    outbound_resources.push(ChartSheetChartOutboundResource::Image(
                        load_chart_image_resource(
                            package,
                            relationship,
                            total,
                            "chartEx direct image",
                        )?,
                    ));
                    continue;
                }
                if relationship.reltype() == conformance.theme_override_rel() {
                    if theme_seen {
                        return Err(invalid("chartEx has multiple themeOverride relationships"));
                    }
                    theme_seen = true;
                    outbound_resources.push(ChartSheetChartOutboundResource::ThemeOverride(
                        load_chart_theme_override_resource(
                            package,
                            relationship,
                            conformance,
                            total,
                        )?,
                    ));
                    continue;
                }
                if relationship.reltype() == conformance.package_rel() {
                    if package_seen {
                        return Err(invalid(
                            "chartEx has multiple embedded package relationships",
                        ));
                    }
                    package_seen = true;
                    outbound_resources.push(ChartSheetChartOutboundResource::EmbeddedPackage(
                        load_chart_embedded_package_resource(package, relationship, total)?,
                    ));
                    continue;
                }
                let (target, root, max_bytes, collection) = match relationship.reltype() {
                    CHART_STYLE_REL => (
                        CHART_STYLE_CT,
                        "chartStyle",
                        MAX_CHART_STYLE_BYTES,
                        &mut styles,
                    ),
                    CHART_COLOR_STYLE_REL => (
                        CHART_COLOR_STYLE_CT,
                        "colorStyle",
                        MAX_CHART_COLOR_STYLE_BYTES,
                        &mut color_styles,
                    ),
                    _ => {
                        return Err(invalid(
                            "chartEx has an unsupported or active outbound relationship",
                        ));
                    },
                };
                if collection.len() >= MAX_CHART_STYLE_PARTS {
                    return Err(limit("chart companion count"));
                }
                if relationship.is_external() {
                    return Err(invalid(
                        "external chartEx companion relationship is rejected",
                    ));
                }
                let companion_name = relationship.target_partname()?;
                if !companion_name.as_str().starts_with("/xl/charts/")
                    || !companion_name.as_str().ends_with(".xml")
                {
                    return Err(invalid(
                        "chart companion target is outside /xl/charts or lacks .xml suffix",
                    ));
                }
                let companion = package.get_part(&companion_name)?;
                require_content_type(companion, target, "chart companion")?;
                validate_chart_companion_xml(companion.blob(), root, max_bytes)?;
                if !companion.rels().is_empty() {
                    return Err(invalid("chart companion must be a relationship-free leaf"));
                }
                add_resource(
                    total,
                    companion.blob().len(),
                    max_bytes,
                    "chart companion bytes",
                )?;
                collection.push(ChartSheetChartCompanionResource {
                    relationship_id: relationship.r_id().to_owned(),
                    part_name: companion_name.to_string(),
                    content_type: companion.content_type().to_owned(),
                    data: companion.blob().to_vec(),
                });
            }
            let image_ids = outbound_resources
                .iter()
                .filter_map(|value| match value {
                    ChartSheetChartOutboundResource::Image(image) => {
                        Some(image.relationship_id.clone())
                    },
                    _ => None,
                })
                .collect::<BTreeSet<_>>();
            let package_id = outbound_resources.iter().find_map(|value| match value {
                ChartSheetChartOutboundResource::EmbeddedPackage(package) => {
                    Some(package.relationship_id.clone())
                },
                _ => None,
            });
            if image_ids != references.images || package_id != references.package {
                return Err(invalid(
                    "chartEx XML relationship references do not close over direct images and embedded package",
                ));
            }
            styles.sort_by(|left, right| left.relationship_id.cmp(&right.relationship_id));
            color_styles.sort_by(|left, right| left.relationship_id.cmp(&right.relationship_id));
            outbound_resources
                .sort_by(|left, right| left.relationship_id().cmp(right.relationship_id()));
            ChartSheetChartResourceKind::Extended {
                styles,
                color_styles,
                user_shapes,
                outbound_resources,
            }
        },
    };
    Ok(ChartSheetChartResource {
        relationship_id: reference.relationship_id.clone(),
        part_name: name.to_string(),
        content_type: part.content_type().to_owned(),
        data: part.blob().to_vec(),
        kind,
    })
}

fn load_chart_user_shapes_resource(
    package: &OpcPackage,
    relationship: &litchi_opc::Relationship,
    conformance: ChartSheetConformance,
    total: &mut usize,
) -> Result<ChartSheetChartUserShapesResource> {
    let name = relationship.target_partname()?;
    if !name.as_str().starts_with("/xl/drawings/") || !name.as_str().ends_with(".xml") {
        return Err(invalid(
            "chartUserShapes target is outside /xl/drawings or lacks .xml suffix",
        ));
    }
    let part = package.get_part(&name)?;
    require_content_type(part, CHART_USER_SHAPES_CT, "chartUserShapes")?;
    let referenced = validate_chart_user_shapes_xml(part.blob(), conformance)?;
    add_resource(
        total,
        part.blob().len(),
        MAX_CHART_USER_SHAPES_BYTES,
        "chartUserShapes bytes",
    )?;
    if referenced.len() > MAX_CHART_USER_SHAPE_IMAGES {
        return Err(limit("chart user-shape image count"));
    }
    if part.rels().iter().count() != referenced.len() {
        return Err(invalid(
            "chartUserShapes image relationships are missing or orphaned",
        ));
    }
    let mut images = Vec::with_capacity(referenced.len());
    for id in referenced {
        let image_relationship = internal_relationship(part, &id, conformance.image_rel())?;
        let image_name = image_relationship.target_partname()?;
        if !image_name.as_str().starts_with("/xl/media/") {
            return Err(invalid("chart user-shape image is outside /xl/media"));
        }
        let image = package.get_part(&image_name)?;
        let content_type = ChartSheetChartUserShapeImageContentType::parse(image.content_type())?;
        if !content_type.validates_part_name(image_name.as_str()) {
            return Err(invalid(
                "chart user-shape image suffix does not match its content type",
            ));
        }
        if !image.rels().is_empty() {
            return Err(invalid(
                "chart user-shape image must be a relationship-free leaf",
            ));
        }
        add_resource(
            total,
            image.blob().len(),
            MAX_CHART_USER_SHAPE_IMAGE_BYTES,
            "chart user-shape image bytes",
        )?;
        images.push(ChartSheetChartUserShapeImageResource {
            relationship_id: id,
            part_name: image_name.to_string(),
            content_type,
            data: image.blob().to_vec(),
        });
    }
    images.sort_by(|left, right| left.relationship_id.cmp(&right.relationship_id));
    Ok(ChartSheetChartUserShapesResource {
        relationship_id: relationship.r_id().to_owned(),
        part_name: name.to_string(),
        content_type: part.content_type().to_owned(),
        data: part.blob().to_vec(),
        images,
    })
}

fn load_chart_image_resource(
    package: &OpcPackage,
    relationship: &litchi_opc::Relationship,
    total: &mut usize,
    label: &str,
) -> Result<ChartSheetChartImageResource> {
    if relationship.is_external() {
        return Err(invalid(format!(
            "external {label} relationship is rejected"
        )));
    }
    let name = relationship.target_partname()?;
    if !name.as_str().starts_with("/xl/media/") {
        return Err(invalid(format!("{label} is outside /xl/media")));
    }
    let part = package.get_part(&name)?;
    let content_type = ChartSheetChartImageContentType::parse(part.content_type())?;
    if !content_type.validates_part_name(name.as_str()) {
        return Err(invalid(format!(
            "{label} suffix does not match its content type"
        )));
    }
    if !part.rels().is_empty() {
        return Err(invalid(format!("{label} must be a relationship-free leaf")));
    }
    add_resource(
        total,
        part.blob().len(),
        MAX_CHART_USER_SHAPE_IMAGE_BYTES,
        "chart image bytes",
    )?;
    Ok(ChartSheetChartImageResource {
        relationship_id: relationship.r_id().to_owned(),
        part_name: name.to_string(),
        content_type,
        data: part.blob().to_vec(),
    })
}

fn load_chart_theme_override_resource(
    package: &OpcPackage,
    relationship: &litchi_opc::Relationship,
    conformance: ChartSheetConformance,
    total: &mut usize,
) -> Result<ChartSheetChartThemeOverrideResource> {
    if relationship.is_external() {
        return Err(invalid("external themeOverride relationship is rejected"));
    }
    let name = relationship.target_partname()?;
    if !name.as_str().starts_with("/xl/theme/") || !name.as_str().ends_with(".xml") {
        return Err(invalid(
            "themeOverride target is outside /xl/theme or lacks .xml suffix",
        ));
    }
    let part = package.get_part(&name)?;
    require_content_type(part, THEME_OVERRIDE_CT, "themeOverride")?;
    let referenced = validate_theme_override_xml(part.blob(), conformance)?;
    add_resource(
        total,
        part.blob().len(),
        MAX_CHART_THEME_OVERRIDE_BYTES,
        "themeOverride bytes",
    )?;
    if referenced.len() > MAX_CHART_THEME_IMAGES {
        return Err(limit("themeOverride image count"));
    }
    if part.rels().iter().count() != referenced.len() {
        return Err(invalid(
            "themeOverride image relationships are missing or orphaned",
        ));
    }
    let mut images = Vec::with_capacity(referenced.len());
    for id in referenced {
        let image_relationship = internal_relationship(part, &id, conformance.image_rel())?;
        images.push(load_chart_image_resource(
            package,
            image_relationship,
            total,
            "themeOverride image",
        )?);
    }
    images.sort_by(|left, right| left.relationship_id.cmp(&right.relationship_id));
    Ok(ChartSheetChartThemeOverrideResource {
        relationship_id: relationship.r_id().to_owned(),
        part_name: name.to_string(),
        content_type: part.content_type().to_owned(),
        data: part.blob().to_vec(),
        images,
    })
}

fn load_chart_embedded_package_resource(
    package: &OpcPackage,
    relationship: &litchi_opc::Relationship,
    total: &mut usize,
) -> Result<ChartSheetChartEmbeddedPackageResource> {
    if relationship.is_external() {
        return Err(invalid(
            "external chartEx embedded-package relationship is rejected",
        ));
    }
    let name = relationship.target_partname()?;
    if !name.as_str().starts_with("/xl/embeddings/") {
        return Err(invalid(
            "chartEx embedded package is outside /xl/embeddings",
        ));
    }
    let part = package.get_part(&name)?;
    let content_type = ChartSheetChartEmbeddedPackageContentType::parse(part.content_type())?;
    if !content_type.validates_part_name(name.as_str()) {
        return Err(invalid(
            "chartEx embedded-package suffix does not match its content type",
        ));
    }
    if !part.rels().is_empty() {
        return Err(invalid(
            "chartEx embedded package must be a relationship-free opaque leaf",
        ));
    }
    add_resource(
        total,
        part.blob().len(),
        MAX_CHART_EMBEDDED_PACKAGE_BYTES,
        "chartEx embedded-package bytes",
    )?;
    Ok(ChartSheetChartEmbeddedPackageResource {
        relationship_id: relationship.r_id().to_owned(),
        part_name: name.to_string(),
        content_type,
        data: part.blob().to_vec(),
    })
}

/// Adds a preflighted chartsheet package graph and workbook sheet entry.
pub fn store_chartsheet(
    package: &mut OpcPackage,
    workbook_name: &PackURI,
    value: &ChartSheetPackage,
    conformance: ChartSheetConformance,
) -> Result<()> {
    validate_package_value(value, conformance)?;
    let workbook = package.get_part(workbook_name)?;
    require_workbook(workbook)?;
    let workbook_root = parse_document(workbook.blob(), MAX_XML_BYTES)?;
    if root_conformance(&workbook_root, "workbook")? != conformance {
        return Err(invalid("requested conformance does not match workbook"));
    }
    validate_new_entry(&workbook_root, conformance, &value.entry)?;
    if workbook
        .rels()
        .get(&value.entry.workbook_relationship_id)
        .is_some()
    {
        return Err(invalid("workbook relationship ID already exists"));
    }
    let chartsheet_uri = new_uri(package, &value.entry.part_name, "/xl/chartsheets/")?;
    let drawing_uri = new_uri(package, &value.drawing.part_name, "/xl/drawings/")?;
    let legacy_uri = value
        .legacy_drawing
        .as_ref()
        .map(|resource| new_uri(package, &resource.part_name, "/xl/drawings/"))
        .transpose()?;
    let legacy_hf_uri = value
        .legacy_header_footer_drawing
        .as_ref()
        .map(|resource| new_uri(package, &resource.part_name, "/xl/drawings/"))
        .transpose()?;
    let picture_uri = value
        .background_picture
        .as_ref()
        .map(|picture| new_uri(package, &picture.part_name, "/xl/media/"))
        .transpose()?;
    let printer_uri = value
        .printer_settings
        .as_ref()
        .map(|settings| {
            let uri = PackURI::new(&settings.resource.part_name).map_err(OoxmlError::InvalidUri)?;
            validate_printer_settings_uri(&uri)?;
            if package.iter_parts().any(|part| part.partname() == &uri) {
                return Err(invalid(format!("part '{uri}' already exists")));
            }
            Ok(uri)
        })
        .transpose()?;
    let mut chart_uris = BTreeMap::new();
    let mut companion_uris = BTreeMap::new();
    let mut user_shape_uris = BTreeMap::new();
    let mut user_shape_image_uris = BTreeMap::new();
    let mut outbound_uris = BTreeMap::new();
    let mut theme_image_uris = BTreeMap::new();
    for chart in &value.drawing.charts {
        chart_uris.insert(
            chart.relationship_id.clone(),
            new_uri(package, &chart.part_name, "/xl/charts/")?,
        );
        if let ChartSheetChartResourceKind::Extended {
            styles,
            color_styles,
            user_shapes,
            outbound_resources,
        } = &chart.kind
        {
            for companion in styles.iter().chain(color_styles) {
                companion_uris.insert(
                    (
                        chart.relationship_id.clone(),
                        companion.relationship_id.clone(),
                    ),
                    new_uri(package, &companion.part_name, "/xl/charts/")?,
                );
            }
            if let Some(user_shapes) = user_shapes {
                user_shape_uris.insert(
                    chart.relationship_id.clone(),
                    new_uri(package, &user_shapes.part_name, "/xl/drawings/")?,
                );
                for image in &user_shapes.images {
                    user_shape_image_uris.insert(
                        (chart.relationship_id.clone(), image.relationship_id.clone()),
                        new_uri(package, &image.part_name, "/xl/media/")?,
                    );
                }
            }
            for resource in outbound_resources {
                let relationship_id = resource.relationship_id().to_owned();
                let (prefix, part_name) = match resource {
                    ChartSheetChartOutboundResource::Image(image) => {
                        ("/xl/media/", image.part_name.as_str())
                    },
                    ChartSheetChartOutboundResource::ThemeOverride(theme) => {
                        ("/xl/theme/", theme.part_name.as_str())
                    },
                    ChartSheetChartOutboundResource::EmbeddedPackage(embedded) => {
                        ("/xl/embeddings/", embedded.part_name.as_str())
                    },
                };
                outbound_uris.insert(
                    (chart.relationship_id.clone(), relationship_id.clone()),
                    new_uri(package, part_name, prefix)?,
                );
                if let ChartSheetChartOutboundResource::ThemeOverride(theme) = resource {
                    for image in &theme.images {
                        theme_image_uris.insert(
                            (
                                chart.relationship_id.clone(),
                                relationship_id.clone(),
                                image.relationship_id.clone(),
                            ),
                            new_uri(package, &image.part_name, "/xl/media/")?,
                        );
                    }
                }
            }
        }
    }
    let updated_workbook = insert_workbook_entry(workbook.blob(), &value.entry, conformance)?;
    let chartsheet_xml = write_chartsheet(&value.chartsheet, conformance)?;
    package
        .get_part_mut(workbook_name)?
        .set_blob(updated_workbook);
    package.add_part(Box::new(BlobPart::new(
        chartsheet_uri.clone(),
        CHARTSHEET_CT.into(),
        chartsheet_xml,
    )));
    package.add_part(Box::new(BlobPart::new(
        drawing_uri.clone(),
        value.drawing.content_type.clone(),
        value.drawing.data.clone(),
    )));
    if let (Some(resource), Some(uri)) = (&value.legacy_drawing, &legacy_uri) {
        package.add_part(Box::new(BlobPart::new(
            uri.clone(),
            resource.content_type.clone(),
            resource.data.clone(),
        )));
    }
    if let (Some(resource), Some(uri)) = (&value.legacy_header_footer_drawing, &legacy_hf_uri) {
        package.add_part(Box::new(BlobPart::new(
            uri.clone(),
            resource.content_type.clone(),
            resource.data.clone(),
        )));
    }
    if let (Some(picture), Some(uri)) = (&value.background_picture, &picture_uri) {
        package.add_part(Box::new(BlobPart::new(
            uri.clone(),
            picture.content_type.as_str().into(),
            picture.data.clone(),
        )));
    }
    if let (Some(settings), Some(uri)) = (&value.printer_settings, &printer_uri) {
        package.add_part(Box::new(BlobPart::new(
            uri.clone(),
            PRINTER_CT.into(),
            settings.resource.data.clone(),
        )));
    }
    for chart in &value.drawing.charts {
        package.add_part(Box::new(BlobPart::new(
            chart_uris[&chart.relationship_id].clone(),
            chart.content_type.clone(),
            chart.data.clone(),
        )));
        if let ChartSheetChartResourceKind::Extended {
            styles,
            color_styles,
            user_shapes,
            outbound_resources,
        } = &chart.kind
        {
            for companion in styles.iter().chain(color_styles) {
                package.add_part(Box::new(BlobPart::new(
                    companion_uris[&(
                        chart.relationship_id.clone(),
                        companion.relationship_id.clone(),
                    )]
                        .clone(),
                    companion.content_type.clone(),
                    companion.data.clone(),
                )));
            }
            if let Some(user_shapes) = user_shapes {
                package.add_part(Box::new(BlobPart::new(
                    user_shape_uris[&chart.relationship_id].clone(),
                    user_shapes.content_type.clone(),
                    user_shapes.data.clone(),
                )));
                for image in &user_shapes.images {
                    package.add_part(Box::new(BlobPart::new(
                        user_shape_image_uris
                            [&(chart.relationship_id.clone(), image.relationship_id.clone())]
                            .clone(),
                        image.content_type.as_str().into(),
                        image.data.clone(),
                    )));
                }
            }
            for resource in outbound_resources {
                let key = (
                    chart.relationship_id.clone(),
                    resource.relationship_id().to_owned(),
                );
                let uri = outbound_uris[&key].clone();
                match resource {
                    ChartSheetChartOutboundResource::Image(image) => package.add_part(Box::new(
                        BlobPart::new(uri, image.content_type.as_str().into(), image.data.clone()),
                    )),
                    ChartSheetChartOutboundResource::ThemeOverride(theme) => {
                        package.add_part(Box::new(BlobPart::new(
                            uri,
                            theme.content_type.clone(),
                            theme.data.clone(),
                        )));
                        for image in &theme.images {
                            let image_uri = theme_image_uris[&(
                                chart.relationship_id.clone(),
                                theme.relationship_id.clone(),
                                image.relationship_id.clone(),
                            )]
                                .clone();
                            package.add_part(Box::new(BlobPart::new(
                                image_uri,
                                image.content_type.as_str().into(),
                                image.data.clone(),
                            )));
                        }
                    },
                    ChartSheetChartOutboundResource::EmbeddedPackage(embedded) => {
                        package.add_part(Box::new(BlobPart::new(
                            uri,
                            embedded.content_type.as_str().into(),
                            embedded.data.clone(),
                        )))
                    },
                }
            }
        }
    }
    package
        .get_part_mut(workbook_name)?
        .rels_mut()
        .add_relationship(
            conformance.chartsheet_rel().into(),
            chartsheet_uri.relative_ref(workbook_name.base_uri()),
            value.entry.workbook_relationship_id.clone(),
            false,
        );
    package
        .get_part_mut(&chartsheet_uri)?
        .rels_mut()
        .add_relationship(
            conformance.drawing_rel().into(),
            drawing_uri.relative_ref(chartsheet_uri.base_uri()),
            value.chartsheet.drawing_relationship_id.clone(),
            false,
        );
    if let (Some(resource), Some(uri)) = (&value.legacy_drawing, &legacy_uri) {
        package
            .get_part_mut(&chartsheet_uri)?
            .rels_mut()
            .add_relationship(
                conformance.vml_drawing_rel().into(),
                uri.relative_ref(chartsheet_uri.base_uri()),
                resource.relationship_id.clone(),
                false,
            );
    }
    if let (Some(resource), Some(uri)) = (&value.legacy_header_footer_drawing, &legacy_hf_uri) {
        package
            .get_part_mut(&chartsheet_uri)?
            .rels_mut()
            .add_relationship(
                conformance.vml_drawing_rel().into(),
                uri.relative_ref(chartsheet_uri.base_uri()),
                resource.relationship_id.clone(),
                false,
            );
    }
    if let (Some(picture), Some(uri)) = (&value.background_picture, &picture_uri) {
        package
            .get_part_mut(&chartsheet_uri)?
            .rels_mut()
            .add_relationship(
                conformance.image_rel().into(),
                uri.relative_ref(chartsheet_uri.base_uri()),
                picture.relationship_id.clone(),
                false,
            );
    }
    if let (Some(settings), Some(uri)) = (&value.printer_settings, &printer_uri) {
        package
            .get_part_mut(&chartsheet_uri)?
            .rels_mut()
            .add_relationship(
                conformance.printer_rel().into(),
                uri.relative_ref(chartsheet_uri.base_uri()),
                settings.relationship_id.clone(),
                false,
            );
    }
    for relationship in &value.extension_relationships {
        let (target, external) = match &relationship.target {
            ChartSheetExtensionRelationshipTarget::Internal { part_name } => (
                PackURI::new(part_name)
                    .map_err(OoxmlError::InvalidUri)?
                    .relative_ref(chartsheet_uri.base_uri()),
                false,
            ),
            ChartSheetExtensionRelationshipTarget::External { target } => (target.clone(), true),
        };
        package
            .get_part_mut(&chartsheet_uri)?
            .rels_mut()
            .add_relationship(
                relationship.relationship_type.clone(),
                target,
                relationship.relationship_id.clone(),
                external,
            );
    }
    for chart in &value.drawing.charts {
        let relationship_type = match &chart.kind {
            ChartSheetChartResourceKind::Classic => conformance.chart_rel(),
            ChartSheetChartResourceKind::Extended { .. } => CHART_EX_REL,
        };
        package
            .get_part_mut(&drawing_uri)?
            .rels_mut()
            .add_relationship(
                relationship_type.into(),
                chart_uris[&chart.relationship_id].relative_ref(drawing_uri.base_uri()),
                chart.relationship_id.clone(),
                false,
            );
        if let ChartSheetChartResourceKind::Extended {
            styles,
            color_styles,
            user_shapes,
            outbound_resources,
        } = &chart.kind
        {
            for (companions, relationship_type) in [
                (styles, CHART_STYLE_REL),
                (color_styles, CHART_COLOR_STYLE_REL),
            ] {
                for companion in companions {
                    let uri = &companion_uris[&(
                        chart.relationship_id.clone(),
                        companion.relationship_id.clone(),
                    )];
                    package
                        .get_part_mut(&chart_uris[&chart.relationship_id])?
                        .rels_mut()
                        .add_relationship(
                            relationship_type.into(),
                            uri.relative_ref(chart_uris[&chart.relationship_id].base_uri()),
                            companion.relationship_id.clone(),
                            false,
                        );
                }
            }
            if let Some(user_shapes) = user_shapes {
                let uri = &user_shape_uris[&chart.relationship_id];
                package
                    .get_part_mut(&chart_uris[&chart.relationship_id])?
                    .rels_mut()
                    .add_relationship(
                        conformance.chart_user_shapes_rel().into(),
                        uri.relative_ref(chart_uris[&chart.relationship_id].base_uri()),
                        user_shapes.relationship_id.clone(),
                        false,
                    );
                for image in &user_shapes.images {
                    let image_uri = &user_shape_image_uris
                        [&(chart.relationship_id.clone(), image.relationship_id.clone())];
                    package.get_part_mut(uri)?.rels_mut().add_relationship(
                        conformance.image_rel().into(),
                        image_uri.relative_ref(uri.base_uri()),
                        image.relationship_id.clone(),
                        false,
                    );
                }
            }
            for resource in outbound_resources {
                let key = (
                    chart.relationship_id.clone(),
                    resource.relationship_id().to_owned(),
                );
                let uri = &outbound_uris[&key];
                let relationship_type = match resource {
                    ChartSheetChartOutboundResource::Image(_) => conformance.image_rel(),
                    ChartSheetChartOutboundResource::ThemeOverride(_) => {
                        conformance.theme_override_rel()
                    },
                    ChartSheetChartOutboundResource::EmbeddedPackage(_) => {
                        conformance.package_rel()
                    },
                };
                package
                    .get_part_mut(&chart_uris[&chart.relationship_id])?
                    .rels_mut()
                    .add_relationship(
                        relationship_type.into(),
                        uri.relative_ref(chart_uris[&chart.relationship_id].base_uri()),
                        resource.relationship_id().to_owned(),
                        false,
                    );
                if let ChartSheetChartOutboundResource::ThemeOverride(theme) = resource {
                    for image in &theme.images {
                        let image_uri = &theme_image_uris[&(
                            chart.relationship_id.clone(),
                            theme.relationship_id.clone(),
                            image.relationship_id.clone(),
                        )];
                        package.get_part_mut(uri)?.rels_mut().add_relationship(
                            conformance.image_rel().into(),
                            image_uri.relative_ref(uri.base_uri()),
                            image.relationship_id.clone(),
                            false,
                        );
                    }
                }
            }
        }
    }
    Ok(())
}

fn validate_package_value(
    value: &ChartSheetPackage,
    conformance: ChartSheetConformance,
) -> Result<()> {
    validate_entry(&value.entry)?;
    validate_chartsheet(&value.chartsheet)?;
    if value.drawing.content_type != DRAWING_CT || value.drawing.data.len() > MAX_DRAWING_BYTES {
        return Err(invalid("invalid or oversized chartsheet drawing resource"));
    }
    let drawing_uri = PackURI::new(&value.drawing.part_name).map_err(OoxmlError::InvalidUri)?;
    if !drawing_uri.as_str().starts_with("/xl/drawings/") {
        return Err(invalid("drawing resource is outside /xl/drawings"));
    }
    let references = drawing_chart_references(&value.drawing.data, conformance)?;
    if references.len() != value.drawing.charts.len() {
        return Err(invalid(
            "drawing chart references and chart resources differ",
        ));
    }
    let mut resources = BTreeMap::new();
    let mut total = value.drawing.data.len();
    for chart in &value.drawing.charts {
        let reference = references
            .iter()
            .find(|reference| reference.relationship_id == chart.relationship_id)
            .ok_or_else(|| {
                invalid(format!(
                    "drawing does not reference chart relationship '{}'",
                    chart.relationship_id
                ))
            })?;
        validate_chart_resource_value(chart, reference, conformance, &mut total, &mut resources)?;
    }
    validate_vml_pair(
        value.chartsheet.legacy_drawing_relationship_id.as_deref(),
        value.legacy_drawing.as_ref(),
        "legacyDrawing",
        &mut total,
        &mut resources,
    )?;
    validate_vml_pair(
        value
            .chartsheet
            .legacy_header_footer_drawing_relationship_id
            .as_deref(),
        value.legacy_header_footer_drawing.as_ref(),
        "legacyDrawingHF",
        &mut total,
        &mut resources,
    )?;
    match (
        &value.chartsheet.background_picture_relationship_id,
        &value.background_picture,
    ) {
        (None, None) => {},
        (Some(id), Some(picture)) => {
            validate_id(id)?;
            validate_id(&picture.relationship_id)?;
            if id != &picture.relationship_id {
                return Err(invalid(
                    "chartsheet picture relationship and resource metadata differ",
                ));
            }
            if id == &value.chartsheet.drawing_relationship_id {
                return Err(invalid(
                    "chartsheet drawing and picture relationship IDs collide",
                ));
            }
            let uri = PackURI::new(&picture.part_name).map_err(OoxmlError::InvalidUri)?;
            if !uri.as_str().starts_with("/xl/media/") {
                return Err(invalid("background image resource is outside /xl/media"));
            }
            add_resource(
                &mut total,
                picture.data.len(),
                MAX_BACKGROUND_IMAGE_BYTES,
                "background image bytes",
            )?;
            if resources
                .insert(picture.part_name.clone(), &picture.data)
                .is_some()
            {
                return Err(invalid("duplicate chartsheet resource part name"));
            }
        },
        _ => {
            return Err(invalid(
                "chartsheet picture relationship and resource must either both be present or both be absent",
            ));
        },
    }
    let printer_id = value
        .chartsheet
        .page_setup
        .as_ref()
        .and_then(|setup| setup.printer_settings_relationship_id.as_deref());
    match (printer_id, value.printer_settings.as_ref()) {
        (None, None) => {},
        (Some(id), Some(settings)) => {
            validate_id(id)?;
            validate_id(&settings.relationship_id)?;
            if id != settings.relationship_id {
                return Err(invalid(
                    "chartsheet pageSetup and Printer Settings relationship IDs differ",
                ));
            }
            validate_settings_bytes(&settings.resource.data)?;
            let uri = PackURI::new(&settings.resource.part_name).map_err(OoxmlError::InvalidUri)?;
            validate_printer_settings_uri(&uri)?;
            add_resource(
                &mut total,
                settings.resource.data.len(),
                MAX_SETTINGS_BYTES,
                "Printer Settings bytes",
            )?;
            if resources
                .insert(settings.resource.part_name.clone(), &settings.resource.data)
                .is_some()
            {
                return Err(invalid("duplicate chartsheet resource part name"));
            }
        },
        _ => {
            return Err(invalid(
                "chartsheet pageSetup relationship and Printer Settings resource must either both be present or both be absent",
            ));
        },
    }
    validate_extension_relationships(value, conformance)?;
    Ok(())
}

fn validate_chart_resource_value<'a>(
    chart: &'a ChartSheetChartResource,
    reference: &DrawingChartReference,
    conformance: ChartSheetConformance,
    total: &mut usize,
    resources: &mut BTreeMap<String, &'a Vec<u8>>,
) -> Result<()> {
    validate_id(&chart.relationship_id)?;
    let uri = PackURI::new(&chart.part_name).map_err(OoxmlError::InvalidUri)?;
    if !uri.as_str().starts_with("/xl/charts/") || !uri.as_str().ends_with(".xml") {
        return Err(invalid(
            "chart resource is outside /xl/charts or lacks .xml suffix",
        ));
    }
    match (&chart.kind, reference.kind) {
        (ChartSheetChartResourceKind::Classic, DrawingChartKind::Classic) => {
            if chart.content_type != CHART_CT {
                return Err(invalid("classic chart has invalid content type"));
            }
            validate_chart_xml(&chart.data, conformance)?;
            add_resource(total, chart.data.len(), MAX_CHART_BYTES, "chart bytes")?;
        },
        (
            ChartSheetChartResourceKind::Extended {
                styles,
                color_styles,
                user_shapes,
                ..
            },
            DrawingChartKind::Extended,
        ) => {
            if chart.content_type != CHART_EX_CT {
                return Err(invalid("chartEx has invalid content type"));
            }
            validate_chart_ex_relationships(&chart.data, conformance)?;
            add_resource(total, chart.data.len(), MAX_CHART_EX_BYTES, "chartEx bytes")?;
            if styles.len() > MAX_CHART_STYLE_PARTS || color_styles.len() > MAX_CHART_STYLE_PARTS {
                return Err(limit("chart companion count"));
            }
            let mut ids = HashSet::new();
            for (companions, content_type, root, max_bytes) in [
                (styles, CHART_STYLE_CT, "chartStyle", MAX_CHART_STYLE_BYTES),
                (
                    color_styles,
                    CHART_COLOR_STYLE_CT,
                    "colorStyle",
                    MAX_CHART_COLOR_STYLE_BYTES,
                ),
            ] {
                for companion in companions {
                    validate_id(&companion.relationship_id)?;
                    if !ids.insert(companion.relationship_id.as_str()) {
                        return Err(invalid("chartEx companion relationship IDs collide"));
                    }
                    if companion.content_type != content_type {
                        return Err(invalid("chart companion has invalid content type"));
                    }
                    let uri = PackURI::new(&companion.part_name).map_err(OoxmlError::InvalidUri)?;
                    if !uri.as_str().starts_with("/xl/charts/") || !uri.as_str().ends_with(".xml") {
                        return Err(invalid(
                            "chart companion is outside /xl/charts or lacks .xml suffix",
                        ));
                    }
                    validate_chart_companion_xml(&companion.data, root, max_bytes)?;
                    add_resource(
                        total,
                        companion.data.len(),
                        max_bytes,
                        "chart companion bytes",
                    )?;
                    if resources
                        .insert(companion.part_name.clone(), &companion.data)
                        .is_some()
                    {
                        return Err(invalid("duplicate chartsheet resource part name"));
                    }
                }
            }
            if let Some(user_shapes) = user_shapes {
                validate_id(&user_shapes.relationship_id)?;
                if !ids.insert(user_shapes.relationship_id.as_str()) {
                    return Err(invalid("chartEx outbound relationship IDs collide"));
                }
                if user_shapes.content_type != CHART_USER_SHAPES_CT {
                    return Err(invalid("chartUserShapes has invalid content type"));
                }
                let uri = PackURI::new(&user_shapes.part_name).map_err(OoxmlError::InvalidUri)?;
                if !uri.as_str().starts_with("/xl/drawings/") || !uri.as_str().ends_with(".xml") {
                    return Err(invalid(
                        "chartUserShapes is outside /xl/drawings or lacks .xml suffix",
                    ));
                }
                let referenced = validate_chart_user_shapes_xml(&user_shapes.data, conformance)?;
                if referenced.len() != user_shapes.images.len()
                    || user_shapes.images.len() > MAX_CHART_USER_SHAPE_IMAGES
                {
                    return Err(invalid(
                        "chartUserShapes image relationship metadata does not match XML references",
                    ));
                }
                add_resource(
                    total,
                    user_shapes.data.len(),
                    MAX_CHART_USER_SHAPES_BYTES,
                    "chartUserShapes bytes",
                )?;
                if resources
                    .insert(user_shapes.part_name.clone(), &user_shapes.data)
                    .is_some()
                {
                    return Err(invalid("duplicate chartsheet resource part name"));
                }
                let mut image_ids = BTreeSet::new();
                for image in &user_shapes.images {
                    validate_id(&image.relationship_id)?;
                    if !referenced.contains(&image.relationship_id)
                        || !image_ids.insert(image.relationship_id.as_str())
                    {
                        return Err(invalid(
                            "chartUserShapes image metadata is duplicate or unreferenced",
                        ));
                    }
                    let image_uri =
                        PackURI::new(&image.part_name).map_err(OoxmlError::InvalidUri)?;
                    if !image_uri.as_str().starts_with("/xl/media/")
                        || !image.content_type.validates_part_name(image_uri.as_str())
                    {
                        return Err(invalid(
                            "invalid chart user-shape image path or content type suffix",
                        ));
                    }
                    add_resource(
                        total,
                        image.data.len(),
                        MAX_CHART_USER_SHAPE_IMAGE_BYTES,
                        "chart user-shape image bytes",
                    )?;
                    if resources
                        .insert(image.part_name.clone(), &image.data)
                        .is_some()
                    {
                        return Err(invalid("duplicate chartsheet resource part name"));
                    }
                }
            }
            validate_chart_outbound_resources(chart, conformance, total, resources)?;
        },
        _ => {
            return Err(invalid(
                "drawing chart reference kind and chart resource kind differ",
            ));
        },
    }
    if resources
        .insert(chart.part_name.clone(), &chart.data)
        .is_some()
    {
        return Err(invalid("duplicate chart resource part name"));
    }
    Ok(())
}

fn validate_chart_outbound_resources<'a>(
    chart: &'a ChartSheetChartResource,
    conformance: ChartSheetConformance,
    total: &mut usize,
    resources: &mut BTreeMap<String, &'a Vec<u8>>,
) -> Result<()> {
    let ChartSheetChartResourceKind::Extended {
        styles,
        color_styles,
        user_shapes,
        outbound_resources,
    } = &chart.kind
    else {
        return Ok(());
    };
    let references = validate_chart_ex_relationships(&chart.data, conformance)?;
    let mut source_ids = HashSet::new();
    for companion in styles.iter().chain(color_styles) {
        source_ids.insert(companion.relationship_id.as_str());
    }
    if let Some(user_shapes) = user_shapes {
        source_ids.insert(user_shapes.relationship_id.as_str());
    }
    let mut direct_ids = BTreeSet::new();
    let mut package_id = None;
    let mut theme_count = 0usize;
    let mut package_count = 0usize;
    if outbound_resources
        .iter()
        .filter(|resource| matches!(resource, ChartSheetChartOutboundResource::Image(_)))
        .count()
        > MAX_CHART_DIRECT_IMAGES
    {
        return Err(limit("chartEx direct image count"));
    }
    for resource in outbound_resources {
        validate_id(resource.relationship_id())?;
        if !source_ids.insert(resource.relationship_id()) {
            return Err(invalid("chartEx outbound relationship IDs collide"));
        }
        match resource {
            ChartSheetChartOutboundResource::Image(image) => {
                if !direct_ids.insert(image.relationship_id.clone()) {
                    return Err(invalid("chartEx direct image relationship IDs collide"));
                }
                validate_chart_image_value(
                    image,
                    total,
                    resources,
                    MAX_CHART_USER_SHAPE_IMAGE_BYTES,
                    "chartEx direct image bytes",
                )?;
            },
            ChartSheetChartOutboundResource::ThemeOverride(theme) => {
                theme_count += 1;
                if theme_count > 1 {
                    return Err(invalid("chartEx has multiple themeOverride relationships"));
                }
                if theme.content_type != THEME_OVERRIDE_CT {
                    return Err(invalid("themeOverride has invalid content type"));
                }
                let uri = PackURI::new(&theme.part_name).map_err(OoxmlError::InvalidUri)?;
                if !uri.as_str().starts_with("/xl/theme/") || !uri.as_str().ends_with(".xml") {
                    return Err(invalid(
                        "themeOverride is outside /xl/theme or lacks .xml suffix",
                    ));
                }
                let theme_references = validate_theme_override_xml(&theme.data, conformance)?;
                if theme.images.len() > MAX_CHART_THEME_IMAGES
                    || theme.images.len() != theme_references.len()
                {
                    return Err(invalid(
                        "themeOverride image relationship metadata does not match XML references",
                    ));
                }
                add_resource(
                    total,
                    theme.data.len(),
                    MAX_CHART_THEME_OVERRIDE_BYTES,
                    "themeOverride bytes",
                )?;
                if resources
                    .insert(theme.part_name.clone(), &theme.data)
                    .is_some()
                {
                    return Err(invalid("duplicate chartsheet resource part name"));
                }
                let mut image_ids = BTreeSet::new();
                for image in &theme.images {
                    validate_id(&image.relationship_id)?;
                    if !image_ids.insert(image.relationship_id.clone())
                        || !theme_references.contains(&image.relationship_id)
                    {
                        return Err(invalid(
                            "themeOverride image metadata is duplicate or unreferenced",
                        ));
                    }
                    validate_chart_image_value(
                        image,
                        total,
                        resources,
                        MAX_CHART_USER_SHAPE_IMAGE_BYTES,
                        "themeOverride image bytes",
                    )?;
                }
            },
            ChartSheetChartOutboundResource::EmbeddedPackage(embedded) => {
                package_count += 1;
                if package_count > 1 {
                    return Err(invalid(
                        "chartEx has multiple embedded package relationships",
                    ));
                }
                let uri = PackURI::new(&embedded.part_name).map_err(OoxmlError::InvalidUri)?;
                if !uri.as_str().starts_with("/xl/embeddings/")
                    || !embedded.content_type.validates_part_name(uri.as_str())
                {
                    return Err(invalid(
                        "invalid chartEx embedded package path or content type suffix",
                    ));
                }
                add_resource(
                    total,
                    embedded.data.len(),
                    MAX_CHART_EMBEDDED_PACKAGE_BYTES,
                    "chartEx embedded package bytes",
                )?;
                if resources
                    .insert(embedded.part_name.clone(), &embedded.data)
                    .is_some()
                {
                    return Err(invalid("duplicate chartsheet resource part name"));
                }
                package_id = Some(embedded.relationship_id.clone());
            },
        }
    }
    if direct_ids != references.images || package_id != references.package {
        return Err(invalid(
            "chartEx outbound relationship metadata does not match XML references",
        ));
    }
    Ok(())
}

fn validate_chart_image_value<'a>(
    image: &'a ChartSheetChartImageResource,
    total: &mut usize,
    resources: &mut BTreeMap<String, &'a Vec<u8>>,
    max_bytes: usize,
    label: &str,
) -> Result<()> {
    let uri = PackURI::new(&image.part_name).map_err(OoxmlError::InvalidUri)?;
    if !uri.as_str().starts_with("/xl/media/")
        || !image.content_type.validates_part_name(uri.as_str())
    {
        return Err(invalid("invalid chart image path or content type suffix"));
    }
    add_resource(total, image.data.len(), max_bytes, label)?;
    if resources
        .insert(image.part_name.clone(), &image.data)
        .is_some()
    {
        return Err(invalid("duplicate chartsheet resource part name"));
    }
    Ok(())
}

fn validate_vml_pair<'a>(
    id: Option<&str>,
    resource: Option<&'a ChartSheetVmlDrawingResource>,
    label: &str,
    total: &mut usize,
    resources: &mut BTreeMap<String, &'a Vec<u8>>,
) -> Result<()> {
    match (id, resource) {
        (None, None) => Ok(()),
        (Some(id), Some(resource)) => {
            validate_id(id)?;
            validate_id(&resource.relationship_id)?;
            if id != resource.relationship_id {
                return Err(invalid(format!(
                    "{label} relationship and resource metadata differ"
                )));
            }
            let uri = PackURI::new(&resource.part_name).map_err(OoxmlError::InvalidUri)?;
            if !uri.as_str().starts_with("/xl/drawings/")
                || !uri.as_str().ends_with(".vml")
                || resource.content_type != VML_DRAWING_CT
            {
                return Err(invalid(format!(
                    "invalid {label} VML resource path or content type"
                )));
            }
            add_resource(
                total,
                resource.data.len(),
                MAX_VML_DRAWING_BYTES,
                "VML drawing bytes",
            )?;
            if resources
                .insert(resource.part_name.clone(), &resource.data)
                .is_some()
            {
                return Err(invalid("duplicate chartsheet resource part name"));
            }
            Ok(())
        },
        _ => Err(invalid(format!(
            "{label} relationship and resource must either both be present or both be absent"
        ))),
    }
}

fn validate_chartsheet(value: &ChartSheet) -> Result<()> {
    if value.views.is_empty() || value.views.len() > MAX_VIEWS {
        return Err(invalid("chartsheet requires a bounded non-empty view list"));
    }
    validate_id(&value.drawing_relationship_id)?;
    let mut relationship_ids = HashSet::new();
    relationship_ids.insert(value.drawing_relationship_id.as_str());
    let printer_id = value
        .page_setup
        .as_ref()
        .and_then(|setup| setup.printer_settings_relationship_id.as_ref());
    for id in [
        value.legacy_drawing_relationship_id.as_ref(),
        value.legacy_header_footer_drawing_relationship_id.as_ref(),
        value.background_picture_relationship_id.as_ref(),
        printer_id,
    ]
    .into_iter()
    .flatten()
    {
        validate_id(id)?;
        if !relationship_ids.insert(id.as_str()) {
            return Err(invalid("chartsheet relationship IDs collide"));
        }
    }
    let mut view_ids = HashSet::new();
    for view in &value.views {
        if view.zoom_scale.is_some_and(|v| !(10..=400).contains(&v)) {
            return Err(invalid("chartsheet zoomScale must be between 10 and 400"));
        }
        if !view_ids.insert(view.workbook_view_id) {
            return Err(invalid("duplicate chartsheet workbookViewId"));
        }
    }
    if let Some(properties) = &value.properties {
        if let Some(name) = &properties.code_name {
            bounded(name)?;
        }
        if let Some(color) = &properties.tab_color {
            let bases = usize::from(color.automatic.is_some())
                + usize::from(color.indexed.is_some())
                + usize::from(color.rgb.is_some())
                + usize::from(color.theme.is_some());
            if bases > 1 {
                return Err(invalid("tab color has multiple base color selectors"));
            }
            if let Some(rgb) = &color.rgb
                && (rgb.len() != 8 || !rgb.bytes().all(|b| b.is_ascii_hexdigit()))
            {
                return Err(invalid("tab color rgb must contain eight hex digits"));
            }
            if color
                .tint
                .is_some_and(|v| !v.is_finite() || !(-1.0..=1.0).contains(&v))
            {
                return Err(invalid("tab color tint is outside [-1, 1]"));
            }
        }
    }
    if let Some(protection) = &value.protection
        && let Some(password) = &protection.password_hash
        && (password.len() != 4 || !password.bytes().all(|b| b.is_ascii_hexdigit()))
    {
        return Err(invalid(
            "chartsheet password hash must contain four hex digits",
        ));
    }
    if let Some(custom_views) = &value.custom_views {
        if custom_views.is_empty() {
            return Err(invalid(
                "customSheetViews requires at least one customSheetView",
            ));
        }
        if custom_views.len() > MAX_CUSTOM_VIEWS {
            return Err(limit("custom view count"));
        }
        let mut guids = HashSet::new();
        for view in custom_views {
            validate_guid(&view.guid)?;
            if !guids.insert(view.guid.to_ascii_lowercase()) {
                return Err(invalid(format!(
                    "duplicate custom chartsheet view GUID '{}'",
                    view.guid
                )));
            }
            if view.scale.is_some_and(|scale| !(10..=400).contains(&scale)) {
                return Err(invalid(
                    "custom chartsheet view scale must be between 10 and 400",
                ));
            }
        }
    }
    if let Some(m) = value.margins {
        for margin in [m.left, m.right, m.top, m.bottom, m.header, m.footer] {
            if !margin.is_finite() || !(0.0..49.0).contains(&margin) {
                return Err(invalid(
                    "chartsheet margin is outside Office's [0, 49) range",
                ));
            }
        }
    }
    if let Some(setup) = &value.page_setup {
        if setup.first_page_number.is_some_and(|v| v > 65_534) {
            return Err(invalid("firstPageNumber exceeds Excel's limit"));
        }
        if setup.copies.is_some_and(|v| !(1..=32_767).contains(&v)) {
            return Err(invalid("copies is outside Excel's supported range"));
        }
        if setup.horizontal_dpi == Some(0) || setup.vertical_dpi == Some(0) {
            return Err(invalid("page setup DPI must be positive"));
        }
    }
    if let Some(hf) = &value.header_footer {
        for text in [
            &hf.odd_header,
            &hf.odd_footer,
            &hf.even_header,
            &hf.even_footer,
            &hf.first_header,
            &hf.first_footer,
        ]
        .into_iter()
        .flatten()
        {
            bounded(text)?;
        }
    }
    if let Some(items) = &value.web_publish_items {
        validate_web_publish_items(items)?;
    }
    if let Some(extensions) = &value.extension_list {
        validate_extension_list(extensions)?;
    }
    Ok(())
}

fn validate_web_publish_items(value: &ChartSheetWebPublishItems) -> Result<()> {
    if value.items.is_empty() {
        return Err(invalid(
            "webPublishItems requires at least one webPublishItem",
        ));
    }
    if value.items.len() > MAX_WEB_PUBLISH_ITEMS {
        return Err(limit("web publish item count"));
    }
    if value
        .count
        .is_some_and(|count| usize::try_from(count).ok() != Some(value.items.len()))
    {
        return Err(invalid(
            "webPublishItems count does not match item cardinality",
        ));
    }
    let mut ids = HashSet::new();
    let mut div_ids = HashSet::new();
    for item in &value.items {
        if !ids.insert(item.id) {
            return Err(invalid(format!(
                "duplicate web publish item id {}",
                item.id
            )));
        }
        if item.div_id.is_empty() || item.destination_file.is_empty() {
            return Err(invalid(
                "web publish divId and destinationFile must be non-empty",
            ));
        }
        for string in [
            Some(&item.div_id),
            Some(&item.destination_file),
            item.source_ref.as_ref(),
            item.source_object.as_ref(),
            item.title.as_ref(),
        ]
        .into_iter()
        .flatten()
        {
            bounded_web_publish(string)?;
        }
        if !div_ids.insert(item.div_id.as_str()) {
            return Err(invalid(format!(
                "duplicate web publish divId '{}'",
                item.div_id
            )));
        }
        if item.source_type == ChartSheetWebSourceType::Range
            && item.source_ref.as_deref().is_none_or(str::is_empty)
        {
            return Err(invalid(
                "range web publish item requires non-empty sourceRef",
            ));
        }
        if matches!(
            item.source_type,
            ChartSheetWebSourceType::PivotTable
                | ChartSheetWebSourceType::Query
                | ChartSheetWebSourceType::Label
        ) && item.source_object.as_deref().is_none_or(str::is_empty)
        {
            return Err(invalid(
                "object web publish item requires non-empty sourceObject",
            ));
        }
    }
    Ok(())
}

fn bounded_web_publish(value: &str) -> Result<()> {
    if value.len() <= MAX_WEB_PUBLISH_STRING_BYTES {
        Ok(())
    } else {
        Err(limit("web publish string bytes"))
    }
}

fn workbook_entry(
    root: &Node,
    conformance: ChartSheetConformance,
    relationship_id: &str,
    part_name: String,
) -> Result<ChartSheetEntry> {
    let sheets = required_child(root, conformance.sml(), "sheets")?;
    let mut found = None;
    for sheet in &sheets.children {
        if sheet.namespace == conformance.sml()
            && sheet.name == "sheet"
            && optional(sheet, conformance.rel(), "id") == Some(relationship_id)
        {
            if found.is_some() {
                return Err(invalid(
                    "multiple workbook sheets reference the chartsheet relationship",
                ));
            }
            found = Some(parse_entry(sheet, conformance, part_name.clone())?);
        }
    }
    found.ok_or_else(|| invalid("workbook has no sheet entry for the chartsheet relationship"))
}

fn parse_entry(
    node: &Node,
    conformance: ChartSheetConformance,
    part_name: String,
) -> Result<ChartSheetEntry> {
    leaf(node, "workbook sheet")?;
    let state = optional(node, "", "state")
        .map(parse_state)
        .transpose()?
        .unwrap_or(ChartSheetState::Visible);
    Ok(ChartSheetEntry {
        name: required(node, "", "name")?.to_owned(),
        sheet_id: required(node, "", "sheetId")?
            .parse()
            .map_err(|_| invalid("invalid workbook sheetId"))?,
        state,
        workbook_relationship_id: required(node, conformance.rel(), "id")?.to_owned(),
        part_name,
    })
}

fn validate_new_entry(
    root: &Node,
    conformance: ChartSheetConformance,
    entry: &ChartSheetEntry,
) -> Result<()> {
    let sheets = required_child(root, conformance.sml(), "sheets")?;
    for sheet in &sheets.children {
        if sheet.namespace == conformance.sml() && sheet.name == "sheet" {
            if optional(sheet, "", "name").is_some_and(|v| v.eq_ignore_ascii_case(&entry.name)) {
                return Err(invalid("workbook sheet name already exists"));
            }
            if optional(sheet, "", "sheetId") == Some(entry.sheet_id.to_string().as_str()) {
                return Err(invalid("workbook sheetId already exists"));
            }
        }
    }
    Ok(())
}

fn validate_entry(entry: &ChartSheetEntry) -> Result<()> {
    bounded(&entry.name)?;
    if entry.name.is_empty()
        || entry.name.chars().count() > 31
        || entry
            .name
            .chars()
            .any(|c| matches!(c, ':' | '\\' | '/' | '?' | '*' | '[' | ']'))
    {
        return Err(invalid("invalid Excel chartsheet name"));
    }
    if entry.sheet_id == 0 {
        return Err(invalid("chartsheet sheetId must be positive"));
    }
    validate_id(&entry.workbook_relationship_id)?;
    let uri = PackURI::new(&entry.part_name).map_err(OoxmlError::InvalidUri)?;
    if !uri.as_str().starts_with("/xl/chartsheets/") {
        return Err(invalid("chartsheet part is outside /xl/chartsheets"));
    }
    Ok(())
}

fn insert_workbook_entry(
    xml: &[u8],
    entry: &ChartSheetEntry,
    conformance: ChartSheetConformance,
) -> Result<Vec<u8>> {
    let mut fragment = Vec::new();
    fragment.extend_from_slice(b"<x:sheet xmlns:x=\"");
    escape(&mut fragment, conformance.sml());
    fragment.extend_from_slice(b"\" xmlns:r=\"");
    escape(&mut fragment, conformance.rel());
    fragment.extend_from_slice(b"\"");
    attr(&mut fragment, "name", &entry.name);
    attr(&mut fragment, "sheetId", &entry.sheet_id.to_string());
    if entry.state != ChartSheetState::Visible {
        attr(
            &mut fragment,
            "state",
            match entry.state {
                ChartSheetState::Visible => "visible",
                ChartSheetState::Hidden => "hidden",
                ChartSheetState::VeryHidden => "veryHidden",
            },
        );
    }
    attr(&mut fragment, "r:id", &entry.workbook_relationship_id);
    fragment.extend_from_slice(b"/>");
    let mut reader = NsReader::from_reader(xml);
    let mut depth = 0usize;
    let mut sheets_depth = None;
    let mut position = None;
    loop {
        let start = usize::try_from(reader.buffer_position())
            .map_err(|_| invalid("workbook XML offset overflow"))?;
        let (namespace, event) = reader.read_resolved_event().map_err(xml_error)?;
        match event {
            Event::Start(element) => {
                let core = matches!(namespace, ResolveResult::Bound(Namespace(v)) if v == conformance.sml().as_bytes());
                depth += 1;
                if core
                    && element.local_name().as_ref() == b"sheets"
                    && sheets_depth.replace(depth).is_some()
                {
                    return Err(invalid("workbook has multiple sheets collections"));
                }
            },
            Event::Empty(element) if element.local_name().as_ref() == b"sheets" => {
                return Err(invalid("cannot insert into empty sheets collection"));
            },
            Event::End(element) => {
                if sheets_depth == Some(depth) && element.local_name().as_ref() == b"sheets" {
                    position = Some(start);
                }
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("unexpected workbook closing element"))?;
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("DTDs and processing instructions are rejected"));
            },
            Event::Eof => break,
            _ => {},
        }
    }
    let position = position.ok_or_else(|| invalid("workbook is missing sheets collection"))?;
    let size = xml
        .len()
        .checked_add(fragment.len())
        .ok_or_else(|| limit("updated workbook bytes"))?;
    if size > MAX_XML_BYTES {
        return Err(limit("updated workbook bytes"));
    }
    let mut out = Vec::with_capacity(size);
    out.extend_from_slice(&xml[..position]);
    out.extend_from_slice(&fragment);
    out.extend_from_slice(&xml[position..]);
    Ok(out)
}

fn validate_extension_uri(uri: &str) -> Result<()> {
    if uri.is_empty() {
        return Err(invalid("extension uri cannot be empty"));
    }
    if uri.len() > MAX_EXTENSION_URI_BYTES {
        return Err(limit("extension uri bytes"));
    }
    if uri.chars().any(|c| c.is_whitespace() || c.is_control()) {
        return Err(invalid(
            "extension uri must not contain whitespace or controls",
        ));
    }
    Ok(())
}

fn validate_extension_list(value: &ChartSheetExtensionList) -> Result<()> {
    if value.extensions.is_empty() {
        return Err(invalid("extLst requires at least one ext"));
    }
    if value.extensions.len() > MAX_EXTENSIONS {
        return Err(limit("extension count"));
    }
    let mut total = 0usize;
    for extension in &value.extensions {
        validate_extension_uri(&extension.uri)?;
        parse_document(&extension.payload_xml, MAX_EXTENSION_PAYLOAD_BYTES)?;
        total = total
            .checked_add(extension.payload_xml.len())
            .ok_or_else(|| limit("retained extension bytes"))?;
        if total > MAX_RETAINED_EXTENSION_BYTES {
            return Err(limit("retained extension bytes"));
        }
    }
    Ok(())
}

fn canonical_extension_payload(xml: &[u8], conformance: ChartSheetConformance) -> Result<Vec<u8>> {
    let node = parse_document(xml, MAX_EXTENSION_PAYLOAD_BYTES)?;
    canonical_extension_payload_node(&node, conformance)
}

fn canonical_extension_payload_node(
    node: &Node,
    conformance: ChartSheetConformance,
) -> Result<Vec<u8>> {
    let mut namespaces = BTreeSet::new();
    collect_namespaces(node, &mut namespaces);
    let mut prefixes = BTreeMap::new();
    if namespaces.remove(conformance.sml()) {
        prefixes.insert(conformance.sml().to_owned(), "x".to_owned());
    }
    if namespaces.remove(conformance.rel()) {
        prefixes.insert(conformance.rel().to_owned(), "r".to_owned());
    }
    namespaces.remove("http://www.w3.org/XML/1998/namespace");
    for (index, namespace) in namespaces
        .into_iter()
        .filter(|namespace| !namespace.is_empty())
        .enumerate()
    {
        prefixes.insert(namespace, format!("e{index}"));
    }
    let mut out = Vec::new();
    write_canonical_node(&mut out, node, &prefixes, true)?;
    if out.len() > MAX_EXTENSION_PAYLOAD_BYTES {
        return Err(limit("extension payload bytes"));
    }
    Ok(out)
}

fn collect_namespaces(node: &Node, namespaces: &mut BTreeSet<String>) {
    namespaces.insert(node.namespace.clone());
    for attribute in &node.attributes {
        namespaces.insert(attribute.namespace.clone());
    }
    for child in &node.children {
        collect_namespaces(child, namespaces);
    }
}
fn qualified_name(
    namespace: &str,
    name: &str,
    prefixes: &BTreeMap<String, String>,
) -> Result<String> {
    if namespace.is_empty() {
        return Ok(name.to_owned());
    }
    if namespace == "http://www.w3.org/XML/1998/namespace" {
        return Ok(format!("xml:{name}"));
    }
    let prefix = prefixes
        .get(namespace)
        .ok_or_else(|| invalid("missing canonical namespace prefix"))?;
    Ok(format!("{prefix}:{name}"))
}
fn write_canonical_node(
    out: &mut Vec<u8>,
    node: &Node,
    prefixes: &BTreeMap<String, String>,
    root: bool,
) -> Result<()> {
    let name = qualified_name(&node.namespace, &node.name, prefixes)?;
    out.push(b'<');
    out.extend_from_slice(name.as_bytes());
    if root {
        let mut declarations = prefixes
            .iter()
            .map(|(namespace, prefix)| (prefix, namespace))
            .collect::<Vec<_>>();
        declarations.sort();
        for (prefix, namespace) in declarations {
            attr(out, &format!("xmlns:{prefix}"), namespace);
        }
    }
    let mut attributes = node.attributes.iter().collect::<Vec<_>>();
    attributes
        .sort_by(|left, right| (&left.namespace, &left.name).cmp(&(&right.namespace, &right.name)));
    for attribute in attributes {
        let name = qualified_name(&attribute.namespace, &attribute.name, prefixes)?;
        attr(out, &name, &attribute.value);
    }
    if node.content.is_empty() {
        out.extend_from_slice(b"/>");
        return Ok(());
    }
    out.push(b'>');
    for content in &node.content {
        match content {
            NodeContent::Text(value) => escape_text(out, value),
            NodeContent::Child(index) => {
                write_canonical_node(out, &node.children[*index], prefixes, false)?
            },
        }
    }
    out.extend_from_slice(b"</");
    out.extend_from_slice(name.as_bytes());
    out.push(b'>');
    Ok(())
}

fn known_chartsheet_relationship_ids(value: &ChartSheet) -> BTreeSet<String> {
    let mut ids = BTreeSet::new();
    ids.insert(value.drawing_relationship_id.clone());
    for id in [
        value.legacy_drawing_relationship_id.as_ref(),
        value.legacy_header_footer_drawing_relationship_id.as_ref(),
        value.background_picture_relationship_id.as_ref(),
        value
            .page_setup
            .as_ref()
            .and_then(|setup| setup.printer_settings_relationship_id.as_ref()),
    ]
    .into_iter()
    .flatten()
    {
        ids.insert(id.clone());
    }
    ids
}
fn extension_relationship_ids(
    value: &ChartSheet,
    conformance: ChartSheetConformance,
) -> Result<BTreeSet<String>> {
    let mut ids = BTreeSet::new();
    if let Some(list) = &value.extension_list {
        for extension in &list.extensions {
            let node = parse_document(&extension.payload_xml, MAX_EXTENSION_PAYLOAD_BYTES)?;
            collect_extension_relationship_ids(&node, conformance.rel(), &mut ids)?;
        }
    }
    Ok(ids)
}
fn collect_extension_relationship_ids(
    node: &Node,
    relationship_namespace: &str,
    ids: &mut BTreeSet<String>,
) -> Result<()> {
    for attribute in &node.attributes {
        if attribute.namespace == relationship_namespace {
            validate_id(&attribute.value)?;
            ids.insert(attribute.value.clone());
        }
    }
    for child in &node.children {
        collect_extension_relationship_ids(child, relationship_namespace, ids)?;
    }
    Ok(())
}

fn validate_extension_relationship_string(value: &str, label: &str) -> Result<()> {
    if value.is_empty() {
        return Err(invalid(format!(
            "extension relationship {label} cannot be empty"
        )));
    }
    if value.len() > MAX_EXTENSION_RELATIONSHIP_STRING_BYTES {
        return Err(limit("extension relationship string bytes"));
    }
    Ok(())
}
fn validate_extension_relationships(
    value: &ChartSheetPackage,
    conformance: ChartSheetConformance,
) -> Result<()> {
    if value.extension_relationships.len() > MAX_EXTENSION_RELATIONSHIPS {
        return Err(limit("extension relationship count"));
    }
    let referenced = extension_relationship_ids(&value.chartsheet, conformance)?;
    let known = known_chartsheet_relationship_ids(&value.chartsheet);
    let unknown = referenced
        .difference(&known)
        .cloned()
        .collect::<BTreeSet<_>>();
    if value.extension_relationships.len() != unknown.len() {
        return Err(invalid(
            "extension relationship metadata does not match referenced unknown relationships",
        ));
    }
    let mut seen = BTreeSet::new();
    for relationship in &value.extension_relationships {
        validate_id(&relationship.relationship_id)?;
        if !unknown.contains(&relationship.relationship_id)
            || !seen.insert(relationship.relationship_id.clone())
        {
            return Err(invalid(
                "extension relationship metadata is duplicate or unreferenced",
            ));
        }
        validate_extension_relationship_string(&relationship.relationship_type, "type")?;
        match &relationship.target {
            ChartSheetExtensionRelationshipTarget::Internal { part_name } => {
                validate_extension_relationship_string(part_name, "target")?;
                let uri = PackURI::new(part_name).map_err(OoxmlError::InvalidUri)?;
                if !uri.as_str().starts_with('/') {
                    return Err(invalid(
                        "internal extension relationship target must be an absolute part name",
                    ));
                }
            },
            ChartSheetExtensionRelationshipTarget::External { target } => {
                validate_extension_relationship_string(target, "target")?
            },
        }
    }
    Ok(())
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum DrawingChartKind {
    Classic,
    Extended,
}
#[derive(Debug, Clone, PartialEq, Eq)]
struct DrawingChartReference {
    relationship_id: String,
    kind: DrawingChartKind,
}

fn chart_ex_mce_capabilities() -> MceCapabilities {
    let mut capabilities = MceCapabilities::ooxml_baseline();
    capabilities
        .understand_namespace(CHART_EX)
        .understand_namespace(CHART_EX_CHOICE);
    capabilities
}
fn drawing_chart_references(
    xml: &[u8],
    conformance: ChartSheetConformance,
) -> Result<Vec<DrawingChartReference>> {
    if xml.len() > MAX_DRAWING_BYTES {
        return Err(limit("drawing bytes"));
    }
    let root =
        parse_document_with_capabilities(xml, MAX_DRAWING_BYTES, &chart_ex_mce_capabilities())?;
    if root.namespace != conformance.xdr() || root.name != "wsDr" {
        return Err(invalid(
            "drawing root does not match chartsheet conformance",
        ));
    }
    let mut references = Vec::new();
    collect_drawing_chart_references(&root, conformance, &mut references)?;
    if references.len() > MAX_CHARTS {
        return Err(limit("chart count"));
    }
    let mut ids = HashSet::new();
    for reference in &references {
        validate_id(&reference.relationship_id)?;
        if !ids.insert(reference.relationship_id.as_str()) {
            return Err(invalid("drawing chart relationship IDs collide"));
        }
    }
    Ok(references)
}
fn collect_drawing_chart_references(
    node: &Node,
    conformance: ChartSheetConformance,
    references: &mut Vec<DrawingChartReference>,
) -> Result<()> {
    if matches!(node.namespace.as_str(), DRAWING_MAIN | STRICT_DRAWING_MAIN)
        && node.name == "graphicData"
    {
        let has_chart_ex = node
            .children
            .iter()
            .any(|child| child.namespace == CHART_EX && child.name == "chart");
        if has_chart_ex || optional(node, "", "uri") == Some(CHART_EX) {
            if optional(node, "", "uri") != Some(CHART_EX) {
                return Err(invalid(
                    "cx:chart requires the exact chartEx graphicData URI",
                ));
            }
            whitespace(node)?;
            no_attributes(node, &[("", "uri")])?;
            if node.children.len() != 1 {
                return Err(invalid(
                    "chartEx graphicData requires exactly one cx:chart child",
                ));
            }
            let chart = &node.children[0];
            if chart.namespace != CHART_EX || chart.name != "chart" {
                return Err(invalid("chartEx graphicData has an invalid root child"));
            }
            leaf(chart, "chartEx drawing reference")?;
            whitespace(chart)?;
            no_attributes(chart, &[(conformance.rel(), "id")])?;
            references.push(DrawingChartReference {
                relationship_id: required(chart, conformance.rel(), "id")?.to_owned(),
                kind: DrawingChartKind::Extended,
            });
            return Ok(());
        }
    }
    if node.namespace == conformance.chart() && node.name == "chart" {
        references.push(DrawingChartReference {
            relationship_id: required(node, conformance.rel(), "id")?.to_owned(),
            kind: DrawingChartKind::Classic,
        });
    }
    for child in &node.children {
        collect_drawing_chart_references(child, conformance, references)?;
    }
    Ok(())
}

fn validate_chart_xml(xml: &[u8], conformance: ChartSheetConformance) -> Result<()> {
    if xml.len() > MAX_CHART_BYTES {
        return Err(limit("chart bytes"));
    }
    let root = parse_document(xml, MAX_CHART_BYTES)?;
    if root.namespace == conformance.chart() && root.name == "chartSpace" {
        Ok(())
    } else {
        Err(invalid("chart root does not match chartsheet conformance"))
    }
}
fn validate_chart_companion_xml(xml: &[u8], root_name: &str, max_bytes: usize) -> Result<()> {
    if xml.len() > max_bytes {
        return Err(limit("chart companion bytes"));
    }
    let root = parse_document(xml, max_bytes)?;
    if root.namespace == CHART_STYLE && root.name == root_name {
        Ok(())
    } else {
        Err(invalid(format!(
            "invalid chart companion root '{root_name}'"
        )))
    }
}
fn validate_chart_user_shapes_xml(
    xml: &[u8],
    conformance: ChartSheetConformance,
) -> Result<BTreeSet<String>> {
    if xml.len() > MAX_CHART_USER_SHAPES_BYTES {
        return Err(limit("chartUserShapes bytes"));
    }
    let root = parse_document(xml, MAX_CHART_USER_SHAPES_BYTES)?;
    if root.namespace != conformance.chart() || root.name != "userShapes" {
        return Err(invalid(
            "chartUserShapes root does not match chartsheet conformance",
        ));
    }
    let mut ids = BTreeSet::new();
    collect_extension_relationship_ids(&root, conformance.rel(), &mut ids)?;
    Ok(ids)
}
#[derive(Default)]
struct ChartExRelationshipReferences {
    images: BTreeSet<String>,
    package: Option<String>,
}
fn validate_chart_ex_relationships(
    xml: &[u8],
    conformance: ChartSheetConformance,
) -> Result<ChartExRelationshipReferences> {
    if xml.len() > MAX_CHART_EX_BYTES {
        return Err(limit("chartEx bytes"));
    }
    let root = parse_document(xml, MAX_CHART_EX_BYTES)?;
    if root.namespace != CHART_EX || root.name != "chartSpace" {
        return Err(invalid("invalid chartEx root"));
    }
    let mut references = ChartExRelationshipReferences::default();
    collect_chart_ex_relationships(&root, conformance, &mut references)?;
    Ok(references)
}
fn collect_chart_ex_relationships(
    node: &Node,
    conformance: ChartSheetConformance,
    references: &mut ChartExRelationshipReferences,
) -> Result<()> {
    let external_data = node.namespace == CHART_EX && node.name == "externalData";
    if external_data
        && optional(node, CHART_EX, "autoUpdate").is_some_and(|value| matches!(value, "1" | "true"))
    {
        return Err(invalid("auto-updating chartEx external data is rejected"));
    }
    for attribute in &node.attributes {
        if attribute.namespace == conformance.rel() {
            validate_id(&attribute.value)?;
            if external_data && attribute.name == "id" {
                if references
                    .package
                    .replace(attribute.value.clone())
                    .is_some()
                {
                    return Err(invalid(
                        "chartEx has multiple externalData package references",
                    ));
                }
            } else {
                if external_data {
                    return Err(invalid(
                        "chartEx externalData has an unsupported relationship attribute",
                    ));
                }
                references.images.insert(attribute.value.clone());
            }
        }
    }
    for child in &node.children {
        collect_chart_ex_relationships(child, conformance, references)?;
    }
    Ok(())
}
fn validate_theme_override_xml(
    xml: &[u8],
    conformance: ChartSheetConformance,
) -> Result<BTreeSet<String>> {
    if xml.len() > MAX_CHART_THEME_OVERRIDE_BYTES {
        return Err(limit("themeOverride bytes"));
    }
    let root = parse_document(xml, MAX_CHART_THEME_OVERRIDE_BYTES)?;
    let namespace = if conformance == ChartSheetConformance::Strict {
        STRICT_DRAWING_MAIN
    } else {
        DRAWING_MAIN
    };
    if root.namespace != namespace || root.name != "themeOverride" {
        return Err(invalid(
            "themeOverride root does not match chartsheet conformance",
        ));
    }
    let mut ids = BTreeSet::new();
    collect_extension_relationship_ids(&root, conformance.rel(), &mut ids)?;
    Ok(ids)
}

fn parse_document(xml: &[u8], max_bytes: usize) -> Result<Node> {
    parse_document_with_capabilities(xml, max_bytes, &MceCapabilities::ooxml_baseline())
}

fn parse_document_with_capabilities(
    xml: &[u8],
    max_bytes: usize,
    capabilities: &MceCapabilities,
) -> Result<Node> {
    if xml.len() > max_bytes {
        return Err(limit("input XML bytes"));
    }
    let limits = MceLimits {
        max_input_bytes: max_bytes,
        max_output_bytes: max_bytes,
        max_depth: MAX_DEPTH,
        max_namespace_bindings: 4096,
        max_directive_tokens: 4096,
        max_choices_per_alternate: 1024,
    };
    let processed = process_markup_compatibility(xml, capabilities, &limits)?.xml;
    let mut reader = NsReader::from_reader(processed.as_ref());
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut stack: Vec<Node> = Vec::new();
    let mut root = None;
    let mut nodes = 0usize;
    let mut strings = 0usize;
    loop {
        let event = reader.read_event_into(&mut buffer).map_err(xml_error)?;
        match event {
            Event::Start(ref element) | Event::Empty(ref element) => {
                nodes += 1;
                if nodes > MAX_NODES || stack.len() >= MAX_DEPTH {
                    return Err(limit("XML structure"));
                }
                let empty = matches!(&event, Event::Empty(_));
                let node = make_node(&reader, element, reader.decoder(), &mut strings)?;
                if empty {
                    attach(node, &mut stack, &mut root)?;
                } else {
                    stack.push(node);
                }
            },
            Event::End(_) => {
                let node = stack
                    .pop()
                    .ok_or_else(|| invalid("unexpected XML closing element"))?;
                attach(node, &mut stack, &mut root)?;
            },
            Event::Text(text) => {
                let decoded = text.decode().map_err(xml_error)?;
                let decoded = quick_xml::escape::unescape(&decoded).map_err(xml_error)?;
                add_strings(&mut strings, decoded.len())?;
                if let Some(node) = stack.last_mut() {
                    add_node_text(node, &decoded);
                } else if !decoded.trim().is_empty() {
                    return Err(invalid("text outside XML root"));
                }
            },
            Event::GeneralRef(reference) => {
                let name = reference.decode().map_err(xml_error)?;
                let value = reference
                    .resolve_char_ref()
                    .map_err(xml_error)?
                    .map(|v| v.to_string())
                    .or_else(|| match name.as_ref() {
                        "amp" => Some("&".into()),
                        "lt" => Some("<".into()),
                        "gt" => Some(">".into()),
                        "apos" => Some("'".into()),
                        "quot" => Some("\"".into()),
                        _ => None,
                    })
                    .ok_or_else(|| invalid("custom XML entity is rejected"))?;
                add_strings(&mut strings, value.len())?;
                if let Some(node) = stack.last_mut() {
                    add_node_text(node, &value);
                }
            },
            Event::CData(_) => return Err(invalid("CDATA is rejected")),
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("DTDs and processing instructions are rejected"));
            },
            Event::Decl(_) | Event::Comment(_) => {},
            Event::Eof => break,
        };
        buffer.clear();
    }
    if !stack.is_empty() {
        return Err(invalid("unterminated XML"));
    }
    root.ok_or_else(|| invalid("missing XML root"))
}

fn make_node(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
    strings: &mut usize,
) -> Result<Node> {
    let namespace = resolved(reader.resolver().resolve_element(element.name()).0)?;
    let name = std::str::from_utf8(element.local_name().as_ref())
        .map_err(xml_error)?
        .to_owned();
    add_strings(strings, namespace.len() + name.len())?;
    let mut attributes = Vec::new();
    for item in element.attributes().with_checks(true) {
        let item = item.map_err(xml_error)?;
        let qname = item.key.as_ref();
        if qname == b"xmlns" || qname.starts_with(b"xmlns:") {
            continue;
        }
        let (namespace, local) = reader.resolver().resolve_attribute(item.key);
        let namespace = resolved(namespace)?;
        let name = std::str::from_utf8(local.as_ref())
            .map_err(xml_error)?
            .to_owned();
        let value = item
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(xml_error)?
            .into_owned();
        add_strings(strings, namespace.len() + name.len() + value.len())?;
        if attributes
            .iter()
            .any(|a: &Attribute| a.namespace == namespace && a.name == name)
        {
            return Err(invalid("duplicate expanded XML attribute"));
        }
        attributes.push(Attribute {
            namespace,
            name,
            value,
        });
    }
    Ok(Node {
        namespace,
        name,
        attributes,
        children: Vec::new(),
        text: String::new(),
        content: Vec::new(),
    })
}

fn add_node_text(node: &mut Node, value: &str) {
    node.text.push_str(value);
    match node.content.last_mut() {
        Some(NodeContent::Text(current)) => current.push_str(value),
        _ => node.content.push(NodeContent::Text(value.to_owned())),
    }
}
fn attach(node: Node, stack: &mut [Node], root: &mut Option<Node>) -> Result<()> {
    if let Some(parent) = stack.last_mut() {
        let index = parent.children.len();
        parent.children.push(node);
        parent.content.push(NodeContent::Child(index));
    } else if root.replace(node).is_some() {
        return Err(invalid("multiple XML roots"));
    }
    Ok(())
}
fn root_conformance(root: &Node, name: &str) -> Result<ChartSheetConformance> {
    if root.name != name {
        return Err(invalid(format!("expected {name} root")));
    }
    match root.namespace.as_str() {
        SML => Ok(ChartSheetConformance::Transitional),
        STRICT_SML => Ok(ChartSheetConformance::Strict),
        _ => Err(invalid("unsupported SpreadsheetML namespace")),
    }
}
fn one_child<'a>(node: &'a Node, namespace: &str, name: &str) -> Result<Option<&'a Node>> {
    let mut values = node
        .children
        .iter()
        .filter(|c| c.namespace == namespace && c.name == name);
    let value = values.next();
    if values.next().is_some() {
        Err(invalid(format!(
            "{} has multiple {name} children",
            node.name
        )))
    } else {
        Ok(value)
    }
}
fn one_child_any_core<'a>(node: &'a Node, name: &str) -> Result<Option<&'a Node>> {
    let mut values = node
        .children
        .iter()
        .filter(|c| is_core(&c.namespace) && c.name == name);
    let value = values.next();
    if values.next().is_some() {
        Err(invalid(format!(
            "{} has multiple {name} children",
            node.name
        )))
    } else {
        Ok(value)
    }
}
fn required_child<'a>(node: &'a Node, namespace: &str, name: &str) -> Result<&'a Node> {
    one_child(node, namespace, name)?
        .ok_or_else(|| invalid(format!("{} is missing {name}", node.name)))
}
fn optional<'a>(node: &'a Node, namespace: &str, name: &str) -> Option<&'a str> {
    node.attributes
        .iter()
        .find(|a| a.namespace == namespace && a.name == name)
        .map(|a| a.value.as_str())
}
fn required<'a>(node: &'a Node, namespace: &str, name: &str) -> Result<&'a str> {
    optional(node, namespace, name)
        .filter(|v| !v.is_empty())
        .ok_or_else(|| invalid(format!("{} is missing attribute '{name}'", node.name)))
}
fn no_attributes(node: &Node, allowed: &[(&str, &str)]) -> Result<()> {
    if let Some(attribute) = node
        .attributes
        .iter()
        .find(|a| !allowed.contains(&(a.namespace.as_str(), a.name.as_str())))
    {
        Err(invalid(format!(
            "unexpected attribute '{}' on {}",
            attribute.name, node.name
        )))
    } else {
        Ok(())
    }
}
fn whitespace(node: &Node) -> Result<()> {
    if node.text.trim().is_empty() {
        Ok(())
    } else {
        Err(invalid(format!("unexpected text in {}", node.name)))
    }
}
fn leaf(node: &Node, label: &str) -> Result<()> {
    if node.children.is_empty() {
        Ok(())
    } else {
        Err(invalid(format!("{label} must not contain child elements")))
    }
}
fn bool_optional(node: &Node, name: &str) -> Result<Option<bool>> {
    optional(node, "", name)
        .map(|v| parse_bool(v, name))
        .transpose()
}
fn u32_optional(node: &Node, name: &str) -> Result<Option<u32>> {
    optional(node, "", name)
        .map(|v| v.parse().map_err(|_| invalid(format!("invalid {name}"))))
        .transpose()
}
fn parse_bool(value: &str, name: &str) -> Result<bool> {
    match value {
        "1" | "true" => Ok(true),
        "0" | "false" => Ok(false),
        _ => Err(invalid(format!("invalid boolean '{value}' for {name}"))),
    }
}
fn parse_orientation(value: &str) -> Result<PageOrientation> {
    match value {
        "default" => Ok(PageOrientation::Default),
        "portrait" => Ok(PageOrientation::Portrait),
        "landscape" => Ok(PageOrientation::Landscape),
        _ => Err(invalid("invalid chartsheet page orientation")),
    }
}
fn parse_state(value: &str) -> Result<ChartSheetState> {
    match value {
        "visible" => Ok(ChartSheetState::Visible),
        "hidden" => Ok(ChartSheetState::Hidden),
        "veryHidden" => Ok(ChartSheetState::VeryHidden),
        _ => Err(invalid("invalid workbook sheet state")),
    }
}
fn validate_guid(value: &str) -> Result<()> {
    let bytes = value.as_bytes();
    if bytes.len() != 38
        || bytes[0] != b'{'
        || bytes[37] != b'}'
        || ![9, 14, 19, 24].iter().all(|index| bytes[*index] == b'-')
        || bytes[1..37].iter().enumerate().any(|(index, byte)| {
            !matches!(index + 1, 9 | 14 | 19 | 24) && !byte.is_ascii_hexdigit()
        })
    {
        return Err(invalid(format!(
            "invalid custom chartsheet view GUID '{value}'"
        )));
    }
    Ok(())
}
fn is_core(value: &str) -> bool {
    matches!(value, SML | STRICT_SML)
}
fn validate_id(value: &str) -> Result<()> {
    let mut bytes = value.bytes();
    let Some(first) = bytes.next() else {
        return Err(invalid("relationship ID cannot be empty"));
    };
    if !(first.is_ascii_alphabetic() || first == b'_')
        || !bytes.all(|b| b.is_ascii_alphanumeric() || matches!(b, b'_' | b'-' | b'.'))
    {
        Err(invalid(format!("invalid relationship ID '{value}'")))
    } else {
        Ok(())
    }
}
fn bounded(value: &str) -> Result<()> {
    if value.len() <= MAX_STRING_BYTES {
        Ok(())
    } else {
        Err(limit("string bytes"))
    }
}
fn add_strings(total: &mut usize, size: usize) -> Result<()> {
    *total = total
        .checked_add(size)
        .ok_or_else(|| limit("XML string bytes"))?;
    if *total > MAX_STRING_BYTES {
        Err(limit("XML string bytes"))
    } else {
        Ok(())
    }
}
fn resolved(value: ResolveResult<'_>) -> Result<String> {
    match value {
        ResolveResult::Bound(Namespace(value)) => {
            Ok(std::str::from_utf8(value).map_err(xml_error)?.to_owned())
        },
        ResolveResult::Unbound => Ok(String::new()),
        ResolveResult::Unknown(prefix) => Err(invalid(format!(
            "unbound XML prefix '{}'",
            String::from_utf8_lossy(prefix.as_ref())
        ))),
    }
}
fn internal_relationship<'a>(
    part: &'a dyn Part,
    id: &str,
    kind: &str,
) -> Result<&'a litchi_opc::Relationship> {
    let relationship = part
        .rels()
        .get(id)
        .ok_or_else(|| invalid(format!("missing relationship '{id}'")))?;
    if relationship.reltype() != kind {
        return Err(invalid(format!("relationship '{id}' has unexpected type")));
    }
    if relationship.is_external() {
        return Err(invalid(format!(
            "external relationship '{id}' is not loaded"
        )));
    }
    Ok(relationship)
}
fn require_workbook(part: &dyn Part) -> Result<()> {
    if matches!(
        part.content_type(),
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
            | "application/vnd.ms-excel.sheet.macroEnabled.main+xml"
            | "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml"
            | "application/vnd.ms-excel.template.macroEnabled.main+xml"
    ) {
        Ok(())
    } else {
        Err(invalid("source part is not a workbook"))
    }
}
fn require_content_type(part: &dyn Part, expected: &str, label: &str) -> Result<()> {
    if part.content_type() == expected {
        Ok(())
    } else {
        Err(invalid(format!(
            "{label} part has content type '{}'",
            part.content_type()
        )))
    }
}
fn new_uri(package: &OpcPackage, value: &str, prefix: &str) -> Result<PackURI> {
    let uri = PackURI::new(value).map_err(OoxmlError::InvalidUri)?;
    if !uri.as_str().starts_with(prefix) {
        return Err(invalid(format!("part '{uri}' is outside {prefix}")));
    }
    if package.iter_parts().any(|part| part.partname() == &uri) {
        return Err(invalid(format!("part '{uri}' already exists")));
    }
    Ok(uri)
}
fn add_resource(total: &mut usize, size: usize, individual: usize, name: &str) -> Result<()> {
    if size > individual {
        return Err(limit(name));
    }
    *total = total
        .checked_add(size)
        .ok_or_else(|| limit("total resource bytes"))?;
    if *total > MAX_TOTAL_RESOURCE_BYTES {
        Err(limit("total resource bytes"))
    } else {
        Ok(())
    }
}
fn bool_attr_opt(out: &mut Vec<u8>, name: &str, value: Option<bool>) {
    if let Some(value) = value {
        attr(out, name, if value { "1" } else { "0" });
    }
}
fn u32_attr_opt(out: &mut Vec<u8>, name: &str, value: Option<u32>) {
    if let Some(value) = value {
        attr(out, name, &value.to_string());
    }
}
fn attr_opt(out: &mut Vec<u8>, name: &str, value: Option<&str>) {
    if let Some(value) = value {
        attr(out, name, value);
    }
}
fn attr(out: &mut Vec<u8>, name: &str, value: &str) {
    out.push(b' ');
    out.extend_from_slice(name.as_bytes());
    out.extend_from_slice(b"=\"");
    escape(out, value);
    out.push(b'\"');
}
fn escape(out: &mut Vec<u8>, value: &str) {
    for c in value.chars() {
        match c {
            '&' => out.extend_from_slice(b"&amp;"),
            '<' => out.extend_from_slice(b"&lt;"),
            '"' => out.extend_from_slice(b"&quot;"),
            '\t' => out.extend_from_slice(b"&#x9;"),
            '\n' => out.extend_from_slice(b"&#xA;"),
            '\r' => out.extend_from_slice(b"&#xD;"),
            _ => {
                let mut bytes = [0; 4];
                out.extend_from_slice(c.encode_utf8(&mut bytes).as_bytes());
            },
        }
    }
}
fn escape_text(out: &mut Vec<u8>, value: &str) {
    for c in value.chars() {
        match c {
            '&' => out.extend_from_slice(b"&amp;"),
            '<' => out.extend_from_slice(b"&lt;"),
            '>' => out.extend_from_slice(b"&gt;"),
            _ => {
                let mut bytes = [0; 4];
                out.extend_from_slice(c.encode_utf8(&mut bytes).as_bytes());
            },
        }
    }
}
fn xml_error(error: impl std::fmt::Display) -> OoxmlError {
    OoxmlError::Xml(error.to_string())
}
fn invalid(message: impl Into<String>) -> OoxmlError {
    OoxmlError::InvalidFormat(message.into())
}
fn limit(name: &str) -> OoxmlError {
    invalid(format!("chartsheet {name} limit exceeded"))
}

#[cfg(test)]
mod tests {
    use super::*;
    const POI_ONE: &[u8] =
        include_bytes!("../../../../test-data/poi/test-data/spreadsheet/WithChartSheet.xlsx");
    const POI_TWO: &[u8] =
        include_bytes!("../../../../test-data/poi/test-data/spreadsheet/chart_sheet.xlsx");
    const LO_CHART_EX: &[u8] = include_bytes!(
        "../../../../test-data/libreoffice-core/chart2/qa/extras/data/xlsx/boxWhisker.xlsx"
    );
    pub(super) const LO_USER_SHAPES_IMAGES: &[u8] = include_bytes!(
        "../../../../test-data/libreoffice-core/chart2/qa/extras/data/xlsx/tdf143127.xlsx"
    );
    fn sheet() -> ChartSheet {
        ChartSheet {
            properties: Some(ChartSheetProperties {
                published: Some(true),
                code_name: Some("ChartCode".into()),
                tab_color: Some(ChartSheetColor {
                    automatic: None,
                    indexed: None,
                    rgb: Some("FF336699".into()),
                    theme: None,
                    tint: Some(0.25),
                }),
            }),
            views: vec![ChartSheetView {
                tab_selected: Some(true),
                zoom_scale: Some(125),
                workbook_view_id: 0,
                zoom_to_fit: Some(false),
            }],
            protection: Some(ChartSheetProtection {
                password_hash: Some("ABCD".into()),
                content: Some(true),
                objects: Some(false),
            }),
            custom_views: Some(vec![
                ChartSheetCustomView {
                    guid: "{00112233-4455-6677-8899-AABBCCDDEEFF}".into(),
                    scale: Some(175),
                    state: Some(ChartSheetState::Hidden),
                    zoom_to_fit: Some(true),
                },
                ChartSheetCustomView {
                    guid: "{10213243-5465-7687-98A9-BACBDCEDFE0F}".into(),
                    scale: None,
                    state: None,
                    zoom_to_fit: Some(false),
                },
            ]),
            margins: Some(ChartSheetMargins {
                left: 0.7,
                right: 0.7,
                top: 0.75,
                bottom: 0.75,
                header: 0.3,
                footer: 0.3,
            }),
            page_setup: Some(ChartSheetPageSetup {
                paper_size: Some(1),
                first_page_number: Some(1),
                orientation: Some(PageOrientation::Landscape),
                use_printer_defaults: Some(true),
                black_and_white: Some(false),
                draft: Some(false),
                use_first_page_number: Some(true),
                horizontal_dpi: Some(600),
                vertical_dpi: Some(600),
                copies: Some(1),
                printer_settings_relationship_id: Some("rIdPrinter".into()),
            }),
            header_footer: Some(ChartSheetHeaderFooter {
                align_with_margins: Some(false),
                odd_header: Some("&CChart & Report".into()),
                ..Default::default()
            }),
            drawing_relationship_id: "rIdDrawing".into(),
            legacy_drawing_relationship_id: Some("rIdLegacy".into()),
            legacy_header_footer_drawing_relationship_id: Some("rIdLegacyHF".into()),
            background_picture_relationship_id: Some("rIdBackground".into()),
            web_publish_items: Some(ChartSheetWebPublishItems {
                count: Some(2),
                items: vec![
                    ChartSheetWebPublishItem {
                        id: 11289,
                        div_id: "Views_11289".into(),
                        source_type: ChartSheetWebSourceType::Range,
                        source_ref: Some("A6:C6".into()),
                        source_object: None,
                        destination_file: "file:///definitely/not/accessed/Publish.htm".into(),
                        title: Some("Range & title".into()),
                        auto_republish: Some(false),
                    },
                    ChartSheetWebPublishItem {
                        id: 6433,
                        div_id: "Views_6433".into(),
                        source_type: ChartSheetWebSourceType::Chart,
                        source_ref: None,
                        source_object: Some("https://example.invalid/Chart 1".into()),
                        destination_file: "https://example.invalid/Publish.mht".into(),
                        title: None,
                        auto_republish: None,
                    },
                ],
            }),
            extension_list: None,
        }
    }
    fn drawing(conformance: ChartSheetConformance) -> Vec<u8> {
        format!("<xdr:wsDr xmlns:xdr=\"{}\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><xdr:absoluteAnchor><a:graphic><a:graphicData><c:chart xmlns:c=\"{}\" xmlns:r=\"{}\" r:id=\"rIdChart\"/></a:graphicData></a:graphic></xdr:absoluteAnchor></xdr:wsDr>", conformance.xdr(), conformance.chart(), conformance.rel()).into_bytes()
    }
    fn chart(conformance: ChartSheetConformance) -> Vec<u8> {
        format!(
            "<c:chartSpace xmlns:c=\"{}\"><c:chart/></c:chartSpace>",
            conformance.chart()
        )
        .into_bytes()
    }
    fn vml(id: &str, name: &str) -> ChartSheetVmlDrawingResource {
        ChartSheetVmlDrawingResource{relationship_id:id.into(),part_name:format!("/xl/drawings/{name}.vml"),content_type:VML_DRAWING_CT.into(),data:format!("<xml xmlns:v=\"urn:schemas-microsoft-com:vml\"><v:shape href=\"https://example.invalid/{name}\"/></xml>").into_bytes()}
    }
    fn value(conformance: ChartSheetConformance) -> ChartSheetPackage {
        ChartSheetPackage {
            entry: ChartSheetEntry {
                name: "Chart 1".into(),
                sheet_id: 2,
                state: ChartSheetState::Visible,
                workbook_relationship_id: "rIdChartSheet".into(),
                part_name: "/xl/chartsheets/sheet1.xml".into(),
            },
            chartsheet: sheet(),
            drawing: ChartSheetDrawingResource {
                part_name: "/xl/drawings/drawing1.xml".into(),
                content_type: DRAWING_CT.into(),
                data: drawing(conformance),
                charts: vec![ChartSheetChartResource {
                    relationship_id: "rIdChart".into(),
                    part_name: "/xl/charts/chart1.xml".into(),
                    content_type: CHART_CT.into(),
                    data: chart(conformance),
                    kind: ChartSheetChartResourceKind::Classic,
                }],
            },
            legacy_drawing: Some(vml("rIdLegacy", "vmlDrawing1")),
            legacy_header_footer_drawing: Some(vml("rIdLegacyHF", "vmlDrawing2")),
            background_picture: Some(ChartSheetBackgroundPicture {
                relationship_id: "rIdBackground".into(),
                part_name: "/xl/media/background1.png".into(),
                content_type: ChartSheetImageContentType::Png,
                data: vec![0, 255, 1, 254],
            }),
            printer_settings: Some(ChartSheetPrinterSettings {
                relationship_id: "rIdPrinter".into(),
                resource: PrinterSettingsResource {
                    part_name: "/xl/printerSettings/printerSettings1.bin".into(),
                    data: vec![0x44, 0x45, 0x56, 0x4d, 0x4f, 0x44, 0x45, 0, 255],
                },
            }),
            extension_relationships: vec![],
        }
    }
    pub(super) fn base_package(conformance: ChartSheetConformance) -> (OpcPackage, PackURI) {
        let mut package = OpcPackage::new();
        let uri = PackURI::new("/xl/workbook.xml").unwrap();
        let xml = format!(
            "<x:workbook xmlns:x=\"{}\" xmlns:r=\"{}\"><x:sheets><x:sheet name=\"Data\" sheetId=\"1\" r:id=\"rIdData\"/></x:sheets></x:workbook>",
            conformance.sml(),
            conformance.rel()
        );
        package.add_part(Box::new(BlobPart::new(
            uri.clone(),
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml".into(),
            xml.into_bytes(),
        )));
        (package, uri)
    }

    fn ext(uri: &str, payload: &str) -> ChartSheetExtension {
        ChartSheetExtension {
            uri: uri.into(),
            payload_xml: payload.as_bytes().to_vec(),
        }
    }
    fn companion(
        id: &str,
        path: &str,
        content_type: &str,
        data: &[u8],
    ) -> ChartSheetChartCompanionResource {
        ChartSheetChartCompanionResource {
            relationship_id: id.into(),
            part_name: path.into(),
            content_type: content_type.into(),
            data: data.to_vec(),
        }
    }

    #[test]
    fn real_libreoffice_chart_ex_fixture_round_trips_as_inert_resources() {
        let source = OpcPackage::from_bytes(LO_CHART_EX).unwrap();
        let blob = |path: &str| {
            source
                .get_part(&PackURI::new(path).unwrap())
                .unwrap()
                .blob()
                .to_vec()
        };
        let mut expected = value(ChartSheetConformance::Transitional);
        expected.drawing.data = blob("/xl/drawings/drawing1.xml");
        expected.drawing.charts = vec![ChartSheetChartResource {
            relationship_id: "rId1".into(),
            part_name: "/xl/charts/chartEx1.xml".into(),
            content_type: CHART_EX_CT.into(),
            data: blob("/xl/charts/chartEx1.xml"),
            kind: ChartSheetChartResourceKind::Extended {
                styles: vec![companion(
                    "rId1",
                    "/xl/charts/style1.xml",
                    CHART_STYLE_CT,
                    &blob("/xl/charts/style1.xml"),
                )],
                color_styles: vec![companion(
                    "rId2",
                    "/xl/charts/colors1.xml",
                    CHART_COLOR_STYLE_CT,
                    &blob("/xl/charts/colors1.xml"),
                )],
                user_shapes: None,
                outbound_resources: vec![],
            },
        }];
        let (mut package, workbook) = base_package(ChartSheetConformance::Transitional);
        store_chartsheet(
            &mut package,
            &workbook,
            &expected,
            ChartSheetConformance::Transitional,
        )
        .unwrap();
        let loaded = load_chartsheet(&package, &workbook, "rIdChartSheet").unwrap();
        assert_eq!(loaded, expected);
        assert!(matches!(
            loaded.drawing.charts[0].kind,
            ChartSheetChartResourceKind::Extended { .. }
        ));
    }

    #[test]
    fn chart_ex_strict_mce_selects_extended_choice_and_preserves_classic_fallback_behavior() {
        let strict = format!(
            "<xdr:wsDr xmlns:xdr=\"{STRICT_XDR}\" xmlns:a=\"{STRICT_DRAWING_MAIN}\" xmlns:c=\"{STRICT_CHART}\" xmlns:r=\"{STRICT_REL}\" xmlns:cx=\"{CHART_EX}\" xmlns:cx1=\"{CHART_EX_CHOICE}\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\"><mc:AlternateContent><mc:Choice Requires=\"cx1\"><a:graphic><a:graphicData uri=\"{CHART_EX}\"><cx:chart r:id=\"rIdExtended\"/></a:graphicData></a:graphic></mc:Choice><mc:Fallback><a:graphic><a:graphicData uri=\"{STRICT_CHART}\"><c:chart r:id=\"rIdClassic\"/></a:graphicData></a:graphic></mc:Fallback></mc:AlternateContent></xdr:wsDr>"
        );
        let references =
            drawing_chart_references(strict.as_bytes(), ChartSheetConformance::Strict).unwrap();
        assert_eq!(
            references,
            vec![DrawingChartReference {
                relationship_id: "rIdExtended".into(),
                kind: DrawingChartKind::Extended
            }]
        );
        let fallback = strict.replace(
            &format!("xmlns:cx1=\"{CHART_EX_CHOICE}\""),
            "xmlns:cx1=\"urn:unsupported-chart-version\"",
        );
        let references =
            drawing_chart_references(fallback.as_bytes(), ChartSheetConformance::Strict).unwrap();
        assert_eq!(
            references,
            vec![DrawingChartReference {
                relationship_id: "rIdClassic".into(),
                kind: DrawingChartKind::Classic
            }]
        );
    }

    #[test]
    fn rejects_chart_ex_drawing_shape_roots_cardinality_relationships_and_caps() {
        let drawing = |body: &str| {
            format!(
                "<xdr:wsDr xmlns:xdr=\"{XDR}\" xmlns:a=\"{DRAWING_MAIN}\" xmlns:r=\"{REL}\" xmlns:cx=\"{CHART_EX}\">{body}</xdr:wsDr>"
            )
        };
        for body in [
            "<a:graphicData uri=\"urn:wrong\"><cx:chart r:id=\"rId1\"/></a:graphicData>"
                .to_string(),
            format!("<a:graphicData uri=\"{CHART_EX}\"><cx:chart/></a:graphicData>"),
            format!(
                "<a:graphicData uri=\"{CHART_EX}\"><cx:chart r:id=\"rId1\"/><cx:chart r:id=\"rId2\"/></a:graphicData>"
            ),
            format!("<a:graphicData uri=\"{CHART_EX}\"><cx:wrong r:id=\"rId1\"/></a:graphicData>"),
            format!(
                "<a:graphicData uri=\"{CHART_EX}\" bad=\"1\"><cx:chart r:id=\"rId1\"/></a:graphicData>"
            ),
        ] {
            assert!(
                drawing_chart_references(
                    drawing(&body).as_bytes(),
                    ChartSheetConformance::Transitional
                )
                .is_err(),
                "accepted {body}"
            );
        }
        assert!(
            validate_chart_ex_relationships(
                b"<cx:chartSpace xmlns:cx=\"urn:wrong\"/>",
                ChartSheetConformance::Transitional
            )
            .is_err()
        );
        assert!(validate_chart_companion_xml(b"<cs:colorStyle xmlns:cs=\"http://schemas.microsoft.com/office/drawing/2012/chartStyle\"/>","chartStyle",MAX_CHART_STYLE_BYTES).is_err());
        assert!(
            validate_chart_ex_relationships(
                &[b' '; MAX_CHART_EX_BYTES + 1],
                ChartSheetConformance::Transitional
            )
            .is_err()
        );
        let mut bad = value(ChartSheetConformance::Transitional);
        bad.drawing.data = drawing(&format!(
            "<a:graphicData uri=\"{CHART_EX}\"><cx:chart r:id=\"rIdEx\"/></a:graphicData>"
        ))
        .into_bytes();
        bad.drawing.charts = vec![ChartSheetChartResource {
            relationship_id: "rIdEx".into(),
            part_name: "/xl/charts/chartEx1.xml".into(),
            content_type: CHART_EX_CT.into(),
            data: format!("<cx:chartSpace xmlns:cx=\"{CHART_EX}\"/>").into_bytes(),
            kind: ChartSheetChartResourceKind::Extended {
                styles: vec![companion(
                    "rIdSame",
                    "/xl/charts/style1.xml",
                    CHART_STYLE_CT,
                    &format!("<cs:chartStyle xmlns:cs=\"{CHART_STYLE}\"/>").into_bytes(),
                )],
                color_styles: vec![companion(
                    "rIdSame",
                    "/xl/charts/colors1.xml",
                    CHART_COLOR_STYLE_CT,
                    &format!("<cs:colorStyle xmlns:cs=\"{CHART_STYLE}\"/>").into_bytes(),
                )],
                user_shapes: None,
                outbound_resources: vec![],
            },
        }];
        let (mut package, workbook) = base_package(ChartSheetConformance::Transitional);
        assert!(
            store_chartsheet(
                &mut package,
                &workbook,
                &bad,
                ChartSheetConformance::Transitional
            )
            .is_err()
        );
        assert!(
            package
                .get_part(&PackURI::new("/xl/chartsheets/sheet1.xml").unwrap())
                .is_err()
        );
    }

    #[test]
    fn rejects_chart_ex_wrong_type_orphan_escape_outbound_and_companion_graphs() {
        let conformance = ChartSheetConformance::Transitional;
        let source = OpcPackage::from_bytes(LO_CHART_EX).unwrap();
        let blob = |path: &str| {
            source
                .get_part(&PackURI::new(path).unwrap())
                .unwrap()
                .blob()
                .to_vec()
        };
        let mut expected = value(conformance);
        expected.drawing.data = blob("/xl/drawings/drawing1.xml");
        expected.drawing.charts = vec![ChartSheetChartResource {
            relationship_id: "rId1".into(),
            part_name: "/xl/charts/chartEx1.xml".into(),
            content_type: CHART_EX_CT.into(),
            data: blob("/xl/charts/chartEx1.xml"),
            kind: ChartSheetChartResourceKind::Extended {
                styles: vec![companion(
                    "rId1",
                    "/xl/charts/style1.xml",
                    CHART_STYLE_CT,
                    &blob("/xl/charts/style1.xml"),
                )],
                color_styles: vec![companion(
                    "rId2",
                    "/xl/charts/colors1.xml",
                    CHART_COLOR_STYLE_CT,
                    &blob("/xl/charts/colors1.xml"),
                )],
                user_shapes: None,
                outbound_resources: vec![],
            },
        }];
        for (kind, target, external) in [
            (rt::CHART, "../charts/chartEx1.xml", false),
            (CHART_EX_REL, "../../../evil.xml", false),
            (CHART_EX_REL, "https://example.invalid/chartEx.xml", true),
        ] {
            let (mut package, workbook) = base_package(conformance);
            store_chartsheet(&mut package, &workbook, &expected, conformance).unwrap();
            let drawing = package
                .get_part_mut(&PackURI::new("/xl/drawings/drawing1.xml").unwrap())
                .unwrap();
            drawing.rels_mut().remove("rId1");
            drawing.rels_mut().add_relationship(
                kind.into(),
                target.into(),
                "rId1".into(),
                external,
            );
            assert!(
                load_chartsheet(&package, &workbook, "rIdChartSheet").is_err(),
                "accepted {kind} {target}"
            );
        }
        let (mut package, workbook) = base_package(conformance);
        store_chartsheet(&mut package, &workbook, &expected, conformance).unwrap();
        package
            .get_part_mut(&PackURI::new("/xl/drawings/drawing1.xml").unwrap())
            .unwrap()
            .rels_mut()
            .add_relationship(
                CHART_EX_REL.into(),
                "../charts/chartEx1.xml".into(),
                "rIdOrphan".into(),
                false,
            );
        assert!(load_chartsheet(&package, &workbook, "rIdChartSheet").is_err());
        let (mut package, workbook) = base_package(conformance);
        store_chartsheet(&mut package, &workbook, &expected, conformance).unwrap();
        package
            .get_part_mut(&PackURI::new("/xl/charts/chartEx1.xml").unwrap())
            .unwrap()
            .rels_mut()
            .add_relationship(
                IMAGE_REL.into(),
                "../media/image1.png".into(),
                "rIdImage".into(),
                false,
            );
        assert!(load_chartsheet(&package, &workbook, "rIdChartSheet").is_err());
        let (mut package, workbook) = base_package(conformance);
        store_chartsheet(&mut package, &workbook, &expected, conformance).unwrap();
        package
            .get_part_mut(&PackURI::new("/xl/charts/style1.xml").unwrap())
            .unwrap()
            .rels_mut()
            .add_relationship(
                IMAGE_REL.into(),
                "../media/image1.png".into(),
                "rIdOutbound".into(),
                false,
            );
        assert!(load_chartsheet(&package, &workbook, "rIdChartSheet").is_err());
    }

    pub(super) fn chart_ex_user_shapes(conformance: ChartSheetConformance) -> ChartSheetPackage {
        let source = OpcPackage::from_bytes(LO_USER_SHAPES_IMAGES).unwrap();
        let blob = |path: &str| {
            source
                .get_part(&PackURI::new(path).unwrap())
                .unwrap()
                .blob()
                .to_vec()
        };
        let mut value = value(conformance);
        value.drawing.data=format!("<xdr:wsDr xmlns:xdr=\"{}\" xmlns:a=\"{}\" xmlns:r=\"{}\" xmlns:cx=\"{CHART_EX}\"><a:graphic><a:graphicData uri=\"{CHART_EX}\"><cx:chart r:id=\"rIdChartEx\"/></a:graphicData></a:graphic></xdr:wsDr>",conformance.xdr(),if conformance==ChartSheetConformance::Strict{STRICT_DRAWING_MAIN}else{DRAWING_MAIN},conformance.rel()).into_bytes();
        let user_shapes_data = if conformance == ChartSheetConformance::Transitional {
            blob("/xl/drawings/drawing2.xml")
        } else {
            String::from_utf8(blob("/xl/drawings/drawing2.xml"))
                .unwrap()
                .replace(CHART, STRICT_CHART)
                .replace(
                    "http://schemas.openxmlformats.org/drawingml/2006/chartDrawing",
                    "http://purl.oclc.org/ooxml/drawingml/chartDrawing",
                )
                .replace(DRAWING_MAIN, STRICT_DRAWING_MAIN)
                .replace(REL, STRICT_REL)
                .into_bytes()
        };
        value.drawing.charts = vec![ChartSheetChartResource {
            relationship_id: "rIdChartEx".into(),
            part_name: "/xl/charts/chartEx1.xml".into(),
            content_type: CHART_EX_CT.into(),
            data: format!("<cx:chartSpace xmlns:cx=\"{CHART_EX}\"/>").into_bytes(),
            kind: ChartSheetChartResourceKind::Extended {
                styles: vec![],
                color_styles: vec![],
                user_shapes: Some(ChartSheetChartUserShapesResource {
                    relationship_id: "rIdUserShapes".into(),
                    part_name: "/xl/drawings/chartDrawing1.xml".into(),
                    content_type: CHART_USER_SHAPES_CT.into(),
                    data: user_shapes_data,
                    images: vec![
                        ChartSheetChartUserShapeImageResource {
                            relationship_id: "rId1".into(),
                            part_name: "/xl/media/image1.png".into(),
                            content_type: ChartSheetChartUserShapeImageContentType::Png,
                            data: blob("/xl/media/image1.png"),
                        },
                        ChartSheetChartUserShapeImageResource {
                            relationship_id: "rId2".into(),
                            part_name: "/xl/media/image2.svg".into(),
                            content_type: ChartSheetChartUserShapeImageContentType::Svg,
                            data: blob("/xl/media/image2.svg"),
                        },
                    ],
                }),
                outbound_resources: vec![],
            },
        }];
        value
    }
    #[test]
    fn chart_ex_user_shapes_png_svg_transitional_and_strict_round_trip() {
        for conformance in [
            ChartSheetConformance::Transitional,
            ChartSheetConformance::Strict,
        ] {
            let expected = chart_ex_user_shapes(conformance);
            let (mut package, workbook) = base_package(conformance);
            store_chartsheet(&mut package, &workbook, &expected, conformance).unwrap();
            let loaded = load_chartsheet(&package, &workbook, "rIdChartSheet").unwrap();
            assert_eq!(loaded, expected);
            let chart = package
                .get_part(&PackURI::new("/xl/charts/chartEx1.xml").unwrap())
                .unwrap();
            assert_eq!(
                chart.rels().get("rIdUserShapes").unwrap().reltype(),
                conformance.chart_user_shapes_rel()
            );
            let shapes = package
                .get_part(&PackURI::new("/xl/drawings/chartDrawing1.xml").unwrap())
                .unwrap();
            assert!(
                shapes
                    .rels()
                    .iter()
                    .all(|relationship| relationship.reltype() == conformance.image_rel())
            );
        }
    }
    #[test]
    fn chart_ex_user_shapes_rejects_graph_mime_namespace_collision_and_caps() {
        let conformance = ChartSheetConformance::Transitional;
        let mut bad = chart_ex_user_shapes(conformance);
        if let ChartSheetChartResourceKind::Extended {
            user_shapes: Some(shapes),
            ..
        } = &mut bad.drawing.charts[0].kind
        {
            shapes.images[0].content_type = ChartSheetChartUserShapeImageContentType::Gif;
        }
        let (mut package, workbook) = base_package(conformance);
        assert!(store_chartsheet(&mut package, &workbook, &bad, conformance).is_err());
        let mut bad = chart_ex_user_shapes(conformance);
        if let ChartSheetChartResourceKind::Extended {
            user_shapes: Some(shapes),
            ..
        } = &mut bad.drawing.charts[0].kind
        {
            shapes.images.pop();
        }
        let (mut package, workbook) = base_package(conformance);
        assert!(store_chartsheet(&mut package, &workbook, &bad, conformance).is_err());
        let mut bad = chart_ex_user_shapes(ChartSheetConformance::Strict);
        if let ChartSheetChartResourceKind::Extended {
            user_shapes: Some(shapes),
            ..
        } = &mut bad.drawing.charts[0].kind
        {
            shapes.data = String::from_utf8(std::mem::take(&mut shapes.data))
                .unwrap()
                .replace(STRICT_CHART, CHART)
                .into_bytes();
        }
        let (mut package, workbook) = base_package(ChartSheetConformance::Strict);
        assert!(
            store_chartsheet(&mut package, &workbook, &bad, ChartSheetConformance::Strict).is_err()
        );
        let mut bad = chart_ex_user_shapes(conformance);
        if let ChartSheetChartResourceKind::Extended {
            styles,
            user_shapes: Some(shapes),
            ..
        } = &mut bad.drawing.charts[0].kind
        {
            styles.push(companion(
                &shapes.relationship_id,
                "/xl/charts/style1.xml",
                CHART_STYLE_CT,
                format!("<cs:chartStyle xmlns:cs=\"{CHART_STYLE}\"/>").as_bytes(),
            ));
        }
        let (mut package, workbook) = base_package(conformance);
        assert!(store_chartsheet(&mut package, &workbook, &bad, conformance).is_err());
        let mut bad = chart_ex_user_shapes(conformance);
        if let ChartSheetChartResourceKind::Extended {
            user_shapes: Some(shapes),
            ..
        } = &mut bad.drawing.charts[0].kind
        {
            shapes.images[0].data = vec![0; MAX_CHART_USER_SHAPE_IMAGE_BYTES + 1];
        }
        let (mut package, workbook) = base_package(conformance);
        assert!(store_chartsheet(&mut package, &workbook, &bad, conformance).is_err());
    }
    fn with_extension_relationships(conformance: ChartSheetConformance) -> ChartSheetPackage {
        let mut value = value(conformance);
        value.chartsheet.extension_list = Some(ChartSheetExtensionList {
            extensions: vec![
                ext(
                    "urn:duplicate",
                    &format!(
                        "<u:payload xmlns:u=\"urn:vendor\" xmlns:r=\"{}\" r:id=\"rIdExtInternal\">before<u:child/>after</u:payload>",
                        conformance.rel()
                    ),
                ),
                ext(
                    "urn:duplicate",
                    &format!(
                        "<v:external xmlns:v=\"urn:vendor-two\" xmlns:r=\"{}\" r:link=\"rIdExtExternal\"/>",
                        conformance.rel()
                    ),
                ),
            ],
        });
        value.extension_relationships = vec![
            ChartSheetExtensionRelationship {
                relationship_id: "rIdExtInternal".into(),
                relationship_type: "urn:relationship:internal".into(),
                target: ChartSheetExtensionRelationshipTarget::Internal {
                    part_name: "/xl/custom/ext.bin".into(),
                },
            },
            ChartSheetExtensionRelationship {
                relationship_id: "rIdExtExternal".into(),
                relationship_type: "urn:relationship:external".into(),
                target: ChartSheetExtensionRelationshipTarget::External {
                    target: "https://example.invalid/not-fetched".into(),
                },
            },
        ];
        value
            .extension_relationships
            .sort_by(|left, right| left.relationship_id.cmp(&right.relationship_id));
        let xml = write_chartsheet(&value.chartsheet, conformance).unwrap();
        value.chartsheet = parse_chartsheet(&xml).unwrap().1;
        value
    }

    #[test]
    fn ext_list_strict_mce_duplicate_uri_and_deterministic_round_trip() {
        let xml = format!(
            "<x:chartsheet xmlns:x=\"{STRICT_SML}\" xmlns:r=\"{STRICT_REL}\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:z=\"urn:unsupported-choice\"><x:sheetViews><x:sheetView workbookViewId=\"0\"/></x:sheetViews><x:drawing r:id=\"rIdD\"/><mc:AlternateContent><mc:Choice Requires=\"z\"><z:ignored/></mc:Choice><mc:Fallback><x:extLst><x:ext uri=\"urn:same\"><u:payload xmlns:u=\"urn:vendor\" r:id=\"rIdExt\">before<u:child a=\"1\"/>after</u:payload></x:ext><x:ext uri=\"urn:same\"><v:other xmlns:v=\"urn:vendor-two\"/></x:ext></x:extLst></mc:Fallback></mc:AlternateContent></x:chartsheet>"
        );
        let (kind, parsed) = parse_chartsheet(xml.as_bytes()).unwrap();
        assert_eq!(kind, ChartSheetConformance::Strict);
        let extensions = &parsed.extension_list.as_ref().unwrap().extensions;
        assert_eq!(extensions.len(), 2);
        assert_eq!(extensions[0].uri, extensions[1].uri);
        let payload = std::str::from_utf8(&extensions[0].payload_xml).unwrap();
        assert!(payload.contains("before<e0:child a=\"1\"/>after"));
        assert!(payload.contains("r:id=\"rIdExt\""));
        let first = write_chartsheet(&parsed, kind).unwrap();
        let reparsed = parse_chartsheet(&first).unwrap().1;
        let second = write_chartsheet(&reparsed, kind).unwrap();
        assert_eq!(first, second);
        assert_eq!(parsed, reparsed);
    }

    #[test]
    fn ext_list_package_round_trip_preserves_inert_internal_and_external_relationships() {
        let conformance = ChartSheetConformance::Transitional;
        let (mut package, workbook) = base_package(conformance);
        let expected = with_extension_relationships(conformance);
        store_chartsheet(&mut package, &workbook, &expected, conformance).unwrap();
        let loaded = load_chartsheet(&package, &workbook, "rIdChartSheet").unwrap();
        assert_eq!(loaded, expected);
        let part = package
            .get_part(&PackURI::new("/xl/chartsheets/sheet1.xml").unwrap())
            .unwrap();
        assert_eq!(
            part.rels()
                .get("rIdExtInternal")
                .unwrap()
                .target_partname()
                .unwrap()
                .as_str(),
            "/xl/custom/ext.bin"
        );
        assert!(part.rels().get("rIdExtExternal").unwrap().is_external());
        assert_eq!(
            part.rels().get("rIdExtExternal").unwrap().target_ref(),
            "https://example.invalid/not-fetched"
        );
    }

    #[test]
    fn rejects_ext_list_schema_order_uri_payload_and_caps() {
        let wrap = |body: &str| {
            format!(
                "<chartsheet xmlns=\"{SML}\" xmlns:r=\"{REL}\"><sheetViews><sheetView workbookViewId=\"0\"/></sheetViews><drawing r:id=\"rIdD\"/>{body}</chartsheet>"
            )
        };
        for body in [
            "<extLst/>",
            "<extLst bad=\"1\"><ext uri=\"u\"><a/></ext></extLst>",
            "<extLst><ext><a/></ext></extLst>",
            "<extLst><ext uri=\"\"><a/></ext></extLst>",
            "<extLst><ext uri=\"bad uri\"><a/></ext></extLst>",
            "<extLst><ext uri=\"u\" bad=\"1\"><a/></ext></extLst>",
            "<extLst><ext uri=\"u\"/></extLst>",
            "<extLst><ext uri=\"u\"><a/><b/></ext></extLst>",
            "<extLst><bad uri=\"u\"><a/></bad></extLst>",
            "<u:extLst xmlns:u=\"urn:foreign\"><u:ext uri=\"u\"><a/></u:ext></u:extLst>",
        ] {
            let xml = wrap(body);
            assert!(parse_chartsheet(xml.as_bytes()).is_err(), "accepted {body}");
        }
        let out_of_order = format!(
            "<chartsheet xmlns=\"{SML}\" xmlns:r=\"{REL}\"><sheetViews><sheetView workbookViewId=\"0\"/></sheetViews><extLst><ext uri=\"u\"><a/></ext></extLst><drawing r:id=\"rIdD\"/></chartsheet>"
        );
        assert!(parse_chartsheet(out_of_order.as_bytes()).is_err());
        let mut value = sheet();
        value.extension_list = Some(ChartSheetExtensionList {
            extensions: vec![ext("u", "<?run?><a/>")],
        });
        assert!(write_chartsheet(&value, ChartSheetConformance::Transitional).is_err());
        value.extension_list = Some(ChartSheetExtensionList {
            extensions: vec![ext(&"u".repeat(MAX_EXTENSION_URI_BYTES + 1), "<a/>")],
        });
        assert!(write_chartsheet(&value, ChartSheetConformance::Transitional).is_err());
        value.extension_list = Some(ChartSheetExtensionList {
            extensions: vec![ext(
                "u",
                &format!("<a>{}</a>", "x".repeat(MAX_EXTENSION_PAYLOAD_BYTES)),
            )],
        });
        assert!(write_chartsheet(&value, ChartSheetConformance::Transitional).is_err());
        value.extension_list = Some(ChartSheetExtensionList {
            extensions: vec![ext("u", "<a/>"); MAX_EXTENSIONS + 1],
        });
        assert!(write_chartsheet(&value, ChartSheetConformance::Transitional).is_err());
    }

    #[test]
    fn rejects_extension_relationship_missing_orphan_mismatch_duplicate_escape_caps_and_wrong_namespace()
     {
        let conformance = ChartSheetConformance::Transitional;
        let (mut package, workbook) = base_package(conformance);
        let mut missing = with_extension_relationships(conformance);
        missing.extension_relationships.clear();
        assert!(store_chartsheet(&mut package, &workbook, &missing, conformance).is_err());
        assert!(
            package
                .get_part(&PackURI::new("/xl/chartsheets/sheet1.xml").unwrap())
                .is_err()
        );
        let (mut package, workbook) = base_package(conformance);
        let mut mismatched = with_extension_relationships(conformance);
        mismatched.extension_relationships[0].relationship_id = "rIdWrong".into();
        assert!(store_chartsheet(&mut package, &workbook, &mismatched, conformance).is_err());
        let (mut package, workbook) = base_package(conformance);
        let mut duplicate = with_extension_relationships(conformance);
        duplicate.extension_relationships[1] = duplicate.extension_relationships[0].clone();
        assert!(store_chartsheet(&mut package, &workbook, &duplicate, conformance).is_err());
        let (mut package, workbook) = base_package(conformance);
        let mut escaped = with_extension_relationships(conformance);
        escaped.extension_relationships[0].target =
            ChartSheetExtensionRelationshipTarget::Internal {
                part_name: "../../../evil.bin".into(),
            };
        assert!(store_chartsheet(&mut package, &workbook, &escaped, conformance).is_err());
        let (mut package, workbook) = base_package(conformance);
        let mut oversized = with_extension_relationships(conformance);
        oversized.extension_relationships[0].relationship_type =
            "x".repeat(MAX_EXTENSION_RELATIONSHIP_STRING_BYTES + 1);
        assert!(store_chartsheet(&mut package, &workbook, &oversized, conformance).is_err());
        let (mut package, workbook) = base_package(ChartSheetConformance::Strict);
        let mut wrong_namespace = with_extension_relationships(ChartSheetConformance::Strict);
        wrong_namespace.chartsheet.extension_list = Some(ChartSheetExtensionList {
            extensions: vec![ext(
                "u",
                &format!("<u:a xmlns:u=\"urn:v\" xmlns:r=\"{REL}\" r:id=\"rIdExtInternal\"/>"),
            )],
        });
        wrong_namespace.extension_relationships.truncate(1);
        assert!(
            store_chartsheet(
                &mut package,
                &workbook,
                &wrong_namespace,
                ChartSheetConformance::Strict
            )
            .is_err()
        );
        let (mut package, workbook) = base_package(conformance);
        let expected = with_extension_relationships(conformance);
        store_chartsheet(&mut package, &workbook, &expected, conformance).unwrap();
        package
            .get_part_mut(&PackURI::new("/xl/chartsheets/sheet1.xml").unwrap())
            .unwrap()
            .rels_mut()
            .add_relationship(
                "urn:orphan".into(),
                "https://example.invalid/orphan".into(),
                "rIdOrphan".into(),
                true,
            );
        assert!(load_chartsheet(&package, &workbook, "rIdChartSheet").is_err());
        let (mut package, workbook) = base_package(conformance);
        store_chartsheet(&mut package, &workbook, &expected, conformance).unwrap();
        package
            .get_part_mut(&PackURI::new("/xl/chartsheets/sheet1.xml").unwrap())
            .unwrap()
            .rels_mut()
            .remove("rIdExtInternal");
        assert!(load_chartsheet(&package, &workbook, "rIdChartSheet").is_err());
    }

    #[test]
    fn strict_typed_xml_round_trip() {
        let expected = sheet();
        let xml = write_chartsheet(&expected, ChartSheetConformance::Strict).unwrap();
        let (kind, parsed) = parse_chartsheet(&xml).unwrap();
        assert_eq!(kind, ChartSheetConformance::Strict);
        assert_eq!(parsed, expected);
    }
    #[test]
    fn transitional_custom_view_reference_round_trip() {
        let xml = format!(
            "<x:chartsheet xmlns:x=\"{SML}\" xmlns:r=\"{REL}\"><x:sheetViews><x:sheetView workbookViewId=\"0\"/></x:sheetViews><x:customSheetViews><x:customSheetView guid=\"{{00112233-4455-6677-8899-AABBCCDDEEFF}}\" scale=\"10\" state=\"veryHidden\" zoomToFit=\"1\"/><x:customSheetView guid=\"{{10213243-5465-7687-98A9-BACBDCEDFE0F}}\"/></x:customSheetViews><x:drawing r:id=\"rId1\"/></x:chartsheet>"
        );
        let (_, parsed) = parse_chartsheet(xml.as_bytes()).unwrap();
        let views = parsed.custom_views.as_ref().unwrap();
        assert_eq!(views.len(), 2);
        assert_eq!(views[0].state, Some(ChartSheetState::VeryHidden));
        assert_eq!(views[0].scale, Some(10));
        let written = write_chartsheet(&parsed, ChartSheetConformance::Transitional).unwrap();
        assert_eq!(parse_chartsheet(&written).unwrap().1, parsed);
    }
    #[test]
    fn mce_fallback_selects_chartsheet_views() {
        let xml = format!(
            "<x:chartsheet xmlns:x=\"{SML}\" xmlns:r=\"{REL}\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:u=\"urn:unsupported\"><mc:AlternateContent><mc:Choice Requires=\"u\"><u:views/></mc:Choice><mc:Fallback><x:sheetViews><x:sheetView workbookViewId=\"0\"/></x:sheetViews></mc:Fallback></mc:AlternateContent><x:drawing r:id=\"rId1\"/></x:chartsheet>"
        );
        assert_eq!(parse_chartsheet(xml.as_bytes()).unwrap().1.views.len(), 1);
    }
    #[test]
    fn mce_fallback_selects_custom_chartsheet_views() {
        let xml = format!(
            "<x:chartsheet xmlns:x=\"{SML}\" xmlns:r=\"{REL}\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:u=\"urn:unsupported\"><x:sheetViews><x:sheetView workbookViewId=\"0\"/></x:sheetViews><mc:AlternateContent><mc:Choice Requires=\"u\"><u:customViews/></mc:Choice><mc:Fallback><x:customSheetViews><x:customSheetView guid=\"{{00112233-4455-6677-8899-AABBCCDDEEFF}}\" scale=\"200\"/></x:customSheetViews></mc:Fallback></mc:AlternateContent><x:drawing r:id=\"rId1\"/></x:chartsheet>"
        );
        let parsed = parse_chartsheet(xml.as_bytes()).unwrap().1;
        assert_eq!(parsed.custom_views.unwrap()[0].scale, Some(200));
    }
    #[test]
    fn loads_both_poi_chartsheet_graphs() {
        for (bytes, name, zoom) in [(POI_ONE, "Chart2", 131), (POI_TWO, "Chart1", 84)] {
            let package = OpcPackage::from_bytes(bytes).unwrap();
            let workbook = PackURI::new("/xl/workbook.xml").unwrap();
            let workbook_part = package.get_part(&workbook).unwrap();
            let id = workbook_part
                .rels()
                .iter()
                .find(|rel| rel.reltype() == CHARTSHEET_REL)
                .unwrap()
                .r_id()
                .to_owned();
            let loaded = load_chartsheet(&package, &workbook, &id).unwrap();
            assert_eq!(loaded.entry.name, name);
            assert_eq!(loaded.chartsheet.views[0].zoom_scale, Some(zoom));
            assert_eq!(loaded.drawing.charts.len(), 1);
            assert!(loaded.drawing.charts[0].data.starts_with(b"<?xml"));
        }
    }
    #[test]
    fn strict_package_writer_round_trips_complete_leaf_graph() {
        let conformance = ChartSheetConformance::Strict;
        let (mut package, workbook) = base_package(conformance);
        let expected = value(conformance);
        store_chartsheet(&mut package, &workbook, &expected, conformance).unwrap();
        assert_eq!(
            load_chartsheet(&package, &workbook, "rIdChartSheet").unwrap(),
            expected
        );
    }
    #[test]
    fn printer_settings_page_setup_strict_mce_and_schema_order() {
        let xml = format!(
            "<x:chartsheet xmlns:x=\"{STRICT_SML}\" xmlns:r=\"{STRICT_REL}\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:u=\"urn:unsupported\"><x:sheetViews><x:sheetView workbookViewId=\"0\"/></x:sheetViews><mc:AlternateContent><mc:Choice Requires=\"u\"><u:pageSetup/></mc:Choice><mc:Fallback><x:pageSetup orientation=\"landscape\" r:id=\"rIdPrinter\"/></mc:Fallback></mc:AlternateContent><x:drawing r:id=\"rIdDrawing\"/></x:chartsheet>"
        );
        let (kind, parsed) = parse_chartsheet(xml.as_bytes()).unwrap();
        assert_eq!(kind, ChartSheetConformance::Strict);
        assert_eq!(
            parsed
                .page_setup
                .as_ref()
                .unwrap()
                .printer_settings_relationship_id
                .as_deref(),
            Some("rIdPrinter")
        );
        let written = write_chartsheet(&parsed, kind).unwrap();
        let text = std::str::from_utf8(&written).unwrap();
        assert!(text.find("pageSetup").unwrap() < text.find("drawing").unwrap());
        assert!(text.contains("r:id=\"rIdPrinter\""));
        assert_eq!(parse_chartsheet(&written).unwrap().1, parsed);
        for body in [
            format!("<pageSetup xmlns:r=\"{REL}\" r:id=\"rIdP\"/><pageSetup/>"),
            format!("<drawing xmlns:r=\"{REL}\" r:id=\"rIdD\"/><pageSetup r:id=\"rIdP\"/>"),
            format!(
                "<pageSetup xmlns:r=\"{STRICT_REL}\" xmlns:t=\"{REL}\" t:id=\"rIdP\"/><drawing xmlns:r=\"{STRICT_REL}\" r:id=\"rIdD\"/>"
            ),
        ] {
            let xml = format!(
                "<chartsheet xmlns=\"{STRICT_SML}\"><sheetViews><sheetView workbookViewId=\"0\"/></sheetViews>{body}</chartsheet>"
            );
            assert!(parse_chartsheet(xml.as_bytes()).is_err(), "accepted {body}");
        }
    }
    #[test]
    fn printer_settings_package_round_trip_preserves_opaque_bytes() {
        let conformance = ChartSheetConformance::Transitional;
        let (mut package, workbook) = base_package(conformance);
        let expected = value(conformance);
        store_chartsheet(&mut package, &workbook, &expected, conformance).unwrap();
        let loaded = load_chartsheet(&package, &workbook, "rIdChartSheet").unwrap();
        assert_eq!(loaded.printer_settings, expected.printer_settings);
        let part = package
            .get_part(&PackURI::new("/xl/printerSettings/printerSettings1.bin").unwrap())
            .unwrap();
        assert_eq!(part.content_type(), PRINTER_CT);
        assert_eq!(
            part.blob(),
            [0x44, 0x45, 0x56, 0x4d, 0x4f, 0x44, 0x45, 0, 255]
        );
    }
    #[test]
    fn rejects_printer_settings_pairing_paths_collisions_and_caps_before_mutation() {
        let conformance = ChartSheetConformance::Transitional;
        let (mut package, workbook) = base_package(conformance);
        let mut bad = value(conformance);
        bad.printer_settings.as_mut().unwrap().relationship_id = "rIdOther".into();
        assert!(store_chartsheet(&mut package, &workbook, &bad, conformance).is_err());
        for path in [
            "/xl/printerSettings/sub/settings.bin",
            "/xl/printerSettings/settings.dat",
            "/xl/media/settings.bin",
        ] {
            let (mut package, workbook) = base_package(conformance);
            let mut bad = value(conformance);
            bad.printer_settings.as_mut().unwrap().resource.part_name = path.into();
            assert!(
                store_chartsheet(&mut package, &workbook, &bad, conformance).is_err(),
                "accepted {path}"
            );
        }
        let (mut package, workbook) = base_package(conformance);
        let mut bad = value(conformance);
        bad.printer_settings.as_mut().unwrap().resource.data = vec![0; MAX_SETTINGS_BYTES + 1];
        assert!(store_chartsheet(&mut package, &workbook, &bad, conformance).is_err());
        assert!(
            package
                .get_part(&PackURI::new("/xl/chartsheets/sheet1.xml").unwrap())
                .is_err()
        );
        let (mut package, workbook) = base_package(conformance);
        package.add_part(Box::new(BlobPart::new(
            PackURI::new("/xl/printerSettings/printerSettings1.bin").unwrap(),
            PRINTER_CT.into(),
            vec![1],
        )));
        assert!(
            store_chartsheet(&mut package, &workbook, &value(conformance), conformance).is_err()
        );
        assert!(
            package
                .get_part(&PackURI::new("/xl/chartsheets/sheet1.xml").unwrap())
                .is_err()
        );
    }
    #[test]
    fn rejects_printer_settings_external_wrong_type_escape_orphan_content_type_and_outbound_graphs()
    {
        for (kind, target, external) in [
            (PRINTER_REL, "https://example.invalid/settings.bin", true),
            (IMAGE_REL, "../printerSettings/printerSettings1.bin", false),
            (PRINTER_REL, "../../../evil.bin", false),
        ] {
            let conformance = ChartSheetConformance::Transitional;
            let (mut package, workbook) = base_package(conformance);
            store_chartsheet(&mut package, &workbook, &value(conformance), conformance).unwrap();
            let chartsheet = package
                .get_part_mut(&PackURI::new("/xl/chartsheets/sheet1.xml").unwrap())
                .unwrap();
            chartsheet.rels_mut().remove("rIdPrinter");
            chartsheet.rels_mut().add_relationship(
                kind.into(),
                target.into(),
                "rIdPrinter".into(),
                external,
            );
            assert!(
                load_chartsheet(&package, &workbook, "rIdChartSheet").is_err(),
                "accepted {kind} {target}"
            );
        }
        let conformance = ChartSheetConformance::Transitional;
        let (mut package, workbook) = base_package(conformance);
        store_chartsheet(&mut package, &workbook, &value(conformance), conformance).unwrap();
        package
            .get_part_mut(&PackURI::new("/xl/chartsheets/sheet1.xml").unwrap())
            .unwrap()
            .rels_mut()
            .add_relationship(
                PRINTER_REL.into(),
                "../printerSettings/printerSettings1.bin".into(),
                "rIdOrphan".into(),
                false,
            );
        assert!(load_chartsheet(&package, &workbook, "rIdChartSheet").is_err());
        let (mut package, workbook) = base_package(conformance);
        store_chartsheet(&mut package, &workbook, &value(conformance), conformance).unwrap();
        package
            .get_part_mut(&PackURI::new("/xl/printerSettings/printerSettings1.bin").unwrap())
            .unwrap()
            .rels_mut()
            .add_relationship(
                IMAGE_REL.into(),
                "../media/evil.png".into(),
                "rIdOutbound".into(),
                false,
            );
        assert!(load_chartsheet(&package, &workbook, "rIdChartSheet").is_err());
        let (mut package, workbook) = base_package(conformance);
        store_chartsheet(&mut package, &workbook, &value(conformance), conformance).unwrap();
        package.add_part(Box::new(BlobPart::new(
            PackURI::new("/xl/printerSettings/printerSettings1.bin").unwrap(),
            "application/octet-stream".into(),
            vec![1],
        )));
        assert!(load_chartsheet(&package, &workbook, "rIdChartSheet").is_err());
    }
    #[test]
    fn web_publish_schema_enum_and_deterministic_round_trip() {
        let mut body = String::from("<webPublishItems count=\"8\">");
        for (index, kind) in [
            "sheet",
            "printArea",
            "autoFilter",
            "range",
            "chart",
            "pivotTable",
            "query",
            "label",
        ]
        .iter()
        .enumerate()
        {
            let source_ref = if *kind == "range" {
                " sourceRef=\"A1:B2\""
            } else {
                ""
            };
            let source_object = if matches!(*kind, "pivotTable" | "query" | "label") {
                " sourceObject=\"OpaqueName\""
            } else {
                ""
            };
            body.push_str(&format!("<webPublishItem id=\"{index}\" divId=\"Div{index}\" sourceType=\"{kind}\"{source_ref}{source_object} destinationFile=\"opaque:{index}\" autoRepublish=\"{}\"/>",if index%2==0{"true"}else{"0"}));
        }
        body.push_str("</webPublishItems>");
        let xml = format!(
            "<chartsheet xmlns=\"{STRICT_SML}\" xmlns:r=\"{STRICT_REL}\"><sheetViews><sheetView workbookViewId=\"0\"/></sheetViews><drawing r:id=\"rIdD\"/>{body}</chartsheet>"
        );
        let (kind, parsed) = parse_chartsheet(xml.as_bytes()).unwrap();
        assert_eq!(kind, ChartSheetConformance::Strict);
        let items = &parsed.web_publish_items.as_ref().unwrap().items;
        assert_eq!(
            items
                .iter()
                .map(|item| item.source_type)
                .collect::<Vec<_>>(),
            vec![
                ChartSheetWebSourceType::Sheet,
                ChartSheetWebSourceType::PrintArea,
                ChartSheetWebSourceType::AutoFilter,
                ChartSheetWebSourceType::Range,
                ChartSheetWebSourceType::Chart,
                ChartSheetWebSourceType::PivotTable,
                ChartSheetWebSourceType::Query,
                ChartSheetWebSourceType::Label
            ]
        );
        let first = write_chartsheet(&parsed, kind).unwrap();
        let reparsed = parse_chartsheet(&first).unwrap().1;
        let second = write_chartsheet(&reparsed, kind).unwrap();
        assert_eq!(first, second);
        assert_eq!(parsed, reparsed);
    }
    #[test]
    fn web_publish_mce_fallback_and_inert_references() {
        let xml = format!(
            "<x:chartsheet xmlns:x=\"{SML}\" xmlns:r=\"{REL}\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:u=\"urn:unsupported\"><x:sheetViews><x:sheetView workbookViewId=\"0\"/></x:sheetViews><x:drawing r:id=\"rIdD\"/><mc:AlternateContent><mc:Choice Requires=\"u\"><u:run href=\"https://example.invalid/execute\"/></mc:Choice><mc:Fallback><x:webPublishItems><x:webPublishItem id=\"0\" divId=\"D\" sourceType=\"chart\" sourceObject=\"file:///not/read\" destinationFile=\"/tmp/not-written\" title=\"$(never-execute)\"/></x:webPublishItems></mc:Fallback></mc:AlternateContent></x:chartsheet>"
        );
        let (_, parsed) = parse_chartsheet(xml.as_bytes()).unwrap();
        let item = &parsed.web_publish_items.as_ref().unwrap().items[0];
        assert_eq!(item.destination_file, "/tmp/not-written");
        assert_eq!(item.auto_republish, None);
        assert_eq!(
            parse_chartsheet(
                &write_chartsheet(&parsed, ChartSheetConformance::Transitional).unwrap()
            )
            .unwrap()
            .1,
            parsed
        );
    }
    #[test]
    fn web_publish_package_load_store_preserves_metadata_only() {
        let conformance = ChartSheetConformance::Transitional;
        let (mut package, workbook) = base_package(conformance);
        let expected = value(conformance);
        store_chartsheet(&mut package, &workbook, &expected, conformance).unwrap();
        let loaded = load_chartsheet(&package, &workbook, "rIdChartSheet").unwrap();
        assert_eq!(
            loaded.chartsheet.web_publish_items,
            expected.chartsheet.web_publish_items
        );
    }
    #[test]
    fn rejects_web_publish_malformed_duplicates_cardinality_and_order() {
        let wrap = |body: &str| {
            format!(
                "<chartsheet xmlns=\"{SML}\" xmlns:r=\"{REL}\"><sheetViews><sheetView workbookViewId=\"0\"/></sheetViews><drawing r:id=\"rIdD\"/>{body}</chartsheet>"
            )
        };
        for body in [
            "<webPublishItems/>",
            "<webPublishItems count=\"2\"><webPublishItem id=\"1\" divId=\"A\" sourceType=\"sheet\" destinationFile=\"x\"/></webPublishItems>",
            "<webPublishItems><webPublishItem id=\"1\" divId=\"A\" sourceType=\"sheet\" destinationFile=\"x\"/><webPublishItem id=\"1\" divId=\"B\" sourceType=\"sheet\" destinationFile=\"y\"/></webPublishItems>",
            "<webPublishItems><webPublishItem id=\"1\" divId=\"A\" sourceType=\"sheet\" destinationFile=\"x\"/><webPublishItem id=\"2\" divId=\"A\" sourceType=\"sheet\" destinationFile=\"y\"/></webPublishItems>",
            "<webPublishItems><webPublishItem id=\"4294967296\" divId=\"A\" sourceType=\"sheet\" destinationFile=\"x\"/></webPublishItems>",
            "<webPublishItems><webPublishItem id=\"1\" divId=\"A\" sourceType=\"bad\" destinationFile=\"x\"/></webPublishItems>",
            "<webPublishItems><webPublishItem id=\"1\" divId=\"A\" sourceType=\"range\" destinationFile=\"x\"/></webPublishItems>",
            "<webPublishItems><webPublishItem id=\"1\" divId=\"A\" sourceType=\"query\" destinationFile=\"x\"/></webPublishItems>",
            "<webPublishItems><webPublishItem id=\"1\" divId=\"A\" sourceType=\"sheet\" destinationFile=\"x\" autoRepublish=\"on\"/></webPublishItems>",
            "<webPublishItems><webPublishItem id=\"1\" divId=\"A\" sourceType=\"sheet\" destinationFile=\"x\" extra=\"1\"/></webPublishItems>",
        ] {
            let xml = wrap(body);
            assert!(parse_chartsheet(xml.as_bytes()).is_err(), "accepted {body}");
        }
        let out_of_order = format!(
            "<chartsheet xmlns=\"{SML}\" xmlns:r=\"{REL}\"><sheetViews><sheetView workbookViewId=\"0\"/></sheetViews><webPublishItems><webPublishItem id=\"1\" divId=\"A\" sourceType=\"sheet\" destinationFile=\"x\"/></webPublishItems><drawing r:id=\"rIdD\"/></chartsheet>"
        );
        assert!(parse_chartsheet(out_of_order.as_bytes()).is_err());
    }
    #[test]
    fn rejects_web_publish_count_and_string_caps() {
        let mut body = String::from("<webPublishItems>");
        for index in 0..=MAX_WEB_PUBLISH_ITEMS {
            body.push_str(&format!("<webPublishItem id=\"{index}\" divId=\"D{index}\" sourceType=\"sheet\" destinationFile=\"x\"/>"));
        }
        body.push_str("</webPublishItems>");
        let xml = format!(
            "<chartsheet xmlns=\"{SML}\" xmlns:r=\"{REL}\"><sheetViews><sheetView workbookViewId=\"0\"/></sheetViews><drawing r:id=\"rIdD\"/>{body}</chartsheet>"
        );
        assert!(parse_chartsheet(xml.as_bytes()).is_err());
        let mut value = sheet();
        value.web_publish_items.as_mut().unwrap().items[0].title =
            Some("x".repeat(MAX_WEB_PUBLISH_STRING_BYTES + 1));
        assert!(write_chartsheet(&value, ChartSheetConformance::Transitional).is_err());
    }
    #[test]
    fn picture_mce_schema_and_inert_round_trip() {
        let xml = format!(
            "<x:chartsheet xmlns:x=\"{SML}\" xmlns:r=\"{REL}\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:u=\"urn:unsupported\"><x:sheetViews><x:sheetView workbookViewId=\"0\"/></x:sheetViews><x:drawing r:id=\"rIdDrawing\"/><mc:AlternateContent><mc:Choice Requires=\"u\"><u:picture/></mc:Choice><mc:Fallback><x:picture r:id=\"rIdBackground\"/></mc:Fallback></mc:AlternateContent></x:chartsheet>"
        );
        let (_, parsed) = parse_chartsheet(xml.as_bytes()).unwrap();
        assert_eq!(
            parsed.background_picture_relationship_id.as_deref(),
            Some("rIdBackground")
        );
        let written = write_chartsheet(&parsed, ChartSheetConformance::Transitional).unwrap();
        assert!(
            String::from_utf8(written.clone())
                .unwrap()
                .contains("<x:drawing r:id=\"rIdDrawing\"/><x:picture r:id=\"rIdBackground\"/>")
        );
        assert_eq!(parse_chartsheet(&written).unwrap().1, parsed);
    }
    #[test]
    fn transitional_picture_package_round_trip_preserves_opaque_bytes() {
        let conformance = ChartSheetConformance::Transitional;
        let (mut package, workbook) = base_package(conformance);
        let expected = value(conformance);
        store_chartsheet(&mut package, &workbook, &expected, conformance).unwrap();
        let loaded = load_chartsheet(&package, &workbook, "rIdChartSheet").unwrap();
        assert_eq!(
            loaded.background_picture.as_ref().unwrap().data,
            vec![0, 255, 1, 254]
        );
        assert_eq!(loaded, expected);
    }
    #[test]
    fn vml_mce_schema_order_and_inert_round_trip() {
        let xml = format!(
            "<x:chartsheet xmlns:x=\"{SML}\" xmlns:r=\"{REL}\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:u=\"urn:unsupported\"><x:sheetViews><x:sheetView workbookViewId=\"0\"/></x:sheetViews><x:drawing r:id=\"rIdDrawing\"/><mc:AlternateContent><mc:Choice Requires=\"u\"><u:vml/></mc:Choice><mc:Fallback><x:legacyDrawing r:id=\"rIdLegacy\"/><x:legacyDrawingHF r:id=\"rIdLegacyHF\"/></mc:Fallback></mc:AlternateContent><x:picture r:id=\"rIdBackground\"/></x:chartsheet>"
        );
        let (_, parsed) = parse_chartsheet(xml.as_bytes()).unwrap();
        assert_eq!(
            parsed.legacy_drawing_relationship_id.as_deref(),
            Some("rIdLegacy")
        );
        assert_eq!(
            parsed
                .legacy_header_footer_drawing_relationship_id
                .as_deref(),
            Some("rIdLegacyHF")
        );
        let written = write_chartsheet(&parsed, ChartSheetConformance::Transitional).unwrap();
        let text = String::from_utf8(written.clone()).unwrap();
        assert!(text.contains("<x:drawing r:id=\"rIdDrawing\"/><x:legacyDrawing r:id=\"rIdLegacy\"/><x:legacyDrawingHF r:id=\"rIdLegacyHF\"/><x:picture"));
        assert_eq!(parse_chartsheet(&written).unwrap().1, parsed);
    }
    #[test]
    fn rejects_vml_schema_duplicates_order_and_missing_ids() {
        for body in [
            "<legacyDrawing xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:id=\"rIdL\"/><drawing xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:id=\"rIdD\"/>",
            "<drawing xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:id=\"rIdD\"/><legacyDrawing r:id=\"rIdL\"/><legacyDrawing r:id=\"rIdL2\"/>",
            "<drawing xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:id=\"rIdD\"/><legacyDrawingHF/>",
            "<drawing xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:id=\"rIdD\"/><legacyDrawingHF r:id=\"rIdHF\"/><legacyDrawing r:id=\"rIdL\"/>",
        ] {
            let xml = format!(
                "<chartsheet xmlns=\"{SML}\"><sheetViews><sheetView workbookViewId=\"0\"/></sheetViews>{body}</chartsheet>"
            );
            assert!(parse_chartsheet(xml.as_bytes()).is_err(), "accepted {body}");
        }
    }
    #[test]
    fn rejects_vml_pairing_content_type_collision_and_caps() {
        let conformance = ChartSheetConformance::Transitional;
        let (mut package, workbook) = base_package(conformance);
        let mut bad = value(conformance);
        bad.legacy_drawing.as_mut().unwrap().relationship_id = "rIdOther".into();
        assert!(store_chartsheet(&mut package, &workbook, &bad, conformance).is_err());
        let (mut package, workbook) = base_package(conformance);
        let mut bad = value(conformance);
        bad.legacy_drawing.as_mut().unwrap().content_type = "application/xml".into();
        assert!(store_chartsheet(&mut package, &workbook, &bad, conformance).is_err());
        let (mut package, workbook) = base_package(conformance);
        let mut bad = value(conformance);
        bad.legacy_header_footer_drawing.as_mut().unwrap().part_name =
            bad.legacy_drawing.as_ref().unwrap().part_name.clone();
        assert!(store_chartsheet(&mut package, &workbook, &bad, conformance).is_err());
        let (mut package, workbook) = base_package(conformance);
        let mut bad = value(conformance);
        bad.legacy_drawing.as_mut().unwrap().data = vec![0; MAX_VML_DRAWING_BYTES + 1];
        assert!(store_chartsheet(&mut package, &workbook, &bad, conformance).is_err());
        let (mut package, workbook) = base_package(conformance);
        store_chartsheet(&mut package, &workbook, &value(conformance), conformance).unwrap();
        package
            .get_part_mut(&PackURI::new("/xl/drawings/vmlDrawing1.vml").unwrap())
            .unwrap()
            .set_blob(vec![0; MAX_VML_DRAWING_BYTES + 1]);
        assert!(load_chartsheet(&package, &workbook, "rIdChartSheet").is_err());
    }
    #[test]
    fn rejects_external_wrong_type_escaped_orphan_and_outbound_vml_graphs() {
        for (kind, target, external) in [
            (
                VML_DRAWING_REL,
                "https://example.invalid/vmlDrawing1.vml",
                true,
            ),
            (IMAGE_REL, "../drawings/vmlDrawing1.vml", false),
            (VML_DRAWING_REL, "../../../evil.vml", false),
        ] {
            let conformance = ChartSheetConformance::Transitional;
            let (mut package, workbook) = base_package(conformance);
            store_chartsheet(&mut package, &workbook, &value(conformance), conformance).unwrap();
            let chartsheet = package
                .get_part_mut(&PackURI::new("/xl/chartsheets/sheet1.xml").unwrap())
                .unwrap();
            chartsheet.rels_mut().remove("rIdLegacy");
            chartsheet.rels_mut().add_relationship(
                kind.into(),
                target.into(),
                "rIdLegacy".into(),
                external,
            );
            assert!(
                load_chartsheet(&package, &workbook, "rIdChartSheet").is_err(),
                "accepted {kind} {target}"
            );
        }
        let conformance = ChartSheetConformance::Transitional;
        let (mut package, workbook) = base_package(conformance);
        store_chartsheet(&mut package, &workbook, &value(conformance), conformance).unwrap();
        package
            .get_part_mut(&PackURI::new("/xl/chartsheets/sheet1.xml").unwrap())
            .unwrap()
            .rels_mut()
            .add_relationship(
                VML_DRAWING_REL.into(),
                "../drawings/vmlDrawing1.vml".into(),
                "rIdOrphan".into(),
                false,
            );
        assert!(load_chartsheet(&package, &workbook, "rIdChartSheet").is_err());
        let (mut package, workbook) = base_package(conformance);
        store_chartsheet(&mut package, &workbook, &value(conformance), conformance).unwrap();
        package
            .get_part_mut(&PackURI::new("/xl/drawings/vmlDrawing1.vml").unwrap())
            .unwrap()
            .rels_mut()
            .add_relationship(
                IMAGE_REL.into(),
                "../media/evil.png".into(),
                "rIdOutbound".into(),
                false,
            );
        assert!(load_chartsheet(&package, &workbook, "rIdChartSheet").is_err());
    }
    #[test]
    fn rejects_picture_cardinality_order_metadata_and_caps() {
        for xml in [
            format!(
                "<chartsheet xmlns=\"{SML}\" xmlns:r=\"{REL}\"><sheetViews><sheetView workbookViewId=\"0\"/></sheetViews><picture r:id=\"rIdP\"/><drawing r:id=\"rIdD\"/></chartsheet>"
            ),
            format!(
                "<chartsheet xmlns=\"{SML}\" xmlns:r=\"{REL}\"><sheetViews><sheetView workbookViewId=\"0\"/></sheetViews><drawing r:id=\"rIdD\"/><picture r:id=\"rIdP\"/><picture r:id=\"rIdQ\"/></chartsheet>"
            ),
            format!(
                "<chartsheet xmlns=\"{SML}\" xmlns:r=\"{REL}\"><sheetViews><sheetView workbookViewId=\"0\"/></sheetViews><drawing r:id=\"rIdD\"/><picture/></chartsheet>"
            ),
        ] {
            assert!(parse_chartsheet(xml.as_bytes()).is_err(), "{xml}");
        }
        let conformance = ChartSheetConformance::Transitional;
        let (mut package, workbook) = base_package(conformance);
        let mut bad = value(conformance);
        bad.background_picture.as_mut().unwrap().relationship_id = "different".into();
        assert!(store_chartsheet(&mut package, &workbook, &bad, conformance).is_err());
        let (mut package, workbook) = base_package(conformance);
        let mut bad = value(conformance);
        bad.background_picture.as_mut().unwrap().part_name = "/xl/charts/chart1.xml".into();
        assert!(store_chartsheet(&mut package, &workbook, &bad, conformance).is_err());
        let (mut package, workbook) = base_package(conformance);
        let mut bad = value(conformance);
        bad.background_picture.as_mut().unwrap().data = vec![0; MAX_BACKGROUND_IMAGE_BYTES + 1];
        assert!(store_chartsheet(&mut package, &workbook, &bad, conformance).is_err());
    }
    #[test]
    fn rejects_external_wrong_type_escaped_and_unreferenced_picture_relationships() {
        for (kind, target, external) in [
            (IMAGE_REL, "https://example.invalid/background.png", true),
            (rt::CHART, "../media/background1.png", false),
            (IMAGE_REL, "../../../evil.png", false),
        ] {
            let conformance = ChartSheetConformance::Transitional;
            let (mut package, workbook) = base_package(conformance);
            store_chartsheet(&mut package, &workbook, &value(conformance), conformance).unwrap();
            let chartsheet = package
                .get_part_mut(&PackURI::new("/xl/chartsheets/sheet1.xml").unwrap())
                .unwrap();
            chartsheet.rels_mut().remove("rIdBackground");
            chartsheet.rels_mut().add_relationship(
                kind.into(),
                target.into(),
                "rIdBackground".into(),
                external,
            );
            assert!(
                load_chartsheet(&package, &workbook, "rIdChartSheet").is_err(),
                "accepted {kind} {target}"
            );
        }
        let conformance = ChartSheetConformance::Transitional;
        let (mut package, workbook) = base_package(conformance);
        store_chartsheet(&mut package, &workbook, &value(conformance), conformance).unwrap();
        package
            .get_part_mut(&PackURI::new("/xl/chartsheets/sheet1.xml").unwrap())
            .unwrap()
            .rels_mut()
            .add_relationship(
                IMAGE_REL.into(),
                "../media/background1.png".into(),
                "rIdExtra".into(),
                false,
            );
        assert!(load_chartsheet(&package, &workbook, "rIdChartSheet").is_err());
    }
    #[test]
    fn rejects_existing_background_part_collision_before_mutation() {
        let conformance = ChartSheetConformance::Transitional;
        let (mut package, workbook) = base_package(conformance);
        package.add_part(Box::new(BlobPart::new(
            PackURI::new("/xl/media/background1.png").unwrap(),
            "image/png".into(),
            vec![9],
        )));
        assert!(
            store_chartsheet(&mut package, &workbook, &value(conformance), conformance).is_err()
        );
        assert!(
            package
                .get_part(&PackURI::new("/xl/chartsheets/sheet1.xml").unwrap())
                .is_err()
        );
    }
    #[test]
    fn rejects_malformed_caps_and_graphs() {
        assert!(parse_chartsheet(b"<!DOCTYPE x><chartsheet/>").is_err());
        assert!(parse_chartsheet(format!("<chartsheet xmlns=\"{SML}\"><sheetViews><sheetView workbookViewId=\"0\" zoomScale=\"401\"/></sheetViews><drawing xmlns:r=\"{REL}\" r:id=\"rId1\"/></chartsheet>").as_bytes()).is_err());
        for custom in [
            "<customSheetViews/>",
            "<customSheetViews><customSheetView guid=\"bad\"/></customSheetViews>",
            "<customSheetViews><customSheetView guid=\"{00112233-4455-6677-8899-AABBCCDDEEFF}\" scale=\"401\"/></customSheetViews>",
            "<customSheetViews><customSheetView guid=\"{00112233-4455-6677-8899-AABBCCDDEEFF}\"/><customSheetView guid=\"{00112233-4455-6677-8899-aabbccddeeff}\"/></customSheetViews>",
        ] {
            let xml = format!(
                "<chartsheet xmlns=\"{SML}\" xmlns:r=\"{REL}\"><sheetViews><sheetView workbookViewId=\"0\"/></sheetViews>{custom}<drawing r:id=\"rId1\"/></chartsheet>"
            );
            assert!(parse_chartsheet(xml.as_bytes()).is_err(), "{custom}");
        }
        assert!(parse_chartsheet(&vec![b' '; MAX_XML_BYTES + 1]).is_err());
        let (mut package, workbook) = base_package(ChartSheetConformance::Transitional);
        let expected = value(ChartSheetConformance::Transitional);
        store_chartsheet(
            &mut package,
            &workbook,
            &expected,
            ChartSheetConformance::Transitional,
        )
        .unwrap();
        package
            .get_part_mut(&PackURI::new("/xl/drawings/drawing1.xml").unwrap())
            .unwrap()
            .rels_mut()
            .add_relationship(
                rt::IMAGE.into(),
                "../media/x.png".into(),
                "rIdBad".into(),
                false,
            );
        assert!(load_chartsheet(&package, &workbook, "rIdChartSheet").is_err());
    }
}

#[cfg(test)]
mod chart_outbound_tests {
    use super::*;

    fn outbound_value(conformance: ChartSheetConformance) -> ChartSheetPackage {
        let source = OpcPackage::from_bytes(tests::LO_USER_SHAPES_IMAGES).unwrap();
        let blob = |path: &str| {
            source
                .get_part(&PackURI::new(path).unwrap())
                .unwrap()
                .blob()
                .to_vec()
        };
        let mut value = tests::chart_ex_user_shapes(conformance);
        let drawing_main = if conformance == ChartSheetConformance::Strict {
            STRICT_DRAWING_MAIN
        } else {
            DRAWING_MAIN
        };
        let chart = &mut value.drawing.charts[0];
        chart.data = format!(
            "<cx:chartSpace xmlns:cx=\"{CHART_EX}\" xmlns:a=\"{drawing_main}\" xmlns:r=\"{}\"><cx:chart><cx:plotArea><cx:plotAreaRegion><cx:series><cx:spPr><a:blipFill><a:blip r:embed=\"rIdDirectImage\"/></a:blipFill></cx:spPr></cx:series></cx:plotAreaRegion></cx:plotArea></cx:chart><cx:externalData r:id=\"rIdPackage\" cx:autoUpdate=\"0\"/></cx:chartSpace>",
            conformance.rel()
        )
        .into_bytes();
        let theme_data = format!(
            "<a:themeOverride xmlns:a=\"{drawing_main}\" xmlns:r=\"{}\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:u=\"urn:unsupported\"><mc:AlternateContent><mc:Choice Requires=\"u\"><u:active/></mc:Choice><mc:Fallback><a:fmtScheme name=\"Inert\"><a:fillStyleLst><a:blipFill><a:blip r:embed=\"rIdThemeImage\"/></a:blipFill></a:fillStyleLst><a:lnStyleLst/><a:effectStyleLst/><a:bgFillStyleLst/></a:fmtScheme></mc:Fallback></mc:AlternateContent></a:themeOverride>",
            conformance.rel()
        )
        .into_bytes();
        let ChartSheetChartResourceKind::Extended {
            outbound_resources, ..
        } = &mut chart.kind
        else {
            unreachable!()
        };
        *outbound_resources = vec![
            ChartSheetChartOutboundResource::Image(ChartSheetChartImageResource {
                relationship_id: "rIdDirectImage".into(),
                part_name: "/xl/media/chartDirect1.png".into(),
                content_type: ChartSheetChartImageContentType::Png,
                data: blob("/xl/media/image1.png"),
            }),
            ChartSheetChartOutboundResource::EmbeddedPackage(
                ChartSheetChartEmbeddedPackageResource {
                    relationship_id: "rIdPackage".into(),
                    part_name: "/xl/embeddings/Microsoft_Excel_Worksheet1.xlsx".into(),
                    content_type: ChartSheetChartEmbeddedPackageContentType::Xlsx,
                    data: tests::LO_USER_SHAPES_IMAGES.to_vec(),
                },
            ),
            ChartSheetChartOutboundResource::ThemeOverride(ChartSheetChartThemeOverrideResource {
                relationship_id: "rIdTheme".into(),
                part_name: "/xl/theme/themeOverride1.xml".into(),
                content_type: THEME_OVERRIDE_CT.into(),
                data: theme_data,
                images: vec![ChartSheetChartImageResource {
                    relationship_id: "rIdThemeImage".into(),
                    part_name: "/xl/media/themeImage1.svg".into(),
                    content_type: ChartSheetChartImageContentType::Svg,
                    data: blob("/xl/media/image2.svg"),
                }],
            }),
        ];
        value
    }

    #[test]
    fn chart_ex_complete_outbound_family_round_trips_strict_and_transitional() {
        for conformance in [
            ChartSheetConformance::Transitional,
            ChartSheetConformance::Strict,
        ] {
            let expected = outbound_value(conformance);
            let (mut package, workbook) = tests::base_package(conformance);
            store_chartsheet(&mut package, &workbook, &expected, conformance).unwrap();
            let loaded = load_chartsheet(&package, &workbook, "rIdChartSheet").unwrap();
            assert_eq!(loaded, expected);
            let chart_uri = PackURI::new("/xl/charts/chartEx1.xml").unwrap();
            let chart = package.get_part(&chart_uri).unwrap();
            assert_eq!(
                chart.rels().get("rIdDirectImage").unwrap().reltype(),
                conformance.image_rel()
            );
            assert_eq!(
                chart.rels().get("rIdTheme").unwrap().reltype(),
                conformance.theme_override_rel()
            );
            assert_eq!(
                chart.rels().get("rIdPackage").unwrap().reltype(),
                conformance.package_rel()
            );
            let theme = package
                .get_part(&PackURI::new("/xl/theme/themeOverride1.xml").unwrap())
                .unwrap();
            assert_eq!(
                theme.rels().get("rIdThemeImage").unwrap().reltype(),
                conformance.image_rel()
            );
            let (mut second, second_workbook) = tests::base_package(conformance);
            store_chartsheet(&mut second, &second_workbook, &loaded, conformance).unwrap();
            assert_eq!(
                load_chartsheet(&second, &second_workbook, "rIdChartSheet").unwrap(),
                loaded
            );
        }
    }

    #[test]
    fn chart_ex_outbound_rejects_active_external_mismatch_collision_roots_and_caps() {
        let conformance = ChartSheetConformance::Transitional;
        let mut bad = outbound_value(conformance);
        bad.drawing.charts[0].data = String::from_utf8(bad.drawing.charts[0].data.clone())
            .unwrap()
            .replace("autoUpdate=\"0\"", "autoUpdate=\"true\"")
            .into_bytes();
        let (mut package, workbook) = tests::base_package(conformance);
        assert!(store_chartsheet(&mut package, &workbook, &bad, conformance).is_err());
        assert!(
            package
                .get_part(&PackURI::new("/xl/chartsheets/sheet1.xml").unwrap())
                .is_err()
        );

        let mut bad = outbound_value(conformance);
        bad.drawing.charts[0].data = String::from_utf8(bad.drawing.charts[0].data.clone())
            .unwrap()
            .replace(" r:embed=\"rIdDirectImage\"", "")
            .into_bytes();
        let (mut package, workbook) = tests::base_package(conformance);
        assert!(store_chartsheet(&mut package, &workbook, &bad, conformance).is_err());

        let mut bad = outbound_value(conformance);
        if let ChartSheetChartResourceKind::Extended {
            outbound_resources, ..
        } = &mut bad.drawing.charts[0].kind
        {
            let theme = outbound_resources
                .iter_mut()
                .find_map(|resource| match resource {
                    ChartSheetChartOutboundResource::ThemeOverride(theme) => Some(theme),
                    _ => None,
                })
                .unwrap();
            theme.data = String::from_utf8(theme.data.clone())
                .unwrap()
                .replace("themeOverride", "theme")
                .into_bytes();
        }
        let (mut package, workbook) = tests::base_package(conformance);
        assert!(store_chartsheet(&mut package, &workbook, &bad, conformance).is_err());

        let mut bad = outbound_value(conformance);
        if let ChartSheetChartResourceKind::Extended {
            user_shapes: Some(shapes),
            outbound_resources,
            ..
        } = &mut bad.drawing.charts[0].kind
        {
            outbound_resources[0] = match outbound_resources[0].clone() {
                ChartSheetChartOutboundResource::Image(mut image) => {
                    image.relationship_id = shapes.relationship_id.clone();
                    ChartSheetChartOutboundResource::Image(image)
                },
                _ => unreachable!(),
            };
        }
        let (mut package, workbook) = tests::base_package(conformance);
        assert!(store_chartsheet(&mut package, &workbook, &bad, conformance).is_err());

        let expected = outbound_value(conformance);
        let (mut package, workbook) = tests::base_package(conformance);
        store_chartsheet(&mut package, &workbook, &expected, conformance).unwrap();
        let chart = package
            .get_part_mut(&PackURI::new("/xl/charts/chartEx1.xml").unwrap())
            .unwrap();
        chart.rels_mut().remove("rIdPackage");
        chart.rels_mut().add_relationship(
            conformance.package_rel().into(),
            "https://example.invalid/active.xlsx".into(),
            "rIdPackage".into(),
            true,
        );
        assert!(load_chartsheet(&package, &workbook, "rIdChartSheet").is_err());

        let mut bad = outbound_value(conformance);
        if let ChartSheetChartResourceKind::Extended {
            outbound_resources, ..
        } = &mut bad.drawing.charts[0].kind
        {
            let embedded = outbound_resources
                .iter_mut()
                .find_map(|resource| match resource {
                    ChartSheetChartOutboundResource::EmbeddedPackage(embedded) => Some(embedded),
                    _ => None,
                })
                .unwrap();
            embedded.data = vec![0; MAX_CHART_EMBEDDED_PACKAGE_BYTES + 1];
        }
        let (mut package, workbook) = tests::base_package(conformance);
        assert!(store_chartsheet(&mut package, &workbook, &bad, conformance).is_err());
        assert!(
            package
                .get_part(&PackURI::new("/xl/chartsheets/sheet1.xml").unwrap())
                .is_err()
        );
    }
}
