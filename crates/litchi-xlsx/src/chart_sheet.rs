//! Typed SpreadsheetML chartsheet semantic model and bounded XML codec.
//!
//! The OOXML package graph, drawing/chart resources, and printer-settings
//! projection remain in the OOXML host. This module owns only the bounded
//! chartsheet part grammar described by [MS-XLSX] and [MS-OE376].

use crate::error::{Error, Result};
use litchi_ooxml_common::custom_xml::valid_guid;
use litchi_ooxml_common::{MceCapabilities, MceLimits, process_markup_compatibility};
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
const DRAWING_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing";
const STRICT_DRAWING_REL: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships/drawing";
const CHART_REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";
const STRICT_CHART_REL: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships/chart";
const CHARTSHEET_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet";
const STRICT_CHARTSHEET_REL: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/chartsheet";
const IMAGE_REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
const STRICT_IMAGE_REL: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships/image";
const CHART_USER_SHAPES_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartUserShapes";
const STRICT_CHART_USER_SHAPES_REL: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/chartUserShapes";
const THEME_OVERRIDE_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/themeOverride";
const STRICT_THEME_OVERRIDE_REL: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/themeOverride";
const PACKAGE_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package";
const STRICT_PACKAGE_REL: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships/package";
const VML_DRAWING_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing";
const STRICT_VML_DRAWING_REL: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/vmlDrawing";
const PRINTER_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings";
const STRICT_PRINTER_REL: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/printerSettings";

const MAX_XML_BYTES: usize = 32 * 1024 * 1024;
const MAX_NODES: usize = 500_000;
const MAX_DEPTH: usize = 256;
const MAX_NAMESPACE_BINDINGS: usize = 4096;
const MAX_STRING_BYTES: usize = 4 * 1024 * 1024;
const MAX_VIEWS: usize = 256;
const MAX_CUSTOM_VIEWS: usize = 1024;
const MAX_WEB_PUBLISH_ITEMS: usize = 4096;
const MAX_WEB_PUBLISH_STRING_BYTES: usize = 64 * 1024;
const MAX_EXTENSIONS: usize = 1024;
const MAX_EXTENSION_URI_BYTES: usize = 1024;
const MAX_EXTENSION_PAYLOAD_BYTES: usize = 1024 * 1024;
const MAX_RETAINED_EXTENSION_BYTES: usize = 4 * 1024 * 1024;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Conformance {
    Transitional,
    Strict,
}

impl Conformance {
    pub fn sml(self) -> &'static str {
        match self {
            Self::Transitional => SML,
            Self::Strict => STRICT_SML,
        }
    }
    pub fn rel(self) -> &'static str {
        match self {
            Self::Transitional => REL,
            Self::Strict => STRICT_REL,
        }
    }
    pub fn xdr(self) -> &'static str {
        match self {
            Self::Transitional => XDR,
            Self::Strict => STRICT_XDR,
        }
    }
    pub fn chart(self) -> &'static str {
        match self {
            Self::Transitional => CHART,
            Self::Strict => STRICT_CHART,
        }
    }
    pub fn chartsheet_rel(self) -> &'static str {
        match self {
            Self::Transitional => CHARTSHEET_REL,
            Self::Strict => STRICT_CHARTSHEET_REL,
        }
    }
    pub fn drawing_rel(self) -> &'static str {
        match self {
            Self::Transitional => DRAWING_REL,
            Self::Strict => STRICT_DRAWING_REL,
        }
    }
    pub fn chart_rel(self) -> &'static str {
        match self {
            Self::Transitional => CHART_REL,
            Self::Strict => STRICT_CHART_REL,
        }
    }
    pub fn image_rel(self) -> &'static str {
        match self {
            Self::Transitional => IMAGE_REL,
            Self::Strict => STRICT_IMAGE_REL,
        }
    }
    pub fn chart_user_shapes_rel(self) -> &'static str {
        match self {
            Self::Transitional => CHART_USER_SHAPES_REL,
            Self::Strict => STRICT_CHART_USER_SHAPES_REL,
        }
    }
    pub fn theme_override_rel(self) -> &'static str {
        match self {
            Self::Transitional => THEME_OVERRIDE_REL,
            Self::Strict => STRICT_THEME_OVERRIDE_REL,
        }
    }
    pub fn package_rel(self) -> &'static str {
        match self {
            Self::Transitional => PACKAGE_REL,
            Self::Strict => STRICT_PACKAGE_REL,
        }
    }
    pub fn vml_drawing_rel(self) -> &'static str {
        match self {
            Self::Transitional => VML_DRAWING_REL,
            Self::Strict => STRICT_VML_DRAWING_REL,
        }
    }
    pub fn printer_rel(self) -> &'static str {
        match self {
            Self::Transitional => PRINTER_REL,
            Self::Strict => STRICT_PRINTER_REL,
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum State {
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
pub struct Color {
    pub automatic: Option<bool>,
    pub indexed: Option<u32>,
    pub rgb: Option<String>,
    pub theme: Option<u32>,
    pub tint: Option<f64>,
}

#[derive(Debug, Clone, PartialEq)]
pub struct Properties {
    pub published: Option<bool>,
    pub code_name: Option<String>,
    pub tab_color: Option<Color>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct View {
    pub tab_selected: Option<bool>,
    pub zoom_scale: Option<u32>,
    pub workbook_view_id: u32,
    pub zoom_to_fit: Option<bool>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Protection {
    pub password_hash: Option<String>,
    pub content: Option<bool>,
    pub objects: Option<bool>,
}

/// One saved chartsheet view from `CT_CustomChartsheetView`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CustomView {
    /// Braced UUID lexical form required by SpreadsheetML `ST_Guid`.
    pub guid: String,
    pub scale: Option<u32>,
    pub state: Option<State>,
    pub zoom_to_fit: Option<bool>,
}

#[derive(Debug, Clone, Copy, PartialEq)]
pub struct Margins {
    pub left: f64,
    pub right: f64,
    pub top: f64,
    pub bottom: f64,
    pub header: f64,
    pub footer: f64,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PageSetup {
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
pub struct HeaderFooter {
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
pub enum WebSourceType {
    Sheet,
    PrintArea,
    AutoFilter,
    Range,
    Chart,
    PivotTable,
    Query,
    Label,
}

impl WebSourceType {
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
pub struct WebPublishItem {
    pub id: u32,
    pub div_id: String,
    pub source_type: WebSourceType,
    pub source_ref: Option<String>,
    pub source_object: Option<String>,
    pub destination_file: String,
    pub title: Option<String>,
    /// `None` preserves the schema default; no publishing action is ever performed.
    pub auto_republish: Option<bool>,
}

/// A present `webPublishItems` collection is schema-required to be non-empty.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct WebPublishItems {
    /// Preserves explicit `count`; when present it must equal `items.len()`.
    pub count: Option<u32>,
    pub items: Vec<WebPublishItem>,
}

#[derive(Debug, Clone, PartialEq)]
pub struct Chart {
    pub properties: Option<Properties>,
    pub views: Vec<View>,
    pub protection: Option<Protection>,
    /// `None` preserves absence; a present collection must be non-empty.
    pub custom_views: Option<Vec<CustomView>>,
    pub margins: Option<Margins>,
    pub page_setup: Option<PageSetup>,
    pub header_footer: Option<HeaderFooter>,
    pub drawing_relationship_id: String,
    pub legacy_drawing_relationship_id: Option<String>,
    pub legacy_header_footer_drawing_relationship_id: Option<String>,
    /// Relationship for the optional tiled chartsheet background image.
    pub background_picture_relationship_id: Option<String>,
    /// Inert web-publishing metadata. Destinations and sources are never resolved or accessed.
    pub web_publish_items: Option<WebPublishItems>,
    /// Inert, canonicalized wildcard extension markup. Payload semantics are never interpreted.
    pub extension_list: Option<ExtensionList>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Extension {
    pub uri: String,
    /// Canonical, namespace-aware XML for the single wildcard child of `ext`.
    pub payload_xml: Vec<u8>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ExtensionList {
    pub extensions: Vec<Extension>,
}

#[derive(Debug, Clone, PartialEq, Eq)]

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
pub fn parse_chartsheet(xml: &[u8]) -> Result<(Conformance, Chart)> {
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
            Ok::<_, Error>(required(picture, conformance.rel(), "id")?.to_owned())
        })
        .transpose()?;
    let web_publish_items = one_child(&root, conformance.sml(), "webPublishItems")?
        .map(|node| parse_web_publish_items(node, conformance))
        .transpose()?;
    let extension_list = one_child(&root, conformance.sml(), "extLst")?
        .map(|node| parse_extension_list(node, conformance))
        .transpose()?;
    let value = Chart {
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
    conformance: Conformance,
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

fn parse_web_publish_items(node: &Node, conformance: Conformance) -> Result<WebPublishItems> {
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
        items.push(WebPublishItem {
            id: required(child, "", "id")?
                .parse()
                .map_err(|_| invalid("invalid web publish item id"))?,
            div_id: required(child, "", "divId")?.to_owned(),
            source_type: WebSourceType::parse(required(child, "", "sourceType")?)?,
            source_ref: optional(child, "", "sourceRef").map(str::to_owned),
            source_object: optional(child, "", "sourceObject").map(str::to_owned),
            destination_file: required(child, "", "destinationFile")?.to_owned(),
            title: optional(child, "", "title").map(str::to_owned),
            auto_republish: bool_optional(child, "autoRepublish")?,
        });
    }
    let value = WebPublishItems { count, items };
    validate_web_publish_items(&value)?;
    Ok(value)
}

fn parse_extension_list(node: &Node, conformance: Conformance) -> Result<ExtensionList> {
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
        let payload = child
            .children
            .first()
            .ok_or_else(|| invalid("extension payload is missing its root"))?;
        let payload_xml = canonical_extension_payload_node(payload, conformance)?;
        total = total
            .checked_add(payload_xml.len())
            .ok_or_else(|| limit("retained extension bytes"))?;
        if total > MAX_RETAINED_EXTENSION_BYTES {
            return Err(limit("retained extension bytes"));
        }
        extensions.push(Extension { uri, payload_xml });
    }
    Ok(ExtensionList { extensions })
}

fn parse_properties(node: &Node) -> Result<Properties> {
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
    Ok(Properties {
        published,
        code_name,
        tab_color,
    })
}

fn parse_color(node: &Node) -> Result<Color> {
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
    Ok(Color {
        automatic: bool_optional(node, "auto")?,
        indexed: u32_optional(node, "indexed")?,
        rgb: optional(node, "", "rgb").map(str::to_owned),
        theme: u32_optional(node, "theme")?,
        tint: optional(node, "", "tint")
            .map(|v| v.parse().map_err(|_| invalid("invalid tab color tint")))
            .transpose()?,
    })
}

fn parse_views(node: &Node) -> Result<Vec<View>> {
    whitespace(node)?;
    no_attributes(node, &[])?;
    if node.children.is_empty() {
        return Err(invalid("sheetViews requires at least one sheetView"));
    }
    if node.children.len() > MAX_VIEWS {
        return Err(limit("view count"));
    }
    let mut views = Vec::with_capacity(node.children.len());
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
        views.push(View {
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

fn parse_protection(node: &Node) -> Result<Protection> {
    leaf(node, "chartsheet protection")?;
    no_attributes(node, &[("", "password"), ("", "content"), ("", "objects")])?;
    Ok(Protection {
        password_hash: optional(node, "", "password").map(str::to_owned),
        content: bool_optional(node, "content")?,
        objects: bool_optional(node, "objects")?,
    })
}

fn parse_custom_views(node: &Node, conformance: Conformance) -> Result<Vec<CustomView>> {
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
        values.push(CustomView {
            guid: required(child, "", "guid")?.to_owned(),
            scale: u32_optional(child, "scale")?,
            state: optional(child, "", "state").map(parse_state).transpose()?,
            zoom_to_fit: bool_optional(child, "zoomToFit")?,
        });
    }
    Ok(values)
}

fn parse_margins(node: &Node) -> Result<Margins> {
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
    Ok(Margins {
        left: number("left")?,
        right: number("right")?,
        top: number("top")?,
        bottom: number("bottom")?,
        header: number("header")?,
        footer: number("footer")?,
    })
}

fn parse_page_setup(node: &Node, conformance: Conformance) -> Result<PageSetup> {
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
    Ok(PageSetup {
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

fn parse_header_footer(node: &Node) -> Result<HeaderFooter> {
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
    let mut value = HeaderFooter {
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
pub fn write_chartsheet(value: &Chart, conformance: Conformance) -> Result<Vec<u8>> {
    validate_chartsheet(value)?;
    let mut out = BoundedXml::new(MAX_XML_BYTES);
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
                        State::Visible => "visible",
                        State::Hidden => "hidden",
                        State::VeryHidden => "veryHidden",
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
    out.finish("serialized XML bytes")
}

fn write_relationship_leaf<T: XmlOutput>(out: &mut T, name: &str, id: &str) {
    out.extend_from_slice(b"<x:");
    out.extend_from_slice(name.as_bytes());
    attr(out, "r:id", id);
    out.extend_from_slice(b"/>");
}

pub fn validate_chartsheet(value: &Chart) -> Result<()> {
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

fn validate_web_publish_items(value: &WebPublishItems) -> Result<()> {
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
        if item.source_type == WebSourceType::Range
            && item.source_ref.as_deref().is_none_or(str::is_empty)
        {
            return Err(invalid(
                "range web publish item requires non-empty sourceRef",
            ));
        }
        if matches!(
            item.source_type,
            WebSourceType::PivotTable | WebSourceType::Query | WebSourceType::Label
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

fn validate_extension_list(value: &ExtensionList) -> Result<()> {
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

fn canonical_extension_payload(xml: &[u8], conformance: Conformance) -> Result<Vec<u8>> {
    let node = parse_document(xml, MAX_EXTENSION_PAYLOAD_BYTES)?;
    canonical_extension_payload_node(&node, conformance)
}

fn canonical_extension_payload_node(node: &Node, conformance: Conformance) -> Result<Vec<u8>> {
    let mut namespaces = BTreeSet::new();
    collect_namespaces(node, &mut namespaces)?;
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
    let mut out = BoundedXml::new(MAX_EXTENSION_PAYLOAD_BYTES);
    write_canonical_node(&mut out, node, &prefixes, true)?;
    out.finish("extension payload bytes")
}

fn collect_namespaces(node: &Node, namespaces: &mut BTreeSet<String>) -> Result<()> {
    namespaces.insert(node.namespace.clone());
    if namespaces.len() > MAX_NAMESPACE_BINDINGS {
        return Err(limit("canonical namespace count"));
    }
    for attribute in &node.attributes {
        namespaces.insert(attribute.namespace.clone());
        if namespaces.len() > MAX_NAMESPACE_BINDINGS {
            return Err(limit("canonical namespace count"));
        }
    }
    for child in &node.children {
        collect_namespaces(child, namespaces)?;
    }
    Ok(())
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
fn write_canonical_node<T: XmlOutput>(
    out: &mut T,
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
                let child = node
                    .children
                    .get(*index)
                    .ok_or_else(|| invalid("invalid canonical XML child index"))?;
                write_canonical_node(out, child, prefixes, false)?
            },
        }
    }
    out.extend_from_slice(b"</");
    out.extend_from_slice(name.as_bytes());
    out.push(b'>');
    Ok(())
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
        max_namespace_bindings: MAX_NAMESPACE_BINDINGS,
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
                nodes = nodes
                    .checked_add(1)
                    .ok_or_else(|| limit("XML node count"))?;
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
fn root_conformance(root: &Node, name: &str) -> Result<Conformance> {
    if root.name != name {
        return Err(invalid(format!("expected {name} root")));
    }
    match root.namespace.as_str() {
        SML => Ok(Conformance::Transitional),
        STRICT_SML => Ok(Conformance::Strict),
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
fn parse_state(value: &str) -> Result<State> {
    match value {
        "visible" => Ok(State::Visible),
        "hidden" => Ok(State::Hidden),
        "veryHidden" => Ok(State::VeryHidden),
        _ => Err(invalid("invalid workbook sheet state")),
    }
}
fn validate_guid(value: &str) -> Result<()> {
    if !valid_guid(value) {
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

trait XmlOutput {
    fn push(&mut self, value: u8);
    fn extend_from_slice(&mut self, value: &[u8]);
}

impl XmlOutput for Vec<u8> {
    fn push(&mut self, value: u8) {
        Vec::push(self, value);
    }

    fn extend_from_slice(&mut self, value: &[u8]) {
        Vec::extend_from_slice(self, value);
    }
}

struct BoundedXml {
    bytes: Vec<u8>,
    max: usize,
    overflowed: bool,
}

impl BoundedXml {
    fn new(max: usize) -> Self {
        Self {
            bytes: Vec::with_capacity(max.min(8192)),
            max,
            overflowed: false,
        }
    }

    fn finish(self, label: &str) -> Result<Vec<u8>> {
        if self.overflowed {
            Err(limit(label))
        } else {
            Ok(self.bytes)
        }
    }
}

impl XmlOutput for BoundedXml {
    fn push(&mut self, value: u8) {
        if self.bytes.len() < self.max {
            self.bytes.push(value);
        } else {
            self.overflowed = true;
        }
    }

    fn extend_from_slice(&mut self, value: &[u8]) {
        let remaining = self.max.saturating_sub(self.bytes.len());
        if value.len() > remaining {
            self.overflowed = true;
            self.bytes.extend_from_slice(&value[..remaining]);
        } else {
            self.bytes.extend_from_slice(value);
        }
    }
}

fn bool_attr_opt<T: XmlOutput>(out: &mut T, name: &str, value: Option<bool>) {
    if let Some(value) = value {
        attr(out, name, if value { "1" } else { "0" });
    }
}
fn u32_attr_opt<T: XmlOutput>(out: &mut T, name: &str, value: Option<u32>) {
    if let Some(value) = value {
        attr(out, name, &value.to_string());
    }
}
fn attr_opt<T: XmlOutput>(out: &mut T, name: &str, value: Option<&str>) {
    if let Some(value) = value {
        attr(out, name, value);
    }
}
fn attr<T: XmlOutput>(out: &mut T, name: &str, value: &str) {
    out.push(b' ');
    out.extend_from_slice(name.as_bytes());
    out.extend_from_slice(b"=\"");
    escape(out, value);
    out.push(b'\"');
}
fn escape<T: XmlOutput>(out: &mut T, value: &str) {
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
fn escape_text<T: XmlOutput>(out: &mut T, value: &str) {
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
fn xml_error(error: impl std::fmt::Display) -> Error {
    Error::Xml(litchi_ooxml_common::XmlError::Malformed(error.to_string()))
}
fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}
fn limit(name: &str) -> Error {
    invalid(format!("chartsheet {name} limit exceeded"))
}
