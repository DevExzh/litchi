//! Typed SpreadsheetML chartsheets and their inert workbook/drawing/chart package graph.

use crate::common::{MceCapabilities, MceLimits, process_markup_compatibility};
use crate::error::{OoxmlError, Result};
use litchi_opc::constants::relationship_type as rt;
use litchi_opc::{BlobPart, OpcPackage, PackURI, Part};
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use std::collections::{BTreeMap, HashSet};

const SML: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const STRICT_SML: &str = "http://purl.oclc.org/ooxml/spreadsheetml/main";
const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const STRICT_REL: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships";
const XDR: &str = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
const STRICT_XDR: &str = "http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing";
const CHART: &str = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const STRICT_CHART: &str = "http://purl.oclc.org/ooxml/drawingml/chart";
const CHARTSHEET_REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet";
const STRICT_CHARTSHEET_REL: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships/chartsheet";
const IMAGE_REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
const STRICT_IMAGE_REL: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships/image";
const CHARTSHEET_CT: &str = "application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml";
const DRAWING_CT: &str = "application/vnd.openxmlformats-officedocument.drawing+xml";
const CHART_CT: &str = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";
const MAX_XML_BYTES: usize = 32 * 1024 * 1024;
const MAX_DRAWING_BYTES: usize = 16 * 1024 * 1024;
const MAX_CHART_BYTES: usize = 32 * 1024 * 1024;
const MAX_BACKGROUND_IMAGE_BYTES: usize = 32 * 1024 * 1024;
const MAX_TOTAL_RESOURCE_BYTES: usize = 128 * 1024 * 1024;
const MAX_NODES: usize = 500_000;
const MAX_DEPTH: usize = 256;
const MAX_STRING_BYTES: usize = 4 * 1024 * 1024;
const MAX_VIEWS: usize = 256;
const MAX_CUSTOM_VIEWS: usize = 1024;
const MAX_CHARTS: usize = 256;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ChartSheetConformance { Transitional, Strict }

impl ChartSheetConformance {
    fn sml(self) -> &'static str { match self { Self::Transitional => SML, Self::Strict => STRICT_SML } }
    fn rel(self) -> &'static str { match self { Self::Transitional => REL, Self::Strict => STRICT_REL } }
    fn xdr(self) -> &'static str { match self { Self::Transitional => XDR, Self::Strict => STRICT_XDR } }
    fn chart(self) -> &'static str { match self { Self::Transitional => CHART, Self::Strict => STRICT_CHART } }
    fn chartsheet_rel(self) -> &'static str { match self { Self::Transitional => CHARTSHEET_REL, Self::Strict => STRICT_CHARTSHEET_REL } }
    fn drawing_rel(self) -> &'static str { match self { Self::Transitional => rt::DRAWING, Self::Strict => rt::STRICT_DRAWING } }
    fn chart_rel(self) -> &'static str { match self { Self::Transitional => rt::CHART, Self::Strict => rt::STRICT_CHART } }
    fn image_rel(self) -> &'static str { match self { Self::Transitional => IMAGE_REL, Self::Strict => STRICT_IMAGE_REL } }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ChartSheetState { Visible, Hidden, VeryHidden }

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PageOrientation { Default, Portrait, Landscape }

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
    pub left: f64, pub right: f64, pub top: f64, pub bottom: f64, pub header: f64, pub footer: f64,
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
}

#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct ChartSheetHeaderFooter {
    pub different_odd_even: Option<bool>,
    pub different_first: Option<bool>,
    pub scale_with_document: Option<bool>,
    pub align_with_margins: Option<bool>,
    pub odd_header: Option<String>, pub odd_footer: Option<String>,
    pub even_header: Option<String>, pub even_footer: Option<String>,
    pub first_header: Option<String>, pub first_footer: Option<String>,
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
    /// Relationship for the optional tiled chartsheet background image.
    pub background_picture_relationship_id: Option<String>,
}

/// Supported inert image media types for chartsheet backgrounds.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ChartSheetImageContentType { Png, Jpeg, Gif, Bmp, Tiff, Emf, Wmf }

impl ChartSheetImageContentType {
    pub fn as_str(self) -> &'static str { match self { Self::Png => "image/png", Self::Jpeg => "image/jpeg", Self::Gif => "image/gif", Self::Bmp => "image/bmp", Self::Tiff => "image/tiff", Self::Emf => "image/x-emf", Self::Wmf => "image/x-wmf" } }
    fn parse(value:&str)->Result<Self>{match value{"image/png"=>Ok(Self::Png),"image/jpeg"=>Ok(Self::Jpeg),"image/gif"=>Ok(Self::Gif),"image/bmp"=>Ok(Self::Bmp),"image/tiff"=>Ok(Self::Tiff),"image/x-emf"=>Ok(Self::Emf),"image/x-wmf"=>Ok(Self::Wmf),_=>Err(invalid(format!("unsupported chartsheet background image content type '{value}'")))}}
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

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChartSheetChartResource {
    pub relationship_id: String,
    pub part_name: String,
    pub content_type: String,
    /// Preserved without chart evaluation, external-data loading, or macro execution.
    pub data: Vec<u8>,
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
    pub background_picture: Option<ChartSheetBackgroundPicture>,
}

#[derive(Clone)]
struct Attribute { namespace: String, name: String, value: String }
#[derive(Clone)]
struct Node { namespace: String, name: String, attributes: Vec<Attribute>, children: Vec<Node>, text: String }

/// Parses the selected, bounded core of a complete Chartsheet part.
pub fn parse_chartsheet(xml: &[u8]) -> Result<(ChartSheetConformance, ChartSheet)> {
    let root = parse_document(xml, MAX_XML_BYTES)?;
    let conformance = root_conformance(&root, "chartsheet")?;
    whitespace(&root)?; no_attributes(&root, &[])?;
    validate_root_order(&root)?;
    let properties = one_child(&root, conformance.sml(), "sheetPr")?.map(|node| parse_properties(node)).transpose()?;
    let views = parse_views(required_child(&root, conformance.sml(), "sheetViews")?)?;
    let protection = one_child(&root, conformance.sml(), "sheetProtection")?.map(parse_protection).transpose()?;
    let custom_views = one_child(&root, conformance.sml(), "customSheetViews")?.map(|node| parse_custom_views(node, conformance)).transpose()?;
    let margins = one_child(&root, conformance.sml(), "pageMargins")?.map(parse_margins).transpose()?;
    let page_setup = one_child(&root, conformance.sml(), "pageSetup")?.map(parse_page_setup).transpose()?;
    let header_footer = one_child(&root, conformance.sml(), "headerFooter")?.map(parse_header_footer).transpose()?;
    let drawing = required_child(&root, conformance.sml(), "drawing")?; leaf(drawing, "chartsheet drawing")?;
    let drawing_relationship_id = required(drawing, conformance.rel(), "id")?.to_owned(); no_attributes(drawing, &[(conformance.rel(), "id")])?;
    let background_picture_relationship_id = one_child(&root, conformance.sml(), "picture")?.map(|picture| { leaf(picture,"chartsheet picture")?; no_attributes(picture,&[(conformance.rel(),"id")])?; Ok::<_,OoxmlError>(required(picture,conformance.rel(),"id")?.to_owned()) }).transpose()?;
    let value = ChartSheet { properties, views, protection, custom_views, margins, page_setup, header_footer, drawing_relationship_id, background_picture_relationship_id };
    validate_chartsheet(&value)?; Ok((conformance, value))
}

fn validate_root_order(root: &Node) -> Result<()> {
    let mut last = 0u8;
    for child in &root.children {
        let order = match child.name.as_str() { "sheetPr" => 1, "sheetViews" => 2, "sheetProtection" => 3, "customSheetViews" => 4, "pageMargins" => 5, "pageSetup" => 6, "headerFooter" => 7, "drawing" => 8, "picture" => 9, name => return Err(invalid(format!("unsupported chartsheet child '{name}'"))) };
        if order <= last { return Err(invalid("chartsheet children are duplicated or out of schema order")); } last = order;
    }
    Ok(())
}

fn parse_properties(node: &Node) -> Result<ChartSheetProperties> {
    whitespace(node)?; no_attributes(node, &[("", "published"), ("", "codeName")])?;
    let published = optional(node, "", "published").map(|v| parse_bool(v, "published")).transpose()?;
    let code_name = optional(node, "", "codeName").map(str::to_owned);
    let tab_color = one_child_any_core(node, "tabColor")?.map(parse_color).transpose()?;
    if node.children.len() > usize::from(tab_color.is_some()) { return Err(invalid("sheetPr contains unsupported children")); }
    Ok(ChartSheetProperties { published, code_name, tab_color })
}

fn parse_color(node: &Node) -> Result<ChartSheetColor> {
    leaf(node, "tab color")?; no_attributes(node, &[("", "auto"), ("", "indexed"), ("", "rgb"), ("", "theme"), ("", "tint")])?;
    Ok(ChartSheetColor { automatic: bool_optional(node, "auto")?, indexed: u32_optional(node, "indexed")?, rgb: optional(node, "", "rgb").map(str::to_owned), theme: u32_optional(node, "theme")?, tint: optional(node, "", "tint").map(|v| v.parse().map_err(|_| invalid("invalid tab color tint"))).transpose()? })
}

fn parse_views(node: &Node) -> Result<Vec<ChartSheetView>> {
    whitespace(node)?; no_attributes(node, &[])?; if node.children.is_empty() { return Err(invalid("sheetViews requires at least one sheetView")); } if node.children.len() > MAX_VIEWS { return Err(limit("view count")); }
    let mut views = Vec::new();
    for child in &node.children {
        if child.name != "sheetView" || !is_core(&child.namespace) { return Err(invalid("sheetViews contains an unsupported child")); }
        leaf(child, "chartsheet view")?; no_attributes(child, &[("", "tabSelected"), ("", "zoomScale"), ("", "workbookViewId"), ("", "zoomToFit")])?;
        views.push(ChartSheetView { tab_selected: bool_optional(child, "tabSelected")?, zoom_scale: u32_optional(child, "zoomScale")?, workbook_view_id: required(child, "", "workbookViewId")?.parse().map_err(|_| invalid("invalid workbookViewId"))?, zoom_to_fit: bool_optional(child, "zoomToFit")? });
    }
    Ok(views)
}

fn parse_protection(node: &Node) -> Result<ChartSheetProtection> {
    leaf(node, "chartsheet protection")?; no_attributes(node, &[("", "password"), ("", "content"), ("", "objects")])?;
    Ok(ChartSheetProtection { password_hash: optional(node, "", "password").map(str::to_owned), content: bool_optional(node, "content")?, objects: bool_optional(node, "objects")? })
}

fn parse_custom_views(node: &Node, conformance: ChartSheetConformance) -> Result<Vec<ChartSheetCustomView>> {
    whitespace(node)?; no_attributes(node, &[])?;
    if node.children.is_empty() { return Err(invalid("customSheetViews requires at least one customSheetView")); }
    if node.children.len() > MAX_CUSTOM_VIEWS { return Err(limit("custom view count")); }
    let mut values = Vec::with_capacity(node.children.len());
    for child in &node.children {
        if child.namespace != conformance.sml() || child.name != "customSheetView" { return Err(invalid("customSheetViews contains an unsupported child")); }
        leaf(child, "custom chartsheet view")?;
        no_attributes(child, &[("", "guid"), ("", "scale"), ("", "state"), ("", "zoomToFit")])?;
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
    leaf(node, "chartsheet margins")?; no_attributes(node, &[("", "left"), ("", "right"), ("", "top"), ("", "bottom"), ("", "header"), ("", "footer")])?;
    let number = |name| required(node, "", name)?.parse().map_err(|_| invalid(format!("invalid {name} page margin")));
    Ok(ChartSheetMargins { left: number("left")?, right: number("right")?, top: number("top")?, bottom: number("bottom")?, header: number("header")?, footer: number("footer")? })
}

fn parse_page_setup(node: &Node) -> Result<ChartSheetPageSetup> {
    leaf(node, "chartsheet page setup")?; no_attributes(node, &[("", "paperSize"), ("", "firstPageNumber"), ("", "orientation"), ("", "usePrinterDefaults"), ("", "blackAndWhite"), ("", "draft"), ("", "useFirstPageNumber"), ("", "horizontalDpi"), ("", "verticalDpi"), ("", "copies")])?;
    Ok(ChartSheetPageSetup { paper_size: u32_optional(node, "paperSize")?, first_page_number: u32_optional(node, "firstPageNumber")?, orientation: optional(node, "", "orientation").map(parse_orientation).transpose()?, use_printer_defaults: bool_optional(node, "usePrinterDefaults")?, black_and_white: bool_optional(node, "blackAndWhite")?, draft: bool_optional(node, "draft")?, use_first_page_number: bool_optional(node, "useFirstPageNumber")?, horizontal_dpi: u32_optional(node, "horizontalDpi")?, vertical_dpi: u32_optional(node, "verticalDpi")?, copies: u32_optional(node, "copies")? })
}

fn parse_header_footer(node: &Node) -> Result<ChartSheetHeaderFooter> {
    whitespace(node)?; no_attributes(node, &[("", "differentOddEven"), ("", "differentFirst"), ("", "scaleWithDoc"), ("", "alignWithMargins")])?;
    let mut value = ChartSheetHeaderFooter { different_odd_even: bool_optional(node, "differentOddEven")?, different_first: bool_optional(node, "differentFirst")?, scale_with_document: bool_optional(node, "scaleWithDoc")?, align_with_margins: bool_optional(node, "alignWithMargins")?, ..Default::default() };
    let mut last = 0u8;
    for child in &node.children {
        if !is_core(&child.namespace) { return Err(invalid("headerFooter has a foreign child")); } leaf(child, "header/footer text")?; no_attributes(child, &[])?;
        let (order, target) = match child.name.as_str() { "oddHeader" => (1, &mut value.odd_header), "oddFooter" => (2, &mut value.odd_footer), "evenHeader" => (3, &mut value.even_header), "evenFooter" => (4, &mut value.even_footer), "firstHeader" => (5, &mut value.first_header), "firstFooter" => (6, &mut value.first_footer), _ => return Err(invalid("unsupported headerFooter child")) };
        if order <= last { return Err(invalid("headerFooter children are duplicated or out of schema order")); } last = order; *target = Some(child.text.clone());
    }
    Ok(value)
}

/// Deterministically serializes one complete Chartsheet part.
pub fn write_chartsheet(value: &ChartSheet, conformance: ChartSheetConformance) -> Result<Vec<u8>> {
    validate_chartsheet(value)?; let mut out = Vec::new();
    out.extend_from_slice(b"<x:chartsheet xmlns:x=\""); escape(&mut out, conformance.sml()); out.extend_from_slice(b"\" xmlns:r=\""); escape(&mut out, conformance.rel()); out.extend_from_slice(b"\">");
    if let Some(properties) = &value.properties { out.extend_from_slice(b"<x:sheetPr"); bool_attr_opt(&mut out, "published", properties.published); attr_opt(&mut out, "codeName", properties.code_name.as_deref()); if let Some(color) = &properties.tab_color { out.push(b'>'); out.extend_from_slice(b"<x:tabColor"); bool_attr_opt(&mut out, "auto", color.automatic); u32_attr_opt(&mut out, "indexed", color.indexed); attr_opt(&mut out, "rgb", color.rgb.as_deref()); u32_attr_opt(&mut out, "theme", color.theme); if let Some(v) = color.tint { attr(&mut out, "tint", &v.to_string()); } out.extend_from_slice(b"/></x:sheetPr>"); } else { out.extend_from_slice(b"/>"); } }
    out.extend_from_slice(b"<x:sheetViews>"); for view in &value.views { out.extend_from_slice(b"<x:sheetView"); bool_attr_opt(&mut out, "tabSelected", view.tab_selected); u32_attr_opt(&mut out, "zoomScale", view.zoom_scale); attr(&mut out, "workbookViewId", &view.workbook_view_id.to_string()); bool_attr_opt(&mut out, "zoomToFit", view.zoom_to_fit); out.extend_from_slice(b"/>"); } out.extend_from_slice(b"</x:sheetViews>");
    if let Some(protection) = &value.protection { out.extend_from_slice(b"<x:sheetProtection"); attr_opt(&mut out, "password", protection.password_hash.as_deref()); bool_attr_opt(&mut out, "content", protection.content); bool_attr_opt(&mut out, "objects", protection.objects); out.extend_from_slice(b"/>"); }
    if let Some(custom_views) = &value.custom_views { out.extend_from_slice(b"<x:customSheetViews>"); for view in custom_views { out.extend_from_slice(b"<x:customSheetView"); attr(&mut out, "guid", &view.guid); u32_attr_opt(&mut out, "scale", view.scale); if let Some(state) = view.state { attr(&mut out, "state", match state { ChartSheetState::Visible => "visible", ChartSheetState::Hidden => "hidden", ChartSheetState::VeryHidden => "veryHidden" }); } bool_attr_opt(&mut out, "zoomToFit", view.zoom_to_fit); out.extend_from_slice(b"/>"); } out.extend_from_slice(b"</x:customSheetViews>"); }
    if let Some(m) = value.margins { out.extend_from_slice(b"<x:pageMargins"); for (name, value) in [("left", m.left), ("right", m.right), ("top", m.top), ("bottom", m.bottom), ("header", m.header), ("footer", m.footer)] { attr(&mut out, name, &value.to_string()); } out.extend_from_slice(b"/>"); }
    if let Some(setup) = &value.page_setup { out.extend_from_slice(b"<x:pageSetup"); u32_attr_opt(&mut out, "paperSize", setup.paper_size); u32_attr_opt(&mut out, "firstPageNumber", setup.first_page_number); if let Some(v) = setup.orientation { attr(&mut out, "orientation", match v { PageOrientation::Default => "default", PageOrientation::Portrait => "portrait", PageOrientation::Landscape => "landscape" }); } bool_attr_opt(&mut out, "usePrinterDefaults", setup.use_printer_defaults); bool_attr_opt(&mut out, "blackAndWhite", setup.black_and_white); bool_attr_opt(&mut out, "draft", setup.draft); bool_attr_opt(&mut out, "useFirstPageNumber", setup.use_first_page_number); u32_attr_opt(&mut out, "horizontalDpi", setup.horizontal_dpi); u32_attr_opt(&mut out, "verticalDpi", setup.vertical_dpi); u32_attr_opt(&mut out, "copies", setup.copies); out.extend_from_slice(b"/>"); }
    if let Some(hf) = &value.header_footer { out.extend_from_slice(b"<x:headerFooter"); bool_attr_opt(&mut out, "differentOddEven", hf.different_odd_even); bool_attr_opt(&mut out, "differentFirst", hf.different_first); bool_attr_opt(&mut out, "scaleWithDoc", hf.scale_with_document); bool_attr_opt(&mut out, "alignWithMargins", hf.align_with_margins); let children = [("oddHeader", &hf.odd_header), ("oddFooter", &hf.odd_footer), ("evenHeader", &hf.even_header), ("evenFooter", &hf.even_footer), ("firstHeader", &hf.first_header), ("firstFooter", &hf.first_footer)]; if children.iter().all(|(_, value)| value.is_none()) { out.extend_from_slice(b"/>"); } else { out.push(b'>'); for (name, value) in children { if let Some(value) = value { out.extend_from_slice(b"<x:"); out.extend_from_slice(name.as_bytes()); out.push(b'>'); escape_text(&mut out, value); out.extend_from_slice(b"</x:"); out.extend_from_slice(name.as_bytes()); out.push(b'>'); } } out.extend_from_slice(b"</x:headerFooter>"); } }
    out.extend_from_slice(b"<x:drawing"); attr(&mut out, "r:id", &value.drawing_relationship_id); out.extend_from_slice(b"/>");
    if let Some(id)=&value.background_picture_relationship_id{out.extend_from_slice(b"<x:picture");attr(&mut out,"r:id",id);out.extend_from_slice(b"/>");}
    out.extend_from_slice(b"</x:chartsheet>");
    if out.len() > MAX_XML_BYTES { return Err(limit("serialized XML bytes")); } Ok(out)
}

/// Loads one workbook-referenced chartsheet and validates its bounded leaf graph.
pub fn load_chartsheet(package: &OpcPackage, workbook_name: &PackURI, workbook_relationship_id: &str) -> Result<ChartSheetPackage> {
    if package.rels().iter().any(|rel| matches!(rel.reltype(), CHARTSHEET_REL | STRICT_CHARTSHEET_REL)) { return Err(invalid("package root cannot source a chartsheet relationship")); }
    let workbook = package.get_part(workbook_name)?; require_workbook(workbook)?; let workbook_root = parse_document(workbook.blob(), MAX_XML_BYTES)?; let conformance = root_conformance(&workbook_root, "workbook")?;
    let relationship = internal_relationship(workbook, workbook_relationship_id, conformance.chartsheet_rel())?; let chartsheet_name = relationship.target_partname()?;
    if !chartsheet_name.as_str().starts_with("/xl/chartsheets/") { return Err(invalid("chartsheet target is outside /xl/chartsheets")); }
    let entry = workbook_entry(&workbook_root, conformance, workbook_relationship_id, chartsheet_name.to_string())?;
    let chartsheet_part = package.get_part(&chartsheet_name)?; require_content_type(chartsheet_part, CHARTSHEET_CT, "chartsheet")?; let (part_conformance, chartsheet) = parse_chartsheet(chartsheet_part.blob())?; if part_conformance != conformance { return Err(invalid("workbook and chartsheet conformance differ")); }
    let drawing_rel = internal_relationship(chartsheet_part, &chartsheet.drawing_relationship_id, conformance.drawing_rel())?; let drawing_name = drawing_rel.target_partname()?; if !drawing_name.as_str().starts_with("/xl/drawings/") { return Err(invalid("chartsheet drawing is outside /xl/drawings")); }
    let background_picture=if let Some(id)=&chartsheet.background_picture_relationship_id{let rel=internal_relationship(chartsheet_part,id,conformance.image_rel())?;let name=rel.target_partname()?;if !name.as_str().starts_with("/xl/media/"){return Err(invalid("chartsheet background image is outside /xl/media"))}let part=package.get_part(&name)?;let content_type=ChartSheetImageContentType::parse(part.content_type())?;if part.blob().len()>MAX_BACKGROUND_IMAGE_BYTES{return Err(limit("background image bytes"))}if !part.rels().is_empty(){return Err(invalid("chartsheet background image must be a relationship-free leaf"))}Some(ChartSheetBackgroundPicture{relationship_id:id.clone(),part_name:name.to_string(),content_type,data:part.blob().to_vec()})}else{None};
    let expected_relationships=1+usize::from(background_picture.is_some());if chartsheet_part.rels().iter().count()!=expected_relationships{return Err(invalid("bounded chartsheet has unsupported or unreferenced relationships"))}
    let drawing_part = package.get_part(&drawing_name)?; require_content_type(drawing_part, DRAWING_CT, "drawing")?; if drawing_part.blob().len() > MAX_DRAWING_BYTES { return Err(limit("drawing bytes")); }
    let chart_ids = drawing_chart_ids(drawing_part.blob(), conformance)?; if drawing_part.rels().iter().count() != chart_ids.len() { return Err(invalid("bounded chartsheet drawing has unsupported or unreferenced relationships")); }
    let mut charts = Vec::new(); let mut total = drawing_part.blob().len();if let Some(picture)=&background_picture{add_resource(&mut total,picture.data.len(),MAX_BACKGROUND_IMAGE_BYTES,"background image bytes")?;}
    for id in chart_ids { let relationship = internal_relationship(drawing_part, &id, conformance.chart_rel())?; let name = relationship.target_partname()?; if !name.as_str().starts_with("/xl/charts/") { return Err(invalid("chart target is outside /xl/charts")); } let part = package.get_part(&name)?; require_content_type(part, CHART_CT, "chart")?; validate_chart_xml(part.blob(), conformance)?; if !part.rels().is_empty() { return Err(invalid("bounded chartsheet chart must be a relationship-free leaf")); } add_resource(&mut total, part.blob().len(), MAX_CHART_BYTES, "chart bytes")?; charts.push(ChartSheetChartResource { relationship_id: id, part_name: name.to_string(), content_type: part.content_type().to_owned(), data: part.blob().to_vec() }); }
    Ok(ChartSheetPackage { entry, chartsheet, drawing: ChartSheetDrawingResource { part_name: drawing_name.to_string(), content_type: drawing_part.content_type().to_owned(), data: drawing_part.blob().to_vec(), charts }, background_picture })
}

/// Adds a preflighted chartsheet package graph and workbook sheet entry.
pub fn store_chartsheet(package: &mut OpcPackage, workbook_name: &PackURI, value: &ChartSheetPackage, conformance: ChartSheetConformance) -> Result<()> {
    validate_package_value(value, conformance)?;
    let workbook = package.get_part(workbook_name)?; require_workbook(workbook)?; let workbook_root = parse_document(workbook.blob(), MAX_XML_BYTES)?; if root_conformance(&workbook_root, "workbook")? != conformance { return Err(invalid("requested conformance does not match workbook")); }
    validate_new_entry(&workbook_root, conformance, &value.entry)?;
    if workbook.rels().get(&value.entry.workbook_relationship_id).is_some() { return Err(invalid("workbook relationship ID already exists")); }
    let chartsheet_uri = new_uri(package, &value.entry.part_name, "/xl/chartsheets/")?; let drawing_uri = new_uri(package, &value.drawing.part_name, "/xl/drawings/")?;
    let picture_uri=value.background_picture.as_ref().map(|picture|new_uri(package,&picture.part_name,"/xl/media/")).transpose()?;
    let mut chart_uris = BTreeMap::new(); for chart in &value.drawing.charts { chart_uris.insert(chart.relationship_id.clone(), new_uri(package, &chart.part_name, "/xl/charts/")?); }
    let updated_workbook = insert_workbook_entry(workbook.blob(), &value.entry, conformance)?; let chartsheet_xml = write_chartsheet(&value.chartsheet, conformance)?;
    package.get_part_mut(workbook_name)?.set_blob(updated_workbook);
    package.add_part(Box::new(BlobPart::new(chartsheet_uri.clone(), CHARTSHEET_CT.into(), chartsheet_xml)));
    package.add_part(Box::new(BlobPart::new(drawing_uri.clone(), value.drawing.content_type.clone(), value.drawing.data.clone())));
    if let (Some(picture),Some(uri))=(&value.background_picture,&picture_uri){package.add_part(Box::new(BlobPart::new(uri.clone(),picture.content_type.as_str().into(),picture.data.clone())));}
    for chart in &value.drawing.charts { package.add_part(Box::new(BlobPart::new(chart_uris[&chart.relationship_id].clone(), chart.content_type.clone(), chart.data.clone()))); }
    package.get_part_mut(workbook_name)?.rels_mut().add_relationship(conformance.chartsheet_rel().into(), chartsheet_uri.relative_ref(workbook_name.base_uri()), value.entry.workbook_relationship_id.clone(), false);
    package.get_part_mut(&chartsheet_uri)?.rels_mut().add_relationship(conformance.drawing_rel().into(), drawing_uri.relative_ref(chartsheet_uri.base_uri()), value.chartsheet.drawing_relationship_id.clone(), false);
    if let (Some(picture),Some(uri))=(&value.background_picture,&picture_uri){package.get_part_mut(&chartsheet_uri)?.rels_mut().add_relationship(conformance.image_rel().into(),uri.relative_ref(chartsheet_uri.base_uri()),picture.relationship_id.clone(),false);}
    for chart in &value.drawing.charts { package.get_part_mut(&drawing_uri)?.rels_mut().add_relationship(conformance.chart_rel().into(), chart_uris[&chart.relationship_id].relative_ref(drawing_uri.base_uri()), chart.relationship_id.clone(), false); }
    Ok(())
}

fn validate_package_value(value: &ChartSheetPackage, conformance: ChartSheetConformance) -> Result<()> {
    validate_entry(&value.entry)?; validate_chartsheet(&value.chartsheet)?;
    if value.drawing.content_type != DRAWING_CT || value.drawing.data.len() > MAX_DRAWING_BYTES { return Err(invalid("invalid or oversized chartsheet drawing resource")); }
    let drawing_uri = PackURI::new(&value.drawing.part_name).map_err(OoxmlError::InvalidUri)?; if !drawing_uri.as_str().starts_with("/xl/drawings/") { return Err(invalid("drawing resource is outside /xl/drawings")); }
    let ids = drawing_chart_ids(&value.drawing.data, conformance)?; if ids.len() != value.drawing.charts.len() { return Err(invalid("drawing chart references and chart resources differ")); }
    let mut resources = BTreeMap::new(); let mut total = value.drawing.data.len();
    for chart in &value.drawing.charts { validate_id(&chart.relationship_id)?; if !ids.contains(&chart.relationship_id) { return Err(invalid(format!("drawing does not reference chart relationship '{}'", chart.relationship_id))); } let uri = PackURI::new(&chart.part_name).map_err(OoxmlError::InvalidUri)?; if !uri.as_str().starts_with("/xl/charts/") || chart.content_type != CHART_CT { return Err(invalid("invalid chart resource path or content type")); } validate_chart_xml(&chart.data, conformance)?; add_resource(&mut total, chart.data.len(), MAX_CHART_BYTES, "chart bytes")?; if resources.insert(chart.part_name.clone(), &chart.data).is_some() { return Err(invalid("duplicate chart resource part name")); } }
    match (&value.chartsheet.background_picture_relationship_id,&value.background_picture){(None,None)=>{},(Some(id),Some(picture))=>{validate_id(id)?;validate_id(&picture.relationship_id)?;if id!=&picture.relationship_id{return Err(invalid("chartsheet picture relationship and resource metadata differ"))}if id==&value.chartsheet.drawing_relationship_id{return Err(invalid("chartsheet drawing and picture relationship IDs collide"))}let uri=PackURI::new(&picture.part_name).map_err(OoxmlError::InvalidUri)?;if !uri.as_str().starts_with("/xl/media/"){return Err(invalid("background image resource is outside /xl/media"))}add_resource(&mut total,picture.data.len(),MAX_BACKGROUND_IMAGE_BYTES,"background image bytes")?;if resources.insert(picture.part_name.clone(),&picture.data).is_some(){return Err(invalid("duplicate chartsheet resource part name"))}},_=>return Err(invalid("chartsheet picture relationship and resource must either both be present or both be absent"))}
    Ok(())
}

fn validate_chartsheet(value: &ChartSheet) -> Result<()> {
    if value.views.is_empty() || value.views.len() > MAX_VIEWS { return Err(invalid("chartsheet requires a bounded non-empty view list")); } validate_id(&value.drawing_relationship_id)?;if let Some(id)=&value.background_picture_relationship_id{validate_id(id)?;if id==&value.drawing_relationship_id{return Err(invalid("chartsheet drawing and picture relationship IDs collide"))}}
    let mut view_ids = HashSet::new(); for view in &value.views { if view.zoom_scale.is_some_and(|v| !(10..=400).contains(&v)) { return Err(invalid("chartsheet zoomScale must be between 10 and 400")); } if !view_ids.insert(view.workbook_view_id) { return Err(invalid("duplicate chartsheet workbookViewId")); } }
    if let Some(properties) = &value.properties { if let Some(name) = &properties.code_name { bounded(name)?; } if let Some(color) = &properties.tab_color { let bases = usize::from(color.automatic.is_some()) + usize::from(color.indexed.is_some()) + usize::from(color.rgb.is_some()) + usize::from(color.theme.is_some()); if bases > 1 { return Err(invalid("tab color has multiple base color selectors")); } if let Some(rgb) = &color.rgb { if rgb.len() != 8 || !rgb.bytes().all(|b| b.is_ascii_hexdigit()) { return Err(invalid("tab color rgb must contain eight hex digits")); } } if color.tint.is_some_and(|v| !v.is_finite() || !(-1.0..=1.0).contains(&v)) { return Err(invalid("tab color tint is outside [-1, 1]")); } } }
    if let Some(protection) = &value.protection { if let Some(password) = &protection.password_hash { if password.len() != 4 || !password.bytes().all(|b| b.is_ascii_hexdigit()) { return Err(invalid("chartsheet password hash must contain four hex digits")); } } }
    if let Some(custom_views) = &value.custom_views {
        if custom_views.is_empty() { return Err(invalid("customSheetViews requires at least one customSheetView")); }
        if custom_views.len() > MAX_CUSTOM_VIEWS { return Err(limit("custom view count")); }
        let mut guids = HashSet::new();
        for view in custom_views {
            validate_guid(&view.guid)?;
            if !guids.insert(view.guid.to_ascii_lowercase()) { return Err(invalid(format!("duplicate custom chartsheet view GUID '{}'", view.guid))); }
            if view.scale.is_some_and(|scale| !(10..=400).contains(&scale)) { return Err(invalid("custom chartsheet view scale must be between 10 and 400")); }
        }
    }
    if let Some(m) = value.margins { for margin in [m.left, m.right, m.top, m.bottom, m.header, m.footer] { if !margin.is_finite() || !(0.0..49.0).contains(&margin) { return Err(invalid("chartsheet margin is outside Office's [0, 49) range")); } } }
    if let Some(setup) = &value.page_setup { if setup.first_page_number.is_some_and(|v| v > 65_534) { return Err(invalid("firstPageNumber exceeds Excel's limit")); } if setup.copies.is_some_and(|v| !(1..=32_767).contains(&v)) { return Err(invalid("copies is outside Excel's supported range")); } if setup.horizontal_dpi == Some(0) || setup.vertical_dpi == Some(0) { return Err(invalid("page setup DPI must be positive")); } }
    if let Some(hf) = &value.header_footer { for text in [&hf.odd_header, &hf.odd_footer, &hf.even_header, &hf.even_footer, &hf.first_header, &hf.first_footer].into_iter().flatten() { bounded(text)?; } }
    Ok(())
}

fn workbook_entry(root: &Node, conformance: ChartSheetConformance, relationship_id: &str, part_name: String) -> Result<ChartSheetEntry> {
    let sheets = required_child(root, conformance.sml(), "sheets")?; let mut found = None;
    for sheet in &sheets.children { if sheet.namespace == conformance.sml() && sheet.name == "sheet" && optional(sheet, conformance.rel(), "id") == Some(relationship_id) { if found.is_some() { return Err(invalid("multiple workbook sheets reference the chartsheet relationship")); } found = Some(parse_entry(sheet, conformance, part_name.clone())?); } }
    found.ok_or_else(|| invalid("workbook has no sheet entry for the chartsheet relationship"))
}

fn parse_entry(node: &Node, conformance: ChartSheetConformance, part_name: String) -> Result<ChartSheetEntry> {
    leaf(node, "workbook sheet")?; let state = optional(node, "", "state").map(parse_state).transpose()?.unwrap_or(ChartSheetState::Visible);
    Ok(ChartSheetEntry { name: required(node, "", "name")?.to_owned(), sheet_id: required(node, "", "sheetId")?.parse().map_err(|_| invalid("invalid workbook sheetId"))?, state, workbook_relationship_id: required(node, conformance.rel(), "id")?.to_owned(), part_name })
}

fn validate_new_entry(root: &Node, conformance: ChartSheetConformance, entry: &ChartSheetEntry) -> Result<()> {
    let sheets = required_child(root, conformance.sml(), "sheets")?;
    for sheet in &sheets.children { if sheet.namespace == conformance.sml() && sheet.name == "sheet" { if optional(sheet, "", "name").is_some_and(|v| v.eq_ignore_ascii_case(&entry.name)) { return Err(invalid("workbook sheet name already exists")); } if optional(sheet, "", "sheetId") == Some(entry.sheet_id.to_string().as_str()) { return Err(invalid("workbook sheetId already exists")); } } }
    Ok(())
}

fn validate_entry(entry: &ChartSheetEntry) -> Result<()> { bounded(&entry.name)?; if entry.name.is_empty() || entry.name.chars().count() > 31 || entry.name.chars().any(|c| matches!(c, ':' | '\\' | '/' | '?' | '*' | '[' | ']')) { return Err(invalid("invalid Excel chartsheet name")); } if entry.sheet_id == 0 { return Err(invalid("chartsheet sheetId must be positive")); } validate_id(&entry.workbook_relationship_id)?; let uri = PackURI::new(&entry.part_name).map_err(OoxmlError::InvalidUri)?; if !uri.as_str().starts_with("/xl/chartsheets/") { return Err(invalid("chartsheet part is outside /xl/chartsheets")); } Ok(()) }

fn insert_workbook_entry(xml: &[u8], entry: &ChartSheetEntry, conformance: ChartSheetConformance) -> Result<Vec<u8>> {
    let mut fragment = Vec::new(); fragment.extend_from_slice(b"<x:sheet xmlns:x=\""); escape(&mut fragment, conformance.sml()); fragment.extend_from_slice(b"\" xmlns:r=\""); escape(&mut fragment, conformance.rel()); fragment.extend_from_slice(b"\""); attr(&mut fragment, "name", &entry.name); attr(&mut fragment, "sheetId", &entry.sheet_id.to_string()); if entry.state != ChartSheetState::Visible { attr(&mut fragment, "state", match entry.state { ChartSheetState::Visible => "visible", ChartSheetState::Hidden => "hidden", ChartSheetState::VeryHidden => "veryHidden" }); } attr(&mut fragment, "r:id", &entry.workbook_relationship_id); fragment.extend_from_slice(b"/>");
    let mut reader = NsReader::from_reader(xml); let mut depth = 0usize; let mut sheets_depth = None; let mut position = None;
    loop { let start = usize::try_from(reader.buffer_position()).map_err(|_| invalid("workbook XML offset overflow"))?; let (namespace, event) = reader.read_resolved_event().map_err(xml_error)?; match event { Event::Start(element) => { let core = matches!(namespace, ResolveResult::Bound(Namespace(v)) if v.as_ref() == conformance.sml().as_bytes()); depth += 1; if core && element.local_name().as_ref() == b"sheets" { if sheets_depth.replace(depth).is_some() { return Err(invalid("workbook has multiple sheets collections")); } } }, Event::Empty(element) if element.local_name().as_ref() == b"sheets" => return Err(invalid("cannot insert into empty sheets collection")), Event::End(element) => { if sheets_depth == Some(depth) && element.local_name().as_ref() == b"sheets" { position = Some(start); } depth = depth.checked_sub(1).ok_or_else(|| invalid("unexpected workbook closing element"))?; }, Event::DocType(_) | Event::PI(_) => return Err(invalid("DTDs and processing instructions are rejected")), Event::Eof => break, _ => {} } }
    let position = position.ok_or_else(|| invalid("workbook is missing sheets collection"))?; let size = xml.len().checked_add(fragment.len()).ok_or_else(|| limit("updated workbook bytes"))?; if size > MAX_XML_BYTES { return Err(limit("updated workbook bytes")); } let mut out = Vec::with_capacity(size); out.extend_from_slice(&xml[..position]); out.extend_from_slice(&fragment); out.extend_from_slice(&xml[position..]); Ok(out)
}

fn drawing_chart_ids(xml: &[u8], conformance: ChartSheetConformance) -> Result<Vec<String>> {
    let root = parse_document(xml, MAX_DRAWING_BYTES)?; if root.namespace != conformance.xdr() || root.name != "wsDr" { return Err(invalid("drawing root does not match chartsheet conformance")); }
    let mut ids = Vec::new(); collect_chart_ids(&root, conformance, &mut ids)?; if ids.is_empty() { return Err(invalid("chartsheet drawing contains no chart")); } if ids.len() > MAX_CHARTS { return Err(limit("chart count")); } let mut unique = HashSet::new(); if ids.iter().any(|id| !unique.insert(id.clone())) { return Err(invalid("drawing repeats a chart relationship ID")); } Ok(ids)
}

fn collect_chart_ids(node: &Node, conformance: ChartSheetConformance, ids: &mut Vec<String>) -> Result<()> { if node.namespace == conformance.chart() && node.name == "chart" { ids.push(required(node, conformance.rel(), "id")?.to_owned()); } for child in &node.children { collect_chart_ids(child, conformance, ids)?; } Ok(()) }
fn validate_chart_xml(xml: &[u8], conformance: ChartSheetConformance) -> Result<()> { if xml.len() > MAX_CHART_BYTES { return Err(limit("chart bytes")); } let root = parse_document(xml, MAX_CHART_BYTES)?; if root.namespace == conformance.chart() && root.name == "chartSpace" { Ok(()) } else { Err(invalid("chart root does not match chartsheet conformance")) } }

fn parse_document(xml: &[u8], max_bytes: usize) -> Result<Node> {
    if xml.len() > max_bytes { return Err(limit("input XML bytes")); } let limits = MceLimits { max_input_bytes: max_bytes, max_output_bytes: max_bytes, max_depth: MAX_DEPTH, max_namespace_bindings: 4096, max_directive_tokens: 4096, max_choices_per_alternate: 1024 }; let processed = process_markup_compatibility(xml, &MceCapabilities::ooxml_baseline(), &limits)?.xml;
    let mut reader = NsReader::from_reader(processed.as_ref()); reader.config_mut().trim_text(false); let mut buffer = Vec::new(); let mut stack: Vec<Node> = Vec::new(); let mut root = None; let mut nodes = 0usize; let mut strings = 0usize;
    loop { let event = reader.read_event_into(&mut buffer).map_err(xml_error)?; match event { Event::Start(ref element) | Event::Empty(ref element) => { nodes += 1; if nodes > MAX_NODES || stack.len() >= MAX_DEPTH { return Err(limit("XML structure")); } let empty = matches!(&event, Event::Empty(_)); let node = make_node(&reader, element, reader.decoder(), &mut strings)?; if empty { attach(node, &mut stack, &mut root)?; } else { stack.push(node); } }, Event::End(_) => { let node = stack.pop().ok_or_else(|| invalid("unexpected XML closing element"))?; attach(node, &mut stack, &mut root)?; }, Event::Text(text) => { let decoded = text.decode().map_err(xml_error)?; let decoded = quick_xml::escape::unescape(&decoded).map_err(xml_error)?; add_strings(&mut strings, decoded.len())?; if let Some(node) = stack.last_mut() { node.text.push_str(&decoded); } else if !decoded.trim().is_empty() { return Err(invalid("text outside XML root")); } }, Event::GeneralRef(reference) => { let name = reference.decode().map_err(xml_error)?; let value = reference.resolve_char_ref().map_err(xml_error)?.map(|v| v.to_string()).or_else(|| match name.as_ref() { "amp" => Some("&".into()), "lt" => Some("<".into()), "gt" => Some(">".into()), "apos" => Some("'".into()), "quot" => Some("\"".into()), _ => None }).ok_or_else(|| invalid("custom XML entity is rejected"))?; if let Some(node) = stack.last_mut() { node.text.push_str(&value); } }, Event::CData(_) => return Err(invalid("CDATA is rejected")), Event::DocType(_) | Event::PI(_) => return Err(invalid("DTDs and processing instructions are rejected")), Event::Decl(_) | Event::Comment(_) => {}, Event::Eof => break }; buffer.clear(); }
    if !stack.is_empty() { return Err(invalid("unterminated XML")); } root.ok_or_else(|| invalid("missing XML root"))
}

fn make_node(reader: &NsReader<&[u8]>, element: &BytesStart<'_>, decoder: quick_xml::encoding::Decoder, strings: &mut usize) -> Result<Node> { let namespace = resolved(reader.resolver().resolve_element(element.name()).0)?; let name = std::str::from_utf8(element.local_name().as_ref()).map_err(xml_error)?.to_owned(); add_strings(strings, namespace.len() + name.len())?; let mut attributes = Vec::new(); for item in element.attributes().with_checks(true) { let item = item.map_err(xml_error)?; let qname = item.key.as_ref(); if qname == b"xmlns" || qname.starts_with(b"xmlns:") { continue; } let (namespace, local) = reader.resolver().resolve_attribute(item.key); let namespace = resolved(namespace)?; let name = std::str::from_utf8(local.as_ref()).map_err(xml_error)?.to_owned(); let value = item.decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder).map_err(xml_error)?.into_owned(); add_strings(strings, namespace.len() + name.len() + value.len())?; if attributes.iter().any(|a: &Attribute| a.namespace == namespace && a.name == name) { return Err(invalid("duplicate expanded XML attribute")); } attributes.push(Attribute { namespace, name, value }); } Ok(Node { namespace, name, attributes, children: Vec::new(), text: String::new() }) }

fn attach(node: Node, stack: &mut [Node], root: &mut Option<Node>) -> Result<()> { if let Some(parent) = stack.last_mut() { parent.children.push(node); } else if root.replace(node).is_some() { return Err(invalid("multiple XML roots")); } Ok(()) }
fn root_conformance(root: &Node, name: &str) -> Result<ChartSheetConformance> { if root.name != name { return Err(invalid(format!("expected {name} root"))); } match root.namespace.as_str() { SML => Ok(ChartSheetConformance::Transitional), STRICT_SML => Ok(ChartSheetConformance::Strict), _ => Err(invalid("unsupported SpreadsheetML namespace")) } }
fn one_child<'a>(node: &'a Node, namespace: &str, name: &str) -> Result<Option<&'a Node>> { let mut values = node.children.iter().filter(|c| c.namespace == namespace && c.name == name); let value = values.next(); if values.next().is_some() { Err(invalid(format!("{} has multiple {name} children", node.name))) } else { Ok(value) } }
fn one_child_any_core<'a>(node: &'a Node, name: &str) -> Result<Option<&'a Node>> { let mut values = node.children.iter().filter(|c| is_core(&c.namespace) && c.name == name); let value = values.next(); if values.next().is_some() { Err(invalid(format!("{} has multiple {name} children", node.name))) } else { Ok(value) } }
fn required_child<'a>(node: &'a Node, namespace: &str, name: &str) -> Result<&'a Node> { one_child(node, namespace, name)?.ok_or_else(|| invalid(format!("{} is missing {name}", node.name))) }
fn optional<'a>(node: &'a Node, namespace: &str, name: &str) -> Option<&'a str> { node.attributes.iter().find(|a| a.namespace == namespace && a.name == name).map(|a| a.value.as_str()) }
fn required<'a>(node: &'a Node, namespace: &str, name: &str) -> Result<&'a str> { optional(node, namespace, name).filter(|v| !v.is_empty()).ok_or_else(|| invalid(format!("{} is missing attribute '{name}'", node.name))) }
fn no_attributes(node: &Node, allowed: &[(&str, &str)]) -> Result<()> { if let Some(attribute) = node.attributes.iter().find(|a| !allowed.contains(&(a.namespace.as_str(), a.name.as_str()))) { Err(invalid(format!("unexpected attribute '{}' on {}", attribute.name, node.name))) } else { Ok(()) } }
fn whitespace(node: &Node) -> Result<()> { if node.text.trim().is_empty() { Ok(()) } else { Err(invalid(format!("unexpected text in {}", node.name))) } }
fn leaf(node: &Node, label: &str) -> Result<()> { if node.children.is_empty() { Ok(()) } else { Err(invalid(format!("{label} must not contain child elements"))) } }
fn bool_optional(node: &Node, name: &str) -> Result<Option<bool>> { optional(node, "", name).map(|v| parse_bool(v, name)).transpose() }
fn u32_optional(node: &Node, name: &str) -> Result<Option<u32>> { optional(node, "", name).map(|v| v.parse().map_err(|_| invalid(format!("invalid {name}")))).transpose() }
fn parse_bool(value: &str, name: &str) -> Result<bool> { match value { "1" | "true" => Ok(true), "0" | "false" => Ok(false), _ => Err(invalid(format!("invalid boolean '{value}' for {name}"))) } }
fn parse_orientation(value: &str) -> Result<PageOrientation> { match value { "default" => Ok(PageOrientation::Default), "portrait" => Ok(PageOrientation::Portrait), "landscape" => Ok(PageOrientation::Landscape), _ => Err(invalid("invalid chartsheet page orientation")) } }
fn parse_state(value: &str) -> Result<ChartSheetState> { match value { "visible" => Ok(ChartSheetState::Visible), "hidden" => Ok(ChartSheetState::Hidden), "veryHidden" => Ok(ChartSheetState::VeryHidden), _ => Err(invalid("invalid workbook sheet state")) } }
fn validate_guid(value: &str) -> Result<()> { let bytes = value.as_bytes(); if bytes.len() != 38 || bytes[0] != b'{' || bytes[37] != b'}' || ![9, 14, 19, 24].iter().all(|index| bytes[*index] == b'-') || bytes[1..37].iter().enumerate().any(|(index, byte)| !matches!(index + 1, 9 | 14 | 19 | 24) && !byte.is_ascii_hexdigit()) { return Err(invalid(format!("invalid custom chartsheet view GUID '{value}'"))); } Ok(()) }
fn is_core(value: &str) -> bool { matches!(value, SML | STRICT_SML) }
fn validate_id(value: &str) -> Result<()> { let mut bytes = value.bytes(); let Some(first) = bytes.next() else { return Err(invalid("relationship ID cannot be empty")); }; if !(first.is_ascii_alphabetic() || first == b'_') || !bytes.all(|b| b.is_ascii_alphanumeric() || matches!(b, b'_' | b'-' | b'.')) { Err(invalid(format!("invalid relationship ID '{value}'"))) } else { Ok(()) } }
fn bounded(value: &str) -> Result<()> { if value.len() <= MAX_STRING_BYTES { Ok(()) } else { Err(limit("string bytes")) } }
fn add_strings(total: &mut usize, size: usize) -> Result<()> { *total = total.checked_add(size).ok_or_else(|| limit("XML string bytes"))?; if *total > MAX_STRING_BYTES { Err(limit("XML string bytes")) } else { Ok(()) } }
fn resolved(value: ResolveResult<'_>) -> Result<String> { match value { ResolveResult::Bound(Namespace(value)) => Ok(std::str::from_utf8(value.as_ref()).map_err(xml_error)?.to_owned()), ResolveResult::Unbound => Ok(String::new()), ResolveResult::Unknown(prefix) => Err(invalid(format!("unbound XML prefix '{}'", String::from_utf8_lossy(prefix.as_ref())))) } }
fn internal_relationship<'a>(part: &'a dyn Part, id: &str, kind: &str) -> Result<&'a litchi_opc::Relationship> { let relationship = part.rels().get(id).ok_or_else(|| invalid(format!("missing relationship '{id}'")))?; if relationship.reltype() != kind { return Err(invalid(format!("relationship '{id}' has unexpected type"))); } if relationship.is_external() { return Err(invalid(format!("external relationship '{id}' is not loaded"))); } Ok(relationship) }
fn require_workbook(part: &dyn Part) -> Result<()> { if matches!(part.content_type(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml" | "application/vnd.ms-excel.sheet.macroEnabled.main+xml" | "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml" | "application/vnd.ms-excel.template.macroEnabled.main+xml") { Ok(()) } else { Err(invalid("source part is not a workbook")) } }
fn require_content_type(part: &dyn Part, expected: &str, label: &str) -> Result<()> { if part.content_type() == expected { Ok(()) } else { Err(invalid(format!("{label} part has content type '{}'", part.content_type()))) } }
fn new_uri(package: &OpcPackage, value: &str, prefix: &str) -> Result<PackURI> { let uri = PackURI::new(value).map_err(OoxmlError::InvalidUri)?; if !uri.as_str().starts_with(prefix) { return Err(invalid(format!("part '{uri}' is outside {prefix}"))); } if package.iter_parts().any(|part| part.partname() == &uri) { return Err(invalid(format!("part '{uri}' already exists"))); } Ok(uri) }
fn add_resource(total: &mut usize, size: usize, individual: usize, name: &str) -> Result<()> { if size > individual { return Err(limit(name)); } *total = total.checked_add(size).ok_or_else(|| limit("total resource bytes"))?; if *total > MAX_TOTAL_RESOURCE_BYTES { Err(limit("total resource bytes")) } else { Ok(()) } }
fn bool_attr_opt(out: &mut Vec<u8>, name: &str, value: Option<bool>) { if let Some(value) = value { attr(out, name, if value { "1" } else { "0" }); } }
fn u32_attr_opt(out: &mut Vec<u8>, name: &str, value: Option<u32>) { if let Some(value) = value { attr(out, name, &value.to_string()); } }
fn attr_opt(out: &mut Vec<u8>, name: &str, value: Option<&str>) { if let Some(value) = value { attr(out, name, value); } }
fn attr(out: &mut Vec<u8>, name: &str, value: &str) { out.push(b' '); out.extend_from_slice(name.as_bytes()); out.extend_from_slice(b"=\""); escape(out, value); out.push(b'\"'); }
fn escape(out: &mut Vec<u8>, value: &str) { for c in value.chars() { match c { '&' => out.extend_from_slice(b"&amp;"), '<' => out.extend_from_slice(b"&lt;"), '"' => out.extend_from_slice(b"&quot;"), '\t' => out.extend_from_slice(b"&#x9;"), '\n' => out.extend_from_slice(b"&#xA;"), '\r' => out.extend_from_slice(b"&#xD;"), _ => { let mut bytes = [0; 4]; out.extend_from_slice(c.encode_utf8(&mut bytes).as_bytes()); } } } }
fn escape_text(out: &mut Vec<u8>, value: &str) { for c in value.chars() { match c { '&' => out.extend_from_slice(b"&amp;"), '<' => out.extend_from_slice(b"&lt;"), '>' => out.extend_from_slice(b"&gt;"), _ => { let mut bytes = [0; 4]; out.extend_from_slice(c.encode_utf8(&mut bytes).as_bytes()); } } } }
fn xml_error(error: impl std::fmt::Display) -> OoxmlError { OoxmlError::Xml(error.to_string()) }
fn invalid(message: impl Into<String>) -> OoxmlError { OoxmlError::InvalidFormat(message.into()) }
fn limit(name: &str) -> OoxmlError { invalid(format!("chartsheet {name} limit exceeded")) }

#[cfg(test)]
mod tests {
    use super::*;
    const POI_ONE: &[u8] = include_bytes!("../../../../3rdparty/poi/test-data/spreadsheet/WithChartSheet.xlsx");
    const POI_TWO: &[u8] = include_bytes!("../../../../3rdparty/poi/test-data/spreadsheet/chart_sheet.xlsx");
    fn sheet() -> ChartSheet { ChartSheet { properties: Some(ChartSheetProperties { published: Some(true), code_name: Some("ChartCode".into()), tab_color: Some(ChartSheetColor { automatic: None, indexed: None, rgb: Some("FF336699".into()), theme: None, tint: Some(0.25) }) }), views: vec![ChartSheetView { tab_selected: Some(true), zoom_scale: Some(125), workbook_view_id: 0, zoom_to_fit: Some(false) }], protection: Some(ChartSheetProtection { password_hash: Some("ABCD".into()), content: Some(true), objects: Some(false) }), custom_views: Some(vec![ChartSheetCustomView { guid: "{00112233-4455-6677-8899-AABBCCDDEEFF}".into(), scale: Some(175), state: Some(ChartSheetState::Hidden), zoom_to_fit: Some(true) }, ChartSheetCustomView { guid: "{10213243-5465-7687-98A9-BACBDCEDFE0F}".into(), scale: None, state: None, zoom_to_fit: Some(false) }]), margins: Some(ChartSheetMargins { left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 }), page_setup: Some(ChartSheetPageSetup { paper_size: Some(1), first_page_number: Some(1), orientation: Some(PageOrientation::Landscape), use_printer_defaults: Some(true), black_and_white: Some(false), draft: Some(false), use_first_page_number: Some(true), horizontal_dpi: Some(600), vertical_dpi: Some(600), copies: Some(1) }), header_footer: Some(ChartSheetHeaderFooter { align_with_margins: Some(false), odd_header: Some("&CChart & Report".into()), ..Default::default() }), drawing_relationship_id: "rIdDrawing".into(), background_picture_relationship_id:Some("rIdBackground".into()) } }
    fn drawing(conformance: ChartSheetConformance) -> Vec<u8> { format!("<xdr:wsDr xmlns:xdr=\"{}\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><xdr:absoluteAnchor><a:graphic><a:graphicData><c:chart xmlns:c=\"{}\" xmlns:r=\"{}\" r:id=\"rIdChart\"/></a:graphicData></a:graphic></xdr:absoluteAnchor></xdr:wsDr>", conformance.xdr(), conformance.chart(), conformance.rel()).into_bytes() }
    fn chart(conformance: ChartSheetConformance) -> Vec<u8> { format!("<c:chartSpace xmlns:c=\"{}\"><c:chart/></c:chartSpace>", conformance.chart()).into_bytes() }
    fn value(conformance: ChartSheetConformance) -> ChartSheetPackage { ChartSheetPackage { entry: ChartSheetEntry { name: "Chart 1".into(), sheet_id: 2, state: ChartSheetState::Visible, workbook_relationship_id: "rIdChartSheet".into(), part_name: "/xl/chartsheets/sheet1.xml".into() }, chartsheet: sheet(), drawing: ChartSheetDrawingResource { part_name: "/xl/drawings/drawing1.xml".into(), content_type: DRAWING_CT.into(), data: drawing(conformance), charts: vec![ChartSheetChartResource { relationship_id: "rIdChart".into(), part_name: "/xl/charts/chart1.xml".into(), content_type: CHART_CT.into(), data: chart(conformance) }] }, background_picture:Some(ChartSheetBackgroundPicture{relationship_id:"rIdBackground".into(),part_name:"/xl/media/background1.png".into(),content_type:ChartSheetImageContentType::Png,data:vec![0,255,1,254]}) } }
    fn base_package(conformance: ChartSheetConformance) -> (OpcPackage, PackURI) { let mut package = OpcPackage::new(); let uri = PackURI::new("/xl/workbook.xml").unwrap(); let xml = format!("<x:workbook xmlns:x=\"{}\" xmlns:r=\"{}\"><x:sheets><x:sheet name=\"Data\" sheetId=\"1\" r:id=\"rIdData\"/></x:sheets></x:workbook>", conformance.sml(), conformance.rel()); package.add_part(Box::new(BlobPart::new(uri.clone(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml".into(), xml.into_bytes()))); (package, uri) }

    #[test] fn strict_typed_xml_round_trip() { let expected = sheet(); let xml = write_chartsheet(&expected, ChartSheetConformance::Strict).unwrap(); let (kind, parsed) = parse_chartsheet(&xml).unwrap(); assert_eq!(kind, ChartSheetConformance::Strict); assert_eq!(parsed, expected); }
    #[test] fn transitional_custom_view_reference_round_trip() { let xml = format!("<x:chartsheet xmlns:x=\"{SML}\" xmlns:r=\"{REL}\"><x:sheetViews><x:sheetView workbookViewId=\"0\"/></x:sheetViews><x:customSheetViews><x:customSheetView guid=\"{{00112233-4455-6677-8899-AABBCCDDEEFF}}\" scale=\"10\" state=\"veryHidden\" zoomToFit=\"1\"/><x:customSheetView guid=\"{{10213243-5465-7687-98A9-BACBDCEDFE0F}}\"/></x:customSheetViews><x:drawing r:id=\"rId1\"/></x:chartsheet>"); let (_, parsed) = parse_chartsheet(xml.as_bytes()).unwrap(); let views = parsed.custom_views.as_ref().unwrap(); assert_eq!(views.len(), 2); assert_eq!(views[0].state, Some(ChartSheetState::VeryHidden)); assert_eq!(views[0].scale, Some(10)); let written = write_chartsheet(&parsed, ChartSheetConformance::Transitional).unwrap(); assert_eq!(parse_chartsheet(&written).unwrap().1, parsed); }
    #[test] fn mce_fallback_selects_chartsheet_views() { let xml = format!("<x:chartsheet xmlns:x=\"{SML}\" xmlns:r=\"{REL}\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:u=\"urn:unsupported\"><mc:AlternateContent><mc:Choice Requires=\"u\"><u:views/></mc:Choice><mc:Fallback><x:sheetViews><x:sheetView workbookViewId=\"0\"/></x:sheetViews></mc:Fallback></mc:AlternateContent><x:drawing r:id=\"rId1\"/></x:chartsheet>"); assert_eq!(parse_chartsheet(xml.as_bytes()).unwrap().1.views.len(), 1); }
    #[test] fn mce_fallback_selects_custom_chartsheet_views() { let xml = format!("<x:chartsheet xmlns:x=\"{SML}\" xmlns:r=\"{REL}\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:u=\"urn:unsupported\"><x:sheetViews><x:sheetView workbookViewId=\"0\"/></x:sheetViews><mc:AlternateContent><mc:Choice Requires=\"u\"><u:customViews/></mc:Choice><mc:Fallback><x:customSheetViews><x:customSheetView guid=\"{{00112233-4455-6677-8899-AABBCCDDEEFF}}\" scale=\"200\"/></x:customSheetViews></mc:Fallback></mc:AlternateContent><x:drawing r:id=\"rId1\"/></x:chartsheet>"); let parsed = parse_chartsheet(xml.as_bytes()).unwrap().1; assert_eq!(parsed.custom_views.unwrap()[0].scale, Some(200)); }
    #[test] fn loads_both_poi_chartsheet_graphs() { for (bytes, name, zoom) in [(POI_ONE, "Chart2", 131), (POI_TWO, "Chart1", 84)] { let package = OpcPackage::from_bytes(bytes).unwrap(); let workbook = PackURI::new("/xl/workbook.xml").unwrap(); let workbook_part = package.get_part(&workbook).unwrap(); let id = workbook_part.rels().iter().find(|rel| rel.reltype() == CHARTSHEET_REL).unwrap().r_id().to_owned(); let loaded = load_chartsheet(&package, &workbook, &id).unwrap(); assert_eq!(loaded.entry.name, name); assert_eq!(loaded.chartsheet.views[0].zoom_scale, Some(zoom)); assert_eq!(loaded.drawing.charts.len(), 1); assert!(loaded.drawing.charts[0].data.starts_with(b"<?xml")); } }
    #[test] fn strict_package_writer_round_trips_complete_leaf_graph() { let conformance = ChartSheetConformance::Strict; let (mut package, workbook) = base_package(conformance); let expected = value(conformance); store_chartsheet(&mut package, &workbook, &expected, conformance).unwrap(); assert_eq!(load_chartsheet(&package, &workbook, "rIdChartSheet").unwrap(), expected); }
    #[test]fn picture_mce_schema_and_inert_round_trip(){let xml=format!("<x:chartsheet xmlns:x=\"{SML}\" xmlns:r=\"{REL}\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:u=\"urn:unsupported\"><x:sheetViews><x:sheetView workbookViewId=\"0\"/></x:sheetViews><x:drawing r:id=\"rIdDrawing\"/><mc:AlternateContent><mc:Choice Requires=\"u\"><u:picture/></mc:Choice><mc:Fallback><x:picture r:id=\"rIdBackground\"/></mc:Fallback></mc:AlternateContent></x:chartsheet>");let(_,parsed)=parse_chartsheet(xml.as_bytes()).unwrap();assert_eq!(parsed.background_picture_relationship_id.as_deref(),Some("rIdBackground"));let written=write_chartsheet(&parsed,ChartSheetConformance::Transitional).unwrap();assert!(String::from_utf8(written.clone()).unwrap().contains("<x:drawing r:id=\"rIdDrawing\"/><x:picture r:id=\"rIdBackground\"/>"));assert_eq!(parse_chartsheet(&written).unwrap().1,parsed);}
    #[test]fn transitional_picture_package_round_trip_preserves_opaque_bytes(){let conformance=ChartSheetConformance::Transitional;let(mut package,workbook)=base_package(conformance);let expected=value(conformance);store_chartsheet(&mut package,&workbook,&expected,conformance).unwrap();let loaded=load_chartsheet(&package,&workbook,"rIdChartSheet").unwrap();assert_eq!(loaded.background_picture.as_ref().unwrap().data,vec![0,255,1,254]);assert_eq!(loaded,expected);}
    #[test]fn rejects_picture_cardinality_order_metadata_and_caps(){for xml in [format!("<chartsheet xmlns=\"{SML}\" xmlns:r=\"{REL}\"><sheetViews><sheetView workbookViewId=\"0\"/></sheetViews><picture r:id=\"rIdP\"/><drawing r:id=\"rIdD\"/></chartsheet>"),format!("<chartsheet xmlns=\"{SML}\" xmlns:r=\"{REL}\"><sheetViews><sheetView workbookViewId=\"0\"/></sheetViews><drawing r:id=\"rIdD\"/><picture r:id=\"rIdP\"/><picture r:id=\"rIdQ\"/></chartsheet>"),format!("<chartsheet xmlns=\"{SML}\" xmlns:r=\"{REL}\"><sheetViews><sheetView workbookViewId=\"0\"/></sheetViews><drawing r:id=\"rIdD\"/><picture/></chartsheet>")]{assert!(parse_chartsheet(xml.as_bytes()).is_err(),"{xml}");}let conformance=ChartSheetConformance::Transitional;let(mut package,workbook)=base_package(conformance);let mut bad=value(conformance);bad.background_picture.as_mut().unwrap().relationship_id="different".into();assert!(store_chartsheet(&mut package,&workbook,&bad,conformance).is_err());let(mut package,workbook)=base_package(conformance);let mut bad=value(conformance);bad.background_picture.as_mut().unwrap().part_name="/xl/charts/chart1.xml".into();assert!(store_chartsheet(&mut package,&workbook,&bad,conformance).is_err());let(mut package,workbook)=base_package(conformance);let mut bad=value(conformance);bad.background_picture.as_mut().unwrap().data=vec![0;MAX_BACKGROUND_IMAGE_BYTES+1];assert!(store_chartsheet(&mut package,&workbook,&bad,conformance).is_err());}
    #[test]fn rejects_external_wrong_type_escaped_and_unreferenced_picture_relationships(){for (kind,target,external) in [(IMAGE_REL,"https://example.invalid/background.png",true),(rt::CHART,"../media/background1.png",false),(IMAGE_REL,"../../../evil.png",false)]{let conformance=ChartSheetConformance::Transitional;let(mut package,workbook)=base_package(conformance);store_chartsheet(&mut package,&workbook,&value(conformance),conformance).unwrap();let chartsheet=package.get_part_mut(&PackURI::new("/xl/chartsheets/sheet1.xml").unwrap()).unwrap();chartsheet.rels_mut().remove("rIdBackground");chartsheet.rels_mut().add_relationship(kind.into(),target.into(),"rIdBackground".into(),external);assert!(load_chartsheet(&package,&workbook,"rIdChartSheet").is_err(),"accepted {kind} {target}");}let conformance=ChartSheetConformance::Transitional;let(mut package,workbook)=base_package(conformance);store_chartsheet(&mut package,&workbook,&value(conformance),conformance).unwrap();package.get_part_mut(&PackURI::new("/xl/chartsheets/sheet1.xml").unwrap()).unwrap().rels_mut().add_relationship(IMAGE_REL.into(),"../media/background1.png".into(),"rIdExtra".into(),false);assert!(load_chartsheet(&package,&workbook,"rIdChartSheet").is_err());}
    #[test]fn rejects_existing_background_part_collision_before_mutation(){let conformance=ChartSheetConformance::Transitional;let(mut package,workbook)=base_package(conformance);package.add_part(Box::new(BlobPart::new(PackURI::new("/xl/media/background1.png").unwrap(),"image/png".into(),vec![9])));assert!(store_chartsheet(&mut package,&workbook,&value(conformance),conformance).is_err());assert!(package.get_part(&PackURI::new("/xl/chartsheets/sheet1.xml").unwrap()).is_err());}
    #[test] fn rejects_malformed_caps_and_graphs() { assert!(parse_chartsheet(b"<!DOCTYPE x><chartsheet/>").is_err()); assert!(parse_chartsheet(format!("<chartsheet xmlns=\"{SML}\"><sheetViews><sheetView workbookViewId=\"0\" zoomScale=\"401\"/></sheetViews><drawing xmlns:r=\"{REL}\" r:id=\"rId1\"/></chartsheet>").as_bytes()).is_err()); for custom in ["<customSheetViews/>", "<customSheetViews><customSheetView guid=\"bad\"/></customSheetViews>", "<customSheetViews><customSheetView guid=\"{00112233-4455-6677-8899-AABBCCDDEEFF}\" scale=\"401\"/></customSheetViews>", "<customSheetViews><customSheetView guid=\"{00112233-4455-6677-8899-AABBCCDDEEFF}\"/><customSheetView guid=\"{00112233-4455-6677-8899-aabbccddeeff}\"/></customSheetViews>"] { let xml = format!("<chartsheet xmlns=\"{SML}\" xmlns:r=\"{REL}\"><sheetViews><sheetView workbookViewId=\"0\"/></sheetViews>{custom}<drawing r:id=\"rId1\"/></chartsheet>"); assert!(parse_chartsheet(xml.as_bytes()).is_err(), "{custom}"); } assert!(parse_chartsheet(&vec![b' '; MAX_XML_BYTES + 1]).is_err()); let (mut package, workbook) = base_package(ChartSheetConformance::Transitional); let expected = value(ChartSheetConformance::Transitional); store_chartsheet(&mut package, &workbook, &expected, ChartSheetConformance::Transitional).unwrap(); package.get_part_mut(&PackURI::new("/xl/drawings/drawing1.xml").unwrap()).unwrap().rels_mut().add_relationship(rt::IMAGE.into(), "../media/x.png".into(), "rIdBad".into(), false); assert!(load_chartsheet(&package, &workbook, "rIdChartSheet").is_err()); }
}
