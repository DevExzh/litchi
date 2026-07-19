//! Inert readers for the Office 2013 chart-style companion parts.

use crate::error::{OoxmlError, Result};
use litchi_opc::OpcPackage;
use litchi_opc::part::Part;
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::ResolveResult;
use quick_xml::reader::NsReader;
use std::collections::HashSet;

pub const CHART_STYLE_CONTENT_TYPE: &str = "application/vnd.ms-office.chartstyle+xml";
pub const CHART_COLOR_STYLE_CONTENT_TYPE: &str = "application/vnd.ms-office.chartcolorstyle+xml";
pub const CHART_STYLE_RELATIONSHIP_TYPE: &str =
    "http://schemas.microsoft.com/office/2012/relationships/chartStyle";
pub const CHART_COLOR_STYLE_RELATIONSHIP_TYPE: &str =
    "http://schemas.microsoft.com/office/2012/relationships/chartColorStyle";

const CS: &str = "http://schemas.microsoft.com/office/drawing/2012/chartStyle";
const A: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const A_STRICT: &str = "http://purl.oclc.org/ooxml/drawingml/main";
const MC: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const MAX_XML_BYTES: usize = 16 * 1024 * 1024;
const MAX_NODES: usize = 100_000;
const MAX_DEPTH: usize = 96;
const MAX_ATTRIBUTES: usize = 64;
const MAX_STRING_BYTES: usize = 4 * 1024 * 1024;
const MAX_MODIFIERS: usize = 256;
const MAX_COLORS: usize = 4_096;
const MAX_VARIATIONS: usize = 4_096;
const MAX_TRANSFORMS: usize = 4_096;

/// A validated Office 2013 chart-style part.
pub struct ChartStylePart<'a> {
    part: &'a dyn Part,
}

/// A validated Office 2013 chart-color-style part.
pub struct ChartColorStylePart<'a> {
    part: &'a dyn Part,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChartStyleDocument {
    info: ChartStyleInfo,
    xml: Vec<u8>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChartColorStyleDocument {
    info: ChartColorStyleInfo,
    xml: Vec<u8>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChartStyleInfo {
    pub id: Option<u32>,
    pub entries: Vec<ChartStyleEntry>,
    pub marker_layout: Option<ChartStyleMarkerLayout>,
    pub has_extension_list: bool,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ChartStyleEntryKind {
    AxisTitle,
    CategoryAxis,
    ChartArea,
    DataLabel,
    DataLabelCallout,
    DataPoint,
    DataPoint3D,
    DataPointLine,
    DataPointMarker,
    DataPointWireframe,
    DataTable,
    DownBar,
    DropLine,
    ErrorBar,
    Floor,
    GridlineMajor,
    GridlineMinor,
    HiLoLine,
    LeaderLine,
    Legend,
    PlotArea,
    PlotArea3D,
    SeriesAxis,
    SeriesLine,
    Title,
    Trendline,
    TrendlineLabel,
    UpBar,
    ValueAxis,
    Wall,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChartStyleEntry {
    pub kind: ChartStyleEntryKind,
    pub modifiers: Vec<String>,
    pub line_reference: ChartStyleReference,
    /// The validated XML Schema double lexical value; defaults to `1.0`.
    pub line_width_scale: String,
    pub fill_reference: ChartStyleReference,
    pub effect_reference: ChartStyleReference,
    pub font_reference: ChartStyleFontReference,
    pub shape_properties: Option<ChartStylePayload>,
    pub default_run_properties: Option<ChartStylePayload>,
    pub body_properties: Option<ChartStylePayload>,
    pub has_extension_list: bool,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChartStyleReference {
    pub index: u32,
    pub modifiers: Vec<String>,
    pub color: Option<ChartStyleColor>,
    pub style_color: Option<ChartStyleColorValue>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ChartStyleFontIndex {
    Major,
    Minor,
    None,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChartStyleFontReference {
    pub index: ChartStyleFontIndex,
    pub modifiers: Vec<String>,
    pub color: Option<ChartStyleColor>,
    pub style_color: Option<ChartStyleColorValue>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChartStyleColorValue {
    pub raw: Option<String>,
    pub index: Option<u32>,
    pub automatic: bool,
    pub transforms: Vec<ChartStyleColorTransform>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ChartStyleMarkerSymbol {
    Circle,
    Dash,
    Diamond,
    Dot,
    Plus,
    Square,
    Star,
    Triangle,
    X,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChartStyleMarkerLayout {
    pub symbol: Option<ChartStyleMarkerSymbol>,
    pub size: Option<u8>,
}

/// A bounded summary of an inert DrawingML formatting subtree.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChartStylePayload {
    pub child_elements: usize,
    pub attributes: usize,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ChartStyleColorKind {
    ScRgb,
    Srgb,
    Hsl,
    System,
    Scheme,
    Preset,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChartStyleColor {
    pub kind: ChartStyleColorKind,
    /// Primary color value where the color model has one; component models use `components`.
    pub value: Option<String>,
    pub components: Vec<(String, String)>,
    pub transforms: Vec<ChartStyleColorTransform>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ChartStyleColorTransformKind {
    Tint,
    Shade,
    Complement,
    Inverse,
    Grayscale,
    Alpha,
    AlphaOffset,
    AlphaModulation,
    Hue,
    HueOffset,
    HueModulation,
    Saturation,
    SaturationOffset,
    SaturationModulation,
    Luminance,
    LuminanceOffset,
    LuminanceModulation,
    Red,
    RedOffset,
    RedModulation,
    Green,
    GreenOffset,
    GreenModulation,
    Blue,
    BlueOffset,
    BlueModulation,
    Gamma,
    InverseGamma,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChartStyleColorTransform {
    pub kind: ChartStyleColorTransformKind,
    /// Preserved integer lexical value for transforms that take `val`.
    pub value: Option<String>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ChartColorStyleMethod {
    Cycle,
    WithinLinear,
    AcrossLinear,
    WithinLinearReversed,
    AcrossLinearReversed,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChartColorStyleInfo {
    pub method: String,
    /// Unknown extension methods have the specified effective behavior `Cycle`.
    pub effective_method: ChartColorStyleMethod,
    pub id: Option<u32>,
    pub colors: Vec<ChartStyleColor>,
    pub variations: Vec<ChartStyleVariation>,
    pub has_extension_list: bool,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChartStyleVariation {
    pub transforms: Vec<ChartStyleColorTransform>,
}

impl<'a> ChartStylePart<'a> {
    pub fn from_part(part: &'a dyn Part) -> Result<Self> {
        require_content_type(part, CHART_STYLE_CONTENT_TYPE, "chart style")?;
        Ok(Self { part })
    }

    pub fn parse(&self) -> Result<ChartStyleDocument> {
        parse_chart_style(self.part.blob())
    }

    pub fn part(&self) -> &'a dyn Part {
        self.part
    }
}

impl<'a> ChartColorStylePart<'a> {
    pub fn from_part(part: &'a dyn Part) -> Result<Self> {
        require_content_type(part, CHART_COLOR_STYLE_CONTENT_TYPE, "chart color style")?;
        Ok(Self { part })
    }

    pub fn parse(&self) -> Result<ChartColorStyleDocument> {
        parse_color_style(self.part.blob())
    }

    pub fn part(&self) -> &'a dyn Part {
        self.part
    }
}

impl ChartStyleDocument {
    pub fn info(&self) -> &ChartStyleInfo {
        &self.info
    }

    pub fn to_xml(&self) -> Vec<u8> {
        self.xml.clone()
    }
}

impl ChartColorStyleDocument {
    pub fn info(&self) -> &ChartColorStyleInfo {
        &self.info
    }

    pub fn to_xml(&self) -> Vec<u8> {
        self.xml.clone()
    }
}

pub(crate) fn discover_chart_styles(
    package: &OpcPackage,
    source: &dyn Part,
) -> Result<(Option<ChartStyleDocument>, Option<ChartColorStyleDocument>)> {
    let mut style = None;
    let mut colors = None;
    for relationship in source.rels().iter() {
        let (expected_type, label) = match relationship.reltype() {
            CHART_STYLE_RELATIONSHIP_TYPE => (CHART_STYLE_CONTENT_TYPE, "chart style"),
            CHART_COLOR_STYLE_RELATIONSHIP_TYPE => {
                (CHART_COLOR_STYLE_CONTENT_TYPE, "chart color style")
            },
            _ => continue,
        };
        if relationship.is_external() {
            return invalid(format!("external {label} relationships are not loaded"));
        }
        reject_target(relationship.target_ref(), label)?;
        let target = relationship.target_partname().map_err(OoxmlError::Opc)?;
        let source_parent = parent_name(source.partname().as_str());
        if parent_name(target.as_str()) != source_parent || target.as_str().ends_with('/') {
            return invalid(format!("{label} part is not a sibling of the ChartEx part"));
        }
        let target_part = package
            .get_part(&target)
            .map_err(|_| invalid_error(format!("{label} target is missing")))?;
        require_content_type(target_part, expected_type, label)?;
        if expected_type == CHART_STYLE_CONTENT_TYPE {
            if style.is_some() {
                return invalid("multiple chart style relationships");
            }
            style = Some(ChartStylePart::from_part(target_part)?.parse()?);
        } else {
            if colors.is_some() {
                return invalid("multiple chart color style relationships");
            }
            colors = Some(ChartColorStylePart::from_part(target_part)?.parse()?);
        }
    }
    Ok((style, colors))
}

#[derive(Debug)]
struct Attribute {
    namespace: String,
    name: String,
    value: String,
}

#[derive(Debug)]
struct Node {
    namespace: String,
    name: String,
    attributes: Vec<Attribute>,
    namespace_declarations: Vec<(String, String)>,
    children: Vec<Node>,
    text: String,
}

fn parse_chart_style(xml: &[u8]) -> Result<ChartStyleDocument> {
    let root = parse_tree(xml)?;
    require_root(&root, "chartStyle")?;
    let ignorable = ignorable_namespaces(&root)?;
    let id = optional_u32_attribute(&root, "id", &ignorable)?;
    reject_root_attributes(&root, &["id"], &ignorable)?;
    let children = semantic_children(&root, &ignorable)?;
    let schema: [(&str, bool, Option<ChartStyleEntryKind>); 31] = [
        ("axisTitle", true, Some(ChartStyleEntryKind::AxisTitle)),
        (
            "categoryAxis",
            true,
            Some(ChartStyleEntryKind::CategoryAxis),
        ),
        ("chartArea", true, Some(ChartStyleEntryKind::ChartArea)),
        ("dataLabel", true, Some(ChartStyleEntryKind::DataLabel)),
        (
            "dataLabelCallout",
            false,
            Some(ChartStyleEntryKind::DataLabelCallout),
        ),
        ("dataPoint", true, Some(ChartStyleEntryKind::DataPoint)),
        ("dataPoint3D", true, Some(ChartStyleEntryKind::DataPoint3D)),
        (
            "dataPointLine",
            true,
            Some(ChartStyleEntryKind::DataPointLine),
        ),
        (
            "dataPointMarker",
            true,
            Some(ChartStyleEntryKind::DataPointMarker),
        ),
        ("dataPointMarkerLayout", false, None),
        (
            "dataPointWireframe",
            true,
            Some(ChartStyleEntryKind::DataPointWireframe),
        ),
        ("dataTable", true, Some(ChartStyleEntryKind::DataTable)),
        ("downBar", true, Some(ChartStyleEntryKind::DownBar)),
        ("dropLine", true, Some(ChartStyleEntryKind::DropLine)),
        ("errorBar", true, Some(ChartStyleEntryKind::ErrorBar)),
        ("floor", true, Some(ChartStyleEntryKind::Floor)),
        (
            "gridlineMajor",
            true,
            Some(ChartStyleEntryKind::GridlineMajor),
        ),
        (
            "gridlineMinor",
            true,
            Some(ChartStyleEntryKind::GridlineMinor),
        ),
        ("hiLoLine", true, Some(ChartStyleEntryKind::HiLoLine)),
        ("leaderLine", true, Some(ChartStyleEntryKind::LeaderLine)),
        ("legend", true, Some(ChartStyleEntryKind::Legend)),
        ("plotArea", true, Some(ChartStyleEntryKind::PlotArea)),
        ("plotArea3D", true, Some(ChartStyleEntryKind::PlotArea3D)),
        ("seriesAxis", true, Some(ChartStyleEntryKind::SeriesAxis)),
        ("seriesLine", true, Some(ChartStyleEntryKind::SeriesLine)),
        ("title", true, Some(ChartStyleEntryKind::Title)),
        ("trendline", true, Some(ChartStyleEntryKind::Trendline)),
        (
            "trendlineLabel",
            true,
            Some(ChartStyleEntryKind::TrendlineLabel),
        ),
        ("upBar", true, Some(ChartStyleEntryKind::UpBar)),
        ("valueAxis", true, Some(ChartStyleEntryKind::ValueAxis)),
        ("wall", true, Some(ChartStyleEntryKind::Wall)),
    ];
    let mut cursor = 0usize;
    let mut entries = Vec::with_capacity(30);
    let mut marker_layout = None;
    for (name, required, kind) in schema {
        let matches = children
            .get(cursor)
            .is_some_and(|child| child.namespace == CS && child.name == name);
        if !matches {
            if required {
                return invalid(format!("missing or out-of-order cs:{name}"));
            }
            continue;
        }
        let child = children[cursor];
        cursor += 1;
        if let Some(kind) = kind {
            entries.push(parse_style_entry(child, kind, &ignorable)?);
        } else {
            marker_layout = Some(parse_marker_layout(child, &ignorable)?);
        }
    }
    let has_extension_list = children
        .get(cursor)
        .is_some_and(|child| child.namespace == CS && child.name == "extLst");
    if has_extension_list {
        parse_payload(children[cursor], &ignorable)?;
        cursor += 1;
    }
    if cursor != children.len() {
        return invalid("unsupported, duplicated, or out-of-order chart style child");
    }
    Ok(ChartStyleDocument {
        info: ChartStyleInfo {
            id,
            entries,
            marker_layout,
            has_extension_list,
        },
        xml: xml.to_vec(),
    })
}

fn parse_color_style(xml: &[u8]) -> Result<ChartColorStyleDocument> {
    let root = parse_tree(xml)?;
    require_root(&root, "colorStyle")?;
    let ignorable = ignorable_namespaces(&root)?;
    let method = required_attribute(&root, "meth", &ignorable)?.to_owned();
    check_string(&method, 256, "chart color style method")?;
    let effective_method = match method.as_str() {
        "cycle" => ChartColorStyleMethod::Cycle,
        "withinLinear" => ChartColorStyleMethod::WithinLinear,
        "acrossLinear" => ChartColorStyleMethod::AcrossLinear,
        "withinLinearReversed" => ChartColorStyleMethod::WithinLinearReversed,
        "acrossLinearReversed" => ChartColorStyleMethod::AcrossLinearReversed,
        _ => ChartColorStyleMethod::Cycle,
    };
    let id = optional_u32_attribute(&root, "id", &ignorable)?;
    reject_root_attributes(&root, &["meth", "id"], &ignorable)?;
    let children = semantic_children(&root, &ignorable)?;
    let mut cursor = 0usize;
    let mut colors = Vec::new();
    while let Some(child) = children.get(cursor) {
        if !is_a(&child.namespace) || color_kind(&child.name).is_none() {
            break;
        }
        if colors.len() >= MAX_COLORS {
            return limit("chart color style colors");
        }
        colors.push(parse_color(child, &ignorable)?);
        cursor += 1;
    }
    if colors.is_empty() {
        return invalid("chart color style requires at least one DrawingML color");
    }
    let mut variations = Vec::new();
    while children
        .get(cursor)
        .is_some_and(|child| child.namespace == CS && child.name == "variation")
    {
        if variations.len() >= MAX_VARIATIONS {
            return limit("chart color style variations");
        }
        let child = children[cursor];
        reject_attributes(child, &[], &ignorable, "variation")?;
        if !child.text.trim().is_empty() {
            return invalid("chart color style variation contains text");
        }
        variations.push(ChartStyleVariation {
            transforms: parse_transforms(child, &ignorable)?,
        });
        cursor += 1;
    }
    let has_extension_list = children
        .get(cursor)
        .is_some_and(|child| child.namespace == CS && child.name == "extLst");
    if has_extension_list {
        parse_payload(children[cursor], &ignorable)?;
        cursor += 1;
    }
    if cursor != children.len() {
        return invalid("unsupported, duplicated, or out-of-order chart color style child");
    }
    Ok(ChartColorStyleDocument {
        info: ChartColorStyleInfo {
            method,
            effective_method,
            id,
            colors,
            variations,
            has_extension_list,
        },
        xml: xml.to_vec(),
    })
}

fn parse_style_entry(
    node: &Node,
    kind: ChartStyleEntryKind,
    ignorable: &HashSet<String>,
) -> Result<ChartStyleEntry> {
    let modifiers = modifiers(node, ignorable)?;
    reject_attributes(node, &["mods"], ignorable, &node.name)?;
    if !node.text.trim().is_empty() {
        return invalid("chart style entry contains text");
    }
    let children = semantic_children(node, ignorable)?;
    let mut cursor = 0usize;
    let line_reference = parse_required_reference(children.get(cursor), "lnRef", ignorable)?;
    cursor += 1;
    let line_width_scale = if children
        .get(cursor)
        .is_some_and(|child| child.namespace == CS && child.name == "lineWidthScale")
    {
        let child = children[cursor];
        reject_attributes(child, &[], ignorable, "lineWidthScale")?;
        if !semantic_children(child, ignorable)?.is_empty() {
            return invalid("lineWidthScale contains child elements");
        }
        validate_double(child.text.trim())?;
        cursor += 1;
        child.text.trim().to_owned()
    } else {
        "1.0".to_owned()
    };
    let fill_reference = parse_required_reference(children.get(cursor), "fillRef", ignorable)?;
    cursor += 1;
    let effect_reference = parse_required_reference(children.get(cursor), "effectRef", ignorable)?;
    cursor += 1;
    let font_reference = parse_required_font_reference(children.get(cursor), ignorable)?;
    cursor += 1;
    let shape_properties = take_payload(&children, &mut cursor, "spPr", ignorable)?;
    let default_run_properties = take_payload(&children, &mut cursor, "defRPr", ignorable)?;
    let body_properties = take_payload(&children, &mut cursor, "bodyPr", ignorable)?;
    let has_extension_list = children
        .get(cursor)
        .is_some_and(|child| child.namespace == CS && child.name == "extLst");
    if has_extension_list {
        parse_payload(children[cursor], ignorable)?;
        cursor += 1;
    }
    if cursor != children.len() {
        return invalid(format!("invalid child order in cs:{}", node.name));
    }
    Ok(ChartStyleEntry {
        kind,
        modifiers,
        line_reference,
        line_width_scale,
        fill_reference,
        effect_reference,
        font_reference,
        shape_properties,
        default_run_properties,
        body_properties,
        has_extension_list,
    })
}

fn parse_required_reference(
    node: Option<&&Node>,
    expected: &str,
    ignorable: &HashSet<String>,
) -> Result<ChartStyleReference> {
    let node = node
        .copied()
        .filter(|node| node.namespace == CS && node.name == expected)
        .ok_or_else(|| invalid_error(format!("missing cs:{expected}")))?;
    let index = required_attribute(node, "idx", ignorable)?
        .parse::<u32>()
        .map_err(|_| invalid_error(format!("invalid {expected} index")))?;
    let modifiers = modifiers(node, ignorable)?;
    reject_attributes(node, &["idx", "mods"], ignorable, expected)?;
    let (color, style_color) = parse_reference_colors(node, ignorable)?;
    Ok(ChartStyleReference {
        index,
        modifiers,
        color,
        style_color,
    })
}

fn parse_required_font_reference(
    node: Option<&&Node>,
    ignorable: &HashSet<String>,
) -> Result<ChartStyleFontReference> {
    let node = node
        .copied()
        .filter(|node| node.namespace == CS && node.name == "fontRef")
        .ok_or_else(|| invalid_error("missing cs:fontRef"))?;
    let index = match required_attribute(node, "idx", ignorable)? {
        "major" => ChartStyleFontIndex::Major,
        "minor" => ChartStyleFontIndex::Minor,
        "none" => ChartStyleFontIndex::None,
        _ => return invalid("invalid fontRef index"),
    };
    let modifiers = modifiers(node, ignorable)?;
    reject_attributes(node, &["idx", "mods"], ignorable, "fontRef")?;
    let (color, style_color) = parse_reference_colors(node, ignorable)?;
    Ok(ChartStyleFontReference {
        index,
        modifiers,
        color,
        style_color,
    })
}

fn parse_reference_colors(
    node: &Node,
    ignorable: &HashSet<String>,
) -> Result<(Option<ChartStyleColor>, Option<ChartStyleColorValue>)> {
    if !node.text.trim().is_empty() {
        return invalid("chart style reference contains text");
    }
    let children = semantic_children(node, ignorable)?;
    let mut cursor = 0usize;
    let color = children
        .get(cursor)
        .filter(|child| is_a(&child.namespace) && color_kind(&child.name).is_some());
    let color = if let Some(color) = color {
        cursor += 1;
        Some(parse_color(color, ignorable)?)
    } else {
        None
    };
    let style_color = children
        .get(cursor)
        .filter(|child| child.namespace == CS && child.name == "styleClr");
    let style_color = if let Some(style_color) = style_color {
        cursor += 1;
        Some(parse_style_color(style_color, ignorable)?)
    } else {
        None
    };
    if cursor != children.len() {
        return invalid("invalid color choice in chart style reference");
    }
    Ok((color, style_color))
}

fn parse_style_color(node: &Node, ignorable: &HashSet<String>) -> Result<ChartStyleColorValue> {
    let raw = optional_attribute(node, "val", ignorable).map(str::to_owned);
    if let Some(raw) = &raw {
        check_string(raw, 256, "style color value")?;
    }
    reject_attributes(node, &["val"], ignorable, "styleClr")?;
    let index = raw.as_deref().and_then(|value| value.parse::<u32>().ok());
    let automatic = raw.as_deref() == Some("auto");
    Ok(ChartStyleColorValue {
        raw,
        index,
        automatic,
        transforms: parse_transforms(node, ignorable)?,
    })
}

fn parse_marker_layout(node: &Node, ignorable: &HashSet<String>) -> Result<ChartStyleMarkerLayout> {
    if !semantic_children(node, ignorable)?.is_empty() || !node.text.trim().is_empty() {
        return invalid("dataPointMarkerLayout must be empty");
    }
    let symbol = match optional_attribute(node, "symbol", ignorable) {
        None => None,
        Some("circle") => Some(ChartStyleMarkerSymbol::Circle),
        Some("dash") => Some(ChartStyleMarkerSymbol::Dash),
        Some("diamond") => Some(ChartStyleMarkerSymbol::Diamond),
        Some("dot") => Some(ChartStyleMarkerSymbol::Dot),
        Some("plus") => Some(ChartStyleMarkerSymbol::Plus),
        Some("square") => Some(ChartStyleMarkerSymbol::Square),
        Some("star") => Some(ChartStyleMarkerSymbol::Star),
        Some("triangle") => Some(ChartStyleMarkerSymbol::Triangle),
        Some("x") => Some(ChartStyleMarkerSymbol::X),
        Some(_) => return invalid("invalid chart style marker symbol"),
    };
    let size = optional_attribute(node, "size", ignorable)
        .map(|value| {
            value
                .parse::<u8>()
                .map_err(|_| invalid_error("invalid marker size"))
        })
        .transpose()?;
    if size.is_some_and(|value| !(2..=72).contains(&value)) {
        return invalid("chart style marker size is outside 2..=72");
    }
    reject_attributes(
        node,
        &["symbol", "size"],
        ignorable,
        "dataPointMarkerLayout",
    )?;
    Ok(ChartStyleMarkerLayout { symbol, size })
}

fn parse_color(node: &Node, ignorable: &HashSet<String>) -> Result<ChartStyleColor> {
    let kind =
        color_kind(&node.name).ok_or_else(|| invalid_error("unsupported DrawingML color"))?;
    let mut components = Vec::new();
    let value = match kind {
        ChartStyleColorKind::ScRgb => {
            for name in ["r", "g", "b"] {
                let value = required_attribute(node, name, ignorable)?;
                validate_integer(value, "scRGB color component")?;
                components.push((name.to_owned(), value.to_owned()));
            }
            reject_attributes(node, &["r", "g", "b"], ignorable, &node.name)?;
            None
        },
        ChartStyleColorKind::Hsl => {
            for name in ["hue", "sat", "lum"] {
                let value = required_attribute(node, name, ignorable)?;
                validate_integer(value, "HSL color component")?;
                components.push((name.to_owned(), value.to_owned()));
            }
            reject_attributes(node, &["hue", "sat", "lum"], ignorable, &node.name)?;
            None
        },
        ChartStyleColorKind::System => {
            let value = required_attribute(node, "val", ignorable)?.to_owned();
            check_string(&value, 256, "system color")?;
            if let Some(last) = optional_attribute(node, "lastClr", ignorable) {
                validate_hex_rgb(last)?;
                components.push(("lastClr".to_owned(), last.to_owned()));
            }
            reject_attributes(node, &["val", "lastClr"], ignorable, &node.name)?;
            Some(value)
        },
        ChartStyleColorKind::Srgb => {
            let value = required_attribute(node, "val", ignorable)?.to_owned();
            validate_hex_rgb(&value)?;
            reject_attributes(node, &["val"], ignorable, &node.name)?;
            Some(value)
        },
        ChartStyleColorKind::Scheme | ChartStyleColorKind::Preset => {
            let value = required_attribute(node, "val", ignorable)?.to_owned();
            check_string(&value, 256, "DrawingML color")?;
            reject_attributes(node, &["val"], ignorable, &node.name)?;
            Some(value)
        },
    };
    Ok(ChartStyleColor {
        kind,
        value,
        components,
        transforms: parse_transforms(node, ignorable)?,
    })
}

fn parse_transforms(
    node: &Node,
    ignorable: &HashSet<String>,
) -> Result<Vec<ChartStyleColorTransform>> {
    if !node.text.trim().is_empty() {
        return invalid("color or variation contains text");
    }
    let mut transforms = Vec::new();
    for child in semantic_children(node, ignorable)? {
        if transforms.len() >= MAX_TRANSFORMS {
            return limit("chart style color transforms");
        }
        if !is_a(&child.namespace) {
            return invalid("color transform has the wrong namespace");
        }
        let (kind, takes_value) = transform_kind(&child.name).ok_or_else(|| {
            invalid_error(format!("unsupported color transform '{}'", child.name))
        })?;
        if !semantic_children(child, ignorable)?.is_empty() || !child.text.trim().is_empty() {
            return invalid("color transform must be empty");
        }
        let value = if takes_value {
            let value = required_attribute(child, "val", ignorable)?.to_owned();
            validate_integer(&value, "color transform value")?;
            reject_attributes(child, &["val"], ignorable, &child.name)?;
            Some(value)
        } else {
            reject_attributes(child, &[], ignorable, &child.name)?;
            None
        };
        transforms.push(ChartStyleColorTransform { kind, value });
    }
    Ok(transforms)
}

fn take_payload(
    children: &[&Node],
    cursor: &mut usize,
    name: &str,
    ignorable: &HashSet<String>,
) -> Result<Option<ChartStylePayload>> {
    if children
        .get(*cursor)
        .is_some_and(|child| child.namespace == CS && child.name == name)
    {
        let payload = parse_payload(children[*cursor], ignorable)?;
        *cursor += 1;
        Ok(Some(payload))
    } else {
        Ok(None)
    }
}

fn parse_payload(node: &Node, ignorable: &HashSet<String>) -> Result<ChartStylePayload> {
    if !node.text.trim().is_empty() {
        return invalid(format!("cs:{} contains direct text", node.name));
    }
    let attributes = node
        .attributes
        .iter()
        .filter(|attribute| !ignorable.contains(&attribute.namespace))
        .count();
    Ok(ChartStylePayload {
        child_elements: count_semantic_descendants(node, ignorable)?,
        attributes,
    })
}

fn count_semantic_descendants(node: &Node, ignorable: &HashSet<String>) -> Result<usize> {
    let mut count = 0usize;
    let mut stack = node.children.iter().collect::<Vec<_>>();
    while let Some(child) = stack.pop() {
        if ignorable.contains(&child.namespace) {
            continue;
        }
        count = count
            .checked_add(1)
            .ok_or_else(|| invalid_error("payload count overflow"))?;
        stack.extend(child.children.iter());
    }
    Ok(count)
}

fn parse_tree(xml: &[u8]) -> Result<Node> {
    if xml.len() > MAX_XML_BYTES {
        return limit("chart style XML bytes");
    }
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut stack = Vec::<Node>::new();
    let mut root = None;
    let mut nodes = 0usize;
    let mut strings = 0usize;
    loop {
        match reader.read_event().map_err(xml_error)? {
            Event::Start(element) => {
                let node = make_node(&reader, &element, &mut strings)?;
                stack.push(node);
                nodes += 1;
                if stack.len() > MAX_DEPTH || nodes > MAX_NODES {
                    return limit("chart style XML structure");
                }
            },
            Event::Empty(element) => {
                let node = make_node(&reader, &element, &mut strings)?;
                nodes += 1;
                if nodes > MAX_NODES {
                    return limit("chart style XML structure");
                }
                append_node(node, &mut stack, &mut root)?;
            },
            Event::End(_) => {
                let node = stack
                    .pop()
                    .ok_or_else(|| invalid_error("unbalanced chart style XML"))?;
                append_node(node, &mut stack, &mut root)?;
            },
            Event::Text(value) => {
                let value = value
                    .xml_content(XmlVersion::Implicit1_0)
                    .map_err(xml_error)?
                    .into_owned();
                add_strings(&mut strings, value.len())?;
                if let Some(node) = stack.last_mut() {
                    node.text.push_str(&value);
                } else if !value.trim().is_empty() {
                    return invalid("text outside chart style root");
                }
            },
            Event::CData(value) => {
                let value = reader
                    .decoder()
                    .decode(value.as_ref())
                    .map_err(xml_error)?
                    .into_owned();
                add_strings(&mut strings, value.len())?;
                if let Some(node) = stack.last_mut() {
                    node.text.push_str(&value);
                } else if !value.trim().is_empty() {
                    return invalid("CDATA outside chart style root");
                }
            },
            Event::DocType(_) | Event::PI(_) => {
                return invalid("DTD and processing instructions are rejected in chart styles");
            },
            Event::Eof => break,
            _ => {},
        }
    }
    if !stack.is_empty() {
        return invalid("unterminated chart style XML");
    }
    root.ok_or_else(|| invalid_error("missing chart style root"))
}

fn make_node(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    strings: &mut usize,
) -> Result<Node> {
    let namespace = resolved(reader.resolver().resolve_element(element.name()).0)?;
    let name = std::str::from_utf8(element.local_name().as_ref())
        .map_err(xml_error)?
        .to_owned();
    add_strings(strings, namespace.len() + name.len())?;
    let mut attributes = Vec::new();
    let mut namespace_declarations = Vec::new();
    for item in element.attributes().with_checks(true) {
        let item = item.map_err(xml_error)?;
        let raw_name = item.key.as_ref();
        let value = item
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
            .map_err(xml_error)?
            .into_owned();
        add_strings(strings, raw_name.len() + value.len())?;
        if raw_name == b"xmlns" {
            namespace_declarations.push((String::new(), value));
            continue;
        }
        if let Some(prefix) = raw_name.strip_prefix(b"xmlns:") {
            namespace_declarations.push((
                std::str::from_utf8(prefix).map_err(xml_error)?.to_owned(),
                value,
            ));
            continue;
        }
        if attributes.len() >= MAX_ATTRIBUTES {
            return limit("chart style element attributes");
        }
        let (namespace, local) = reader.resolver().resolve_attribute(item.key);
        let namespace = resolved(namespace)?;
        let name = std::str::from_utf8(local.as_ref())
            .map_err(xml_error)?
            .to_owned();
        if attributes
            .iter()
            .any(|existing: &Attribute| existing.namespace == namespace && existing.name == name)
        {
            return invalid("duplicate expanded chart style attribute");
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
        namespace_declarations,
        children: Vec::new(),
        text: String::new(),
    })
}

fn append_node(node: Node, stack: &mut [Node], root: &mut Option<Node>) -> Result<()> {
    if let Some(parent) = stack.last_mut() {
        parent.children.push(node);
    } else if root.replace(node).is_some() {
        return invalid("multiple chart style roots");
    }
    Ok(())
}

fn require_root(root: &Node, name: &str) -> Result<()> {
    if root.namespace != CS || root.name != name {
        return invalid(format!("expected cs:{name} root"));
    }
    if !root.text.trim().is_empty() {
        return invalid(format!("cs:{name} contains direct text"));
    }
    Ok(())
}

fn ignorable_namespaces(root: &Node) -> Result<HashSet<String>> {
    let mut ignorable = HashSet::new();
    for attribute in &root.attributes {
        if attribute.namespace != MC {
            continue;
        }
        match attribute.name.as_str() {
            "Ignorable" => {
                for prefix in attribute.value.split_ascii_whitespace() {
                    let namespace = root
                        .namespace_declarations
                        .iter()
                        .find_map(|(declared, namespace)| (declared == prefix).then_some(namespace))
                        .ok_or_else(|| {
                            invalid_error(format!("unbound mc:Ignorable prefix '{prefix}'"))
                        })?;
                    if [CS, A, A_STRICT, MC].contains(&namespace.as_str()) {
                        return invalid("known chart-style namespaces cannot be mc:Ignorable");
                    }
                    ignorable.insert(namespace.clone());
                }
            },
            "MustUnderstand" if !attribute.value.trim().is_empty() => {
                return invalid("unsupported mc:MustUnderstand namespaces in chart style");
            },
            "MustUnderstand" => {},
            _ => return invalid(format!("unsupported chart style mc:{}", attribute.name)),
        }
    }
    Ok(ignorable)
}

fn semantic_children<'a>(node: &'a Node, ignorable: &HashSet<String>) -> Result<Vec<&'a Node>> {
    let mut children = Vec::new();
    for child in &node.children {
        if child.namespace == MC && child.name == "AlternateContent" {
            return invalid("mc:AlternateContent is not evaluated in inert chart styles");
        }
        if ignorable.contains(&child.namespace) {
            continue;
        }
        children.push(child);
    }
    Ok(children)
}

fn reject_root_attributes(node: &Node, known: &[&str], ignorable: &HashSet<String>) -> Result<()> {
    for attribute in &node.attributes {
        if attribute.namespace == MC || ignorable.contains(&attribute.namespace) {
            continue;
        }
        if !attribute.namespace.is_empty() || !known.contains(&attribute.name.as_str()) {
            return invalid(format!(
                "unsupported cs:{} attribute '{}'",
                node.name, attribute.name
            ));
        }
    }
    Ok(())
}

fn reject_attributes(
    node: &Node,
    known: &[&str],
    ignorable: &HashSet<String>,
    label: &str,
) -> Result<()> {
    for attribute in &node.attributes {
        if ignorable.contains(&attribute.namespace) {
            continue;
        }
        if !attribute.namespace.is_empty() || !known.contains(&attribute.name.as_str()) {
            return invalid(format!(
                "unsupported {label} attribute '{}'",
                attribute.name
            ));
        }
    }
    Ok(())
}

fn required_attribute<'a>(
    node: &'a Node,
    name: &str,
    ignorable: &HashSet<String>,
) -> Result<&'a str> {
    optional_attribute(node, name, ignorable)
        .ok_or_else(|| invalid_error(format!("missing {} {name} attribute", node.name)))
}

fn optional_attribute<'a>(
    node: &'a Node,
    name: &str,
    _ignorable: &HashSet<String>,
) -> Option<&'a str> {
    node.attributes
        .iter()
        .find(|attribute| attribute.namespace.is_empty() && attribute.name == name)
        .map(|attribute| attribute.value.as_str())
}

fn optional_u32_attribute(
    node: &Node,
    name: &str,
    ignorable: &HashSet<String>,
) -> Result<Option<u32>> {
    optional_attribute(node, name, ignorable)
        .map(|value| {
            value
                .parse::<u32>()
                .map_err(|_| invalid_error(format!("invalid {} {name}", node.name)))
        })
        .transpose()
}

fn modifiers(node: &Node, ignorable: &HashSet<String>) -> Result<Vec<String>> {
    let values = optional_attribute(node, "mods", ignorable)
        .unwrap_or_default()
        .split_ascii_whitespace()
        .map(str::to_owned)
        .collect::<Vec<_>>();
    if values.len() > MAX_MODIFIERS || values.iter().any(|value| value.len() > 256) {
        return limit("chart style modifiers");
    }
    Ok(values)
}

fn color_kind(name: &str) -> Option<ChartStyleColorKind> {
    match name {
        "scrgbClr" => Some(ChartStyleColorKind::ScRgb),
        "srgbClr" => Some(ChartStyleColorKind::Srgb),
        "hslClr" => Some(ChartStyleColorKind::Hsl),
        "sysClr" => Some(ChartStyleColorKind::System),
        "schemeClr" => Some(ChartStyleColorKind::Scheme),
        "prstClr" => Some(ChartStyleColorKind::Preset),
        _ => None,
    }
}

fn transform_kind(name: &str) -> Option<(ChartStyleColorTransformKind, bool)> {
    use ChartStyleColorTransformKind::*;
    Some(match name {
        "tint" => (Tint, true),
        "shade" => (Shade, true),
        "comp" => (Complement, false),
        "inv" => (Inverse, false),
        "gray" => (Grayscale, false),
        "alpha" => (Alpha, true),
        "alphaOff" => (AlphaOffset, true),
        "alphaMod" => (AlphaModulation, true),
        "hue" => (Hue, true),
        "hueOff" => (HueOffset, true),
        "hueMod" => (HueModulation, true),
        "sat" => (Saturation, true),
        "satOff" => (SaturationOffset, true),
        "satMod" => (SaturationModulation, true),
        "lum" => (Luminance, true),
        "lumOff" => (LuminanceOffset, true),
        "lumMod" => (LuminanceModulation, true),
        "red" => (Red, true),
        "redOff" => (RedOffset, true),
        "redMod" => (RedModulation, true),
        "green" => (Green, true),
        "greenOff" => (GreenOffset, true),
        "greenMod" => (GreenModulation, true),
        "blue" => (Blue, true),
        "blueOff" => (BlueOffset, true),
        "blueMod" => (BlueModulation, true),
        "gamma" => (Gamma, false),
        "invGamma" => (InverseGamma, false),
        _ => return None,
    })
}

fn is_a(namespace: &str) -> bool {
    namespace == A || namespace == A_STRICT
}

fn validate_hex_rgb(value: &str) -> Result<()> {
    if value.len() != 6 || !value.bytes().all(|byte| byte.is_ascii_hexdigit()) {
        return invalid("invalid six-digit sRGB color");
    }
    Ok(())
}

fn validate_integer(value: &str, label: &str) -> Result<()> {
    value
        .parse::<i64>()
        .map_err(|_| invalid_error(format!("invalid {label}")))?;
    Ok(())
}

fn validate_double(value: &str) -> Result<()> {
    if value.is_empty() {
        return invalid("empty lineWidthScale");
    }
    if !matches!(value, "INF" | "-INF" | "NaN") && value.parse::<f64>().is_err() {
        return invalid("invalid lineWidthScale double");
    }
    Ok(())
}

fn require_content_type(part: &dyn Part, expected: &str, label: &str) -> Result<()> {
    if part.content_type() != expected {
        return invalid(format!("{label} part has the wrong content type"));
    }
    Ok(())
}

fn reject_target(value: &str, label: &str) -> Result<()> {
    let lower = value.to_ascii_lowercase();
    if value.is_empty()
        || value.contains(['?', '#', '\\'])
        || lower.contains("%2e")
        || lower.contains("%2f")
        || lower.contains("%5c")
    {
        return invalid(format!("ambiguous or encoded {label} relationship target"));
    }
    Ok(())
}

fn parent_name(value: &str) -> &str {
    value.rsplit_once('/').map_or("", |(parent, _)| parent)
}

fn check_string(value: &str, maximum: usize, label: &str) -> Result<()> {
    if value.len() > maximum {
        return limit(label);
    }
    Ok(())
}

fn add_strings(total: &mut usize, amount: usize) -> Result<()> {
    *total = total
        .checked_add(amount)
        .ok_or_else(|| invalid_error("chart style string count overflow"))?;
    if *total > MAX_STRING_BYTES {
        return limit("chart style strings");
    }
    Ok(())
}

fn resolved(result: ResolveResult<'_>) -> Result<String> {
    match result {
        ResolveResult::Bound(namespace) => std::str::from_utf8(namespace.as_ref())
            .map(str::to_owned)
            .map_err(xml_error),
        ResolveResult::Unbound => Ok(String::new()),
        ResolveResult::Unknown(prefix) => invalid(format!(
            "unbound chart style namespace prefix '{}'",
            String::from_utf8_lossy(prefix.as_ref())
        )),
    }
}

fn invalid<T>(message: impl Into<String>) -> Result<T> {
    Err(invalid_error(message))
}

fn invalid_error(message: impl Into<String>) -> OoxmlError {
    OoxmlError::InvalidFormat(message.into())
}

fn limit<T>(label: impl Into<String>) -> Result<T> {
    invalid(format!("{} exceeds safety limit", label.into()))
}

fn xml_error(error: impl std::fmt::Display) -> OoxmlError {
    OoxmlError::Xml(error.to_string())
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_opc::{BlobPart, PackURI};

    fn entry(name: &str) -> String {
        format!(
            r#"<cs:{name} mods="alpha beta"><cs:lnRef idx="0"><a:schemeClr val="accent1"/></cs:lnRef><cs:lineWidthScale>1.25</cs:lineWidthScale><cs:fillRef idx="1"/><cs:effectRef idx="2"><cs:styleClr val="auto"><a:tint val="20000"/></cs:styleClr></cs:effectRef><cs:fontRef idx="minor"/><cs:spPr><a:solidFill><a:srgbClr val="112233"/></a:solidFill></cs:spPr></cs:{name}>"#
        )
    }

    fn style_xml() -> String {
        let names = [
            "axisTitle",
            "categoryAxis",
            "chartArea",
            "dataLabel",
            "dataLabelCallout",
            "dataPoint",
            "dataPoint3D",
            "dataPointLine",
            "dataPointMarker",
        ];
        let trailing = [
            "dataPointWireframe",
            "dataTable",
            "downBar",
            "dropLine",
            "errorBar",
            "floor",
            "gridlineMajor",
            "gridlineMinor",
            "hiLoLine",
            "leaderLine",
            "legend",
            "plotArea",
            "plotArea3D",
            "seriesAxis",
            "seriesLine",
            "title",
            "trendline",
            "trendlineLabel",
            "upBar",
            "valueAxis",
            "wall",
        ];
        let mut body = names.into_iter().map(entry).collect::<String>();
        body.push_str(r#"<cs:dataPointMarkerLayout symbol="diamond" size="9"/>"#);
        body.push_str(&trailing.into_iter().map(entry).collect::<String>());
        body.push_str("<cs:extLst/><a14:future/>");
        format!(
            r#"<?xml version="1.0"?><cs:chartStyle xmlns:cs="{CS}" xmlns:a="{A}" xmlns:mc="{MC}" xmlns:a14="urn:ignored:office" mc:Ignorable="a14" id="42">{body}</cs:chartStyle>"#
        )
    }

    fn color_xml(method: &str) -> String {
        format!(
            r#"<cs:colorStyle xmlns:cs="{CS}" xmlns:a="{A}" meth="{method}" id="7"><a:srgbClr val="FF0000"><a:shade val="50000"/></a:srgbClr><a:schemeClr val="accent2"/><cs:variation><a:tint val="20000"/><a:gamma/></cs:variation><cs:extLst/></cs:colorStyle>"#
        )
    }

    fn part(name: &str, content_type: &str, xml: String) -> BlobPart {
        BlobPart::new(
            PackURI::new(name).unwrap(),
            content_type.into(),
            xml.into_bytes(),
        )
    }

    #[test]
    fn parses_producer_styles_as_inert_lossless_metadata() {
        let style = part(
            "/ppt/charts/style1.xml",
            CHART_STYLE_CONTENT_TYPE,
            style_xml(),
        );
        let document = ChartStylePart::from_part(&style).unwrap().parse().unwrap();
        assert_eq!(document.info().id, Some(42));
        assert_eq!(document.info().entries.len(), 30);
        assert_eq!(document.info().entries[0].line_width_scale, "1.25");
        assert_eq!(document.info().entries[0].modifiers, ["alpha", "beta"]);
        assert_eq!(
            document.info().entries[0]
                .shape_properties
                .as_ref()
                .unwrap()
                .child_elements,
            2
        );
        assert_eq!(
            document.info().marker_layout.as_ref().unwrap().size,
            Some(9)
        );
        assert!(document.info().has_extension_list);
        assert_eq!(document.to_xml(), style.blob());

        let colors = part(
            "/ppt/charts/colors1.xml",
            CHART_COLOR_STYLE_CONTENT_TYPE,
            color_xml("vendorMethod"),
        );
        let document = ChartColorStylePart::from_part(&colors)
            .unwrap()
            .parse()
            .unwrap();
        assert_eq!(document.info().method, "vendorMethod");
        assert_eq!(
            document.info().effective_method,
            ChartColorStyleMethod::Cycle
        );
        assert_eq!(document.info().colors.len(), 2);
        assert_eq!(document.info().variations[0].transforms.len(), 2);
        assert_eq!(document.to_xml(), colors.blob());
    }

    #[test]
    fn validates_package_discovery_without_following_external_targets() {
        let mut source = part(
            "/ppt/charts/chartEx1.xml",
            "application/vnd.ms-office.chartex+xml",
            String::new(),
        );
        source.rels_mut().add_relationship(
            CHART_STYLE_RELATIONSHIP_TYPE.into(),
            "style1.xml".into(),
            "rIdStyle".into(),
            false,
        );
        source.rels_mut().add_relationship(
            CHART_COLOR_STYLE_RELATIONSHIP_TYPE.into(),
            "colors1.xml".into(),
            "rIdColors".into(),
            false,
        );
        let mut package = OpcPackage::new();
        package.add_part(Box::new(part(
            "/ppt/charts/style1.xml",
            CHART_STYLE_CONTENT_TYPE,
            style_xml(),
        )));
        package.add_part(Box::new(part(
            "/ppt/charts/colors1.xml",
            CHART_COLOR_STYLE_CONTENT_TYPE,
            color_xml("cycle"),
        )));
        let (style, colors) = discover_chart_styles(&package, &source).unwrap();
        assert_eq!(style.unwrap().info().id, Some(42));
        assert_eq!(colors.unwrap().info().id, Some(7));

        let mut external = part(
            "/ppt/charts/chartEx2.xml",
            "application/vnd.ms-office.chartex+xml",
            String::new(),
        );
        external.rels_mut().add_relationship(
            CHART_STYLE_RELATIONSHIP_TYPE.into(),
            "https://example.test/style.xml".into(),
            "rIdStyle".into(),
            true,
        );
        assert!(discover_chart_styles(&package, &external).is_err());
    }

    #[test]
    fn rejects_hostile_style_grammar_and_relationships() {
        let base = style_xml();
        let cases = [
            base.replace(CS, "urn:spoofed"),
            base.replacen("<cs:axisTitle", "<cs:vendor", 1),
            base.replacen("<cs:axisTitle", "<cs:axisTitle vendor=\"1\"", 1),
            base.replace(
                "<cs:lineWidthScale>1.25</cs:lineWidthScale>",
                "<cs:lineWidthScale>bad</cs:lineWidthScale>",
            ),
            base.replace("size=\"9\"", "size=\"73\""),
            base.replace("<cs:extLst/>", "<cs:extLst/><cs:extLst/>"),
            format!("<!DOCTYPE x [<!ENTITY e SYSTEM 'file:///etc/passwd'>]>{base}"),
        ];
        for xml in cases {
            assert!(
                ChartStylePart::from_part(&part(
                    "/ppt/charts/style.xml",
                    CHART_STYLE_CONTENT_TYPE,
                    xml
                ))
                .unwrap()
                .parse()
                .is_err()
            );
        }
        let colors = color_xml("cycle");
        for xml in [
            colors.replace("<a:srgbClr", "<cs:srgbClr"),
            colors.replace("val=\"FF0000\"", "val=\"GG0000\""),
            colors.replace("<cs:variation>", "<cs:variation><a:vendor/>"),
            colors.replace("meth=\"cycle\"", ""),
        ] {
            assert!(
                ChartColorStylePart::from_part(&part(
                    "/ppt/charts/colors.xml",
                    CHART_COLOR_STYLE_CONTENT_TYPE,
                    xml
                ))
                .unwrap()
                .parse()
                .is_err()
            );
        }

        let mut source = part(
            "/ppt/charts/chartEx.xml",
            "application/vnd.ms-office.chartex+xml",
            String::new(),
        );
        source.rels_mut().add_relationship(
            CHART_STYLE_RELATIONSHIP_TYPE.into(),
            "../styles/style.xml".into(),
            "rId1".into(),
            false,
        );
        assert!(discover_chart_styles(&OpcPackage::new(), &source).is_err());
    }
}
