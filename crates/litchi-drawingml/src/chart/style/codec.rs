//! Bounded XML codecs for Office 2013 chart-style companion parts.

use super::model::{
    Color, ColorDocument, ColorInfo, ColorKind, ColorMethod, ColorValue, Document, Entry,
    EntryKind, FontIndex, FontReference, Info, MarkerLayout, MarkerSymbol, Payload, Reference,
    Transform, TransformKind, Variation,
};
use crate::chart::Limits;
use crate::chart::limits::{
    decoded_attribute_value_bytes, decoded_cdata_value_bytes, decoded_text_value_bytes,
};
use crate::{Error, LimitKind, Result};
use litchi_ooxml_common::xml::attributes::BytesStartExt as _;
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::ResolveResult;
use quick_xml::reader::NsReader;
use std::collections::HashSet;

const CS: &str = "http://schemas.microsoft.com/office/drawing/2012/chartStyle";
const A: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const A_STRICT: &str = "http://purl.oclc.org/ooxml/drawingml/main";
const MC: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const MAX_XML_BYTES: usize = 16 * 1024 * 1024;

impl Document {
    #[must_use]
    pub fn info(&self) -> &Info {
        &self.info
    }

    #[must_use]
    pub fn to_xml(&self) -> Vec<u8> {
        self.xml.clone()
    }
}

impl ColorDocument {
    #[must_use]
    pub fn info(&self) -> &ColorInfo {
        &self.info
    }

    #[must_use]
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

#[derive(Debug)]
struct Node {
    namespace: String,
    name: String,
    attributes: Vec<Attribute>,
    namespace_declarations: Vec<(String, String)>,
    children: Vec<Node>,
    text: String,
}

/// Parse and validate a `cs:chartStyle` XML document.
/// # Errors
///
/// Returns an error when input violates DrawingML constraints, exceeds a configured
/// bound, or an underlying XML, MCE, I/O, or formatting operation fails.
pub fn parse(xml: &[u8], limits: &Limits) -> Result<Document> {
    let limits = limits.validate()?;
    check_source(xml, limits.max_style_bytes)?;
    let processed = process_mce(xml, limits)?;
    let mut document = parse_document(processed.as_ref(), limits)?;
    document.xml = xml.to_vec();
    limits.validate_style_document(&document)?;
    Ok(document)
}

fn parse_document(xml: &[u8], limits: Limits) -> Result<Document> {
    let root = parse_tree(xml, limits)?;
    require_root(&root, "chartStyle")?;
    let ignorable = ignorable_namespaces(&root)?;
    let id = optional_u32_attribute(&root, "id", &ignorable)?;
    reject_root_attributes(&root, &["id"], &ignorable)?;
    let children = semantic_children(&root, &ignorable)?;
    let schema: [(&str, bool, Option<EntryKind>); 31] = [
        ("axisTitle", true, Some(EntryKind::AxisTitle)),
        ("categoryAxis", true, Some(EntryKind::CategoryAxis)),
        ("chartArea", true, Some(EntryKind::ChartArea)),
        ("dataLabel", true, Some(EntryKind::DataLabel)),
        ("dataLabelCallout", false, Some(EntryKind::DataLabelCallout)),
        ("dataPoint", true, Some(EntryKind::DataPoint)),
        ("dataPoint3D", true, Some(EntryKind::DataPoint3D)),
        ("dataPointLine", true, Some(EntryKind::DataPointLine)),
        ("dataPointMarker", true, Some(EntryKind::DataPointMarker)),
        ("dataPointMarkerLayout", false, None),
        (
            "dataPointWireframe",
            true,
            Some(EntryKind::DataPointWireframe),
        ),
        ("dataTable", true, Some(EntryKind::DataTable)),
        ("downBar", true, Some(EntryKind::DownBar)),
        ("dropLine", true, Some(EntryKind::DropLine)),
        ("errorBar", true, Some(EntryKind::ErrorBar)),
        ("floor", true, Some(EntryKind::Floor)),
        ("gridlineMajor", true, Some(EntryKind::GridlineMajor)),
        ("gridlineMinor", true, Some(EntryKind::GridlineMinor)),
        ("hiLoLine", true, Some(EntryKind::HiLoLine)),
        ("leaderLine", true, Some(EntryKind::LeaderLine)),
        ("legend", true, Some(EntryKind::Legend)),
        ("plotArea", true, Some(EntryKind::PlotArea)),
        ("plotArea3D", true, Some(EntryKind::PlotArea3D)),
        ("seriesAxis", true, Some(EntryKind::SeriesAxis)),
        ("seriesLine", true, Some(EntryKind::SeriesLine)),
        ("title", true, Some(EntryKind::Title)),
        ("trendline", true, Some(EntryKind::Trendline)),
        ("trendlineLabel", true, Some(EntryKind::TrendlineLabel)),
        ("upBar", true, Some(EntryKind::UpBar)),
        ("valueAxis", true, Some(EntryKind::ValueAxis)),
        ("wall", true, Some(EntryKind::Wall)),
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
            entries.push(parse_style_entry(child, kind, &ignorable, limits)?);
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
    Ok(Document {
        info: Info {
            id,
            entries,
            marker_layout,
            has_extension_list,
        },
        xml: xml.to_vec(),
    })
}

/// Parse and validate a `cs:colorStyle` XML document.
/// # Errors
///
/// Returns an error when input violates DrawingML constraints, exceeds a configured
/// bound, or an underlying XML, MCE, I/O, or formatting operation fails.
pub fn parse_color(xml: &[u8], limits: &Limits) -> Result<ColorDocument> {
    let limits = limits.validate()?;
    check_source(xml, limits.max_style_bytes)?;
    let processed = process_mce(xml, limits)?;
    let mut document = parse_color_document(processed.as_ref(), limits)?;
    document.xml = xml.to_vec();
    limits.validate_color_document(&document)?;
    Ok(document)
}

fn parse_color_document(xml: &[u8], limits: Limits) -> Result<ColorDocument> {
    let root = parse_tree(xml, limits)?;
    require_root(&root, "colorStyle")?;
    let ignorable = ignorable_namespaces(&root)?;
    let method = required_attribute(&root, "meth", &ignorable)?.to_owned();
    check_string(&method, 256, "chart color style method")?;
    let effective_method = match method.as_str() {
        "withinLinear" => ColorMethod::WithinLinear,
        "acrossLinear" => ColorMethod::AcrossLinear,
        "withinLinearReversed" => ColorMethod::WithinLinearReversed,
        "acrossLinearReversed" => ColorMethod::AcrossLinearReversed,
        _ => ColorMethod::Cycle,
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
        if colors.len() >= limits.max_style_colors {
            return limit(
                LimitKind::ChartStyleColors,
                colors.len() + 1,
                limits.max_style_colors,
            );
        }
        colors.push(parse_drawing_color(child, &ignorable, limits)?);
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
        if variations.len() >= limits.max_style_variations {
            return limit(
                LimitKind::ChartStyleVariations,
                variations.len() + 1,
                limits.max_style_variations,
            );
        }
        let child = children[cursor];
        reject_attributes(child, &[], &ignorable, "variation")?;
        if !child.text.trim().is_empty() {
            return invalid("chart color style variation contains text");
        }
        variations.push(Variation {
            transforms: parse_transforms(child, &ignorable, limits)?,
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
    Ok(ColorDocument {
        info: ColorInfo {
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
    kind: EntryKind,
    ignorable: &HashSet<String>,
    limits: Limits,
) -> Result<Entry> {
    let modifiers = modifiers(node, ignorable, limits)?;
    reject_attributes(node, &["mods"], ignorable, &node.name)?;
    if !node.text.trim().is_empty() {
        return invalid("chart style entry contains text");
    }
    let children = semantic_children(node, ignorable)?;
    let mut cursor = 0usize;
    let line_reference =
        parse_required_reference(children.get(cursor), "lnRef", ignorable, limits)?;
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
    let fill_reference =
        parse_required_reference(children.get(cursor), "fillRef", ignorable, limits)?;
    cursor += 1;
    let effect_reference =
        parse_required_reference(children.get(cursor), "effectRef", ignorable, limits)?;
    cursor += 1;
    let font_reference = parse_required_font_reference(children.get(cursor), ignorable, limits)?;
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
    Ok(Entry {
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
    limits: Limits,
) -> Result<Reference> {
    let node = node
        .copied()
        .filter(|node| node.namespace == CS && node.name == expected)
        .ok_or_else(|| invalid_error(format!("missing cs:{expected}")))?;
    let index = required_attribute(node, "idx", ignorable)?
        .parse::<u32>()
        .map_err(|_error| invalid_error(format!("invalid {expected} index")))?;
    let modifiers = modifiers(node, ignorable, limits)?;
    reject_attributes(node, &["idx", "mods"], ignorable, expected)?;
    let (color, style_color) = parse_reference_colors(node, ignorable, limits)?;
    Ok(Reference {
        index,
        modifiers,
        color,
        style_color,
    })
}

fn parse_required_font_reference(
    node: Option<&&Node>,
    ignorable: &HashSet<String>,
    limits: Limits,
) -> Result<FontReference> {
    let node = node
        .copied()
        .filter(|node| node.namespace == CS && node.name == "fontRef")
        .ok_or_else(|| invalid_error("missing cs:fontRef"))?;
    let index = match required_attribute(node, "idx", ignorable)? {
        "major" => FontIndex::Major,
        "minor" => FontIndex::Minor,
        "none" => FontIndex::None,
        _ => return invalid("invalid fontRef index"),
    };
    let modifiers = modifiers(node, ignorable, limits)?;
    reject_attributes(node, &["idx", "mods"], ignorable, "fontRef")?;
    let (color, style_color) = parse_reference_colors(node, ignorable, limits)?;
    Ok(FontReference {
        index,
        modifiers,
        color,
        style_color,
    })
}

fn parse_reference_colors(
    node: &Node,
    ignorable: &HashSet<String>,
    limits: Limits,
) -> Result<(Option<Color>, Option<ColorValue>)> {
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
        Some(parse_drawing_color(color, ignorable, limits)?)
    } else {
        None
    };
    let style_color = children
        .get(cursor)
        .filter(|child| child.namespace == CS && child.name == "styleClr");
    let style_color = if let Some(style_color) = style_color {
        cursor += 1;
        Some(parse_style_color(style_color, ignorable, limits)?)
    } else {
        None
    };
    if cursor != children.len() {
        return invalid("invalid color choice in chart style reference");
    }
    Ok((color, style_color))
}

fn parse_style_color(
    node: &Node,
    ignorable: &HashSet<String>,
    limits: Limits,
) -> Result<ColorValue> {
    let raw = optional_attribute(node, "val", ignorable).map(str::to_owned);
    if let Some(raw) = &raw {
        check_string(raw, 256, "style color value")?;
    }
    reject_attributes(node, &["val"], ignorable, "styleClr")?;
    let index = raw.as_deref().and_then(|value| value.parse::<u32>().ok());
    let automatic = raw.as_deref() == Some("auto");
    Ok(ColorValue {
        raw,
        index,
        automatic,
        transforms: parse_transforms(node, ignorable, limits)?,
    })
}

fn parse_marker_layout(node: &Node, ignorable: &HashSet<String>) -> Result<MarkerLayout> {
    if !semantic_children(node, ignorable)?.is_empty() || !node.text.trim().is_empty() {
        return invalid("dataPointMarkerLayout must be empty");
    }
    let symbol = match optional_attribute(node, "symbol", ignorable) {
        None => None,
        Some("circle") => Some(MarkerSymbol::Circle),
        Some("dash") => Some(MarkerSymbol::Dash),
        Some("diamond") => Some(MarkerSymbol::Diamond),
        Some("dot") => Some(MarkerSymbol::Dot),
        Some("plus") => Some(MarkerSymbol::Plus),
        Some("square") => Some(MarkerSymbol::Square),
        Some("star") => Some(MarkerSymbol::Star),
        Some("triangle") => Some(MarkerSymbol::Triangle),
        Some("x") => Some(MarkerSymbol::X),
        Some(_) => return invalid("invalid chart style marker symbol"),
    };
    let size = optional_attribute(node, "size", ignorable)
        .map(|value| {
            value
                .parse::<u8>()
                .map_err(|_error| invalid_error("invalid marker size"))
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
    Ok(MarkerLayout { symbol, size })
}

fn parse_drawing_color(node: &Node, ignorable: &HashSet<String>, limits: Limits) -> Result<Color> {
    let kind =
        color_kind(&node.name).ok_or_else(|| invalid_error("unsupported DrawingML color"))?;
    let mut components = Vec::new();
    let value = match kind {
        ColorKind::ScRgb => {
            for name in ["r", "g", "b"] {
                let value = required_attribute(node, name, ignorable)?;
                validate_integer(value, "scRGB color component")?;
                components.push((name.to_owned(), value.to_owned()));
            }
            reject_attributes(node, &["r", "g", "b"], ignorable, &node.name)?;
            None
        },
        ColorKind::Hsl => {
            for name in ["hue", "sat", "lum"] {
                let value = required_attribute(node, name, ignorable)?;
                validate_integer(value, "HSL color component")?;
                components.push((name.to_owned(), value.to_owned()));
            }
            reject_attributes(node, &["hue", "sat", "lum"], ignorable, &node.name)?;
            None
        },
        ColorKind::System => {
            let value = required_attribute(node, "val", ignorable)?.to_owned();
            check_string(&value, 256, "system color")?;
            if let Some(last) = optional_attribute(node, "lastClr", ignorable) {
                validate_hex_rgb(last)?;
                components.push(("lastClr".to_owned(), last.to_owned()));
            }
            reject_attributes(node, &["val", "lastClr"], ignorable, &node.name)?;
            Some(value)
        },
        ColorKind::Srgb => {
            let value = required_attribute(node, "val", ignorable)?.to_owned();
            validate_hex_rgb(&value)?;
            reject_attributes(node, &["val"], ignorable, &node.name)?;
            Some(value)
        },
        ColorKind::Scheme | ColorKind::Preset => {
            let value = required_attribute(node, "val", ignorable)?.to_owned();
            check_string(&value, 256, "DrawingML color")?;
            reject_attributes(node, &["val"], ignorable, &node.name)?;
            Some(value)
        },
    };
    Ok(Color {
        kind,
        value,
        components,
        transforms: parse_transforms(node, ignorable, limits)?,
    })
}

fn parse_transforms(
    node: &Node,
    ignorable: &HashSet<String>,
    limits: Limits,
) -> Result<Vec<Transform>> {
    if !node.text.trim().is_empty() {
        return invalid("color or variation contains text");
    }
    let mut transforms = Vec::new();
    for child in semantic_children(node, ignorable)? {
        if transforms.len() >= limits.max_style_transforms {
            return limit(
                LimitKind::ChartStyleTransforms,
                transforms.len() + 1,
                limits.max_style_transforms,
            );
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
        transforms.push(Transform { kind, value });
    }
    Ok(transforms)
}

fn take_payload(
    children: &[&Node],
    cursor: &mut usize,
    name: &str,
    ignorable: &HashSet<String>,
) -> Result<Option<Payload>> {
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

fn parse_payload(node: &Node, ignorable: &HashSet<String>) -> Result<Payload> {
    if !node.text.trim().is_empty() {
        return invalid(format!("cs:{} contains direct text", node.name));
    }
    let attributes = node
        .attributes
        .iter()
        .filter(|attribute| !ignorable.contains(&attribute.namespace))
        .count();
    Ok(Payload {
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

fn parse_tree(xml: &[u8], limits: Limits) -> Result<Node> {
    let maximum = limits.max_style_bytes.min(MAX_XML_BYTES);
    if xml.len() > maximum {
        return limit(LimitKind::XmlBytes, xml.len(), maximum);
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
                let depth = stack.len().checked_add(1).ok_or_else(|| {
                    Error::limit(LimitKind::XmlDepth, usize::MAX, limits.max_style_depth)
                })?;
                if depth > limits.max_style_depth {
                    return limit(LimitKind::XmlDepth, depth, limits.max_style_depth);
                }
                nodes = nodes.checked_add(1).ok_or_else(|| {
                    Error::limit(LimitKind::XmlNodes, usize::MAX, limits.max_style_nodes)
                })?;
                if nodes > limits.max_style_nodes {
                    return limit(LimitKind::XmlNodes, nodes, limits.max_style_nodes);
                }
                let node = make_node(&reader, &element, &mut strings, limits)?;
                stack.push(node);
            },
            Event::Empty(element) => {
                let depth = stack.len().checked_add(1).ok_or_else(|| {
                    Error::limit(LimitKind::XmlDepth, usize::MAX, limits.max_style_depth)
                })?;
                if depth > limits.max_style_depth {
                    return limit(LimitKind::XmlDepth, depth, limits.max_style_depth);
                }
                nodes = nodes.checked_add(1).ok_or_else(|| {
                    Error::limit(LimitKind::XmlNodes, usize::MAX, limits.max_style_nodes)
                })?;
                if nodes > limits.max_style_nodes {
                    return limit(LimitKind::XmlNodes, nodes, limits.max_style_nodes);
                }
                let node = make_node(&reader, &element, &mut strings, limits)?;
                append_node(node, &mut stack, &mut root)?;
            },
            Event::End(_) => {
                let node = stack
                    .pop()
                    .ok_or_else(|| invalid_error("unbalanced chart style XML"))?;
                append_node(node, &mut stack, &mut root)?;
            },
            Event::Text(value) => {
                let decoded_length = decoded_text_value_bytes(
                    value.as_ref(),
                    reader.decoder(),
                    XmlVersion::Implicit1_0,
                )
                .map_err(Error::Xml)?;
                preflight_strings(strings, decoded_length, limits.max_style_string_bytes)?;
                let value = value
                    .xml_content(XmlVersion::Implicit1_0)
                    .map_err(xml_error)?
                    .into_owned();
                add_strings(&mut strings, value.len(), limits.max_style_string_bytes)?;
                if let Some(node) = stack.last_mut() {
                    node.text.push_str(&value);
                } else if !value.trim().is_empty() {
                    return invalid("text outside chart style root");
                }
            },
            Event::CData(value) => {
                let decoded_length = decoded_cdata_value_bytes(value.as_ref(), reader.decoder())
                    .map_err(Error::Xml)?;
                preflight_strings(strings, decoded_length, limits.max_style_string_bytes)?;
                let value = reader
                    .decoder()
                    .decode(value.as_ref())
                    .map_err(xml_error)?
                    .into_owned();
                add_strings(&mut strings, value.len(), limits.max_style_string_bytes)?;
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
            Event::GeneralRef(reference) => {
                let value = match reference
                    .resolve_char_ref()
                    .map_err(|error| Error::Xml(error.to_string()))?
                {
                    Some(character) => character.to_string(),
                    None => {
                        let name = reference
                            .decode()
                            .map_err(|error| Error::Xml(error.to_string()))?;
                        quick_xml::escape::resolve_predefined_entity(&name)
                            .ok_or_else(|| Error::Xml("unknown XML entity".into()))?
                            .to_owned()
                    },
                };
                preflight_strings(strings, value.len(), limits.max_style_string_bytes)?;
                if let Some(node) = stack.last_mut() {
                    node.text.push_str(&value);
                } else if !value.trim().is_empty() {
                    return invalid("text outside chart style root");
                }
                add_strings(&mut strings, value.len(), limits.max_style_string_bytes)?;
            },
            Event::Comment(_) | Event::Decl(_) => {},
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
    limits: Limits,
) -> Result<Node> {
    let namespace = resolved(reader.resolver().resolve_element(element.name()).0)?;
    let name = std::str::from_utf8(element.local_name().as_ref())
        .map_err(xml_error)?
        .to_owned();
    add_strings(
        strings,
        namespace.len() + name.len(),
        limits.max_style_string_bytes,
    )?;
    let mut attribute_count = 0usize;
    // Admit the complete attribute budget before decoding any values into
    // owned strings. The semantic loop below still uses the checked iterator
    // so expanded-name duplicates retain the style codec's strict grammar.
    for item in element.unchecked_attributes() {
        let item = item.map_err(xml_error)?;
        let raw_name = item.key.as_ref();
        if raw_name == b"xmlns" || raw_name.starts_with(b"xmlns:") {
            continue;
        }
        attribute_count = attribute_count.checked_add(1).ok_or_else(|| {
            Error::limit(
                LimitKind::XmlAttributes,
                usize::MAX,
                limits.max_style_attributes,
            )
        })?;
        if attribute_count > limits.max_style_attributes {
            return limit(
                LimitKind::XmlAttributes,
                attribute_count,
                limits.max_style_attributes,
            );
        }
    }
    let mut attributes = Vec::new();
    let mut namespace_declarations = Vec::new();
    for item in element.checked_attributes() {
        let item = item.map_err(xml_error)?;
        let raw_name = item.key.as_ref();
        let value_length = decoded_attribute_value_bytes(item.value.as_ref(), reader.decoder())
            .map_err(xml_error)?;
        if value_length > limits.max_attribute_value_bytes {
            return limit(
                LimitKind::AttributeValueBytes,
                value_length,
                limits.max_attribute_value_bytes,
            );
        }
        preflight_strings(
            *strings,
            raw_name.len().saturating_add(value_length),
            limits.max_style_string_bytes,
        )?;
        let value = item
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
            .map_err(xml_error)?
            .into_owned();
        add_strings(
            strings,
            raw_name.len() + value.len(),
            limits.max_style_string_bytes,
        )?;
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
        if attributes.len() >= limits.max_style_attributes {
            return limit(
                LimitKind::XmlAttributes,
                attributes.len() + 1,
                limits.max_style_attributes,
            );
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
                .map_err(|_error| invalid_error(format!("invalid {} {name}", node.name)))
        })
        .transpose()
}

fn modifiers(node: &Node, ignorable: &HashSet<String>, limits: Limits) -> Result<Vec<String>> {
    let mut values = Vec::new();
    for modifier in optional_attribute(node, "mods", ignorable)
        .unwrap_or_default()
        .split_ascii_whitespace()
    {
        let observed = values.len().checked_add(1).ok_or_else(|| {
            Error::limit(
                LimitKind::ChartStyleModifiers,
                usize::MAX,
                limits.max_style_modifiers,
            )
        })?;
        if observed > limits.max_style_modifiers {
            return limit(
                LimitKind::ChartStyleModifiers,
                observed,
                limits.max_style_modifiers,
            );
        }
        if modifier.len() > 256 {
            return limit(LimitKind::ChartStyleModifierBytes, modifier.len(), 256);
        }
        values.push(modifier.to_owned());
    }
    Ok(values)
}

fn color_kind(name: &str) -> Option<ColorKind> {
    match name {
        "scrgbClr" => Some(ColorKind::ScRgb),
        "srgbClr" => Some(ColorKind::Srgb),
        "hslClr" => Some(ColorKind::Hsl),
        "sysClr" => Some(ColorKind::System),
        "schemeClr" => Some(ColorKind::Scheme),
        "prstClr" => Some(ColorKind::Preset),
        _ => None,
    }
}

fn transform_kind(name: &str) -> Option<(TransformKind, bool)> {
    use TransformKind::{
        Alpha, AlphaModulation, AlphaOffset, Blue, BlueModulation, BlueOffset, Complement, Gamma,
        Grayscale, Green, GreenModulation, GreenOffset, Hue, HueModulation, HueOffset, Inverse,
        InverseGamma, Luminance, LuminanceModulation, LuminanceOffset, Red, RedModulation,
        RedOffset, Saturation, SaturationModulation, SaturationOffset, Shade, Tint,
    };
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
        .map_err(|_error| invalid_error(format!("invalid {label}")))?;
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

fn check_string(value: &str, maximum: usize, _label: &str) -> Result<()> {
    if value.len() > maximum {
        return limit(LimitKind::TextBytes, value.len(), maximum);
    }
    Ok(())
}

fn add_strings(total: &mut usize, amount: usize, maximum: usize) -> Result<()> {
    *total = total
        .checked_add(amount)
        .ok_or_else(|| invalid_error("chart style string count overflow"))?;
    if *total > maximum {
        return limit(LimitKind::ChartStyleStrings, *total, maximum);
    }
    Ok(())
}

fn preflight_strings(total: usize, amount: usize, maximum: usize) -> Result<()> {
    let observed = total
        .checked_add(amount)
        .ok_or_else(|| Error::limit(LimitKind::ChartStyleStrings, usize::MAX, maximum))?;
    if observed > maximum {
        return limit(LimitKind::ChartStyleStrings, observed, maximum);
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

fn invalid_error(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

fn limit<T>(kind: LimitKind, observed: usize, maximum: usize) -> Result<T> {
    Err(Error::limit(kind, observed, maximum))
}

fn xml_error(error: impl std::fmt::Display) -> Error {
    Error::Xml(error.to_string())
}

fn check_source(xml: &[u8], maximum: usize) -> Result<()> {
    if xml.len() > maximum {
        return limit(LimitKind::XmlBytes, xml.len(), maximum);
    }
    Ok(())
}

fn process_mce(xml: &[u8], limits: Limits) -> Result<std::borrow::Cow<'_, [u8]>> {
    litchi_ooxml_common::mce::process_markup_compatibility(
        xml,
        &litchi_ooxml_common::mce::Capabilities::default(),
        &litchi_ooxml_common::mce::Limits {
            max_input_bytes: limits.max_style_bytes,
            max_output_bytes: limits.max_style_bytes,
            max_depth: limits.max_style_depth,
            max_namespace_bindings: 4_096,
            max_directive_tokens: 4_096,
            max_choices_per_alternate: 1_024,
            max_attributes_per_element: limits.max_style_attributes,
        },
    )
    .map(|output| output.xml)
    .map_err(|error| match error {
        litchi_ooxml_common::mce::Error::LimitExceeded {
            kind,
            observed,
            maximum,
        } => Error::limit(
            match kind {
                litchi_ooxml_common::mce::LimitKind::InputBytes => LimitKind::XmlBytes,
                litchi_ooxml_common::mce::LimitKind::OutputBytes => LimitKind::XmlBytes,
                litchi_ooxml_common::mce::LimitKind::Depth => LimitKind::XmlDepth,
                litchi_ooxml_common::mce::LimitKind::Attributes => LimitKind::XmlAttributes,
                _ => LimitKind::XmlNodes,
            },
            observed,
            maximum,
        ),
        litchi_ooxml_common::mce::Error::InvalidLimit {
            kind,
            value,
            maximum,
            ..
        } => Error::invalid_limit(
            match kind {
                litchi_ooxml_common::mce::LimitKind::InputBytes => LimitKind::XmlBytes,
                litchi_ooxml_common::mce::LimitKind::OutputBytes => LimitKind::XmlBytes,
                litchi_ooxml_common::mce::LimitKind::Depth => LimitKind::XmlDepth,
                litchi_ooxml_common::mce::LimitKind::Attributes => LimitKind::XmlAttributes,
                _ => LimitKind::XmlNodes,
            },
            value,
            maximum,
        ),
        other => Error::Mce(other),
    })
}

#[cfg(test)]
mod tests {
    use super::*;

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

    #[test]
    fn parses_producer_styles_as_inert_lossless_metadata() {
        let style = style_xml();
        let document = parse(style.as_bytes(), &Limits::DEFAULT).unwrap();
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
        assert_eq!(document.to_xml(), style.as_bytes());

        let colors = color_xml("vendorMethod");
        let document = parse_color(colors.as_bytes(), &Limits::DEFAULT).unwrap();
        assert_eq!(document.info().method, "vendorMethod");
        assert_eq!(document.info().effective_method, ColorMethod::Cycle);
        assert_eq!(document.info().colors.len(), 2);
        assert_eq!(document.info().variations[0].transforms.len(), 2);
        assert_eq!(document.to_xml(), colors.as_bytes());
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
            assert!(parse(xml.as_bytes(), &Limits::DEFAULT).is_err());
        }
        let colors = color_xml("cycle");
        for xml in [
            colors.replace("<a:srgbClr", "<cs:srgbClr"),
            colors.replace("val=\"FF0000\"", "val=\"GG0000\""),
            colors.replace("<cs:variation>", "<cs:variation><a:vendor/>"),
            colors.replace("meth=\"cycle\"", ""),
        ] {
            assert!(parse_color(xml.as_bytes(), &Limits::DEFAULT).is_err());
        }
    }

    #[test]
    fn decodes_style_text_references_before_budgeting() {
        let xml = style_xml().replace(
            "<cs:lineWidthScale>1.25</cs:lineWidthScale>",
            "<cs:lineWidthScale>1&#46;25</cs:lineWidthScale>",
        );
        let document = parse(xml.as_bytes(), &Limits::DEFAULT).unwrap();
        assert_eq!(document.info().entries[0].line_width_scale, "1.25");
    }
}
