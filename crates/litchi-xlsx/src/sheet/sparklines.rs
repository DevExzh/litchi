#![allow(dead_code)]

use litchi_core::sheet::Result as SheetResult;
use litchi_core::{id::generate_guid_braced, xml::escape_xml};
use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, QName, ResolveResult};
use quick_xml::reader::NsReader;
use ryu::Buffer as RyuBuffer;
use std::fmt::Write as FmtWrite;

use crate::raw::namespace::is_spreadsheetml_name;
use litchi_ooxml_common::xml::{decode_xml_reference, unqualified_attribute_value};
pub use litchi_sheet::sparkline::{AxisType, EmptyCells, SparklineType};

/// Backwards-compatible XLSX spelling for [`SparklineType`].
pub type Type = SparklineType;
/// Backwards-compatible XLSX spelling for [`EmptyCells`].
pub type DisplayEmptyCellsAs = EmptyCells;
/// Backwards-compatible XLSX spelling for [`AxisType`].
pub type AxisMinMax = AxisType;

fn sparkline_type_as_str(value: SparklineType) -> &'static str {
    match value {
        SparklineType::Line => "line",
        SparklineType::Column => "column",
        SparklineType::WinLoss => "stacked",
    }
}

fn parse_sparkline_type(value: &str) -> Option<SparklineType> {
    match value {
        "line" => Some(SparklineType::Line),
        "column" => Some(SparklineType::Column),
        "stacked" => Some(SparklineType::WinLoss),
        _ => None,
    }
}

fn empty_cells_as_str(value: EmptyCells) -> &'static str {
    match value {
        EmptyCells::Zero => "zero",
        EmptyCells::Gap => "gap",
        EmptyCells::Span => "span",
    }
}

fn parse_empty_cells(value: &str) -> Option<EmptyCells> {
    match value {
        "zero" => Some(EmptyCells::Zero),
        "gap" => Some(EmptyCells::Gap),
        "span" => Some(EmptyCells::Span),
        _ => None,
    }
}

fn axis_type_as_str(value: AxisType) -> &'static str {
    match value {
        AxisType::Individual => "individual",
        AxisType::Group => "group",
        AxisType::Custom => "custom",
    }
}

fn parse_axis_type(value: &str) -> Option<AxisType> {
    match value {
        "individual" => Some(AxisType::Individual),
        "group" => Some(AxisType::Group),
        "custom" => Some(AxisType::Custom),
        _ => None,
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct Color {
    pub rgb: String,
    pub indexed: Option<u32>,
    pub automatic: Option<bool>,
    pub theme: Option<u32>,
    pub tint: Option<f64>,
}

impl Color {
    pub fn new(rgb: String) -> Self {
        Self {
            rgb,
            indexed: None,
            automatic: None,
            theme: None,
            tint: None,
        }
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct GroupColors {
    pub series: Option<Color>,
    pub negative: Option<Color>,
    pub axis: Option<Color>,
    pub markers: Option<Color>,
    pub first: Option<Color>,
    pub last: Option<Color>,
    pub high: Option<Color>,
    pub low: Option<Color>,
}

impl Default for GroupColors {
    fn default() -> Self {
        Self {
            series: Some(Color::new("FF000000".to_string())),
            negative: None,
            axis: None,
            markers: None,
            first: None,
            last: None,
            high: None,
            low: None,
        }
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct GroupOptions {
    pub display_empty_cells_as: DisplayEmptyCellsAs,
    pub date_axis: bool,
    pub display_hidden: bool,
    pub display_x_axis: bool,
    pub markers: bool,
    pub high: bool,
    pub low: bool,
    pub first: bool,
    pub last: bool,
    pub negative: bool,
    pub right_to_left: bool,
    pub min_axis_type: AxisMinMax,
    pub max_axis_type: AxisMinMax,
    pub manual_min: Option<f64>,
    pub manual_max: Option<f64>,
    pub line_weight: Option<f64>,
}

impl Default for GroupOptions {
    fn default() -> Self {
        Self {
            display_empty_cells_as: DisplayEmptyCellsAs::Zero,
            date_axis: false,
            display_hidden: false,
            display_x_axis: false,
            markers: false,
            high: false,
            low: false,
            first: false,
            last: false,
            negative: false,
            right_to_left: false,
            min_axis_type: AxisMinMax::Individual,
            max_axis_type: AxisMinMax::Individual,
            manual_min: None,
            manual_max: None,
            line_weight: None,
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Item {
    pub data_range: String,
    pub location: String,
}

#[derive(Debug, Clone, PartialEq)]
pub struct Group {
    pub sparkline_type: Type,
    pub sparklines: Vec<Item>,
    pub options: GroupOptions,
    pub colors: GroupColors,
    pub extra_attributes: Vec<(String, String)>,
}

impl Group {
    pub fn new(sparkline_type: Type) -> Self {
        Self {
            sparkline_type,
            sparklines: Vec::new(),
            options: GroupOptions::default(),
            colors: GroupColors::default(),
            extra_attributes: Vec::new(),
        }
    }

    pub fn push(&mut self, sparkline: Item) {
        self.sparklines.push(sparkline);
    }
}

pub fn parse_groups(content: &str) -> SheetResult<Vec<Group>> {
    let content = litchi_ooxml_common::mce::process_str(content)?;
    Parser::parse(content.as_ref())
}

pub fn write_groups_ext(xml: &mut String, groups: &[Group]) -> SheetResult<()> {
    if groups.is_empty() {
        return Ok(());
    }

    xml.push_str(r#"<ext uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">"#);
    xml.push_str(
        r#"<x14:sparklineGroups xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">"#,
    );

    if groups.len() >= 231 {
        return Err("sparklineGroups must contain fewer than 231 sparklineGroup elements".into());
    }
    for group in groups {
        if group.sparklines.is_empty() {
            return Err("sparklineGroup must contain at least one sparkline".into());
        }
        validate_group_options(&group.options)?;
        xml.push_str("<x14:sparklineGroup");

        // Excel omits the attribute for the default type (line).
        if group.sparkline_type != Type::Line {
            xml.push_str(" type=\"");
            xml.push_str(sparkline_type_as_str(group.sparkline_type));
            xml.push('"');
        }

        let has_attr = |name: &str| group.extra_attributes.iter().any(|(k, _)| k == name);

        if !has_attr("displayEmptyCellsAs") {
            xml.push_str(" displayEmptyCellsAs=\"");
            xml.push_str(empty_cells_as_str(group.options.display_empty_cells_as));
            xml.push('"');
        }
        if group.options.date_axis && !has_attr("dateAxis") {
            xml.push_str(" dateAxis=\"1\"");
        }
        if group.options.display_hidden && !has_attr("displayHidden") {
            xml.push_str(" displayHidden=\"1\"");
        }
        if group.options.display_x_axis && !has_attr("displayXAxis") {
            xml.push_str(" displayXAxis=\"1\"");
        }
        if group.options.markers && !has_attr("markers") {
            xml.push_str(" markers=\"1\"");
        }
        if group.options.high && !has_attr("high") {
            xml.push_str(" high=\"1\"");
        }
        if group.options.low && !has_attr("low") {
            xml.push_str(" low=\"1\"");
        }
        if group.options.first && !has_attr("first") {
            xml.push_str(" first=\"1\"");
        }
        if group.options.last && !has_attr("last") {
            xml.push_str(" last=\"1\"");
        }
        if group.options.negative && !has_attr("negative") {
            xml.push_str(" negative=\"1\"");
        }
        if group.options.right_to_left && !has_attr("rightToLeft") {
            xml.push_str(" rightToLeft=\"1\"");
        }
        if group.options.min_axis_type != AxisMinMax::Individual && !has_attr("minAxisType") {
            xml.push_str(" minAxisType=\"");
            xml.push_str(axis_type_as_str(group.options.min_axis_type));
            xml.push('"');
        }
        if group.options.max_axis_type != AxisMinMax::Individual && !has_attr("maxAxisType") {
            xml.push_str(" maxAxisType=\"");
            xml.push_str(axis_type_as_str(group.options.max_axis_type));
            xml.push('"');
        }

        if let Some(min) = group.options.manual_min {
            if group.options.min_axis_type != AxisMinMax::Custom && !has_attr("minAxisType") {
                return Err("manualMin requires minAxisType=custom".into());
            }
            if !has_attr("manualMin") {
                let mut b = RyuBuffer::new();
                xml.push_str(" manualMin=\"");
                xml.push_str(b.format(min));
                xml.push('"');
            }
        }
        if let Some(max) = group.options.manual_max {
            if group.options.max_axis_type != AxisMinMax::Custom && !has_attr("maxAxisType") {
                return Err("manualMax requires maxAxisType=custom".into());
            }
            if !has_attr("manualMax") {
                let mut b = RyuBuffer::new();
                xml.push_str(" manualMax=\"");
                xml.push_str(b.format(max));
                xml.push('"');
            }
        }
        if let Some(w) = group.options.line_weight
            && !has_attr("lineWeight")
        {
            let mut b = RyuBuffer::new();
            xml.push_str(" lineWeight=\"");
            xml.push_str(b.format(w));
            xml.push('"');
        }

        if !has_attr("xr2:uid") {
            xml.push_str(" xr2:uid=\"");
            xml.push_str(&generate_guid_braced());
            xml.push('"');
        }

        for (index, (k, v)) in group.extra_attributes.iter().enumerate() {
            validate_extra_attribute_name(k)?;
            if group.extra_attributes[..index]
                .iter()
                .any(|(previous, _)| previous == k)
            {
                return Err(format!("duplicate sparkline group attribute '{k}'").into());
            }
            xml.push(' ');
            xml.push_str(k);
            xml.push_str("=\"");
            xml.push_str(&escape_xml(v));
            xml.push('"');
        }

        xml.push('>');

        for (name, color) in [
            ("colorSeries", group.colors.series.as_ref()),
            ("colorNegative", group.colors.negative.as_ref()),
            ("colorAxis", group.colors.axis.as_ref()),
            ("colorMarkers", group.colors.markers.as_ref()),
            ("colorFirst", group.colors.first.as_ref()),
            ("colorLast", group.colors.last.as_ref()),
            ("colorHigh", group.colors.high.as_ref()),
            ("colorLow", group.colors.low.as_ref()),
        ] {
            if let Some(color) = color {
                write_color(xml, name, color)?;
            }
        }
        xml.push_str("<x14:sparklines>");

        for sp in &group.sparklines {
            if sp.data_range.trim().is_empty() {
                return Err("sparkline data formula cannot be empty".into());
            }
            litchi_sheet::Cell::from_a1(sp.location.trim())?;
            xml.push_str("<x14:sparkline>");
            xml.push_str("<xm:f>");
            xml.push_str(&escape_xml(&sp.data_range));
            xml.push_str("</xm:f>");
            xml.push_str("<xm:sqref>");
            xml.push_str(&escape_xml(&sp.location));
            xml.push_str("</xm:sqref>");
            xml.push_str("</x14:sparkline>");
        }

        xml.push_str("</x14:sparklines>");
        xml.push_str("</x14:sparklineGroup>");
    }

    xml.push_str("</x14:sparklineGroups>");
    xml.push_str("</ext>");
    Ok(())
}

fn write_color(xml: &mut String, name: &str, color: &Color) -> SheetResult<()> {
    write!(xml, "<x14:{name}").map_err(|error| format!("XML write error: {error}"))?;
    if !color.rgb.is_empty() {
        if color.rgb.len() != 8 || !color.rgb.bytes().all(|byte| byte.is_ascii_hexdigit()) {
            return Err(format!("invalid sparkline RGB color '{}'", color.rgb).into());
        }
        write!(xml, r#" rgb="{}""#, escape_xml(&color.rgb))
            .map_err(|error| format!("XML write error: {error}"))?;
    }
    if let Some(indexed) = color.indexed {
        write!(xml, r#" indexed="{indexed}""#)
            .map_err(|error| format!("XML write error: {error}"))?;
    }
    if let Some(automatic) = color.automatic {
        write!(xml, r#" auto="{}""#, if automatic { 1 } else { 0 })
            .map_err(|error| format!("XML write error: {error}"))?;
    }
    if let Some(theme) = color.theme {
        write!(xml, r#" theme="{theme}""#).map_err(|error| format!("XML write error: {error}"))?;
    }
    if let Some(tint) = color.tint {
        if !tint.is_finite() || !(-1.0..=1.0).contains(&tint) {
            return Err("sparkline color tint must be finite and between -1 and 1".into());
        }
        let mut buffer = RyuBuffer::new();
        write!(xml, r#" tint="{}""#, buffer.format(tint))
            .map_err(|error| format!("XML write error: {error}"))?;
    }
    xml.push_str("/>");
    Ok(())
}

fn validate_extra_attribute_name(name: &str) -> SheetResult<()> {
    let mut parts = name.split(':');
    let first = parts.next().unwrap_or_default();
    let second = parts.next();
    if parts.next().is_some()
        || !valid_xml_name_part(first)
        || second.is_some_and(|part| !valid_xml_name_part(part))
    {
        return Err(format!("invalid sparkline group attribute name '{name}'").into());
    }
    if let Some(prefix) = second.map(|_| first)
        && !matches!(prefix, "xr" | "xr2" | "xr3" | "x14ac" | "x14" | "xm")
    {
        return Err(format!("undeclared sparkline group attribute prefix '{prefix}'").into());
    }
    if name == "xmlns" || name.starts_with("xmlns:") {
        return Err("sparkline group attributes cannot declare namespaces".into());
    }
    Ok(())
}

fn valid_xml_name_part(part: &str) -> bool {
    let mut bytes = part.bytes();
    bytes
        .next()
        .is_some_and(|byte| byte.is_ascii_alphabetic() || byte == b'_')
        && bytes.all(|byte| byte.is_ascii_alphanumeric() || matches!(byte, b'_' | b'-' | b'.'))
}

const SPARKLINE_EXT_URI: &str = "{05C60535-1F16-4fd2-B633-F4F36F0B64E0}";
const SPARKLINE_NAMESPACE: &[u8] = b"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
const EXCEL_MAIN_NAMESPACE: &[u8] = b"http://schemas.microsoft.com/office/excel/2006/main";
const MAX_SPARKLINE_GROUPS: usize = 230;

#[derive(Clone, Copy, PartialEq, Eq)]
enum Context {
    Worksheet,
    ExtensionList,
    SparklineExtension,
    Groups,
    Group,
    Sparklines,
    Item,
    Formula,
    Location,
    Other,
}

#[derive(Default)]
struct Pending {
    data_range: String,
    location: String,
    saw_formula: bool,
    saw_location: bool,
}

struct Parser {
    groups: Vec<Group>,
    group: Option<Group>,
    sparkline: Option<Pending>,
    groups_start: Option<usize>,
    group_saw_sparklines: bool,
    seen_extension: bool,
    extension_saw_groups: bool,
    seen_colors: u8,
}

impl Parser {
    fn new() -> Self {
        Self {
            groups: Vec::new(),
            group: None,
            sparkline: None,
            groups_start: None,
            group_saw_sparklines: false,
            seen_extension: false,
            extension_saw_groups: false,
            seen_colors: 0,
        }
    }

    fn parse(content: &str) -> SheetResult<Vec<Group>> {
        let mut reader = NsReader::from_reader(content.as_bytes());
        let mut parser = Self::new();
        let mut stack = Vec::new();
        let mut closed_root = false;

        loop {
            let decoder = reader.decoder();
            let event = reader.read_event()?.into_owned();
            let resolver = reader.resolver().clone();
            let (namespace, event) = resolver.resolve_event(event);
            match event {
                Event::Start(element) if stack.is_empty() => {
                    if closed_root
                        || !is_spreadsheetml_name(&namespace, element.name(), b"worksheet")
                    {
                        return Err(
                            "sparkline source must have one SpreadsheetML worksheet root".into(),
                        );
                    }
                    stack.push(Context::Worksheet);
                },
                Event::Empty(element) if stack.is_empty() => {
                    if closed_root
                        || !is_spreadsheetml_name(&namespace, element.name(), b"worksheet")
                    {
                        return Err(
                            "sparkline source must have one SpreadsheetML worksheet root".into(),
                        );
                    }
                    closed_root = true;
                },
                Event::Start(element) => {
                    let parent = *stack
                        .last()
                        .ok_or("sparkline parser is missing its root context")?;
                    let context = parser.start(parent, &namespace, &element, decoder)?;
                    stack.push(context);
                },
                Event::Empty(element) => {
                    let parent = *stack
                        .last()
                        .ok_or("sparkline parser is missing its root context")?;
                    let context = parser.start(parent, &namespace, &element, decoder)?;
                    parser.finish(context)?;
                },
                Event::Text(text) => {
                    if let Some(context) = sparkline_text_context(&stack) {
                        parser.push_text(&text.decode()?, context)?;
                    }
                },
                Event::CData(text) => {
                    if let Some(context) = sparkline_text_context(&stack) {
                        parser.push_text(&text.decode()?, context)?;
                    }
                },
                Event::GeneralRef(reference) => {
                    if let Some(context) = sparkline_text_context(&stack) {
                        parser.push_text(&decode_xml_reference(&reference)?, context)?;
                    }
                },
                Event::End(element) => {
                    let context = stack
                        .pop()
                        .ok_or("sparkline XML has a closing element outside its root")?;
                    parser.finish(context)?;
                    if context == Context::Worksheet {
                        if !is_spreadsheetml_name(&namespace, element.name(), b"worksheet") {
                            return Err(
                                "sparkline XML has an invalid worksheet closing element".into()
                            );
                        }
                        closed_root = true;
                    }
                },
                Event::Eof if !closed_root || !stack.is_empty() => {
                    return Err(
                        "sparkline source has a missing or unterminated worksheet root".into(),
                    );
                },
                Event::Eof => break,
                _ => {},
            }
        }
        Ok(parser.groups)
    }

    fn start(
        &mut self,
        parent: Context,
        namespace: &ResolveResult<'_>,
        element: &BytesStart<'_>,
        decoder: Decoder,
    ) -> SheetResult<Context> {
        if parent == Context::Worksheet
            && is_spreadsheetml_name(namespace, element.name(), b"extLst")
        {
            return Ok(Context::ExtensionList);
        }
        if parent == Context::ExtensionList
            && is_spreadsheetml_name(namespace, element.name(), b"ext")
        {
            let uri = unqualified_attribute_value(element, b"uri", decoder)?;
            if uri
                .as_deref()
                .is_some_and(|uri| uri.eq_ignore_ascii_case(SPARKLINE_EXT_URI))
            {
                if self.seen_extension {
                    return Err("duplicate worksheet sparkline extension".into());
                }
                self.seen_extension = true;
                return Ok(Context::SparklineExtension);
            }
            return Ok(Context::Other);
        }
        if parent == Context::SparklineExtension
            && is_name(
                namespace,
                element.name(),
                b"sparklineGroups",
                SPARKLINE_NAMESPACE,
            )
        {
            if self.extension_saw_groups {
                return Err("duplicate sparklineGroups element".into());
            }
            self.extension_saw_groups = true;
            self.groups_start = Some(self.groups.len());
            return Ok(Context::Groups);
        }
        if parent == Context::Groups
            && is_name(
                namespace,
                element.name(),
                b"sparklineGroup",
                SPARKLINE_NAMESPACE,
            )
        {
            self.start_group(element, decoder)?;
            return Ok(Context::Group);
        }
        if parent == Context::Group {
            if let Some((slot, bit)) = color_slot(element.name().local_name().as_ref())
                && is_name(
                    namespace,
                    element.name(),
                    element.name().local_name().as_ref(),
                    SPARKLINE_NAMESPACE,
                )
            {
                self.color(slot, bit, element, decoder)?;
                return Ok(Context::Other);
            }
            if is_name(
                namespace,
                element.name(),
                b"sparklines",
                SPARKLINE_NAMESPACE,
            ) {
                if self.group_saw_sparklines {
                    return Err("duplicate sparklines element in sparkline group".into());
                }
                self.group_saw_sparklines = true;
                return Ok(Context::Sparklines);
            }
        }
        if parent == Context::Sparklines
            && is_name(namespace, element.name(), b"sparkline", SPARKLINE_NAMESPACE)
        {
            if self.sparkline.is_some() {
                return Err("nested worksheet sparkline".into());
            }
            self.sparkline = Some(Pending::default());
            return Ok(Context::Item);
        }
        if parent == Context::Item && is_name(namespace, element.name(), b"f", EXCEL_MAIN_NAMESPACE)
        {
            let sparkline = self
                .sparkline
                .as_mut()
                .ok_or("sparkline formula outside a sparkline")?;
            if sparkline.saw_formula {
                return Err("duplicate sparkline formula".into());
            }
            sparkline.saw_formula = true;
            return Ok(Context::Formula);
        }
        if parent == Context::Item
            && is_name(namespace, element.name(), b"sqref", EXCEL_MAIN_NAMESPACE)
        {
            let sparkline = self
                .sparkline
                .as_mut()
                .ok_or("sparkline location outside a sparkline")?;
            if sparkline.saw_location {
                return Err("duplicate sparkline location".into());
            }
            sparkline.saw_location = true;
            return Ok(Context::Location);
        }
        Ok(Context::Other)
    }

    fn start_group(&mut self, element: &BytesStart<'_>, decoder: Decoder) -> SheetResult<()> {
        if self.group.is_some() {
            return Err("nested sparkline group".into());
        }
        if self.groups.len() >= MAX_SPARKLINE_GROUPS {
            return Err(
                "sparklineGroups must contain fewer than 231 sparklineGroup elements".into(),
            );
        }
        let sparkline_type = parse_enum_attribute(
            element,
            b"type",
            decoder,
            "sparkline type",
            parse_sparkline_type,
        )?
        .unwrap_or(Type::Line);
        let mut group = Group::new(sparkline_type);
        group.options.display_empty_cells_as = parse_enum_attribute(
            element,
            b"displayEmptyCellsAs",
            decoder,
            "sparkline empty-cell mode",
            parse_empty_cells,
        )?
        .unwrap_or(DisplayEmptyCellsAs::Zero);
        group.options.date_axis = sparkline_bool(element, b"dateAxis", decoder)?.unwrap_or(false);
        group.options.display_hidden =
            sparkline_bool(element, b"displayHidden", decoder)?.unwrap_or(false);
        group.options.display_x_axis =
            sparkline_bool(element, b"displayXAxis", decoder)?.unwrap_or(false);
        group.options.markers = sparkline_bool(element, b"markers", decoder)?.unwrap_or(false);
        group.options.high = sparkline_bool(element, b"high", decoder)?.unwrap_or(false);
        group.options.low = sparkline_bool(element, b"low", decoder)?.unwrap_or(false);
        group.options.first = sparkline_bool(element, b"first", decoder)?.unwrap_or(false);
        group.options.last = sparkline_bool(element, b"last", decoder)?.unwrap_or(false);
        group.options.negative = sparkline_bool(element, b"negative", decoder)?.unwrap_or(false);
        group.options.right_to_left =
            sparkline_bool(element, b"rightToLeft", decoder)?.unwrap_or(false);
        group.options.min_axis_type = parse_enum_attribute(
            element,
            b"minAxisType",
            decoder,
            "sparkline minimum-axis type",
            parse_axis_type,
        )?
        .unwrap_or(AxisMinMax::Individual);
        group.options.max_axis_type = parse_enum_attribute(
            element,
            b"maxAxisType",
            decoder,
            "sparkline maximum-axis type",
            parse_axis_type,
        )?
        .unwrap_or(AxisMinMax::Individual);
        group.options.manual_min = sparkline_f64(element, b"manualMin", decoder)?;
        group.options.manual_max = sparkline_f64(element, b"manualMax", decoder)?;
        group.options.line_weight = sparkline_f64(element, b"lineWeight", decoder)?;
        group.extra_attributes = extra_group_attributes(element, decoder)?;
        self.group = Some(group);
        self.group_saw_sparklines = false;
        self.seen_colors = 0;
        Ok(())
    }

    fn color(
        &mut self,
        slot: ColorSlot,
        bit: u8,
        element: &BytesStart<'_>,
        decoder: Decoder,
    ) -> SheetResult<()> {
        if self.seen_colors & bit != 0 {
            return Err("duplicate sparkline group color element".into());
        }
        self.seen_colors |= bit;
        let rgb = unqualified_attribute_value(element, b"rgb", decoder)?.unwrap_or_default();
        if !rgb.is_empty() && (rgb.len() != 8 || !rgb.bytes().all(|byte| byte.is_ascii_hexdigit()))
        {
            return Err(format!("invalid sparkline RGB color '{rgb}'").into());
        }
        let mut color = Color::new(rgb);
        color.indexed = sparkline_u32(element, b"indexed", decoder)?;
        color.automatic = sparkline_bool(element, b"auto", decoder)?;
        color.theme = sparkline_u32(element, b"theme", decoder)?;
        color.tint = sparkline_f64(element, b"tint", decoder)?;
        if color.tint.is_some_and(|tint| !(-1.0..=1.0).contains(&tint)) {
            return Err("sparkline color tint must be between -1 and 1".into());
        }
        let color = Some(color);
        let colors = &mut self
            .group
            .as_mut()
            .ok_or("sparkline color outside a group")?
            .colors;
        match slot {
            ColorSlot::Series => colors.series = color,
            ColorSlot::Negative => colors.negative = color,
            ColorSlot::Axis => colors.axis = color,
            ColorSlot::Markers => colors.markers = color,
            ColorSlot::First => colors.first = color,
            ColorSlot::Last => colors.last = color,
            ColorSlot::High => colors.high = color,
            ColorSlot::Low => colors.low = color,
        }
        Ok(())
    }

    fn push_text(&mut self, value: &str, context: Context) -> SheetResult<()> {
        let sparkline = self
            .sparkline
            .as_mut()
            .ok_or("sparkline text outside a sparkline")?;
        match context {
            Context::Formula => sparkline.data_range.push_str(value),
            Context::Location => sparkline.location.push_str(value),
            _ => return Err("unexpected sparkline text context".into()),
        }
        Ok(())
    }

    fn finish(&mut self, context: Context) -> SheetResult<()> {
        match context {
            Context::Item => {
                let pending = self.sparkline.take().ok_or("missing pending sparkline")?;
                let data_range = pending.data_range.trim().to_string();
                let location = pending.location.trim().to_string();
                if !pending.saw_formula || data_range.is_empty() {
                    return Err("sparkline is missing its data formula".into());
                }
                if !pending.saw_location || location.is_empty() {
                    return Err("sparkline is missing its location".into());
                }
                litchi_sheet::Cell::from_a1(&location)?;
                self.group
                    .as_mut()
                    .ok_or("sparkline outside a group")?
                    .sparklines
                    .push(Item {
                        data_range,
                        location,
                    });
            },
            Context::Sparklines
                if self
                    .group
                    .as_ref()
                    .is_none_or(|group| group.sparklines.is_empty()) =>
            {
                return Err("sparklines element contains no sparkline".into());
            },
            Context::Group => {
                let group = self.group.take().ok_or("missing pending sparkline group")?;
                if !self.group_saw_sparklines || group.sparklines.is_empty() {
                    return Err("sparklineGroup contains no sparklines".into());
                }
                validate_group_options(&group.options)?;
                self.groups.push(group);
            },
            Context::Groups => {
                let start = self
                    .groups_start
                    .take()
                    .ok_or("missing sparklineGroups start")?;
                if self.groups.len() == start {
                    return Err("sparklineGroups contains no sparklineGroup".into());
                }
            },
            Context::SparklineExtension if !self.extension_saw_groups => {
                return Err("sparkline extension contains no sparklineGroups".into());
            },
            _ => {},
        }
        Ok(())
    }
}

#[derive(Clone, Copy)]
enum ColorSlot {
    Series,
    Negative,
    Axis,
    Markers,
    First,
    Last,
    High,
    Low,
}

fn color_slot(local_name: &[u8]) -> Option<(ColorSlot, u8)> {
    match local_name {
        b"colorSeries" => Some((ColorSlot::Series, 1)),
        b"colorNegative" => Some((ColorSlot::Negative, 2)),
        b"colorAxis" => Some((ColorSlot::Axis, 4)),
        b"colorMarkers" => Some((ColorSlot::Markers, 8)),
        b"colorFirst" => Some((ColorSlot::First, 16)),
        b"colorLast" => Some((ColorSlot::Last, 32)),
        b"colorHigh" => Some((ColorSlot::High, 64)),
        b"colorLow" => Some((ColorSlot::Low, 128)),
        _ => None,
    }
}

fn is_name(
    namespace: &ResolveResult<'_>,
    name: QName<'_>,
    local_name: &[u8],
    expected_namespace: &[u8],
) -> bool {
    name.local_name().as_ref() == local_name
        && matches!(namespace, ResolveResult::Bound(Namespace(value)) if *value == expected_namespace)
}

fn sparkline_text_context(stack: &[Context]) -> Option<Context> {
    stack
        .last()
        .copied()
        .filter(|context| matches!(context, Context::Formula | Context::Location))
}

fn sparkline_bool(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
) -> SheetResult<Option<bool>> {
    unqualified_attribute_value(element, name, decoder)?
        .map(|value| match value.as_str() {
            "1" | "true" => Ok(true),
            "0" | "false" => Ok(false),
            _ => Err(format!(
                "invalid sparkline boolean {}='{value}'",
                String::from_utf8_lossy(name)
            )
            .into()),
        })
        .transpose()
}

fn sparkline_f64(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
) -> SheetResult<Option<f64>> {
    unqualified_attribute_value(element, name, decoder)?
        .map(|value| {
            let number = value.parse::<f64>().map_err(|_| {
                format!(
                    "invalid sparkline number {}='{value}'",
                    String::from_utf8_lossy(name)
                )
            })?;
            if !number.is_finite() {
                return Err(format!(
                    "sparkline number {} must be finite",
                    String::from_utf8_lossy(name)
                )
                .into());
            }
            Ok(number)
        })
        .transpose()
}

fn sparkline_u32(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
) -> SheetResult<Option<u32>> {
    unqualified_attribute_value(element, name, decoder)?
        .map(|value| {
            value.parse::<u32>().map_err(|_| {
                format!(
                    "invalid sparkline unsigned integer {}='{value}'",
                    String::from_utf8_lossy(name)
                )
                .into()
            })
        })
        .transpose()
}

fn parse_enum_attribute<T>(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    description: &str,
    parse: impl FnOnce(&str) -> Option<T>,
) -> SheetResult<Option<T>> {
    unqualified_attribute_value(element, name, decoder)?
        .map(|value| parse(&value).ok_or_else(|| format!("invalid {description} '{value}'").into()))
        .transpose()
}

fn validate_group_options(options: &GroupOptions) -> SheetResult<()> {
    if options.manual_min.is_some() && options.min_axis_type != AxisMinMax::Custom {
        return Err("manualMin requires minAxisType=custom".into());
    }
    if options.manual_max.is_some() && options.max_axis_type != AxisMinMax::Custom {
        return Err("manualMax requires maxAxisType=custom".into());
    }
    for (name, value) in [
        ("manualMin", options.manual_min),
        ("manualMax", options.manual_max),
        ("lineWeight", options.line_weight),
    ] {
        if value.is_some_and(|value| !value.is_finite()) {
            return Err(format!("sparkline {name} must be finite").into());
        }
    }
    Ok(())
}

fn extra_group_attributes(
    element: &BytesStart<'_>,
    decoder: Decoder,
) -> SheetResult<Vec<(String, String)>> {
    const KNOWN: &[&[u8]] = &[
        b"type",
        b"displayEmptyCellsAs",
        b"dateAxis",
        b"displayHidden",
        b"displayXAxis",
        b"markers",
        b"high",
        b"low",
        b"first",
        b"last",
        b"negative",
        b"rightToLeft",
        b"minAxisType",
        b"maxAxisType",
        b"manualMin",
        b"manualMax",
        b"lineWeight",
    ];
    let mut extra = Vec::new();
    for attribute in element.attributes() {
        let attribute = attribute?;
        let key = attribute.key.as_ref();
        if key == b"xmlns" || key.starts_with(b"xmlns:") {
            continue;
        }
        if !key.contains(&b':') && KNOWN.contains(&key) {
            continue;
        }
        let name = std::str::from_utf8(key)?.to_string();
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)?
            .into_owned();
        extra.push((name, value));
    }
    Ok(extra)
}

#[cfg(test)]
mod tests {
    use super::*;

    const STRICT_S: &str = "http://purl.oclc.org/ooxml/spreadsheetml/main";

    fn worksheet(extension_contents: &str) -> String {
        format!(
            r#"<s:worksheet xmlns:s="{STRICT_S}"
                    xmlns:sl="{x14}" xmlns:m="{xm}"
                    xmlns:f="urn:foreign" xmlns:xr2="urn:revision">
                <f:extLst><s:ext uri="{uri}"><sl:sparklineGroups/></s:ext></f:extLst>
                <s:extLst><s:ext uri="{uri}">{extension_contents}</s:ext></s:extLst>
            </s:worksheet>"#,
            x14 = String::from_utf8_lossy(SPARKLINE_NAMESPACE),
            xm = String::from_utf8_lossy(EXCEL_MAIN_NAMESPACE),
            uri = SPARKLINE_EXT_URI.to_ascii_lowercase(),
        )
    }

    fn group_xml(attributes: &str, contents: &str) -> String {
        worksheet(&format!(
            r#"<sl:sparklineGroups><sl:sparklineGroup {attributes}>{contents}</sl:sparklineGroup></sl:sparklineGroups>"#
        ))
    }

    fn sparkline_xml(formula: &str, location: &str) -> String {
        format!(
            "<sl:sparklines><sl:sparkline><m:f>{formula}</m:f><m:sqref>{location}</m:sqref></sl:sparkline></sl:sparklines>"
        )
    }

    #[test]
    fn parses_namespace_aware_sparkline_groups() {
        let xml = group_xml(
            r#"type="stacked" displayEmptyCellsAs="span" dateAxis="1"
                displayHidden="true" displayXAxis="1" markers="1" high="1"
                low="0" first="true" last="false" negative="1" rightToLeft="0"
                minAxisType="custom" maxAxisType="group" manualMin="-2.5"
                lineWeight="1.25" xr2:uid="{ABC}""#,
            r#"<sl:colorSeries rgb="FF102030"/>
                    <sl:colorAxis theme="4" tint="-0.25"/>
                    <sl:sparklines>
                        <sl:sparkline><m:f>'Data &amp; More'!A1:A3</m:f><m:sqref>B2</m:sqref></sl:sparkline>
                        <sl:sparkline><m:f>Sheet1!C1:C3</m:f><m:sqref>C2</m:sqref></sl:sparkline>
                    </sl:sparklines>"#,
        );
        let groups = parse_groups(&xml).unwrap();

        assert_eq!(groups.len(), 1);
        let group = &groups[0];
        assert_eq!(group.sparkline_type, Type::WinLoss);
        assert_eq!(
            group.options.display_empty_cells_as,
            DisplayEmptyCellsAs::Span
        );
        assert!(group.options.date_axis);
        assert!(group.options.display_hidden);
        assert!(group.options.display_x_axis);
        assert!(group.options.markers);
        assert!(group.options.high);
        assert!(!group.options.low);
        assert_eq!(group.options.min_axis_type, AxisMinMax::Custom);
        assert_eq!(group.options.max_axis_type, AxisMinMax::Group);
        assert_eq!(group.options.manual_min, Some(-2.5));
        assert_eq!(group.options.line_weight, Some(1.25));
        assert_eq!(group.colors.series.as_ref().unwrap().rgb, "FF102030");
        assert_eq!(group.colors.axis.as_ref().unwrap().theme, Some(4));
        assert_eq!(group.colors.axis.as_ref().unwrap().tint, Some(-0.25));
        assert_eq!(group.sparklines.len(), 2);
        assert_eq!(group.sparklines[0].data_range, "'Data & More'!A1:A3");
        assert_eq!(group.sparklines[0].location, "B2");
        assert_eq!(
            group.extra_attributes,
            vec![("xr2:uid".to_string(), "{ABC}".to_string())]
        );
    }

    #[test]
    fn rejects_malformed_sparkline_extensions() {
        let valid_sparkline = sparkline_xml("Sheet1!A1:A3", "B2");
        let invalid = [
            worksheet(""),
            worksheet("<sl:sparklineGroups/>"),
            group_xml("", ""),
            group_xml("type=\"winLoss\"", &valid_sparkline),
            group_xml("displayHidden=\"TRUE\"", &valid_sparkline),
            group_xml("displayEmptyCellsAs=\"blank\"", &valid_sparkline),
            group_xml("minAxisType=\"customish\"", &valid_sparkline),
            group_xml("manualMin=\"NaN\" minAxisType=\"custom\"", &valid_sparkline),
            group_xml("manualMin=\"1\"", &valid_sparkline),
            group_xml("", "<sl:sparklines/>"),
            group_xml(
                "",
                "<sl:sparklines><sl:sparkline><m:sqref>B2</m:sqref></sl:sparkline></sl:sparklines>",
            ),
            group_xml(
                "",
                "<sl:sparklines><sl:sparkline><m:f>Sheet1!A1:A3</m:f></sl:sparkline></sl:sparklines>",
            ),
            group_xml("", &sparkline_xml("Sheet1!A1:A3", "B2:C3")),
            group_xml(
                "",
                &format!(
                    "<sl:colorSeries rgb=\"FF000000\"/><sl:colorSeries rgb=\"FFFFFFFF\"/>{valid_sparkline}"
                ),
            ),
            group_xml(
                "",
                &format!("<sl:colorSeries rgb=\"123\"/>{valid_sparkline}"),
            ),
            group_xml(
                "",
                &format!("<sl:colorAxis theme=\"4\" tint=\"2\"/>{valid_sparkline}"),
            ),
            group_xml("", &format!("{valid_sparkline}{valid_sparkline}")),
        ];

        for xml in invalid {
            assert!(
                parse_groups(&xml).is_err(),
                "accepted invalid sparkline XML: {xml}"
            );
        }

        let group = format!("<sl:sparklineGroup>{valid_sparkline}</sl:sparklineGroup>");
        let too_many = worksheet(&format!(
            "<sl:sparklineGroups>{}</sl:sparklineGroups>",
            group.repeat(MAX_SPARKLINE_GROUPS + 1)
        ));
        assert!(parse_groups(&too_many).is_err());

        let duplicate_extension = format!(
            r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                    xmlns:sl="{}" xmlns:m="{}">
                <extLst><ext uri="{SPARKLINE_EXT_URI}"><sl:sparklineGroups><sl:sparklineGroup>{valid_sparkline}</sl:sparklineGroup></sl:sparklineGroups></ext>
                <ext uri="{SPARKLINE_EXT_URI}"><sl:sparklineGroups><sl:sparklineGroup>{valid_sparkline}</sl:sparklineGroup></sl:sparklineGroups></ext></extLst>
            </worksheet>"#,
            String::from_utf8_lossy(SPARKLINE_NAMESPACE),
            String::from_utf8_lossy(EXCEL_MAIN_NAMESPACE),
        );
        assert!(parse_groups(&duplicate_extension).is_err());
    }

    #[test]
    fn writer_preserves_group_membership_and_round_trips() {
        let mut group = Group::new(Type::WinLoss);
        group.options.date_axis = true;
        group.colors.axis = Some(Color {
            rgb: String::new(),
            indexed: None,
            automatic: None,
            theme: Some(4),
            tint: Some(-0.25),
        });
        group.push(Item {
            data_range: "Sheet1!A1:A3".to_string(),
            location: "B2".to_string(),
        });
        group.push(Item {
            data_range: "Sheet1!C1:C3".to_string(),
            location: "C2".to_string(),
        });

        let mut extension = String::new();
        write_groups_ext(&mut extension, &[group]).unwrap();
        assert_eq!(extension.matches("<x14:sparklineGroup ").count(), 1);
        assert_eq!(extension.matches("<x14:sparkline>").count(), 2);
        assert!(extension.contains("type=\"stacked\""));
        assert!(extension.contains("dateAxis=\"1\""));

        let xml = format!(
            r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                xmlns:xr2="urn:revision"><extLst>{extension}</extLst></worksheet>"#
        );
        let parsed = parse_groups(&xml).unwrap();
        assert_eq!(parsed.len(), 1);
        assert_eq!(parsed[0].sparklines.len(), 2);
        assert_eq!(parsed[0].sparkline_type, Type::WinLoss);
        assert!(parsed[0].options.date_axis);
        assert_eq!(parsed[0].colors.axis.as_ref().unwrap().theme, Some(4));
        assert_eq!(parsed[0].colors.axis.as_ref().unwrap().tint, Some(-0.25));
    }

    #[test]
    fn writer_rejects_invalid_groups() {
        let empty = Group::new(Type::Line);
        let mut xml = String::new();
        assert!(write_groups_ext(&mut xml, &[empty]).is_err());

        let mut invalid = Group::new(Type::Line);
        invalid.options.manual_min = Some(1.0);
        invalid.push(Item {
            data_range: "Sheet1!A1:A2".to_string(),
            location: "B2".to_string(),
        });
        let mut xml = String::new();
        assert!(write_groups_ext(&mut xml, &[invalid]).is_err());

        let mut valid = Group::new(Type::Line);
        valid.push(Item {
            data_range: "Sheet1!A1:A2".to_string(),
            location: "B2".to_string(),
        });
        let mut invalid_attribute = valid.clone();
        invalid_attribute
            .extra_attributes
            .push(("bad name".to_string(), "value".to_string()));
        let mut xml = String::new();
        assert!(write_groups_ext(&mut xml, &[invalid_attribute]).is_err());

        let mut xml = String::new();
        assert!(write_groups_ext(&mut xml, &vec![valid; MAX_SPARKLINE_GROUPS + 1]).is_err());
    }
}
