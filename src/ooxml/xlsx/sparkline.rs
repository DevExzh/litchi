use crate::common::{
    id::generate_guid_braced,
    xml::{escape_xml, unescape_xml},
};
use crate::sheet::Result as SheetResult;
use memchr;
use ryu::Buffer as RyuBuffer;
use std::fmt::Write as FmtWrite;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SparklineType {
    Line,
    Column,
    WinLoss,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SparklineDisplayEmptyCellsAs {
    Gap,
    Zero,
    Span,
}

impl SparklineDisplayEmptyCellsAs {
    fn as_str(self) -> &'static str {
        match self {
            Self::Gap => "gap",
            Self::Zero => "zero",
            Self::Span => "span",
        }
    }

    fn parse(s: &str) -> Option<Self> {
        match s {
            "gap" => Some(Self::Gap),
            "zero" => Some(Self::Zero),
            "span" => Some(Self::Span),
            _ => None,
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SparklineAxisMinMax {
    Individual,
    Group,
    Custom,
}

impl SparklineAxisMinMax {
    fn as_str(self) -> &'static str {
        match self {
            Self::Individual => "individual",
            Self::Group => "group",
            Self::Custom => "custom",
        }
    }

    fn parse(s: &str) -> Option<Self> {
        match s {
            "individual" => Some(Self::Individual),
            "group" => Some(Self::Group),
            "custom" => Some(Self::Custom),
            _ => None,
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SparklineColor {
    pub rgb: String,
}

impl SparklineColor {
    pub fn new(rgb: String) -> Self {
        Self { rgb }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SparklineGroupColors {
    pub series: Option<SparklineColor>,
    pub negative: Option<SparklineColor>,
    pub axis: Option<SparklineColor>,
    pub markers: Option<SparklineColor>,
    pub first: Option<SparklineColor>,
    pub last: Option<SparklineColor>,
    pub high: Option<SparklineColor>,
    pub low: Option<SparklineColor>,
}

impl Default for SparklineGroupColors {
    fn default() -> Self {
        Self {
            series: Some(SparklineColor::new("FF000000".to_string())),
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
pub struct SparklineGroupOptions {
    pub display_empty_cells_as: SparklineDisplayEmptyCellsAs,
    pub display_hidden: bool,
    pub display_x_axis: bool,
    pub markers: bool,
    pub high: bool,
    pub low: bool,
    pub first: bool,
    pub last: bool,
    pub negative: bool,
    pub right_to_left: bool,
    pub min_axis_type: SparklineAxisMinMax,
    pub max_axis_type: SparklineAxisMinMax,
    pub manual_min: Option<f64>,
    pub manual_max: Option<f64>,
    pub line_weight: Option<f64>,
}

impl Default for SparklineGroupOptions {
    fn default() -> Self {
        Self {
            display_empty_cells_as: SparklineDisplayEmptyCellsAs::Gap,
            display_hidden: false,
            display_x_axis: false,
            markers: false,
            high: false,
            low: false,
            first: false,
            last: false,
            negative: false,
            right_to_left: false,
            min_axis_type: SparklineAxisMinMax::Individual,
            max_axis_type: SparklineAxisMinMax::Individual,
            manual_min: None,
            manual_max: None,
            line_weight: None,
        }
    }
}

impl SparklineType {
    fn as_str(self) -> &'static str {
        match self {
            Self::Line => "line",
            Self::Column => "column",
            Self::WinLoss => "winLoss",
        }
    }

    fn parse(s: &str) -> Option<Self> {
        match s {
            "line" => Some(Self::Line),
            "column" => Some(Self::Column),
            "winLoss" => Some(Self::WinLoss),
            _ => None,
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Sparkline {
    pub data_range: String,
    pub location: String,
}

#[derive(Debug, Clone, PartialEq)]
pub struct SparklineGroup {
    pub sparkline_type: SparklineType,
    pub sparklines: Vec<Sparkline>,
    pub options: SparklineGroupOptions,
    pub colors: SparklineGroupColors,
    pub extra_attributes: Vec<(String, String)>,
}

impl SparklineGroup {
    pub fn new(sparkline_type: SparklineType) -> Self {
        Self {
            sparkline_type,
            sparklines: Vec::new(),
            options: SparklineGroupOptions::default(),
            colors: SparklineGroupColors::default(),
            extra_attributes: Vec::new(),
        }
    }

    pub fn push(&mut self, sparkline: Sparkline) {
        self.sparklines.push(sparkline);
    }
}

pub(crate) fn parse_sparkline_groups_from_worksheet_xml(
    content: &str,
) -> SheetResult<Vec<SparklineGroup>> {
    // Find extLst and scan for the sparklineGroups ext.
    let ext_lst_start = match content.find("<extLst") {
        Some(p) => p,
        None => return Ok(Vec::new()),
    };
    let ext_lst_end_rel = match content[ext_lst_start..].find("</extLst>") {
        Some(p) => p,
        None => return Ok(Vec::new()),
    };
    let ext_lst = &content[ext_lst_start..ext_lst_start + ext_lst_end_rel + "</extLst>".len()];

    const SPARKLINE_EXT_URI: &str = "{05C60535-1F16-4fd2-B633-F4F36F0B64E0}";

    let mut groups: Vec<SparklineGroup> = Vec::new();
    let mut pos = 0usize;

    while let Some(ext_pos_rel) = ext_lst[pos..].find("<ext ") {
        let ext_pos = pos + ext_pos_rel;
        let ext_end_rel = match ext_lst[ext_pos..].find("</ext>") {
            Some(p) => p,
            None => break,
        };
        let ext_end = ext_pos + ext_end_rel + "</ext>".len();
        let ext_xml = &ext_lst[ext_pos..ext_end];

        if extract_attribute(ext_xml, "uri")
            .as_deref()
            .is_some_and(|u| u.eq_ignore_ascii_case(SPARKLINE_EXT_URI))
            && let Some(spark_groups_xml) = extract_element(ext_xml, "x14:sparklineGroups")
        {
            groups.extend(parse_sparkline_groups(spark_groups_xml)?);
        }

        pos = ext_end;
    }

    if groups.len() >= 231 {
        return Err("sparklineGroups must contain fewer than 231 sparklineGroup elements".into());
    }

    Ok(groups)
}

pub(crate) fn write_sparkline_groups_ext(
    xml: &mut String,
    groups: &[SparklineGroup],
) -> SheetResult<()> {
    if groups.is_empty() {
        return Ok(());
    }

    xml.push_str(r#"<ext uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">"#);
    xml.push_str(
        r#"<x14:sparklineGroups xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">"#,
    );

    for group in groups {
        for sp in &group.sparklines {
            xml.push_str("<x14:sparklineGroup");

            // Excel omits the attribute for the default type (line).
            if group.sparkline_type != SparklineType::Line {
                xml.push_str(" type=\"");
                xml.push_str(group.sparkline_type.as_str());
                xml.push('"');
            }

            let has_attr = |name: &str| group.extra_attributes.iter().any(|(k, _)| k == name);

            if !has_attr("displayEmptyCellsAs") {
                xml.push_str(" displayEmptyCellsAs=\"");
                xml.push_str(group.options.display_empty_cells_as.as_str());
                xml.push('"');
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
            if group.options.min_axis_type != SparklineAxisMinMax::Individual
                && !has_attr("minAxisType")
            {
                xml.push_str(" minAxisType=\"");
                xml.push_str(group.options.min_axis_type.as_str());
                xml.push('"');
            }
            if group.options.max_axis_type != SparklineAxisMinMax::Individual
                && !has_attr("maxAxisType")
            {
                xml.push_str(" maxAxisType=\"");
                xml.push_str(group.options.max_axis_type.as_str());
                xml.push('"');
            }

            if let Some(min) = group.options.manual_min {
                if group.options.min_axis_type != SparklineAxisMinMax::Custom
                    && !has_attr("minAxisType")
                {
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
                if group.options.max_axis_type != SparklineAxisMinMax::Custom
                    && !has_attr("maxAxisType")
                {
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

            for (k, v) in &group.extra_attributes {
                xml.push(' ');
                xml.push_str(k);
                xml.push_str("=\"");
                xml.push_str(&escape_xml(v));
                xml.push('"');
            }

            xml.push('>');

            if let Some(ref c) = group.colors.series {
                write!(xml, r#"<x14:colorSeries rgb="{}"/>"#, escape_xml(&c.rgb))
                    .map_err(|e| format!("XML write error: {}", e))?;
            }
            if let Some(ref c) = group.colors.negative {
                write!(xml, r#"<x14:colorNegative rgb="{}"/>"#, escape_xml(&c.rgb))
                    .map_err(|e| format!("XML write error: {}", e))?;
            }
            if let Some(ref c) = group.colors.axis {
                write!(xml, r#"<x14:colorAxis rgb="{}"/>"#, escape_xml(&c.rgb))
                    .map_err(|e| format!("XML write error: {}", e))?;
            }
            if let Some(ref c) = group.colors.markers {
                write!(xml, r#"<x14:colorMarkers rgb="{}"/>"#, escape_xml(&c.rgb))
                    .map_err(|e| format!("XML write error: {}", e))?;
            }
            if let Some(ref c) = group.colors.first {
                write!(xml, r#"<x14:colorFirst rgb="{}"/>"#, escape_xml(&c.rgb))
                    .map_err(|e| format!("XML write error: {}", e))?;
            }
            if let Some(ref c) = group.colors.last {
                write!(xml, r#"<x14:colorLast rgb="{}"/>"#, escape_xml(&c.rgb))
                    .map_err(|e| format!("XML write error: {}", e))?;
            }
            if let Some(ref c) = group.colors.high {
                write!(xml, r#"<x14:colorHigh rgb="{}"/>"#, escape_xml(&c.rgb))
                    .map_err(|e| format!("XML write error: {}", e))?;
            }
            if let Some(ref c) = group.colors.low {
                write!(xml, r#"<x14:colorLow rgb="{}"/>"#, escape_xml(&c.rgb))
                    .map_err(|e| format!("XML write error: {}", e))?;
            }
            xml.push_str("<x14:sparklines>");

            xml.push_str("<x14:sparkline>");
            xml.push_str("<xm:f>");
            xml.push_str(&escape_xml(&sp.data_range));
            xml.push_str("</xm:f>");
            xml.push_str("<xm:sqref>");
            xml.push_str(&escape_xml(&sp.location));
            xml.push_str("</xm:sqref>");
            xml.push_str("</x14:sparkline>");

            xml.push_str("</x14:sparklines>");
            xml.push_str("</x14:sparklineGroup>");
        }
    }

    xml.push_str("</x14:sparklineGroups>");
    xml.push_str("</ext>");
    Ok(())
}

fn parse_sparkline_groups(content: &str) -> SheetResult<Vec<SparklineGroup>> {
    let mut groups: Vec<SparklineGroup> = Vec::new();

    let mut pos = 0usize;
    while let Some(group_pos_rel) = content[pos..].find("<x14:sparklineGroup") {
        let group_pos = pos + group_pos_rel;
        let open_end_rel = match content[group_pos..].find('>') {
            Some(p) => p,
            None => break,
        };
        let open_end = group_pos + open_end_rel + 1;
        let open_tag = &content[group_pos..open_end];

        let close_rel = match content[open_end..].find("</x14:sparklineGroup>") {
            Some(p) => p,
            None => break,
        };
        let group_end = open_end + close_rel + "</x14:sparklineGroup>".len();
        let group_inner = &content[open_end..open_end + close_rel];

        let sparkline_type = extract_attribute(open_tag, "type")
            .as_deref()
            .and_then(SparklineType::parse)
            .unwrap_or(SparklineType::Line);

        let mut group = SparklineGroup::new(sparkline_type);

        if let Some(v) = extract_attribute(open_tag, "displayEmptyCellsAs")
            .as_deref()
            .and_then(SparklineDisplayEmptyCellsAs::parse)
        {
            group.options.display_empty_cells_as = v;
        }
        if let Some(v) = extract_attribute(open_tag, "displayHidden").and_then(parse_bool_attr) {
            group.options.display_hidden = v;
        }
        if let Some(v) = extract_attribute(open_tag, "displayXAxis").and_then(parse_bool_attr) {
            group.options.display_x_axis = v;
        }
        if let Some(v) = extract_attribute(open_tag, "markers").and_then(parse_bool_attr) {
            group.options.markers = v;
        }
        if let Some(v) = extract_attribute(open_tag, "high").and_then(parse_bool_attr) {
            group.options.high = v;
        }
        if let Some(v) = extract_attribute(open_tag, "low").and_then(parse_bool_attr) {
            group.options.low = v;
        }
        if let Some(v) = extract_attribute(open_tag, "first").and_then(parse_bool_attr) {
            group.options.first = v;
        }
        if let Some(v) = extract_attribute(open_tag, "last").and_then(parse_bool_attr) {
            group.options.last = v;
        }
        if let Some(v) = extract_attribute(open_tag, "negative").and_then(parse_bool_attr) {
            group.options.negative = v;
        }
        if let Some(v) = extract_attribute(open_tag, "rightToLeft").and_then(parse_bool_attr) {
            group.options.right_to_left = v;
        }
        if let Some(v) = extract_attribute(open_tag, "minAxisType")
            .as_deref()
            .and_then(SparklineAxisMinMax::parse)
        {
            group.options.min_axis_type = v;
        }
        if let Some(v) = extract_attribute(open_tag, "maxAxisType")
            .as_deref()
            .and_then(SparklineAxisMinMax::parse)
        {
            group.options.max_axis_type = v;
        }
        if let Some(v) = extract_attribute(open_tag, "manualMin").and_then(parse_f64_attr) {
            group.options.manual_min = Some(v);
        }
        if let Some(v) = extract_attribute(open_tag, "manualMax").and_then(parse_f64_attr) {
            group.options.manual_max = Some(v);
        }
        if let Some(v) = extract_attribute(open_tag, "lineWeight").and_then(parse_f64_attr) {
            group.options.line_weight = Some(v);
        }

        group.extra_attributes = extract_extra_attributes(
            open_tag,
            &[
                "type",
                "displayEmptyCellsAs",
                "displayHidden",
                "displayXAxis",
                "markers",
                "high",
                "low",
                "first",
                "last",
                "negative",
                "rightToLeft",
                "minAxisType",
                "maxAxisType",
                "manualMin",
                "manualMax",
                "lineWeight",
            ],
        );

        if let Some(rgb) = extract_child_rgb(group_inner, "x14:colorSeries") {
            group.colors.series = Some(SparklineColor::new(rgb));
        }
        if let Some(rgb) = extract_child_rgb(group_inner, "x14:colorNegative") {
            group.colors.negative = Some(SparklineColor::new(rgb));
        }
        if let Some(rgb) = extract_child_rgb(group_inner, "x14:colorAxis") {
            group.colors.axis = Some(SparklineColor::new(rgb));
        }
        if let Some(rgb) = extract_child_rgb(group_inner, "x14:colorMarkers") {
            group.colors.markers = Some(SparklineColor::new(rgb));
        }
        if let Some(rgb) = extract_child_rgb(group_inner, "x14:colorFirst") {
            group.colors.first = Some(SparklineColor::new(rgb));
        }
        if let Some(rgb) = extract_child_rgb(group_inner, "x14:colorLast") {
            group.colors.last = Some(SparklineColor::new(rgb));
        }
        if let Some(rgb) = extract_child_rgb(group_inner, "x14:colorHigh") {
            group.colors.high = Some(SparklineColor::new(rgb));
        }
        if let Some(rgb) = extract_child_rgb(group_inner, "x14:colorLow") {
            group.colors.low = Some(SparklineColor::new(rgb));
        }

        if let Some(sparklines_xml) = extract_element(group_inner, "x14:sparklines") {
            group.sparklines = parse_sparklines(sparklines_xml);
        }

        groups.push(group);
        pos = group_end;
    }

    Ok(groups)
}

fn parse_sparklines(content: &str) -> Vec<Sparkline> {
    let mut sparklines: Vec<Sparkline> = Vec::new();

    let mut pos = 0usize;
    while let Some(sp_pos_rel) = content[pos..].find("<x14:sparkline") {
        let sp_pos = pos + sp_pos_rel;
        let open_end_rel = match content[sp_pos..].find('>') {
            Some(p) => p,
            None => break,
        };
        let open_end = sp_pos + open_end_rel + 1;

        let close_rel = match content[open_end..].find("</x14:sparkline>") {
            Some(p) => p,
            None => break,
        };
        let sp_end = open_end + close_rel + "</x14:sparkline>".len();
        let sp_inner = &content[open_end..open_end + close_rel];

        let data_range = extract_text_element(sp_inner, "xm:f")
            .or_else(|| extract_text_element(sp_inner, "f"))
            .unwrap_or_default();
        let location = extract_text_element(sp_inner, "xm:sqref")
            .or_else(|| extract_text_element(sp_inner, "sqref"))
            .or_else(|| extract_text_element(sp_inner, "xm:ref"))
            .or_else(|| extract_text_element(sp_inner, "ref"))
            .unwrap_or_default();

        if !data_range.is_empty() && !location.is_empty() {
            sparklines.push(Sparkline {
                data_range,
                location,
            });
        }

        pos = sp_end;
    }

    sparklines
}

fn extract_element<'a>(xml: &'a str, tag: &str) -> Option<&'a str> {
    let open_pat = format!("<{}", tag);
    let start = xml.find(&open_pat)?;
    let after_open = &xml[start..];
    let gt_rel = after_open.find('>')?;
    let inner_start = start + gt_rel + 1;

    let close_pat = format!("</{}>", tag);
    let close_rel = xml[inner_start..].find(&close_pat)?;
    let inner_end = inner_start + close_rel;
    Some(&xml[inner_start..inner_end])
}

fn extract_text_element(xml: &str, tag: &str) -> Option<String> {
    let open = format!("<{}>", tag);
    let start = xml.find(&open)? + open.len();
    let close = format!("</{}>", tag);
    let end_rel = xml[start..].find(&close)?;
    Some(unescape_xml(&xml[start..start + end_rel]))
}

fn extract_attribute(tag: &str, attr: &str) -> Option<String> {
    let pat = format!("{}=\"", attr);
    let start = tag.find(&pat)? + pat.len();
    let tail = &tag[start..];
    let end = tail.find('"')?;
    Some(tail[..end].to_string())
}

fn extract_extra_attributes(tag: &str, known: &[&str]) -> Vec<(String, String)> {
    // Very small attribute scanner: only handles quoted attributes and ignores namespaces.
    let mut out: Vec<(String, String)> = Vec::new();
    let mut i = 0usize;
    let bytes = tag.as_bytes();

    while i < bytes.len() {
        // Find '='
        let eq_rel = match memchr::memchr(b'=', &bytes[i..]) {
            Some(p) => p,
            None => break,
        };
        let eq = i + eq_rel;
        // Find attribute name start by scanning backwards to whitespace or '<'
        let name_end = eq;
        let mut name_start = name_end;
        while name_start > 0 {
            let b = bytes[name_start - 1];
            if b == b' ' || b == b'\n' || b == b'\r' || b == b'\t' || b == b'<' {
                break;
            }
            name_start -= 1;
        }
        let name = &tag[name_start..name_end];
        if known.contains(&name) {
            i = eq + 1;
            continue;
        }
        // Expect "
        if bytes.get(eq + 1) != Some(&b'"') {
            i = eq + 1;
            continue;
        }
        let val_start = eq + 2;
        let quote_rel = match memchr::memchr(b'"', &bytes[val_start..]) {
            Some(p) => p,
            None => break,
        };
        let val_end = val_start + quote_rel;
        let val = &tag[val_start..val_end];
        if !name.is_empty() {
            out.push((name.to_string(), unescape_xml(val)));
        }
        i = val_end + 1;
    }

    out
}

fn parse_bool_attr(v: String) -> Option<bool> {
    match v.as_str() {
        "1" | "true" | "TRUE" => Some(true),
        "0" | "false" | "FALSE" => Some(false),
        _ => None,
    }
}

fn parse_f64_attr(v: String) -> Option<f64> {
    v.parse().ok()
}

fn extract_child_rgb(xml: &str, tag: &str) -> Option<String> {
    let open_pat = format!("<{}", tag);
    let start = xml.find(&open_pat)?;
    let after_open = &xml[start..];
    let gt_rel = after_open.find('>')?;
    let open_tag = &after_open[..gt_rel + 1];
    extract_attribute(open_tag, "rgb")
}
