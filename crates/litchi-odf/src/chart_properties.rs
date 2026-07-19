//! Complete typed ODF `style:chart-properties` support.

use crate::{FlatOpenDocument, OpenDocumentPackage};
use litchi_core::{Error, Result, xml::escape_xml};
use quick_xml::{
    XmlVersion,
    events::{BytesStart, Event},
    name::{QName, ResolveResult},
    reader::NsReader,
};
use std::collections::HashSet;

const OFFICE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const STYLE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:style:1.0";
const CHART: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:chart:1.0";
const TEXT: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const XLINK: &[u8] = b"http://www.w3.org/1999/xlink";
const OFFICE_NS: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const STYLE_NS: &str = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
const CHART_NS: &str = "urn:oasis:names:tc:opendocument:xmlns:chart:1.0";
const TEXT_NS: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const XLINK_NS: &str = "http://www.w3.org/1999/xlink";
const MAX_XML: usize = 32 * 1024 * 1024;
const MAX_VALUE: usize = 1024 * 1024;
const MAX_ATTRIBUTES: usize = 96;
const MAX_DEPTH: usize = 128;
const MAX_STYLES: usize = 65_536;
const MAX_TOTAL: usize = 16 * 1024 * 1024;
const MAX_EVENTS: usize = 1_000_000;

fn bad(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}
fn safe(value: &str, name: &str, empty: bool) -> Result<()> {
    if (!empty && value.is_empty())
        || value.len() > MAX_VALUE
        || value.chars().any(
            |c| matches!(c,'\0'..='\u{8}'|'\u{b}'|'\u{c}'|'\u{e}'..='\u{1f}'|'\u{fffe}'|'\u{ffff}'),
        )
    {
        return Err(bad(format!("invalid {name}")));
    }
    Ok(())
}
fn ncname(value: &str, name: &str) -> Result<()> {
    safe(value, name, false)?;
    let mut chars = value.chars();
    let first = chars.next().ok_or_else(|| bad(format!("invalid {name}")))?;
    if !(first == '_' || first.is_alphabetic())
        || chars.any(|c| !(c == '_' || c == '-' || c == '.' || c.is_alphanumeric()))
    {
        return Err(bad(format!("invalid {name}")));
    }
    Ok(())
}
fn decimal(value: &str, signed: bool) -> bool {
    let value = if signed {
        value.strip_prefix(['+', '-']).unwrap_or(value)
    } else {
        value
    };
    if value.is_empty() {
        return false;
    }
    let mut parts = value.split('.');
    let left = parts.next().unwrap_or_default();
    let right = parts.next();
    if parts.next().is_some() {
        return false;
    }
    match right {
        None => !left.is_empty() && left.bytes().all(|b| b.is_ascii_digit()),
        Some(right) => {
            (!left.is_empty() || !right.is_empty())
                && left.bytes().all(|b| b.is_ascii_digit())
                && right.bytes().all(|b| b.is_ascii_digit())
        },
    }
}
fn integer(value: &str) -> bool {
    let value = value.strip_prefix(['+', '-']).unwrap_or(value);
    !value.is_empty() && value.bytes().all(|b| b.is_ascii_digit())
}
fn positive_integer(value: &str) -> bool {
    let value = value.strip_prefix('+').unwrap_or(value);
    !value.is_empty()
        && value.bytes().all(|b| b.is_ascii_digit())
        && value.bytes().any(|b| b != b'0')
}
fn nonnegative_integer(value: &str) -> bool {
    let value = value.strip_prefix('+').unwrap_or(value);
    !value.is_empty() && value.bytes().all(|b| b.is_ascii_digit())
}
fn percent(value: &str) -> bool {
    value.strip_suffix('%').is_some_and(|v| {
        decimal(v, false) || v.strip_prefix('-').is_some_and(|v| decimal(v, false))
    })
}
fn nonnegative_length(value: &str) -> bool {
    ["cm", "mm", "in", "pt", "pc", "px"].iter().any(|unit| {
        value
            .strip_suffix(unit)
            .is_some_and(|number| decimal(number, false))
    })
}
fn double(value: &str) -> bool {
    if matches!(value, "INF" | "-INF" | "NaN") {
        return true;
    }
    let value = value.strip_prefix(['+', '-']).unwrap_or(value);
    let (mut base, mut exponent) = (value, None);
    if let Some(index) = value.find(['e', 'E']) {
        base = &value[..index];
        exponent = Some(&value[index + 1..])
    }
    decimal(base, false) && exponent.is_none_or(integer)
}

macro_rules! lexical {
    ($name:ident,$validator:expr,$label:literal,$empty:expr) => {
        #[derive(Debug, Clone, PartialEq, Eq)]
        pub struct $name(String);
        impl $name {
            pub fn new(value: impl Into<String>) -> Result<Self> {
                let value = value.into();
                safe(&value, $label, $empty)?;
                if !($validator)(&value) {
                    return Err(bad(concat!("invalid ", $label)));
                }
                Ok(Self(value))
            }
            pub fn as_str(&self) -> &str {
                &self.0
            }
        }
    };
}
lexical!(
    ChartNonNegativeLength,
    nonnegative_length,
    "chart non-negative length",
    false
);
lexical!(ChartInteger, integer, "chart integer", false);
lexical!(
    ChartPositiveInteger,
    positive_integer,
    "chart positive integer",
    false
);
lexical!(
    ChartNonNegativeInteger,
    nonnegative_integer,
    "chart non-negative integer",
    false
);
lexical!(ChartPercent, percent, "chart percentage", false);
lexical!(ChartDouble, double, "chart double", false);
lexical!(ChartAngle, |_: &str| true, "chart angle", true);

macro_rules! keyword_enum{($name:ident{$($variant:ident=>$value:literal),+$(,)?})=>{#[derive(Debug,Clone,Copy,PartialEq,Eq)]pub enum $name{$($variant),+}impl $name{fn parse(value:&str)->Result<Self>{match value{$($value=>Ok(Self::$variant),)+_=>Err(bad(concat!("invalid ",stringify!($name))))}}fn xml(self)->&'static str{match self{$(Self::$variant=>$value),+}}}};}
keyword_enum!(ChartSymbolType{None=>"none",Automatic=>"automatic",NamedSymbol=>"named-symbol",Image=>"image"});
keyword_enum!(ChartSymbolName{Square=>"square",Diamond=>"diamond",ArrowDown=>"arrow-down",ArrowUp=>"arrow-up",ArrowRight=>"arrow-right",ArrowLeft=>"arrow-left",BowTie=>"bow-tie",Hourglass=>"hourglass",Circle=>"circle",Star=>"star",X=>"x",Plus=>"plus",Asterisk=>"asterisk",HorizontalBar=>"horizontal-bar",VerticalBar=>"vertical-bar"});
keyword_enum!(ChartInterpolation{None=>"none",CubicSpline=>"cubic-spline",BSpline=>"b-spline"});
keyword_enum!(ChartSolidType{Cuboid=>"cuboid",Cylinder=>"cylinder",Cone=>"cone",Pyramid=>"pyramid"});
keyword_enum!(ChartEmptyCellTreatment{UseZero=>"use-zero",LeaveGap=>"leave-gap",Ignore=>"ignore"});
keyword_enum!(ChartLabelArrangement{SideBySide=>"side-by-side",StaggerEven=>"stagger-even",StaggerOdd=>"stagger-odd"});
keyword_enum!(ChartDirection{Ltr=>"ltr",Ttb=>"ttb"});
keyword_enum!(ChartDataLabelNumber{None=>"none",Value=>"value",Percentage=>"percentage",ValueAndPercentage=>"value-and-percentage"});
keyword_enum!(ChartLabelPosition{AvoidOverlap=>"avoid-overlap",Center=>"center",Top=>"top",TopRight=>"top-right",Right=>"right",BottomRight=>"bottom-right",Bottom=>"bottom",BottomLeft=>"bottom-left",Left=>"left",TopLeft=>"top-left",Inside=>"inside",Outside=>"outside",NearOrigin=>"near-origin"});
keyword_enum!(ChartErrorCategory{None=>"none",Variance=>"variance",StandardDeviation=>"standard-deviation",Percentage=>"percentage",ErrorMargin=>"error-margin",Constant=>"constant",StandardError=>"standard-error",CellRange=>"cell-range"});
keyword_enum!(ChartSeriesSource{Columns=>"columns",Rows=>"rows"});
keyword_enum!(ChartRegressionType{None=>"none",Linear=>"linear",Logarithmic=>"logarithmic",Exponential=>"exponential",Power=>"power"});
keyword_enum!(ChartAxisLabelPosition{NearAxis=>"near-axis",NearAxisOtherSide=>"near-axis-other-side",OutsideStart=>"outside-start",OutsideEnd=>"outside-end"});
keyword_enum!(ChartTickMarkPosition{AtLabels=>"at-labels",AtAxis=>"at-axis",AtLabelsAndAxis=>"at-labels-and-axis"});

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ChartAxisPosition {
    Start,
    End,
    Value(ChartDouble),
}
impl ChartAxisPosition {
    fn parse(value: &str) -> Result<Self> {
        match value {
            "start" => Ok(Self::Start),
            "end" => Ok(Self::End),
            _ => Ok(Self::Value(ChartDouble::new(value)?)),
        }
    }
    fn xml(&self) -> &str {
        match self {
            Self::Start => "start",
            Self::End => "end",
            Self::Value(value) => value.as_str(),
        }
    }
}

/// Inert image marker reference. The URI is never fetched.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChartSymbolImage {
    pub href: String,
}
impl ChartSymbolImage {
    pub fn new(href: impl Into<String>) -> Result<Self> {
        let value = Self { href: href.into() };
        value.validate()?;
        Ok(value)
    }
    pub fn validate(&self) -> Result<()> {
        safe(&self.href, "chart:symbol-image xlink:href", true)
    }
    pub fn to_xml_fragment(&self) -> Result<String> {
        self.validate()?;
        Ok(format!(
            r#"<chart:symbol-image xmlns:chart="{CHART_NS}" xmlns:xlink="{XLINK_NS}" xlink:href="{}"/>"#,
            escape_xml(&self.href)
        ))
    }
}

/// Inert, bounded `text:p` XML inside `chart:label-separator`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChartLabelSeparator {
    paragraph_xml: String,
}
impl ChartLabelSeparator {
    pub fn new_text(text: &str) -> Result<Self> {
        safe(text, "chart label separator text", true)?;
        Self::from_paragraph_xml(format!(
            r#"<text:p xmlns:text="{TEXT_NS}">{}</text:p>"#,
            escape_xml(text)
        ))
    }
    pub fn from_paragraph_xml(paragraph_xml: impl Into<String>) -> Result<Self> {
        let paragraph_xml = paragraph_xml.into();
        validate_paragraph(&paragraph_xml)?;
        Ok(Self { paragraph_xml })
    }
    pub fn paragraph_xml(&self) -> &str {
        &self.paragraph_xml
    }
    pub fn to_xml_fragment(&self) -> Result<String> {
        validate_paragraph(&self.paragraph_xml)?;
        Ok(format!(
            r#"<chart:label-separator xmlns:chart="{CHART_NS}">{} </chart:label-separator>"#,
            self.paragraph_xml
        )
        .replace("</text:p> </chart:", "</text:p></chart:"))
    }
}
fn validate_paragraph(xml: &str) -> Result<()> {
    if xml.len() > MAX_VALUE {
        return Err(bad("chart label paragraph is too large"));
    }
    let wrapped = format!(
        r#"<wrapper xmlns:text="{TEXT_NS}" xmlns:chart="{CHART_NS}" xmlns:style="{STYLE_NS}" xmlns:xlink="{XLINK_NS}">{xml}</wrapper>"#
    );
    let mut reader = NsReader::from_reader(wrapped.as_bytes());
    let mut depth = 0usize;
    let mut paragraph = false;
    let mut events = 0;
    loop {
        events += 1;
        if events > MAX_EVENTS {
            return Err(bad("chart label paragraph has too many events"));
        }
        match reader.read_event() {
            Ok(Event::Start(start)) => {
                if depth >= MAX_DEPTH {
                    return Err(bad("chart label paragraph is too deep"));
                }
                let current = element(&reader, start.name());
                if depth == 0 {
                    if current.0 != Ns::Other || current.1 != b"wrapper" {
                        return Err(bad("invalid chart label wrapper"));
                    }
                } else if depth == 1 {
                    if paragraph || current.0 != Ns::Text || current.1 != b"p" {
                        return Err(bad("chart label separator requires one text:p"));
                    }
                    paragraph = true
                }
                depth += 1
            },
            Ok(Event::Empty(start)) => {
                let current = element(&reader, start.name());
                if depth == 1 {
                    if paragraph || current.0 != Ns::Text || current.1 != b"p" {
                        return Err(bad("chart label separator requires one text:p"));
                    }
                    paragraph = true
                }
            },
            Ok(Event::Text(text)) => {
                let bytes: &[u8] = text.as_ref();
                if depth == 1 && !bytes.iter().all(u8::is_ascii_whitespace) {
                    return Err(bad("chart label separator allows only one text:p"));
                }
            },
            Ok(Event::CData(text)) => {
                let bytes: &[u8] = text.as_ref();
                if depth == 1 && !bytes.iter().all(u8::is_ascii_whitespace) {
                    return Err(bad("chart label separator allows only one text:p"));
                }
            },
            Ok(Event::End(_)) => {
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| bad("invalid chart label paragraph"))?
            },
            Ok(Event::Decl(_)) | Ok(Event::DocType(_)) | Ok(Event::PI(_)) => {
                return Err(bad(
                    "declarations, DTDs, and processing instructions are not allowed in chart labels",
                ));
            },
            Ok(Event::Eof) => break,
            Ok(_) => {},
            Err(error) => return Err(bad(format!("invalid chart label paragraph: {error}"))),
        }
    }
    if !paragraph || depth != 0 {
        return Err(bad("truncated chart label paragraph"));
    }
    Ok(())
}

/// All 66 attributes and both children allowed by `style:chart-properties`.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct ChartStyleProperties {
    pub scale_text: Option<bool>,
    pub three_dimensional: Option<bool>,
    pub deep: Option<bool>,
    pub right_angled_axes: Option<bool>,
    pub symbol_type: Option<ChartSymbolType>,
    pub symbol_name: Option<ChartSymbolName>,
    pub symbol_image: Option<ChartSymbolImage>,
    pub symbol_width: Option<ChartNonNegativeLength>,
    pub symbol_height: Option<ChartNonNegativeLength>,
    pub sort_by_x_values: Option<bool>,
    pub vertical: Option<bool>,
    pub connect_bars: Option<bool>,
    pub gap_width: Option<ChartInteger>,
    pub overlap: Option<ChartInteger>,
    pub group_bars_per_axis: Option<bool>,
    pub japanese_candle_stick: Option<bool>,
    pub interpolation: Option<ChartInterpolation>,
    pub spline_order: Option<ChartPositiveInteger>,
    pub spline_resolution: Option<ChartPositiveInteger>,
    pub pie_offset: Option<ChartNonNegativeInteger>,
    pub angle_offset: Option<ChartAngle>,
    pub hole_size: Option<ChartPercent>,
    pub lines: Option<bool>,
    pub solid_type: Option<ChartSolidType>,
    pub stacked: Option<bool>,
    pub percentage: Option<bool>,
    pub treat_empty_cells: Option<ChartEmptyCellTreatment>,
    pub link_data_style_to_source: Option<bool>,
    pub logarithmic: Option<bool>,
    pub maximum: Option<ChartDouble>,
    pub minimum: Option<ChartDouble>,
    pub origin: Option<ChartDouble>,
    pub interval_major: Option<ChartDouble>,
    pub interval_minor_divisor: Option<ChartPositiveInteger>,
    pub tick_marks_major_inner: Option<bool>,
    pub tick_marks_major_outer: Option<bool>,
    pub tick_marks_minor_inner: Option<bool>,
    pub tick_marks_minor_outer: Option<bool>,
    pub reverse_direction: Option<bool>,
    pub display_label: Option<bool>,
    pub text_overlap: Option<bool>,
    pub line_break: Option<bool>,
    pub label_arrangement: Option<ChartLabelArrangement>,
    pub direction: Option<ChartDirection>,
    pub rotation_angle: Option<ChartAngle>,
    pub data_label_number: Option<ChartDataLabelNumber>,
    pub data_label_text: Option<bool>,
    pub data_label_symbol: Option<bool>,
    pub label_separator: Option<ChartLabelSeparator>,
    pub label_position: Option<ChartLabelPosition>,
    pub label_position_negative: Option<ChartLabelPosition>,
    pub visible: Option<bool>,
    pub auto_position: Option<bool>,
    pub auto_size: Option<bool>,
    pub mean_value: Option<bool>,
    pub error_category: Option<ChartErrorCategory>,
    pub error_percentage: Option<ChartDouble>,
    pub error_margin: Option<ChartDouble>,
    pub error_lower_limit: Option<ChartDouble>,
    pub error_upper_limit: Option<ChartDouble>,
    pub error_upper_indicator: Option<bool>,
    pub error_lower_indicator: Option<bool>,
    pub series_source: Option<ChartSeriesSource>,
    pub regression_type: Option<ChartRegressionType>,
    pub axis_position: Option<ChartAxisPosition>,
    pub axis_label_position: Option<ChartAxisLabelPosition>,
    pub tick_mark_position: Option<ChartTickMarkPosition>,
    pub include_hidden_cells: Option<bool>,
}
impl ChartStyleProperties {
    pub fn validate(&self) -> Result<()> {
        match self.symbol_type {
            Some(ChartSymbolType::NamedSymbol)
                if self.symbol_name.is_some() && self.symbol_image.is_none() => {},
            Some(ChartSymbolType::Image)
                if self.symbol_image.is_some() && self.symbol_name.is_none() => {},
            Some(ChartSymbolType::None | ChartSymbolType::Automatic)
                if self.symbol_name.is_none() && self.symbol_image.is_none() => {},
            None if self.symbol_name.is_none() && self.symbol_image.is_none() => {},
            _ => {
                return Err(bad(
                    "chart symbol type/name/image combination violates ODF grammar",
                ));
            },
        }
        if let Some(image) = &self.symbol_image {
            image.validate()?
        }
        if let Some(label) = &self.label_separator {
            validate_paragraph(label.paragraph_xml())?
        }
        Ok(())
    }
    pub fn from_xml_fragment(fragment: &str) -> Result<Self> {
        let xml = format!(
            r#"<office:document xmlns:office="{OFFICE_NS}" xmlns:style="{STYLE_NS}"><office:styles><style:style style:name="fragment" style:family="chart">{fragment}</style:style></office:styles></office:document>"#
        );
        let mut set = parse_chart_style_properties(&xml)?;
        set.styles
            .pop()
            .and_then(|style| style.properties)
            .ok_or_else(|| bad("fragment does not contain style:chart-properties"))
    }
    pub fn to_xml_fragment(&self) -> Result<String> {
        self.validate()?;
        let mut xml = format!(
            r#"<style:chart-properties xmlns:style="{STYLE_NS}" xmlns:chart="{CHART_NS}" xmlns:text="{TEXT_NS}" xmlns:xlink="{XLINK_NS}""#
        );
        macro_rules! a {
            ($field:expr,$name:literal,$render:expr) => {
                if let Some(value) = $field {
                    let rendered = ($render)(value);
                    xml.push_str(&format!(concat!(" ", $name, "=\"{}\""), rendered))
                }
            };
        }
        macro_rules! b {
            ($field:expr,$name:literal) => {
                a!($field, $name, |value: bool| if value {
                    "true"
                } else {
                    "false"
                })
            };
        }
        b!(self.scale_text, "chart:scale-text");
        b!(self.three_dimensional, "chart:three-dimensional");
        b!(self.deep, "chart:deep");
        b!(self.right_angled_axes, "chart:right-angled-axes");
        a!(
            self.symbol_type,
            "chart:symbol-type",
            |v: ChartSymbolType| v.xml()
        );
        a!(
            self.symbol_name,
            "chart:symbol-name",
            |v: ChartSymbolName| v.xml()
        );
        a!(
            self.symbol_width.as_ref(),
            "chart:symbol-width",
            |v: &ChartNonNegativeLength| v.as_str().to_owned()
        );
        a!(
            self.symbol_height.as_ref(),
            "chart:symbol-height",
            |v: &ChartNonNegativeLength| v.as_str().to_owned()
        );
        b!(self.sort_by_x_values, "chart:sort-by-x-values");
        b!(self.vertical, "chart:vertical");
        b!(self.connect_bars, "chart:connect-bars");
        a!(
            self.gap_width.as_ref(),
            "chart:gap-width",
            |v: &ChartInteger| v.as_str().to_owned()
        );
        a!(
            self.overlap.as_ref(),
            "chart:overlap",
            |v: &ChartInteger| v.as_str().to_owned()
        );
        b!(self.group_bars_per_axis, "chart:group-bars-per-axis");
        b!(self.japanese_candle_stick, "chart:japanese-candle-stick");
        a!(
            self.interpolation,
            "chart:interpolation",
            |v: ChartInterpolation| v.xml()
        );
        a!(
            self.spline_order.as_ref(),
            "chart:spline-order",
            |v: &ChartPositiveInteger| v.as_str().to_owned()
        );
        a!(
            self.spline_resolution.as_ref(),
            "chart:spline-resolution",
            |v: &ChartPositiveInteger| v.as_str().to_owned()
        );
        a!(
            self.pie_offset.as_ref(),
            "chart:pie-offset",
            |v: &ChartNonNegativeInteger| v.as_str().to_owned()
        );
        a!(
            self.angle_offset.as_ref(),
            "chart:angle-offset",
            |v: &ChartAngle| escape_xml(v.as_str())
        );
        a!(
            self.hole_size.as_ref(),
            "chart:hole-size",
            |v: &ChartPercent| v.as_str().to_owned()
        );
        b!(self.lines, "chart:lines");
        a!(self.solid_type, "chart:solid-type", |v: ChartSolidType| v
            .xml());
        b!(self.stacked, "chart:stacked");
        b!(self.percentage, "chart:percentage");
        a!(
            self.treat_empty_cells,
            "chart:treat-empty-cells",
            |v: ChartEmptyCellTreatment| v.xml()
        );
        b!(
            self.link_data_style_to_source,
            "chart:link-data-style-to-source"
        );
        b!(self.logarithmic, "chart:logarithmic");
        a!(self.maximum.as_ref(), "chart:maximum", |v: &ChartDouble| v
            .as_str()
            .to_owned());
        a!(self.minimum.as_ref(), "chart:minimum", |v: &ChartDouble| v
            .as_str()
            .to_owned());
        a!(self.origin.as_ref(), "chart:origin", |v: &ChartDouble| v
            .as_str()
            .to_owned());
        a!(
            self.interval_major.as_ref(),
            "chart:interval-major",
            |v: &ChartDouble| v.as_str().to_owned()
        );
        a!(
            self.interval_minor_divisor.as_ref(),
            "chart:interval-minor-divisor",
            |v: &ChartPositiveInteger| v.as_str().to_owned()
        );
        b!(self.tick_marks_major_inner, "chart:tick-marks-major-inner");
        b!(self.tick_marks_major_outer, "chart:tick-marks-major-outer");
        b!(self.tick_marks_minor_inner, "chart:tick-marks-minor-inner");
        b!(self.tick_marks_minor_outer, "chart:tick-marks-minor-outer");
        b!(self.reverse_direction, "chart:reverse-direction");
        b!(self.display_label, "chart:display-label");
        b!(self.text_overlap, "chart:text-overlap");
        b!(self.line_break, "text:line-break");
        a!(
            self.label_arrangement,
            "chart:label-arrangement",
            |v: ChartLabelArrangement| v.xml()
        );
        a!(self.direction, "style:direction", |v: ChartDirection| v
            .xml());
        a!(
            self.rotation_angle.as_ref(),
            "style:rotation-angle",
            |v: &ChartAngle| escape_xml(v.as_str())
        );
        a!(
            self.data_label_number,
            "chart:data-label-number",
            |v: ChartDataLabelNumber| v.xml()
        );
        b!(self.data_label_text, "chart:data-label-text");
        b!(self.data_label_symbol, "chart:data-label-symbol");
        a!(
            self.label_position,
            "chart:label-position",
            |v: ChartLabelPosition| v.xml()
        );
        a!(
            self.label_position_negative,
            "chart:label-position-negative",
            |v: ChartLabelPosition| v.xml()
        );
        b!(self.visible, "chart:visible");
        b!(self.auto_position, "chart:auto-position");
        b!(self.auto_size, "chart:auto-size");
        b!(self.mean_value, "chart:mean-value");
        a!(
            self.error_category,
            "chart:error-category",
            |v: ChartErrorCategory| v.xml()
        );
        a!(
            self.error_percentage.as_ref(),
            "chart:error-percentage",
            |v: &ChartDouble| v.as_str().to_owned()
        );
        a!(
            self.error_margin.as_ref(),
            "chart:error-margin",
            |v: &ChartDouble| v.as_str().to_owned()
        );
        a!(
            self.error_lower_limit.as_ref(),
            "chart:error-lower-limit",
            |v: &ChartDouble| v.as_str().to_owned()
        );
        a!(
            self.error_upper_limit.as_ref(),
            "chart:error-upper-limit",
            |v: &ChartDouble| v.as_str().to_owned()
        );
        b!(self.error_upper_indicator, "chart:error-upper-indicator");
        b!(self.error_lower_indicator, "chart:error-lower-indicator");
        a!(
            self.series_source,
            "chart:series-source",
            |v: ChartSeriesSource| v.xml()
        );
        a!(
            self.regression_type,
            "chart:regression-type",
            |v: ChartRegressionType| v.xml()
        );
        a!(
            self.axis_position.as_ref(),
            "chart:axis-position",
            |v: &ChartAxisPosition| v.xml().to_owned()
        );
        a!(
            self.axis_label_position,
            "chart:axis-label-position",
            |v: ChartAxisLabelPosition| v.xml()
        );
        a!(
            self.tick_mark_position,
            "chart:tick-mark-position",
            |v: ChartTickMarkPosition| v.xml()
        );
        b!(self.include_hidden_cells, "chart:include-hidden-cells");
        if self.symbol_image.is_some() || self.label_separator.is_some() {
            xml.push('>');
            if let Some(image) = &self.symbol_image {
                xml.push_str(&image.to_xml_fragment()?)
            }
            if let Some(label) = &self.label_separator {
                xml.push_str(&label.to_xml_fragment()?)
            }
            xml.push_str("</style:chart-properties>")
        } else {
            xml.push_str("/>")
        }
        Ok(xml)
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChartStyleRecord {
    pub name: Option<String>,
    pub parent_style_name: Option<String>,
    pub is_default_style: bool,
    pub properties: Option<ChartStyleProperties>,
}
impl ChartStyleRecord {
    pub fn named(
        name: impl Into<String>,
        properties: Option<ChartStyleProperties>,
    ) -> Result<Self> {
        let value = Self {
            name: Some(name.into()),
            parent_style_name: None,
            is_default_style: false,
            properties,
        };
        value.validate()?;
        Ok(value)
    }
    pub fn default_style(properties: Option<ChartStyleProperties>) -> Self {
        Self {
            name: None,
            parent_style_name: None,
            is_default_style: true,
            properties,
        }
    }
    pub fn validate(&self) -> Result<()> {
        match (&self.name, self.is_default_style) {
            (Some(value), false) => ncname(value, "chart style name")?,
            (None, true) => {},
            _ => return Err(bad("invalid chart style identity")),
        }
        if let Some(value) = &self.parent_style_name {
            if self.is_default_style {
                return Err(bad("default chart style cannot have a parent"));
            }
            ncname(value, "parent chart style name")?
        }
        if let Some(value) = &self.properties {
            value.validate()?
        }
        Ok(())
    }
    pub fn to_xml_fragment(&self) -> Result<String> {
        self.validate()?;
        let tag = if self.is_default_style {
            "default-style"
        } else {
            "style"
        };
        let mut xml = format!(r#"<style:{tag} xmlns:style="{STYLE_NS}" style:family="chart""#);
        if let Some(value) = &self.name {
            xml.push_str(&format!(r#" style:name="{}""#, escape_xml(value)))
        }
        if let Some(value) = &self.parent_style_name {
            xml.push_str(&format!(
                r#" style:parent-style-name="{}""#,
                escape_xml(value)
            ))
        }
        if let Some(value) = &self.properties {
            xml.push('>');
            xml.push_str(&value.to_xml_fragment()?);
            xml.push_str(&format!("</style:{tag}>"))
        } else {
            xml.push_str("/>")
        }
        Ok(xml)
    }
}
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct ChartStylePropertiesSet {
    pub styles: Vec<ChartStyleRecord>,
}
impl ChartStylePropertiesSet {
    pub fn get(&self, name: &str) -> Option<&ChartStyleRecord> {
        self.styles
            .iter()
            .find(|style| style.name.as_deref() == Some(name))
    }
    pub fn default_style(&self) -> Option<&ChartStyleRecord> {
        self.styles.iter().find(|style| style.is_default_style)
    }
}

#[derive(Clone, Copy, PartialEq, Eq, Hash)]
enum Ns {
    Office,
    Style,
    Chart,
    Text,
    Xlink,
    Other,
}
fn ns(value: ResolveResult<'_>) -> Ns {
    match value {
        ResolveResult::Bound(value) if value.as_ref() == OFFICE => Ns::Office,
        ResolveResult::Bound(value) if value.as_ref() == STYLE => Ns::Style,
        ResolveResult::Bound(value) if value.as_ref() == CHART => Ns::Chart,
        ResolveResult::Bound(value) if value.as_ref() == TEXT => Ns::Text,
        ResolveResult::Bound(value) if value.as_ref() == XLINK => Ns::Xlink,
        _ => Ns::Other,
    }
}
fn element(reader: &NsReader<&[u8]>, name: QName<'_>) -> (Ns, Vec<u8>) {
    let (namespace, local) = reader.resolver().resolve_element(name);
    (ns(namespace), local.as_ref().to_vec())
}
fn attrs(
    reader: &NsReader<&[u8]>,
    version: XmlVersion,
    start: &BytesStart<'_>,
) -> Result<Vec<(Ns, Vec<u8>, String)>> {
    let mut out = Vec::new();
    let mut seen = HashSet::new();
    for attribute in start.attributes().with_checks(true) {
        let attribute =
            attribute.map_err(|error| bad(format!("invalid chart property attribute: {error}")))?;
        if attribute.key.as_ref() == b"xmlns" || attribute.key.as_ref().starts_with(b"xmlns:") {
            continue;
        }
        if out.len() >= MAX_ATTRIBUTES {
            return Err(bad("too many chart property attributes"));
        }
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        let key = (ns(namespace), local.as_ref().to_vec());
        if !seen.insert(key.clone()) {
            return Err(bad("duplicate chart property attribute"));
        }
        let value = attribute
            .decoded_and_normalized_value(version, reader.decoder())
            .map_err(|error| bad(format!("invalid chart property value: {error}")))?
            .into_owned();
        safe(&value, "chart property value", true)?;
        out.push((key.0, key.1, value))
    }
    Ok(out)
}
fn take(a: &mut Vec<(Ns, Vec<u8>, String)>, namespace: Ns, local: &[u8]) -> Option<String> {
    a.iter()
        .position(|value| value.0 == namespace && value.1 == local)
        .map(|index| a.remove(index).2)
}
fn boolean(value: &str) -> Result<bool> {
    match value {
        "true" => Ok(true),
        "false" => Ok(false),
        _ => Err(bad("ODF boolean must be true or false")),
    }
}
fn e<T>(value: Option<String>, parse: fn(&str) -> Result<T>) -> Result<Option<T>> {
    value.map(|value| parse(&value)).transpose()
}
fn header(
    reader: &NsReader<&[u8]>,
    version: XmlVersion,
    start: &BytesStart<'_>,
    default: bool,
) -> Result<Option<ChartStyleRecord>> {
    let mut a = attrs(reader, version, start)?;
    if take(&mut a, Ns::Style, b"family").as_deref() != Some("chart") {
        return Ok(None);
    }
    let style = ChartStyleRecord {
        name: take(&mut a, Ns::Style, b"name"),
        parent_style_name: take(&mut a, Ns::Style, b"parent-style-name"),
        is_default_style: default,
        properties: None,
    };
    style.validate()?;
    Ok(Some(style))
}
fn properties(
    reader: &NsReader<&[u8]>,
    version: XmlVersion,
    start: &BytesStart<'_>,
) -> Result<ChartStyleProperties> {
    let mut a = attrs(reader, version, start)?;
    macro_rules! b {
        ($n:literal) => {
            take(&mut a, Ns::Chart, $n)
                .map(|v| boolean(&v))
                .transpose()?
        };
    }
    let value = ChartStyleProperties {
        scale_text: b!(b"scale-text"),
        three_dimensional: b!(b"three-dimensional"),
        deep: b!(b"deep"),
        right_angled_axes: b!(b"right-angled-axes"),
        symbol_type: e(
            take(&mut a, Ns::Chart, b"symbol-type"),
            ChartSymbolType::parse,
        )?,
        symbol_name: e(
            take(&mut a, Ns::Chart, b"symbol-name"),
            ChartSymbolName::parse,
        )?,
        symbol_image: None,
        symbol_width: take(&mut a, Ns::Chart, b"symbol-width")
            .map(ChartNonNegativeLength::new)
            .transpose()?,
        symbol_height: take(&mut a, Ns::Chart, b"symbol-height")
            .map(ChartNonNegativeLength::new)
            .transpose()?,
        sort_by_x_values: b!(b"sort-by-x-values"),
        vertical: b!(b"vertical"),
        connect_bars: b!(b"connect-bars"),
        gap_width: take(&mut a, Ns::Chart, b"gap-width")
            .map(ChartInteger::new)
            .transpose()?,
        overlap: take(&mut a, Ns::Chart, b"overlap")
            .map(ChartInteger::new)
            .transpose()?,
        group_bars_per_axis: b!(b"group-bars-per-axis"),
        japanese_candle_stick: b!(b"japanese-candle-stick"),
        interpolation: e(
            take(&mut a, Ns::Chart, b"interpolation"),
            ChartInterpolation::parse,
        )?,
        spline_order: take(&mut a, Ns::Chart, b"spline-order")
            .map(ChartPositiveInteger::new)
            .transpose()?,
        spline_resolution: take(&mut a, Ns::Chart, b"spline-resolution")
            .map(ChartPositiveInteger::new)
            .transpose()?,
        pie_offset: take(&mut a, Ns::Chart, b"pie-offset")
            .map(ChartNonNegativeInteger::new)
            .transpose()?,
        angle_offset: take(&mut a, Ns::Chart, b"angle-offset")
            .map(ChartAngle::new)
            .transpose()?,
        hole_size: take(&mut a, Ns::Chart, b"hole-size")
            .map(ChartPercent::new)
            .transpose()?,
        lines: b!(b"lines"),
        solid_type: e(
            take(&mut a, Ns::Chart, b"solid-type"),
            ChartSolidType::parse,
        )?,
        stacked: b!(b"stacked"),
        percentage: b!(b"percentage"),
        treat_empty_cells: e(
            take(&mut a, Ns::Chart, b"treat-empty-cells"),
            ChartEmptyCellTreatment::parse,
        )?,
        link_data_style_to_source: b!(b"link-data-style-to-source"),
        logarithmic: b!(b"logarithmic"),
        maximum: take(&mut a, Ns::Chart, b"maximum")
            .map(ChartDouble::new)
            .transpose()?,
        minimum: take(&mut a, Ns::Chart, b"minimum")
            .map(ChartDouble::new)
            .transpose()?,
        origin: take(&mut a, Ns::Chart, b"origin")
            .map(ChartDouble::new)
            .transpose()?,
        interval_major: take(&mut a, Ns::Chart, b"interval-major")
            .map(ChartDouble::new)
            .transpose()?,
        interval_minor_divisor: take(&mut a, Ns::Chart, b"interval-minor-divisor")
            .map(ChartPositiveInteger::new)
            .transpose()?,
        tick_marks_major_inner: b!(b"tick-marks-major-inner"),
        tick_marks_major_outer: b!(b"tick-marks-major-outer"),
        tick_marks_minor_inner: b!(b"tick-marks-minor-inner"),
        tick_marks_minor_outer: b!(b"tick-marks-minor-outer"),
        reverse_direction: b!(b"reverse-direction"),
        display_label: b!(b"display-label"),
        text_overlap: b!(b"text-overlap"),
        line_break: take(&mut a, Ns::Text, b"line-break")
            .map(|v| boolean(&v))
            .transpose()?,
        label_arrangement: e(
            take(&mut a, Ns::Chart, b"label-arrangement"),
            ChartLabelArrangement::parse,
        )?,
        direction: e(take(&mut a, Ns::Style, b"direction"), ChartDirection::parse)?,
        rotation_angle: take(&mut a, Ns::Style, b"rotation-angle")
            .map(ChartAngle::new)
            .transpose()?,
        data_label_number: e(
            take(&mut a, Ns::Chart, b"data-label-number"),
            ChartDataLabelNumber::parse,
        )?,
        data_label_text: b!(b"data-label-text"),
        data_label_symbol: b!(b"data-label-symbol"),
        label_separator: None,
        label_position: e(
            take(&mut a, Ns::Chart, b"label-position"),
            ChartLabelPosition::parse,
        )?,
        label_position_negative: e(
            take(&mut a, Ns::Chart, b"label-position-negative"),
            ChartLabelPosition::parse,
        )?,
        visible: b!(b"visible"),
        auto_position: b!(b"auto-position"),
        auto_size: b!(b"auto-size"),
        mean_value: b!(b"mean-value"),
        error_category: e(
            take(&mut a, Ns::Chart, b"error-category"),
            ChartErrorCategory::parse,
        )?,
        error_percentage: take(&mut a, Ns::Chart, b"error-percentage")
            .map(ChartDouble::new)
            .transpose()?,
        error_margin: take(&mut a, Ns::Chart, b"error-margin")
            .map(ChartDouble::new)
            .transpose()?,
        error_lower_limit: take(&mut a, Ns::Chart, b"error-lower-limit")
            .map(ChartDouble::new)
            .transpose()?,
        error_upper_limit: take(&mut a, Ns::Chart, b"error-upper-limit")
            .map(ChartDouble::new)
            .transpose()?,
        error_upper_indicator: b!(b"error-upper-indicator"),
        error_lower_indicator: b!(b"error-lower-indicator"),
        series_source: e(
            take(&mut a, Ns::Chart, b"series-source"),
            ChartSeriesSource::parse,
        )?,
        regression_type: e(
            take(&mut a, Ns::Chart, b"regression-type"),
            ChartRegressionType::parse,
        )?,
        axis_position: take(&mut a, Ns::Chart, b"axis-position")
            .map(|v| ChartAxisPosition::parse(&v))
            .transpose()?,
        axis_label_position: e(
            take(&mut a, Ns::Chart, b"axis-label-position"),
            ChartAxisLabelPosition::parse,
        )?,
        tick_mark_position: e(
            take(&mut a, Ns::Chart, b"tick-mark-position"),
            ChartTickMarkPosition::parse,
        )?,
        include_hidden_cells: b!(b"include-hidden-cells"),
    };
    if !a.is_empty() {
        return Err(bad(
            "unknown style:chart-properties attribute or wrong namespace",
        ));
    }
    Ok(value)
}
fn symbol_image(
    reader: &NsReader<&[u8]>,
    version: XmlVersion,
    start: &BytesStart<'_>,
) -> Result<ChartSymbolImage> {
    let mut a = attrs(reader, version, start)?;
    let href = take(&mut a, Ns::Xlink, b"href")
        .ok_or_else(|| bad("chart:symbol-image requires xlink:href"))?;
    if !a.is_empty() {
        return Err(bad(
            "unknown chart:symbol-image attribute or wrong namespace",
        ));
    }
    ChartSymbolImage::new(href)
}
fn no_attrs(
    reader: &NsReader<&[u8]>,
    version: XmlVersion,
    start: &BytesStart<'_>,
    name: &str,
) -> Result<()> {
    if !attrs(reader, version, start)?.is_empty() {
        return Err(bad(format!("{name} does not allow attributes")));
    }
    Ok(())
}
fn boundary(xml: &str, end: usize) -> Result<usize> {
    xml[..end]
        .rfind('<')
        .ok_or_else(|| bad("invalid XML event boundary"))
}

struct Active {
    depth: usize,
    style: ChartStyleRecord,
    seen: bool,
    property_depth: Option<usize>,
    symbol_depth: Option<usize>,
    label_depth: Option<usize>,
    paragraph_depth: Option<usize>,
    paragraph_start: Option<usize>,
}
fn push(out: &mut Vec<ChartStyleRecord>, style: ChartStyleRecord, total: &mut usize) -> Result<()> {
    if out.len() >= MAX_STYLES
        || out.iter().any(|value| {
            value.name == style.name && value.is_default_style == style.is_default_style
        })
    {
        return Err(bad("duplicate or excessive chart style"));
    }
    *total += style.to_xml_fragment()?.len();
    if *total > MAX_TOTAL {
        return Err(bad("chart style data is too large"));
    }
    out.push(style);
    Ok(())
}
/// Parse direct chart-family styles in standard style containers.
pub fn parse_chart_style_properties(xml: &str) -> Result<ChartStylePropertiesSet> {
    if xml.len() > MAX_XML {
        return Err(bad("styles XML is too large"));
    }
    let mut reader = NsReader::from_reader(xml.as_bytes());
    reader.config_mut().trim_text(false);
    let mut version = XmlVersion::Implicit1_0;
    let mut stack: Vec<(Ns, Vec<u8>)> = Vec::new();
    let mut active: Option<Active> = None;
    let mut out = Vec::new();
    let mut total = 0;
    let mut events = 0;
    loop {
        events += 1;
        if events > MAX_EVENTS {
            return Err(bad("styles XML has too many events"));
        }
        match reader.read_event() {
            Ok(Event::Start(start)) => {
                if stack.len() >= MAX_DEPTH {
                    return Err(bad("styles XML is too deep"));
                }
                let end = reader.buffer_position() as usize;
                let begin = boundary(xml, end)?;
                let current = element(&reader, start.name());
                let parent = stack.last();
                let direct = parent.is_some_and(|value| {
                    value.0 == Ns::Office
                        && matches!(value.1.as_slice(), b"styles" | b"automatic-styles")
                }) && current.0 == Ns::Style
                    && matches!(current.1.as_slice(), b"style" | b"default-style");
                stack.push(current.clone());
                let depth = stack.len();
                if direct {
                    if let Some(style) =
                        header(&reader, version, &start, current.1 == b"default-style")?
                    {
                        active = Some(Active {
                            depth,
                            style,
                            seen: false,
                            property_depth: None,
                            symbol_depth: None,
                            label_depth: None,
                            paragraph_depth: None,
                            paragraph_start: None,
                        })
                    }
                    continue;
                }
                if let Some(value) = active.as_mut() {
                    if value.paragraph_depth.is_some() {
                        continue;
                    }
                    if depth == value.depth + 1
                        && current.0 == Ns::Style
                        && current.1 == b"chart-properties"
                    {
                        if value.seen {
                            return Err(bad("duplicate style:chart-properties"));
                        }
                        value.seen = true;
                        value.style.properties = Some(properties(&reader, version, &start)?);
                        value.property_depth = Some(depth)
                    } else if current.1 == b"chart-properties" {
                        return Err(bad(
                            "style:chart-properties has invalid namespace or parent",
                        ));
                    } else if value.property_depth.is_some_and(|p| depth == p + 1)
                        && current.0 == Ns::Chart
                        && current.1 == b"symbol-image"
                    {
                        if value
                            .style
                            .properties
                            .as_ref()
                            .unwrap()
                            .symbol_image
                            .is_some()
                        {
                            return Err(bad("duplicate chart:symbol-image"));
                        }
                        value.style.properties.as_mut().unwrap().symbol_image =
                            Some(symbol_image(&reader, version, &start)?);
                        value.symbol_depth = Some(depth)
                    } else if value.property_depth.is_some_and(|p| depth == p + 1)
                        && current.0 == Ns::Chart
                        && current.1 == b"label-separator"
                    {
                        if value.label_depth.is_some()
                            || value
                                .style
                                .properties
                                .as_ref()
                                .unwrap()
                                .label_separator
                                .is_some()
                        {
                            return Err(bad("duplicate chart:label-separator"));
                        }
                        no_attrs(&reader, version, &start, "chart:label-separator")?;
                        value.label_depth = Some(depth)
                    } else if value.label_depth.is_some_and(|l| depth == l + 1)
                        && current.0 == Ns::Text
                        && current.1 == b"p"
                    {
                        if value.paragraph_start.is_some() {
                            return Err(bad("duplicate text:p in chart:label-separator"));
                        }
                        value.paragraph_depth = Some(depth);
                        value.paragraph_start = Some(begin)
                    } else if value.property_depth.is_some_and(|p| depth > p) {
                        return Err(bad("unexpected style:chart-properties child"));
                    }
                }
            },
            Ok(Event::Empty(start)) => {
                let end = reader.buffer_position() as usize;
                let begin = boundary(xml, end)?;
                let current = element(&reader, start.name());
                let parent = stack.last();
                let depth = stack.len() + 1;
                let direct = parent.is_some_and(|value| {
                    value.0 == Ns::Office
                        && matches!(value.1.as_slice(), b"styles" | b"automatic-styles")
                }) && current.0 == Ns::Style
                    && matches!(current.1.as_slice(), b"style" | b"default-style");
                if direct {
                    if let Some(style) =
                        header(&reader, version, &start, current.1 == b"default-style")?
                    {
                        push(&mut out, style, &mut total)?
                    }
                    continue;
                }
                if let Some(value) = active.as_mut() {
                    if value.paragraph_depth.is_some() {
                        continue;
                    }
                    if depth == value.depth + 1
                        && current.0 == Ns::Style
                        && current.1 == b"chart-properties"
                    {
                        if value.seen {
                            return Err(bad("duplicate style:chart-properties"));
                        }
                        value.seen = true;
                        let parsed = properties(&reader, version, &start)?;
                        parsed.validate()?;
                        value.style.properties = Some(parsed)
                    } else if current.1 == b"chart-properties" {
                        return Err(bad(
                            "style:chart-properties has invalid namespace or parent",
                        ));
                    } else if value.property_depth.is_some_and(|p| depth == p + 1)
                        && current.0 == Ns::Chart
                        && current.1 == b"symbol-image"
                    {
                        if value
                            .style
                            .properties
                            .as_ref()
                            .unwrap()
                            .symbol_image
                            .is_some()
                        {
                            return Err(bad("duplicate chart:symbol-image"));
                        }
                        value.style.properties.as_mut().unwrap().symbol_image =
                            Some(symbol_image(&reader, version, &start)?)
                    } else if value.property_depth.is_some_and(|p| depth == p + 1)
                        && current.0 == Ns::Chart
                        && current.1 == b"label-separator"
                    {
                        return Err(bad("chart:label-separator requires text:p"));
                    } else if value.label_depth.is_some_and(|l| depth == l + 1)
                        && current.0 == Ns::Text
                        && current.1 == b"p"
                    {
                        if value
                            .style
                            .properties
                            .as_ref()
                            .unwrap()
                            .label_separator
                            .is_some()
                        {
                            return Err(bad("duplicate text:p in chart:label-separator"));
                        }
                        value.style.properties.as_mut().unwrap().label_separator =
                            Some(ChartLabelSeparator::from_paragraph_xml(&xml[begin..end])?)
                    } else if value.property_depth.is_some_and(|p| depth > p) {
                        return Err(bad("unexpected style:chart-properties child"));
                    }
                }
            },
            Ok(Event::Text(text)) => {
                let bytes: &[u8] = text.as_ref();
                if active
                    .as_ref()
                    .is_some_and(|v| v.property_depth.is_some() && v.paragraph_depth.is_none())
                    && !bytes.iter().all(u8::is_ascii_whitespace)
                {
                    return Err(bad("unexpected text in style:chart-properties"));
                }
            },
            Ok(Event::CData(text)) => {
                let bytes: &[u8] = text.as_ref();
                if active
                    .as_ref()
                    .is_some_and(|v| v.property_depth.is_some() && v.paragraph_depth.is_none())
                    && !bytes.iter().all(u8::is_ascii_whitespace)
                {
                    return Err(bad("unexpected text in style:chart-properties"));
                }
            },
            Ok(Event::End(_)) => {
                let end = reader.buffer_position() as usize;
                let depth = stack.len();
                if let Some(value) = active.as_mut() {
                    if value.paragraph_depth == Some(depth) {
                        let begin = value.paragraph_start.take().unwrap();
                        value.style.properties.as_mut().unwrap().label_separator =
                            Some(ChartLabelSeparator::from_paragraph_xml(&xml[begin..end])?);
                        value.paragraph_depth = None
                    }
                    if value.symbol_depth == Some(depth) {
                        value.symbol_depth = None
                    }
                    if value.label_depth == Some(depth) {
                        if value
                            .style
                            .properties
                            .as_ref()
                            .unwrap()
                            .label_separator
                            .is_none()
                        {
                            return Err(bad("chart:label-separator requires text:p"));
                        }
                        value.label_depth = None
                    }
                    if value.property_depth == Some(depth) {
                        value.style.properties.as_ref().unwrap().validate()?;
                        value.property_depth = None
                    }
                }
                if active.as_ref().is_some_and(|value| value.depth == depth) {
                    push(&mut out, active.take().unwrap().style, &mut total)?
                }
                stack.pop();
            },
            Ok(Event::Decl(decl)) => {
                version = decl
                    .xml_version()
                    .map_err(|error| bad(format!("unsupported XML version: {error}")))?
            },
            Ok(Event::DocType(_)) | Ok(Event::PI(_)) => {
                return Err(bad("DTD and processing instructions are not allowed"));
            },
            Ok(Event::Eof) => break,
            Ok(_) => {},
            Err(error) => return Err(bad(format!("invalid styles XML: {error}"))),
        }
    }
    if !stack.is_empty() || active.is_some() {
        return Err(bad("truncated styles XML"));
    }
    Ok(ChartStylePropertiesSet { styles: out })
}

#[derive(Default)]
struct Span {
    start: usize,
    end: usize,
    end_start: usize,
    qname: String,
    empty: bool,
}
#[derive(Default)]
struct Target {
    style: Span,
    properties: Option<Span>,
}
fn replace(xml: &str, span: &Span, value: &str) -> String {
    format!("{}{}{}", &xml[..span.start], value, &xml[span.end..])
}
fn expand(xml: &str, span: &Span, value: &str) -> Result<String> {
    let raw = &xml[span.start..span.end];
    let slash = raw.rfind("/>").ok_or_else(|| bad("invalid empty style"))?;
    Ok(replace(
        xml,
        span,
        &format!("{}>{value}</{}>", &raw[..slash], span.qname),
    ))
}
/// Losslessly replace, insert, or remove one existing chart style property element.
pub fn set_chart_style_properties_xml(xml: &str, requested: &ChartStyleRecord) -> Result<String> {
    requested.validate()?;
    if xml.len() > MAX_XML {
        return Err(bad("styles XML is too large"));
    }
    let mut reader = NsReader::from_reader(xml.as_bytes());
    let mut version = XmlVersion::Implicit1_0;
    let mut stack: Vec<(Ns, Vec<u8>)> = Vec::new();
    let mut target_depth = None;
    let mut active: Option<Target> = None;
    let mut found = None;
    loop {
        match reader.read_event() {
            Ok(Event::Start(start)) => {
                let end = reader.buffer_position() as usize;
                let begin = boundary(xml, end)?;
                let current = element(&reader, start.name());
                let parent = stack.last();
                let direct = parent.is_some_and(|v| {
                    v.0 == Ns::Office && matches!(v.1.as_slice(), b"styles" | b"automatic-styles")
                }) && current.0 == Ns::Style
                    && matches!(current.1.as_slice(), b"style" | b"default-style");
                stack.push(current.clone());
                let depth = stack.len();
                if direct {
                    if let Some(style) =
                        header(&reader, version, &start, current.1 == b"default-style")?
                        && style.name == requested.name
                        && style.is_default_style == requested.is_default_style
                    {
                        if active.is_some() || found.is_some() {
                            return Err(bad("duplicate target chart style"));
                        }
                        target_depth = Some(depth);
                        active = Some(Target {
                            style: Span {
                                start: begin,
                                qname: String::from_utf8_lossy(start.name().as_ref()).into_owned(),
                                ..Default::default()
                            },
                            ..Default::default()
                        })
                    }
                } else if target_depth.is_some_and(|d| depth == d + 1)
                    && current.0 == Ns::Style
                    && current.1 == b"chart-properties"
                {
                    let span = Span {
                        start: begin,
                        qname: String::from_utf8_lossy(start.name().as_ref()).into_owned(),
                        ..Default::default()
                    };
                    if active.as_mut().unwrap().properties.replace(span).is_some() {
                        return Err(bad("duplicate style:chart-properties"));
                    }
                }
            },
            Ok(Event::Empty(start)) => {
                let end = reader.buffer_position() as usize;
                let begin = boundary(xml, end)?;
                let current = element(&reader, start.name());
                let parent = stack.last();
                let depth = stack.len() + 1;
                let direct = parent.is_some_and(|v| {
                    v.0 == Ns::Office && matches!(v.1.as_slice(), b"styles" | b"automatic-styles")
                }) && current.0 == Ns::Style
                    && matches!(current.1.as_slice(), b"style" | b"default-style");
                let span = Span {
                    start: begin,
                    end,
                    end_start: begin,
                    qname: String::from_utf8_lossy(start.name().as_ref()).into_owned(),
                    empty: true,
                };
                if direct {
                    if let Some(style) =
                        header(&reader, version, &start, current.1 == b"default-style")?
                        && style.name == requested.name
                        && style.is_default_style == requested.is_default_style
                    {
                        if active.is_some() || found.is_some() {
                            return Err(bad("duplicate target chart style"));
                        }
                        found = Some(Target {
                            style: span,
                            ..Default::default()
                        })
                    }
                } else if target_depth.is_some_and(|d| depth == d + 1)
                    && current.0 == Ns::Style
                    && current.1 == b"chart-properties"
                {
                    if active.as_mut().unwrap().properties.replace(span).is_some() {
                        return Err(bad("duplicate style:chart-properties"));
                    }
                }
            },
            Ok(Event::End(_)) => {
                let end = reader.buffer_position() as usize;
                let begin = boundary(xml, end)?;
                let depth = stack.len();
                if let Some(spans) = active.as_mut() {
                    if spans.properties.as_ref().is_some_and(|s| s.end == 0)
                        && target_depth.is_some_and(|d| depth == d + 1)
                    {
                        let span = spans.properties.as_mut().unwrap();
                        span.end_start = begin;
                        span.end = end
                    }
                    if target_depth == Some(depth) {
                        spans.style.end_start = begin;
                        spans.style.end = end;
                        found = active.take();
                        target_depth = None
                    }
                }
                stack.pop();
            },
            Ok(Event::Decl(decl)) => {
                version = decl
                    .xml_version()
                    .map_err(|error| bad(format!("unsupported XML version: {error}")))?
            },
            Ok(Event::DocType(_)) | Ok(Event::PI(_)) => {
                return Err(bad("DTD and processing instructions are not allowed"));
            },
            Ok(Event::Eof) => break,
            Ok(_) => {},
            Err(error) => return Err(bad(format!("invalid styles XML: {error}"))),
        }
    }
    let spans = found.ok_or_else(|| bad("target chart style does not exist"))?;
    let replacement = requested
        .properties
        .as_ref()
        .map(ChartStyleProperties::to_xml_fragment)
        .transpose()?;
    if let Some(properties) = &spans.properties {
        return Ok(replace(
            xml,
            properties,
            replacement.as_deref().unwrap_or(""),
        ));
    }
    let Some(replacement) = replacement else {
        return Ok(xml.to_owned());
    };
    if spans.style.empty {
        return expand(xml, &spans.style, &replacement);
    }
    let mut out = xml.to_owned();
    out.insert_str(spans.style.end_start, &replacement);
    Ok(out)
}

impl OpenDocumentPackage {
    pub fn chart_style_properties(&self) -> Result<ChartStylePropertiesSet> {
        self.styles_xml()?.map_or_else(
            || Ok(Default::default()),
            |xml| parse_chart_style_properties(&xml),
        )
    }
}
impl FlatOpenDocument {
    pub fn chart_style_properties(&self) -> Result<ChartStylePropertiesSet> {
        parse_chart_style_properties(self.xml())
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    const HEAD: &str = r#"<office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:xlink="http://www.w3.org/1999/xlink"><office:automatic-styles>"#;
    fn doc(body: &str) -> String {
        format!("{HEAD}{body}</office:automatic-styles></office:document>")
    }
    #[test]
    fn complete_family_round_trips() {
        let xml = doc(
            r#"<style:style style:name="ch1" style:family="chart"><style:chart-properties chart:scale-text="true" chart:three-dimensional="false" chart:deep="true" chart:right-angled-axes="false" chart:symbol-type="image" chart:symbol-width=".5cm" chart:symbol-height="0px" chart:sort-by-x-values="true" chart:vertical="false" chart:connect-bars="true" chart:gap-width="-001" chart:overlap="+2" chart:group-bars-per-axis="false" chart:japanese-candle-stick="true" chart:interpolation="b-spline" chart:spline-order="+03" chart:spline-resolution="4" chart:pie-offset="000" chart:angle-offset="344.61" chart:hole-size="-.5%" chart:lines="true" chart:solid-type="pyramid" chart:stacked="false" chart:percentage="true" chart:treat-empty-cells="leave-gap" chart:link-data-style-to-source="false" chart:logarithmic="true" chart:maximum="INF" chart:minimum="-INF" chart:origin="NaN" chart:interval-major="+1.2E-3" chart:interval-minor-divisor="2" chart:tick-marks-major-inner="true" chart:tick-marks-major-outer="false" chart:tick-marks-minor-inner="true" chart:tick-marks-minor-outer="false" chart:reverse-direction="true" chart:display-label="false" chart:text-overlap="true" text:line-break="false" chart:label-arrangement="stagger-even" style:direction="ttb" style:rotation-angle="" chart:data-label-number="value-and-percentage" chart:data-label-text="true" chart:data-label-symbol="false" chart:label-position="near-origin" chart:label-position-negative="avoid-overlap" chart:visible="true" chart:auto-position="false" chart:auto-size="true" chart:mean-value="false" chart:error-category="cell-range" chart:error-percentage="5" chart:error-margin=".1" chart:error-lower-limit="-2" chart:error-upper-limit="2" chart:error-upper-indicator="true" chart:error-lower-indicator="false" chart:series-source="columns" chart:regression-type="power" chart:axis-position="-0.0" chart:axis-label-position="near-axis-other-side" chart:tick-mark-position="at-labels-and-axis" chart:include-hidden-cells="false"><chart:symbol-image xlink:href="Pictures/a&amp;b.svg"/><chart:label-separator><text:p> / <text:span>rich</text:span></text:p></chart:label-separator></style:chart-properties></style:style>"#,
        );
        let set = parse_chart_style_properties(&xml).unwrap();
        let value = set.get("ch1").unwrap().properties.as_ref().unwrap();
        assert_eq!(value.maximum.as_ref().unwrap().as_str(), "INF");
        assert_eq!(
            value.symbol_image.as_ref().unwrap().href,
            "Pictures/a&b.svg"
        );
        let fragment = value.to_xml_fragment().unwrap();
        assert_eq!(
            ChartStyleProperties::from_xml_fragment(&fragment).unwrap(),
            *value
        )
    }
    #[test]
    fn parses_real_libreoffice_standard_style() {
        let fixture = include_str!(
            "../../../3rdparty/libreoffice-core/chart2/qa/extras/data/fods/stacked-column-chart.fods"
        );
        let begin = fixture
            .find(r#"<style:style style:name="ch3" style:family="chart">"#)
            .unwrap();
        let end = begin + fixture[begin..].find("</style:style>").unwrap() + "</style:style>".len();
        let set = parse_chart_style_properties(&doc(&fixture[begin..end])).unwrap();
        let value = set.get("ch3").unwrap().properties.as_ref().unwrap();
        assert_eq!(value.stacked, Some(true));
        assert_eq!(value.series_source, Some(ChartSeriesSource::Rows));
        assert_eq!(
            value.treat_empty_cells,
            Some(ChartEmptyCellTreatment::LeaveGap)
        )
    }
    #[test]
    fn lossless_replace_insert_remove() {
        let original = doc(
            "<!--keep--><style:style style:name=\"a\" style:family=\"chart\"><x:k xmlns:x=\"urn:k\"/></style:style><style:style style:name=\"b\" style:family=\"chart\"><style:chart-properties chart:lines=\"true\"/></style:style>",
        );
        let mut a = ChartStyleRecord::named(
            "a",
            Some(ChartStyleProperties {
                stacked: Some(true),
                ..Default::default()
            }),
        )
        .unwrap();
        let inserted = set_chart_style_properties_xml(&original, &a).unwrap();
        assert!(inserted.contains("<x:k xmlns:x=\"urn:k\"/><style:chart-properties"));
        a.properties = None;
        let restored = set_chart_style_properties_xml(&inserted, &a).unwrap();
        assert_eq!(restored, original);
        let b = ChartStyleRecord::named("b", None).unwrap();
        let removed = set_chart_style_properties_xml(&restored, &b).unwrap();
        assert!(!removed.contains("chart:lines=\"true\""));
        assert!(removed.contains("<!--keep-->"))
    }
    #[test]
    fn rejects_malformed_lexicals_namespaces_placement_and_children() {
        let cases = [
            r#"<style:style style:name="a" style:family="chart"><style:chart-properties chart:stacked="1"/></style:style>"#,
            r#"<style:style style:name="a" style:family="chart"><style:chart-properties chart:spline-order="0"/></style:style>"#,
            r#"<style:style style:name="a" style:family="chart"><style:chart-properties chart:symbol-type="named-symbol"/></style:style>"#,
            r#"<style:style style:name="a" style:family="chart"><style:chart-properties chart:symbol-type="image"><chart:symbol-image/></style:chart-properties></style:style>"#,
            r#"<style:style style:name="a" style:family="chart"><style:chart-properties chart:lines="true" chart:lines="false"/></style:style>"#,
            r#"<style:style style:name="a" style:family="chart"><chart:chart-properties/></style:style>"#,
            r#"<style:style style:name="a" style:family="chart"><style:chart-properties><chart:label-separator/></style:chart-properties></style:style>"#,
            r#"<style:style style:name="a" style:family="chart"><style:chart-properties><chart:label-separator><chart:p/></chart:label-separator></style:chart-properties></style:style>"#,
            r#"<style:style style:name="a" style:family="chart"><style:chart-properties lo:extension="1" xmlns:lo="urn:extension"/></style:style>"#,
        ];
        for case in cases {
            assert!(
                parse_chart_style_properties(&doc(case)).is_err(),
                "accepted {case}"
            )
        }
    }
}
