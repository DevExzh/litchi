//! Typed ODF chart-style property values and invariants.

use litchi_core::{Error, Result, xml::escape_xml};

use super::{MAX_VALUE, TEXT_NS};

pub(super) fn bad(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}
pub(super) fn safe(value: &str, name: &str, empty: bool) -> Result<()> {
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
pub(super) fn ncname(value: &str, name: &str) -> Result<()> {
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
        exponent = Some(&value[index + 1..]);
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
    NonNegativeLength,
    nonnegative_length,
    "chart non-negative length",
    false
);
lexical!(Integer, integer, "chart integer", false);
lexical!(
    PositiveInteger,
    positive_integer,
    "chart positive integer",
    false
);
lexical!(
    NonNegativeInteger,
    nonnegative_integer,
    "chart non-negative integer",
    false
);
lexical!(Percent, percent, "chart percentage", false);
lexical!(Double, double, "chart double", false);
lexical!(Angle, |_: &str| true, "chart angle", true);

macro_rules! keyword_enum{($name:ident{$($variant:ident=>$value:literal),+$(,)?})=>{#[derive(Debug,Clone,Copy,PartialEq,Eq)]pub enum $name{$($variant),+}impl $name{pub(super) fn parse(value:&str)->Result<Self>{match value{$($value=>Ok(Self::$variant),)+_=>Err(bad(concat!("invalid ",stringify!($name))))}}pub(super) fn xml(self)->&'static str{match self{$(Self::$variant=>$value),+}}}};}
keyword_enum!(SymbolType{None=>"none",Automatic=>"automatic",NamedSymbol=>"named-symbol",Image=>"image"});
keyword_enum!(SymbolName{Square=>"square",Diamond=>"diamond",ArrowDown=>"arrow-down",ArrowUp=>"arrow-up",ArrowRight=>"arrow-right",ArrowLeft=>"arrow-left",BowTie=>"bow-tie",Hourglass=>"hourglass",Circle=>"circle",Star=>"star",X=>"x",Plus=>"plus",Asterisk=>"asterisk",HorizontalBar=>"horizontal-bar",VerticalBar=>"vertical-bar"});
keyword_enum!(Interpolation{None=>"none",CubicSpline=>"cubic-spline",BSpline=>"b-spline"});
keyword_enum!(SolidType{Cuboid=>"cuboid",Cylinder=>"cylinder",Cone=>"cone",Pyramid=>"pyramid"});
keyword_enum!(EmptyCellTreatment{UseZero=>"use-zero",LeaveGap=>"leave-gap",Ignore=>"ignore"});
keyword_enum!(LabelArrangement{SideBySide=>"side-by-side",StaggerEven=>"stagger-even",StaggerOdd=>"stagger-odd"});
keyword_enum!(Direction{Ltr=>"ltr",Ttb=>"ttb"});
keyword_enum!(DataLabelNumber{None=>"none",Value=>"value",Percentage=>"percentage",ValueAndPercentage=>"value-and-percentage"});
keyword_enum!(LabelPosition{AvoidOverlap=>"avoid-overlap",Center=>"center",Top=>"top",TopRight=>"top-right",Right=>"right",BottomRight=>"bottom-right",Bottom=>"bottom",BottomLeft=>"bottom-left",Left=>"left",TopLeft=>"top-left",Inside=>"inside",Outside=>"outside",NearOrigin=>"near-origin"});
keyword_enum!(ErrorCategory{None=>"none",Variance=>"variance",StandardDeviation=>"standard-deviation",Percentage=>"percentage",ErrorMargin=>"error-margin",Constant=>"constant",StandardError=>"standard-error",CellRange=>"cell-range"});
keyword_enum!(SeriesSource{Columns=>"columns",Rows=>"rows"});
keyword_enum!(RegressionType{None=>"none",Linear=>"linear",Logarithmic=>"logarithmic",Exponential=>"exponential",Power=>"power"});
keyword_enum!(AxisLabelPosition{NearAxis=>"near-axis",NearAxisOtherSide=>"near-axis-other-side",OutsideStart=>"outside-start",OutsideEnd=>"outside-end"});
keyword_enum!(TickMarkPosition{AtLabels=>"at-labels",AtAxis=>"at-axis",AtLabelsAndAxis=>"at-labels-and-axis"});

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum AxisPosition {
    Start,
    End,
    Value(Double),
}
impl AxisPosition {
    pub(super) fn parse(value: &str) -> Result<Self> {
        match value {
            "start" => Ok(Self::Start),
            "end" => Ok(Self::End),
            _ => Ok(Self::Value(Double::new(value)?)),
        }
    }
    pub(super) fn xml(&self) -> &str {
        match self {
            Self::Start => "start",
            Self::End => "end",
            Self::Value(value) => value.as_str(),
        }
    }
}

/// Inert image marker reference. The URI is never fetched.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SymbolImage {
    pub href: String,
}
impl SymbolImage {
    pub fn new(href: impl Into<String>) -> Result<Self> {
        let value = Self { href: href.into() };
        value.validate()?;
        Ok(value)
    }
    pub fn validate(&self) -> Result<()> {
        safe(&self.href, "chart:symbol-image xlink:href", true)
    }
}
/// Inert, bounded `text:p` XML inside `chart:label-separator`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct LabelSeparator {
    paragraph_xml: String,
}
impl LabelSeparator {
    pub fn new_text(text: &str) -> Result<Self> {
        safe(text, "chart label separator text", true)?;
        Self::from_paragraph_xml(format!(
            r#"<text:p xmlns:text="{TEXT_NS}">{}</text:p>"#,
            escape_xml(text)
        ))
    }
    pub fn from_paragraph_xml(paragraph_xml: impl Into<String>) -> Result<Self> {
        let paragraph_xml = paragraph_xml.into();
        super::codec::validate_paragraph(&paragraph_xml)?;
        Ok(Self { paragraph_xml })
    }
    pub fn paragraph_xml(&self) -> &str {
        &self.paragraph_xml
    }
}

/// All 66 attributes and both children allowed by `style:chart-properties`.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct StyleProperties {
    pub scale_text: Option<bool>,
    pub three_dimensional: Option<bool>,
    pub deep: Option<bool>,
    pub right_angled_axes: Option<bool>,
    pub symbol_type: Option<SymbolType>,
    pub symbol_name: Option<SymbolName>,
    pub symbol_image: Option<SymbolImage>,
    pub symbol_width: Option<NonNegativeLength>,
    pub symbol_height: Option<NonNegativeLength>,
    pub sort_by_x_values: Option<bool>,
    pub vertical: Option<bool>,
    pub connect_bars: Option<bool>,
    pub gap_width: Option<Integer>,
    pub overlap: Option<Integer>,
    pub group_bars_per_axis: Option<bool>,
    pub japanese_candle_stick: Option<bool>,
    pub interpolation: Option<Interpolation>,
    pub spline_order: Option<PositiveInteger>,
    pub spline_resolution: Option<PositiveInteger>,
    pub pie_offset: Option<NonNegativeInteger>,
    pub angle_offset: Option<Angle>,
    pub hole_size: Option<Percent>,
    pub lines: Option<bool>,
    pub solid_type: Option<SolidType>,
    pub stacked: Option<bool>,
    pub percentage: Option<bool>,
    pub treat_empty_cells: Option<EmptyCellTreatment>,
    pub link_data_style_to_source: Option<bool>,
    pub logarithmic: Option<bool>,
    pub maximum: Option<Double>,
    pub minimum: Option<Double>,
    pub origin: Option<Double>,
    pub interval_major: Option<Double>,
    pub interval_minor_divisor: Option<PositiveInteger>,
    pub tick_marks_major_inner: Option<bool>,
    pub tick_marks_major_outer: Option<bool>,
    pub tick_marks_minor_inner: Option<bool>,
    pub tick_marks_minor_outer: Option<bool>,
    pub reverse_direction: Option<bool>,
    pub display_label: Option<bool>,
    pub text_overlap: Option<bool>,
    pub line_break: Option<bool>,
    pub label_arrangement: Option<LabelArrangement>,
    pub direction: Option<Direction>,
    pub rotation_angle: Option<Angle>,
    pub data_label_number: Option<DataLabelNumber>,
    pub data_label_text: Option<bool>,
    pub data_label_symbol: Option<bool>,
    pub label_separator: Option<LabelSeparator>,
    pub label_position: Option<LabelPosition>,
    pub label_position_negative: Option<LabelPosition>,
    pub visible: Option<bool>,
    pub auto_position: Option<bool>,
    pub auto_size: Option<bool>,
    pub mean_value: Option<bool>,
    pub error_category: Option<ErrorCategory>,
    pub error_percentage: Option<Double>,
    pub error_margin: Option<Double>,
    pub error_lower_limit: Option<Double>,
    pub error_upper_limit: Option<Double>,
    pub error_upper_indicator: Option<bool>,
    pub error_lower_indicator: Option<bool>,
    pub series_source: Option<SeriesSource>,
    pub regression_type: Option<RegressionType>,
    pub axis_position: Option<AxisPosition>,
    pub axis_label_position: Option<AxisLabelPosition>,
    pub tick_mark_position: Option<TickMarkPosition>,
    pub include_hidden_cells: Option<bool>,
}

impl StyleProperties {
    pub fn validate(&self) -> Result<()> {
        match self.symbol_type {
            Some(SymbolType::NamedSymbol)
                if self.symbol_name.is_some() && self.symbol_image.is_none() => {},
            Some(SymbolType::Image)
                if self.symbol_image.is_some() && self.symbol_name.is_none() => {},
            Some(SymbolType::None | SymbolType::Automatic)
                if self.symbol_name.is_none() && self.symbol_image.is_none() => {},
            None if self.symbol_name.is_none() && self.symbol_image.is_none() => {},
            _ => {
                return Err(bad(
                    "chart symbol type/name/image combination violates ODF grammar",
                ));
            },
        }
        if let Some(image) = &self.symbol_image {
            image.validate()?;
        }
        if let Some(label) = &self.label_separator {
            super::codec::validate_paragraph(label.paragraph_xml())?;
        }
        Ok(())
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct StyleRecord {
    pub name: Option<String>,
    pub parent_style_name: Option<String>,
    pub is_default_style: bool,
    pub properties: Option<StyleProperties>,
}
impl StyleRecord {
    pub fn named(name: impl Into<String>, properties: Option<StyleProperties>) -> Result<Self> {
        let value = Self {
            name: Some(name.into()),
            parent_style_name: None,
            is_default_style: false,
            properties,
        };
        value.validate()?;
        Ok(value)
    }
    pub fn default_style(properties: Option<StyleProperties>) -> Self {
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
            ncname(value, "parent chart style name")?;
        }
        if let Some(value) = &self.properties {
            value.validate()?;
        }
        Ok(())
    }
}
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct StylePropertiesSet {
    pub styles: Vec<StyleRecord>,
}
impl StylePropertiesSet {
    pub fn get(&self, name: &str) -> Option<&StyleRecord> {
        self.styles
            .iter()
            .find(|style| style.name.as_deref() == Some(name))
    }
    pub fn default_style(&self) -> Option<&StyleRecord> {
        self.styles.iter().find(|style| style.is_default_style)
    }
}
