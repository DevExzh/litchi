//! Semantic SpreadsheetML conditional-formatting values and differential styles.

use crate::color::Rgb;

use smallvec::SmallVec;

use std::fmt;

use std::str::FromStr;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[repr(u8)]
pub enum Source {
    Core,
    Office2010,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Range(String);

impl Range {
    pub(crate) fn from_raw(value: String) -> Self {
        Self(value)
    }

    pub fn as_str(&self) -> &str {
        &self.0
    }
}

/// Error returned when a closed conditional-formatting token is invalid.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TokenError {
    domain: &'static str,
    token: Box<str>,
}

impl TokenError {
    /// Construct a token error for a closed SpreadsheetML vocabulary.
    #[must_use]
    pub fn new(domain: &'static str, token: &str) -> Self {
        Self {
            domain,
            token: token.into(),
        }
    }

    #[must_use]
    pub const fn domain(&self) -> &'static str {
        self.domain
    }

    #[must_use]
    pub fn token(&self) -> &str {
        &self.token
    }
}

impl fmt::Display for TokenError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(
            formatter,
            "invalid SpreadsheetML {} token '{}'",
            self.domain, self.token
        )
    }
}

impl std::error::Error for TokenError {}

macro_rules! token_enum {
    ($(#[$meta:meta])* pub enum $name:ident, $domain:literal { $($variant:ident => $token:literal),+ $(,)? }) => {
        $(#[$meta])*
        #[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
        #[repr(u8)]
        pub enum $name { $($variant),+ }

        impl $name {
            #[must_use]
            pub const fn as_str(self) -> &'static str {
                match self { $(Self::$variant => $token),+ }
            }
        }

        impl FromStr for $name {
            type Err = TokenError;

            fn from_str(value: &str) -> std::result::Result<Self, Self::Err> {
                match value {
                    $($token => Ok(Self::$variant),)+
                    _ => Err(TokenError::new($domain, value)),
                }
            }
        }

        impl fmt::Display for $name {
            fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
                formatter.write_str(self.as_str())
            }
        }
    };
}

token_enum! {
    /// Conditional-formatting rule kind (`ST_CfType`).
    pub enum Kind, "conditional-formatting rule kind" {
        Expression => "expression",
        CellIs => "cellIs",
        ColorScale => "colorScale",
        DataBar => "dataBar",
        IconSet => "iconSet",
        Top10 => "top10",
        UniqueValues => "uniqueValues",
        DuplicateValues => "duplicateValues",
        ContainsText => "containsText",
        NotContainsText => "notContainsText",
        BeginsWith => "beginsWith",
        EndsWith => "endsWith",
        ContainsBlanks => "containsBlanks",
        NotContainsBlanks => "notContainsBlanks",
        ContainsErrors => "containsErrors",
        NotContainsErrors => "notContainsErrors",
        TimePeriod => "timePeriod",
        AboveAverage => "aboveAverage",
    }
}

token_enum! {
    /// Cell comparison operator (`ST_ConditionalFormattingOperator`).
    pub enum Operator, "conditional-formatting operator" {
        LessThan => "lessThan",
        LessThanOrEqual => "lessThanOrEqual",
        Equal => "equal",
        NotEqual => "notEqual",
        GreaterThanOrEqual => "greaterThanOrEqual",
        GreaterThan => "greaterThan",
        Between => "between",
        NotBetween => "notBetween",
        ContainsText => "containsText",
        NotContains => "notContains",
        BeginsWith => "beginsWith",
        EndsWith => "endsWith",
    }
}

token_enum! {
    /// Conditional-formatting value-object kind (core and x14 `ST_CfvoType`).
    pub enum ValueKind, "conditional-formatting value kind" {
        Number => "num",
        Percent => "percent",
        Max => "max",
        Min => "min",
        Formula => "formula",
        Percentile => "percentile",
        AutoMin => "autoMin",
        AutoMax => "autoMax",
    }
}

impl ValueKind {
    #[must_use]
    pub const fn is_core(self) -> bool {
        !matches!(self, Self::AutoMin | Self::AutoMax)
    }
}

token_enum! {
    /// Relative date period (`ST_TimePeriod`).
    pub enum Period, "conditional-formatting time period" {
        Today => "today",
        Yesterday => "yesterday",
        Tomorrow => "tomorrow",
        Last7Days => "last7Days",
        ThisMonth => "thisMonth",
        LastMonth => "lastMonth",
        NextMonth => "nextMonth",
        ThisWeek => "thisWeek",
        LastWeek => "lastWeek",
        NextWeek => "nextWeek",
    }
}

token_enum! {
    /// Office 2010 data-bar direction (`x14:ST_DataBarDirection`).
    pub enum Direction, "data-bar direction" {
        Context => "context",
        LeftToRight => "leftToRight",
        RightToLeft => "rightToLeft",
    }
}

token_enum! {
    /// Office 2010 data-bar axis position (`x14:ST_DataBarAxisPosition`).
    pub enum Axis, "data-bar axis position" {
        Automatic => "automatic",
        Middle => "middle",
        None => "none",
    }
}

token_enum! {
    /// Conditional-formatting color element role.
    pub enum ColorRole, "conditional-formatting color role" {
        Color => "color",
        Fill => "fillColor",
        Border => "borderColor",
        NegativeFill => "negativeFillColor",
        NegativeBorder => "negativeBorderColor",
        Axis => "axisColor",
    }
}

token_enum! {
    /// Core SpreadsheetML icon set (`ST_IconSetType`).
    pub enum IconSet, "core icon set" {
        ThreeArrows => "3Arrows",
        ThreeArrowsGray => "3ArrowsGray",
        ThreeFlags => "3Flags",
        ThreeTrafficLights1 => "3TrafficLights1",
        ThreeTrafficLights2 => "3TrafficLights2",
        ThreeSigns => "3Signs",
        ThreeSymbols => "3Symbols",
        ThreeSymbols2 => "3Symbols2",
        FourArrows => "4Arrows",
        FourArrowsGray => "4ArrowsGray",
        FourRedToBlack => "4RedToBlack",
        FourRating => "4Rating",
        FourTrafficLights => "4TrafficLights",
        FiveArrows => "5Arrows",
        FiveArrowsGray => "5ArrowsGray",
        FiveRating => "5Rating",
        FiveQuarters => "5Quarters",
    }
}

impl IconSet {
    #[must_use]
    pub const fn len(self) -> u8 {
        match self {
            Self::ThreeArrows
            | Self::ThreeArrowsGray
            | Self::ThreeFlags
            | Self::ThreeTrafficLights1
            | Self::ThreeTrafficLights2
            | Self::ThreeSigns
            | Self::ThreeSymbols
            | Self::ThreeSymbols2 => 3,
            Self::FourArrows
            | Self::FourArrowsGray
            | Self::FourRedToBlack
            | Self::FourRating
            | Self::FourTrafficLights => 4,
            Self::FiveArrows | Self::FiveArrowsGray | Self::FiveRating | Self::FiveQuarters => 5,
        }
    }

    #[must_use]
    pub const fn is_empty(self) -> bool {
        false
    }
}

token_enum! {
    /// Office 2010 icon set (`x14:ST_IconSetType`).
    pub enum IconSet14, "Office 2010 icon set" {
        ThreeArrows => "3Arrows",
        ThreeArrowsGray => "3ArrowsGray",
        ThreeFlags => "3Flags",
        ThreeTrafficLights1 => "3TrafficLights1",
        ThreeTrafficLights2 => "3TrafficLights2",
        ThreeSigns => "3Signs",
        ThreeSymbols => "3Symbols",
        ThreeSymbols2 => "3Symbols2",
        FourArrows => "4Arrows",
        FourArrowsGray => "4ArrowsGray",
        FourRedToBlack => "4RedToBlack",
        FourRating => "4Rating",
        FourTrafficLights => "4TrafficLights",
        FiveArrows => "5Arrows",
        FiveArrowsGray => "5ArrowsGray",
        FiveRating => "5Rating",
        FiveQuarters => "5Quarters",
        ThreeStars => "3Stars",
        ThreeTriangles => "3Triangles",
        FiveBoxes => "5Boxes",
        NoIcons => "NoIcons",
    }
}

impl IconSet14 {
    #[must_use]
    pub const fn len(self) -> Option<u8> {
        match self {
            Self::ThreeArrows
            | Self::ThreeArrowsGray
            | Self::ThreeFlags
            | Self::ThreeTrafficLights1
            | Self::ThreeTrafficLights2
            | Self::ThreeSigns
            | Self::ThreeSymbols
            | Self::ThreeSymbols2
            | Self::ThreeStars
            | Self::ThreeTriangles => Some(3),
            Self::FourArrows
            | Self::FourArrowsGray
            | Self::FourRedToBlack
            | Self::FourRating
            | Self::FourTrafficLights => Some(4),
            Self::FiveArrows
            | Self::FiveArrowsGray
            | Self::FiveRating
            | Self::FiveQuarters
            | Self::FiveBoxes => Some(5),
            Self::NoIcons => None,
        }
    }

    #[must_use]
    pub const fn is_empty(self) -> bool {
        matches!(self, Self::NoIcons)
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct Color {
    pub rgb: Option<Rgb>,
    pub indexed: Option<u32>,
    pub theme: Option<u32>,
    pub tint: Option<f64>,
    pub automatic: Option<bool>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Value {
    pub kind: ValueKind,
    pub value: Option<String>,
    pub formula: Option<String>,
    pub greater_than_or_equal: bool,
}

#[derive(Debug, Clone, PartialEq)]
pub struct ColorScale {
    pub thresholds: Vec<Value>,
    pub colors: Vec<Color>,
}

#[derive(Debug, Clone, PartialEq)]
pub struct NamedColor {
    pub role: ColorRole,
    pub color: Color,
}

#[derive(Debug, Clone, PartialEq)]
pub struct DataBar {
    pub thresholds: Vec<Value>,
    pub colors: Vec<NamedColor>,
    pub min_length: u32,
    pub max_length: u32,
    pub show_value: bool,
    pub border: bool,
    pub gradient: bool,
    pub direction: Direction,
    pub axis_position: Axis,
}

/// Typed icon-set payload whose vocabulary is carried by `S`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Icons<S> {
    pub set: S,
    pub thresholds: Vec<Value>,
    pub show_value: bool,
    pub percent: bool,
    pub reverse: bool,
}

#[derive(Debug, Clone, PartialEq)]
pub enum Payload {
    ColorScale(ColorScale),
    DataBar(DataBar),
    IconSet(Icons<IconSet>),
    IconSet14(Icons<IconSet14>),
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Component {
    pub(crate) raw_xml: Box<[u8]>,
}

impl Component {
    pub fn raw_xml(&self) -> &[u8] {
        &self.raw_xml
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct NumberFormat {
    pub id: u32,
    pub code: String,
    pub raw_xml: Box<[u8]>,
}

#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Differential {
    pub number_format: Option<NumberFormat>,
    pub font: Option<Component>,
    pub fill: Option<Component>,
    pub border: Option<Component>,
    pub alignment: Option<Component>,
    pub protection: Option<Component>,
    pub extensions: Vec<Component>,
    pub(crate) raw_xml: Box<[u8]>,
}

impl Differential {
    pub fn raw_xml(&self) -> &[u8] {
        &self.raw_xml
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum DifferentialRef {
    StylesIndex(u32),
    Inline(Differential),
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Association {
    Independent,
    EnhancesCore { priority: i32 },
    UnmatchedIgnored,
}

#[derive(Debug, Clone, PartialEq)]
pub struct Rule {
    pub source: Source,
    pub rule_type: Option<Kind>,
    pub priority: Option<i32>,
    pub differential_format: Option<DifferentialRef>,
    pub formulas: SmallVec<[String; 3]>,
    pub stop_if_true: bool,
    pub above_average: bool,
    pub equal_average: bool,
    pub percent: bool,
    pub bottom: bool,
    pub operator: Option<Operator>,
    pub text: Option<String>,
    pub time_period: Option<Period>,
    pub rank: Option<u32>,
    pub standard_deviations: Option<i32>,
    pub payload: Option<Payload>,
    pub extension_id: Option<String>,
    pub extension_association: Association,
}

#[derive(Debug, Clone, PartialEq)]
pub struct Formatting {
    pub ranges: Vec<Range>,
    pub pivot: bool,
    pub rules: Vec<Rule>,
}
