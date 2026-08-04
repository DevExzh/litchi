//! Package-neutral conditional-formatting values for XLSB.
//!
//! The canonical vocabulary is deliberately concise. Binary framing, formula
//! resolution, and record validation remain in `codec.rs`; worksheet/package
//! traversal remains in the host.

use crate::formula::CellParsedFormula;

/// Conditional formatting rule type (CFType per MS-XLSB 2.5.18)
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum RuleType {
    /// CF_TYPE_CELLIS = 1: Cell value comparison
    CellIs = 1,
    /// CF_TYPE_EXPRIS = 2: Expression evaluation
    Expression = 2,
    /// CF_TYPE_GRADIENT = 3: Color scale (2-3 colors)
    ColorScale = 3,
    /// CF_TYPE_DATABAR = 4: Data bar
    DataBar = 4,
    /// CF_TYPE_FILTER = 5: Top/bottom N values
    TopN = 5,
    /// CF_TYPE_MULTISTATE = 6: Icon set
    IconSet = 6,
}

/// Binary record family used by a conditional-formatting collection.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum RecordKind {
    /// Original XLSB conditional-formatting records.
    #[default]
    Classic,
    /// Office 2013 future-record conditional-formatting records.
    Extension14,
}

/// Fields unique to an Office 2013 `BrtBeginCFRule14` record.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct RuleMetadata {
    /// Signed priority. `-1` denotes an extension of a classic data-bar rule.
    pub priority: i32,
    /// Undefined field preserved for lossless roundtrips.
    pub unused: u32,
    /// Raw GUID bytes in MS-DTYP wire order.
    pub guid: [u8; 16],
    /// Whether `guid` is semantically present.
    pub guid_present: bool,
    /// Priority of the classic rule resolved through `BrtCFRuleExt`, if any.
    pub linked_classic_priority: Option<u32>,
}

/// Conditional formatting value object (CFVO)
#[derive(Debug, Clone, PartialEq)]
pub struct Value {
    /// Type: 1=num, 2=min, 3=max, 4=percent, 5=percentile, 7=formula.
    pub cfvo_type: u8,
    /// Human-readable numeric value or formula.
    pub value: Option<String>,
    /// Stored numeric parameter when the CFVO does not use a formula.
    pub numeric_value: f64,
    /// Whether the greater-than/equal flag is meaningful (icon sets).
    pub save_greater_than_or_equal: bool,
    /// Whether threshold comparison is greater-than-or-equal.
    pub greater_than_or_equal: bool,
    /// Original binary formula including ancillary data.
    pub formula_binary: Option<CellParsedFormula>,
}

/// Lossless XLSB color used by conditional formatting visualizations.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Color {
    pub color_type: u8,
    pub index: u8,
    pub tint: i16,
    pub argb: Option<u32>,
    pub(crate) raw: [u8; 8],
}

/// Color scale conditional formatting
#[derive(Debug, Clone)]
pub struct Scale {
    /// Minimum CFVO
    pub min_cfvo: Value,
    /// Middle CFVO (optional for 2-color scale)
    pub mid_cfvo: Option<Value>,
    /// Maximum CFVO
    pub max_cfvo: Value,
    /// Minimum color (ARGB)
    pub min_color: u32,
    /// Middle color (ARGB, optional)
    pub mid_color: Option<u32>,
    /// Maximum color (ARGB)
    pub max_color: u32,
    pub min_color_record: Color,
    pub mid_color_record: Option<Color>,
    pub max_color_record: Color,
}

/// Data bar conditional formatting
#[derive(Debug, Clone)]
pub struct Bar {
    /// Minimum CFVO
    pub min_cfvo: Value,
    /// Maximum CFVO
    pub max_cfvo: Value,
    /// Bar color (ARGB)
    pub color: u32,
    /// Show value alongside bar
    pub show_value: bool,
    pub min_length: u8,
    pub max_length: u8,
    pub color_record: Color,
}

/// Icon set conditional formatting
#[derive(Debug, Clone)]
pub struct IconSet {
    /// Icon set type (3Arrows, 3Flags, 3TrafficLights, etc.)
    pub icon_set_type: u8,
    /// CFVOs for thresholds
    pub cfvos: Vec<Value>,
    /// Show values alongside icons
    pub show_value: bool,
    /// Reverse icon order
    pub reverse: bool,
}

/// Direction of an Office 2013 data bar.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
#[repr(u8)]
pub enum Direction14 {
    /// Resolve direction from worksheet context.
    #[default]
    Context = 0,
    LeftToRight = 1,
    RightToLeft = 2,
}

/// Axis placement of an Office 2013 data bar.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
#[repr(u8)]
pub enum AxisPosition14 {
    #[default]
    Automatic = 0,
    Midpoint = 1,
    None = 2,
}

/// Office 2013 extended data-bar visualization.
#[derive(Debug, Clone, PartialEq)]
pub struct Bar14 {
    pub min_cfvo: Value,
    pub max_cfvo: Value,
    /// Absent only when this augments a classic data-bar rule (`iPri = -1`).
    pub positive_color: Option<Color>,
    pub border_color: Option<Color>,
    pub negative_color: Option<Color>,
    pub negative_border_color: Option<Color>,
    pub axis_color: Option<Color>,
    pub min_length: u8,
    pub max_length: u8,
    pub show_value: bool,
    pub direction: Direction14,
    pub axis_position: AxisPosition14,
    pub border: bool,
    pub gradient: bool,
    pub custom_negative_fill: bool,
    pub custom_negative_border: bool,
    /// Undefined upper flag bits preserved for lossless roundtrips.
    pub unused_flags: u16,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct BarHeader14 {
    pub min_length: u8,
    pub max_length: u8,
    pub show_value: bool,
    pub direction: Direction14,
    pub axis_position: AxisPosition14,
    pub border: bool,
    pub gradient: bool,
    pub custom_negative_fill: bool,
    pub custom_negative_border: bool,
    pub unused_flags: u16,
}

/// One custom icon in an Office 2013 icon-set rule.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Icon {
    /// Icon-set identifier, or `-1` for no icon.
    pub icon_set: i32,
    /// Zero-based icon index, or `-1` when `icon_set` is `-1`.
    pub index: i32,
}

/// Office 2013 extended icon-set visualization.
#[derive(Debug, Clone, PartialEq)]
pub struct IconSet14 {
    pub icon_set_type: u8,
    pub cfvos: Vec<Value>,
    /// Present only for a custom icon set and one-for-one with `cfvos`.
    pub custom_icons: Option<Vec<Icon>>,
    pub show_value: bool,
    pub reverse: bool,
    /// Undefined flag bits 3 through 6 preserved for lossless roundtrips.
    pub unused_flags: u16,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct IconHeader14 {
    pub icon_set_type: u8,
    pub custom: bool,
    pub show_value: bool,
    pub reverse: bool,
    pub unused_flags: u16,
}

/// Conditional formatting rule
#[derive(Debug, Clone)]
pub struct Rule {
    /// Rule type
    pub rule_type: RuleType,
    /// DXF index (differential formatting)
    pub dxf_id: Option<u32>,
    /// Priority (lower = higher priority)
    pub priority: u32,
    /// Stop if true
    pub stop_if_true: bool,
    /// Formula(s) for the rule (binary PTG tokens)
    pub formulas: Vec<Vec<u8>>,
    /// Ancillary streams corresponding one-for-one with `formulas`.
    pub formula_extras: Vec<Vec<u8>>,
    /// Human-readable formulas, compiled when binary formulas are absent.
    pub formula_texts: Vec<String>,
    /// Color scale (for Scale type)
    pub color_scale: Option<Scale>,
    /// Data bar (for Bar type)
    pub data_bar: Option<Bar>,
    /// Icon set (for IconSet type)
    pub icon_set: Option<IconSet>,
    /// Office 2013 color scale.
    pub color_scale14: Option<Scale>,
    /// Office 2013 data bar.
    pub data_bar14: Option<Bar14>,
    /// Office 2013 icon set.
    pub icon_set14: Option<IconSet14>,
    /// Operator (for CellIs type): 1=between, 2=not between, 3=equal, etc.
    pub operator: Option<u8>,
    /// Exact 32-bit rule parameter (operator, rank, date operation, or standard deviation).
    pub parameter: u32,
    /// Exact `CFTemp` template identifier.
    pub template: u32,
    /// String parameter used by contains-text templates.
    pub text: Option<String>,
    pub above_average: bool,
    pub bottom: bool,
    pub percent: bool,
    /// Office 2013 record metadata when this rule came from `BrtBeginCFRule14`.
    pub extension14: Option<RuleMetadata>,
    /// GUID linking a classic rule to an Office 2013 data-bar augmentation.
    pub classic_extension_guid: Option<[u8; 16]>,
}

/// Conditional formatting for a range
#[derive(Debug, Clone)]
pub struct Formatting {
    /// Cell ranges (e.g., "A1:B10")
    pub ranges: Vec<String>,
    /// Rules
    pub rules: Vec<Rule>,
    /// Whether the ranges are confined to a PivotTable data area.
    pub pivot_only: bool,
    /// Binary record family used to encode this collection.
    pub record_kind: RecordKind,
}

impl RuleType {
    pub fn from_u32(value: u32) -> Option<Self> {
        match value {
            1 => Some(Self::CellIs),
            2 => Some(Self::Expression),
            3 => Some(Self::ColorScale),
            4 => Some(Self::DataBar),
            5 => Some(Self::TopN),
            6 => Some(Self::IconSet),
            _ => None,
        }
    }

    pub fn from_u8(value: u8) -> Option<Self> {
        Self::from_u32(u32::from(value))
    }
}

fn default_template(rule_type: RuleType) -> u32 {
    match rule_type {
        RuleType::CellIs => 0,
        RuleType::Expression => 1,
        RuleType::ColorScale => 2,
        RuleType::DataBar => 3,
        RuleType::TopN => 5,
        RuleType::IconSet => 4,
    }
}

impl Value {
    pub fn new(cfvo_type: u8, value: Option<String>) -> Self {
        let numeric_value = value
            .as_deref()
            .and_then(|value| value.parse().ok())
            .unwrap_or(0.0);
        Value {
            cfvo_type,
            value,
            numeric_value,
            save_greater_than_or_equal: false,
            greater_than_or_equal: true,
            formula_binary: None,
        }
    }
}

impl Color {
    pub fn automatic() -> Self {
        Self {
            color_type: 0,
            index: 0,
            tint: 0,
            argb: None,
            raw: [0; 8],
        }
    }

    pub fn indexed(index: u8, tint: i16) -> Self {
        let tint_bytes = tint.to_le_bytes();
        Self {
            color_type: 1,
            index,
            tint,
            argb: None,
            raw: [2, index, tint_bytes[0], tint_bytes[1], 0, 0, 0, 0],
        }
    }

    pub fn from_argb(argb: u32) -> Self {
        let raw = [
            5,
            0,
            0,
            0,
            ((argb >> 16) & 0xff) as u8,
            ((argb >> 8) & 0xff) as u8,
            (argb & 0xff) as u8,
            ((argb >> 24) & 0xff) as u8,
        ];
        Self {
            color_type: 2,
            index: 0,
            tint: 0,
            argb: Some(argb),
            raw,
        }
    }
}

impl Scale {
    pub fn new(min_cfvo: Value, max_cfvo: Value, min_color: u32, max_color: u32) -> Self {
        Scale {
            min_cfvo,
            mid_cfvo: None,
            max_cfvo,
            min_color,
            mid_color: None,
            max_color,
            min_color_record: Color::from_argb(min_color),
            mid_color_record: None,
            max_color_record: Color::from_argb(max_color),
        }
    }

    pub fn with_middle(mut self, mid_cfvo: Value, mid_color: u32) -> Self {
        self.mid_cfvo = Some(mid_cfvo);
        self.mid_color = Some(mid_color);
        self.mid_color_record = Some(Color::from_argb(mid_color));
        self
    }
}

impl Bar {
    pub fn new(min_cfvo: Value, max_cfvo: Value, color: u32) -> Self {
        Bar {
            min_cfvo,
            max_cfvo,
            color,
            show_value: true,
            min_length: 10,
            max_length: 90,
            color_record: Color::from_argb(color),
        }
    }
}

impl IconSet {
    pub fn new(icon_set_type: u8, cfvos: Vec<Value>) -> Self {
        IconSet {
            icon_set_type,
            cfvos,
            show_value: true,
            reverse: false,
        }
    }
}

impl Bar14 {
    pub fn new(min_cfvo: Value, max_cfvo: Value, positive_color: Color) -> Self {
        Self {
            min_cfvo,
            max_cfvo,
            positive_color: Some(positive_color),
            border_color: None,
            negative_color: None,
            negative_border_color: None,
            axis_color: Some(Color::automatic()),
            min_length: 10,
            max_length: 90,
            show_value: true,
            direction: Direction14::Context,
            axis_position: AxisPosition14::Automatic,
            border: false,
            gradient: true,
            custom_negative_fill: false,
            custom_negative_border: false,
            unused_flags: 0,
        }
    }
}

impl IconSet14 {
    pub fn new(icon_set_type: u8, cfvos: Vec<Value>) -> Self {
        Self {
            icon_set_type,
            cfvos,
            custom_icons: None,
            show_value: true,
            reverse: false,
            unused_flags: 0,
        }
    }
}

impl Rule {
    pub fn new(rule_type: RuleType, priority: u32) -> Self {
        Rule {
            rule_type,
            dxf_id: None,
            priority,
            stop_if_true: false,
            formulas: Vec::new(),
            formula_extras: Vec::new(),
            formula_texts: Vec::new(),
            color_scale: None,
            data_bar: None,
            icon_set: None,
            color_scale14: None,
            data_bar14: None,
            icon_set14: None,
            operator: None,
            parameter: 0,
            template: default_template(rule_type),
            text: None,
            above_average: false,
            bottom: false,
            percent: false,
            extension14: None,
            classic_extension_guid: None,
        }
    }
}

impl Formatting {
    pub fn new(ranges: Vec<String>) -> Self {
        Formatting {
            ranges,
            rules: Vec::new(),
            pivot_only: false,
            record_kind: RecordKind::Classic,
        }
    }

    /// Create an Office 2013 conditional-formatting collection.
    pub fn new_extension14(ranges: Vec<String>) -> Self {
        Self {
            ranges,
            rules: Vec::new(),
            pivot_only: false,
            record_kind: RecordKind::Extension14,
        }
    }

    pub fn add_rule(&mut self, rule: Rule) {
        self.rules.push(rule);
    }
}
