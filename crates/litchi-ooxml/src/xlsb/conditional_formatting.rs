//! Conditional formatting support for XLSB

use crate::xlsb::error::{XlsbError, XlsbResult};
use crate::xlsb::formula::{
    CellParsedFormula, FormulaConverter, FormulaParser, FormulaResolutionContext,
    MAX_CELL_FORMULA_BYTES,
};
use crate::xlsb::utils::cell_reference;

/// Conditional formatting rule type (CFType per MS-XLSB 2.5.18)
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum CfRuleType {
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

impl CfRuleType {
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

/// Conditional formatting value object (CFVO)
#[derive(Debug, Clone, PartialEq)]
pub struct Cfvo {
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

impl Cfvo {
    pub fn new(cfvo_type: u8, value: Option<String>) -> Self {
        let numeric_value = value
            .as_deref()
            .and_then(|value| value.parse().ok())
            .unwrap_or(0.0);
        Cfvo {
            cfvo_type,
            value,
            numeric_value,
            save_greater_than_or_equal: false,
            greater_than_or_equal: true,
            formula_binary: None,
        }
    }

    pub fn parse(data: &[u8]) -> XlsbResult<Self> {
        if data.len() < 24 {
            return Err(XlsbError::InvalidLength {
                expected: 24,
                found: data.len(),
            });
        }
        let context = FormulaResolutionContext::default();
        Self::parse_with_context(data, (0, 0), &context)
    }

    pub(crate) fn parse_with_context(
        data: &[u8],
        base: (u32, u32),
        context: &FormulaResolutionContext,
    ) -> XlsbResult<Self> {
        let mut cursor = CfCursor::new(data, "BrtCFVO");
        let cfvo_type = u8::try_from(cursor.read_u32()?)
            .map_err(|_| invalid("BrtCFVO", "CFVO type overflow"))?;
        if !matches!(cfvo_type, 1 | 2 | 3 | 4 | 5 | 7) {
            return Err(invalid("BrtCFVO", format!("invalid type {cfvo_type}")));
        }
        let numeric_value = cursor.read_f64()?;
        if !numeric_value.is_finite() {
            return Err(invalid("BrtCFVO", "non-finite numeric parameter"));
        }
        if matches!(cfvo_type, 4 | 5) && !(0.0..=100.0).contains(&numeric_value) {
            return Err(invalid(
                "BrtCFVO",
                format!("percentage parameter {numeric_value} outside 0..=100"),
            ));
        }
        let save_greater_than_or_equal = cursor.read_bool32()?;
        let greater_than_or_equal = cursor.read_bool32()?;
        let declared_formula_size = cursor.read_u32()? as usize;
        let formula_binary = if declared_formula_size == 0 {
            None
        } else {
            let formula = cursor.read_formula()?;
            if formula.rgce.len() != declared_formula_size {
                return Err(invalid(
                    "BrtCFVO",
                    "declared formula size does not match token stream",
                ));
            }
            Some(formula)
        };
        cursor.finish()?;
        if matches!(cfvo_type, 2 | 3) && formula_binary.is_some() {
            return Err(invalid("BrtCFVO", "min/max threshold contains a formula"));
        }
        if cfvo_type == 7 && formula_binary.is_none() {
            return Err(invalid("BrtCFVO", "formula threshold omits its formula"));
        }
        let value = if let Some(formula) = &formula_binary {
            Some(render_formula(formula, base, context)?)
        } else if matches!(cfvo_type, 1 | 4 | 5) {
            Some(format_number(numeric_value))
        } else {
            None
        };
        Ok(Self {
            cfvo_type,
            value,
            numeric_value,
            save_greater_than_or_equal,
            greater_than_or_equal,
            formula_binary,
        })
    }
}

/// Lossless XLSB color used by conditional formatting visualizations.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ConditionalFormatColor {
    pub color_type: u8,
    pub index: u8,
    pub tint: i16,
    pub argb: Option<u32>,
    pub(crate) raw: [u8; 8],
}

impl ConditionalFormatColor {
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

    pub fn theme(index: u8, tint: i16) -> XlsbResult<Self> {
        if index > 0x0b {
            return Err(invalid("BrtColor", format!("theme color index {index}")));
        }
        let tint_bytes = tint.to_le_bytes();
        Ok(Self {
            color_type: 3,
            index,
            tint,
            argb: None,
            raw: [6, index, tint_bytes[0], tint_bytes[1], 0, 0, 0, 0],
        })
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

    pub(crate) fn parse(data: &[u8]) -> XlsbResult<Self> {
        if data.len() != 8 {
            return Err(XlsbError::InvalidLength {
                expected: 8,
                found: data.len(),
            });
        }
        let raw: [u8; 8] = data.try_into().map_err(|_| XlsbError::InvalidLength {
            expected: 8,
            found: data.len(),
        })?;
        let color_type = raw[0] >> 1;
        if color_type > 3 {
            return Err(invalid("BrtColor", format!("color type {color_type}")));
        }
        let argb = if color_type == 2 {
            if raw[0] & 1 == 0 {
                return Err(invalid("BrtColor", "direct color is not marked valid"));
            }
            Some(
                (u32::from(raw[7]) << 24)
                    | (u32::from(raw[4]) << 16)
                    | (u32::from(raw[5]) << 8)
                    | u32::from(raw[6]),
            )
        } else {
            None
        };
        if color_type == 3 && raw[1] > 0x0b {
            return Err(invalid("BrtColor", format!("theme color index {}", raw[1])));
        }
        Ok(Self {
            color_type,
            index: raw[1],
            tint: i16::from_le_bytes([raw[2], raw[3]]),
            argb,
            raw,
        })
    }

    pub(crate) fn to_bytes(self) -> XlsbResult<[u8; 8]> {
        if self.color_type > 3 || (self.color_type == 3 && self.index > 0x0b) {
            return Err(invalid("BrtColor", "invalid color type or theme index"));
        }
        if self.color_type == 2 && self.argb.is_none() {
            return Err(invalid("BrtColor", "direct color has no ARGB value"));
        }
        if self.color_type != 2 && self.argb.is_some() {
            return Err(invalid("BrtColor", "non-direct color has an ARGB value"));
        }
        let parsed_raw = Self::parse(&self.raw).ok();
        if parsed_raw.as_ref().is_some_and(|raw| {
            raw.color_type == self.color_type
                && raw.index == self.index
                && raw.tint == self.tint
                && raw.argb == self.argb
        }) {
            return Ok(self.raw);
        }
        let tint = self.tint.to_le_bytes();
        let mut raw = [
            self.color_type << 1,
            self.index,
            tint[0],
            tint[1],
            0,
            0,
            0,
            0,
        ];
        if let Some(argb) = self.argb {
            raw[0] |= 1;
            raw[4] = ((argb >> 16) & 0xff) as u8;
            raw[5] = ((argb >> 8) & 0xff) as u8;
            raw[6] = (argb & 0xff) as u8;
            raw[7] = ((argb >> 24) & 0xff) as u8;
        }
        Ok(raw)
    }
}

/// Color scale conditional formatting
#[derive(Debug, Clone)]
pub struct ColorScale {
    /// Minimum CFVO
    pub min_cfvo: Cfvo,
    /// Middle CFVO (optional for 2-color scale)
    pub mid_cfvo: Option<Cfvo>,
    /// Maximum CFVO
    pub max_cfvo: Cfvo,
    /// Minimum color (ARGB)
    pub min_color: u32,
    /// Middle color (ARGB, optional)
    pub mid_color: Option<u32>,
    /// Maximum color (ARGB)
    pub max_color: u32,
    pub min_color_record: ConditionalFormatColor,
    pub mid_color_record: Option<ConditionalFormatColor>,
    pub max_color_record: ConditionalFormatColor,
}

impl ColorScale {
    pub fn new(min_cfvo: Cfvo, max_cfvo: Cfvo, min_color: u32, max_color: u32) -> Self {
        ColorScale {
            min_cfvo,
            mid_cfvo: None,
            max_cfvo,
            min_color,
            mid_color: None,
            max_color,
            min_color_record: ConditionalFormatColor::from_argb(min_color),
            mid_color_record: None,
            max_color_record: ConditionalFormatColor::from_argb(max_color),
        }
    }

    pub fn with_middle(mut self, mid_cfvo: Cfvo, mid_color: u32) -> Self {
        self.mid_cfvo = Some(mid_cfvo);
        self.mid_color = Some(mid_color);
        self.mid_color_record = Some(ConditionalFormatColor::from_argb(mid_color));
        self
    }
}

/// Data bar conditional formatting
#[derive(Debug, Clone)]
pub struct DataBar {
    /// Minimum CFVO
    pub min_cfvo: Cfvo,
    /// Maximum CFVO
    pub max_cfvo: Cfvo,
    /// Bar color (ARGB)
    pub color: u32,
    /// Show value alongside bar
    pub show_value: bool,
    pub min_length: u8,
    pub max_length: u8,
    pub color_record: ConditionalFormatColor,
}

impl DataBar {
    pub fn new(min_cfvo: Cfvo, max_cfvo: Cfvo, color: u32) -> Self {
        DataBar {
            min_cfvo,
            max_cfvo,
            color,
            show_value: true,
            min_length: 10,
            max_length: 90,
            color_record: ConditionalFormatColor::from_argb(color),
        }
    }
}

/// Icon set conditional formatting
#[derive(Debug, Clone)]
pub struct IconSet {
    /// Icon set type (3Arrows, 3Flags, 3TrafficLights, etc.)
    pub icon_set_type: u8,
    /// CFVOs for thresholds
    pub cfvos: Vec<Cfvo>,
    /// Show values alongside icons
    pub show_value: bool,
    /// Reverse icon order
    pub reverse: bool,
}

impl IconSet {
    pub fn new(icon_set_type: u8, cfvos: Vec<Cfvo>) -> Self {
        IconSet {
            icon_set_type,
            cfvos,
            show_value: true,
            reverse: false,
        }
    }
}

/// Conditional formatting rule
#[derive(Debug, Clone)]
pub struct ConditionalFormattingRule {
    /// Rule type
    pub rule_type: CfRuleType,
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
    /// Color scale (for ColorScale type)
    pub color_scale: Option<ColorScale>,
    /// Data bar (for DataBar type)
    pub data_bar: Option<DataBar>,
    /// Icon set (for IconSet type)
    pub icon_set: Option<IconSet>,
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
}

impl ConditionalFormattingRule {
    pub fn new(rule_type: CfRuleType, priority: u32) -> Self {
        ConditionalFormattingRule {
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
            operator: None,
            parameter: 0,
            template: default_template(rule_type),
            text: None,
            above_average: false,
            bottom: false,
            percent: false,
        }
    }

    pub fn parse(data: &[u8]) -> XlsbResult<Self> {
        let context = FormulaResolutionContext::default();
        Self::parse_with_context(data, (0, 0), &context)
    }

    pub(crate) fn parse_with_context(
        data: &[u8],
        base: (u32, u32),
        context: &FormulaResolutionContext,
    ) -> XlsbResult<Self> {
        let mut cursor = CfCursor::new(data, "BrtBeginCFRule");
        let rule_type_raw = cursor.read_u32()?;
        let rule_type = CfRuleType::from_u32(rule_type_raw).ok_or_else(|| {
            invalid(
                "BrtBeginCFRule",
                format!("invalid rule type {rule_type_raw}"),
            )
        })?;
        let template = cursor.read_u32()?;
        validate_template(rule_type, template)?;
        let raw_dxf = cursor.read_u32()?;
        let dxf_id = (raw_dxf != u32::MAX).then_some(raw_dxf);
        if matches!(
            rule_type,
            CfRuleType::ColorScale | CfRuleType::DataBar | CfRuleType::IconSet
        ) && dxf_id.is_some()
        {
            return Err(invalid(
                "BrtBeginCFRule",
                "visual rule has a differential-format index",
            ));
        }
        let priority = cursor.read_u32()?;
        if priority == 0 || priority > i32::MAX as u32 {
            return Err(invalid(
                "BrtBeginCFRule",
                format!("invalid priority {priority}"),
            ));
        }
        let parameter = cursor.read_u32()?;
        let reserved1 = cursor.read_u32()?;
        let reserved2 = cursor.read_u32()?;
        let flags = cursor.read_u16()?;
        if reserved1 != 0 || reserved2 != 0 || flags & !0x1e != 0 {
            return Err(invalid("BrtBeginCFRule", "reserved fields are nonzero"));
        }
        let stop_if_true = flags & 0x02 != 0;
        let above_average = flags & 0x04 != 0;
        let bottom = flags & 0x08 != 0;
        let percent = flags & 0x10 != 0;
        if matches!(
            rule_type,
            CfRuleType::ColorScale | CfRuleType::DataBar | CfRuleType::IconSet
        ) && stop_if_true
        {
            return Err(invalid("BrtBeginCFRule", "visual rule sets stop-if-true"));
        }
        if rule_type != CfRuleType::TopN && (bottom || percent) {
            return Err(invalid(
                "BrtBeginCFRule",
                "non-filter rule sets bottom/percent flags",
            ));
        }
        validate_parameter_and_flags(
            rule_type,
            template,
            parameter,
            above_average,
            bottom,
            percent,
        )?;
        let declared = [cursor.read_u32()?, cursor.read_u32()?, cursor.read_u32()?];
        let text = cursor.read_nullable_string()?;
        if template == 8 {
            if text
                .as_ref()
                .is_none_or(|text| text.is_empty() || text.encode_utf16().count() > 255)
            {
                return Err(invalid(
                    "BrtBeginCFRule",
                    "contains-text template has an invalid text parameter",
                ));
            }
        } else if text.is_some() {
            return Err(invalid(
                "BrtBeginCFRule",
                "non-text template has a string parameter",
            ));
        }
        let mut formula_slots: [Option<CellParsedFormula>; 3] = [None, None, None];
        for (index, size) in declared.into_iter().enumerate() {
            if size == 0 {
                continue;
            }
            let formula = cursor.read_formula()?;
            if formula.rgce.len() != size as usize {
                return Err(invalid(
                    "BrtBeginCFRule",
                    format!(
                        "formula {} declared {size} token bytes, found {}",
                        index + 1,
                        formula.rgce.len()
                    ),
                ));
            }
            formula_slots[index] = Some(formula);
        }
        cursor.finish()?;
        validate_formula_slots(rule_type, template, parameter, &formula_slots)?;

        let mut formulas = Vec::new();
        let mut formula_extras = Vec::new();
        let mut formula_texts = Vec::new();
        for formula in formula_slots.into_iter().flatten() {
            formulas.push(formula.rgce.clone());
            formula_extras.push(formula.rgcb.clone());
            formula_texts.push(render_formula(&formula, base, context)?);
        }
        let operator = (rule_type == CfRuleType::CellIs)
            .then(|| u8::try_from(parameter).ok())
            .flatten();
        if rule_type == CfRuleType::CellIs && !matches!(operator, Some(1..=8)) {
            return Err(invalid(
                "BrtBeginCFRule",
                format!("invalid cell comparison operator {parameter}"),
            ));
        }

        Ok(ConditionalFormattingRule {
            rule_type,
            dxf_id,
            priority,
            stop_if_true,
            formulas,
            formula_extras,
            formula_texts,
            color_scale: None,
            data_bar: None,
            icon_set: None,
            operator,
            parameter,
            template,
            text,
            above_average,
            bottom,
            percent,
        })
    }
}

/// Conditional formatting for a range
#[derive(Debug, Clone)]
pub struct ConditionalFormatting {
    /// Cell ranges (e.g., "A1:B10")
    pub ranges: Vec<String>,
    /// Rules
    pub rules: Vec<ConditionalFormattingRule>,
    /// Whether the ranges are confined to a PivotTable data area.
    pub pivot_only: bool,
}

impl ConditionalFormatting {
    pub fn new(ranges: Vec<String>) -> Self {
        ConditionalFormatting {
            ranges,
            rules: Vec::new(),
            pivot_only: false,
        }
    }

    pub fn add_rule(&mut self, rule: ConditionalFormattingRule) {
        self.rules.push(rule);
    }
}

pub(crate) fn parse_classic_header(
    data: &[u8],
) -> XlsbResult<(ConditionalFormatting, u32, (u32, u32))> {
    let mut cursor = CfCursor::new(data, "BrtBeginConditionalFormatting");
    let count = cursor.read_u32()?;
    let pivot_only = cursor.read_bool32()?;
    let ranges = cursor.read_ranges(1, 8_192)?;
    cursor.finish()?;
    let base = (ranges[0].0, ranges[0].2);
    let ranges = ranges
        .into_iter()
        .map(|(first_row, last_row, first_col, last_col)| {
            let first = cell_reference(first_row, first_col);
            let last = cell_reference(last_row, last_col);
            if first == last {
                first
            } else {
                format!("{first}:{last}")
            }
        })
        .collect();
    Ok((
        ConditionalFormatting {
            ranges,
            rules: Vec::new(),
            pivot_only,
        },
        count,
        base,
    ))
}

fn default_template(rule_type: CfRuleType) -> u32 {
    match rule_type {
        CfRuleType::CellIs => 0,
        CfRuleType::Expression => 1,
        CfRuleType::ColorScale => 2,
        CfRuleType::DataBar => 3,
        CfRuleType::TopN => 5,
        CfRuleType::IconSet => 4,
    }
}

pub(crate) fn validate_template(rule_type: CfRuleType, template: u32) -> XlsbResult<()> {
    let valid = match rule_type {
        CfRuleType::CellIs => template == 0,
        CfRuleType::Expression => matches!(template, 1 | 7..=12 | 15..=27 | 29 | 30),
        CfRuleType::ColorScale => template == 2,
        CfRuleType::DataBar => template == 3,
        CfRuleType::TopN => template == 5,
        CfRuleType::IconSet => template == 4,
    };
    if valid {
        Ok(())
    } else {
        Err(invalid(
            "BrtBeginCFRule",
            format!("template {template} is invalid for {rule_type:?}"),
        ))
    }
}

pub(crate) fn validate_formula_count(
    rule_type: CfRuleType,
    template: u32,
    parameter: u32,
    count: usize,
) -> XlsbResult<()> {
    let expected = if rule_type == CfRuleType::CellIs {
        if matches!(parameter, 1 | 2) { 2 } else { 1 }
    } else if rule_type == CfRuleType::Expression && matches!(template, 1 | 8..=12 | 15..=24) {
        1
    } else {
        0
    };
    let valid = if matches!(
        rule_type,
        CfRuleType::ColorScale | CfRuleType::DataBar | CfRuleType::IconSet
    ) {
        count <= 1
    } else {
        count == expected
    };
    if valid {
        Ok(())
    } else {
        Err(invalid(
            "BrtBeginCFRule",
            format!("formula count {count} does not match required {expected}"),
        ))
    }
}

fn validate_formula_slots(
    rule_type: CfRuleType,
    template: u32,
    parameter: u32,
    slots: &[Option<CellParsedFormula>; 3],
) -> XlsbResult<()> {
    let expected = if rule_type == CfRuleType::CellIs {
        [true, matches!(parameter, 1 | 2), false]
    } else if rule_type == CfRuleType::Expression && matches!(template, 1 | 8..=12 | 15..=24) {
        [true, false, false]
    } else if matches!(
        rule_type,
        CfRuleType::ColorScale | CfRuleType::DataBar | CfRuleType::IconSet
    ) {
        [false, false, slots[2].is_some()]
    } else {
        [false, false, false]
    };
    let found = slots.each_ref().map(Option::is_some);
    if found == expected {
        Ok(())
    } else {
        Err(invalid(
            "BrtBeginCFRule",
            format!("formula slots {found:?} do not match required {expected:?}"),
        ))
    }
}

fn validate_parameter_and_flags(
    rule_type: CfRuleType,
    template: u32,
    parameter: u32,
    above_average: bool,
    bottom: bool,
    percent: bool,
) -> XlsbResult<()> {
    let valid_parameter = match (rule_type, template) {
        (CfRuleType::CellIs, 0) => (1..=8).contains(&parameter),
        (CfRuleType::Expression, 8) => parameter <= 3,
        (CfRuleType::Expression, 15) => parameter == 0,
        (CfRuleType::Expression, 16) => parameter == 6,
        (CfRuleType::Expression, 17) => parameter == 1,
        (CfRuleType::Expression, 18) => parameter == 2,
        (CfRuleType::Expression, 19) => parameter == 5,
        (CfRuleType::Expression, 20) => parameter == 8,
        (CfRuleType::Expression, 21) => parameter == 3,
        (CfRuleType::Expression, 22) => parameter == 7,
        (CfRuleType::Expression, 23) => parameter == 4,
        (CfRuleType::Expression, 24) => parameter == 9,
        (CfRuleType::Expression, 25 | 26) => parameter < 4,
        (CfRuleType::TopN, 5) if percent => parameter <= 100,
        (CfRuleType::TopN, 5) => (1..=1_000).contains(&parameter),
        _ => parameter == 0,
    };
    if !valid_parameter {
        return Err(invalid(
            "BrtBeginCFRule",
            format!("invalid parameter {parameter} for template {template}"),
        ));
    }
    if above_average != matches!(template, 25 | 29) {
        return Err(invalid(
            "BrtBeginCFRule",
            format!("invalid above-average flag for template {template}"),
        ));
    }
    if rule_type != CfRuleType::TopN && (bottom || percent) {
        return Err(invalid(
            "BrtBeginCFRule",
            "bottom/percent flags are set on a non-filter rule",
        ));
    }
    Ok(())
}

fn render_formula(
    formula: &CellParsedFormula,
    base: (u32, u32),
    context: &FormulaResolutionContext,
) -> XlsbResult<String> {
    let tokens =
        FormulaParser::with_base_cell_and_extra(&formula.rgce, &formula.rgcb, base.0, base.1)
            .parse()?;
    FormulaConverter::try_tokens_to_string_with_context(&tokens, context)
}

fn format_number(value: f64) -> String {
    if value.fract() == 0.0 {
        format!("{value:.0}")
    } else {
        value.to_string()
    }
}

fn invalid(typ: &'static str, val: impl Into<String>) -> XlsbError {
    XlsbError::Unrecognized {
        typ: typ.to_string(),
        val: val.into(),
    }
}

struct CfCursor<'a> {
    data: &'a [u8],
    offset: usize,
    record: &'static str,
}

impl<'a> CfCursor<'a> {
    fn new(data: &'a [u8], record: &'static str) -> Self {
        Self {
            data,
            offset: 0,
            record,
        }
    }

    fn take(&mut self, size: usize) -> XlsbResult<&'a [u8]> {
        let end = self
            .offset
            .checked_add(size)
            .ok_or_else(|| invalid(self.record, "field size overflow"))?;
        let bytes = self
            .data
            .get(self.offset..end)
            .ok_or(XlsbError::InvalidLength {
                expected: end,
                found: self.data.len(),
            })?;
        self.offset = end;
        Ok(bytes)
    }

    fn read_u16(&mut self) -> XlsbResult<u16> {
        let bytes = self.take(2)?;
        Ok(u16::from_le_bytes([bytes[0], bytes[1]]))
    }

    fn read_u32(&mut self) -> XlsbResult<u32> {
        let bytes = self.take(4)?;
        Ok(u32::from_le_bytes(
            bytes.try_into().expect("four-byte field"),
        ))
    }

    fn read_f64(&mut self) -> XlsbResult<f64> {
        let bytes = self.take(8)?;
        Ok(f64::from_le_bytes(
            bytes.try_into().expect("eight-byte field"),
        ))
    }

    fn read_bool32(&mut self) -> XlsbResult<bool> {
        match self.read_u32()? {
            0 => Ok(false),
            1 => Ok(true),
            value => Err(invalid(self.record, format!("invalid Boolean {value}"))),
        }
    }

    fn read_nullable_string(&mut self) -> XlsbResult<Option<String>> {
        let count = self.read_u32()?;
        if count == u32::MAX {
            return Ok(None);
        }
        let count = count as usize;
        let bytes = self.take(
            count
                .checked_mul(2)
                .ok_or_else(|| invalid(self.record, "string size overflow"))?,
        )?;
        let units = bytes
            .chunks_exact(2)
            .map(|pair| u16::from_le_bytes([pair[0], pair[1]]))
            .collect::<Vec<_>>();
        String::from_utf16(&units)
            .map(Some)
            .map_err(|error| XlsbError::Encoding(format!("invalid UTF-16: {error}")))
    }

    fn read_formula(&mut self) -> XlsbResult<CellParsedFormula> {
        let cce = self.read_u32()? as usize;
        if cce == 0 || cce > MAX_CELL_FORMULA_BYTES {
            return Err(XlsbError::InvalidFormula(format!(
                "conditional-format formula length {cce} is outside 1..={MAX_CELL_FORMULA_BYTES}"
            )));
        }
        let rgce = self.take(cce)?.to_vec();
        let cb = self.read_u32()? as usize;
        let rgcb = self.take(cb)?.to_vec();
        Ok(CellParsedFormula { rgce, rgcb })
    }

    fn read_ranges(
        &mut self,
        minimum: usize,
        maximum: usize,
    ) -> XlsbResult<Vec<(u32, u32, u32, u32)>> {
        let raw_count = self.read_u32()? as i32;
        let count = usize::try_from(raw_count)
            .map_err(|_| invalid(self.record, "NULL range collection"))?;
        if !(minimum..=maximum).contains(&count)
            || count > self.data.len().saturating_sub(self.offset) / 16
        {
            return Err(invalid(self.record, format!("invalid range count {count}")));
        }
        let mut ranges = Vec::with_capacity(count);
        for _ in 0..count {
            let first_row = self.read_u32()?;
            let last_row = self.read_u32()?;
            let first_col = self.read_u32()?;
            let last_col = self.read_u32()?;
            if first_row > last_row
                || first_col > last_col
                || last_row >= 1_048_576
                || last_col >= 16_384
            {
                return Err(invalid(self.record, "invalid target range"));
            }
            ranges.push((first_row, last_row, first_col, last_col));
        }
        Ok(ranges)
    }

    fn finish(self) -> XlsbResult<()> {
        if self.offset == self.data.len() {
            Ok(())
        } else {
            Err(XlsbError::InvalidLength {
                expected: self.offset,
                found: self.data.len(),
            })
        }
    }
}

#[cfg(test)]
mod model_tests {
    use super::*;
    use crate::xlsb::formula::FormulaCompiler;

    fn numeric_cfvo_payload(cfvo_type: u32, value: f64) -> Vec<u8> {
        let mut data = Vec::new();
        data.extend_from_slice(&cfvo_type.to_le_bytes());
        data.extend_from_slice(&value.to_le_bytes());
        data.extend_from_slice(&0u32.to_le_bytes());
        data.extend_from_slice(&1u32.to_le_bytes());
        data.extend_from_slice(&0u32.to_le_bytes());
        data
    }

    fn cell_rule_payload(dxf_id: u32, priority: u32, stop: bool, operator: u32) -> Vec<u8> {
        let formula = FormulaCompiler::compile("1").unwrap();
        let mut data = Vec::new();
        data.extend_from_slice(&1u32.to_le_bytes());
        data.extend_from_slice(&0u32.to_le_bytes());
        data.extend_from_slice(&dxf_id.to_le_bytes());
        data.extend_from_slice(&priority.to_le_bytes());
        data.extend_from_slice(&operator.to_le_bytes());
        data.extend_from_slice(&[0; 8]);
        data.extend_from_slice(&(u16::from(stop) << 1).to_le_bytes());
        data.extend_from_slice(&(formula.rgce.len() as u32).to_le_bytes());
        data.extend_from_slice(&0u32.to_le_bytes());
        data.extend_from_slice(&0u32.to_le_bytes());
        data.extend_from_slice(&u32::MAX.to_le_bytes());
        data.extend_from_slice(&formula.to_bytes().unwrap());
        data
    }

    #[test]
    fn test_cf_rule_type_from_u8() {
        assert_eq!(CfRuleType::from_u8(1), Some(CfRuleType::CellIs));
        assert_eq!(CfRuleType::from_u8(2), Some(CfRuleType::Expression));
        assert_eq!(CfRuleType::from_u8(3), Some(CfRuleType::ColorScale));
        assert_eq!(CfRuleType::from_u8(4), Some(CfRuleType::DataBar));
        assert_eq!(CfRuleType::from_u8(5), Some(CfRuleType::TopN));
        assert_eq!(CfRuleType::from_u8(6), Some(CfRuleType::IconSet));
        assert_eq!(CfRuleType::from_u8(0), None);
        assert_eq!(CfRuleType::from_u8(7), None);
        assert_eq!(CfRuleType::from_u8(255), None);
    }

    #[test]
    fn test_cfvo_new() {
        let cfvo = Cfvo::new(1, Some("10".to_string()));
        assert_eq!(cfvo.cfvo_type, 1);
        assert_eq!(cfvo.value, Some("10".to_string()));
    }

    #[test]
    fn test_cfvo_serialize_roundtrip() {
        let parsed = Cfvo::parse(&numeric_cfvo_payload(1, 50.0)).unwrap();
        assert_eq!(parsed.cfvo_type, 1);
        assert_eq!(parsed.value.as_deref(), Some("50"));
        assert_eq!(parsed.numeric_value, 50.0);
    }

    #[test]
    fn test_cfvo_serialize_none_value() {
        let parsed = Cfvo::parse(&numeric_cfvo_payload(2, 0.0)).unwrap();
        assert_eq!(parsed.cfvo_type, 2);
        assert!(parsed.value.is_none());
        assert!(parsed.formula_binary.is_none());
    }

    #[test]
    fn test_cfvo_parse_too_short() {
        let result = Cfvo::parse(&[0x01]);
        assert!(result.is_err());
    }

    #[test]
    fn test_color_scale_new() {
        let min_cfvo = Cfvo::new(2, None); // min
        let max_cfvo = Cfvo::new(3, None); // max
        let cs = ColorScale::new(min_cfvo, max_cfvo, 0xFFFF0000, 0xFF00FF00);

        assert_eq!(cs.min_cfvo.cfvo_type, 2);
        assert_eq!(cs.max_cfvo.cfvo_type, 3);
        assert_eq!(cs.min_color, 0xFFFF0000);
        assert_eq!(cs.max_color, 0xFF00FF00);
        assert!(cs.mid_cfvo.is_none());
        assert!(cs.mid_color.is_none());
    }

    #[test]
    fn test_color_scale_with_middle() {
        let min_cfvo = Cfvo::new(2, None);
        let mid_cfvo = Cfvo::new(1, Some("50".to_string()));
        let max_cfvo = Cfvo::new(3, None);
        let cs = ColorScale::new(min_cfvo, max_cfvo, 0xFFFF0000, 0xFF00FF00)
            .with_middle(mid_cfvo, 0xFFFFFF00);

        assert!(cs.mid_cfvo.is_some());
        assert!(cs.mid_color.is_some());
        assert_eq!(cs.mid_color.unwrap(), 0xFFFFFF00);
    }

    #[test]
    fn test_data_bar_new() {
        let min_cfvo = Cfvo::new(2, None);
        let max_cfvo = Cfvo::new(3, None);
        let db = DataBar::new(min_cfvo, max_cfvo, 0xFF4472C4);

        assert_eq!(db.min_cfvo.cfvo_type, 2);
        assert_eq!(db.max_cfvo.cfvo_type, 3);
        assert_eq!(db.color, 0xFF4472C4);
        assert!(db.show_value);
    }

    #[test]
    fn test_icon_set_new() {
        let cfvos = vec![
            Cfvo::new(1, Some("0".to_string())),
            Cfvo::new(1, Some("33".to_string())),
            Cfvo::new(1, Some("67".to_string())),
        ];
        let icon_set = IconSet::new(0x01, cfvos); // 3Arrows

        assert_eq!(icon_set.icon_set_type, 0x01);
        assert_eq!(icon_set.cfvos.len(), 3);
        assert!(icon_set.show_value);
        assert!(!icon_set.reverse);
    }

    #[test]
    fn test_conditional_formatting_rule_new() {
        let rule = ConditionalFormattingRule::new(CfRuleType::CellIs, 1);

        assert_eq!(rule.rule_type, CfRuleType::CellIs);
        assert_eq!(rule.priority, 1);
        assert!(rule.dxf_id.is_none());
        assert!(!rule.stop_if_true);
        assert!(rule.formulas.is_empty());
        assert!(rule.color_scale.is_none());
        assert!(rule.data_bar.is_none());
        assert!(rule.icon_set.is_none());
        assert!(rule.operator.is_none());
    }

    #[test]
    fn test_conditional_formatting_new() {
        let ranges = vec!["A1:B10".to_string()];
        let cf = ConditionalFormatting::new(ranges);

        assert_eq!(cf.ranges.len(), 1);
        assert_eq!(cf.ranges[0], "A1:B10");
        assert!(cf.rules.is_empty());
    }

    #[test]
    fn test_conditional_formatting_add_rule() {
        let mut cf = ConditionalFormatting::new(vec!["A1:A10".to_string()]);
        let rule = ConditionalFormattingRule::new(CfRuleType::CellIs, 1);
        cf.add_rule(rule);

        assert_eq!(cf.rules.len(), 1);
        assert_eq!(cf.rules[0].rule_type, CfRuleType::CellIs);
    }

    #[test]
    fn test_conditional_formatting_rule_parse() {
        let rule =
            ConditionalFormattingRule::parse(&cell_rule_payload(u32::MAX, 1, false, 5)).unwrap();
        assert_eq!(rule.rule_type, CfRuleType::CellIs);
        assert!(rule.dxf_id.is_none());
        assert_eq!(rule.priority, 1);
        assert!(!rule.stop_if_true);
        assert_eq!(rule.operator, Some(5));
    }

    #[test]
    fn test_conditional_formatting_rule_parse_with_dxf() {
        let rule = ConditionalFormattingRule::parse(&cell_rule_payload(5, 10, true, 3)).unwrap();
        assert_eq!(rule.dxf_id, Some(5));
        assert_eq!(rule.priority, 10);
        assert!(rule.stop_if_true);
    }

    #[test]
    fn test_conditional_formatting_rule_parse_too_short() {
        let data = [0x01, 0x02, 0x03]; // too short
        let result = ConditionalFormattingRule::parse(&data);
        assert!(result.is_err());
    }

    #[test]
    fn test_conditional_formatting_rule_parse_invalid_type() {
        let data = [
            0xFF, // invalid type
            0xFF, 0xFF, 0xFF, 0xFF, 0x01, 0x00, 0x00, 0x00, 0x00,
        ];
        let result = ConditionalFormattingRule::parse(&data);
        assert!(result.is_err());
    }

    #[test]
    fn test_read_optional_string_none() {
        let data = u32::MAX.to_le_bytes();
        let mut cursor = CfCursor::new(&data, "test");
        assert_eq!(cursor.read_nullable_string().unwrap(), None);
        cursor.finish().unwrap();
    }

    #[test]
    fn test_read_optional_string_some() {
        // "Hi" encoded as UTF-16LE with length prefix
        let data = [
            0x02, 0x00, 0x00, 0x00, // length = 2
            0x48, 0x00, // 'H'
            0x69, 0x00, // 'i'
        ];
        let mut cursor = CfCursor::new(&data, "test");
        assert_eq!(
            cursor.read_nullable_string().unwrap().as_deref(),
            Some("Hi")
        );
        cursor.finish().unwrap();
    }

    #[test]
    fn test_read_optional_string_too_short() {
        let data = [0x01]; // too short
        let mut cursor = CfCursor::new(&data, "test");
        assert!(cursor.read_nullable_string().is_err());
    }

    #[test]
    fn test_write_optional_string_none() {
        let data = u32::MAX.to_le_bytes();
        let mut cursor = CfCursor::new(&data, "test");
        assert!(cursor.read_nullable_string().unwrap().is_none());
    }

    #[test]
    fn test_write_optional_string_some() {
        let data = [0x04, 0x00, 0x00, 0x00, b'T', 0, b'e', 0, b's', 0, b't', 0];
        let mut cursor = CfCursor::new(&data, "test");
        assert_eq!(
            cursor.read_nullable_string().unwrap().as_deref(),
            Some("Test")
        );
    }

    #[test]
    fn test_cf_rule_type_variants() {
        // Verify all enum variants have correct discriminant values
        assert_eq!(CfRuleType::CellIs as u8, 1);
        assert_eq!(CfRuleType::Expression as u8, 2);
        assert_eq!(CfRuleType::ColorScale as u8, 3);
        assert_eq!(CfRuleType::DataBar as u8, 4);
        assert_eq!(CfRuleType::TopN as u8, 5);
        assert_eq!(CfRuleType::IconSet as u8, 6);
    }

    #[test]
    fn test_conditional_formatting_clone() {
        let mut cf = ConditionalFormatting::new(vec!["A1:A10".to_string()]);
        let rule = ConditionalFormattingRule::new(CfRuleType::CellIs, 1);
        cf.add_rule(rule);

        let cloned = cf.clone();
        assert_eq!(cloned.ranges.len(), cf.ranges.len());
        assert_eq!(cloned.rules.len(), cf.rules.len());
    }

    #[test]
    fn test_color_scale_clone() {
        let min_cfvo = Cfvo::new(2, None);
        let max_cfvo = Cfvo::new(3, None);
        let cs = ColorScale::new(min_cfvo, max_cfvo, 0xFFFF0000, 0xFF00FF00);
        let cloned = cs.clone();

        assert_eq!(cloned.min_color, cs.min_color);
        assert_eq!(cloned.max_color, cs.max_color);
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::xlsb::formula::FormulaCompiler;

    fn fixture_cell_is_payload() -> Vec<u8> {
        let formula = FormulaCompiler::compile("5").unwrap();
        let mut data = Vec::new();
        data.extend_from_slice(&1u32.to_le_bytes());
        data.extend_from_slice(&0u32.to_le_bytes());
        data.extend_from_slice(&0u32.to_le_bytes());
        data.extend_from_slice(&1u32.to_le_bytes());
        data.extend_from_slice(&5u32.to_le_bytes());
        data.extend_from_slice(&[0; 8]);
        data.extend_from_slice(&0u16.to_le_bytes());
        data.extend_from_slice(&(formula.rgce.len() as u32).to_le_bytes());
        data.extend_from_slice(&0u32.to_le_bytes());
        data.extend_from_slice(&0u32.to_le_bytes());
        data.extend_from_slice(&u32::MAX.to_le_bytes());
        data.extend_from_slice(&formula.to_bytes().unwrap());
        data
    }

    #[test]
    fn parses_normative_cell_is_rule() {
        let rule = ConditionalFormattingRule::parse(&fixture_cell_is_payload()).unwrap();
        assert_eq!(rule.rule_type, CfRuleType::CellIs);
        assert_eq!(rule.template, 0);
        assert_eq!(rule.parameter, 5);
        assert_eq!(rule.operator, Some(5));
        assert_eq!(rule.formula_texts, ["5"]);
        assert_eq!(rule.formulas.len(), 1);
        assert_eq!(rule.formula_extras, [Vec::<u8>::new()]);
    }

    #[test]
    fn rejects_formula_in_wrong_slot_or_with_wrong_declared_size() {
        let mut wrong_slot = fixture_cell_is_payload();
        let size = wrong_slot[30..34].to_vec();
        wrong_slot[30..34].fill(0);
        wrong_slot[34..38].copy_from_slice(&size);
        assert!(ConditionalFormattingRule::parse(&wrong_slot).is_err());

        let mut wrong_size = fixture_cell_is_payload();
        wrong_size[30..34].copy_from_slice(&4u32.to_le_bytes());
        assert!(ConditionalFormattingRule::parse(&wrong_size).is_err());
    }

    #[test]
    fn parses_cfvo_with_ancillary_formula_losslessly() {
        let formula = FormulaCompiler::compile("{1,2}").unwrap();
        assert!(!formula.rgcb.is_empty());
        let mut data = Vec::new();
        data.extend_from_slice(&7u32.to_le_bytes());
        data.extend_from_slice(&0f64.to_le_bytes());
        data.extend_from_slice(&0u32.to_le_bytes());
        data.extend_from_slice(&1u32.to_le_bytes());
        data.extend_from_slice(&(formula.rgce.len() as u32).to_le_bytes());
        data.extend_from_slice(&formula.to_bytes().unwrap());
        let parsed = Cfvo::parse(&data).unwrap();
        assert_eq!(parsed.formula_binary.as_ref().unwrap(), &formula);
    }

    #[test]
    fn parses_direct_and_theme_colors() {
        let direct = ConditionalFormatColor::parse(&[5, 0, 0, 0, 0x11, 0x22, 0x33, 0xff]).unwrap();
        assert_eq!(direct.argb, Some(0xff11_2233));
        let theme = ConditionalFormatColor::parse(&[6, 4, 0, 0, 0, 0, 0, 0]).unwrap();
        assert_eq!(theme.color_type, 3);
        assert_eq!(theme.index, 4);
        assert!(ConditionalFormatColor::parse(&[6, 12, 0, 0, 0, 0, 0, 0]).is_err());

        let theme = ConditionalFormatColor::theme(5, -1_000).unwrap();
        assert_eq!(
            ConditionalFormatColor::parse(&theme.to_bytes().unwrap()).unwrap(),
            theme
        );
        let mut indexed = ConditionalFormatColor::indexed(42, 0);
        indexed.tint = 2_000;
        let reparsed = ConditionalFormatColor::parse(&indexed.to_bytes().unwrap()).unwrap();
        assert_eq!(reparsed.index, 42);
        assert_eq!(reparsed.tint, 2_000);
    }

    #[test]
    fn parses_classic_header_with_pivot_and_range() {
        let mut data = Vec::new();
        data.extend_from_slice(&1u32.to_le_bytes());
        data.extend_from_slice(&1u32.to_le_bytes());
        data.extend_from_slice(&1u32.to_le_bytes());
        data.extend_from_slice(&0u32.to_le_bytes());
        data.extend_from_slice(&1u32.to_le_bytes());
        data.extend_from_slice(&0u32.to_le_bytes());
        data.extend_from_slice(&1u32.to_le_bytes());
        let (formatting, count, base) = parse_classic_header(&data).unwrap();
        assert_eq!(count, 1);
        assert!(formatting.pivot_only);
        assert_eq!(formatting.ranges, ["A1:B2"]);
        assert_eq!(base, (0, 0));
    }
}
