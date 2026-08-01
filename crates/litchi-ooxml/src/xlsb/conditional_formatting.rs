//! Conditional formatting support for XLSB

use crate::xlsb::error::{XlsbError, XlsbResult};
use crate::xlsb::formula::{
    CellParsedFormula, FormulaConverter, FormulaParser, FormulaResolutionContext,
    MAX_CELL_FORMULA_BYTES,
};
use crate::xlsb::frt::{
    parse_formula_header, parse_sqref_header, serialize_formula_header, serialize_sqref_header,
};
use crate::xlsb::utils::{cell_reference, parse_cell_reference};

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

/// Binary record family used by a conditional-formatting collection.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum ConditionalFormattingRecordKind {
    /// Original XLSB conditional-formatting records.
    #[default]
    Classic,
    /// Office 2013 future-record conditional-formatting records.
    Extension14,
}

/// Fields unique to an Office 2013 `BrtBeginCFRule14` record.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ConditionalFormattingRule14Metadata {
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

    /// Parse an Office 2013 `BrtCFVO14` record.
    pub fn parse_extension14(data: &[u8]) -> XlsbResult<Self> {
        let context = FormulaResolutionContext::default();
        Self::parse_extension14_with_context(data, (0, 0), &context)
    }

    pub(crate) fn parse_extension14_with_context(
        data: &[u8],
        base: (u32, u32),
        context: &FormulaResolutionContext,
    ) -> XlsbResult<Self> {
        let (formulas, header_size) = parse_formula_header(data, "BrtCFVO14", 1)?;
        let mut cursor = CfCursor::new(&data[header_size..], "BrtCFVO14");
        let cfvo_type = u8::try_from(cursor.read_u32()?)
            .map_err(|_| invalid("BrtCFVO14", "CFVO type overflow"))?;
        if !matches!(cfvo_type, 1 | 2 | 3 | 4 | 5 | 7 | 8 | 9) {
            return Err(invalid("BrtCFVO14", format!("invalid type {cfvo_type}")));
        }
        let numeric_value = cursor.read_f64()?;
        if !numeric_value.is_finite() {
            return Err(invalid("BrtCFVO14", "non-finite numeric parameter"));
        }
        let save_greater_than_or_equal = cursor.read_bool32()?;
        let greater_than_or_equal = cursor.read_bool32()?;
        let declared_formula_size = cursor.read_u32()? as usize;
        cursor.finish()?;
        let formula_binary = formulas.into_iter().next();
        if formula_binary
            .as_ref()
            .map_or(0, |formula| formula.rgce.len())
            != declared_formula_size
        {
            return Err(invalid(
                "BrtCFVO14",
                "FRT formula and declared token size disagree",
            ));
        }
        if matches!(cfvo_type, 2 | 3 | 8 | 9) && formula_binary.is_some() {
            return Err(invalid(
                "BrtCFVO14",
                "automatic/min/max threshold contains a formula",
            ));
        }
        if cfvo_type == 7 && formula_binary.is_none() {
            return Err(invalid("BrtCFVO14", "formula threshold omits its formula"));
        }
        if formula_binary.is_none()
            && matches!(cfvo_type, 4 | 5)
            && !(0.0..=100.0).contains(&numeric_value)
        {
            return Err(invalid(
                "BrtCFVO14",
                format!("percentage parameter {numeric_value} outside 0..=100"),
            ));
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

    /// Serialize an Office 2013 `BrtCFVO14` payload using its binary formula.
    pub fn serialize_extension14(&self) -> XlsbResult<Vec<u8>> {
        if !matches!(self.cfvo_type, 1 | 2 | 3 | 4 | 5 | 7 | 8 | 9) {
            return Err(invalid(
                "BrtCFVO14",
                format!("invalid type {}", self.cfvo_type),
            ));
        }
        if !self.numeric_value.is_finite() {
            return Err(invalid("BrtCFVO14", "non-finite numeric parameter"));
        }
        if self.formula_binary.is_none()
            && matches!(self.cfvo_type, 4 | 5)
            && !(0.0..=100.0).contains(&self.numeric_value)
        {
            return Err(invalid(
                "BrtCFVO14",
                format!(
                    "percentage parameter {} outside 0..=100",
                    self.numeric_value
                ),
            ));
        }
        if matches!(self.cfvo_type, 2 | 3 | 8 | 9) && self.formula_binary.is_some() {
            return Err(invalid(
                "BrtCFVO14",
                "automatic/min/max threshold contains a formula",
            ));
        }
        if self.cfvo_type == 7 && self.formula_binary.is_none() {
            return Err(invalid("BrtCFVO14", "formula threshold omits its formula"));
        }
        let formulas = self.formula_binary.as_slice();
        let mut data = serialize_formula_header(formulas, 1)?;
        data.extend_from_slice(&u32::from(self.cfvo_type).to_le_bytes());
        data.extend_from_slice(&self.numeric_value.to_le_bytes());
        data.extend_from_slice(&u32::from(self.save_greater_than_or_equal).to_le_bytes());
        data.extend_from_slice(&u32::from(self.greater_than_or_equal).to_le_bytes());
        data.extend_from_slice(
            &u32::try_from(
                self.formula_binary
                    .as_ref()
                    .map_or(0, |formula| formula.rgce.len()),
            )
            .map_err(|_| XlsbError::InvalidFormula("formula is too large".to_string()))?
            .to_le_bytes(),
        );
        Ok(data)
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

    /// Parse an Office 2013 `BrtColor14` payload.
    pub fn parse_extension14(data: &[u8]) -> XlsbResult<Self> {
        if data.len() != 12 {
            return Err(XlsbError::InvalidLength {
                expected: 12,
                found: data.len(),
            });
        }
        if data[..4] != [0; 4] {
            return Err(invalid("BrtColor14", "nonzero FRTBlank"));
        }
        Self::parse(&data[4..])
    }

    /// Serialize an Office 2013 `BrtColor14` payload.
    pub fn serialize_extension14(self) -> XlsbResult<[u8; 12]> {
        let mut data = [0; 12];
        data[4..].copy_from_slice(&self.to_bytes()?);
        Ok(data)
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

/// Direction of an Office 2013 data bar.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
#[repr(u8)]
pub enum DataBarDirection14 {
    /// Resolve direction from worksheet context.
    #[default]
    Context = 0,
    LeftToRight = 1,
    RightToLeft = 2,
}

impl DataBarDirection14 {
    fn from_u8(value: u8) -> Option<Self> {
        match value {
            0 => Some(Self::Context),
            1 => Some(Self::LeftToRight),
            2 => Some(Self::RightToLeft),
            _ => None,
        }
    }
}

/// Axis placement of an Office 2013 data bar.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
#[repr(u8)]
pub enum DataBarAxisPosition14 {
    #[default]
    Automatic = 0,
    Midpoint = 1,
    None = 2,
}

impl DataBarAxisPosition14 {
    fn from_u8(value: u8) -> Option<Self> {
        match value {
            0 => Some(Self::Automatic),
            1 => Some(Self::Midpoint),
            2 => Some(Self::None),
            _ => None,
        }
    }
}

/// Office 2013 extended data-bar visualization.
#[derive(Debug, Clone, PartialEq)]
pub struct DataBar14 {
    pub min_cfvo: Cfvo,
    pub max_cfvo: Cfvo,
    /// Absent only when this augments a classic data-bar rule (`iPri = -1`).
    pub positive_color: Option<ConditionalFormatColor>,
    pub border_color: Option<ConditionalFormatColor>,
    pub negative_color: Option<ConditionalFormatColor>,
    pub negative_border_color: Option<ConditionalFormatColor>,
    pub axis_color: Option<ConditionalFormatColor>,
    pub min_length: u8,
    pub max_length: u8,
    pub show_value: bool,
    pub direction: DataBarDirection14,
    pub axis_position: DataBarAxisPosition14,
    pub border: bool,
    pub gradient: bool,
    pub custom_negative_fill: bool,
    pub custom_negative_border: bool,
    /// Undefined upper flag bits preserved for lossless roundtrips.
    pub unused_flags: u16,
}

impl DataBar14 {
    pub fn new(min_cfvo: Cfvo, max_cfvo: Cfvo, positive_color: ConditionalFormatColor) -> Self {
        Self {
            min_cfvo,
            max_cfvo,
            positive_color: Some(positive_color),
            border_color: None,
            negative_color: None,
            negative_border_color: None,
            axis_color: Some(ConditionalFormatColor::automatic()),
            min_length: 10,
            max_length: 90,
            show_value: true,
            direction: DataBarDirection14::Context,
            axis_position: DataBarAxisPosition14::Automatic,
            border: false,
            gradient: true,
            custom_negative_fill: false,
            custom_negative_border: false,
            unused_flags: 0,
        }
    }

    pub fn parse_header(data: &[u8]) -> XlsbResult<DataBar14Header> {
        let mut cursor = CfCursor::new(data, "BrtBeginDatabar14");
        if cursor.read_u32()? != 0 {
            return Err(invalid("BrtBeginDatabar14", "nonzero FRTBlank"));
        }
        let min_length = cursor.read_u8()?;
        let max_length = cursor.read_u8()?;
        let show_value = cursor.read_bool8()?;
        let direction = DataBarDirection14::from_u8(cursor.read_u8()?)
            .ok_or_else(|| invalid("BrtBeginDatabar14", "invalid direction"))?;
        let axis_position = DataBarAxisPosition14::from_u8(cursor.read_u8()?)
            .ok_or_else(|| invalid("BrtBeginDatabar14", "invalid axis position"))?;
        let flags = cursor.read_u16()?;
        cursor.finish()?;
        if min_length > max_length || max_length > 100 {
            return Err(invalid(
                "BrtBeginDatabar14",
                "invalid minimum/maximum length",
            ));
        }
        Ok(DataBar14Header {
            min_length,
            max_length,
            show_value,
            direction,
            axis_position,
            border: flags & 0x01 != 0,
            gradient: flags & 0x02 != 0,
            custom_negative_fill: flags & 0x04 != 0,
            custom_negative_border: flags & 0x08 != 0,
            unused_flags: flags & 0xfff0,
        })
    }

    pub fn serialize_header(&self) -> XlsbResult<Vec<u8>> {
        if self.min_length > self.max_length
            || self.max_length > 100
            || self.unused_flags & 0x0f != 0
        {
            return Err(invalid("BrtBeginDatabar14", "invalid data-bar header"));
        }
        let mut flags = self.unused_flags;
        flags |= u16::from(self.border);
        flags |= u16::from(self.gradient) << 1;
        flags |= u16::from(self.custom_negative_fill) << 2;
        flags |= u16::from(self.custom_negative_border) << 3;
        let mut data = Vec::with_capacity(11);
        data.extend_from_slice(&0u32.to_le_bytes());
        data.extend_from_slice(&[
            self.min_length,
            self.max_length,
            u8::from(self.show_value),
            self.direction as u8,
            self.axis_position as u8,
        ]);
        data.extend_from_slice(&flags.to_le_bytes());
        Ok(data)
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DataBar14Header {
    pub min_length: u8,
    pub max_length: u8,
    pub show_value: bool,
    pub direction: DataBarDirection14,
    pub axis_position: DataBarAxisPosition14,
    pub border: bool,
    pub gradient: bool,
    pub custom_negative_fill: bool,
    pub custom_negative_border: bool,
    pub unused_flags: u16,
}

/// One custom icon in an Office 2013 icon-set rule.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ConditionalFormatIcon {
    /// Icon-set identifier, or `-1` for no icon.
    pub icon_set: i32,
    /// Zero-based icon index, or `-1` when `icon_set` is `-1`.
    pub index: i32,
}

impl ConditionalFormatIcon {
    pub fn parse(data: &[u8]) -> XlsbResult<Self> {
        let mut cursor = CfCursor::new(data, "BrtCFIcon");
        if cursor.read_u32()? != 0 {
            return Err(invalid("BrtCFIcon", "nonzero FRTBlank"));
        }
        let value = Self {
            icon_set: cursor.read_i32()?,
            index: cursor.read_i32()?,
        };
        cursor.finish()?;
        value.validate()?;
        Ok(value)
    }

    pub fn serialize(self) -> XlsbResult<[u8; 12]> {
        self.validate()?;
        let mut data = [0; 12];
        data[4..8].copy_from_slice(&self.icon_set.to_le_bytes());
        data[8..].copy_from_slice(&self.index.to_le_bytes());
        Ok(data)
    }

    fn validate(self) -> XlsbResult<()> {
        if self.icon_set == -1 {
            if self.index == -1 {
                return Ok(());
            }
        } else if let Ok(icon_set) = u8::try_from(self.icon_set)
            && icon_set <= 19
            && (0..icon_count14(icon_set) as i32).contains(&self.index)
        {
            return Ok(());
        }
        Err(invalid("BrtCFIcon", "invalid icon set or index"))
    }
}

/// Office 2013 extended icon-set visualization.
#[derive(Debug, Clone, PartialEq)]
pub struct IconSet14 {
    pub icon_set_type: u8,
    pub cfvos: Vec<Cfvo>,
    /// Present only for a custom icon set and one-for-one with `cfvos`.
    pub custom_icons: Option<Vec<ConditionalFormatIcon>>,
    pub show_value: bool,
    pub reverse: bool,
    /// Undefined flag bits 3 through 6 preserved for lossless roundtrips.
    pub unused_flags: u16,
}

impl IconSet14 {
    pub fn new(icon_set_type: u8, cfvos: Vec<Cfvo>) -> Self {
        Self {
            icon_set_type,
            cfvos,
            custom_icons: None,
            show_value: true,
            reverse: false,
            unused_flags: 0,
        }
    }

    pub fn parse_header(data: &[u8]) -> XlsbResult<IconSet14Header> {
        let mut cursor = CfCursor::new(data, "BrtBeginIconSet14");
        if cursor.read_u32()? != 0 {
            return Err(invalid("BrtBeginIconSet14", "nonzero FRTBlank"));
        }
        let icon_set_type = u8::try_from(cursor.read_u32()?)
            .map_err(|_| invalid("BrtBeginIconSet14", "icon-set type overflow"))?;
        if icon_set_type > 19 {
            return Err(invalid("BrtBeginIconSet14", "invalid icon-set type"));
        }
        let flags = cursor.read_u16()?;
        cursor.finish()?;
        if flags & 0xff80 != 0 {
            return Err(invalid("BrtBeginIconSet14", "reserved flags are nonzero"));
        }
        Ok(IconSet14Header {
            icon_set_type,
            custom: flags & 0x01 != 0,
            show_value: flags & 0x02 == 0,
            reverse: flags & 0x04 == 0,
            unused_flags: flags & 0x78,
        })
    }

    pub fn serialize_header(&self) -> XlsbResult<Vec<u8>> {
        if self.icon_set_type > 19 || self.unused_flags & !0x78 != 0 {
            return Err(invalid("BrtBeginIconSet14", "invalid icon-set header"));
        }
        let mut flags = self.unused_flags;
        flags |= u16::from(self.custom_icons.is_some());
        flags |= u16::from(!self.show_value) << 1;
        flags |= u16::from(!self.reverse) << 2;
        let mut data = Vec::with_capacity(10);
        data.extend_from_slice(&0u32.to_le_bytes());
        data.extend_from_slice(&u32::from(self.icon_set_type).to_le_bytes());
        data.extend_from_slice(&flags.to_le_bytes());
        Ok(data)
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct IconSet14Header {
    pub icon_set_type: u8,
    pub custom: bool,
    pub show_value: bool,
    pub reverse: bool,
    pub unused_flags: u16,
}

pub(crate) fn icon_count14(icon_set_type: u8) -> usize {
    match icon_set_type {
        0..=7 | 17 | 18 => 3,
        8..=12 => 4,
        13..=16 | 19 => 5,
        _ => 0,
    }
}

pub fn parse_rule_extension_guid(data: &[u8]) -> XlsbResult<[u8; 16]> {
    if data.len() != 20 {
        return Err(XlsbError::InvalidLength {
            expected: 20,
            found: data.len(),
        });
    }
    if data[..4] != [0; 4] {
        return Err(invalid("BrtCFRuleExt", "nonzero FRTBlank"));
    }
    Ok(data[4..].try_into().expect("sixteen-byte GUID"))
}

pub fn serialize_rule_extension_guid(guid: [u8; 16]) -> [u8; 20] {
    let mut data = [0; 20];
    data[4..].copy_from_slice(&guid);
    data
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
    /// Office 2013 color scale.
    pub color_scale14: Option<ColorScale>,
    /// Office 2013 data bar.
    pub data_bar14: Option<DataBar14>,
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
    pub extension14: Option<ConditionalFormattingRule14Metadata>,
    /// GUID linking a classic rule to an Office 2013 data-bar augmentation.
    pub classic_extension_guid: Option<[u8; 16]>,
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
            color_scale14: None,
            data_bar14: None,
            icon_set14: None,
            operator,
            parameter,
            template,
            text,
            above_average,
            bottom,
            percent,
            extension14: None,
            classic_extension_guid: None,
        })
    }

    /// Parse an Office 2013 `BrtBeginCFRule14` payload.
    pub fn parse_extension14(data: &[u8]) -> XlsbResult<Self> {
        let context = FormulaResolutionContext::default();
        Self::parse_extension14_with_context(data, (0, 0), &context)
    }

    pub(crate) fn parse_extension14_with_context(
        data: &[u8],
        base: (u32, u32),
        context: &FormulaResolutionContext,
    ) -> XlsbResult<Self> {
        let (formulas, header_size) = parse_formula_header(data, "BrtBeginCFRule14", 2)?;
        let mut cursor = CfCursor::new(&data[header_size..], "BrtBeginCFRule14");
        let rule_type_raw = cursor.read_u32()?;
        let rule_type = CfRuleType::from_u32(rule_type_raw).ok_or_else(|| {
            invalid(
                "BrtBeginCFRule14",
                format!("invalid rule type {rule_type_raw}"),
            )
        })?;
        let template = cursor.read_u32()?;
        validate_extension14_template(rule_type, template)?;
        let raw_dxf = cursor.read_u32()?;
        let signed_priority = cursor.read_i32()?;
        if signed_priority != -1 && signed_priority <= 0 {
            return Err(invalid(
                "BrtBeginCFRule14",
                format!("invalid priority {signed_priority}"),
            ));
        }
        if signed_priority == -1 && (rule_type != CfRuleType::DataBar || raw_dxf != 0) {
            return Err(invalid(
                "BrtBeginCFRule14",
                "priority -1 requires a data-bar rule and zero DXF index",
            ));
        }
        let visual = matches!(
            rule_type,
            CfRuleType::ColorScale | CfRuleType::DataBar | CfRuleType::IconSet
        );
        if signed_priority > 0 && visual && raw_dxf != u32::MAX {
            return Err(invalid(
                "BrtBeginCFRule14",
                "visual rule has a differential-format index",
            ));
        }
        let dxf_id = if signed_priority == -1 || raw_dxf == u32::MAX {
            None
        } else {
            Some(raw_dxf)
        };
        let parameter = cursor.read_u32()?;
        let reserved1 = cursor.read_u32()?;
        let reserved2 = cursor.read_u32()?;
        let flags = cursor.read_u16()?;
        if reserved1 != 0 || reserved2 != 0 || flags & !0x1e != 0 {
            return Err(invalid("BrtBeginCFRule14", "reserved fields are nonzero"));
        }
        let stop_if_true = flags & 0x02 != 0;
        let above_average = flags & 0x04 != 0;
        let bottom = flags & 0x08 != 0;
        let percent = flags & 0x10 != 0;
        if visual && stop_if_true {
            return Err(invalid("BrtBeginCFRule14", "visual rule sets stop-if-true"));
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
        let unused = cursor.read_u32()?;
        let guid = cursor.read_array::<16>()?;
        let guid_present = cursor.read_bool32()?;
        let text = cursor.read_nullable_string()?;
        cursor.finish()?;

        if template == 8 {
            if text
                .as_ref()
                .is_none_or(|text| text.is_empty() || text.encode_utf16().count() > 255)
            {
                return Err(invalid(
                    "BrtBeginCFRule14",
                    "contains-text template has an invalid text parameter",
                ));
            }
        } else if text.is_some() {
            return Err(invalid(
                "BrtBeginCFRule14",
                "non-text template has a string parameter",
            ));
        }

        let mut formula_slots: [Option<CellParsedFormula>; 3] = [None, None, None];
        let mut formula_iter = formulas.into_iter();
        for (index, declared_size) in declared.into_iter().enumerate() {
            if declared_size == 0 {
                continue;
            }
            let formula = formula_iter.next().ok_or_else(|| {
                invalid(
                    "BrtBeginCFRule14",
                    "declared formula is absent from FRTHeader",
                )
            })?;
            if formula.rgce.len() != declared_size as usize {
                return Err(invalid(
                    "BrtBeginCFRule14",
                    format!(
                        "formula {} declared {declared_size} token bytes, found {}",
                        index + 1,
                        formula.rgce.len()
                    ),
                ));
            }
            formula_slots[index] = Some(formula);
        }
        if formula_iter.next().is_some() {
            return Err(invalid(
                "BrtBeginCFRule14",
                "FRTHeader contains an undeclared formula",
            ));
        }
        validate_formula_slots(rule_type, template, parameter, &formula_slots)?;

        let mut binary_formulas = Vec::new();
        let mut formula_extras = Vec::new();
        let mut formula_texts = Vec::new();
        for formula in formula_slots.into_iter().flatten() {
            binary_formulas.push(formula.rgce.clone());
            formula_extras.push(formula.rgcb.clone());
            formula_texts.push(render_formula(&formula, base, context)?);
        }
        let operator = (rule_type == CfRuleType::CellIs)
            .then(|| u8::try_from(parameter).ok())
            .flatten();

        Ok(Self {
            rule_type,
            dxf_id,
            priority: u32::try_from(signed_priority).unwrap_or(0),
            stop_if_true,
            formulas: binary_formulas,
            formula_extras,
            formula_texts,
            color_scale: None,
            data_bar: None,
            icon_set: None,
            color_scale14: None,
            data_bar14: None,
            icon_set14: None,
            operator,
            parameter,
            template,
            text,
            above_average,
            bottom,
            percent,
            extension14: Some(ConditionalFormattingRule14Metadata {
                priority: signed_priority,
                unused,
                guid,
                guid_present,
                linked_classic_priority: None,
            }),
            classic_extension_guid: None,
        })
    }

    /// Serialize an Office 2013 `BrtBeginCFRule14` payload.
    pub fn serialize_extension14(&self) -> XlsbResult<Vec<u8>> {
        let metadata = self.extension14.ok_or_else(|| {
            invalid(
                "BrtBeginCFRule14",
                "rule does not contain Office 2013 metadata",
            )
        })?;
        validate_extension14_template(self.rule_type, self.template)?;
        if metadata.priority != -1 && metadata.priority <= 0 {
            return Err(invalid(
                "BrtBeginCFRule14",
                format!("invalid priority {}", metadata.priority),
            ));
        }
        if metadata.priority > 0 && self.priority != metadata.priority as u32 {
            return Err(invalid(
                "BrtBeginCFRule14",
                "classic and extension priorities disagree",
            ));
        }
        if metadata.priority == -1 && self.rule_type != CfRuleType::DataBar {
            return Err(invalid(
                "BrtBeginCFRule14",
                "priority -1 is only valid for a data-bar extension",
            ));
        }
        let parameter = effective_rule_parameter(self)?;
        validate_parameter_and_flags(
            self.rule_type,
            self.template,
            parameter,
            self.above_average,
            self.bottom,
            self.percent,
        )?;
        let visual = matches!(
            self.rule_type,
            CfRuleType::ColorScale | CfRuleType::DataBar | CfRuleType::IconSet
        );
        if visual && (self.stop_if_true || (metadata.priority > 0 && self.dxf_id.is_some())) {
            return Err(invalid(
                "BrtBeginCFRule14",
                "visual rule has a DXF or stop-if-true flag",
            ));
        }
        if metadata.priority == -1 && self.dxf_id.is_some() {
            return Err(invalid(
                "BrtBeginCFRule14",
                "data-bar extension has a DXF index",
            ));
        }
        validate_rule_text(self.template, self.text.as_deref(), "BrtBeginCFRule14")?;

        let formulas = effective_rule_formulas(self)?;
        validate_formula_count(self.rule_type, self.template, parameter, formulas.len())?;
        let mut slots: [Option<&CellParsedFormula>; 3] = [None, None, None];
        let start = if visual { 2 } else { 0 };
        for (index, formula) in formulas.iter().enumerate() {
            slots[start + index] = Some(formula);
        }
        let owned_slots = slots.each_ref().map(|formula| formula.cloned());
        validate_formula_slots(self.rule_type, self.template, parameter, &owned_slots)?;

        let mut payload = serialize_formula_header(&formulas, 2)?;
        payload.extend_from_slice(&(self.rule_type as u32).to_le_bytes());
        payload.extend_from_slice(&self.template.to_le_bytes());
        let raw_dxf = if metadata.priority == -1 {
            0
        } else {
            self.dxf_id.unwrap_or(u32::MAX)
        };
        payload.extend_from_slice(&raw_dxf.to_le_bytes());
        payload.extend_from_slice(&metadata.priority.to_le_bytes());
        payload.extend_from_slice(&parameter.to_le_bytes());
        payload.extend_from_slice(&0u32.to_le_bytes());
        payload.extend_from_slice(&0u32.to_le_bytes());
        let mut flags = 0u16;
        flags |= u16::from(self.stop_if_true) << 1;
        flags |= u16::from(self.above_average) << 2;
        flags |= u16::from(self.bottom) << 3;
        flags |= u16::from(self.percent) << 4;
        payload.extend_from_slice(&flags.to_le_bytes());
        for formula in &slots {
            payload.extend_from_slice(
                &u32::try_from(formula.map_or(0, |formula| formula.rgce.len()))
                    .map_err(|_| XlsbError::InvalidFormula("formula is too large".to_string()))?
                    .to_le_bytes(),
            );
        }
        payload.extend_from_slice(&metadata.unused.to_le_bytes());
        payload.extend_from_slice(&metadata.guid);
        payload.extend_from_slice(&u32::from(metadata.guid_present).to_le_bytes());
        write_nullable_string(&mut payload, self.text.as_deref())?;
        Ok(payload)
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
    /// Binary record family used to encode this collection.
    pub record_kind: ConditionalFormattingRecordKind,
}

impl ConditionalFormatting {
    pub fn new(ranges: Vec<String>) -> Self {
        ConditionalFormatting {
            ranges,
            rules: Vec::new(),
            pivot_only: false,
            record_kind: ConditionalFormattingRecordKind::Classic,
        }
    }

    /// Create an Office 2013 conditional-formatting collection.
    pub fn new_extension14(ranges: Vec<String>) -> Self {
        Self {
            ranges,
            rules: Vec::new(),
            pivot_only: false,
            record_kind: ConditionalFormattingRecordKind::Extension14,
        }
    }

    pub fn add_rule(&mut self, rule: ConditionalFormattingRule) {
        self.rules.push(rule);
    }

    /// Parse an Office 2013 `BrtBeginConditionalFormatting14` payload.
    pub fn parse_extension14_header(data: &[u8]) -> XlsbResult<(Self, u32)> {
        let (formatting, count, _) = Self::parse_extension14_header_with_base(data)?;
        Ok((formatting, count))
    }

    pub(crate) fn parse_extension14_header_with_base(
        data: &[u8],
    ) -> XlsbResult<(Self, u32, (u32, u32))> {
        let (ranges, header_size) =
            parse_sqref_header(data, "BrtBeginConditionalFormatting14", i32::MAX as usize)?;
        let mut cursor = CfCursor::new(&data[header_size..], "BrtBeginConditionalFormatting14");
        let count = cursor.read_u32()?;
        let pivot_only = cursor.read_bool32()?;
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
            Self {
                ranges,
                rules: Vec::new(),
                pivot_only,
                record_kind: ConditionalFormattingRecordKind::Extension14,
            },
            count,
            base,
        ))
    }

    /// Serialize an Office 2013 `BrtBeginConditionalFormatting14` payload.
    pub fn serialize_extension14_header(&self) -> XlsbResult<Vec<u8>> {
        let mut ranges = Vec::new();
        for range_list in &self.ranges {
            for range in range_list
                .split([',', ' '])
                .filter(|range| !range.is_empty())
            {
                let (first, last) = range.split_once(':').unwrap_or((range, range));
                let (first_row, first_col) = parse_cell_reference(first)?;
                let (last_row, last_col) = parse_cell_reference(last)?;
                ranges.push((first_row, last_row, first_col, last_col));
            }
        }
        let mut data = serialize_sqref_header(&ranges)?;
        data.extend_from_slice(
            &u32::try_from(self.rules.len())
                .map_err(|_| invalid("BrtBeginConditionalFormatting14", "rule count overflow"))?
                .to_le_bytes(),
        );
        data.extend_from_slice(&u32::from(self.pivot_only).to_le_bytes());
        Ok(data)
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
            record_kind: ConditionalFormattingRecordKind::Classic,
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

fn validate_extension14_template(rule_type: CfRuleType, template: u32) -> XlsbResult<()> {
    if rule_type == CfRuleType::DataBar && template == 0 {
        Ok(())
    } else {
        validate_template(rule_type, template).map_err(|_| {
            invalid(
                "BrtBeginCFRule14",
                format!("template {template} is invalid for {rule_type:?}"),
            )
        })
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

fn effective_rule_parameter(rule: &ConditionalFormattingRule) -> XlsbResult<u32> {
    if rule.rule_type != CfRuleType::CellIs {
        if rule.operator.is_some() {
            return Err(invalid(
                "BrtBeginCFRule14",
                "operator is set on a non-cell-comparison rule",
            ));
        }
        return Ok(rule.parameter);
    }
    let parameter = rule.operator.map_or(rule.parameter, u32::from);
    if rule.parameter != 0 && rule.parameter != parameter {
        return Err(invalid(
            "BrtBeginCFRule14",
            "operator and exact parameter disagree",
        ));
    }
    Ok(parameter)
}

fn effective_rule_formulas(rule: &ConditionalFormattingRule) -> XlsbResult<Vec<CellParsedFormula>> {
    if !rule.formulas.is_empty() {
        if !rule.formula_extras.is_empty() && rule.formula_extras.len() != rule.formulas.len() {
            return Err(XlsbError::InvalidFormula(
                "conditional-format ancillary stream count does not match formulas".to_string(),
            ));
        }
        return rule
            .formulas
            .iter()
            .enumerate()
            .map(|(index, rgce)| {
                if rgce.is_empty() || rgce.len() > MAX_CELL_FORMULA_BYTES {
                    return Err(XlsbError::InvalidFormula(format!(
                        "conditional-format formula length {} is outside 1..={MAX_CELL_FORMULA_BYTES}",
                        rgce.len()
                    )));
                }
                Ok(CellParsedFormula {
                    rgce: rgce.clone(),
                    rgcb: rule.formula_extras.get(index).cloned().unwrap_or_default(),
                })
            })
            .collect();
    }
    rule.formula_texts
        .iter()
        .map(|formula| crate::xlsb::formula::FormulaCompiler::compile(formula))
        .collect()
}

fn validate_rule_text(template: u32, text: Option<&str>, record: &'static str) -> XlsbResult<()> {
    if template == 8 {
        if text.is_none_or(|text| text.is_empty() || text.encode_utf16().count() > 255) {
            return Err(invalid(record, "invalid text parameter"));
        }
    } else if text.is_some() {
        return Err(invalid(record, "non-text template has a text parameter"));
    }
    Ok(())
}

fn write_nullable_string(data: &mut Vec<u8>, value: Option<&str>) -> XlsbResult<()> {
    let Some(value) = value else {
        data.extend_from_slice(&u32::MAX.to_le_bytes());
        return Ok(());
    };
    let units = value.encode_utf16().collect::<Vec<_>>();
    data.extend_from_slice(
        &u32::try_from(units.len())
            .map_err(|_| invalid("XLNullableWideString", "string length overflow"))?
            .to_le_bytes(),
    );
    for unit in units {
        data.extend_from_slice(&unit.to_le_bytes());
    }
    Ok(())
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

    fn read_u8(&mut self) -> XlsbResult<u8> {
        Ok(self.take(1)?[0])
    }

    fn read_bool8(&mut self) -> XlsbResult<bool> {
        match self.read_u8()? {
            0 => Ok(false),
            1 => Ok(true),
            value => Err(invalid(self.record, format!("invalid Boolean {value}"))),
        }
    }

    fn read_u32(&mut self) -> XlsbResult<u32> {
        let bytes = self.take(4)?;
        Ok(u32::from_le_bytes(
            bytes.try_into().expect("four-byte field"),
        ))
    }

    fn read_i32(&mut self) -> XlsbResult<i32> {
        let bytes = self.take(4)?;
        Ok(i32::from_le_bytes(
            bytes.try_into().expect("four-byte field"),
        ))
    }

    fn read_array<const N: usize>(&mut self) -> XlsbResult<[u8; N]> {
        self.take(N)?
            .try_into()
            .map_err(|_| XlsbError::InvalidLength {
                expected: N,
                found: self.data.len().saturating_sub(self.offset),
            })
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
    fn extension_cfvo_roundtrips_formula_and_automatic_bounds() {
        let formula = FormulaCompiler::compile("$A$1").unwrap();
        let formula_value = Cfvo {
            cfvo_type: 7,
            value: Some("$A$1".to_string()),
            numeric_value: 0.0,
            save_greater_than_or_equal: true,
            greater_than_or_equal: false,
            formula_binary: Some(formula.clone()),
        };
        let encoded = formula_value.serialize_extension14().unwrap();
        let parsed = Cfvo::parse_extension14(&encoded).unwrap();
        assert_eq!(parsed.cfvo_type, 7);
        assert_eq!(parsed.formula_binary, Some(formula));
        assert!(!parsed.greater_than_or_equal);

        for cfvo_type in [8, 9] {
            let automatic = Cfvo {
                cfvo_type,
                value: None,
                numeric_value: 0.0,
                save_greater_than_or_equal: false,
                greater_than_or_equal: true,
                formula_binary: None,
            };
            let encoded = automatic.serialize_extension14().unwrap();
            assert_eq!(Cfvo::parse_extension14(&encoded).unwrap(), automatic);
        }
    }

    #[test]
    fn extension_cfvo_rejects_inconsistent_formula_metadata() {
        let formula = FormulaCompiler::compile("1").unwrap();
        let value = Cfvo {
            cfvo_type: 7,
            value: Some("1".to_string()),
            numeric_value: 0.0,
            save_greater_than_or_equal: false,
            greater_than_or_equal: true,
            formula_binary: Some(formula),
        };
        let mut encoded = value.serialize_extension14().unwrap();
        let declared_offset = encoded.len() - 4;
        encoded[declared_offset..].copy_from_slice(&999u32.to_le_bytes());
        assert!(Cfvo::parse_extension14(&encoded).is_err());
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
    fn extension_color_and_rule_guid_roundtrip_exactly() {
        let color = ConditionalFormatColor::theme(4, -2_500).unwrap();
        let encoded = color.serialize_extension14().unwrap();
        assert_eq!(
            ConditionalFormatColor::parse_extension14(&encoded).unwrap(),
            color
        );
        let mut malformed = encoded;
        malformed[0] = 1;
        assert!(ConditionalFormatColor::parse_extension14(&malformed).is_err());

        let guid = [0x42; 16];
        let encoded = serialize_rule_extension_guid(guid);
        assert_eq!(parse_rule_extension_guid(&encoded).unwrap(), guid);
        let mut malformed = encoded;
        malformed[3] = 1;
        assert!(parse_rule_extension_guid(&malformed).is_err());
    }

    #[test]
    fn extension_data_bar_header_preserves_flags() {
        let mut bar = DataBar14::new(
            Cfvo::new(8, None),
            Cfvo::new(9, None),
            ConditionalFormatColor::from_argb(0xff44_72c4),
        );
        bar.min_length = 3;
        bar.max_length = 97;
        bar.show_value = false;
        bar.direction = DataBarDirection14::RightToLeft;
        bar.axis_position = DataBarAxisPosition14::Midpoint;
        bar.border = true;
        bar.custom_negative_fill = true;
        bar.unused_flags = 0xA5F0;
        let encoded = bar.serialize_header().unwrap();
        let parsed = DataBar14::parse_header(&encoded).unwrap();
        assert_eq!(parsed.min_length, 3);
        assert_eq!(parsed.max_length, 97);
        assert!(!parsed.show_value);
        assert_eq!(parsed.direction, DataBarDirection14::RightToLeft);
        assert_eq!(parsed.axis_position, DataBarAxisPosition14::Midpoint);
        assert!(parsed.border);
        assert!(parsed.gradient);
        assert!(parsed.custom_negative_fill);
        assert_eq!(parsed.unused_flags, 0xA5F0);

        let mut malformed = encoded;
        malformed[6] = 2;
        assert!(DataBar14::parse_header(&malformed).is_err());
    }

    #[test]
    fn extension_icon_set_and_custom_icons_roundtrip() {
        let mut set = IconSet14::new(19, vec![Cfvo::new(1, Some("0".to_string())); 5]);
        set.show_value = false;
        set.reverse = true;
        set.unused_flags = 0x38;
        set.custom_icons = Some(vec![
            ConditionalFormatIcon {
                icon_set: -1,
                index: -1,
            };
            5
        ]);
        let encoded = set.serialize_header().unwrap();
        let parsed = IconSet14::parse_header(&encoded).unwrap();
        assert_eq!(parsed.icon_set_type, 19);
        assert!(parsed.custom);
        assert!(!parsed.show_value);
        assert!(parsed.reverse);
        assert_eq!(parsed.unused_flags, 0x38);

        for icon in set.custom_icons.unwrap() {
            let encoded = icon.serialize().unwrap();
            assert_eq!(ConditionalFormatIcon::parse(&encoded).unwrap(), icon);
        }
        assert!(
            ConditionalFormatIcon {
                icon_set: 0,
                index: 3,
            }
            .serialize()
            .is_err()
        );
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

    #[test]
    fn extension_header_roundtrips_ranges_and_pivot_flag() {
        let mut formatting = ConditionalFormatting::new(vec!["A1:B2 C3".to_string()]);
        formatting.pivot_only = true;
        formatting
            .rules
            .push(ConditionalFormattingRule::new(CfRuleType::Expression, 1));
        let encoded = formatting.serialize_extension14_header().unwrap();
        let (parsed, count) = ConditionalFormatting::parse_extension14_header(&encoded).unwrap();
        assert_eq!(count, 1);
        assert_eq!(parsed.ranges, ["A1:B2", "C3"]);
        assert!(parsed.pivot_only);
        assert_eq!(
            parsed.record_kind,
            ConditionalFormattingRecordKind::Extension14
        );
    }

    #[test]
    fn extension_rule_roundtrips_two_formulas_and_ancillary_data() {
        let first = FormulaCompiler::compile("{1,2}").unwrap();
        let second = FormulaCompiler::compile("10").unwrap();
        assert!(!first.rgcb.is_empty());
        let mut rule = ConditionalFormattingRule::new(CfRuleType::CellIs, 7);
        rule.operator = Some(1);
        rule.parameter = 1;
        rule.formulas = vec![first.rgce.clone(), second.rgce.clone()];
        rule.formula_extras = vec![first.rgcb.clone(), second.rgcb.clone()];
        rule.dxf_id = Some(4);
        rule.extension14 = Some(ConditionalFormattingRule14Metadata {
            priority: 7,
            unused: 0xA5A5_5A5A,
            guid: [0x3c; 16],
            guid_present: true,
            linked_classic_priority: None,
        });

        let encoded = rule.serialize_extension14().unwrap();
        let parsed = ConditionalFormattingRule::parse_extension14(&encoded).unwrap();
        assert_eq!(parsed.priority, 7);
        assert_eq!(parsed.operator, Some(1));
        assert_eq!(parsed.formulas, [first.rgce, second.rgce]);
        assert_eq!(parsed.formula_extras, [first.rgcb, second.rgcb]);
        assert_eq!(parsed.extension14, rule.extension14);
        assert_eq!(parsed.serialize_extension14().unwrap(), encoded);
    }

    #[test]
    fn extension_rule_preserves_signed_data_bar_linkage() {
        let mut rule = ConditionalFormattingRule::new(CfRuleType::DataBar, 0);
        rule.template = 0;
        rule.extension14 = Some(ConditionalFormattingRule14Metadata {
            priority: -1,
            unused: 0xDEAD_BEEF,
            guid: [0x96; 16],
            guid_present: true,
            linked_classic_priority: None,
        });

        let encoded = rule.serialize_extension14().unwrap();
        let parsed = ConditionalFormattingRule::parse_extension14(&encoded).unwrap();
        assert_eq!(parsed.priority, 0);
        assert_eq!(parsed.template, 0);
        assert_eq!(parsed.extension14, rule.extension14);
        assert_eq!(parsed.serialize_extension14().unwrap(), encoded);
    }

    #[test]
    fn extension_rule_rejects_malformed_fixed_and_formula_fields() {
        let mut rule = ConditionalFormattingRule::new(CfRuleType::Expression, 2);
        rule.formula_texts.push("1".to_string());
        rule.extension14 = Some(ConditionalFormattingRule14Metadata {
            priority: 2,
            unused: 0,
            guid: [0; 16],
            guid_present: false,
            linked_classic_priority: None,
        });
        let encoded = rule.serialize_extension14().unwrap();
        let (_, fixed_offset) = parse_formula_header(&encoded, "test", 2).unwrap();

        let mut reserved = encoded.clone();
        reserved[fixed_offset + 20..fixed_offset + 24].copy_from_slice(&1u32.to_le_bytes());
        assert!(ConditionalFormattingRule::parse_extension14(&reserved).is_err());

        let mut priority = encoded.clone();
        priority[fixed_offset + 12..fixed_offset + 16].copy_from_slice(&0i32.to_le_bytes());
        assert!(ConditionalFormattingRule::parse_extension14(&priority).is_err());

        let mut declared = encoded;
        declared[fixed_offset + 30..fixed_offset + 34].copy_from_slice(&999u32.to_le_bytes());
        assert!(ConditionalFormattingRule::parse_extension14(&declared).is_err());
    }
}
