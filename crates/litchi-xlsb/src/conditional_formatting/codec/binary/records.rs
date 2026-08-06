//! Typed Brt* conditional-formatting record codecs.

use crate::conditional_formatting::model::*;
use crate::formula::{ParsedFormula, Resolution};

use super::super::semantic::{
    EmptyFormulaResolution, effective_rule_formulas, effective_rule_parameter, format_number,
    icon_count14, render_formula, validate_extension14_template, validate_formula_count,
    validate_formula_slots, validate_parameter_and_flags, validate_rule_text, validate_template,
};
use super::super::{Error, Result, invalid};
use super::wire::{
    CfCursor, cell_reference, parse_cell_reference, parse_formula_header, parse_sqref_header,
    serialize_formula_header, serialize_sqref_header, write_nullable_string,
};

impl Value {
    pub fn parse(data: &[u8]) -> Result<Self> {
        if data.len() < 24 {
            return Err(Error::InvalidLength {
                expected: 24,
                found: data.len(),
            });
        }
        let context = EmptyFormulaResolution;
        Self::parse_with_context(data, (0, 0), &context)
    }

    pub fn parse_with_context(
        data: &[u8],
        base: (u32, u32),
        context: &impl Resolution,
    ) -> Result<Self> {
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
    pub fn parse_extension14(data: &[u8]) -> Result<Self> {
        let context = EmptyFormulaResolution;
        Self::parse_extension14_with_context(data, (0, 0), &context)
    }

    pub fn parse_extension14_with_context(
        data: &[u8],
        base: (u32, u32),
        context: &impl Resolution,
    ) -> Result<Self> {
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
    pub fn serialize_extension14(&self) -> Result<Vec<u8>> {
        self.serialize_extension14_with(
            self.formula_binary.as_ref(),
            self.numeric_value,
            self.save_greater_than_or_equal,
        )
    }

    pub(super) fn serialize_extension14_with(
        &self,
        formula_binary: Option<&ParsedFormula>,
        numeric_value: f64,
        save_greater_than_or_equal: bool,
    ) -> Result<Vec<u8>> {
        if !matches!(self.cfvo_type, 1 | 2 | 3 | 4 | 5 | 7 | 8 | 9) {
            return Err(invalid(
                "BrtCFVO14",
                format!("invalid type {}", self.cfvo_type),
            ));
        }
        if !numeric_value.is_finite() {
            return Err(invalid("BrtCFVO14", "non-finite numeric parameter"));
        }
        if formula_binary.is_none()
            && matches!(self.cfvo_type, 4 | 5)
            && !(0.0..=100.0).contains(&numeric_value)
        {
            return Err(invalid(
                "BrtCFVO14",
                format!("percentage parameter {} outside 0..=100", numeric_value),
            ));
        }
        if matches!(self.cfvo_type, 2 | 3 | 8 | 9) && formula_binary.is_some() {
            return Err(invalid(
                "BrtCFVO14",
                "automatic/min/max threshold contains a formula",
            ));
        }
        if self.cfvo_type == 7 && formula_binary.is_none() {
            return Err(invalid("BrtCFVO14", "formula threshold omits its formula"));
        }
        let formulas = formula_binary.map_or(&[][..], std::slice::from_ref);
        let mut data = serialize_formula_header(formulas, 1)?;
        data.extend_from_slice(&u32::from(self.cfvo_type).to_le_bytes());
        data.extend_from_slice(&numeric_value.to_le_bytes());
        data.extend_from_slice(&u32::from(save_greater_than_or_equal).to_le_bytes());
        data.extend_from_slice(&u32::from(self.greater_than_or_equal).to_le_bytes());
        data.extend_from_slice(
            &u32::try_from(formula_binary.map_or(0, |formula| formula.rgce.len()))
                .map_err(|_| Error::InvalidFormula("formula is too large".to_string()))?
                .to_le_bytes(),
        );
        Ok(data)
    }
}

impl Color {
    pub fn theme(index: u8, tint: i16) -> Result<Self> {
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

    pub fn parse(data: &[u8]) -> Result<Self> {
        if data.len() != 8 {
            return Err(Error::InvalidLength {
                expected: 8,
                found: data.len(),
            });
        }
        let raw: [u8; 8] = data.try_into().map_err(|_| Error::InvalidLength {
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

    pub fn to_bytes(self) -> Result<[u8; 8]> {
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
    pub fn parse_extension14(data: &[u8]) -> Result<Self> {
        if data.len() != 12 {
            return Err(Error::InvalidLength {
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
    pub fn serialize_extension14(self) -> Result<[u8; 12]> {
        let mut data = [0; 12];
        data[4..].copy_from_slice(&self.to_bytes()?);
        Ok(data)
    }
}

impl Direction14 {
    fn from_u8(value: u8) -> Option<Self> {
        match value {
            0 => Some(Self::Context),
            1 => Some(Self::LeftToRight),
            2 => Some(Self::RightToLeft),
            _ => None,
        }
    }
}

impl AxisPosition14 {
    fn from_u8(value: u8) -> Option<Self> {
        match value {
            0 => Some(Self::Automatic),
            1 => Some(Self::Midpoint),
            2 => Some(Self::None),
            _ => None,
        }
    }
}

impl Bar14 {
    pub fn parse_header(data: &[u8]) -> Result<BarHeader14> {
        let mut cursor = CfCursor::new(data, "BrtBeginDatabar14");
        if cursor.read_u32()? != 0 {
            return Err(invalid("BrtBeginDatabar14", "nonzero FRTBlank"));
        }
        let min_length = cursor.read_u8()?;
        let max_length = cursor.read_u8()?;
        let show_value = cursor.read_bool8()?;
        let direction = Direction14::from_u8(cursor.read_u8()?)
            .ok_or_else(|| invalid("BrtBeginDatabar14", "invalid direction"))?;
        let axis_position = AxisPosition14::from_u8(cursor.read_u8()?)
            .ok_or_else(|| invalid("BrtBeginDatabar14", "invalid axis position"))?;
        let flags = cursor.read_u16()?;
        cursor.finish()?;
        if min_length > max_length || max_length > 100 {
            return Err(invalid(
                "BrtBeginDatabar14",
                "invalid minimum/maximum length",
            ));
        }
        Ok(BarHeader14 {
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

    pub fn serialize_header(&self) -> Result<Vec<u8>> {
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

impl Icon {
    pub fn parse(data: &[u8]) -> Result<Self> {
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

    pub fn serialize(self) -> Result<[u8; 12]> {
        self.validate()?;
        let mut data = [0; 12];
        data[4..8].copy_from_slice(&self.icon_set.to_le_bytes());
        data[8..].copy_from_slice(&self.index.to_le_bytes());
        Ok(data)
    }

    fn validate(self) -> Result<()> {
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

impl IconSet14 {
    pub fn parse_header(data: &[u8]) -> Result<IconHeader14> {
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
        Ok(IconHeader14 {
            icon_set_type,
            custom: flags & 0x01 != 0,
            show_value: flags & 0x02 == 0,
            reverse: flags & 0x04 == 0,
            unused_flags: flags & 0x78,
        })
    }

    pub fn serialize_header(&self) -> Result<Vec<u8>> {
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

impl Rule {
    pub fn parse(data: &[u8]) -> Result<Self> {
        let context = EmptyFormulaResolution;
        Self::parse_with_context(data, (0, 0), &context)
    }

    pub fn parse_with_context(
        data: &[u8],
        base: (u32, u32),
        context: &impl Resolution,
    ) -> Result<Self> {
        let mut cursor = CfCursor::new(data, "BrtBeginCFRule");
        let rule_type_raw = cursor.read_u32()?;
        let rule_type = RuleType::from_u32(rule_type_raw).ok_or_else(|| {
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
            RuleType::ColorScale | RuleType::DataBar | RuleType::IconSet
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
            RuleType::ColorScale | RuleType::DataBar | RuleType::IconSet
        ) && stop_if_true
        {
            return Err(invalid("BrtBeginCFRule", "visual rule sets stop-if-true"));
        }
        if rule_type != RuleType::TopN && (bottom || percent) {
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
        let mut formula_slots: [Option<ParsedFormula>; 3] = [None, None, None];
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
        let operator = (rule_type == RuleType::CellIs)
            .then(|| u8::try_from(parameter).ok())
            .flatten();
        if rule_type == RuleType::CellIs && !matches!(operator, Some(1..=8)) {
            return Err(invalid(
                "BrtBeginCFRule",
                format!("invalid cell comparison operator {parameter}"),
            ));
        }

        Ok(Rule {
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
    pub fn parse_extension14(data: &[u8]) -> Result<Self> {
        let context = EmptyFormulaResolution;
        Self::parse_extension14_with_context(data, (0, 0), &context)
    }

    pub fn parse_extension14_with_context(
        data: &[u8],
        base: (u32, u32),
        context: &impl Resolution,
    ) -> Result<Self> {
        let (formulas, header_size) = parse_formula_header(data, "BrtBeginCFRule14", 2)?;
        let mut cursor = CfCursor::new(&data[header_size..], "BrtBeginCFRule14");
        let rule_type_raw = cursor.read_u32()?;
        let rule_type = RuleType::from_u32(rule_type_raw).ok_or_else(|| {
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
        if signed_priority == -1 && (rule_type != RuleType::DataBar || raw_dxf != 0) {
            return Err(invalid(
                "BrtBeginCFRule14",
                "priority -1 requires a data-bar rule and zero DXF index",
            ));
        }
        let visual = matches!(
            rule_type,
            RuleType::ColorScale | RuleType::DataBar | RuleType::IconSet
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

        let mut formula_slots: [Option<ParsedFormula>; 3] = [None, None, None];
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
        let operator = (rule_type == RuleType::CellIs)
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
            extension14: Some(RuleMetadata {
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
    pub fn serialize_extension14(&self) -> Result<Vec<u8>> {
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
        if metadata.priority == -1 && self.rule_type != RuleType::DataBar {
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
            RuleType::ColorScale | RuleType::DataBar | RuleType::IconSet
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
        let mut slots: [Option<&ParsedFormula>; 3] = [None, None, None];
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
                    .map_err(|_| Error::InvalidFormula("formula is too large".to_string()))?
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

impl Formatting {
    /// Parse an Office 2013 `BrtBeginConditionalFormatting14` payload.
    pub fn parse_extension14_header(data: &[u8]) -> Result<(Self, u32)> {
        let (formatting, count, _) = Self::parse_extension14_header_with_base(data)?;
        Ok((formatting, count))
    }

    pub fn parse_extension14_header_with_base(data: &[u8]) -> Result<(Self, u32, (u32, u32))> {
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
                record_kind: RecordKind::Extension14,
            },
            count,
            base,
        ))
    }

    /// Serialize an Office 2013 `BrtBeginConditionalFormatting14` payload.
    pub fn serialize_extension14_header(&self) -> Result<Vec<u8>> {
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

pub fn parse_classic_header(data: &[u8]) -> Result<(Formatting, u32, (u32, u32))> {
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
        Formatting {
            ranges,
            rules: Vec::new(),
            pivot_only,
            record_kind: RecordKind::Classic,
        },
        count,
        base,
    ))
}
