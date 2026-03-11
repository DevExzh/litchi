//! Conditional formatting support for XLSB

use crate::common::binary;
use crate::ooxml::xlsb::error::{XlsbError, XlsbResult};

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
    pub fn from_u8(value: u8) -> Option<Self> {
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
}

/// Conditional formatting value object (CFVO)
#[derive(Debug, Clone)]
pub struct Cfvo {
    /// Type: 1=num, 2=percent, 3=max, 4=min, 5=formula, 6=percentile
    pub cfvo_type: u8,
    /// Value (for num, percent, formula, percentile)
    pub value: Option<String>,
}

impl Cfvo {
    pub fn new(cfvo_type: u8, value: Option<String>) -> Self {
        Cfvo { cfvo_type, value }
    }

    pub fn parse(data: &[u8]) -> XlsbResult<Self> {
        if data.len() < 2 {
            return Err(XlsbError::InvalidLength {
                expected: 2,
                found: data.len(),
            });
        }

        let cfvo_type = data[0];
        let mut offset = 1;

        // Skip flags
        offset += 1;

        let (value, _) = if offset < data.len() {
            read_optional_string(&data[offset..])?
        } else {
            (None, 0)
        };

        Ok(Cfvo { cfvo_type, value })
    }

    pub fn serialize(&self) -> Vec<u8> {
        let mut data = Vec::new();
        data.push(self.cfvo_type);
        data.push(0); // Flags
        write_optional_string(&mut data, self.value.as_deref());
        data
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
        }
    }

    pub fn with_middle(mut self, mid_cfvo: Cfvo, mid_color: u32) -> Self {
        self.mid_cfvo = Some(mid_cfvo);
        self.mid_color = Some(mid_color);
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
}

impl DataBar {
    pub fn new(min_cfvo: Cfvo, max_cfvo: Cfvo, color: u32) -> Self {
        DataBar {
            min_cfvo,
            max_cfvo,
            color,
            show_value: true,
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
    /// Color scale (for ColorScale type)
    pub color_scale: Option<ColorScale>,
    /// Data bar (for DataBar type)
    pub data_bar: Option<DataBar>,
    /// Icon set (for IconSet type)
    pub icon_set: Option<IconSet>,
    /// Operator (for CellIs type): 1=between, 2=not between, 3=equal, etc.
    pub operator: Option<u8>,
}

impl ConditionalFormattingRule {
    pub fn new(rule_type: CfRuleType, priority: u32) -> Self {
        ConditionalFormattingRule {
            rule_type,
            dxf_id: None,
            priority,
            stop_if_true: false,
            formulas: Vec::new(),
            color_scale: None,
            data_bar: None,
            icon_set: None,
            operator: None,
        }
    }

    pub fn parse(data: &[u8]) -> XlsbResult<Self> {
        if data.len() < 12 {
            return Err(XlsbError::InvalidLength {
                expected: 12,
                found: data.len(),
            });
        }

        let mut offset = 0;

        // Type
        let rule_type = CfRuleType::from_u8(data[offset])
            .ok_or_else(|| XlsbError::UnsupportedFeature("Invalid CF rule type".to_string()))?;
        offset += 1;

        // DXF ID (optional, -1 means none)
        let dxf_id_raw = binary::read_i32_le(data, offset)?;
        let dxf_id = if dxf_id_raw >= 0 {
            Some(dxf_id_raw as u32)
        } else {
            None
        };
        offset += 4;

        // Priority
        let priority = binary::read_u32_le_at(data, offset)?;
        offset += 4;

        // Flags
        let flags = data[offset];
        offset += 1;
        let stop_if_true = (flags & 0x01) != 0;

        // Operator (for CellIs type)
        let operator = if offset < data.len() {
            Some(data[offset])
        } else {
            None
        };

        Ok(ConditionalFormattingRule {
            rule_type,
            dxf_id,
            priority,
            stop_if_true,
            formulas: Vec::new(),
            color_scale: None,
            data_bar: None,
            icon_set: None,
            operator,
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
}

impl ConditionalFormatting {
    pub fn new(ranges: Vec<String>) -> Self {
        ConditionalFormatting {
            ranges,
            rules: Vec::new(),
        }
    }

    pub fn add_rule(&mut self, rule: ConditionalFormattingRule) {
        self.rules.push(rule);
    }
}

/// Read optional string from XLSB data
fn read_optional_string(data: &[u8]) -> XlsbResult<(Option<String>, usize)> {
    if data.len() < 4 {
        return Ok((None, 0));
    }

    let len = binary::read_u32_le_at(data, 0)? as usize;
    if len == 0 {
        return Ok((None, 4));
    }

    if data.len() < 4 + len * 2 {
        return Ok((None, 4));
    }

    let mut chars = Vec::with_capacity(len);
    for i in 0..len {
        let ch = binary::read_u16_le_at(data, 4 + i * 2)?;
        chars.push(ch);
    }

    let string = String::from_utf16_lossy(&chars);
    Ok((Some(string), 4 + len * 2))
}

/// Write optional string to XLSB data
fn write_optional_string(data: &mut Vec<u8>, s: Option<&str>) {
    if let Some(s) = s {
        let chars: Vec<u16> = s.encode_utf16().collect();
        data.extend_from_slice(&(chars.len() as u32).to_le_bytes());
        for &ch in &chars {
            data.extend_from_slice(&ch.to_le_bytes());
        }
    } else {
        data.extend_from_slice(&0u32.to_le_bytes());
    }
}

#[cfg(test)]
mod tests {
    use super::*;

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
        let cfvo = Cfvo::new(2, Some("50".to_string()));
        let serialized = cfvo.serialize();
        let parsed = Cfvo::parse(&serialized).unwrap();
        assert_eq!(parsed.cfvo_type, cfvo.cfvo_type);
        assert_eq!(parsed.value, cfvo.value);
    }

    #[test]
    fn test_cfvo_serialize_none_value() {
        let cfvo = Cfvo::new(4, None); // min type
        let serialized = cfvo.serialize();
        assert_eq!(serialized[0], 4); // type
        assert_eq!(serialized[1], 0); // flags
        // Should have 4 bytes of length (0)
        assert_eq!(&serialized[2..6], &[0, 0, 0, 0]);
    }

    #[test]
    fn test_cfvo_parse_too_short() {
        let result = Cfvo::parse(&[0x01]);
        assert!(result.is_err());
    }

    #[test]
    fn test_color_scale_new() {
        let min_cfvo = Cfvo::new(4, None); // min
        let max_cfvo = Cfvo::new(3, None); // max
        let cs = ColorScale::new(min_cfvo, max_cfvo, 0xFFFF0000, 0xFF00FF00);

        assert_eq!(cs.min_cfvo.cfvo_type, 4);
        assert_eq!(cs.max_cfvo.cfvo_type, 3);
        assert_eq!(cs.min_color, 0xFFFF0000);
        assert_eq!(cs.max_color, 0xFF00FF00);
        assert!(cs.mid_cfvo.is_none());
        assert!(cs.mid_color.is_none());
    }

    #[test]
    fn test_color_scale_with_middle() {
        let min_cfvo = Cfvo::new(4, None);
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
        let min_cfvo = Cfvo::new(4, None);
        let max_cfvo = Cfvo::new(3, None);
        let db = DataBar::new(min_cfvo, max_cfvo, 0xFF4472C4);

        assert_eq!(db.min_cfvo.cfvo_type, 4);
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
        // Create minimal valid data for parsing (need at least 12 bytes)
        let data = [
            0x01, // type = CellIs
            0xFF, 0xFF, 0xFF, 0xFF, // dxf_id = -1 (none)
            0x01, 0x00, 0x00, 0x00, // priority = 1
            0x00, // flags (stop_if_true = false)
            0x02, // operator = greater than
            0x00, // padding
            0x00, // padding
        ];

        let rule = ConditionalFormattingRule::parse(&data).unwrap();
        assert_eq!(rule.rule_type, CfRuleType::CellIs);
        assert!(rule.dxf_id.is_none());
        assert_eq!(rule.priority, 1);
        assert!(!rule.stop_if_true);
        assert_eq!(rule.operator, Some(0x02));
    }

    #[test]
    fn test_conditional_formatting_rule_parse_with_dxf() {
        let data = [
            0x01, // type = CellIs
            0x05, 0x00, 0x00, 0x00, // dxf_id = 5
            0x0A, 0x00, 0x00, 0x00, // priority = 10
            0x01, // flags (stop_if_true = true)
            0x00, // operator
            0x00, // padding
            0x00, // padding
        ];

        let rule = ConditionalFormattingRule::parse(&data).unwrap();
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
        let data = [0x00, 0x00, 0x00, 0x00]; // length = 0
        let (result, bytes_read) = read_optional_string(&data).unwrap();
        assert!(result.is_none());
        assert_eq!(bytes_read, 4);
    }

    #[test]
    fn test_read_optional_string_some() {
        // "Hi" encoded as UTF-16LE with length prefix
        let data = [
            0x02, 0x00, 0x00, 0x00, // length = 2
            0x48, 0x00, // 'H'
            0x69, 0x00, // 'i'
        ];
        let (result, bytes_read) = read_optional_string(&data).unwrap();
        assert_eq!(result, Some("Hi".to_string()));
        assert_eq!(bytes_read, 8);
    }

    #[test]
    fn test_read_optional_string_too_short() {
        let data = [0x01]; // too short
        let (result, bytes_read) = read_optional_string(&data).unwrap();
        assert!(result.is_none());
        assert_eq!(bytes_read, 0);
    }

    #[test]
    fn test_write_optional_string_none() {
        let mut data = Vec::new();
        write_optional_string(&mut data, None);
        assert_eq!(data, vec![0x00, 0x00, 0x00, 0x00]);
    }

    #[test]
    fn test_write_optional_string_some() {
        let mut data = Vec::new();
        write_optional_string(&mut data, Some("Test"));

        // Should have 4-byte length (4) followed by UTF-16LE chars
        assert_eq!(data[0..4], [0x04, 0x00, 0x00, 0x00]);
        assert_eq!(data.len(), 4 + 8); // 4 bytes length + 4 chars * 2 bytes
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
        let min_cfvo = Cfvo::new(4, None);
        let max_cfvo = Cfvo::new(3, None);
        let cs = ColorScale::new(min_cfvo, max_cfvo, 0xFFFF0000, 0xFF00FF00);
        let cloned = cs.clone();

        assert_eq!(cloned.min_color, cs.min_color);
        assert_eq!(cloned.max_color, cs.max_color);
    }
}
