//! Bounded Brt* record codec tests.

use super::super::super::model::*;
use super::super::{binary::*, semantic::*};
use crate::formula::ParsedFormula;
use crate::raw::{Writer, kind};

#[cfg(test)]
mod model_tests {
    use super::*;

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
        let formula = TextCompiler::compile("1").unwrap();
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
        assert_eq!(RuleType::from_u8(1), Some(RuleType::CellIs));
        assert_eq!(RuleType::from_u8(2), Some(RuleType::Expression));
        assert_eq!(RuleType::from_u8(3), Some(RuleType::ColorScale));
        assert_eq!(RuleType::from_u8(4), Some(RuleType::DataBar));
        assert_eq!(RuleType::from_u8(5), Some(RuleType::TopN));
        assert_eq!(RuleType::from_u8(6), Some(RuleType::IconSet));
        assert_eq!(RuleType::from_u8(0), None);
        assert_eq!(RuleType::from_u8(7), None);
        assert_eq!(RuleType::from_u8(255), None);
    }

    #[test]
    fn test_cfvo_new() {
        let cfvo = Value::new(1, Some("10".to_string()));
        assert_eq!(cfvo.cfvo_type, 1);
        assert_eq!(cfvo.value, Some("10".to_string()));
    }

    #[test]
    fn test_cfvo_serialize_roundtrip() {
        let parsed = Value::parse(&numeric_cfvo_payload(1, 50.0)).unwrap();
        assert_eq!(parsed.cfvo_type, 1);
        assert_eq!(parsed.value.as_deref(), Some("50"));
        assert_eq!(parsed.numeric_value, 50.0);
    }

    #[test]
    fn test_cfvo_serialize_none_value() {
        let parsed = Value::parse(&numeric_cfvo_payload(2, 0.0)).unwrap();
        assert_eq!(parsed.cfvo_type, 2);
        assert!(parsed.value.is_none());
        assert!(parsed.formula_binary.is_none());
    }

    #[test]
    fn test_cfvo_parse_too_short() {
        let result = Value::parse(&[0x01]);
        assert!(result.is_err());
    }

    #[test]
    fn test_color_scale_new() {
        let min_cfvo = Value::new(2, None); // min
        let max_cfvo = Value::new(3, None); // max
        let cs = Scale::new(min_cfvo, max_cfvo, 0xFFFF0000, 0xFF00FF00);

        assert_eq!(cs.min_cfvo.cfvo_type, 2);
        assert_eq!(cs.max_cfvo.cfvo_type, 3);
        assert_eq!(cs.min_color, 0xFFFF0000);
        assert_eq!(cs.max_color, 0xFF00FF00);
        assert!(cs.mid_cfvo.is_none());
        assert!(cs.mid_color.is_none());
    }

    #[test]
    fn test_color_scale_with_middle() {
        let min_cfvo = Value::new(2, None);
        let mid_cfvo = Value::new(1, Some("50".to_string()));
        let max_cfvo = Value::new(3, None);
        let cs = Scale::new(min_cfvo, max_cfvo, 0xFFFF0000, 0xFF00FF00)
            .with_middle(mid_cfvo, 0xFFFFFF00);

        assert!(cs.mid_cfvo.is_some());
        assert!(cs.mid_color.is_some());
        assert_eq!(cs.mid_color.unwrap(), 0xFFFFFF00);
    }

    #[test]
    fn test_data_bar_new() {
        let min_cfvo = Value::new(2, None);
        let max_cfvo = Value::new(3, None);
        let db = Bar::new(min_cfvo, max_cfvo, 0xFF4472C4);

        assert_eq!(db.min_cfvo.cfvo_type, 2);
        assert_eq!(db.max_cfvo.cfvo_type, 3);
        assert_eq!(db.color, 0xFF4472C4);
        assert!(db.show_value);
    }

    #[test]
    fn test_icon_set_new() {
        let cfvos = vec![
            Value::new(1, Some("0".to_string())),
            Value::new(1, Some("33".to_string())),
            Value::new(1, Some("67".to_string())),
        ];
        let icon_set = IconSet::new(0x01, cfvos); // 3Arrows

        assert_eq!(icon_set.icon_set_type, 0x01);
        assert_eq!(icon_set.cfvos.len(), 3);
        assert!(icon_set.show_value);
        assert!(!icon_set.reverse);
    }

    #[test]
    fn test_conditional_formatting_rule_new() {
        let rule = Rule::new(RuleType::CellIs, 1);

        assert_eq!(rule.rule_type, RuleType::CellIs);
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
        let cf = Formatting::new(ranges);

        assert_eq!(cf.ranges.len(), 1);
        assert_eq!(cf.ranges[0], "A1:B10");
        assert!(cf.rules.is_empty());
    }

    #[test]
    fn test_conditional_formatting_add_rule() {
        let mut cf = Formatting::new(vec!["A1:A10".to_string()]);
        let rule = Rule::new(RuleType::CellIs, 1);
        cf.add_rule(rule);

        assert_eq!(cf.rules.len(), 1);
        assert_eq!(cf.rules[0].rule_type, RuleType::CellIs);
    }

    #[test]
    fn test_conditional_formatting_rule_parse() {
        let rule = Rule::parse(&cell_rule_payload(u32::MAX, 1, false, 5)).unwrap();
        assert_eq!(rule.rule_type, RuleType::CellIs);
        assert!(rule.dxf_id.is_none());
        assert_eq!(rule.priority, 1);
        assert!(!rule.stop_if_true);
        assert_eq!(rule.operator, Some(5));
    }

    #[test]
    fn test_conditional_formatting_rule_parse_with_dxf() {
        let rule = Rule::parse(&cell_rule_payload(5, 10, true, 3)).unwrap();
        assert_eq!(rule.dxf_id, Some(5));
        assert_eq!(rule.priority, 10);
        assert!(rule.stop_if_true);
    }

    #[test]
    fn test_conditional_formatting_rule_parse_too_short() {
        let data = [0x01, 0x02, 0x03]; // too short
        let result = Rule::parse(&data);
        assert!(result.is_err());
    }

    #[test]
    fn test_conditional_formatting_rule_parse_invalid_type() {
        let data = [
            0xFF, // invalid type
            0xFF, 0xFF, 0xFF, 0xFF, 0x01, 0x00, 0x00, 0x00, 0x00,
        ];
        let result = Rule::parse(&data);
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
        assert_eq!(RuleType::CellIs as u8, 1);
        assert_eq!(RuleType::Expression as u8, 2);
        assert_eq!(RuleType::ColorScale as u8, 3);
        assert_eq!(RuleType::DataBar as u8, 4);
        assert_eq!(RuleType::TopN as u8, 5);
        assert_eq!(RuleType::IconSet as u8, 6);
    }

    #[test]
    fn test_conditional_formatting_clone() {
        let mut cf = Formatting::new(vec!["A1:A10".to_string()]);
        let rule = Rule::new(RuleType::CellIs, 1);
        cf.add_rule(rule);

        let cloned = cf.clone();
        assert_eq!(cloned.ranges.len(), cf.ranges.len());
        assert_eq!(cloned.rules.len(), cf.rules.len());
    }

    #[test]
    fn test_color_scale_clone() {
        let min_cfvo = Value::new(2, None);
        let max_cfvo = Value::new(3, None);
        let cs = Scale::new(min_cfvo, max_cfvo, 0xFFFF0000, 0xFF00FF00);
        let cloned = cs.clone();

        assert_eq!(cloned.min_color, cs.min_color);
        assert_eq!(cloned.max_color, cs.max_color);
    }
}

#[cfg(test)]
mod writer_tests {
    use super::*;

    fn fixture_cell_is_payload() -> Vec<u8> {
        let formula = TextCompiler::compile("5").unwrap();
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
        let rule = Rule::parse(&fixture_cell_is_payload()).unwrap();
        assert_eq!(rule.rule_type, RuleType::CellIs);
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
        assert!(Rule::parse(&wrong_slot).is_err());

        let mut wrong_size = fixture_cell_is_payload();
        wrong_size[30..34].copy_from_slice(&4u32.to_le_bytes());
        assert!(Rule::parse(&wrong_size).is_err());
    }

    #[test]
    fn parses_cfvo_with_ancillary_formula_losslessly() {
        let formula = TextCompiler::compile("{1,2}").unwrap();
        assert!(!formula.rgcb.is_empty());
        let mut data = Vec::new();
        data.extend_from_slice(&7u32.to_le_bytes());
        data.extend_from_slice(&0f64.to_le_bytes());
        data.extend_from_slice(&0u32.to_le_bytes());
        data.extend_from_slice(&1u32.to_le_bytes());
        data.extend_from_slice(&(formula.rgce.len() as u32).to_le_bytes());
        data.extend_from_slice(&formula.to_bytes().unwrap());
        let parsed = Value::parse(&data).unwrap();
        assert_eq!(parsed.formula_binary.as_ref().unwrap(), &formula);
    }

    #[test]
    fn extension_cfvo_roundtrips_formula_and_automatic_bounds() {
        let formula = TextCompiler::compile("$A$1").unwrap();
        let formula_value = Value {
            cfvo_type: 7,
            value: Some("$A$1".to_string()),
            numeric_value: 0.0,
            save_greater_than_or_equal: true,
            greater_than_or_equal: false,
            formula_binary: Some(formula.clone()),
        };
        let encoded = formula_value.serialize_extension14().unwrap();
        let parsed = Value::parse_extension14(&encoded).unwrap();
        assert_eq!(parsed.cfvo_type, 7);
        assert_eq!(parsed.formula_binary, Some(formula));
        assert!(!parsed.greater_than_or_equal);

        for cfvo_type in [8, 9] {
            let automatic = Value {
                cfvo_type,
                value: None,
                numeric_value: 0.0,
                save_greater_than_or_equal: false,
                greater_than_or_equal: true,
                formula_binary: None,
            };
            let encoded = automatic.serialize_extension14().unwrap();
            assert_eq!(Value::parse_extension14(&encoded).unwrap(), automatic);
        }
    }

    #[test]
    fn extension_cfvo_rejects_inconsistent_formula_metadata() {
        let formula = TextCompiler::compile("1").unwrap();
        let value = Value {
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
        assert!(Value::parse_extension14(&encoded).is_err());
    }

    #[test]
    fn parses_direct_and_theme_colors() {
        let direct = Color::parse(&[5, 0, 0, 0, 0x11, 0x22, 0x33, 0xff]).unwrap();
        assert_eq!(direct.argb, Some(0xff11_2233));
        let theme = Color::parse(&[6, 4, 0, 0, 0, 0, 0, 0]).unwrap();
        assert_eq!(theme.color_type, 3);
        assert_eq!(theme.index, 4);
        assert!(Color::parse(&[6, 12, 0, 0, 0, 0, 0, 0]).is_err());

        let theme = Color::theme(5, -1_000).unwrap();
        assert_eq!(Color::parse(&theme.to_bytes().unwrap()).unwrap(), theme);
        let mut indexed = Color::indexed(42, 0);
        indexed.tint = 2_000;
        let reparsed = Color::parse(&indexed.to_bytes().unwrap()).unwrap();
        assert_eq!(reparsed.index, 42);
        assert_eq!(reparsed.tint, 2_000);
    }

    #[test]
    fn extension_color_and_rule_guid_roundtrip_exactly() {
        let color = Color::theme(4, -2_500).unwrap();
        let encoded = color.serialize_extension14().unwrap();
        assert_eq!(Color::parse_extension14(&encoded).unwrap(), color);
        let mut malformed = encoded;
        malformed[0] = 1;
        assert!(Color::parse_extension14(&malformed).is_err());

        let guid = [0x42; 16];
        let encoded = serialize_rule_extension_guid(guid);
        assert_eq!(parse_rule_extension_guid(&encoded).unwrap(), guid);
        let mut malformed = encoded;
        malformed[3] = 1;
        assert!(parse_rule_extension_guid(&malformed).is_err());
    }

    #[test]
    fn extension_data_bar_header_preserves_flags() {
        let mut bar = Bar14::new(
            Value::new(8, None),
            Value::new(9, None),
            Color::from_argb(0xff44_72c4),
        );
        bar.min_length = 3;
        bar.max_length = 97;
        bar.show_value = false;
        bar.direction = Direction14::RightToLeft;
        bar.axis_position = AxisPosition14::Midpoint;
        bar.border = true;
        bar.custom_negative_fill = true;
        bar.unused_flags = 0xA5F0;
        let encoded = bar.serialize_header().unwrap();
        let parsed = Bar14::parse_header(&encoded).unwrap();
        assert_eq!(parsed.min_length, 3);
        assert_eq!(parsed.max_length, 97);
        assert!(!parsed.show_value);
        assert_eq!(parsed.direction, Direction14::RightToLeft);
        assert_eq!(parsed.axis_position, AxisPosition14::Midpoint);
        assert!(parsed.border);
        assert!(parsed.gradient);
        assert!(parsed.custom_negative_fill);
        assert_eq!(parsed.unused_flags, 0xA5F0);

        let mut malformed = encoded;
        malformed[6] = 2;
        assert!(Bar14::parse_header(&malformed).is_err());
    }

    #[test]
    fn extension_icon_set_and_custom_icons_roundtrip() {
        let mut set = IconSet14::new(19, vec![Value::new(1, Some("0".to_string())); 5]);
        set.show_value = false;
        set.reverse = true;
        set.unused_flags = 0x38;
        set.custom_icons = Some(vec![
            Icon {
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
            assert_eq!(Icon::parse(&encoded).unwrap(), icon);
        }
        assert!(
            Icon {
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
        let mut formatting = Formatting::new(vec!["A1:B2 C3".to_string()]);
        formatting.pivot_only = true;
        formatting.rules.push(Rule::new(RuleType::Expression, 1));
        let encoded = formatting.serialize_extension14_header().unwrap();
        let (parsed, count) = Formatting::parse_extension14_header(&encoded).unwrap();
        assert_eq!(count, 1);
        assert_eq!(parsed.ranges, ["A1:B2", "C3"]);
        assert!(parsed.pivot_only);
        assert_eq!(parsed.record_kind, RecordKind::Extension14);
    }

    #[test]
    fn extension_rule_roundtrips_two_formulas_and_ancillary_data() {
        let first = TextCompiler::compile("{1,2}").unwrap();
        let second = TextCompiler::compile("10").unwrap();
        assert!(!first.rgcb.is_empty());
        let mut rule = Rule::new(RuleType::CellIs, 7);
        rule.operator = Some(1);
        rule.parameter = 1;
        rule.formulas = vec![first.rgce.clone(), second.rgce.clone()];
        rule.formula_extras = vec![first.rgcb.clone(), second.rgcb.clone()];
        rule.dxf_id = Some(4);
        rule.extension14 = Some(RuleMetadata {
            priority: 7,
            unused: 0xA5A5_5A5A,
            guid: [0x3c; 16],
            guid_present: true,
            linked_classic_priority: None,
        });

        let encoded = rule.serialize_extension14().unwrap();
        let parsed = Rule::parse_extension14(&encoded).unwrap();
        assert_eq!(parsed.priority, 7);
        assert_eq!(parsed.operator, Some(1));
        assert_eq!(parsed.formulas, [first.rgce, second.rgce]);
        assert_eq!(parsed.formula_extras, [first.rgcb, second.rgcb]);
        assert_eq!(parsed.extension14, rule.extension14);
        assert_eq!(parsed.serialize_extension14().unwrap(), encoded);
    }

    #[test]
    fn extension_rule_preserves_signed_data_bar_linkage() {
        let mut rule = Rule::new(RuleType::DataBar, 0);
        rule.template = 0;
        rule.extension14 = Some(RuleMetadata {
            priority: -1,
            unused: 0xDEAD_BEEF,
            guid: [0x96; 16],
            guid_present: true,
            linked_classic_priority: None,
        });

        let encoded = rule.serialize_extension14().unwrap();
        let parsed = Rule::parse_extension14(&encoded).unwrap();
        assert_eq!(parsed.priority, 0);
        assert_eq!(parsed.template, 0);
        assert_eq!(parsed.extension14, rule.extension14);
        assert_eq!(parsed.serialize_extension14().unwrap(), encoded);
    }

    #[test]
    fn extension_rule_rejects_malformed_fixed_and_formula_fields() {
        let mut rule = Rule::new(RuleType::Expression, 2);
        rule.formula_texts.push("1".to_string());
        rule.extension14 = Some(RuleMetadata {
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
        assert!(Rule::parse_extension14(&reserved).is_err());

        let mut priority = encoded.clone();
        priority[fixed_offset + 12..fixed_offset + 16].copy_from_slice(&0i32.to_le_bytes());
        assert!(Rule::parse_extension14(&priority).is_err());

        let mut declared = encoded;
        declared[fixed_offset + 30..fixed_offset + 34].copy_from_slice(&999u32.to_le_bytes());
        assert!(Rule::parse_extension14(&declared).is_err());
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use super::{Bar, Bar14, IconSet, RuleMetadata, Scale};
    use crate::raw::Records;

    fn compiled(text: &str) -> ParsedFormula {
        TextCompiler::compile(text).unwrap()
    }

    #[test]
    fn fixture_rule_header_matches_libreoffice_sample() {
        let formula = compiled("5");
        let mut rule = Rule::new(RuleType::CellIs, 1);
        rule.dxf_id = Some(0);
        rule.operator = Some(5);
        rule.parameter = 5;
        rule.formulas.push(formula.rgce);
        let payload = serialize_cf_rule(&rule).unwrap();
        assert_eq!(payload.len(), 57);
        assert_eq!(u32::from_le_bytes(payload[30..34].try_into().unwrap()), 3);
        assert_eq!(&payload[42..46], &u32::MAX.to_le_bytes());
        let parsed = Rule::parse(&payload).unwrap();
        assert_eq!(parsed.operator, Some(5));
        assert_eq!(parsed.formula_texts, ["5"]);
    }

    #[test]
    fn header_preserves_pivot_and_multiple_ranges() {
        let mut formatting = Formatting::new(vec!["A1:B10".into(), "D4".into()]);
        formatting.pivot_only = true;
        let payload = serialize_cond_formatting_header(&formatting).unwrap();
        assert_eq!(u32::from_le_bytes(payload[4..8].try_into().unwrap()), 1);
        assert_eq!(u32::from_le_bytes(payload[8..12].try_into().unwrap()), 2);
    }

    #[test]
    fn writes_color_scale_data_bar_and_icon_set_subrecords() {
        let mut formatting = Formatting::new(vec!["A1:A10".into()]);
        let mut scale = Rule::new(RuleType::ColorScale, 1);
        scale.color_scale = Some(Scale::new(
            Value::new(2, None),
            Value::new(3, None),
            0xffff0000,
            0xff00ff00,
        ));
        formatting.add_rule(scale);

        let mut bar = Rule::new(RuleType::DataBar, 2);
        bar.data_bar = Some(Bar::new(
            Value::new(2, None),
            Value::new(3, None),
            0xff4472c4,
        ));
        formatting.add_rule(bar);

        let mut icons = Rule::new(RuleType::IconSet, 3);
        icons.icon_set = Some(IconSet::new(
            0,
            vec![
                Value::new(1, Some("0".into())),
                Value::new(4, Some("33".into())),
                Value::new(4, Some("67".into())),
            ],
        ));
        formatting.add_rule(icons);

        let mut bytes = Vec::new();
        write_conditional_formattings(&mut Writer::new(&mut bytes), &[formatting]).unwrap();
        let records = Records::new(&bytes);
        let mut found = Vec::new();
        for record in records {
            found.push(record.unwrap().kind());
        }
        for typ in [
            kind::BEGIN_COLOR_SCALE,
            kind::BEGIN_DATABAR,
            kind::BEGIN_ICON_SET,
            kind::CFVO,
            kind::COLOR,
        ] {
            assert!(found.contains(&typ), "record 0x{typ:04x}");
        }
    }

    #[test]
    fn rejects_duplicate_priority_and_wrong_formula_slot_count() {
        let mut first = Formatting::new(vec!["A1".into()]);
        let mut rule = Rule::new(RuleType::CellIs, 1);
        rule.operator = Some(1);
        rule.formulas.push(compiled("1").rgce);
        first.add_rule(rule.clone());
        let mut second = Formatting::new(vec!["B1".into()]);
        second.add_rule(rule);
        assert!(
            write_conditional_formattings(&mut Writer::new(Vec::new()), &[first, second]).is_err()
        );
    }

    #[test]
    fn writes_office_2013_visualization_records() {
        let mut formatting = Formatting::new(vec!["A1:A10".into()]);
        formatting.record_kind = RecordKind::Extension14;
        let mut rule = Rule::new(RuleType::DataBar, 1);
        rule.extension14 = Some(RuleMetadata {
            priority: 1,
            unused: 7,
            guid: [0x24; 16],
            guid_present: true,
            linked_classic_priority: None,
        });
        rule.data_bar14 = Some(Bar14::new(
            Value::new(8, None),
            Value::new(9, None),
            Color::from_argb(0xff44_72c4),
        ));
        formatting.add_rule(rule);

        let mut bytes = Vec::new();
        write_conditional_formattings(&mut Writer::new(&mut bytes), &[formatting]).unwrap();
        let found = Records::new(&bytes)
            .map(|record| record.unwrap().kind())
            .collect::<Vec<_>>();
        assert_eq!(
            found,
            [
                kind::BEGIN_COND_FORMATTING14,
                kind::BEGIN_CF_RULE14,
                kind::BEGIN_DATABAR14,
                kind::CFVO14,
                kind::CFVO14,
                kind::COLOR14,
                kind::COLOR14,
                kind::END_DATABAR14,
                kind::END_CF_RULE14,
                kind::END_COND_FORMATTING14,
            ]
        );
    }

    #[test]
    fn rejects_extension_cross_record_violations() {
        let mut extension = Formatting::new_extension14(vec!["A1".into()]);
        let mut rule = Rule::new(RuleType::ColorScale, 1);
        rule.extension14 = Some(RuleMetadata {
            priority: 1,
            unused: 0,
            guid: [0; 16],
            guid_present: false,
            linked_classic_priority: None,
        });
        rule.color_scale14 = Some(Scale::new(
            Value::new(8, None),
            Value::new(9, None),
            0xffff_0000,
            0xff00_ff00,
        ));
        extension.add_rule(rule);
        assert!(write_conditional_formattings(&mut Writer::new(Vec::new()), &[extension]).is_err());

        let mut classic = Formatting::new(vec!["A1".into()]);
        let mut bar = Rule::new(RuleType::DataBar, 1);
        bar.classic_extension_guid = Some([1; 16]);
        bar.data_bar = Some(Bar::new(
            Value::new(2, None),
            Value::new(3, None),
            0xff44_72c4,
        ));
        classic.add_rule(bar);
        assert!(write_conditional_formattings(&mut Writer::new(Vec::new()), &[classic]).is_err());
    }
}
