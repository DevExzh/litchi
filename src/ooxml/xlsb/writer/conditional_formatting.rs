//! Conditional formatting binary serialization for XLSB writer.
//!
//! Serializes conditional formatting as a `BrtBeginCondFormatting` (0x01CD) /
//! `BrtBeginCFRule`* (0x01CF) / `BrtEndCondFormatting` (0x01CE) block,
//! following the binary layout documented by LibreOffice `condformatbuffer.cxx`
//! and `condformatcontext.cxx`.
//!
//! # BrtBeginCondFormatting binary layout (per [MS-XLSB] 2.4.34)
//!
//! ```text
//! ccf         u32   — count of BrtBeginCFRule records
//! fPivot      u32   — 0 = normal, 1 = pivot-table only
//! sqrfx       UncheckedSqRfX (BinRangeList)
//! ```
//!
//! # BrtBeginCFRule binary layout (per LibreOffice `CondFormatRule::importCfRule`)
//!
//! ```text
//! iType       i32   — CF rule type (1=CellIs, 2=Expression, 3=ColorScale, …)
//! iSubType    i32   — CF rule sub-type (template)
//! dxfId       i32   — differential formatting ID (-1 = none)
//! iPri        i32   — priority (> 0)
//! iParam      i32   — operator / rank / stddev depending on type
//! reserved    8 bytes (zeros)
//! flags       u16   — bit 1=stopIfTrue, bit 2=aboveAverage, bit 3=bottom, bit 4=percent
//! cbFmla1     i32   — formula 1 total byte count (0 if absent, else 8+ptg_len)
//! cbFmla2     i32   — formula 2 total byte count
//! cbFmla3     i32   — formula 3 total byte count
//! strParam    XLNullableWideString — text parameter (-1 = null)
//! fmla1       BIFF12 formula (cb_ptg(i32) + ptg_bytes + cb_adddata(i32) + adddata)
//! fmla2       BIFF12 formula
//! fmla3       BIFF12 formula
//! ```

use crate::ooxml::xlsb::conditional_formatting::{
    CfRuleType, ConditionalFormatting, ConditionalFormattingRule,
};
use crate::ooxml::xlsb::error::XlsbResult;
use crate::ooxml::xlsb::records::record_types;
use crate::ooxml::xlsb::writer::RecordWriter;
use crate::ooxml::xlsb::writer::bin_range::{parse_range_list, write_bin_range_list};
use std::io::Write;

/// Write all conditional formatting blocks for a worksheet.
///
/// Each [`ConditionalFormatting`] becomes one
/// `BrtBeginCondFormatting` / `BrtBeginCFRule`* / `BrtEndCondFormatting` group.
pub fn write_conditional_formattings<W: Write>(
    writer: &mut RecordWriter<W>,
    cond_fmts: &[ConditionalFormatting],
) -> XlsbResult<()> {
    for cf in cond_fmts {
        write_single_cond_formatting(writer, cf)?;
    }
    Ok(())
}

/// Write a single conditional formatting block.
fn write_single_cond_formatting<W: Write>(
    writer: &mut RecordWriter<W>,
    cf: &ConditionalFormatting,
) -> XlsbResult<()> {
    // --- BrtBeginCondFormatting payload ---
    let header_payload = serialize_cond_formatting_header(cf)?;
    writer.write_record(record_types::BEGIN_COND_FORMATTING, &header_payload)?;

    // --- BrtBeginCFRule / BrtEndCFRule pairs ---
    for rule in &cf.rules {
        let rule_payload = serialize_cf_rule(rule)?;
        writer.write_record(record_types::BEGIN_CF_RULE, &rule_payload)?;
        writer.write_record(record_types::END_CF_RULE, &[])?;
    }

    writer.write_record(record_types::END_COND_FORMATTING, &[])?;
    Ok(())
}

/// Serialize the `BrtBeginCondFormatting` payload.
///
/// Layout per [MS-XLSB] 2.4.34:
///   `ccf(u32)` + `fPivot(u32)` + `sqrfx(UncheckedSqRfX)`
///
/// - `ccf`: count of `BrtBeginCFRule` records in this block
/// - `fPivot`: 0 = not pivot-only, 1 = pivot-table only
fn serialize_cond_formatting_header(cf: &ConditionalFormatting) -> XlsbResult<Vec<u8>> {
    let mut buf = Vec::with_capacity(64);

    // ccf (u32): count of CF rules in this block
    buf.extend_from_slice(&(cf.rules.len() as u32).to_le_bytes());
    // fPivot (u32): not pivot-table-only
    buf.extend_from_slice(&0u32.to_le_bytes());

    // BinRangeList from the range strings
    let mut all_ranges = Vec::new();
    for range_str in &cf.ranges {
        all_ranges.extend(parse_range_list(range_str)?);
    }
    write_bin_range_list(&all_ranges, &mut buf)?;

    Ok(buf)
}

/// Serialize a single `BrtBeginCFRule` payload.
///
/// Follows the field order observed in LibreOffice's `importCfRule` reader.
fn serialize_cf_rule(rule: &ConditionalFormattingRule) -> XlsbResult<Vec<u8>> {
    let mut buf = Vec::with_capacity(64);

    // iType (i32): CF rule type
    buf.extend_from_slice(&(rule.rule_type as i32).to_le_bytes());

    // iSubType (i32): CF sub-type / template
    let sub_type = cf_rule_sub_type(rule);
    buf.extend_from_slice(&sub_type.to_le_bytes());

    // dxfId (i32): differential formatting ID, -1 if absent
    let dxf_id = rule.dxf_id.map_or(-1i32, |id| id as i32);
    buf.extend_from_slice(&dxf_id.to_le_bytes());

    // iPri (i32): priority
    buf.extend_from_slice(&(rule.priority as i32).to_le_bytes());

    // iParam (i32): operator for CellIs, 0 otherwise
    let param = rule.operator.map_or(0i32, |op| op as i32);
    buf.extend_from_slice(&param.to_le_bytes());

    // reserved (8 bytes)
    buf.extend_from_slice(&[0u8; 8]);

    // flags (u16)
    let mut flags: u16 = 0;
    if rule.stop_if_true {
        flags |= 0x0002;
    }
    buf.extend_from_slice(&flags.to_le_bytes());

    // cbFmla1, cbFmla2, cbFmla3 (3 × i32)
    // Per LO comment: "For no obvious reason, the sizes of the formulas are
    // already stored before. Nevertheless the following formulas contain
    // their own sizes."  So cbFmla = total bytes consumed by importFormula(),
    // which is 4 (cb_ptg) + ptg_len + 4 (cb_adddata) + adddata_len.
    // For simple formulas with no additional data: cbFmla = 8 + ptg_len.
    let raw_ptg_sizes: [usize; 3] = [
        rule.formulas.first().map_or(0, Vec::len),
        rule.formulas.get(1).map_or(0, Vec::len),
        rule.formulas.get(2).map_or(0, Vec::len),
    ];
    for &ptg_len in &raw_ptg_sizes {
        let cb_fmla = if ptg_len > 0 { 8 + ptg_len } else { 0 };
        buf.extend_from_slice(&(cb_fmla as i32).to_le_bytes());
    }

    // strParam: XLNullableWideString — NULL = 0xFFFFFFFF (i32 = -1)
    buf.extend_from_slice(&(-1i32).to_le_bytes());

    // fmla1, fmla2, fmla3 — each in BIFF12 formula format:
    // cb_ptg(i32) + ptg_bytes + cb_adddata(i32) + adddata
    for (i, &ptg_len) in raw_ptg_sizes.iter().enumerate() {
        if ptg_len > 0 {
            // cb_ptg
            buf.extend_from_slice(&(ptg_len as i32).to_le_bytes());
            // PTG bytes
            buf.extend_from_slice(&rule.formulas[i]);
            // cb_adddata (0 = no additional data)
            buf.extend_from_slice(&0i32.to_le_bytes());
        }
    }

    Ok(buf)
}

/// Determine the BIFF12 CF rule sub-type from the rule type.
///
/// Maps [`CfRuleType`] to the `BIFF12_CFRULE_SUB_*` constants used by
/// LibreOffice (see `condformatbuffer.cxx`).
fn cf_rule_sub_type(rule: &ConditionalFormattingRule) -> i32 {
    match rule.rule_type {
        CfRuleType::CellIs => 0,     // BIFF12_CFRULE_SUB_CELLIS
        CfRuleType::Expression => 1, // BIFF12_CFRULE_SUB_EXPRESSION
        CfRuleType::ColorScale => 2, // BIFF12_CFRULE_SUB_COLORSCALE
        CfRuleType::DataBar => 3,    // BIFF12_CFRULE_SUB_DATABAR
        CfRuleType::IconSet => 4,    // BIFF12_CFRULE_SUB_ICONSET
        CfRuleType::TopN => 5,       // BIFF12_CFRULE_SUB_TOPTEN
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::ooxml::xlsb::conditional_formatting::{
        CfRuleType, ConditionalFormatting, ConditionalFormattingRule,
    };
    use crate::ooxml::xlsb::writer::RecordWriter;

    #[test]
    fn test_serialize_cond_formatting_header() {
        let cf = ConditionalFormatting {
            ranges: vec!["A1:B10".to_string()],
            rules: Vec::new(),
        };
        let payload = serialize_cond_formatting_header(&cf).unwrap();
        // ccf(4) + fPivot(4) + range_count(4) + 1 range(16) = 28
        assert_eq!(payload.len(), 28);
        // ccf = 0 (no rules in this test CF)
        assert_eq!(u32::from_le_bytes(payload[0..4].try_into().unwrap()), 0);
        // fPivot = 0
        assert_eq!(u32::from_le_bytes(payload[4..8].try_into().unwrap()), 0);
        // Range count = 1
        assert_eq!(i32::from_le_bytes(payload[8..12].try_into().unwrap()), 1);
    }

    #[test]
    fn test_serialize_cf_rule_cellis() {
        let rule = ConditionalFormattingRule {
            rule_type: CfRuleType::CellIs,
            dxf_id: Some(0),
            priority: 1,
            stop_if_true: false,
            formulas: Vec::new(),
            color_scale: None,
            data_bar: None,
            icon_set: None,
            operator: Some(3), // equal
        };
        let payload = serialize_cf_rule(&rule).unwrap();

        // iType = 1
        assert_eq!(i32::from_le_bytes(payload[0..4].try_into().unwrap()), 1);
        // iSubType = 0 (CELLIS)
        assert_eq!(i32::from_le_bytes(payload[4..8].try_into().unwrap()), 0);
        // dxfId = 0
        assert_eq!(i32::from_le_bytes(payload[8..12].try_into().unwrap()), 0);
        // iPri = 1
        assert_eq!(i32::from_le_bytes(payload[12..16].try_into().unwrap()), 1);
        // iParam = 3 (equal)
        assert_eq!(i32::from_le_bytes(payload[16..20].try_into().unwrap()), 3);
    }

    #[test]
    fn test_cf_rule_sub_type() {
        let rule_cellis = ConditionalFormattingRule {
            rule_type: CfRuleType::CellIs,
            dxf_id: None,
            priority: 1,
            stop_if_true: false,
            formulas: Vec::new(),
            color_scale: None,
            data_bar: None,
            icon_set: None,
            operator: None,
        };
        assert_eq!(cf_rule_sub_type(&rule_cellis), 0);

        let rule_expr = ConditionalFormattingRule {
            rule_type: CfRuleType::Expression,
            dxf_id: None,
            priority: 1,
            stop_if_true: false,
            formulas: Vec::new(),
            color_scale: None,
            data_bar: None,
            icon_set: None,
            operator: None,
        };
        assert_eq!(cf_rule_sub_type(&rule_expr), 1);

        let rule_colorscale = ConditionalFormattingRule {
            rule_type: CfRuleType::ColorScale,
            dxf_id: None,
            priority: 1,
            stop_if_true: false,
            formulas: Vec::new(),
            color_scale: None,
            data_bar: None,
            icon_set: None,
            operator: None,
        };
        assert_eq!(cf_rule_sub_type(&rule_colorscale), 2);

        let rule_databar = ConditionalFormattingRule {
            rule_type: CfRuleType::DataBar,
            dxf_id: None,
            priority: 1,
            stop_if_true: false,
            formulas: Vec::new(),
            color_scale: None,
            data_bar: None,
            icon_set: None,
            operator: None,
        };
        assert_eq!(cf_rule_sub_type(&rule_databar), 3);

        let rule_iconset = ConditionalFormattingRule {
            rule_type: CfRuleType::IconSet,
            dxf_id: None,
            priority: 1,
            stop_if_true: false,
            formulas: Vec::new(),
            color_scale: None,
            data_bar: None,
            icon_set: None,
            operator: None,
        };
        assert_eq!(cf_rule_sub_type(&rule_iconset), 4);

        let rule_topn = ConditionalFormattingRule {
            rule_type: CfRuleType::TopN,
            dxf_id: None,
            priority: 1,
            stop_if_true: false,
            formulas: Vec::new(),
            color_scale: None,
            data_bar: None,
            icon_set: None,
            operator: None,
        };
        assert_eq!(cf_rule_sub_type(&rule_topn), 5);
    }

    #[test]
    fn test_serialize_cf_rule_with_formulas() {
        let formula_bytes: Vec<u8> = vec![0x01, 0x02, 0x03];
        let rule = ConditionalFormattingRule {
            rule_type: CfRuleType::Expression,
            dxf_id: Some(5),
            priority: 3,
            stop_if_true: true,
            formulas: vec![formula_bytes.clone()],
            color_scale: None,
            data_bar: None,
            icon_set: None,
            operator: None,
        };

        let payload = serialize_cf_rule(&rule).unwrap();

        // Verify type = Expression (2)
        assert_eq!(i32::from_le_bytes(payload[0..4].try_into().unwrap()), 2);
        // iSubType = 1 (Expression)
        assert_eq!(i32::from_le_bytes(payload[4..8].try_into().unwrap()), 1);
        // dxfId = 5
        assert_eq!(i32::from_le_bytes(payload[8..12].try_into().unwrap()), 5);
        // iPri = 3
        assert_eq!(i32::from_le_bytes(payload[12..16].try_into().unwrap()), 3);

        // Check flags for stop_if_true
        let flags = u16::from_le_bytes(payload[28..30].try_into().unwrap());
        assert_ne!(flags & 0x0002, 0); // stopIfTrue bit

        // cbFmla1 should be 8 + ptg_len = 8 + 3 = 11
        let cb_fmla1 = i32::from_le_bytes(payload[30..34].try_into().unwrap());
        assert_eq!(cb_fmla1, 11);
    }

    #[test]
    fn test_serialize_cf_rule_no_dxf() {
        let rule = ConditionalFormattingRule {
            rule_type: CfRuleType::CellIs,
            dxf_id: None, // No differential format
            priority: 1,
            stop_if_true: false,
            formulas: Vec::new(),
            color_scale: None,
            data_bar: None,
            icon_set: None,
            operator: Some(1),
        };

        let payload = serialize_cf_rule(&rule).unwrap();
        // dxfId = -1 when not present
        assert_eq!(i32::from_le_bytes(payload[8..12].try_into().unwrap()), -1);
    }

    #[test]
    fn test_serialize_cond_formatting_header_multiple_ranges() {
        let cf = ConditionalFormatting {
            ranges: vec!["A1:B10".to_string(), "C1:D10".to_string()],
            rules: Vec::new(),
        };
        let payload = serialize_cond_formatting_header(&cf).unwrap();
        // ccf(4) + fPivot(4) + range_count(4) + 2 ranges(32) = 44
        assert_eq!(payload.len(), 44);
        // Range count = 2
        assert_eq!(i32::from_le_bytes(payload[8..12].try_into().unwrap()), 2);
    }

    #[test]
    fn test_write_conditional_formattings_empty() {
        let mut buffer = Vec::new();
        let mut writer = RecordWriter::new(&mut buffer);
        let cfs: Vec<ConditionalFormatting> = vec![];

        let result = write_conditional_formattings(&mut writer, &cfs);
        assert!(result.is_ok());
        assert!(buffer.is_empty()); // No records for empty list
    }

    #[test]
    fn test_write_single_cond_formatting() {
        let mut buffer = Vec::new();
        let mut writer = RecordWriter::new(&mut buffer);
        let cf = ConditionalFormatting::new(vec!["A1:A10".to_string()]);

        let result = write_single_cond_formatting(&mut writer, &cf);
        assert!(result.is_ok());
        assert!(!buffer.is_empty());
    }

    #[test]
    fn test_write_conditional_formattings_with_rules() {
        let mut buffer = Vec::new();
        let mut writer = RecordWriter::new(&mut buffer);

        let mut cf = ConditionalFormatting::new(vec!["A1:A10".to_string()]);
        let mut rule = ConditionalFormattingRule::new(CfRuleType::CellIs, 1);
        rule.dxf_id = Some(0);
        rule.operator = Some(2); // greater than
        cf.add_rule(rule);

        let cfs = vec![cf];
        let result = write_conditional_formattings(&mut writer, &cfs);
        assert!(result.is_ok());
        assert!(!buffer.is_empty());
    }
}
