use litchi_ole::xls::writer::{
    FillPattern, XlsConditionalFormatGroup, XlsConditionalFormatOperator,
    XlsConditionalFormatRange, XlsConditionalFormatRule, XlsConditionalFormatType,
    XlsConditionalPattern, XlsWriter,
};
use litchi_ole::xls::{XlsConditionalComparison, XlsConditionalRuleKind, XlsWorkbook};
use std::io::Cursor;

#[test]
fn grouped_ordered_legacy_rules_round_trip() {
    let mut writer = XlsWriter::new();
    let sheet = writer.add_worksheet("CF").unwrap();
    writer
        .add_conditional_format_group(
            sheet,
            XlsConditionalFormatGroup {
                ranges: vec![
                    XlsConditionalFormatRange {
                        first_row: 8,
                        last_row: 9,
                        first_col: 2,
                        last_col: 3,
                    },
                    XlsConditionalFormatRange {
                        first_row: 1,
                        last_row: 2,
                        first_col: 0,
                        last_col: 1,
                    },
                    XlsConditionalFormatRange {
                        first_row: 8,
                        last_row: 9,
                        first_col: 2,
                        last_col: 3,
                    },
                ],
                rules: vec![
                    XlsConditionalFormatRule {
                        format_type: XlsConditionalFormatType::CellValue {
                            operator: XlsConditionalFormatOperator::Between,
                            formula1: "1".into(),
                            formula2: Some("10".into()),
                        },
                        pattern: Some(XlsConditionalPattern {
                            pattern: FillPattern::Solid,
                            foreground_color: 2,
                            background_color: 3,
                        }),
                    },
                    XlsConditionalFormatRule {
                        format_type: XlsConditionalFormatType::Formula {
                            formula: "A1>0".into(),
                        },
                        pattern: None,
                    },
                ],
            },
        )
        .unwrap();
    let mut bytes = Cursor::new(Vec::new());
    writer.write_to(&mut bytes).unwrap();
    let workbook = XlsWorkbook::new(Cursor::new(bytes.into_inner())).unwrap();
    let groups = workbook.xls_worksheet(0).unwrap().conditional_formattings();
    assert_eq!(groups.len(), 1);
    assert_eq!(groups[0].ranges().len(), 3);
    assert_eq!(groups[0].ranges()[0].first_row(), 8);
    assert_eq!(groups[0].ranges()[1].first_row(), 1);
    assert_eq!(groups[0].ranges()[2].first_row(), 8);
    assert_eq!(groups[0].rules().len(), 2);
    assert_eq!(
        groups[0].rules()[0].kind(),
        XlsConditionalRuleKind::CellValue(XlsConditionalComparison::Between)
    );
    assert_eq!(groups[0].rules()[0].formula1_rendered(), Some("=1"));
    assert_eq!(groups[0].rules()[0].formula2_rendered(), Some("=10"));
    assert_eq!(groups[0].rules()[1].kind(), XlsConditionalRuleKind::Formula);
}

#[test]
fn grouped_api_rejects_cardinality_and_formula_shape() {
    let mut writer = XlsWriter::new();
    let sheet = writer.add_worksheet("CF").unwrap();
    let range = XlsConditionalFormatRange {
        first_row: 0,
        last_row: 0,
        first_col: 0,
        last_col: 0,
    };
    let bad = XlsConditionalFormatRule {
        format_type: XlsConditionalFormatType::CellValue {
            operator: XlsConditionalFormatOperator::Between,
            formula1: "1".into(),
            formula2: None,
        },
        pattern: None,
    };
    assert!(
        writer
            .add_conditional_format_group(
                sheet,
                XlsConditionalFormatGroup {
                    ranges: vec![range],
                    rules: vec![bad]
                }
            )
            .is_err()
    );
    assert!(
        writer
            .add_conditional_format_group(
                sheet,
                XlsConditionalFormatGroup {
                    ranges: Vec::new(),
                    rules: Vec::new()
                }
            )
            .is_err()
    );
}
