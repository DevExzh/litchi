//! Regression coverage for the conditional-format owner.

use std::collections::HashSet;

use super::codec::{parse_cf, parse_condfmt};
use super::model::{CF_RECORD_TYPE, CONDFMT_RECORD_TYPE};
use super::*;
use litchi_core::sheet::WorkbookTrait;

fn header(rule_count: u16) -> Vec<u8> {
    let mut data = Vec::new();
    data.extend_from_slice(&rule_count.to_le_bytes());
    data.extend_from_slice(&0u16.to_le_bytes());
    data.extend_from_slice(&[0, 0, 7, 0, 0, 0, 0, 0]);
    data.extend_from_slice(&1u16.to_le_bytes());
    data.extend_from_slice(&[0, 0, 7, 0, 0, 0, 0, 0]);
    data
}

fn rule(condition: u8, operator: u8, formula1: &[u8], formula2: &[u8]) -> Vec<u8> {
    let mut data = vec![condition, operator];
    data.extend_from_slice(&(formula1.len() as u16).to_le_bytes());
    data.extend_from_slice(&(formula2.len() as u16).to_le_bytes());
    data.extend_from_slice(&0x003f_ffffu32.to_le_bytes());
    data.extend_from_slice(&0x8002u16.to_le_bytes());
    data.extend_from_slice(formula1);
    data.extend_from_slice(formula2);
    data
}

#[test]
fn parses_and_collects_legacy_rules() {
    let mut collector = ConditionalFormatCollector::new();
    collector
        .feed_record(CONDFMT_RECORD_TYPE, &header(2), None)
        .unwrap();
    collector
        .feed_record(
            CF_RECORD_TYPE,
            &rule(1, 1, &[0x1e, 1, 0], &[0x1e, 5, 0]),
            None,
        )
        .unwrap();
    collector
        .feed_record(CF_RECORD_TYPE, &rule(2, 0, &[0x1d], &[]), None)
        .unwrap();
    let groups = collector.finish().unwrap().0;
    assert_eq!(groups[0].rules().len(), 2);
    assert_eq!(
        groups[0].rules()[0].kind(),
        ConditionalRuleKind::CellValue(ConditionalComparison::Between)
    );
    assert_eq!(groups[0].rules()[1].formula1_tokens(), &[0x1d]);
}

#[test]
fn rejects_malformed_ranges_formulas_and_sequences() {
    assert!(parse_condfmt(&header(0)).is_err());
    assert!(parse_cf(&rule(2, 1, &[0x1d], &[]), None).is_err());
    assert!(parse_cf(&rule(1, 5, &[0x1e, 1, 0], &[0x1e, 2, 0]), None).is_err());

    let mut collector = ConditionalFormatCollector::new();
    assert!(
        collector
            .feed_record(CF_RECORD_TYPE, &rule(2, 0, &[0x1d], &[]), None)
            .is_err()
    );
    let mut collector = ConditionalFormatCollector::new();
    collector
        .feed_record(CONDFMT_RECORD_TYPE, &header(1), None)
        .unwrap();
    assert!(collector.feed_record(0x000a, &[], None).is_err());
}

#[test]
fn reads_poi_legacy_conditional_formatting_fixture() {
    use crate::Workbook;
    use std::fs::File;
    use std::path::Path;

    let fixture = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/poi/test-data/spreadsheet/WithConditionalFormatting.xls");
    let workbook = Workbook::new(File::open(fixture).unwrap()).unwrap();
    let groups = workbook.xls_worksheet(0).unwrap().conditional_formattings();
    assert_eq!(groups.len(), 3);
    assert_eq!(groups[0].rules().len(), 2);
    assert_eq!(groups[0].ranges()[0].last_row(), 7);
    assert_eq!(
        groups[0].rules()[0].kind(),
        ConditionalRuleKind::CellValue(ConditionalComparison::GreaterThan)
    );
    assert!(groups[0].rules()[0].style().font().is_some());
    assert_eq!(groups[1].rules()[0].kind(), ConditionalRuleKind::Formula);
    assert!(
        groups
            .iter()
            .flat_map(|group| group.rules())
            .any(|rule| rule.formula1_rendered().is_some())
    );
    assert_eq!(
        groups[2].rules()[1].kind(),
        ConditionalRuleKind::CellValue(ConditionalComparison::Between)
    );
    assert!(!groups[2].rules()[1].formula2_tokens().is_empty());
}

#[test]
fn reads_poi_future_conditional_formatting_fixture() {
    use crate::Workbook;
    use std::fs::File;
    use std::path::Path;
    let fixture = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/poi/test-data/spreadsheet/NewStyleConditionalFormattings.xls");
    let workbook = Workbook::new(File::open(fixture).unwrap()).unwrap();
    let mut count = 0usize;
    let mut priorities = HashSet::new();
    for sheet in 0..workbook.worksheet_count() {
        let worksheet = workbook.xls_worksheet(sheet).unwrap();
        for group in worksheet.conditional_formattings12() {
            assert!(!group.ranges().is_empty());
            for rule in group.rules() {
                assert!(priorities.insert(rule.priority()));
                assert!(rule.differential_format().len() >= 6);
                count += 1;
            }
        }
        for extension in worksheet.conditional_format_extensions() {
            assert!(priorities.insert(extension.priority()));
            count += 1;
        }
    }
    assert!(count > 0);
}
