//! Workbook-writer serialization and round-trip tests.

use super::super::WorkbookWriter;
use crate::conditional_formatting::{
    Bar, Bar14, Color, Formatting, IconSet, RecordKind as ConditionalRecordKind, Rule,
    RuleMetadata, RuleType, Scale, Value,
};
use crate::named_ranges::{Definition, area3d_formula};
use crate::package::data_validation::{RecordKind, Settings, Validation};
use crate::raw::kind;
use crate::writer::{MutableChartSheet, MutableWorksheet, SheetProtection};
use litchi_core::sheet::{CellValue, WorkbookTrait};
use litchi_opc::PackURI;
use litchi_opc::constants::{content_type as ct, relationship_type as rel};
use std::io::Cursor;

#[test]
fn row_and_column_formatting_survive_package_roundtrip() {
    let mut workbook = WorkbookWriter::new();
    let mut sheet = MutableWorksheet::new("Layout");
    sheet.set_cell(3, 2, "value");
    sheet.set_column_width(2, 18.25);
    sheet.set_column_hidden(2, true);
    sheet.set_column_best_fit(2, true);
    sheet.set_row_height(3, 24.5);
    sheet.set_row_hidden(3, true);
    workbook.add_worksheet(sheet);

    let mut output = Cursor::new(Vec::new());
    workbook.save(&mut output).unwrap();
    let reader = crate::Workbook::new(Cursor::new(output.into_inner())).unwrap();
    let worksheet = reader.worksheet(0).unwrap();

    assert_eq!(worksheet.column_infos().len(), 1);
    let column = &worksheet.column_infos()[0];
    assert_eq!((column.first_column, column.last_column), (2, 2));
    assert_eq!(column.width, 18.25);
    assert!(column.user_set_width);
    assert!(column.hidden);
    assert!(column.best_fit);

    assert_eq!(worksheet.row_infos().len(), 1);
    let row = &worksheet.row_infos()[0];
    assert_eq!(row.row, 3);
    assert_eq!(row.height, Some(24.5));
    assert!(row.hidden);
    assert_eq!(row.column_spans, vec![(2, 2)]);
}

#[test]
fn auto_filter_range_survives_package_roundtrip() {
    let mut workbook = WorkbookWriter::new();
    let mut sheet = MutableWorksheet::new("Filtered");
    sheet.set_cell(0, 0, "Header");
    sheet.set_auto_filter(0, 20, 0, 4);
    workbook.add_worksheet(sheet);

    let mut output = Cursor::new(Vec::new());
    workbook.save(&mut output).unwrap();
    let reader = crate::Workbook::new(Cursor::new(output.into_inner())).unwrap();
    let auto_filter = reader.worksheet(0).unwrap().auto_filter().unwrap();
    assert_eq!(
        auto_filter,
        crate::sheet::AutoFilter {
            first_row: 0,
            last_row: 20,
            first_column: 0,
            last_column: 4,
        }
    );
}

#[test]
fn classic_and_extension_validations_survive_package_roundtrip() {
    let mut workbook = WorkbookWriter::new();
    let mut sheet = MutableWorksheet::new("Validated");
    sheet.set_cell(0, 0, 5);

    let mut classic = Validation::new(1, "A1:A10 C1:C10".to_string());
    classic.operator = 0;
    classic.formula1 = Some("1".to_string());
    classic.formula2 = Some("10".to_string());
    classic.ime_mode = 4;
    classic.show_input_message = true;
    classic.input_title = Some("Number".to_string());
    classic.input_text = Some("Enter 1 through 10".to_string());
    sheet.add_data_validation(classic);

    let mut extension = Validation::new(7, "B1:B20".to_string());
    extension.formula1 = Some("Source!A1>0".to_string());
    extension.record_kind = RecordKind::Extension14;
    sheet.add_data_validation(extension);
    sheet.set_data_validation_settings(Settings {
        input_prompts_disabled: true,
        prompt_x: 120,
        prompt_y: 240,
    });
    sheet.set_data_validation14_settings(Settings {
        input_prompts_disabled: false,
        prompt_x: 12,
        prompt_y: 24,
    });
    workbook.add_worksheet(sheet);
    let mut source = MutableWorksheet::new("Source");
    source.set_cell(0, 0, 1);
    workbook.add_worksheet(source);

    let mut output = Cursor::new(Vec::new());
    workbook.save(&mut output).unwrap();
    let reader = crate::Workbook::new(Cursor::new(output.into_inner())).unwrap();
    let worksheet = reader.worksheet(0).unwrap();
    assert_eq!(worksheet.data_validations().len(), 2);
    assert_eq!(
        worksheet.data_validations()[0].record_kind,
        RecordKind::Classic
    );
    assert_eq!(worksheet.data_validations()[0].cell_ranges, "A1:A10 C1:C10");
    assert_eq!(
        worksheet.data_validations()[0].formula1.as_deref(),
        Some("1")
    );
    assert_eq!(
        worksheet.data_validations()[0].formula2.as_deref(),
        Some("10")
    );
    assert_eq!(worksheet.data_validations()[0].ime_mode, 4);
    assert_eq!(
        worksheet.data_validations()[1].record_kind,
        RecordKind::Extension14
    );
    assert_eq!(
        worksheet.data_validations()[1].formula1.as_deref(),
        Some("(Source!A1>0)")
    );
    assert_eq!(
        worksheet.data_validation_settings(),
        Some(Settings {
            input_prompts_disabled: true,
            prompt_x: 120,
            prompt_y: 240,
        })
    );
    assert_eq!(
        worksheet.data_validation14_settings(),
        Some(Settings {
            input_prompts_disabled: false,
            prompt_x: 12,
            prompt_y: 24,
        })
    );
}

#[test]
fn classic_conditional_formatting_survives_package_roundtrip() {
    let mut workbook = WorkbookWriter::new();
    let mut sheet = MutableWorksheet::new("Formatted");
    sheet.set_cell(0, 0, 5);
    let mut formatting = Formatting::new(vec!["A1:A10 C1:C10".to_string()]);
    formatting.pivot_only = true;

    let mut expression = Rule::new(RuleType::Expression, 1);
    expression.formula_texts.push("Source!A1>0".to_string());
    expression.stop_if_true = true;
    formatting.add_rule(expression);

    let mut scale = Rule::new(RuleType::ColorScale, 2);
    scale.color_scale = Some(Scale::new(
        Value::new(2, None),
        Value::new(7, Some("Source!A1".to_string())),
        0xffff_0000,
        0xff00_ff00,
    ));
    formatting.add_rule(scale);

    let mut bar = Rule::new(RuleType::DataBar, 3);
    bar.data_bar = Some(Bar::new(
        Value::new(2, None),
        Value::new(3, None),
        0xff44_72c4,
    ));
    formatting.add_rule(bar);

    let mut icons = Rule::new(RuleType::IconSet, 4);
    icons.icon_set = Some(IconSet::new(
        0,
        vec![
            Value::new(1, Some("0".to_string())),
            Value::new(4, Some("33".to_string())),
            Value::new(4, Some("67".to_string())),
        ],
    ));
    formatting.add_rule(icons);
    sheet.add_conditional_formatting(formatting);
    workbook.add_worksheet(sheet);

    let mut source = MutableWorksheet::new("Source");
    source.set_cell(0, 0, 10);
    workbook.add_worksheet(source);

    let mut output = Cursor::new(Vec::new());
    workbook.save(&mut output).unwrap();
    let reader = crate::Workbook::new(Cursor::new(output.into_inner())).unwrap();
    let worksheet = reader.worksheet(0).unwrap();
    let formatting = &worksheet.conditional_formattings()[0];
    assert_eq!(formatting.ranges, ["A1:A10", "C1:C10"]);
    assert!(formatting.pivot_only);
    assert_eq!(formatting.rules.len(), 4);
    assert_eq!(formatting.rules[0].formula_texts, ["(Source!A1>0)"]);
    assert!(formatting.rules[0].stop_if_true);
    let scale = formatting.rules[1].color_scale.as_ref().unwrap();
    assert_eq!(scale.max_cfvo.value.as_deref(), Some("Source!A1"));
    assert_eq!(scale.min_color, 0xffff_0000);
    assert_eq!(scale.max_color, 0xff00_ff00);
    assert!(formatting.rules[2].data_bar.is_some());
    let icons = formatting.rules[3].icon_set.as_ref().unwrap();
    assert_eq!(icons.icon_set_type, 0);
    assert_eq!(icons.cfvos.len(), 3);
}

#[test]
fn extension_conditional_formatting_survives_package_roundtrip() {
    let mut workbook = WorkbookWriter::new();
    let mut sheet = MutableWorksheet::new("Formatted");
    sheet.set_cell(0, 0, 5);
    let mut formatting = Formatting::new(vec!["A1:A10".to_string()]);
    formatting.record_kind = ConditionalRecordKind::Extension14;

    let mut rule = Rule::new(RuleType::DataBar, 1);
    rule.extension14 = Some(RuleMetadata {
        priority: 1,
        unused: 0xCAFE_BABE,
        guid: [0x2a; 16],
        guid_present: true,
        linked_classic_priority: None,
    });
    let mut maximum = Value::new(7, Some("Source!A1".to_string()));
    maximum.greater_than_or_equal = false;
    let mut bar = Bar14::new(Value::new(8, None), maximum, Color::from_argb(0xff44_72c4));
    bar.min_length = 4;
    bar.max_length = 96;
    bar.gradient = false;
    rule.data_bar14 = Some(bar);
    formatting.add_rule(rule);
    sheet.add_conditional_formatting(formatting);
    workbook.add_worksheet(sheet);

    let mut source = MutableWorksheet::new("Source");
    source.set_cell(0, 0, 10);
    workbook.add_worksheet(source);

    let mut output = Cursor::new(Vec::new());
    workbook.save(&mut output).unwrap();
    let reader = crate::Workbook::new(Cursor::new(output.into_inner())).unwrap();
    let worksheet = reader.worksheet(0).unwrap();
    let formatting = &worksheet.conditional_formattings()[0];
    assert_eq!(formatting.record_kind, ConditionalRecordKind::Extension14);
    let rule = &formatting.rules[0];
    assert_eq!(rule.extension14.unwrap().unused, 0xCAFE_BABE);
    let bar = rule.data_bar14.as_ref().unwrap();
    assert_eq!(bar.min_cfvo.cfvo_type, 8);
    assert_eq!(bar.max_cfvo.value.as_deref(), Some("Source!A1"));
    assert!(!bar.max_cfvo.greater_than_or_equal);
    assert_eq!((bar.min_length, bar.max_length), (4, 96));
    assert!(!bar.gradient);
    assert_eq!(bar.positive_color.unwrap().argb, Some(0xff44_72c4));
}

#[test]
fn extended_data_bar_resolves_its_classic_rule_guid() {
    let guid = [0x7b; 16];
    let mut workbook = WorkbookWriter::new();
    let mut sheet = MutableWorksheet::new("Formatted");
    sheet.set_cell(0, 0, -5);

    let mut classic = Formatting::new(vec!["A1:A10".to_string()]);
    let mut classic_rule = Rule::new(RuleType::DataBar, 1);
    classic_rule.classic_extension_guid = Some(guid);
    classic_rule.data_bar = Some(Bar::new(
        Value::new(2, None),
        Value::new(3, None),
        0xff44_72c4,
    ));
    classic.add_rule(classic_rule);
    sheet.add_conditional_formatting(classic);

    let mut extension = Formatting::new(vec!["A1:A10".to_string()]);
    extension.record_kind = ConditionalRecordKind::Extension14;
    let mut extension_rule = Rule::new(RuleType::DataBar, 0);
    extension_rule.template = 0;
    extension_rule.extension14 = Some(RuleMetadata {
        priority: -1,
        unused: 0,
        guid,
        guid_present: true,
        linked_classic_priority: Some(1),
    });
    let mut bar = Bar14::new(
        Value::new(8, None),
        Value::new(9, None),
        Color::from_argb(0xff44_72c4),
    );
    bar.min_length = 0;
    bar.max_length = 100;
    bar.positive_color = None;
    bar.negative_color = Some(Color::from_argb(0xffff_0000));
    bar.custom_negative_fill = true;
    extension_rule.data_bar14 = Some(bar);
    extension.add_rule(extension_rule);
    sheet.add_conditional_formatting(extension);
    workbook.add_worksheet(sheet);

    let mut output = Cursor::new(Vec::new());
    workbook.save(&mut output).unwrap();
    let reader = crate::Workbook::new(Cursor::new(output.into_inner())).unwrap();
    let worksheet = reader.worksheet(0).unwrap();
    assert_eq!(worksheet.conditional_formattings().len(), 2);
    assert_eq!(
        worksheet.conditional_formattings()[0].rules[0].classic_extension_guid,
        Some(guid)
    );
    let extension_rule = &worksheet.conditional_formattings()[1].rules[0];
    assert_eq!(
        extension_rule.extension14.unwrap().linked_classic_priority,
        Some(1)
    );
    assert_eq!(
        extension_rule
            .data_bar14
            .as_ref()
            .unwrap()
            .negative_color
            .unwrap()
            .argb,
        Some(0xffff_0000)
    );
}

#[test]
fn sheet_protection_survives_package_roundtrip() {
    let mut workbook = WorkbookWriter::new();
    let mut sheet = MutableWorksheet::new("Protected");
    sheet.set_cell(0, 0, "locked");
    sheet.set_sheet_protection(Some(SheetProtection {
        password_hash: Some(0x5A3C),
        objects: Some(true),
        format_cells: Some(false),
        sort: Some(false),
        ..SheetProtection::default()
    }));
    workbook.add_worksheet(sheet);

    let mut output = Cursor::new(Vec::new());
    workbook.save(&mut output).unwrap();
    let reader = crate::Workbook::new(Cursor::new(output.into_inner())).unwrap();
    let protection = reader.worksheet(0).unwrap().sheet_protection().unwrap();
    assert_eq!(protection.password_hash, Some(0x5A3C));
    assert!(protection.locked);
    assert!(!protection.allow_edit_objects);
    assert!(protection.allow_edit_scenarios);
    assert!(protection.allow_format_cells);
    assert!(!protection.allow_format_columns);
    assert!(protection.allow_sort);
    assert!(protection.allow_select_locked_cells);
}

#[test]
fn sheet_range_formula_survives_package_roundtrip() {
    let mut workbook = WorkbookWriter::new();
    workbook.add_worksheet(MutableWorksheet::new("Data Sheet"));
    workbook.add_worksheet(MutableWorksheet::new("Middle"));
    let mut summary = MutableWorksheet::new("Summary");
    for col in 0..2 {
        summary.set_cell(
            0,
            col,
            CellValue::Formula {
                formula: "SUM('Data Sheet:Summary'!A1)".to_string(),
                cached_value: Some(Box::new(CellValue::Float(0.0))),
                is_array: false,
                array_range: None,
            },
        );
    }
    workbook.add_worksheet(summary);

    let mut output = Cursor::new(Vec::new());
    workbook.save(&mut output).unwrap();
    let reader = crate::Workbook::new(Cursor::new(output.into_inner())).unwrap();
    let summary = reader.worksheet_by_index(2).unwrap();
    for col in 0..2 {
        assert!(matches!(
            summary.cell_value(0, col).unwrap().as_ref(),
            CellValue::Formula { formula, .. }
                if formula == "SUM('Data Sheet:Summary'!A1)"
        ));
    }
}

#[test]
fn contextual_grouped_formulas_survive_package_roundtrip() {
    let mut workbook = WorkbookWriter::new();
    workbook.add_worksheet(MutableWorksheet::new("Data"));
    workbook.add_worksheet(MutableWorksheet::new("Middle"));
    workbook.add_named_range(
        Definition::new("Rate".to_string(), None)
            .with_formula(area3d_formula(0, 0, 0, 0, 0).unwrap()),
    );
    let mut summary = MutableWorksheet::new("Summary");
    summary
        .set_array_formula(0, 0, 0, 1, "SUM('Data:Middle'!A1)+Rate")
        .unwrap();
    summary
        .set_shared_formula(0, 2, 1, 2, "Data!A1+$A1")
        .unwrap();
    summary.set_cell(
        0,
        3,
        CellValue::Formula {
            formula: "Middle!A1".to_string(),
            cached_value: None,
            is_array: true,
            array_range: Some("D1:E1".to_string()),
        },
    );
    workbook.add_worksheet(summary);

    let mut output = Cursor::new(Vec::new());
    workbook.save(&mut output).unwrap();
    let reader = crate::Workbook::new(Cursor::new(output.into_inner())).unwrap();
    let summary = reader.worksheet_by_index(2).unwrap();
    for col in 0..=1 {
        assert!(matches!(
            summary.cell_value(0, col).unwrap().as_ref(),
            CellValue::Formula {
                formula,
                is_array: true,
                array_range: Some(range),
                ..
            } if formula == "(SUM('Data:Middle'!A1)+Rate)" && range == "A1:B1"
        ));
    }
    assert!(matches!(
        summary.cell_value(0, 2).unwrap().as_ref(),
        CellValue::Formula { formula, is_array: false, .. }
            if formula == "(Data!A1+$A1)"
    ));
    assert!(matches!(
        summary.cell_value(1, 2).unwrap().as_ref(),
        CellValue::Formula { formula, is_array: false, .. }
            if formula == "(Data!A1+$A2)"
    ));
    for col in 3..=4 {
        assert!(matches!(
            summary.cell_value(0, col).unwrap().as_ref(),
            CellValue::Formula {
                formula,
                is_array: true,
                array_range: Some(range),
                ..
            } if formula == "Middle!A1" && range == "D1:E1"
        ));
    }

    workbook.get_worksheet_mut(0).unwrap().set_name("Renamed");
    assert!(workbook.save(Cursor::new(Vec::new())).is_err());

    let mut invalid = MutableWorksheet::new("Invalid");
    assert!(
        invalid
            .set_array_formula(0, 0, 0, 0, "NOT_A_REAL_FUNCTION(1)")
            .is_err()
    );
}

#[test]
fn rejects_ambiguous_formula_metadata_before_writing() {
    let mut duplicate_sheets = WorkbookWriter::new();
    duplicate_sheets.add_worksheet(MutableWorksheet::new("Data"));
    duplicate_sheets.add_worksheet(MutableWorksheet::new("data"));
    assert!(duplicate_sheets.save(Cursor::new(Vec::new())).is_err());

    let mut invalid_sheet = WorkbookWriter::new();
    invalid_sheet.add_worksheet(MutableWorksheet::new("Data/2026"));
    assert!(invalid_sheet.save(Cursor::new(Vec::new())).is_err());

    let mut duplicate_names = WorkbookWriter::new();
    duplicate_names.add_worksheet(MutableWorksheet::new("Data"));
    duplicate_names.add_named_range(
        Definition::new("Rate".to_string(), None)
            .with_formula(area3d_formula(0, 0, 0, 0, 0).unwrap()),
    );
    duplicate_names.add_named_range(
        Definition::new("rate".to_string(), None)
            .with_formula(area3d_formula(0, 1, 1, 0, 0).unwrap()),
    );
    assert!(duplicate_names.save(Cursor::new(Vec::new())).is_err());
}

#[test]
fn contextual_formula_tokens_are_not_cached_across_saves() {
    let mut workbook = WorkbookWriter::new();
    workbook.add_worksheet(MutableWorksheet::new("Data"));
    let mut summary = MutableWorksheet::new("Summary");
    summary.set_cell(
        0,
        0,
        CellValue::Formula {
            formula: "Data!A1".to_string(),
            cached_value: None,
            is_array: false,
            array_range: None,
        },
    );
    workbook.add_worksheet(summary);
    workbook.save(Cursor::new(Vec::new())).unwrap();

    workbook.get_worksheet_mut(0).unwrap().set_name("Renamed");
    assert!(workbook.save(Cursor::new(Vec::new())).is_err());
}

#[test]
fn formula_survives_workbook_package_roundtrip() {
    let mut workbook = WorkbookWriter::new();
    let mut sheet = MutableWorksheet::new("Calculations");
    sheet.set_cell(0, 0, 2.0);
    sheet.set_cell(0, 1, 3.0);
    sheet.set_cell(
        0,
        2,
        CellValue::Formula {
            formula: "A1+B1".to_string(),
            cached_value: Some(Box::new(CellValue::Float(5.0))),
            is_array: false,
            array_range: None,
        },
    );
    sheet.set_cell(
        1,
        0,
        CellValue::Formula {
            formula: "\"result\"".to_string(),
            cached_value: Some(Box::new(CellValue::String("result".to_string()))),
            is_array: false,
            array_range: None,
        },
    );
    sheet.set_cell(
        1,
        1,
        CellValue::Formula {
            formula: "1=1".to_string(),
            cached_value: Some(Box::new(CellValue::Bool(true))),
            is_array: false,
            array_range: None,
        },
    );
    sheet.set_cell(
        1,
        2,
        CellValue::Formula {
            formula: "1/0".to_string(),
            cached_value: Some(Box::new(CellValue::Error("#DIV/0!".to_string()))),
            is_array: false,
            array_range: None,
        },
    );
    sheet.set_cell(
        2,
        0,
        CellValue::Formula {
            formula: "#REF!".to_string(),
            cached_value: Some(Box::new(CellValue::Error("#REF!".to_string()))),
            is_array: false,
            array_range: None,
        },
    );
    sheet.set_cell(
        2,
        1,
        CellValue::Formula {
            formula: "IF(TRUE,1,2)".to_string(),
            cached_value: Some(Box::new(CellValue::Float(1.0))),
            is_array: false,
            array_range: None,
        },
    );
    workbook.add_worksheet(sheet);

    let mut output = Cursor::new(Vec::new());
    workbook.save(&mut output).unwrap();
    let reader = crate::Workbook::new(Cursor::new(output.into_inner())).unwrap();
    let worksheet = reader.worksheet_by_index(0).unwrap();
    let value = worksheet.cell_value(0, 2).unwrap();
    assert!(matches!(
        value.as_ref(),
        CellValue::Formula {
            formula,
            cached_value: Some(cached),
            ..
        } if formula == "(A1+B1)" && matches!(cached.as_ref(), CellValue::Float(5.0))
    ));
    assert!(matches!(
        worksheet.cell_value(1, 0).unwrap().as_ref(),
        CellValue::Formula {
            cached_value: Some(cached),
            ..
        } if matches!(cached.as_ref(), CellValue::String(value) if value == "result")
    ));
    assert!(matches!(
        worksheet.cell_value(1, 1).unwrap().as_ref(),
        CellValue::Formula {
            cached_value: Some(cached),
            ..
        } if matches!(cached.as_ref(), CellValue::Bool(true))
    ));
    assert!(matches!(
        worksheet.cell_value(1, 2).unwrap().as_ref(),
        CellValue::Formula {
            cached_value: Some(cached),
            ..
        } if matches!(cached.as_ref(), CellValue::Error(error) if error == "#DIV/0!")
    ));
    assert!(matches!(
        worksheet.cell_value(2, 0).unwrap().as_ref(),
        CellValue::Formula {
            formula,
            cached_value: Some(cached),
            ..
        } if formula == "#REF!"
            && matches!(cached.as_ref(), CellValue::Error(error) if error == "#REF!")
    ));
    assert!(matches!(
        worksheet.cell_value(2, 1).unwrap().as_ref(),
        CellValue::Formula {
            formula,
            cached_value: Some(cached),
            ..
        } if formula == "IF(TRUE,1,2)"
            && matches!(cached.as_ref(), CellValue::Float(1.0))
    ));
}

#[test]
fn array_and_shared_formulas_survive_package_roundtrip() {
    let mut workbook = WorkbookWriter::new();
    let mut sheet = MutableWorksheet::new("Grouped formulas");
    sheet.set_cell(2, 2, 10.0);
    sheet.set_cell(3, 2, 20.0);
    sheet.set_shared_formula(2, 2, 3, 2, "B3").unwrap();
    sheet.set_array_formula(0, 4, 1, 5, "A1*2").unwrap();
    // The core CellValue representation remains a supported compatibility
    // path; the writer fills missing follower records from array_range.
    sheet.set_cell(
        5,
        0,
        CellValue::Formula {
            formula: "1+1".to_string(),
            cached_value: None,
            is_array: true,
            array_range: Some("A6:B6".to_string()),
        },
    );
    workbook.add_worksheet(sheet);

    let mut output = Cursor::new(Vec::new());
    workbook.save(&mut output).unwrap();
    let reader = crate::Workbook::new(Cursor::new(output.into_inner())).unwrap();
    let worksheet = reader.worksheet_by_index(0).unwrap();

    assert!(matches!(
        worksheet.cell_value(2, 2).unwrap().as_ref(),
        CellValue::Formula {
            formula,
            cached_value: Some(cached),
            is_array: false,
            ..
        } if formula == "B3" && matches!(cached.as_ref(), CellValue::Float(10.0))
    ));
    assert!(matches!(
        worksheet.cell_value(3, 2).unwrap().as_ref(),
        CellValue::Formula {
            formula,
            cached_value: Some(cached),
            is_array: false,
            ..
        } if formula == "B4" && matches!(cached.as_ref(), CellValue::Float(20.0))
    ));
    for row in 0..=1 {
        for col in 4..=5 {
            assert!(matches!(
                worksheet.cell_value(row, col).unwrap().as_ref(),
                CellValue::Formula {
                    formula,
                    is_array: true,
                    array_range: Some(range),
                    ..
                } if formula == "(A1*2)" && range == "E1:F2"
            ));
        }
    }
    for col in 0..=1 {
        assert!(matches!(
            worksheet.cell_value(5, col).unwrap().as_ref(),
            CellValue::Formula {
                formula,
                is_array: true,
                array_range: Some(range),
                ..
            } if formula == "(1+1)" && range == "A6:B6"
        ));
    }
}

#[test]
fn array_constant_survives_package_roundtrip() {
    let mut workbook = WorkbookWriter::new();
    let mut sheet = MutableWorksheet::new("Array constant");
    sheet.set_cell(
        0,
        0,
        CellValue::Formula {
            formula: "SUM({1,2;3,4})".to_string(),
            cached_value: Some(Box::new(CellValue::Float(10.0))),
            is_array: false,
            array_range: None,
        },
    );
    workbook.add_worksheet(sheet);

    let mut output = Cursor::new(Vec::new());
    workbook.save(&mut output).unwrap();
    let reader = crate::Workbook::new(Cursor::new(output.into_inner())).unwrap();
    let worksheet = reader.worksheet_by_index(0).unwrap();
    assert!(matches!(
        worksheet.cell_value(0, 0).unwrap().as_ref(),
        CellValue::Formula {
            formula,
            cached_value: Some(cached),
            ..
        } if formula == "SUM({1,2;3,4})"
            && matches!(cached.as_ref(), CellValue::Float(10.0))
    ));
}

#[test]
fn worksheet_charts_round_trip_through_binary_drawing_graphs() {
    use crate::chart::{Anchor, Chart};
    use crate::package::drawing::AnchorKind;
    use litchi_drawingml::chart::plot_area::TypeGroup;

    let bar = Chart::bar_chart_with_cache(
        "Quarterly sales",
        "Charts!$A$2:$A$4",
        &["Q1", "Q2", "Q3"],
        "Charts!$B$2:$B$4",
        &[10.0, 20.0, 30.0],
        Anchor::with_offsets(1, 10, 1, 20, 8, 30, 15, 40),
    )
    .unwrap();
    let line = Chart::line_chart(
        "Trend",
        "Charts!$A$2:$A$4",
        "Charts!$B$2:$B$4",
        Anchor::new(9, 1, 16, 15),
    )
    .unwrap();

    let mut sheet = MutableWorksheet::new("Charts");
    sheet.set_cell(0, 0, "Quarter");
    sheet.set_cell(0, 1, "Sales");
    sheet.add_chart(bar).unwrap();
    sheet.add_chart(line).unwrap();
    assert_eq!(sheet.charts().len(), 2);

    let pie = Chart::pie_chart(
        "Share",
        "Summary!$A$1:$A$3",
        "Summary!$B$1:$B$3",
        Anchor::new(0, 0, 7, 12),
    )
    .unwrap();
    let mut summary = MutableWorksheet::new("Summary");
    summary.add_chart(pie).unwrap();

    let mut workbook = WorkbookWriter::new();
    workbook.add_worksheet(sheet);
    workbook.add_worksheet(summary);
    let mut output = Cursor::new(Vec::new());
    workbook.save(&mut output).unwrap();
    let reader = crate::Workbook::new(Cursor::new(output.into_inner())).unwrap();

    let drawing = reader.sheet_drawing(0).expect("sheet drawing missing");
    assert_eq!(drawing.drawing.anchors.len(), 2);
    assert_eq!(drawing.charts.len(), 2);
    assert!(matches!(
        drawing.charts[0].chart.plot_area.type_groups.as_slice(),
        [TypeGroup::Bar(_)]
    ));
    assert!(matches!(
        drawing.charts[1].chart.plot_area.type_groups.as_slice(),
        [TypeGroup::Line(_)]
    ));
    let summary_drawing = reader
        .sheet_drawing(1)
        .expect("second sheet drawing missing");
    assert!(matches!(
        summary_drawing.charts[0]
            .chart
            .plot_area
            .type_groups
            .as_slice(),
        [TypeGroup::Pie(_)]
    ));
    match &drawing.drawing.anchors[0].anchor {
        AnchorKind::TwoCell { from, to, edit_as } => {
            assert_eq!((from.column, from.row), (1, 1));
            assert_eq!((from.column_offset, from.row_offset), (10, 20));
            assert_eq!((to.column, to.row), (8, 15));
            assert_eq!((to.column_offset, to.row_offset), (30, 40));
            assert!(edit_as.is_none());
        },
        other => panic!("unexpected chart anchor: {other:?}"),
    }

    let package = reader.opc_package();
    assert_eq!(
        package
            .iter_parts()
            .filter(|part| part.content_type() == ct::OFC_DRAWING)
            .count(),
        2
    );
    assert_eq!(
        package
            .iter_parts()
            .filter(|part| part.content_type() == ct::DML_CHART)
            .count(),
        3
    );
    let sheet_uri = PackURI::new("/xl/worksheets/sheet1.bin").unwrap();
    let sheet_part = package.get_part(&sheet_uri).unwrap();
    let drawing_record = crate::raw::Records::new(sheet_part.blob())
        .find_map(|record| {
            let record = record.unwrap();
            (record.kind() == kind::DRAWING).then_some(record)
        })
        .expect("BrtDrawing missing");
    let mut cursor = crate::raw::Cursor::new(drawing_record.payload(), "BrtDrawing");
    let drawing_rel_id = cursor.read_wide_string().unwrap();
    cursor.finish().unwrap();
    let relationship = sheet_part.rels().get(&drawing_rel_id).unwrap();
    assert_eq!(relationship.reltype(), rel::DRAWING);
    assert!(!relationship.is_external());
}

#[test]
fn chart_resource_graphs_round_trip_for_worksheets_and_chart_sheets() {
    use crate::chart::{
        Anchor, Chart, ExternalDataPart, ExternalDataTarget, Relationship, Target, UserShapesPart,
    };
    use litchi_drawingml::chart::{ExtensionList, ShapeProperties};

    let mut worksheet_chart = Chart::bar_chart(
        "Resources",
        "Data!$A$1:$A$2",
        "Data!$B$1:$B$2",
        Anchor::new(1, 1, 8, 15),
    )
    .unwrap();
    worksheet_chart.chart.shape_properties = Some(
            ShapeProperties::from_xml(
                br#"<c:spPr xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><a:blipFill><a:blip r:embed="rId9"/></a:blipFill></c:spPr>"#.to_vec(),
            )
            .unwrap(),
        );
    worksheet_chart.chart.extension_list = Some(
            ExtensionList::from_xml(
                br#"<c:extLst xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:x="urn:example"><c:ext uri="resources"><x:reference r:id="rId1" r:link="rId10"/></c:ext></c:extLst>"#.to_vec(),
            )
            .unwrap(),
        );
    worksheet_chart = worksheet_chart
            .with_additional_relationship(Relationship {
                relationship_id: "rId9".to_string(),
                relationship_type: rel::IMAGE.to_string(),
                target: Target::Embedded {
                    data: b"chart background".to_vec(),
                    content_type: "image/png".to_string(),
                    extension: "png".to_string(),
                },
            })
            .with_additional_relationship(Relationship {
                relationship_id: "rId10".to_string(),
                relationship_type: rel::HYPERLINK.to_string(),
                target: Target::External {
                    target: "https://example.test/chart".to_string(),
                },
            })
            .with_external_data_part(
                ExternalDataPart::embedded_workbook(b"PK chart workbook".to_vec()),
                Some(false),
            )
            .with_user_shapes_part(UserShapesPart {
                xml: br#"<c:userShapes xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:cdr="http://schemas.openxmlformats.org/drawingml/2006/chartDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><cdr:relSizeAnchor><cdr:from><cdr:x>0</cdr:x><cdr:y>0</cdr:y></cdr:from><cdr:to><cdr:x>1</cdr:x><cdr:y>1</cdr:y></cdr:to><cdr:pic><a:blip r:embed="rId5"/></cdr:pic></cdr:relSizeAnchor></c:userShapes>"#.to_vec(),
                relationships: vec![Relationship {
                    relationship_id: "rId5".to_string(),
                    relationship_type: rel::IMAGE.to_string(),
                    target: Target::Embedded {
                        data: b"shape image".to_vec(),
                        content_type: "image/png".to_string(),
                        extension: "png".to_string(),
                    },
                }],
            });

    let chart_sheet_chart = Chart::line_chart(
        "Linked",
        "Data!$A$1:$A$2",
        "Data!$B$1:$B$2",
        Anchor::new(0, 0, 5, 10),
    )
    .unwrap()
    .with_external_data_part(
        ExternalDataPart::linked_package("https://example.test/data.xlsx"),
        Some(true),
    );

    let mut workbook = WorkbookWriter::new();
    let mut data = MutableWorksheet::new("Data");
    data.add_chart(worksheet_chart).unwrap();
    workbook.add_worksheet(data);
    workbook
        .add_chart_sheet(MutableChartSheet::new("Linked Chart", chart_sheet_chart))
        .unwrap();

    let mut output = Cursor::new(Vec::new());
    workbook.save(&mut output).unwrap();
    let reader = crate::Workbook::new(Cursor::new(output.into_inner())).unwrap();

    let worksheet_chart = &reader.sheet_drawing(0).unwrap().charts[0];
    match &worksheet_chart.external_data_part.as_ref().unwrap().target {
        ExternalDataTarget::Embedded { data, .. } => {
            assert_eq!(data, b"PK chart workbook");
        },
        other => panic!("unexpected worksheet chart external data: {other:?}"),
    }
    let user_shapes = worksheet_chart.user_shapes_part.as_ref().unwrap();
    assert_eq!(user_shapes.relationships.len(), 1);
    match &user_shapes.relationships[0].target {
        Target::Embedded { data, .. } => {
            assert_eq!(data, b"shape image");
        },
        other => panic!("unexpected user-shapes target: {other:?}"),
    }
    assert_eq!(worksheet_chart.additional_relationships.len(), 2);
    let background = worksheet_chart
        .additional_relationships
        .iter()
        .find(|relationship| relationship.relationship_id == "rId9")
        .unwrap();
    match &background.target {
        Target::Embedded { data, .. } => {
            assert_eq!(data, b"chart background");
        },
        other => panic!("unexpected background target: {other:?}"),
    }
    let hyperlink = worksheet_chart
        .additional_relationships
        .iter()
        .find(|relationship| relationship.relationship_id == "rId10")
        .unwrap();
    match &hyperlink.target {
        Target::External { target } => {
            assert_eq!(target, "https://example.test/chart");
        },
        other => panic!("unexpected hyperlink target: {other:?}"),
    }

    let chart_sheet_chart = &reader.sheet_drawing(1).unwrap().charts[0];
    match &chart_sheet_chart
        .external_data_part
        .as_ref()
        .unwrap()
        .target
    {
        ExternalDataTarget::Linked { target } => {
            assert_eq!(target, "https://example.test/data.xlsx");
        },
        other => panic!("unexpected chart-sheet external data: {other:?}"),
    }
    assert!(chart_sheet_chart.user_shapes_part.is_none());
    assert!(chart_sheet_chart.additional_relationships.is_empty());
}

#[test]
fn worksheet_chart_validation_and_crud_are_lossless_or_refuse() {
    use crate::chart::{Anchor, Chart, Relationship, Target, UserShapesPart};

    let mut sheet = MutableWorksheet::new("Charts");
    let valid = Chart::bar_chart(
        "Valid",
        "Charts!$A$1:$A$2",
        "Charts!$B$1:$B$2",
        Anchor::new(1, 1, 8, 15),
    )
    .unwrap();
    sheet.add_chart(valid.clone()).unwrap();

    let mut descending = valid.clone();
    descending.anchor.to_col = 0;
    assert!(sheet.add_chart(descending).is_err());
    assert_eq!(sheet.charts().len(), 1);

    let mut mismatched_external_data = valid.clone();
    mismatched_external_data.chart.external_data =
        Some(litchi_drawingml::chart::ExternalData::pending());
    assert!(sheet.add_chart(mismatched_external_data).is_err());
    assert_eq!(sheet.charts().len(), 1);

    let invalid_user_shapes = valid.clone().with_user_shapes_part(UserShapesPart::new(
            br#"<c:userShapes xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><a:blip r:embed="rId5"/></c:userShapes>"#.to_vec(),
        ));
    assert!(sheet.add_chart(invalid_user_shapes).is_err());
    let invalid_relationship = valid.clone().with_additional_relationship(Relationship {
        relationship_id: "not an id".to_string(),
        relationship_type: rel::HYPERLINK.to_string(),
        target: Target::External {
            target: "https://example.test".to_string(),
        },
    });
    assert!(sheet.add_chart(invalid_relationship).is_err());
    assert_eq!(sheet.charts().len(), 1);

    let removed = sheet.remove_chart(0).unwrap();
    assert_eq!(removed.anchor.from_col, 1);
    assert!(sheet.charts().is_empty());
    assert!(sheet.remove_chart(0).is_err());
    sheet.add_chart(valid).unwrap();
    sheet.clear_charts();
    assert!(sheet.charts().is_empty());
}

#[test]
fn worksheet_images_round_trip_with_charts_in_one_drawing_graph() {
    use crate::chart::{Anchor, Chart};
    use crate::package::drawing::{AnchorKind, Object};
    use crate::shapes::{DrawingObject, Emu, EmuExtent, EmuOffset, Preset, ShapeAnchor};
    use crate::writer::{Image, ImageFormat};

    const PNG_1X1: &[u8] = &[
        0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44,
        0x52, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1F,
        0x15, 0xC4, 0x89, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x44, 0x41, 0x54, 0x08, 0xD7, 0x63, 0xF8,
        0xCF, 0xC0, 0xF0, 0x1F, 0x00, 0x05, 0x00, 0x01, 0xFF, 0x89, 0x99, 0x03, 0x5D, 0x00, 0x00,
        0x00, 0x00, 0x49, 0x45, 0x4E, 0x44, 0xAE, 0x42, 0x60, 0x82,
    ];
    const SVG: &[u8] =
            br#"<svg xmlns="http://www.w3.org/2000/svg" width="1" height="1"><path d="M0 0h1v1z"/></svg>"#;
    const GIF_1X1: &[u8] = &[
        0x47, 0x49, 0x46, 0x38, 0x39, 0x61, 0x01, 0x00, 0x01, 0x00, 0x80, 0x00, 0x00, 0x00, 0x00,
        0x00, 0xFF, 0xFF, 0xFF, 0x21, 0xF9, 0x04, 0x01, 0x00, 0x00, 0x00, 0x00, 0x2C, 0x00, 0x00,
        0x00, 0x00, 0x01, 0x00, 0x01, 0x00, 0x00, 0x02, 0x02, 0x44, 0x01, 0x00, 0x3B,
    ];

    let png_anchor = Anchor::with_offsets(1, 10, 2, 20, 5, 30, 8, 40);
    let png = Image::new(PNG_1X1.to_vec(), ImageFormat::Png, png_anchor)
        .unwrap()
        .with_description("Logo & <mark>")
        .unwrap();
    let svg = Image::new(SVG.to_vec(), ImageFormat::Svg, Anchor::new(6, 2, 9, 8)).unwrap();
    let chart = Chart::line_chart(
        "Trend",
        "Pictures!$A$1:$A$2",
        "Pictures!$B$1:$B$2",
        Anchor::new(10, 2, 17, 16),
    )
    .unwrap();

    let mut sheet = MutableWorksheet::new("Pictures");
    sheet.add_image(png).unwrap();
    sheet.add_image(svg).unwrap();
    sheet.add_chart(chart).unwrap();
    sheet
        .add_text_box(
            "Caption",
            ShapeAnchor::Absolute {
                position: EmuOffset {
                    x: Emu(100_000),
                    y: Emu(100_000),
                },
                extent: EmuExtent {
                    width: Emu(1_000_000),
                    height: Emu(500_000),
                },
            },
            Preset::Rect,
            "Mixed drawing",
        )
        .unwrap();
    let mut image_only = MutableWorksheet::new("Image only");
    image_only
        .add_image(Image::new(GIF_1X1.to_vec(), ImageFormat::Gif, Anchor::new(0, 0, 2, 3)).unwrap())
        .unwrap();
    let mut workbook = WorkbookWriter::new();
    workbook.add_worksheet(sheet);
    workbook.add_worksheet(image_only);
    let mut output = Cursor::new(Vec::new());
    workbook.save(&mut output).unwrap();

    let reader = crate::Workbook::new(Cursor::new(output.into_inner())).unwrap();
    let drawing = reader.sheet_drawing(0).expect("sheet drawing missing");
    assert_eq!(drawing.images.len(), 2);
    assert_eq!(drawing.charts.len(), 1);
    assert_eq!(drawing.images[0].format, ImageFormat::Png);
    assert_eq!(drawing.images[0].data.as_ref(), PNG_1X1);
    assert_eq!(
        drawing.images[0].description.as_deref(),
        Some("Logo & <mark>")
    );
    assert_eq!(drawing.images[1].format, ImageFormat::Svg);
    assert_eq!(drawing.images[1].data.as_ref(), SVG);
    assert_eq!(drawing.images[0].rel_id, "rId1");
    assert_eq!(drawing.images[1].rel_id, "rId2");
    assert_eq!(drawing.charts[0].rel_id, "rId3");
    assert_eq!(drawing.drawing.anchors.len(), 4);
    assert_eq!(drawing.shapes.len(), 1);
    let DrawingObject::Shape(caption) = &drawing.shapes[0].object else {
        panic!("expected mixed-drawing caption");
    };
    assert_eq!(caption.non_visual.id, Some(4));
    match &drawing.drawing.anchors[0].anchor {
        AnchorKind::TwoCell { from, to, .. } => {
            assert_eq!(
                (from.column, from.column_offset, from.row, from.row_offset),
                (1, 10, 2, 20)
            );
            assert_eq!(
                (to.column, to.column_offset, to.row, to.row_offset),
                (5, 30, 8, 40)
            );
        },
        other => panic!("unexpected image anchor: {other:?}"),
    }
    assert!(matches!(
        &drawing.drawing.anchors[0].object,
        Object::Picture {
            embed_rel_id: Some(rel_id),
            ..
        } if rel_id == "rId1"
    ));
    assert!(matches!(
        &drawing.drawing.anchors[2].object,
        Object::GraphicFrame(frame)
            if frame.rel_id.as_deref() == Some("rId3")
    ));
    let second_drawing = reader
        .sheet_drawing(1)
        .expect("image-only sheet drawing missing");
    assert_eq!(second_drawing.images.len(), 1);
    assert!(second_drawing.charts.is_empty());
    assert_eq!(second_drawing.images[0].format, ImageFormat::Gif);
    assert_eq!(second_drawing.images[0].data.as_ref(), GIF_1X1);

    let package = reader.opc_package();
    assert_eq!(
        package
            .iter_parts()
            .filter(|part| matches!(part.content_type(), ct::PNG | ct::GIF | "image/svg+xml"))
            .count(),
        3
    );
    for part_name in [
        "/xl/media/image1.png",
        "/xl/media/image2.svg",
        "/xl/media/image3.gif",
    ] {
        assert!(package.get_part(&PackURI::new(part_name).unwrap()).is_ok());
    }
}

#[test]
fn worksheet_image_validation_and_crud_are_lossless_or_refuse() {
    use crate::chart::Anchor;
    use crate::writer::{Image, ImageFormat};

    assert!(
        Image::new(
            b"not a png".to_vec(),
            ImageFormat::Png,
            Anchor::new(0, 0, 1, 1),
        )
        .is_err()
    );
    assert!(
        Image::new(
            b"<not-svg/>".to_vec(),
            ImageFormat::Svg,
            Anchor::new(0, 0, 1, 1),
        )
        .is_err()
    );
    assert!(
        Image::new(
            b"GIF89a".to_vec(),
            ImageFormat::Gif,
            Anchor::new(2, 2, 1, 1),
        )
        .is_err()
    );

    let valid = Image::new(
        b"GIF89a".to_vec(),
        ImageFormat::Gif,
        Anchor::new(0, 0, 1, 1),
    )
    .unwrap();
    let mut sheet = MutableWorksheet::new("Pictures");
    sheet.add_image(valid.clone()).unwrap();
    assert_eq!(sheet.images().len(), 1);
    assert!(
        valid
            .clone()
            .with_description("invalid\u{0}description")
            .is_err()
    );
    assert_eq!(sheet.images().len(), 1);
    let removed = sheet.remove_image(0).unwrap();
    assert_eq!(removed.format(), ImageFormat::Gif);
    assert!(sheet.images().is_empty());
    assert!(sheet.remove_image(0).is_err());
    sheet.add_image(valid).unwrap();
    sheet.clear_images();
    assert!(sheet.images().is_empty());
}

#[test]
fn worksheet_shapes_groups_and_connectors_round_trip() {
    use crate::shapes::{
        CellMarker, DrawingObject, EditAs, Emu, EmuExtent, EmuOffset, GroupTransform, Preset,
        ShapeAnchor, TextSize,
    };
    use crate::writer::{ConnectionEndSpec, ConnectionShapeSpec, GroupSpec, ShapeSpec};

    fn marker(column: u32, row: u32) -> CellMarker {
        CellMarker {
            column,
            row,
            column_offset: Emu(0),
            row_offset: Emu(0),
        }
    }

    let two_cell = ShapeAnchor::TwoCell {
        from: marker(0, 0),
        to: marker(3, 4),
        edit_as: EditAs::OneCell,
    };
    let child_anchor = ShapeAnchor::TwoCell {
        from: marker(0, 0),
        to: marker(1, 1),
        edit_as: EditAs::TwoCell,
    };
    let mut standalone = ShapeSpec::text_box("Standalone", two_cell, Preset::RoundRect, "A\nB");
    standalone.description = Some("Typed XLSB text box".to_string());
    standalone.paragraphs[0].runs[0].bold = Some(true);
    standalone.paragraphs[0].runs[0].font_size = Some(TextSize::new(1_400).unwrap());

    let group_anchor = ShapeAnchor::OneCell {
        from: marker(4, 1),
        extent: EmuExtent {
            width: Emu(4_000_000),
            height: Emu(2_000_000),
        },
    };
    let mut group = GroupSpec::new("Pair", group_anchor)
        .with_child(ShapeSpec::shape("Left", child_anchor, Preset::Rect, "L").into())
        .with_child(ShapeSpec::shape("Right", child_anchor, Preset::Ellipse, "R").into());
    group.transform = Some(GroupTransform {
        offset: Some(EmuOffset {
            x: Emu(0),
            y: Emu(0),
        }),
        extent: Some(EmuExtent {
            width: Emu(4_000_000),
            height: Emu(2_000_000),
        }),
        child_offset: Some(EmuOffset {
            x: Emu(0),
            y: Emu(0),
        }),
        child_extent: Some(EmuExtent {
            width: Emu(4_000_000),
            height: Emu(2_000_000),
        }),
    });

    let connection = ConnectionShapeSpec::new(
        "Bridge",
        ShapeAnchor::Absolute {
            position: EmuOffset {
                x: Emu(500_000),
                y: Emu(500_000),
            },
            extent: EmuExtent {
                width: Emu(1_000_000),
                height: Emu(1_000_000),
            },
        },
        Preset::StraightConnector1,
        ConnectionEndSpec {
            shape_name: "Left".to_string(),
            site: 1,
        },
        ConnectionEndSpec {
            shape_name: "Right".to_string(),
            site: 2,
        },
    );

    let mut sheet = MutableWorksheet::new("Shapes");
    sheet.add_shape(standalone).unwrap();
    sheet.add_group(group).unwrap();
    sheet.add_connection(connection).unwrap();
    let mut workbook = WorkbookWriter::new();
    workbook.add_worksheet(sheet);
    let mut output = Cursor::new(Vec::new());
    workbook.save(&mut output).unwrap();

    let reader = crate::Workbook::new(Cursor::new(output.into_inner())).unwrap();
    let drawing = reader.sheet_drawing(0).expect("shape drawing missing");
    assert!(drawing.images.is_empty());
    assert!(drawing.charts.is_empty());
    assert_eq!(drawing.drawing.anchors.len(), 3);
    assert_eq!(drawing.shapes.len(), 3);

    let DrawingObject::Shape(shape) = &drawing.shapes[0].object else {
        panic!("expected standalone shape");
    };
    assert_eq!(shape.non_visual.id, Some(1));
    assert_eq!(shape.non_visual.name.as_deref(), Some("Standalone"));
    assert_eq!(
        shape.non_visual.description.as_deref(),
        Some("Typed XLSB text box")
    );
    assert_eq!(shape.text_body.as_ref().unwrap().text(), "A\nB");
    assert_eq!(
        shape.text_body.as_ref().unwrap().paragraphs[0].runs[0].bold,
        Some(true)
    );

    let DrawingObject::Group(group) = &drawing.shapes[1].object else {
        panic!("expected shape group");
    };
    assert_eq!(group.non_visual.id, Some(2));
    assert_eq!(group.children.len(), 2);
    let DrawingObject::Shape(left) = &group.children[0] else {
        panic!("expected left group child");
    };
    let DrawingObject::Shape(right) = &group.children[1] else {
        panic!("expected right group child");
    };
    assert_eq!(left.non_visual.id, Some(3));
    assert_eq!(right.non_visual.id, Some(4));

    let DrawingObject::ConnectionShape(connection) = &drawing.shapes[2].object else {
        panic!("expected connection shape");
    };
    assert_eq!(connection.non_visual.id, Some(5));
    assert_eq!(connection.start.unwrap().shape_id, 3);
    assert_eq!(connection.end.unwrap().shape_id, 4);

    let drawing_part = reader
        .opc_package()
        .iter_parts()
        .find(|part| part.content_type() == ct::OFC_DRAWING)
        .unwrap();
    assert!(drawing_part.rels().is_empty());
}

#[test]
fn worksheet_shape_crud_and_save_validation_are_lossless_or_refuse() {
    use crate::shapes::{CellMarker, Columns, EditAs, Emu, Preset, ShapeAnchor, TextSize};
    use crate::writer::{ConnectionEndSpec, ConnectionShapeSpec, GroupSpec, ShapeSpec};

    fn anchor(from: (u32, u32), to: (u32, u32)) -> ShapeAnchor {
        let marker = |(column, row)| CellMarker {
            column,
            row,
            column_offset: Emu(0),
            row_offset: Emu(0),
        };
        ShapeAnchor::TwoCell {
            from: marker(from),
            to: marker(to),
            edit_as: EditAs::TwoCell,
        }
    }

    let valid = ShapeSpec::shape("Target", anchor((0, 0), (2, 2)), Preset::Rect, "target");
    let invalid = ShapeSpec::shape("Descending", anchor((2, 2), (1, 1)), Preset::Rect, "");
    let mut sheet = MutableWorksheet::new("Shapes");
    assert!(sheet.add_shape(invalid).is_err());
    assert!(sheet.shapes().is_empty());
    let mut invalid_xml = valid.clone();
    invalid_xml.name = "invalid\u{0}name".to_string();
    assert!(sheet.add_shape(invalid_xml).is_err());
    assert!(TextSize::new(0).is_err());
    assert!(Columns::new(17).is_err());
    assert!(sheet.shapes().is_empty());
    sheet.add_shape(valid.clone()).unwrap();
    sheet
        .add_group(GroupSpec::new("Group", anchor((3, 3), (6, 6))).with_child(
            ShapeSpec::shape("Nested", anchor((3, 3), (4, 4)), Preset::Ellipse, "").into(),
        ))
        .unwrap();
    sheet
        .add_connection(ConnectionShapeSpec::new(
            "Dangling",
            anchor((1, 1), (2, 2)),
            Preset::StraightConnector1,
            ConnectionEndSpec {
                shape_name: "Missing".to_string(),
                site: 0,
            },
            ConnectionEndSpec {
                shape_name: "Target".to_string(),
                site: 0,
            },
        ))
        .unwrap();

    let mut workbook = WorkbookWriter::new();
    workbook.add_worksheet(sheet);
    assert!(workbook.save(&mut Cursor::new(Vec::new())).is_err());
    let sheet = workbook.get_worksheet_mut(0).unwrap();
    assert_eq!(sheet.shapes().len(), 1);
    assert_eq!(sheet.groups().len(), 1);
    assert_eq!(sheet.connections().len(), 1);
    assert!(sheet.remove_shape(4).is_err());
    assert!(sheet.remove_group(4).is_err());
    assert!(sheet.remove_connection(4).is_err());
    sheet.remove_connection(0).unwrap();
    assert!(workbook.save(&mut Cursor::new(Vec::new())).is_ok());
    let sheet = workbook.get_worksheet_mut(0).unwrap();
    sheet.clear_drawing_shapes();
    assert!(sheet.shapes().is_empty());
    assert!(sheet.groups().is_empty());
    assert!(sheet.connections().is_empty());
    sheet.add_shape(valid).unwrap();
    assert_eq!(sheet.remove_shape(0).unwrap().name, "Target");
}
