use std::fs::File;
use std::io::Cursor;
use std::path::{Path, PathBuf};

use litchi_cfb::OleFile;
use litchi_xls::writer::{DefinedNameRecordOptions, FunctionGroupOptions, Writer};
use litchi_xls::{
    BuiltInFunctionCategories, DefinedNameFutureRecords, DefinedNameKind, NameFnGrp12, NamePublish,
    NameScope, Workbook,
};

fn fixture(name: &str) -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/poi/test-data/spreadsheet")
        .join(name)
}

fn future_record_file() -> Vec<u8> {
    let mut writer = Writer::new();
    writer.add_worksheet("Sheet1").unwrap();
    let categories = (0..19).map(|index| format!("Category{index}")).collect();
    writer
        .set_function_groups(FunctionGroupOptions {
            built_in: BuiltInFunctionCategories::Fourteen,
            custom_categories: categories,
        })
        .unwrap();
    writer
        .add_defined_name_record_with_future_records(
            DefinedNameRecordOptions {
                name: "ΣRate".to_string(),
                kind: DefinedNameKind::User,
                scope: NameScope::Workbook,
                hidden: true,
                function: true,
                vba_procedure: false,
                procedure: true,
                calculated_expression: true,
                function_group: 7,
                published: true,
                workbook_parameter: false,
                shortcut_key: None,
                formula_tokens: vec![],
                formula_extra: vec![],
                custom_menu: String::new(),
                description: String::new(),
                help_topic: String::new(),
                status_bar: String::new(),
                comment: Some("x".to_string()),
            },
            DefinedNameFutureRecords {
                function_group: Some(NameFnGrp12 {
                    function_name: "σRate".to_string(),
                    category: 32,
                }),
                publication: Some(NamePublish {
                    published: false,
                    workbook_parameter: true,
                    name: "σRate".to_string(),
                }),
            },
        )
        .unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}

fn workbook_records(file: &[u8]) -> Vec<(u16, Vec<u8>)> {
    let mut ole = OleFile::open(Cursor::new(file.to_vec())).unwrap();
    let stream = ole.open_stream(&["Workbook"]).unwrap();
    let mut records = Vec::new();
    let mut offset = 0;
    while offset + 4 <= stream.len() {
        let kind = u16::from_le_bytes([stream[offset], stream[offset + 1]]);
        let len = usize::from(u16::from_le_bytes([stream[offset + 2], stream[offset + 3]]));
        let end = offset + 4 + len;
        if end > stream.len() {
            break;
        }
        records.push((kind, stream[offset..end].to_vec()));
        offset = end;
        if kind == 0x000a {
            break;
        }
    }
    records
}

fn replace_group(file: &[u8], old: &[u8], new: &[u8]) -> Vec<u8> {
    assert_eq!(old.len(), new.len());
    let matches = file
        .windows(old.len())
        .enumerate()
        .filter_map(|(index, value)| (value == old).then_some(index))
        .collect::<Vec<_>>();
    assert_eq!(matches.len(), 1);
    let mut output = file.to_vec();
    output[matches[0]..matches[0] + new.len()].copy_from_slice(new);
    output
}

#[test]
fn rich_name_and_comment_round_trip_as_inert_metadata() {
    let mut writer = Writer::new();
    writer.add_worksheet("Sheet1").unwrap();
    writer
        .define_name_with_comment("Rate", "A1", "Unicode \u{7a0e}\u{7387}")
        .unwrap();
    writer
        .add_defined_name_record(DefinedNameRecordOptions {
            name: "MacroCommand".to_string(),
            kind: DefinedNameKind::User,
            scope: NameScope::Workbook,
            hidden: true,
            function: false,
            vba_procedure: true,
            procedure: true,
            calculated_expression: true,
            function_group: 14,
            published: true,
            workbook_parameter: true,
            shortcut_key: Some(b'K'),
            formula_tokens: vec![],
            formula_extra: vec![7; 9_000],
            custom_menu: "Menu".to_string(),
            description: "Description".to_string(),
            help_topic: "Help".to_string(),
            status_bar: "Status".to_string(),
            comment: Some("Macro metadata only".to_string()),
        })
        .unwrap();
    let mut categories = (0..19)
        .map(|index| format!("Category{index}"))
        .collect::<Vec<_>>();
    categories[18] = "Extended".to_string();
    writer
        .set_function_groups(FunctionGroupOptions {
            built_in: BuiltInFunctionCategories::Fourteen,
            custom_categories: categories,
        })
        .unwrap();
    writer
        .add_defined_name_record_with_future_records(
            DefinedNameRecordOptions {
                name: "ΣRate".to_string(),
                kind: DefinedNameKind::User,
                scope: NameScope::Workbook,
                hidden: false,
                function: true,
                vba_procedure: false,
                procedure: true,
                calculated_expression: false,
                function_group: 0,
                published: false,
                workbook_parameter: false,
                shortcut_key: None,
                formula_tokens: vec![],
                formula_extra: vec![],
                custom_menu: String::new(),
                description: String::new(),
                help_topic: String::new(),
                status_bar: String::new(),
                comment: Some("x".to_string()),
            },
            DefinedNameFutureRecords {
                function_group: Some(NameFnGrp12 {
                    function_name: "σRate".to_string(),
                    category: 32,
                }),
                publication: Some(NamePublish {
                    published: true,
                    workbook_parameter: true,
                    name: "σRate".to_string(),
                }),
            },
        )
        .unwrap();

    let mut bytes = Cursor::new(Vec::new());
    writer.write_to(&mut bytes).unwrap();
    let workbook = Workbook::new(Cursor::new(bytes.into_inner())).unwrap();
    assert_eq!(workbook.defined_names().len(), 1);
    assert_eq!(
        workbook.defined_names()[0].comment.as_deref(),
        Some("Unicode \u{7a0e}\u{7387}")
    );
    assert_eq!(workbook.defined_name_records().len(), 3);
    let macro_name = &workbook.defined_name_records()[1];
    assert!(macro_name.is_macro());
    assert!(macro_name.vba_procedure);
    assert_eq!(macro_name.shortcut_key, Some(b'K'));
    assert_eq!(macro_name.function_group, 14);
    assert_eq!(macro_name.formula_extra, vec![7; 9_000]);
    assert!(!macro_name.continuation_chunks.is_empty());
    assert_eq!(macro_name.custom_menu, "Menu");
    assert_eq!(macro_name.comment.as_deref(), Some("Macro metadata only"));
    let function = &workbook.defined_name_records()[2];
    assert!(function.is_macro());
    let future = &function.future_records;
    assert_eq!(future.function_group.as_ref().unwrap().category_index(), 0);
    assert_eq!(
        future.function_group.as_ref().unwrap().function_name,
        "σRate"
    );
    assert!(future.publication.as_ref().unwrap().published);
    assert!(future.publication.as_ref().unwrap().workbook_parameter);
}

#[test]
fn reads_poi_unicode_names_and_formula_extra() {
    let workbook = Workbook::new(File::open(fixture("testNames.xls")).unwrap()).unwrap();
    assert_eq!(workbook.defined_name_records().len(), 8);
    assert!(workbook.defined_name_records()[1].is_macro());
    let array_name = workbook
        .defined_name_records()
        .iter()
        .find(|name| name.name == "n_array")
        .unwrap();
    assert!(!array_name.formula_extra.is_empty());

    let workbook = Workbook::new(File::open(fixture("unicodeNameRecord.xls")).unwrap()).unwrap();
    assert!(
        workbook
            .defined_name_records()
            .iter()
            .any(|name| name.name == "日本語")
    );
}

#[test]
fn poi_name_comment_corpus_round_trips_through_rich_inert_options() {
    let source = Workbook::new(File::open(fixture("53109.xls")).unwrap()).unwrap();
    let name = source
        .defined_name_records()
        .iter()
        .find(|name| name.comment.is_some() && !name.is_macro())
        .unwrap();
    let mut writer = Writer::new();
    for index in 0..source.sheets().len() {
        writer
            .add_worksheet(&format!("Sheet{}", index + 1))
            .unwrap();
    }
    writer
        .add_defined_name_record(DefinedNameRecordOptions {
            name: name.name.clone(),
            kind: name.kind,
            scope: name.scope,
            hidden: name.hidden,
            function: name.function,
            vba_procedure: name.vba_procedure,
            procedure: name.procedure,
            calculated_expression: name.calculated_expression,
            function_group: name.function_group,
            published: name.published,
            workbook_parameter: name.workbook_parameter,
            shortcut_key: name.shortcut_key,
            formula_tokens: name.formula_tokens.clone(),
            formula_extra: name.formula_extra.clone(),
            custom_menu: name.custom_menu.clone(),
            description: name.description.clone(),
            help_topic: name.help_topic.clone(),
            status_bar: name.status_bar.clone(),
            comment: name.comment.clone(),
        })
        .unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    let reparsed = Workbook::new(Cursor::new(output.into_inner())).unwrap();
    let round_tripped = &reparsed.defined_name_records()[0];
    assert_eq!(round_tripped.name, name.name);
    assert_eq!(round_tripped.comment, name.comment);
    assert_eq!(round_tripped.formula_tokens, name.formula_tokens);
    assert_eq!(round_tripped.formula_extra, name.formula_extra);
}

#[test]
fn exact_future_records_reject_malformed_order_cardinality_names_headers_and_continue() {
    let file = future_record_file();
    let records = workbook_records(&file);
    let function_index = records
        .iter()
        .position(|(kind, _)| *kind == 0x0899)
        .unwrap();
    assert_eq!(records[function_index - 1].0, 0x0894);
    assert_eq!(records[function_index + 1].0, 0x0893);
    let comment = &records[function_index - 1].1;
    let function = &records[function_index].1;
    let publish = &records[function_index + 1].1;
    assert_eq!(comment.len(), function.len());
    assert_eq!(
        &function[4..16],
        &[0x99, 0x08, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]
    );
    assert_eq!(u16::from_le_bytes([function[16], function[17]]), 5);
    assert_eq!(u16::from_le_bytes([function[18], function[19]]), 32);
    assert_eq!(u16::from_le_bytes([publish[16], publish[17]]) & 3, 2);
    let mut group = Vec::new();
    group.extend_from_slice(comment);
    group.extend_from_slice(function);
    group.extend_from_slice(publish);
    let parses =
        |replacement: &[u8]| Workbook::new(Cursor::new(replace_group(&file, &group, replacement)));
    let mut out_of_order = Vec::new();
    out_of_order.extend_from_slice(function);
    out_of_order.extend_from_slice(comment);
    out_of_order.extend_from_slice(publish);
    assert!(parses(&out_of_order).is_err());
    let mut duplicate = group.clone();
    duplicate[comment.len()..comment.len() + function.len()].copy_from_slice(comment);
    assert!(parses(&duplicate).is_err());
    let mut bad_header = group.clone();
    bad_header[comment.len() + 6] = 1;
    assert!(parses(&bad_header).is_err());
    let mut bad_cached_count = group.clone();
    bad_cached_count[comment.len() + 16] = 4;
    assert!(parses(&bad_cached_count).is_err());
    let mut bad_function_name = group.clone();
    bad_function_name[comment.len() + 23] = 0xc4;
    assert!(parses(&bad_function_name).is_err());
    let publish_start = comment.len() + function.len();
    let mut bad_publish_name = group.clone();
    bad_publish_name[publish_start + 21] = 0xc4;
    assert!(parses(&bad_publish_name).is_err());
    let mut continued = group.clone();
    continued[publish_start..publish_start + 2].copy_from_slice(&0x003cu16.to_le_bytes());
    assert!(parses(&continued).is_err());
    let mut ignored_unused = group.clone();
    ignored_unused[publish_start + 16] |= 4;
    let workbook = parses(&ignored_unused).unwrap();
    assert_eq!(
        workbook.defined_name_records()[0]
            .future_records
            .publication
            .as_ref()
            .unwrap()
            .name,
        "σRate"
    );
}

#[test]
fn future_record_writer_enforces_caps_names_and_emitted_category_references() {
    let options = || DefinedNameRecordOptions {
        name: "Fn".to_string(),
        kind: DefinedNameKind::User,
        scope: NameScope::Workbook,
        hidden: false,
        function: true,
        vba_procedure: false,
        procedure: true,
        calculated_expression: false,
        function_group: 0,
        published: false,
        workbook_parameter: false,
        shortcut_key: None,
        formula_tokens: vec![],
        formula_extra: vec![],
        custom_menu: String::new(),
        description: String::new(),
        help_topic: String::new(),
        status_bar: String::new(),
        comment: None,
    };
    let mut writer = Writer::new();
    writer.add_worksheet("Sheet1").unwrap();
    assert!(
        writer
            .add_defined_name_record_with_future_records(
                options(),
                DefinedNameFutureRecords {
                    function_group: Some(NameFnGrp12 {
                        function_name: "Fn".to_string(),
                        category: 31
                    }),
                    publication: None
                }
            )
            .is_err()
    );
    assert!(
        writer
            .add_defined_name_record_with_future_records(
                options(),
                DefinedNameFutureRecords {
                    function_group: None,
                    publication: Some(NamePublish {
                        published: false,
                        workbook_parameter: false,
                        name: "X".repeat(256)
                    })
                }
            )
            .is_err()
    );
    writer
        .add_defined_name_record_with_future_records(
            options(),
            DefinedNameFutureRecords {
                function_group: Some(NameFnGrp12 {
                    function_name: "Fn".to_string(),
                    category: 32,
                }),
                publication: None,
            },
        )
        .unwrap();
    let mut output = Cursor::new(Vec::new());
    assert!(writer.write_to(&mut output).is_err());
}
