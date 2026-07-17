use std::fs::File;
use std::io::Cursor;
use std::path::{Path, PathBuf};

use litchi_ole::xls::writer::{XlsDefinedNameRecordOptions, XlsWriter};
use litchi_ole::xls::{XlsDefinedNameKind, XlsNameScope, XlsWorkbook};

fn fixture(name: &str) -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../3rdparty/poi/test-data/spreadsheet")
        .join(name)
}

#[test]
fn rich_name_and_comment_round_trip_as_inert_metadata() {
    let mut writer = XlsWriter::new();
    writer.add_worksheet("Sheet1").unwrap();
    writer.define_name_with_comment("Rate", "A1", "Unicode \u{7a0e}\u{7387}").unwrap();
    writer.add_defined_name_record(XlsDefinedNameRecordOptions {
        name: "MacroCommand".to_string(),
        kind: XlsDefinedNameKind::User,
        scope: XlsNameScope::Workbook,
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
    }).unwrap();

    let mut bytes = Cursor::new(Vec::new());
    writer.write_to(&mut bytes).unwrap();
    let workbook = XlsWorkbook::new(Cursor::new(bytes.into_inner())).unwrap();
    assert_eq!(workbook.defined_names().len(), 1);
    assert_eq!(workbook.defined_names()[0].comment.as_deref(), Some("Unicode \u{7a0e}\u{7387}"));
    assert_eq!(workbook.defined_name_records().len(), 2);
    let macro_name = &workbook.defined_name_records()[1];
    assert!(macro_name.is_macro());
    assert!(macro_name.vba_procedure);
    assert_eq!(macro_name.shortcut_key, Some(b'K'));
    assert_eq!(macro_name.function_group, 14);
    assert_eq!(macro_name.formula_extra, vec![7; 9_000]);
    assert!(!macro_name.continuation_chunks.is_empty());
    assert_eq!(macro_name.custom_menu, "Menu");
    assert_eq!(macro_name.comment.as_deref(), Some("Macro metadata only"));
}

#[test]
fn reads_poi_unicode_names_and_formula_extra() {
    let workbook = XlsWorkbook::new(File::open(fixture("testNames.xls")).unwrap()).unwrap();
    assert_eq!(workbook.defined_name_records().len(), 8);
    assert!(workbook.defined_name_records()[1].is_macro());
    let array_name = workbook.defined_name_records().iter().find(|name| name.name == "n_array").unwrap();
    assert!(!array_name.formula_extra.is_empty());

    let workbook = XlsWorkbook::new(File::open(fixture("unicodeNameRecord.xls")).unwrap()).unwrap();
    assert!(workbook.defined_name_records().iter().any(|name| name.name == "日本語"));
}
