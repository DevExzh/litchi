//! Native Excel names, validation, protection and CFB property transactions.
use litchi_core::{
    Metadata,
    sheet::{Cell as _, CellValue},
};
use litchi_ole_common::property_set::summary_properties::MetadataError;
use litchi_ole_common::property_set::{PropertySetReader, SummaryProperties, Value};
use litchi_xls::{
    BuiltInName, DataValidationErrorStyle as ErrorStyle, DataValidationImeMode as ImeMode,
    DataValidationKind as Kind, DataValidationOperator as Operator, DefinedNameKind, NameScope,
    Package, raw::Reader,
};
use std::{collections::BTreeMap, io::Cursor};

const RULES: &str = "excel-interaction-print.xls";
fn native(name: &str) -> Vec<u8> {
    std::fs::read(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/microsoft-office/excel")
            .join(name),
    )
    .unwrap()
}
fn streams(bytes: &[u8]) -> BTreeMap<Vec<String>, Vec<u8>> {
    let mut ole =
        litchi_cfb::OleFile::open(Cursor::new(bytes), litchi_cfb::OleFileLimits::default())
            .unwrap();
    ole.list_streams()
        .into_iter()
        .map(|path| {
            let names = path.iter().map(String::as_str).collect::<Vec<_>>();
            let data = ole.open_stream(&names).unwrap();
            (path, data)
        })
        .collect()
}
fn properties(p: &Package) -> SummaryProperties {
    p.edit_metadata().unwrap().properties().unwrap().clone()
}
fn custom(p: &Package) -> BTreeMap<String, Value> {
    let bytes = p.to_bytes().unwrap();
    let mut ole =
        litchi_cfb::OleFile::open(Cursor::new(bytes), litchi_cfb::OleFileLimits::default())
            .unwrap();
    ole.get_metadata()
        .unwrap()
        .custom_properties
        .into_iter()
        .collect()
}
fn clsid(bytes: &[u8]) -> String {
    litchi_cfb::OleFile::open(Cursor::new(bytes), litchi_cfb::OleFileLimits::default())
        .unwrap()
        .root_entry()
        .unwrap()
        .clsid
        .clone()
}
fn reopen_metadata(source: &Package, commit: &litchi_xls::package::metadata::Commit) -> Package {
    let file = tempfile::NamedTempFile::with_suffix(".xls").unwrap();
    commit.package().save(file.path()).unwrap();
    let after = Package::open(file.path()).unwrap();
    let source_bytes = source.to_bytes().unwrap();
    let after_bytes = after.to_bytes().unwrap();
    let a = streams(&source_bytes);
    let b = streams(&after_bytes);
    assert_eq!(a.keys().collect::<Vec<_>>(), b.keys().collect::<Vec<_>>());
    for (name, data) in &a {
        if !matches!(name.as_slice(),[s] if matches!(s.as_str(),"\u{5}SummaryInformation"|"\u{5}DocumentSummaryInformation"))
        {
            assert_eq!(*data, b[name], "changed {}", name.join("/"));
        }
    }
    assert_eq!(clsid(&source_bytes), clsid(&after_bytes));
    assert_eq!(
        properties(source).document_summary_information(),
        properties(&after).document_summary_information()
    );
    assert_eq!(
        source
            .apply_metadata(commit.patch())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap(),
        after_bytes
    );
    assert!(matches!(
        after.apply_metadata(commit.patch()),
        Err(litchi_xls::package::metadata::Error::PatchConflict)
    ));
    assert_eq!(
        after
            .apply_metadata(&commit.patch().inverse())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap(),
        source_bytes
    );
    after
}

#[test]
fn native_defined_names_keep_workbook_and_sheet_scope_with_exact_tokens() {
    let bytes = native(RULES);
    let r = Reader::new(Cursor::new(&bytes)).unwrap();
    let names = r.defined_names();
    assert_eq!(names.len(), 3);
    let expected = [
        (
            "FixtureAmounts",
            NameScope::Workbook,
            DefinedNameKind::User,
            "='Sheet1'!$E$2:$E$5",
            vec![59, 0, 0, 1, 0, 4, 0, 4, 0, 4, 0],
        ),
        (
            "Print_Area",
            NameScope::Worksheet(0),
            DefinedNameKind::BuiltIn(BuiltInName::PrintArea),
            "='Sheet1'!$A$1:$I$10",
            vec![59, 0, 0, 0, 0, 9, 0, 0, 0, 8, 0],
        ),
        (
            "Print_Titles",
            NameScope::Worksheet(0),
            DefinedNameKind::BuiltIn(BuiltInName::PrintTitles),
            "=('Sheet1'!$A$1:$A$65536,'Sheet1'!$A$1:$IV$1)",
            vec![
                41, 23, 0, 59, 0, 0, 0, 0, 255, 255, 0, 0, 0, 0, 59, 0, 0, 0, 0, 0, 0, 0, 0, 255,
                0, 16,
            ],
        ),
    ];
    for (index, (name, (text, scope, kind, formula, tokens))) in
        names.iter().zip(expected).enumerate()
    {
        assert_eq!(name.record_index, index as u32 + 1);
        assert_eq!(name.name, text);
        assert_eq!(name.scope, scope);
        assert_eq!(name.kind, kind);
        assert_eq!(name.formula.as_deref(), Some(formula));
        assert_eq!(name.formula_tokens, tokens);
        assert!(
            !name.hidden
                && !name.is_macro()
                && !name.is_deleted()
                && !name.calculated_expression
                && !name.published
                && !name.workbook_parameter
        );
        assert!(name.formula_extra.is_empty() && name.continuation_chunks.is_empty());
        assert_eq!(name.comment, None);
    }
    assert_eq!(
        Package::from_bytes(&bytes).unwrap().to_bytes().unwrap(),
        bytes
    );
}

#[test]
fn native_seven_validation_kinds_keep_ranges_messages_and_inert_formulas() {
    let bytes = native(RULES);
    let r = Reader::new(Cursor::new(&bytes)).unwrap();
    let sheet = r.xls_worksheet(0).unwrap();
    let settings = sheet.data_validation_settings().unwrap();
    assert!(!settings.window_closed());
    assert_eq!(
        (
            settings.x_left(),
            settings.y_top(),
            settings.dropdown_object_id(),
            settings.declared_rule_count()
        ),
        (0, 0, None, 7)
    );
    let rules = sheet.data_validations();
    assert_eq!(rules.len(), 7);
    for (index, (rule, (kind, column, last_row, first, second))) in rules
        .iter()
        .zip([
            (Kind::Whole, 1, 4, "=1", Some("=10")),
            (Kind::List, 2, 4, "=\"Open\0Closed\0Hold\"", None),
            (Kind::Date, 3, 4, "=46023", Some("=46387")),
            (Kind::Decimal, 4, 4, "=0", Some("=100")),
            (Kind::Time, 6, 3, "=0.375", Some("=0.75")),
            (Kind::TextLength, 7, 3, "=3", Some("=5")),
            (Kind::Custom, 8, 3, "=(MOD(I2,2)=0)", None),
        ])
        .enumerate()
    {
        assert_eq!(rule.kind(), kind);
        assert_eq!(rule.operator(), Operator::Between);
        assert_eq!(rule.error_style(), ErrorStyle::Stop);
        assert_eq!(rule.ime_mode(), ImeMode::NoControl);
        assert!(rule.allow_blank() && rule.show_input_message() && rule.show_error_message());
        assert!(!rule.suppress_dropdown());
        assert_eq!(rule.explicit_list(), index == 1);
        assert_eq!(rule.ranges().len(), 1);
        let range = rule.ranges()[0];
        assert_eq!(
            (
                range.first_row(),
                range.last_row(),
                range.first_column(),
                range.last_column()
            ),
            (1, last_row, column, column)
        );
        assert_eq!(rule.formula1().unwrap().rendered(), Some(first));
        assert_eq!(rule.formula2().and_then(|f| f.rendered()), second);
        assert_eq!(rule.prompt(), None);
        assert_eq!(
            rule.prompt_title(),
            if index == 0 {
                Some("Quantity input")
            } else {
                None
            }
        );
        assert_eq!(
            rule.error_title(),
            if index == 0 {
                Some("Invalid quantity")
            } else {
                None
            }
        );
        assert_eq!(
            rule.error(),
            if index == 0 {
                Some("Quantity must be an integer from 1 to 10.")
            } else {
                None
            }
        );
    }
    assert_eq!(
        rules[1].formula1().unwrap().tokens(),
        [
            23, 16, 0, 79, 112, 101, 110, 0, 67, 108, 111, 115, 101, 100, 0, 72, 111, 108, 100
        ]
    );
    assert_eq!(
        rules[6].formula1().unwrap().tokens(),
        [76, 0, 0, 0, 192, 30, 2, 0, 65, 39, 0, 30, 0, 0, 11]
    );
    for (row, value) in [(1, 2.), (2, 4.), (3, 6.)] {
        assert_eq!(
            sheet.get_cell(row, 8).unwrap().value(),
            &CellValue::Float(value)
        );
    }
    // Formula display is anchored at the first target; no rule was evaluated.
    assert_eq!(
        Package::from_bytes(&bytes).unwrap().to_bytes().unwrap(),
        bytes
    );
}

#[test]
fn native_passwordless_protection_keeps_flags_and_refuses_tab_rename() {
    use litchi_xls::package::edit::Error;
    let bytes = native(RULES);
    let r = Reader::new(Cursor::new(&bytes)).unwrap();
    let book = r.protection();
    assert!(book.structure_protected());
    assert!(!book.windows_protected() && !book.revisions_protected() && !book.write_protected());
    assert_eq!(book.password().raw(), 0);
    assert_eq!(book.revision_password().raw(), 0);
    assert!(book.file_sharing().is_none());
    let sheet = r.xls_worksheet(0).unwrap().protection();
    assert!(sheet.is_protected() && sheet.objects_protected() && sheet.scenarios_protected());
    assert_eq!(sheet.password(), None);
    let p = Package::from_bytes(&bytes).unwrap();
    let mut edit = p.edit().unwrap();
    edit.rename_sheet("Sheet1", "Sheet1").unwrap();
    let error = edit.rename_sheet("Sheet1", "Renamed").unwrap_err();
    assert!(
        matches!(error,Error::CellValues(litchi_xls::Error::Refused(r)) if r.kind()==litchi_core::edit::RefusalKind::Protected)
    );
    assert_eq!(edit.commit().unwrap().package().to_bytes().unwrap(), bytes);
    assert_eq!(p.to_bytes().unwrap(), bytes);
}

#[test]
fn native_mac_property_streams_keep_producer_dates_and_reserved_hyperlinks() {
    for (name, created, modified) in [
        (RULES, "2026-10-15T06:03:06Z", "2026-10-15T06:13:31Z"),
        (
            "excel-metadata.xls",
            "2026-10-15T11:10:49Z",
            "2026-10-15T11:17:10Z",
        ),
        (
            "excel-recorded-macro.xls",
            "2026-10-15T06:29:36Z",
            "2026-10-15T06:31:28Z",
        ),
    ] {
        let bytes = native(name);
        let p = Package::from_bytes(&bytes).unwrap();
        let m = p.metadata().unwrap();
        assert_eq!(m.author.as_deref(), Some("Ryker Zhu"));
        assert_eq!(m.last_modified_by, m.author);
        assert_eq!(m.application.as_deref(), Some("Microsoft Macintosh Excel"));
        assert_eq!(m.created.unwrap().to_string(), created);
        assert_eq!(m.modified.unwrap().to_string(), modified);
        assert_eq!(m.company.as_deref(), Some(""));
        assert_eq!(m.title, None);
        assert_eq!(m.security, Some(0));
        assert_eq!(m.codepage, Some(10000));
        let props = properties(&p);
        assert_eq!(props.summary_information().codepage().unwrap().id(), 10000);
        let values = custom(&p);
        if name == RULES {
            assert!(props.user_defined_properties().is_some());
            assert_eq!(values.len(), 1);
            let Value::Blob(blob) = &values["_PID_HLINKS"] else {
                panic!("reserved hyperlink blob");
            };
            assert_eq!(blob.len(), 140);
            let target = "https://example.com/litchi-excel-fixture"
                .encode_utf16()
                .flat_map(u16::to_le_bytes)
                .collect::<Vec<_>>();
            assert!(blob.windows(target.len()).any(|bytes| bytes == target));
        } else {
            assert!(props.user_defined_properties().is_none());
            assert!(values.is_empty());
        }
        assert_eq!(
            p.edit_metadata()
                .unwrap()
                .commit()
                .unwrap()
                .package()
                .to_bytes()
                .unwrap(),
            bytes
        );
    }
}

#[test]
fn native_property_crud_preserves_workbook_rules_names_protection_and_macros() {
    for name in [RULES, "excel-metadata.xls", "excel-recorded-macro.xls"] {
        let bytes = native(name);
        let p = Package::from_bytes(&bytes).unwrap();
        let before = p.metadata().unwrap();
        let original_custom = custom(&p);
        let mut edit = p.edit_metadata().unwrap();
        edit.merge_metadata(&Metadata {
            title: Some("Native validation review".into()),
            ..Metadata::default()
        })
        .unwrap();
        edit.properties_mut()
            .unwrap()
            .set_user_defined("NativeReviewed", Value::Bool(true))
            .unwrap();
        let staged = edit.properties().unwrap().clone();
        assert!(matches!(
            edit.merge_metadata(&Metadata {
                title: Some("must not be staged".into()),
                identifier: Some("urn:unsupported".into()),
                ..Metadata::default()
            }),
            Err(litchi_xls::package::metadata::Error::Metadata(
                MetadataError::Unsupported {
                    field: "identifier"
                }
            ))
        ));
        assert_eq!(edit.properties().unwrap(), &staged);
        let commit = edit.commit().unwrap();
        let after = reopen_metadata(&p, &commit);
        let mut expected = before.clone();
        expected.title = Some("Native validation review".into());
        assert_eq!(after.metadata().unwrap(), expected);
        let mut expected_custom = original_custom.clone();
        expected_custom.insert("NativeReviewed".into(), Value::Bool(true));
        assert_eq!(custom(&after), expected_custom);
        let mut update = after.edit_metadata().unwrap();
        update
            .properties_mut()
            .unwrap()
            .set_user_defined("NativeReviewed", Value::Bool(false))
            .unwrap();
        let changed = update.commit().unwrap();
        let updated = reopen_metadata(&after, &changed);
        assert_eq!(custom(&updated)["NativeReviewed"], Value::Bool(false));
        let mut remove = updated.edit_metadata().unwrap();
        assert!(remove.properties_mut().unwrap().remove("title").unwrap());
        assert!(
            remove
                .properties_mut()
                .unwrap()
                .remove("NativeReviewed")
                .unwrap()
        );
        let removed = remove.commit().unwrap();
        let restored = reopen_metadata(&updated, &removed);
        assert_eq!(restored.metadata().unwrap(), before);
        assert_eq!(custom(&restored), original_custom);
        assert_eq!(p.to_bytes().unwrap(), bytes);
    }
}

#[test]
fn native_mac_custom_unicode_edit_changes_only_the_custom_section_codepage() {
    let bytes = native(RULES);
    let p = Package::from_bytes(&bytes).unwrap();
    let before = properties(&p);
    let original_custom = custom(&p);
    let mut edit = p.edit_metadata().unwrap();
    edit.properties_mut()
        .unwrap()
        .set_user_defined("审核", Value::Lpwstr("完成".into()))
        .unwrap();
    let commit = edit.commit().unwrap();
    let after = reopen_metadata(&p, &commit);
    let props = properties(&after);
    assert_eq!(
        props.user_defined_properties().unwrap().page(),
        Some(litchi_ole_common::property_set::CodePage::Utf16Le)
    );
    assert_eq!(props.summary_information(), before.summary_information());
    let summary_key = vec!["\u{5}SummaryInformation".to_owned()];
    assert_eq!(
        streams(&bytes)[&summary_key],
        streams(&after.to_bytes().unwrap())[&summary_key]
    );
    assert_eq!(
        props.document_summary_information(),
        before.document_summary_information()
    );
    let mut expected = original_custom;
    expected.insert("审核".into(), Value::Lpwstr("完成".into()));
    assert_eq!(custom(&after), expected);
    assert_eq!(p.to_bytes().unwrap(), bytes);
}
