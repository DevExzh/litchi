//! Native Microsoft Excel files with independently recorded expectations.

use litchi_xlsx::Package;
use litchi_xlsx::{Cell, Worksheet, WorksheetKind};

fn native(name: &str) -> Package {
    Package::open(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/microsoft-office/excel")
            .join(name),
    )
    .unwrap()
}

#[test]
fn native_excel_hyperlink_retains_target_tooltip_and_revision_identity() {
    let source = native("excel-interaction-print.xlsx");
    let workbook = source.workbook();
    let sheet = workbook.sheet("Sheet1").unwrap().unwrap();
    let links = sheet.hyperlinks().unwrap();
    assert_eq!(links.len(), 1);
    assert_eq!(links[0].reference().as_str(), "A10");
    assert_eq!(
        links[0].external_target(),
        Some("https://example.com/litchi-excel-fixture")
    );
    assert_eq!(links[0].tooltip(), Some("Native hyperlink screen tip 002"));
    assert_eq!(
        links[0].revision_uid(),
        Some("{3A497593-3366-B24C-A7AF-6952EF70F853}")
    );
}

fn members(bytes: &[u8]) -> std::collections::BTreeMap<String, Vec<u8>> {
    let archive = litchi_opc::phys_pkg::OwnedPhysPkgReader::from_bytes(
        bytes.to_vec(),
        litchi_opc::SourceLimits::default(),
    )
    .unwrap();
    archive
        .member_names()
        .unwrap()
        .into_iter()
        .map(|name| {
            let data = archive.read_member(&name).unwrap();
            (name, data)
        })
        .collect()
}

fn assert_members_equal(before: &[u8], after: &[u8]) {
    let before = members(before);
    let after = members(after);
    assert_eq!(
        before.keys().collect::<Vec<_>>(),
        after.keys().collect::<Vec<_>>()
    );
    for (name, data) in before {
        assert!(after[&name] == data, "changed {name}");
    }
}

fn manifest_children(
    bytes: &[u8],
    omit_chain: bool,
) -> Vec<std::collections::BTreeMap<String, String>> {
    use quick_xml::events::Event;
    let mut reader = quick_xml::Reader::from_reader(bytes);
    let mut children = Vec::new();
    let mut depth = 0;
    loop {
        let event = reader.read_event().unwrap();
        match &event {
            Event::Start(element) | Event::Empty(element) if depth <= 1 => {
                let mut attributes = std::collections::BTreeMap::new();
                attributes.insert(
                    "element".into(),
                    String::from_utf8(element.local_name().as_ref().to_vec()).unwrap(),
                );
                for attribute in element.attributes() {
                    let attribute = attribute.unwrap();
                    attributes.insert(
                        String::from_utf8(attribute.key.as_ref().to_vec()).unwrap(),
                        attribute
                            .decoded_and_normalized_value(
                                quick_xml::XmlVersion::Explicit1_0,
                                reader.decoder(),
                            )
                            .unwrap()
                            .into_owned(),
                    );
                }
                if !omit_chain
                    || !attributes.values().any(|value| {
                        value.ends_with("/calcChain") || value.ends_with("calcChain.xml")
                    })
                {
                    children.push(attributes);
                }
            },
            Event::Eof => break,
            _ => {},
        }
        match event {
            Event::Start(_) => depth += 1,
            Event::End(_) => depth -= 1,
            _ => {},
        }
    }
    children.sort();
    children
}

fn workbook_events_without_calculation_properties(bytes: &[u8]) -> Vec<Vec<u8>> {
    use quick_xml::events::Event;
    let mut reader = quick_xml::Reader::from_reader(bytes);
    let mut events = Vec::new();
    loop {
        let event = reader.read_event().unwrap();
        match &event {
            Event::Eof => break,
            Event::Text(text) if text.as_ref().iter().all(u8::is_ascii_whitespace) => {},
            Event::Empty(element) if element.local_name().as_ref() == b"calcPr" => {
                let mut kept = std::collections::BTreeMap::new();
                for attribute in element.attributes() {
                    let attribute = attribute.unwrap();
                    if !matches!(
                        attribute.key.as_ref(),
                        b"calcId"
                            | b"fullCalcOnLoad"
                            | b"forceFullCalc"
                            | b"calcCompleted"
                            | b"calcOnSave"
                    ) {
                        kept.insert(attribute.key.as_ref().to_vec(), attribute.value.to_vec());
                    }
                }
                if !kept.is_empty() {
                    events.push(format!("calcPr:{kept:?}").into_bytes());
                }
            },
            _ => events.push(event.as_ref().to_vec()),
        }
    }
    events
}

#[test]
fn native_excel_1904_dates_breaks_and_hidden_rows_keep_stored_values() {
    let source = native("excel-date1904-breaks.xlsx");
    let workbook = source.workbook();
    assert_eq!(workbook.date_system(), litchi_xlsx::DateSystem::Excel1904);
    let sheet = workbook.sheet("Sheet1").unwrap().unwrap();
    scalar_number(&sheet, "B3", "0");
    scalar_number(&sheet, "B4", "1");
    scalar_number(&sheet, "B7", "-1");
    formula(&sheet, "B5", "DATE(2024,2,29)", "43889");
    formula(
        &sheet,
        "B6",
        "DATE(2026,10,15)+TIME(13,14,15)",
        "44848.551562499997",
    );
    formula(&sheet, "B8", "TIME(6,30,0)", "0.27083333333333331");
    assert!(sheet.row(13u32).unwrap().hidden());
    assert!(sheet.row(14u32).unwrap().hidden());
    assert!(!sheet.row(12u32).unwrap().hidden());
    assert!(!sheet.row(15u32).unwrap().hidden());
    assert!(sheet.column(9u32).unwrap().hidden());
    let breaks = sheet.page_breaks().unwrap();
    for (collection, id, maximum) in [
        (breaks.horizontal().unwrap(), 19, 16383),
        (breaks.vertical().unwrap(), 6, 1048575),
    ] {
        assert_eq!(collection.len(), 1);
        let value = collection.breaks()[0];
        assert_eq!(
            (value.id(), value.minimum(), value.maximum()),
            (id, 0, maximum)
        );
        assert!(value.is_manual());
        assert!(!value.is_pivot());
    }
}

#[test]
fn native_excel_hyperlink_edit_keeps_revision_uid_and_other_parts() {
    let source = native("excel-interaction-print.xlsx");
    let before = source.to_bytes().unwrap();
    let sheet = source.workbook().sheet("Sheet1").unwrap().unwrap();
    let link = sheet.hyperlinks().unwrap().remove(0);
    let changed = litchi_xlsx::hyperlinks::Hyperlink::new(
        link.reference().clone(),
        link.location().map(str::to_owned),
        link.external_target().map(str::to_owned),
        link.display().map(str::to_owned),
        Some("Native hyperlink edited".into()),
    )
    .unwrap()
    .with_revision_uid(link.revision_uid().map(str::to_owned))
    .unwrap();
    let mut edit = source.edit().unwrap();
    edit.sheet("Sheet1")
        .unwrap()
        .unwrap()
        .put_hyperlink(changed.clone())
        .unwrap();
    let commit = edit.commit().unwrap();
    let after = commit.package().to_bytes().unwrap();
    let reopened = Package::from_bytes(&after).unwrap();
    assert_eq!(
        reopened
            .workbook()
            .sheet("Sheet1")
            .unwrap()
            .unwrap()
            .hyperlinks()
            .unwrap(),
        [changed]
    );
    let old = members(&before);
    let new = members(&after);
    assert_eq!(
        old.keys().collect::<Vec<_>>(),
        new.keys().collect::<Vec<_>>()
    );
    for (name, data) in old {
        if name != "xl/worksheets/sheet1.xml" {
            assert!(new[&name] == data, "changed {name}");
        }
    }
    assert_members_equal(
        &reopened
            .apply(&commit.patch().inverse())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap(),
        &before,
    );
    assert!(source.to_bytes().unwrap() == before);
}

#[test]
fn native_excel_variants_keep_catalog_kinds_and_every_package_member() {
    let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/microsoft-office/excel");
    for (name, names) in [
        (
            "excel-values-formulas-chart.xlsx",
            vec!["Sales chart", "Sheet1", "Sheet2"],
        ),
        ("excel-interaction-print.xlsx", vec!["Sheet1"]),
        ("excel-recorded-macro.xlsm", vec!["Sheet1"]),
        ("excel-recorded-macro.xltm", vec!["Sheet1"]),
        (
            "excel-visual-analytics.xlsx",
            vec!["Sheet2", "Sheet1", "Sheet3"],
        ),
        (
            "excel-visual-analytics.xltx",
            vec!["Sheet2", "Sheet1", "Sheet3"],
        ),
        (
            "excel-visual-analytics-strict.xlsx",
            vec!["Sheet2", "Sheet1", "Sheet3"],
        ),
        ("excel-conditional-drawings.xlsx", vec!["Sheet1"]),
        ("excel-links-objects.xlsx", vec!["Sheet1"]),
        (
            "excel-pivot-controls.xlsx",
            vec!["PivotChart007", "Sheet2", "Sheet1", "Sheet3"],
        ),
        ("excel-form-controls.xlsx", vec!["Sheet1"]),
        ("excel-formula-owners.xlsx", vec!["Sheet1", "Sheet2"]),
        ("excel-connections.xlsx", vec!["native-query-010", "Sheet1"]),
        ("excel-date1904-breaks.xlsx", vec!["Sheet1"]),
        (
            "excel-chart-families.xlsx",
            vec![
                "Line_013",
                "Area_013",
                "Pie_013",
                "Radar_013",
                "Surface_013",
                "Scatter_013",
                "Bubble_013",
                "DrawingSVG_013",
                "Bar_013",
                "Bar3D_013",
                "Area3D_013",
                "Line3D_013",
                "Contour_013",
                "Doughnut_013",
                "OfPie_013",
                "Pie3D_013",
                "Stock_013",
                "Combo_013",
            ],
        ),
        ("excel-revisions.xlsx", vec!["Sheet1", "Review_014"]),
        (
            "excel-metadata.xlsx",
            vec![
                "Sheet1",
                "Source_A_015",
                "Source_B_015",
                "Consolidated_015",
                "Print_015",
            ],
        ),
        ("excel-legacy-sheets.xlsm", vec!["Macro_016", "Sheet1"]),
    ] {
        let source = std::fs::read(root.join(name)).unwrap();
        let package = Package::from_bytes(&source).unwrap();
        let workbook = package.workbook();
        assert_eq!(
            workbook
                .sheets()
                .map(|sheet| sheet.name().to_owned())
                .collect::<Vec<_>>(),
            names,
            "{name}"
        );
        let expected_flavor = if name.ends_with(".xlsm") {
            litchi_xlsx::Flavor::MacroWorkbook
        } else if name.ends_with(".xltm") {
            litchi_xlsx::Flavor::MacroTemplate
        } else if name.ends_with(".xltx") {
            litchi_xlsx::Flavor::Template
        } else {
            litchi_xlsx::Flavor::Workbook
        };
        assert_eq!(workbook.flavor(), expected_flavor, "{name}");
        for sheet in workbook.sheets() {
            let kind = match sheet.name() {
                "Sales chart" | "PivotChart007" => WorksheetKind::Chart,
                "Macro_016" => WorksheetKind::Macro,
                _ => WorksheetKind::Worksheet,
            };
            assert_eq!(sheet.kind(), kind, "{name}: {}", sheet.name());
            if kind == WorksheetKind::Worksheet {
                assert!(
                    sheet.cells("A1:Z104").unwrap().next().is_some(),
                    "{name}: {}",
                    sheet.name()
                );
            } else {
                assert!(matches!(
                    sheet.cell("A1"),
                    Err(litchi_xlsx::Error::NotWorksheet { .. })
                ));
            }
        }
        assert_members_equal(&source, &package.to_bytes().unwrap());
    }
}

fn scalar_number(sheet: &Worksheet, at: &str, expected: &str) {
    let view = sheet.cell(at).unwrap();
    let Some(Cell::Value(litchi_xlsx::Value::Number(number))) = view.stored() else {
        panic!("{at}: expected numeric value, got {view:?}")
    };
    assert_eq!(number.as_str(), expected, "{at}");
}

fn formula(sheet: &Worksheet, at: &str, text: &str, cache: &str) -> litchi_xlsx::Formula {
    let view = sheet.cell(at).unwrap();
    let Some(Cell::Formula(formula)) = view.stored() else {
        panic!("{at}: expected formula, got {view:?}")
    };
    assert_eq!(formula.text(), text, "{at}");
    let cached = formula.cached().unwrap();
    let litchi_xlsx::Value::Number(number) = cached.value() else {
        panic!("{at}: expected number cache")
    };
    assert_eq!(number.as_str(), cache, "{at}");
    assert_eq!(cached.origin(), litchi_xlsx::formula::Origin::Stored);
    assert_eq!(cached.freshness(), litchi_xlsx::formula::Freshness::Unknown);
    formula.clone()
}

#[test]
fn native_excel_array_shared_dynamic_and_what_if_formula_owners_keep_stored_caches() {
    use litchi_xlsx::formula::Kind;
    let package = native("excel-formula-owners.xlsx");
    let workbook = package.workbook();
    let main = workbook.sheet("Sheet1").unwrap().unwrap();
    let dynamic = workbook.sheet("Sheet2").unwrap().unwrap();
    let array = formula(&main, "C4", "A4:A7*B4:B7", "20");
    assert!(matches!(array.kind(),Kind::Array {range:Some(range)} if range.as_str()=="C4:C7"));
    for (at, value) in [("C5", "60"), ("C6", "120"), ("C7", "200")] {
        scalar_number(&main, at, value);
    }
    for row in 4..=103 {
        let value = formula(&main, &format!("D{row}"), "ROW()*2", &(row * 2).to_string());
        assert_eq!(value.kind(), &Kind::Scalar);
    }
    let sequence = formula(&dynamic, "A4", "_xlfn.SEQUENCE(3,2,10,2)", "10");
    assert!(matches!(sequence.kind(),Kind::Array {range:Some(range)} if range.as_str()=="A4:B6"));
    for (at, value) in [
        ("B4", "12"),
        ("A5", "14"),
        ("B5", "16"),
        ("A6", "18"),
        ("B6", "20"),
        ("E4", "30"),
        ("D5", "5"),
        ("E5", "40"),
    ] {
        scalar_number(&dynamic, at, value);
    }
    let filter = formula(
        &dynamic,
        "D4",
        "_xlfn._xlws.FILTER(Sheet1!A4:B7,Sheet1!A4:A7>=4)",
        "4",
    );
    assert!(matches!(filter.kind(),Kind::Array {range:Some(range)} if range.as_str()=="D4:E5"));
    formula(
        &dynamic,
        "F4",
        "_xlfn.LET(_xlpm.x,7,_xlpm.x*_xlpm.x+1)",
        "50",
    );
    formula(&main, "H6", "H4*H5+5", "25");
    for (at, range, cache) in [("J4", "J4:J7", "7"), ("M4", "M4:P7", "15")] {
        let owner = main.cell(at).unwrap();
        let Some(Cell::Formula(value)) = owner.stored() else {
            panic!("missing what-if owner")
        };
        assert!(
            matches!(value.kind(),Kind::DataTable {range:Some(value)} if value.as_str()==range)
        );
        assert_eq!(
            value.cached().unwrap().value(),
            &litchi_xlsx::Value::Number(litchi_xlsx::cell::Number::new(cache).unwrap())
        );
    }
}

#[cfg(feature = "encryption")]
#[test]
fn native_excel_agile_password_ingress_and_rekey_keep_clear_members() {
    use litchi_crypto::ooxml::{Error as CryptoError, Mode};
    let bytes = include_bytes!("../../../test-data/microsoft-office/excel/excel-encrypted.xlsx");
    for password in ["", "wrong"] {
        assert!(matches!(
            Package::from_bytes_with_password(bytes, password),
            Err(litchi_xlsx::Error::Encryption(CryptoError::Password))
        ));
    }
    let source = Package::from_bytes_with_password(bytes, "OfficeFixture42!").unwrap();
    assert_eq!(source.encryption(), Some(Mode::AgileAes256Sha512));
    let sheet = source.workbook().sheet("Sheet1").unwrap().unwrap();
    assert_eq!(
        sheet.cell("A1").unwrap().stored(),
        Some(&Cell::Value(litchi_xlsx::Value::text(
            "EXCEL_ENCRYPTED_011"
        )))
    );
    scalar_number(&sheet, "B2", "42");
    formula(&sheet, "B3", "B2*2", "84");
    assert!(matches!(
        source.to_bytes(),
        Err(litchi_xlsx::Error::EncryptionPolicy { .. })
    ));
    let before = source.to_plain_bytes().unwrap();
    let encrypted = source.to_reencrypted("ChangedOfficeFixture42!").unwrap();
    assert!(matches!(
        Package::from_bytes_with_password(&encrypted, "OfficeFixture42!"),
        Err(litchi_xlsx::Error::Encryption(CryptoError::Password))
    ));
    let reopened =
        Package::from_bytes_with_password(&encrypted, "ChangedOfficeFixture42!").unwrap();
    assert_eq!(reopened.encryption(), Some(Mode::AgileAes256Sha512));
    assert_members_equal(&before, &reopened.to_plain_bytes().unwrap());
}

#[test]
fn native_excel_cell_create_update_remove_preserves_other_parts_and_inverts() {
    for name in [
        "excel-values-formulas-chart.xlsx",
        "excel-interaction-print.xlsx",
        "excel-recorded-macro.xlsm",
        "excel-recorded-macro.xltm",
        "excel-visual-analytics.xlsx",
        "excel-visual-analytics.xltx",
        "excel-visual-analytics-strict.xlsx",
        "excel-conditional-drawings.xlsx",
        "excel-links-objects.xlsx",
        "excel-pivot-controls.xlsx",
        "excel-form-controls.xlsx",
        "excel-formula-owners.xlsx",
        "excel-connections.xlsx",
        "excel-date1904-breaks.xlsx",
        "excel-chart-families.xlsx",
        "excel-revisions.xlsx",
        "excel-metadata.xlsx",
        "excel-legacy-sheets.xlsm",
    ] {
        let source = native(name);
        let before = source.to_bytes().unwrap();
        let sheet_name = if name == "excel-chart-families.xlsx" {
            "Line_013"
        } else {
            "Sheet1"
        };
        let sheet_part = if name.starts_with("excel-visual-analytics")
            || matches!(name, "excel-pivot-controls.xlsx" | "excel-connections.xlsx")
        {
            "xl/worksheets/sheet2.xml"
        } else {
            "xl/worksheets/sheet1.xml"
        };
        assert!(
            source
                .workbook()
                .sheet(sheet_name)
                .unwrap()
                .unwrap()
                .cell("Z100")
                .unwrap()
                .is_missing()
        );
        let mut edit = source.edit().unwrap();
        edit.sheet(sheet_name)
            .unwrap()
            .unwrap()
            .set("Z100", 42)
            .unwrap();
        if name == "excel-interaction-print.xlsx" {
            assert!(matches!(
                edit.commit(),
                Err(litchi_xlsx::Error::EditBlocked {
                    reason: litchi_xlsx::EditBlock::ProtectedSheet,
                    ..
                })
            ));
            assert!(source.to_bytes().unwrap() == before);
            continue;
        }
        let created = edit.commit().unwrap_or_else(|e| panic!("{name}: {e:?}"));
        let after = created.package().to_bytes().unwrap();
        let reopened = Package::from_bytes(&after).unwrap();
        scalar_number(
            &reopened.workbook().sheet(sheet_name).unwrap().unwrap(),
            "Z100",
            "42",
        );
        let old = members(&before);
        let new = members(&after);
        // Value changes invalidate the calculation-chain graph by contract.
        // Its part and two manifest entries are the only related removals.
        let had_chain = old.contains_key("xl/calcChain.xml");
        assert_eq!(
            old.keys()
                .filter(|part| part.as_str() != "xl/calcChain.xml")
                .collect::<Vec<_>>(),
            new.keys().collect::<Vec<_>>(),
            "{name}"
        );
        for (part, bytes) in &old {
            if part == sheet_part || part == "xl/calcChain.xml" {
                continue;
            }
            if part == "xl/workbook.xml" {
                assert_eq!(
                    workbook_events_without_calculation_properties(bytes),
                    workbook_events_without_calculation_properties(&new[part]),
                    "{name}: changed unrelated workbook content"
                );
                let metadata = reopened.workbook().calculation_metadata().unwrap();
                let properties = metadata.properties().unwrap();
                assert_eq!(properties.calculation_id(), 0);
                assert!(properties.full_calculation_on_load());
                assert!(properties.force_full_calculation());
                assert!(!properties.calculation_completed());
                assert!(properties.calculate_on_save());
            } else if had_chain
                && matches!(
                    part.as_str(),
                    "[Content_Types].xml" | "xl/_rels/workbook.xml.rels"
                )
            {
                assert_eq!(
                    manifest_children(bytes, true),
                    manifest_children(&new[part], true),
                    "{name}: changed manifest semantics {part}"
                );
                assert!(
                    !std::str::from_utf8(&new[part])
                        .unwrap()
                        .contains("calcChain")
                );
            } else {
                assert!(new[part] == *bytes, "{name}: changed {part}");
            }
        }
        let mut edit = reopened.edit().unwrap();
        edit.sheet(sheet_name)
            .unwrap()
            .unwrap()
            .set("Z100", 43)
            .unwrap();
        let updated = edit.commit().unwrap();
        scalar_number(
            &updated
                .package()
                .workbook()
                .sheet(sheet_name)
                .unwrap()
                .unwrap(),
            "Z100",
            "43",
        );
        let mut edit = updated.package().edit().unwrap();
        edit.sheet(sheet_name)
            .unwrap()
            .unwrap()
            .remove("Z100")
            .unwrap();
        let removed = edit.commit().unwrap();
        assert!(
            removed
                .package()
                .workbook()
                .sheet(sheet_name)
                .unwrap()
                .unwrap()
                .cell("Z100")
                .unwrap()
                .is_missing()
        );
        assert_members_equal(
            &created
                .package()
                .apply(&created.patch().inverse())
                .unwrap()
                .package()
                .to_bytes()
                .unwrap(),
            &before,
        );
        let restored = members(
            &reopened
                .apply(&created.patch().inverse())
                .unwrap()
                .package()
                .to_bytes()
                .unwrap(),
        );
        assert_eq!(
            old.keys().collect::<Vec<_>>(),
            restored.keys().collect::<Vec<_>>()
        );
        for (part, bytes) in &old {
            if part == "xl/_rels/workbook.xml.rels" {
                assert_eq!(
                    manifest_children(bytes, false),
                    manifest_children(&restored[part], false),
                    "{name}: inverse changed manifest fields"
                );
            } else {
                assert!(restored[part] == *bytes, "{name}: inverse changed {part}");
            }
        }
        assert!(
            source.to_bytes().unwrap() == before,
            "{name}: source mutated"
        );
    }
}

#[test]
fn native_excel_validation_notes_names_and_print_settings_are_typed() {
    use litchi_xlsx::data_validation::{ListSource, ValidationType};
    let source = native("excel-interaction-print.xlsx");
    let workbook = source.workbook();
    let sheet = workbook.sheet("Sheet1").unwrap().unwrap();
    let collections = sheet.data_validations().unwrap();
    let rules = collections
        .iter()
        .flat_map(|collection| collection.validations())
        .collect::<Vec<_>>();
    assert_eq!(rules.len(), 7);
    for (range, kind, first, second) in [
        ("B2:B5", ValidationType::Whole, "1", Some("10")),
        ("C2:C5", ValidationType::List, "\"Open,Closed,Hold\"", None),
        ("D2:D5", ValidationType::Date, "46023", Some("46387")),
        ("E2:E5", ValidationType::Decimal, "0", Some("100")),
        ("G2:G4", ValidationType::Time, "0.375", Some("0.75")),
        ("H2:H4", ValidationType::TextLength, "3", Some("5")),
        ("I2:I4", ValidationType::Custom, "MOD(I2,2)=0", None),
    ] {
        let rule = rules
            .iter()
            .find(|rule| {
                rule.sqref().ranges().len() == 1 && rule.sqref().ranges()[0].as_str() == range
            })
            .unwrap();
        assert_eq!(rule.validation_type(), kind);
        let Some(ListSource::Formula(value)) = rule.formula1() else {
            panic!("{range}: expected formula source")
        };
        assert_eq!(value.as_str(), first);
        assert_eq!(rule.formula2().map(|v| v.as_str()), second);
        assert!(rule.uid().is_some());
    }
    let notes = sheet.comments().unwrap().unwrap();
    assert_eq!(notes.len(), 2);
    assert_eq!(
        notes.get("A2").unwrap().text,
        "Ryker Zhu:\r\nLEGACY_NOTE_002: native Excel cell note."
    );
    assert_eq!(
        notes.get("A3").unwrap().text,
        "Ryker Zhu:\r\nCOMMENT_COMMAND_002: comment inserted through Review ribbon."
    );
    for (name, reference) in [
        ("FixtureAmounts", "Sheet1!$E$2:$E$5"),
        ("_xlnm.Print_Area", "Sheet1!$A$1:$I$10"),
        ("_xlnm.Print_Titles", "Sheet1!$A:$A,Sheet1!$1:$1"),
    ] {
        assert_eq!(
            workbook
                .defined_names()
                .iter()
                .find(|value| value.name == name)
                .unwrap()
                .reference,
            reference
        );
    }
    let setup = sheet.page_setup().unwrap().unwrap();
    assert_eq!(
        setup.orientation,
        Some(litchi_xlsx::page_setup::Orientation::Landscape)
    );
    assert_eq!(setup.scale.unwrap().get(), 80);
    assert_eq!(setup.first_page.unwrap().get(), 3);
    assert_eq!(
        setup.horizontal_dpi,
        Some(litchi_xlsx::page_setup::Dpi::UNSPECIFIED)
    );
    assert_eq!(
        setup.vertical_dpi,
        Some(litchi_xlsx::page_setup::Dpi::UNSPECIFIED)
    );
    let margins = sheet.page_margins().unwrap().unwrap();
    assert_eq!(
        margins.left(),
        litchi_xlsx::page_margins::PageMargin::from_inches(0.7).unwrap()
    );
    assert_eq!(
        margins.top(),
        litchi_xlsx::page_margins::PageMargin::from_inches(0.5).unwrap()
    );
    formula(&sheet, "B6", "SUM(B2:B5)", "14");
    formula(&sheet, "E6", "SUM(E2:E5)", "95");
}
