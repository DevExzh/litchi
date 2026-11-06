//! Native styles, rich strings and merge lifecycle; no rendering or evaluation.
use litchi_xlsx::Package;
fn native(name: &str) -> Vec<u8> {
    std::fs::read(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/microsoft-office/excel")
            .join(name),
    )
    .unwrap()
}
fn members(bytes: &[u8]) -> std::collections::BTreeMap<String, Vec<u8>> {
    let reader = litchi_opc::phys_pkg::OwnedPhysPkgReader::from_bytes(
        bytes.to_vec(),
        litchi_opc::SourceLimits::default(),
    )
    .unwrap();
    reader
        .member_names()
        .unwrap()
        .into_iter()
        .map(|name| {
            let bytes = reader.read_member(&name).unwrap();
            (name, bytes)
        })
        .collect()
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
fn native_merge_anchor_and_number_formats_retain_local_style_values() {
    let bytes = native("excel-values-formulas-chart.xlsx");
    let p = Package::from_bytes(&bytes).unwrap();
    let book = p.workbook();
    let s = book.sheet("Sheet2").unwrap().unwrap();
    assert_eq!(
        s.merges()
            .unwrap()
            .map(|r| r.to_string())
            .collect::<Vec<_>>(),
        ["A1:C1"]
    );
    assert_eq!(
        s.cell("A1").unwrap().stored(),
        Some(&litchi_xlsx::Cell::Value(litchi_xlsx::Value::text(
            "Native workbook summary"
        )))
    );
    let style = s.style("A1").unwrap().unwrap().format().unwrap();
    let font = style.font().unwrap();
    assert_eq!(font.name(), Some("Aptos Narrow"));
    assert_eq!(font.size(), Some(12.));
    assert_eq!(font.bold(), Some(true));
    assert_eq!(font.italic(), None);
    assert_eq!(
        style.alignment().unwrap().horizontal,
        Some(litchi_xlsx::style::stylesheet::alignment::Horizontal::Center)
    );
    let data = book.sheet("Sheet1").unwrap().unwrap();
    for at in ["B13", "C13", "B14", "C14"] {
        assert!(
            data.style(at)
                .unwrap()
                .unwrap()
                .format()
                .unwrap()
                .number_format()
                .is_date_time()
        );
    }
    for (at, code) in [("B15", "0%"), ("C15", "0.00%")] {
        assert_eq!(
            data.style(at)
                .unwrap()
                .unwrap()
                .format()
                .unwrap()
                .number_format()
                .code(),
            Some(code)
        );
    }
    assert!(
        data.style("B20")
            .unwrap()
            .unwrap()
            .format()
            .unwrap()
            .quote_prefix()
    );
    assert!(p.to_bytes().unwrap() == bytes);
}
#[test]
fn native_rich_text_runs_keep_fonts_and_explicit_vs_inherited_properties() {
    let p = Package::from_bytes(&native("excel-metadata.xlsx")).unwrap();
    let s = p.workbook().sheet("Sheet1").unwrap().unwrap();
    let rich = s.rich_text("A2").unwrap().unwrap();
    assert_eq!(
        rich.runs().iter().map(|r| r.text()).collect::<Vec<_>>(),
        ["RICH_TEXT_015 ", "bold", " ", "italic"]
    );
    assert!(!rich.runs()[0].is_formatted());
    for (i, (bold, italic)) in [(false, false), (true, false), (false, false), (false, true)]
        .into_iter()
        .enumerate()
    {
        let font = rich.runs()[i].font().unwrap();
        assert_eq!((font.bold, font.italic), (bold, italic));
        if i == 0 {
            assert!(font.name.is_none() && font.size.is_none());
        } else {
            assert_eq!(font.name.as_deref(), Some("Aptos Narrow"));
            assert_eq!(font.size, Some(12.));
        }
    }
}
#[test]
fn native_merge_remove_create_and_inverse_keep_unrelated_package_parts() {
    let bytes = native("excel-values-formulas-chart.xlsx");
    let p = Package::from_bytes(&bytes).unwrap();
    let mut edit = p.edit().unwrap();
    {
        let mut s = edit.sheet("Sheet2").unwrap().unwrap();
        s.unmerge("A1").unwrap();
        s.merge("A10:C10").unwrap();
    }
    let commit = edit.commit().unwrap();
    let output = tempfile::NamedTempFile::with_suffix(".xlsx").unwrap();
    commit.package().save(output.path()).unwrap();
    let after = std::fs::read(output.path()).unwrap();
    let read = Package::from_bytes(&after).unwrap();
    let s = read.workbook().sheet("Sheet2").unwrap().unwrap();
    assert_eq!(
        s.merges()
            .unwrap()
            .map(|r| r.to_string())
            .collect::<Vec<_>>(),
        ["A10:C10"]
    );
    assert_eq!(
        s.cell("A1").unwrap().stored(),
        p.workbook()
            .sheet("Sheet2")
            .unwrap()
            .unwrap()
            .cell("A1")
            .unwrap()
            .stored()
    );
    let original = members(&bytes);
    let changed = members(&after);
    let diff = original
        .iter()
        .filter(|(n, b)| changed[*n] != **b)
        .map(|(n, _)| n.as_str())
        .collect::<Vec<_>>();
    assert_eq!(
        original.keys().collect::<Vec<_>>(),
        changed.keys().collect::<Vec<_>>()
    );
    assert_eq!(diff, ["xl/worksheets/sheet2.xml"]);
    let restored = members(
        &read
            .apply(&commit.patch().inverse())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap(),
    );
    assert_eq!(
        original.keys().collect::<Vec<_>>(),
        restored.keys().collect::<Vec<_>>()
    );
    for (n, b) in original {
        assert!(restored[&n] == b, "inverse changed {n}");
    }
}
#[test]
fn native_rich_run_text_edit_retains_fonts_and_unrelated_native_parts() {
    let bytes = native("excel-metadata.xlsx");
    let p = Package::from_bytes(&bytes).unwrap();
    let mut rich = p
        .workbook()
        .sheet("Sheet1")
        .unwrap()
        .unwrap()
        .rich_text("A2")
        .unwrap()
        .unwrap();
    rich.set_run_text(1, "BOLD").unwrap();
    let mut edit = p.edit().unwrap();
    edit.sheet("Sheet1")
        .unwrap()
        .unwrap()
        .set_rich_text("A2", rich.clone())
        .unwrap();
    let commit = edit.commit().unwrap();
    let after = commit.package().to_bytes().unwrap();
    let read = Package::from_bytes(&after).unwrap();
    assert_eq!(
        read.workbook()
            .sheet("Sheet1")
            .unwrap()
            .unwrap()
            .rich_text("A2")
            .unwrap()
            .unwrap(),
        rich
    );
    let original = members(&bytes);
    let changed = members(&after);
    assert!(original.contains_key("xl/calcChain.xml"));
    assert_eq!(
        original
            .keys()
            .filter(|n| n.as_str() != "xl/calcChain.xml")
            .collect::<Vec<_>>(),
        changed.keys().collect::<Vec<_>>()
    );
    for (part, data) in &original {
        match part.as_str() {
            "xl/worksheets/sheet1.xml" | "xl/calcChain.xml" => {},
            "xl/workbook.xml" => assert_eq!(
                workbook_events_without_calculation_properties(data),
                workbook_events_without_calculation_properties(&changed[part])
            ),
            "[Content_Types].xml" | "xl/_rels/workbook.xml.rels" => {
                assert_eq!(
                    manifest_children(data, true),
                    manifest_children(&changed[part], true)
                );
                assert!(
                    !std::str::from_utf8(&changed[part])
                        .unwrap()
                        .contains("calcChain")
                );
            },
            _ => assert!(changed[part] == *data, "changed {part}"),
        }
    }
    let metadata = read.workbook().calculation_metadata().unwrap();
    let props = metadata.properties().unwrap();
    assert_eq!(props.calculation_id(), 0);
    assert!(props.full_calculation_on_load() && props.force_full_calculation());
    assert!(!props.calculation_completed());
    assert!(props.calculate_on_save());
    let restored = members(
        &read
            .apply(&commit.patch().inverse())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap(),
    );
    assert_eq!(
        original.keys().collect::<Vec<_>>(),
        restored.keys().collect::<Vec<_>>()
    );
    for (part, data) in original {
        if part == "xl/_rels/workbook.xml.rels" {
            assert_eq!(
                manifest_children(&data, false),
                manifest_children(&restored[&part], false)
            );
        } else {
            assert!(restored[&part] == data, "inverse changed {part}");
        }
    }
    assert!(p.to_bytes().unwrap() == bytes);
}
