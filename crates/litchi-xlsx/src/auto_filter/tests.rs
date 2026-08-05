use super::codec::*;
use super::model::*;
use super::package::parse_auto_filter;
use crate::sort::SortMethod;
use litchi_opc::PackURI;
fn fixture(path: &str) -> Definition {
    let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../..");
    let p = litchi_opc::phys_pkg::OwnedPhysPkgReader::open(root.join(path)).unwrap();
    let u = PackURI::new("/xl/worksheets/sheet1.xml").unwrap();
    parse_auto_filter(&p.blob_for(&u).unwrap())
        .unwrap()
        .unwrap()
}
#[test]
fn parses_bundled_fixtures() {
    let custom = fixture("test-data/ooxml/xlsx/autofilter.xlsx");
    assert_eq!(custom.reference.unwrap().as_str(), "A1:C5");
    assert!(matches!(
        custom.columns[0].payload,
        Some(Payload::Custom(_))
    ));
    let values = fixture("test-data/ooxml/xlsx/autofilternamedrange.xlsx");
    assert!(matches!(&values.columns[0].payload,Some(Payload::Values(v))if v.items.len()==2));
    let date = fixture("test-data/libreoffice-core/sc/qa/unit/data/xlsx/dateAutofilter.xlsx");
    assert!(matches!(&date.columns[0].payload,Some(Payload::Values(v))if v.items.len()==2));
    let top = fixture("test-data/libreoffice-core/sc/qa/unit/data/xlsx/tdf143068_top10filter.xlsx");
    assert!(
        matches!(&top.columns[0].payload,Some(Payload::Top10(v))if v.value==4.0&&v.filter_value==Some(7.0))
    );
    let buttons =
        fixture("test-data/libreoffice-core/sc/qa/unit/data/xlsx/autofilterShowButton.xlsx");
    assert_eq!(buttons.columns.len(), 4);
    assert!(buttons.columns.iter().all(|v| !v.show_button));
    for f in [
        "test-data/ooxml/xlsx/sortconditionref.xlsx",
        "test-data/ooxml/xlsx/sortconditionref2.xlsx",
    ] {
        let v = fixture(f);
        assert_eq!(v.sort_state.unwrap().conditions.len(), 1);
    }
}
#[test]
fn parses_all_variants_strict_and_mce() {
    let xml=br#"<s:worksheet xmlns:s="http://purl.oclc.org/ooxml/spreadsheetml/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><s:autoFilter ref="A1:F9"><s:filterColumn colId="0"><s:filters blank="1"><s:filter val="x"/><s:dateGroupItem year="2024" month="2" dateTimeGrouping="month"/></s:filters></s:filterColumn><s:filterColumn colId="1"><s:customFilters and="1"><s:customFilter operator="greaterThan" val="2"/></s:customFilters></s:filterColumn><s:filterColumn colId="2"><s:dynamicFilter type="today" val="4"/></s:filterColumn><s:filterColumn colId="3"><s:colorFilter dxfId="2" cellColor="0"/></s:filterColumn><s:filterColumn colId="4"><mc:AlternateContent><mc:Choice Requires="x14"><x14:iconFilter iconSet="3Arrows" iconId="2"/></mc:Choice><mc:Fallback><s:customFilters><s:customFilter val="fallback"/></s:customFilters></mc:Fallback></mc:AlternateContent></s:filterColumn><s:filterColumn colId="5"><s:top10 percent="1" val="10"/></s:filterColumn><s:sortState ref="A2:F9" caseSensitive="1" sortMethod="none"><s:sortCondition ref="D2:D9" sortBy="cellColor" dxfId="2"/><s:sortCondition ref="E2:E9" sortBy="icon" iconSet="3Arrows" iconId="1"/></s:sortState></s:autoFilter></s:worksheet>"#;
    let v = parse_auto_filter(xml).unwrap().unwrap();
    assert_eq!(v.columns.len(), 6);
    assert!(matches!(v.columns[4].payload, Some(Payload::Custom(_))));
    let sort = v.sort_state.unwrap();
    assert_eq!(sort.sort_method, Some(SortMethod::None));
    assert_eq!(sort.conditions.len(), 2);
}

#[test]
fn authored_filter_payloads_round_trip_through_shared_serializer() {
    let payloads = vec![
        Payload::Values(
            Values::new(
                true,
                Calendar::Gregorian,
                vec![
                    Item::Value("North".into()),
                    Item::DateGroup(
                        DateGroup::new(2026, Some(7), Some(26), None, None, None, Grouping::Day)
                            .unwrap(),
                    ),
                ],
            )
            .unwrap(),
        ),
        Payload::Custom(
            Customs::new(
                true,
                vec![
                    Custom::new(Operator::GreaterThan, "10").unwrap(),
                    Custom::new(Operator::LessThan, "20").unwrap(),
                ],
            )
            .unwrap(),
        ),
        Payload::Dynamic(Dynamic::new(DynamicType::ThisMonth, Some(1.5), Some(2.5)).unwrap()),
        Payload::Color(Color::new(4, false)),
        Payload::Icon(Icon::new(IconSet::ThreeArrows, 2).unwrap()),
        Payload::Top10(Top10::new(false, true, 25.0, Some(9.0)).unwrap()),
    ];
    let mut authored = Definition::new(Some(Range::new("A1:F20").unwrap()));
    for (column_id, payload) in payloads.into_iter().enumerate() {
        let mut column = Column::new(column_id as u32).unwrap();
        column.set_payload(Some(payload));
        authored.columns.push(column);
    }

    let xml = write_auto_filter_fragment(&authored).unwrap();
    assert_eq!(parse_auto_filter_fragment(&xml).unwrap(), authored);
}

#[test]
fn rejects_malformed_and_security_cases() {
    for xml in [
        r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><autoFilter ref="B2:A1"/></worksheet>"#,
        r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><autoFilter ref="A1:A2"><filterColumn colId="1"/></autoFilter></worksheet>"#,
        r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><autoFilter ref="A1:B2"><filterColumn colId="0"/><filterColumn colId="0"/></autoFilter></worksheet>"#,
        r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><autoFilter ref="A1:B2"><filterColumn colId="0"><customFilters/></filterColumn></autoFilter></worksheet>"#,
        r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><autoFilter ref="A1:B2"><filterColumn colId="0"><top10 percent="1" val="101"/></filterColumn></autoFilter></worksheet>"#,
        r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><autoFilter><sortState ref="A1"><sortCondition ref="A1" sortBy="icon"/></sortState></autoFilter></worksheet>"#,
        r#"<!DOCTYPE x><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>"#,
        r#"<?bad x?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>"#,
        r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><autoFilter ref="A1"><filterColumn colId="0"><filters><filter val="&bogus;"/></filters></filterColumn></autoFilter></worksheet>"#,
    ] {
        assert!(parse_auto_filter(xml.as_bytes()).is_err(), "{xml}");
    }
    let conditions = "<sortCondition ref=\"A1\"/>".repeat(MAX_SORT_CONDITIONS + 1);
    let xml = format!(
        "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><autoFilter><sortState ref=\"A1\">{conditions}</sortState></autoFilter></worksheet>"
    );
    assert!(parse_auto_filter(xml.as_bytes()).is_err());
}
