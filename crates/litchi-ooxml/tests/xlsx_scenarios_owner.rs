use litchi_ooxml::xlsx::scenarios::{
    ScenarioCellReference, WorksheetScenario, WorksheetScenarioConformance,
    WorksheetScenarioInputCell, WorksheetScenarios, parse_worksheet_scenarios,
    write_worksheet_scenarios,
};

#[test]
fn host_reexports_the_canonical_scenarios_owner() {
    let scenario = WorksheetScenario::new("baseline")
        .unwrap()
        .with_input_cells(vec![
            WorksheetScenarioInputCell::new(ScenarioCellReference::new("A1").unwrap(), "10")
                .unwrap(),
        ])
        .unwrap();
    let value = WorksheetScenarios::new(vec![scenario]).unwrap();

    fn accepts_canonical_owner(_: &litchi_xlsx::scenarios::WorksheetScenarios) {}
    accepts_canonical_owner(&value);

    for conformance in [
        WorksheetScenarioConformance::Transitional,
        WorksheetScenarioConformance::Strict,
    ] {
        let fragment = write_worksheet_scenarios(&value, conformance).unwrap();
        let namespace = match conformance {
            WorksheetScenarioConformance::Transitional => {
                "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
            },
            WorksheetScenarioConformance::Strict => "http://purl.oclc.org/ooxml/spreadsheetml/main",
        };
        let document = format!(r#"<worksheet xmlns="{namespace}">{fragment}</worksheet>"#);
        let parsed = parse_worksheet_scenarios(document.as_bytes())
            .unwrap()
            .unwrap();
        assert_eq!(parsed, value);
    }
}
