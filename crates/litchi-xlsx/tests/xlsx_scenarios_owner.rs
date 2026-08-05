use litchi_xlsx::scenarios::{
    Conformance, InputCell, Scenario, ScenarioCellReference, Scenarios, parse_worksheet_scenarios,
    write_worksheet_scenarios,
};

#[test]
fn standalone_scenarios_round_trip_through_the_owner() {
    let scenario = Scenario::new("baseline")
        .unwrap()
        .with_input_cells(vec![
            InputCell::new(ScenarioCellReference::new("A1").unwrap(), "10").unwrap(),
        ])
        .unwrap();
    let value = Scenarios::new(vec![scenario]).unwrap();

    fn accepts_canonical_owner(_: &Scenarios) {}
    accepts_canonical_owner(&value);

    for conformance in [Conformance::Transitional, Conformance::Strict] {
        let fragment = write_worksheet_scenarios(&value, conformance).unwrap();
        let namespace = match conformance {
            Conformance::Transitional => {
                "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
            },
            Conformance::Strict => "http://purl.oclc.org/ooxml/spreadsheetml/main",
        };
        let document = format!(r#"<worksheet xmlns="{namespace}">{fragment}</worksheet>"#);
        let parsed = parse_worksheet_scenarios(document.as_bytes())
            .unwrap()
            .unwrap();
        assert_eq!(parsed, value);
    }
}
