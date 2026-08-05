use litchi_xlsx::cell_watches::{
    CellWatchConformance, CellWatchReference, CellWatches, parse_cell_watches, write_cell_watches,
};

#[test]
fn standalone_cell_watches_round_trip_through_the_owner() {
    let value = CellWatches::new(vec![CellWatchReference::new("B2").unwrap()]).unwrap();

    fn accepts_canonical_owner(_: &CellWatches) {}
    accepts_canonical_owner(&value);

    for conformance in [
        CellWatchConformance::Transitional,
        CellWatchConformance::Strict,
    ] {
        let fragment = write_cell_watches(&value, conformance).unwrap();
        let namespace = match conformance {
            CellWatchConformance::Transitional => {
                "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
            },
            CellWatchConformance::Strict => "http://purl.oclc.org/ooxml/spreadsheetml/main",
        };
        let document = format!(r#"<worksheet xmlns="{namespace}">{fragment}</worksheet>"#);
        let parsed = parse_cell_watches(document.as_bytes()).unwrap().unwrap();
        assert_eq!(parsed, value);
    }
}
