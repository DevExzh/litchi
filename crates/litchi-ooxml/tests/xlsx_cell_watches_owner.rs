use litchi_ooxml::xlsx::cell_watches::{
    CellWatchReference, WorksheetCellWatchConformance, WorksheetCellWatches,
    parse_worksheet_cell_watches, write_worksheet_cell_watches,
};

#[test]
fn host_reexports_the_canonical_cell_watches_owner() {
    let value = WorksheetCellWatches::new(vec![CellWatchReference::new("B2").unwrap()]).unwrap();

    fn accepts_canonical_owner(_: &litchi_xlsx::cell_watches::WorksheetCellWatches) {}
    accepts_canonical_owner(&value);

    for conformance in [
        WorksheetCellWatchConformance::Transitional,
        WorksheetCellWatchConformance::Strict,
    ] {
        let fragment = write_worksheet_cell_watches(&value, conformance).unwrap();
        let namespace = match conformance {
            WorksheetCellWatchConformance::Transitional => {
                "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
            },
            WorksheetCellWatchConformance::Strict => {
                "http://purl.oclc.org/ooxml/spreadsheetml/main"
            },
        };
        let document = format!(r#"<worksheet xmlns="{namespace}">{fragment}</worksheet>"#);
        let parsed = parse_worksheet_cell_watches(document.as_bytes())
            .unwrap()
            .unwrap();
        assert_eq!(parsed, value);
    }
}
