//! Focused semantic chartsheet facade coverage.

use super::{Conformance, View, parse_chartsheet, write_chartsheet};

#[test]
fn semantic_view_children_round_trip_through_the_flat_facade() {
    let xml = br#"<x:chartsheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><x:sheetViews><x:sheetView tabSelected="1" zoomScale="125" workbookViewId="0" zoomToFit="0"/></x:sheetViews><x:drawing r:id="rIdDrawing"/></x:chartsheet>"#;
    let (conformance, sheet) = parse_chartsheet(xml).unwrap();

    assert_eq!(conformance, Conformance::Transitional);
    assert_eq!(
        sheet.views,
        vec![View {
            tab_selected: Some(true),
            zoom_scale: Some(125),
            workbook_view_id: 0,
            zoom_to_fit: Some(false),
        }]
    );
    assert_eq!(
        parse_chartsheet(&write_chartsheet(&sheet, conformance).unwrap())
            .unwrap()
            .1,
        sheet
    );
}
