use super::*;
use litchi_opc::{OpcPackage, PackURI};
const T: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const S: &str = "http://purl.oclc.org/ooxml/spreadsheetml/main";
fn fixture(bytes: &[u8]) -> Views {
    let package = OpcPackage::from_bytes(bytes).unwrap();
    let part = package
        .get_part(&PackURI::new("/xl/worksheets/sheet1.xml").unwrap())
        .unwrap();
    parse_worksheet_views(part.blob()).unwrap().unwrap()
}

#[test]
fn reads_poi_and_libreoffice_view_fixtures() {
    let poi = fixture(include_bytes!(
        "../../../../test-data/poi/test-data/spreadsheet/right-to-left.xlsx"
    ));
    assert_eq!(poi.views()[0].selections()[0].active_cell().as_str(), "A4");
    let lo = fixture(include_bytes!(
        "../../../../test-data/libreoffice-core/sc/qa/unit/data/xlsx/freezePaneStartCell.xlsx"
    ));
    let view = &lo.views()[0];
    assert_eq!(view.selections().len(), 4);
    assert_eq!(view.pane().unwrap().x_split(), Some(5.0));
    assert_eq!(view.pane().unwrap().state(), PaneState::Frozen);
}

#[test]
fn reads_strict_mce_pivot_selection_and_defaults() {
    let xml = format!(
        r#"<s:worksheet xmlns:s="{S}" xmlns:r="http://purl.oclc.org/ooxml/officeDocument/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:u="urn:unsupported"><mc:AlternateContent><mc:Choice Requires="u"><u:no/></mc:Choice><mc:Fallback><s:sheetViews><s:sheetView workbookViewId="7" colorId="64" zoomScale="125" zoomScaleNormal="0"><s:pane xSplit="2" topLeftCell="C1" state="frozen"/><s:selection activeCell="C1" sqref="C1:D2 F3"/><s:pivotSelection pane="bottomRight" showHeader="1" axis="axisRow" dimension="2" activeRow="11" activeCol="1" r:id="rId1"><s:pivotArea dataOnly="0" labelOnly="1" fieldPosition="0" offset="A1:B2"><s:references count="1"><s:reference field="9" count="0"/></s:references></s:pivotArea></s:pivotSelection><s:extLst><s:ext uri="urn:view"><u:ignored/></s:ext></s:extLst></s:sheetView></s:sheetViews></mc:Fallback></mc:AlternateContent></s:worksheet>"#
    );
    let collection = parse_worksheet_views(xml.as_bytes()).unwrap().unwrap();
    let view = &collection.views()[0];
    assert!(view.show_grid_lines());
    assert_eq!(view.view_type(), ViewType::Normal);
    assert_eq!(view.zoom_scale_normal(), 0);
    assert_eq!(view.selections()[0].sqref().ranges().len(), 2);
    let pivot = &view.pivot_selections()[0];
    assert_eq!(pivot.axis(), Some(PivotSelectionAxis::Row));
    assert_eq!(pivot.relationship_id(), Some("rId1"));
    assert_eq!(pivot.area().offset().unwrap().as_str(), "A1:B2");
    assert!(!pivot.area().markup().is_empty());
    assert_eq!(view.extensions()[0].uri(), "urn:view");
}

#[test]
fn accepts_automatic_remembered_zoom_defaults() {
    let xml = format!(
        r#"<worksheet xmlns="{T}"><sheetViews><sheetView workbookViewId="0" zoomScale="100" zoomScaleNormal="0" zoomScaleSheetLayoutView="0" zoomScalePageLayoutView="0"/></sheetViews></worksheet>"#
    );
    let parsed = parse_worksheet_views(xml.as_bytes()).unwrap().unwrap();
    let view = &parsed.views()[0];
    assert_eq!(view.zoom_scale(), 100);
    assert_eq!(view.zoom_scale_normal(), 0);
    assert_eq!(view.zoom_scale_sheet_layout_view(), 0);
    assert_eq!(view.zoom_scale_page_layout_view(), 0);
}

#[test]
fn accepts_spec_absolute_a1_references() {
    let xml = format!(
        r#"<worksheet xmlns="{T}"><sheetViews><sheetView workbookViewId="0" topLeftCell="$XFD$1048576"><selection activeCell="$A$1" sqref="$A$1:$XFD$1048576"/></sheetView></sheetViews></worksheet>"#
    );
    let parsed = parse_worksheet_views(xml.as_bytes()).unwrap().unwrap();
    let view = &parsed.views()[0];
    assert_eq!(view.top_left_cell().unwrap().as_str(), "$XFD$1048576");
    assert_eq!(view.selections()[0].active_cell().as_str(), "$A$1");
}

#[test]
fn rejects_invalid_view_grammar_values_and_security() {
    let cases = [
        format!(
            r#"<worksheet xmlns="{T}"><sheetViews><sheetView workbookViewId="0" view="bad"/></sheetViews></worksheet>"#
        ),
        format!(
            r#"<worksheet xmlns="{T}"><sheetViews><sheetView workbookViewId="0"><selection activeCell="XFE1"/></sheetView></sheetViews></worksheet>"#
        ),
        format!(
            r#"<worksheet xmlns="{T}"><sheetViews><sheetView workbookViewId="0"><selection sqref="A1" activeCellId="1"/></sheetView></sheetViews></worksheet>"#
        ),
        format!(
            r#"<worksheet xmlns="{T}"><sheetViews><sheetView workbookViewId="0"><pivotSelection axis="bad"><pivotArea/></pivotSelection></sheetView></sheetViews></worksheet>"#
        ),
        format!(
            r#"<worksheet xmlns="{T}"><sheetViews><sheetView workbookViewId="0"><pivotSelection/></sheetView></sheetViews></worksheet>"#
        ),
        format!(r#"<!DOCTYPE x><worksheet xmlns="{T}"/>"#),
    ];
    for xml in cases {
        assert!(parse_worksheet_views(xml.as_bytes()).is_err(), "{xml}");
    }
}
