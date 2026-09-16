use super::*;
use litchi_opc::{OpcPackage, PackURI};
use litchi_sheet::Cell;
use litchi_sheet::Rect;
use litchi_sheet::view::{Color, Display, Mode, Position, Scale, Split, State};
use quick_xml::events::Event;
use quick_xml::name::ResolveResult;
use quick_xml::reader::NsReader;
const T: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const S: &str = "http://purl.oclc.org/ooxml/spreadsheetml/main";
const MCE: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";
fn fixture(bytes: &[u8]) -> Collection {
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
    assert_eq!(
        poi.entries()[0].view().selections[0].active_cell(),
        Cell::from_a1("A4").unwrap()
    );
    let lo = fixture(include_bytes!(
        "../../../../test-data/libreoffice-core/sc/qa/unit/data/xlsx/freezePaneStartCell.xlsx"
    ));
    let view = lo.entries()[0].view();
    assert_eq!(view.selections.len(), 4);
    assert_eq!(
        view.pane.as_ref().unwrap().horizontal,
        Some(Split::new(5.0).unwrap())
    );
    assert_eq!(view.pane.as_ref().unwrap().state, State::Frozen);
}

#[test]
fn reads_strict_mce_pivot_selection_and_defaults() {
    let xml = format!(
        r#"<s:worksheet xmlns:s="{S}" xmlns:r="http://purl.oclc.org/ooxml/officeDocument/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:u="urn:unsupported"><mc:AlternateContent><mc:Choice Requires="u"><u:no/></mc:Choice><mc:Fallback><s:sheetViews><s:sheetView workbookViewId="7" colorId="64" zoomScale="125" zoomScaleNormal="0" zoomScaleSheetLayoutView="135" zoomScalePageLayoutView="145"><s:pane xSplit="2" topLeftCell="C1" state="frozen"/><s:selection activeCell="C1" sqref="C1:D2 F3"/><s:pivotSelection pane="bottomRight" showHeader="1" axis="axisRow" dimension="2" activeRow="11" activeCol="1" r:id="rId1"><s:pivotArea dataOnly="0" labelOnly="1" fieldPosition="0" offset="A1:B2"><s:references count="1"><s:reference field="9" count="0"/></s:references></s:pivotArea></s:pivotSelection><s:extLst><s:ext uri="urn:view"><u:ignored/></s:ext></s:extLst></s:sheetView></s:sheetViews></mc:Fallback></mc:AlternateContent></s:worksheet>"#
    );
    let collection = parse_worksheet_views(xml.as_bytes()).unwrap().unwrap();
    let entry = &collection.entries()[0];
    let view = entry.view();
    assert_eq!(view.display, Display::default());
    assert_eq!(view.mode, Mode::Normal);
    assert_eq!(view.color, Color::DEFAULT);
    assert_eq!(view.zoom.current, Scale::new(125).unwrap());
    assert_eq!(view.zoom.normal, None);
    assert_eq!(view.zoom.page_break_preview, Some(Scale::new(135).unwrap()));
    assert_eq!(view.zoom.page_layout, Some(Scale::new(145).unwrap()));
    assert_eq!(view.origin, Cell::from_a1("A1").unwrap());
    assert_eq!(view.selections[0].ranges().len(), 2);
    assert_eq!(
        view.selections[0].ranges(),
        &[
            Rect::from_a1("C1:D2").unwrap(),
            Rect::from_a1("F3").unwrap()
        ]
    );
    assert_eq!(view.pane.as_ref().unwrap().position, Position::TopLeft);
    assert_eq!(
        view.pane.as_ref().unwrap().top_left,
        Cell::from_a1("C1").unwrap()
    );
    let pivot = &entry.pivot_selections()[0];
    assert_eq!(pivot.pane(), Position::BottomRight);
    assert_eq!(pivot.axis(), Some(PivotSelectionAxis::Row));
    assert_eq!(pivot.relationship_id(), Some("rId1"));
    assert_eq!(pivot.area().offset(), Some(Rect::from_a1("A1:B2").unwrap()));
    assert!(!pivot.area().markup().is_empty());
    assert_eq!(entry.extensions()[0].uri(), "urn:view");
}

#[test]
fn accepts_automatic_remembered_zoom_defaults() {
    let xml = format!(
        r#"<worksheet xmlns="{T}"><sheetViews><sheetView workbookViewId="0" zoomScale="100" zoomScaleNormal="0" zoomScaleSheetLayoutView="0" zoomScalePageLayoutView="0"/></sheetViews></worksheet>"#
    );
    let parsed = parse_worksheet_views(xml.as_bytes()).unwrap().unwrap();
    let view = parsed.entries()[0].view();
    assert_eq!(view.zoom.current, Scale::DEFAULT);
    assert_eq!(view.zoom.normal, None);
    assert_eq!(view.zoom.page_layout, None);
    assert_eq!(view.zoom.page_break_preview, None);
}

#[test]
fn accepts_spec_absolute_a1_references() {
    let xml = format!(
        r#"<worksheet xmlns="{T}"><sheetViews><sheetView workbookViewId="0" topLeftCell="$XFD$1048576" zoomScaleNormal="0"><pane topLeftCell="A1"/><selection activeCell="$A$1" sqref="$A$1:$XFD$1048576"/></sheetView></sheetViews></worksheet>"#
    );
    let parsed = parse_worksheet_views(xml.as_bytes()).unwrap().unwrap();
    let view = parsed.entries()[0].view();
    assert_eq!(view.origin, Cell::from_a1("$XFD$1048576").unwrap());
    assert_eq!(
        view.selections[0].active_cell(),
        Cell::from_a1("$A$1").unwrap()
    );
    let retained = std::str::from_utf8(parsed.entries()[0].retained_xml()).unwrap();
    // The retained sheetView re-declares what it inherits from the worksheet
    // root, so it parses standalone (change 0653).
    assert!(retained.starts_with(&format!(
        r#"<sheetView workbookViewId="0" topLeftCell="$XFD$1048576" zoomScaleNormal="0" xmlns="{T}">"#
    )));
    assert!(retained.contains(r#"topLeftCell="A1""#));
    assert!(retained.contains(r#"sqref="$A$1:$XFD$1048576""#));
    assert!(retained.find("<pane").unwrap() < retained.find("<selection").unwrap());
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

/// Every name in `markup` that does not resolve the way it resolves in the part
/// the fragment was cut from: an element whose name reaches no namespace at
/// all, or a prefixed name whose prefix has no declaration in the fragment.
///
/// A fragment cut out of a processed part has to carry the declarations it
/// inherits, so this is empty for the markup the accessors publish.
fn unresolved_names(markup: &[u8]) -> Vec<String> {
    let mut reader = NsReader::from_reader(markup);
    let mut unresolved = Vec::new();
    loop {
        let event = reader.read_event().unwrap().into_owned();
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(element) | Event::Empty(element) => {
                if !matches!(namespace, ResolveResult::Bound(_)) {
                    unresolved.push(String::from_utf8_lossy(element.name().as_ref()).into_owned());
                }
                for attribute in element.attributes() {
                    let attribute = attribute.unwrap();
                    if matches!(
                        resolver.resolve_attribute(attribute.key).0,
                        ResolveResult::Unknown(_)
                    ) {
                        unresolved
                            .push(String::from_utf8_lossy(attribute.key.as_ref()).into_owned());
                    }
                }
            },
            Event::Eof => break,
            _ => {},
        }
    }
    unresolved
}

#[test]
fn retained_view_markup_resolves_standalone() {
    let xml = format!(
        concat!(
            r#"<worksheet xmlns="{}" xmlns:mc="{}" xmlns:p="urn:payload"><sheetViews>"#,
            r#"<sheetView workbookViewId="0"><pivotSelection pane="topLeft">"#,
            r#"<pivotArea dataOnly="0"><p:hint p:v="1"/></pivotArea></pivotSelection>"#,
            r#"<extLst><ext uri="urn:view"><p:payload/></ext></extLst>"#,
            r#"</sheetView></sheetViews></worksheet>"#,
        ),
        T, MCE
    );
    let collection = parse_worksheet_views(xml.as_bytes()).unwrap().unwrap();
    let entry = &collection.entries()[0];
    for markup in [
        entry.retained_xml(),
        entry.pivot_selections()[0].area().markup(),
        entry.extensions()[0].markup(),
    ] {
        assert!(
            unresolved_names(markup).is_empty(),
            "{}",
            String::from_utf8_lossy(markup)
        );
        assert!(String::from_utf8_lossy(markup).contains(&format!(r#"xmlns="{T}""#)));
    }
    assert!(
        String::from_utf8_lossy(entry.pivot_selections()[0].area().markup())
            .contains(r#"xmlns:p="urn:payload""#)
    );
}
