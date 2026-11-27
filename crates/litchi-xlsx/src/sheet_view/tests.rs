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
    let package = OpcPackage::from_bytes(bytes, litchi_opc::SourceLimits::default()).unwrap();
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
    // root, so it parses standalone.
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

fn views_outcome(xml: &str) -> String {
    match parse_worksheet_views(xml.as_bytes()) {
        Ok(Some(value)) => {
            let entries = value
                .entries()
                .iter()
                .map(|entry| {
                    format!(
                        "[pane={} selections={} pivots={} extensions={}]",
                        entry.view().pane.is_some(),
                        entry.view().selections.len(),
                        entry.pivot_selections().len(),
                        entry.extensions().len()
                    )
                })
                .collect::<String>();
            format!("ok views={entries} extensions={}", value.extensions().len())
        },
        Ok(None) => "none".to_owned(),
        Err(error) => format!("{error:?}"),
    }
}

fn assert_rows(actual: &[(String, String)], expected: &[(&str, &str)]) {
    let same = actual.len() == expected.len()
        && actual.iter().zip(expected).all(
            |((label, outcome), (expected_label, expected_outcome))| {
                label == expected_label && outcome == expected_outcome
            },
        );
    let observed: String = actual
        .iter()
        .map(|(label, outcome)| format!("({label:?}, {outcome:?}),\n"))
        .collect();
    assert!(same, "observed rows:\n{observed}");
}

#[test]
fn collection_child_order_matrix_is_stable() {
    let view = ("view", r#"<sheetView workbookViewId="0"/>"#);
    let open_view = ("open view", r#"<sheetView workbookViewId="0"></sheetView>"#);
    let list = ("extLst", r#"<extLst><ext uri="urn:a"/></extLst>"#);
    let open_list = (
        "open extLst",
        r#"<extLst><ext uri="urn:b"><y:data xmlns:y="urn:y"/></ext></extLst>"#,
    );
    let empty_list = ("empty extLst", "<extLst/>");
    let unknown = ("unknown", "<unknown/>");
    let children = [view, open_view, list, open_list, empty_list];
    let mut cases: Vec<Vec<(&str, &str)>> = vec![Vec::new()];
    for first in children {
        cases.push(vec![first]);
    }
    for first in children {
        for second in children {
            cases.push(vec![first, second]);
        }
    }
    // Every pair again behind one view, so that an accepted order is not
    // hidden by the refusal of a collection without views.
    for second in children {
        for third in children {
            cases.push(vec![view, second, third]);
        }
    }
    cases.push(vec![unknown]);
    cases.push(vec![view, unknown]);
    cases.push(vec![view, list, unknown]);
    cases.push(vec![(
        "foreign view",
        r#"<q:sheetView xmlns:q="urn:other" workbookViewId="0"/>"#,
    )]);
    let mut actual: Vec<(String, String)> = cases
        .iter()
        .map(|case| {
            let label = case.iter().map(|(name, _)| *name).collect::<Vec<_>>();
            let body: String = case.iter().map(|(_, xml)| *xml).collect();
            (
                label.join(","),
                views_outcome(&format!(
                    r#"<worksheet xmlns="{T}"><sheetViews>{body}</sheetViews></worksheet>"#
                )),
            )
        })
        .collect();
    // A second collection is refused whatever the first one reached.
    for (label, body) in [
        ("no collection", String::new()),
        ("self-closing collection", "<sheetViews/>".to_owned()),
        (
            "two collections",
            format!(
                "<sheetViews>{0}</sheetViews><sheetViews>{0}</sheetViews>",
                view.1
            ),
        ),
        (
            "collection with extLst, then collection",
            format!(
                "<sheetViews>{}{}</sheetViews><sheetViews>{}</sheetViews>",
                view.1, list.1, view.1
            ),
        ),
        (
            "collection with extLst, then self-closing collection",
            format!("<sheetViews>{}{}</sheetViews><sheetViews/>", view.1, list.1),
        ),
    ] {
        actual.push((
            label.to_owned(),
            views_outcome(&format!(r#"<worksheet xmlns="{T}">{body}</worksheet>"#)),
        ));
    }
    let expected: &[(&str, &str)] = &[
        (
            "",
            "Invalid(\"sheetViews must contain at least one sheetView\")",
        ),
        (
            "view",
            "ok views=[pane=false selections=0 pivots=0 extensions=0] extensions=0",
        ),
        (
            "open view",
            "ok views=[pane=false selections=0 pivots=0 extensions=0] extensions=0",
        ),
        (
            "extLst",
            "Invalid(\"sheetViews must contain at least one sheetView\")",
        ),
        (
            "open extLst",
            "Invalid(\"sheetViews must contain at least one sheetView\")",
        ),
        (
            "empty extLst",
            "Invalid(\"sheetViews must contain at least one sheetView\")",
        ),
        (
            "view,view",
            "ok views=[pane=false selections=0 pivots=0 extensions=0][pane=false selections=0 pivots=0 extensions=0] extensions=0",
        ),
        (
            "view,open view",
            "ok views=[pane=false selections=0 pivots=0 extensions=0][pane=false selections=0 pivots=0 extensions=0] extensions=0",
        ),
        (
            "view,extLst",
            "ok views=[pane=false selections=0 pivots=0 extensions=0] extensions=1",
        ),
        (
            "view,open extLst",
            "ok views=[pane=false selections=0 pivots=0 extensions=0] extensions=1",
        ),
        (
            "view,empty extLst",
            "ok views=[pane=false selections=0 pivots=0 extensions=0] extensions=0",
        ),
        (
            "open view,view",
            "ok views=[pane=false selections=0 pivots=0 extensions=0][pane=false selections=0 pivots=0 extensions=0] extensions=0",
        ),
        (
            "open view,open view",
            "ok views=[pane=false selections=0 pivots=0 extensions=0][pane=false selections=0 pivots=0 extensions=0] extensions=0",
        ),
        (
            "open view,extLst",
            "ok views=[pane=false selections=0 pivots=0 extensions=0] extensions=1",
        ),
        (
            "open view,open extLst",
            "ok views=[pane=false selections=0 pivots=0 extensions=0] extensions=1",
        ),
        (
            "open view,empty extLst",
            "ok views=[pane=false selections=0 pivots=0 extensions=0] extensions=0",
        ),
        (
            "extLst,view",
            "Invalid(\"sheetView appears after sheetViews extLst\")",
        ),
        (
            "extLst,open view",
            "Invalid(\"sheetView appears after sheetViews extLst\")",
        ),
        ("extLst,extLst", "Invalid(\"duplicate sheetViews extLst\")"),
        (
            "extLst,open extLst",
            "Invalid(\"duplicate sheetViews extLst\")",
        ),
        (
            "extLst,empty extLst",
            "Invalid(\"duplicate sheetViews extLst\")",
        ),
        (
            "open extLst,view",
            "Invalid(\"sheetView appears after sheetViews extLst\")",
        ),
        (
            "open extLst,open view",
            "Invalid(\"sheetView appears after sheetViews extLst\")",
        ),
        (
            "open extLst,extLst",
            "Invalid(\"duplicate sheetViews extLst\")",
        ),
        (
            "open extLst,open extLst",
            "Invalid(\"duplicate sheetViews extLst\")",
        ),
        (
            "open extLst,empty extLst",
            "Invalid(\"duplicate sheetViews extLst\")",
        ),
        (
            "empty extLst,view",
            "Invalid(\"sheetView appears after sheetViews extLst\")",
        ),
        (
            "empty extLst,open view",
            "Invalid(\"sheetView appears after sheetViews extLst\")",
        ),
        (
            "empty extLst,extLst",
            "Invalid(\"duplicate sheetViews extLst\")",
        ),
        (
            "empty extLst,open extLst",
            "Invalid(\"duplicate sheetViews extLst\")",
        ),
        (
            "empty extLst,empty extLst",
            "Invalid(\"duplicate sheetViews extLst\")",
        ),
        (
            "view,view,view",
            "ok views=[pane=false selections=0 pivots=0 extensions=0][pane=false selections=0 pivots=0 extensions=0][pane=false selections=0 pivots=0 extensions=0] extensions=0",
        ),
        (
            "view,view,open view",
            "ok views=[pane=false selections=0 pivots=0 extensions=0][pane=false selections=0 pivots=0 extensions=0][pane=false selections=0 pivots=0 extensions=0] extensions=0",
        ),
        (
            "view,view,extLst",
            "ok views=[pane=false selections=0 pivots=0 extensions=0][pane=false selections=0 pivots=0 extensions=0] extensions=1",
        ),
        (
            "view,view,open extLst",
            "ok views=[pane=false selections=0 pivots=0 extensions=0][pane=false selections=0 pivots=0 extensions=0] extensions=1",
        ),
        (
            "view,view,empty extLst",
            "ok views=[pane=false selections=0 pivots=0 extensions=0][pane=false selections=0 pivots=0 extensions=0] extensions=0",
        ),
        (
            "view,open view,view",
            "ok views=[pane=false selections=0 pivots=0 extensions=0][pane=false selections=0 pivots=0 extensions=0][pane=false selections=0 pivots=0 extensions=0] extensions=0",
        ),
        (
            "view,open view,open view",
            "ok views=[pane=false selections=0 pivots=0 extensions=0][pane=false selections=0 pivots=0 extensions=0][pane=false selections=0 pivots=0 extensions=0] extensions=0",
        ),
        (
            "view,open view,extLst",
            "ok views=[pane=false selections=0 pivots=0 extensions=0][pane=false selections=0 pivots=0 extensions=0] extensions=1",
        ),
        (
            "view,open view,open extLst",
            "ok views=[pane=false selections=0 pivots=0 extensions=0][pane=false selections=0 pivots=0 extensions=0] extensions=1",
        ),
        (
            "view,open view,empty extLst",
            "ok views=[pane=false selections=0 pivots=0 extensions=0][pane=false selections=0 pivots=0 extensions=0] extensions=0",
        ),
        (
            "view,extLst,view",
            "Invalid(\"sheetView appears after sheetViews extLst\")",
        ),
        (
            "view,extLst,open view",
            "Invalid(\"sheetView appears after sheetViews extLst\")",
        ),
        (
            "view,extLst,extLst",
            "Invalid(\"duplicate sheetViews extLst\")",
        ),
        (
            "view,extLst,open extLst",
            "Invalid(\"duplicate sheetViews extLst\")",
        ),
        (
            "view,extLst,empty extLst",
            "Invalid(\"duplicate sheetViews extLst\")",
        ),
        (
            "view,open extLst,view",
            "Invalid(\"sheetView appears after sheetViews extLst\")",
        ),
        (
            "view,open extLst,open view",
            "Invalid(\"sheetView appears after sheetViews extLst\")",
        ),
        (
            "view,open extLst,extLst",
            "Invalid(\"duplicate sheetViews extLst\")",
        ),
        (
            "view,open extLst,open extLst",
            "Invalid(\"duplicate sheetViews extLst\")",
        ),
        (
            "view,open extLst,empty extLst",
            "Invalid(\"duplicate sheetViews extLst\")",
        ),
        (
            "view,empty extLst,view",
            "Invalid(\"sheetView appears after sheetViews extLst\")",
        ),
        (
            "view,empty extLst,open view",
            "Invalid(\"sheetView appears after sheetViews extLst\")",
        ),
        (
            "view,empty extLst,extLst",
            "Invalid(\"duplicate sheetViews extLst\")",
        ),
        (
            "view,empty extLst,open extLst",
            "Invalid(\"duplicate sheetViews extLst\")",
        ),
        (
            "view,empty extLst,empty extLst",
            "Invalid(\"duplicate sheetViews extLst\")",
        ),
        (
            "unknown",
            "Invalid(\"unexpected worksheet-view element 'unknown'\")",
        ),
        (
            "view,unknown",
            "Invalid(\"unexpected worksheet-view element 'unknown'\")",
        ),
        (
            "view,extLst,unknown",
            "Invalid(\"unexpected worksheet-view element 'unknown'\")",
        ),
        (
            "foreign view",
            "Invalid(\"unexpected worksheet-view element 'sheetView'\")",
        ),
        ("no collection", "none"),
        (
            "self-closing collection",
            "Invalid(\"sheetViews must contain at least one sheetView\")",
        ),
        (
            "two collections",
            "Invalid(\"duplicate worksheet sheetViews element\")",
        ),
        (
            "collection with extLst, then collection",
            "Invalid(\"duplicate worksheet sheetViews element\")",
        ),
        (
            "collection with extLst, then self-closing collection",
            "Invalid(\"sheetViews must contain at least one sheetView\")",
        ),
    ];
    assert_rows(&actual, expected);
}

#[test]
fn view_child_order_matrix_is_stable() {
    let pane = ("pane", r#"<pane xSplit="1"/>"#);
    let open_pane = ("open pane", r#"<pane xSplit="1"></pane>"#);
    let selection = ("selection", "<selection/>");
    let open_selection = ("open selection", "<selection></selection>");
    let pivot = ("pivot", "<pivotSelection><pivotArea/></pivotSelection>");
    let empty_pivot = ("empty pivot", "<pivotSelection/>");
    let list = ("extLst", r#"<extLst><ext uri="urn:a"/></extLst>"#);
    let empty_list = ("empty extLst", "<extLst/>");
    let unknown = ("unknown", "<unknown/>");
    let children = [
        pane,
        open_pane,
        selection,
        open_selection,
        pivot,
        empty_pivot,
        list,
        empty_list,
    ];
    let mut cases: Vec<Vec<(&str, &str)>> = vec![Vec::new()];
    for first in children {
        cases.push(vec![first]);
    }
    for first in children {
        for second in children {
            cases.push(vec![first, second]);
        }
    }
    cases.push(vec![pane, selection, pivot, list]);
    cases.push(vec![pane, selection, selection, pivot, pivot, empty_list]);
    cases.push(vec![selection, pivot, selection]);
    cases.push(vec![pivot, list, pivot]);
    cases.push(vec![unknown]);
    cases.push(vec![selection, unknown]);
    cases.push(vec![list, unknown]);
    cases.push(vec![(
        "foreign pane",
        r#"<q:pane xmlns:q="urn:other" xSplit="1"/>"#,
    )]);
    let mut actual: Vec<(String, String)> = cases
        .iter()
        .map(|case| {
            let label = case.iter().map(|(name, _)| *name).collect::<Vec<_>>();
            let body: String = case.iter().map(|(_, xml)| *xml).collect();
            (
                label.join(","),
                views_outcome(&format!(
                    r#"<worksheet xmlns="{T}"><sheetViews><sheetView workbookViewId="0">{body}</sheetView></sheetViews></worksheet>"#
                )),
            )
        })
        .collect();
    // Each view starts its own child order, whatever the previous view
    // reached and whichever form the previous view was written in.
    let last_children = [pane, selection, pivot, list];
    for (previous_label, previous_body) in [
        ("self-closing view", None),
        ("view with selection", Some(selection.1)),
        ("view with pivot", Some(pivot.1)),
        ("view with extLst", Some(list.1)),
        ("view with empty extLst", Some(empty_list.1)),
    ] {
        let previous = match previous_body {
            Some(body) => format!(r#"<sheetView workbookViewId="0">{body}</sheetView>"#),
            None => r#"<sheetView workbookViewId="0"/>"#.to_owned(),
        };
        for (label, child) in last_children {
            actual.push((
                format!("{previous_label}, then view with {label}"),
                views_outcome(&format!(
                    r#"<worksheet xmlns="{T}"><sheetViews>{previous}<sheetView workbookViewId="1">{child}</sheetView></sheetViews></worksheet>"#
                )),
            ));
        }
    }
    let expected: &[(&str, &str)] = &[
        (
            "",
            "ok views=[pane=false selections=0 pivots=0 extensions=0] extensions=0",
        ),
        (
            "pane",
            "ok views=[pane=true selections=0 pivots=0 extensions=0] extensions=0",
        ),
        (
            "open pane",
            "ok views=[pane=true selections=0 pivots=0 extensions=0] extensions=0",
        ),
        (
            "selection",
            "ok views=[pane=false selections=1 pivots=0 extensions=0] extensions=0",
        ),
        (
            "open selection",
            "ok views=[pane=false selections=1 pivots=0 extensions=0] extensions=0",
        ),
        (
            "pivot",
            "ok views=[pane=false selections=0 pivots=1 extensions=0] extensions=0",
        ),
        (
            "empty pivot",
            "Invalid(\"pivotSelection requires one pivotArea\")",
        ),
        (
            "extLst",
            "ok views=[pane=false selections=0 pivots=0 extensions=1] extensions=0",
        ),
        (
            "empty extLst",
            "ok views=[pane=false selections=0 pivots=0 extensions=0] extensions=0",
        ),
        ("pane,pane", "Invalid(\"duplicate worksheet-view pane\")"),
        (
            "pane,open pane",
            "Invalid(\"duplicate worksheet-view pane\")",
        ),
        (
            "pane,selection",
            "ok views=[pane=true selections=1 pivots=0 extensions=0] extensions=0",
        ),
        (
            "pane,open selection",
            "ok views=[pane=true selections=1 pivots=0 extensions=0] extensions=0",
        ),
        (
            "pane,pivot",
            "ok views=[pane=true selections=0 pivots=1 extensions=0] extensions=0",
        ),
        (
            "pane,empty pivot",
            "Invalid(\"pivotSelection requires one pivotArea\")",
        ),
        (
            "pane,extLst",
            "ok views=[pane=true selections=0 pivots=0 extensions=1] extensions=0",
        ),
        (
            "pane,empty extLst",
            "ok views=[pane=true selections=0 pivots=0 extensions=0] extensions=0",
        ),
        (
            "open pane,pane",
            "Invalid(\"duplicate worksheet-view pane\")",
        ),
        (
            "open pane,open pane",
            "Invalid(\"duplicate worksheet-view pane\")",
        ),
        (
            "open pane,selection",
            "ok views=[pane=true selections=1 pivots=0 extensions=0] extensions=0",
        ),
        (
            "open pane,open selection",
            "ok views=[pane=true selections=1 pivots=0 extensions=0] extensions=0",
        ),
        (
            "open pane,pivot",
            "ok views=[pane=true selections=0 pivots=1 extensions=0] extensions=0",
        ),
        (
            "open pane,empty pivot",
            "Invalid(\"pivotSelection requires one pivotArea\")",
        ),
        (
            "open pane,extLst",
            "ok views=[pane=true selections=0 pivots=0 extensions=1] extensions=0",
        ),
        (
            "open pane,empty extLst",
            "ok views=[pane=true selections=0 pivots=0 extensions=0] extensions=0",
        ),
        ("selection,pane", "Invalid(\"pane is out of schema order\")"),
        (
            "selection,open pane",
            "Invalid(\"pane is out of schema order\")",
        ),
        (
            "selection,selection",
            "ok views=[pane=false selections=2 pivots=0 extensions=0] extensions=0",
        ),
        (
            "selection,open selection",
            "ok views=[pane=false selections=2 pivots=0 extensions=0] extensions=0",
        ),
        (
            "selection,pivot",
            "ok views=[pane=false selections=1 pivots=1 extensions=0] extensions=0",
        ),
        (
            "selection,empty pivot",
            "Invalid(\"pivotSelection requires one pivotArea\")",
        ),
        (
            "selection,extLst",
            "ok views=[pane=false selections=1 pivots=0 extensions=1] extensions=0",
        ),
        (
            "selection,empty extLst",
            "ok views=[pane=false selections=1 pivots=0 extensions=0] extensions=0",
        ),
        (
            "open selection,pane",
            "Invalid(\"pane is out of schema order\")",
        ),
        (
            "open selection,open pane",
            "Invalid(\"pane is out of schema order\")",
        ),
        (
            "open selection,selection",
            "ok views=[pane=false selections=2 pivots=0 extensions=0] extensions=0",
        ),
        (
            "open selection,open selection",
            "ok views=[pane=false selections=2 pivots=0 extensions=0] extensions=0",
        ),
        (
            "open selection,pivot",
            "ok views=[pane=false selections=1 pivots=1 extensions=0] extensions=0",
        ),
        (
            "open selection,empty pivot",
            "Invalid(\"pivotSelection requires one pivotArea\")",
        ),
        (
            "open selection,extLst",
            "ok views=[pane=false selections=1 pivots=0 extensions=1] extensions=0",
        ),
        (
            "open selection,empty extLst",
            "ok views=[pane=false selections=1 pivots=0 extensions=0] extensions=0",
        ),
        ("pivot,pane", "Invalid(\"pane is out of schema order\")"),
        (
            "pivot,open pane",
            "Invalid(\"pane is out of schema order\")",
        ),
        (
            "pivot,selection",
            "Invalid(\"selection is out of schema order\")",
        ),
        (
            "pivot,open selection",
            "Invalid(\"selection is out of schema order\")",
        ),
        (
            "pivot,pivot",
            "ok views=[pane=false selections=0 pivots=2 extensions=0] extensions=0",
        ),
        (
            "pivot,empty pivot",
            "Invalid(\"pivotSelection requires one pivotArea\")",
        ),
        (
            "pivot,extLst",
            "ok views=[pane=false selections=0 pivots=1 extensions=1] extensions=0",
        ),
        (
            "pivot,empty extLst",
            "ok views=[pane=false selections=0 pivots=1 extensions=0] extensions=0",
        ),
        (
            "empty pivot,pane",
            "Invalid(\"pivotSelection requires one pivotArea\")",
        ),
        (
            "empty pivot,open pane",
            "Invalid(\"pivotSelection requires one pivotArea\")",
        ),
        (
            "empty pivot,selection",
            "Invalid(\"pivotSelection requires one pivotArea\")",
        ),
        (
            "empty pivot,open selection",
            "Invalid(\"pivotSelection requires one pivotArea\")",
        ),
        (
            "empty pivot,pivot",
            "Invalid(\"pivotSelection requires one pivotArea\")",
        ),
        (
            "empty pivot,empty pivot",
            "Invalid(\"pivotSelection requires one pivotArea\")",
        ),
        (
            "empty pivot,extLst",
            "Invalid(\"pivotSelection requires one pivotArea\")",
        ),
        (
            "empty pivot,empty extLst",
            "Invalid(\"pivotSelection requires one pivotArea\")",
        ),
        ("extLst,pane", "Invalid(\"pane is out of schema order\")"),
        (
            "extLst,open pane",
            "Invalid(\"pane is out of schema order\")",
        ),
        (
            "extLst,selection",
            "Invalid(\"selection is out of schema order\")",
        ),
        (
            "extLst,open selection",
            "Invalid(\"selection is out of schema order\")",
        ),
        (
            "extLst,pivot",
            "Invalid(\"pivotSelection is out of schema order\")",
        ),
        (
            "extLst,empty pivot",
            "Invalid(\"pivotSelection requires one pivotArea\")",
        ),
        ("extLst,extLst", "Invalid(\"duplicate sheetView extLst\")"),
        (
            "extLst,empty extLst",
            "Invalid(\"duplicate sheetView extLst\")",
        ),
        (
            "empty extLst,pane",
            "Invalid(\"pane is out of schema order\")",
        ),
        (
            "empty extLst,open pane",
            "Invalid(\"pane is out of schema order\")",
        ),
        (
            "empty extLst,selection",
            "Invalid(\"selection is out of schema order\")",
        ),
        (
            "empty extLst,open selection",
            "Invalid(\"selection is out of schema order\")",
        ),
        (
            "empty extLst,pivot",
            "Invalid(\"pivotSelection is out of schema order\")",
        ),
        (
            "empty extLst,empty pivot",
            "Invalid(\"pivotSelection requires one pivotArea\")",
        ),
        (
            "empty extLst,extLst",
            "Invalid(\"duplicate sheetView extLst\")",
        ),
        (
            "empty extLst,empty extLst",
            "Invalid(\"duplicate sheetView extLst\")",
        ),
        (
            "pane,selection,pivot,extLst",
            "ok views=[pane=true selections=1 pivots=1 extensions=1] extensions=0",
        ),
        (
            "pane,selection,selection,pivot,pivot,empty extLst",
            "ok views=[pane=true selections=2 pivots=2 extensions=0] extensions=0",
        ),
        (
            "selection,pivot,selection",
            "Invalid(\"selection is out of schema order\")",
        ),
        (
            "pivot,extLst,pivot",
            "Invalid(\"pivotSelection is out of schema order\")",
        ),
        (
            "unknown",
            "Invalid(\"unexpected worksheet-view element 'unknown'\")",
        ),
        (
            "selection,unknown",
            "Invalid(\"unexpected worksheet-view element 'unknown'\")",
        ),
        (
            "extLst,unknown",
            "Invalid(\"unexpected worksheet-view element 'unknown'\")",
        ),
        (
            "foreign pane",
            "Invalid(\"unexpected worksheet-view element 'pane'\")",
        ),
        (
            "self-closing view, then view with pane",
            "ok views=[pane=false selections=0 pivots=0 extensions=0][pane=true selections=0 pivots=0 extensions=0] extensions=0",
        ),
        (
            "self-closing view, then view with selection",
            "ok views=[pane=false selections=0 pivots=0 extensions=0][pane=false selections=1 pivots=0 extensions=0] extensions=0",
        ),
        (
            "self-closing view, then view with pivot",
            "ok views=[pane=false selections=0 pivots=0 extensions=0][pane=false selections=0 pivots=1 extensions=0] extensions=0",
        ),
        (
            "self-closing view, then view with extLst",
            "ok views=[pane=false selections=0 pivots=0 extensions=0][pane=false selections=0 pivots=0 extensions=1] extensions=0",
        ),
        (
            "view with selection, then view with pane",
            "ok views=[pane=false selections=1 pivots=0 extensions=0][pane=true selections=0 pivots=0 extensions=0] extensions=0",
        ),
        (
            "view with selection, then view with selection",
            "ok views=[pane=false selections=1 pivots=0 extensions=0][pane=false selections=1 pivots=0 extensions=0] extensions=0",
        ),
        (
            "view with selection, then view with pivot",
            "ok views=[pane=false selections=1 pivots=0 extensions=0][pane=false selections=0 pivots=1 extensions=0] extensions=0",
        ),
        (
            "view with selection, then view with extLst",
            "ok views=[pane=false selections=1 pivots=0 extensions=0][pane=false selections=0 pivots=0 extensions=1] extensions=0",
        ),
        (
            "view with pivot, then view with pane",
            "ok views=[pane=false selections=0 pivots=1 extensions=0][pane=true selections=0 pivots=0 extensions=0] extensions=0",
        ),
        (
            "view with pivot, then view with selection",
            "ok views=[pane=false selections=0 pivots=1 extensions=0][pane=false selections=1 pivots=0 extensions=0] extensions=0",
        ),
        (
            "view with pivot, then view with pivot",
            "ok views=[pane=false selections=0 pivots=1 extensions=0][pane=false selections=0 pivots=1 extensions=0] extensions=0",
        ),
        (
            "view with pivot, then view with extLst",
            "ok views=[pane=false selections=0 pivots=1 extensions=0][pane=false selections=0 pivots=0 extensions=1] extensions=0",
        ),
        (
            "view with extLst, then view with pane",
            "ok views=[pane=false selections=0 pivots=0 extensions=1][pane=true selections=0 pivots=0 extensions=0] extensions=0",
        ),
        (
            "view with extLst, then view with selection",
            "ok views=[pane=false selections=0 pivots=0 extensions=1][pane=false selections=1 pivots=0 extensions=0] extensions=0",
        ),
        (
            "view with extLst, then view with pivot",
            "ok views=[pane=false selections=0 pivots=0 extensions=1][pane=false selections=0 pivots=1 extensions=0] extensions=0",
        ),
        (
            "view with extLst, then view with extLst",
            "ok views=[pane=false selections=0 pivots=0 extensions=1][pane=false selections=0 pivots=0 extensions=1] extensions=0",
        ),
        (
            "view with empty extLst, then view with pane",
            "ok views=[pane=false selections=0 pivots=0 extensions=0][pane=true selections=0 pivots=0 extensions=0] extensions=0",
        ),
        (
            "view with empty extLst, then view with selection",
            "ok views=[pane=false selections=0 pivots=0 extensions=0][pane=false selections=1 pivots=0 extensions=0] extensions=0",
        ),
        (
            "view with empty extLst, then view with pivot",
            "ok views=[pane=false selections=0 pivots=0 extensions=0][pane=false selections=0 pivots=1 extensions=0] extensions=0",
        ),
        (
            "view with empty extLst, then view with extLst",
            "ok views=[pane=false selections=0 pivots=0 extensions=0][pane=false selections=0 pivots=0 extensions=1] extensions=0",
        ),
    ];
    assert_rows(&actual, expected);
}
