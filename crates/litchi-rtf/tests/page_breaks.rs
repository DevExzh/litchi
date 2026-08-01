use litchi_rtf::{BodyStoryEvent, CellStoryEvent, RtfDocument, RtfWriter, StoryEvent};

fn write(document: &RtfDocument<'_>) -> Vec<u8> {
    let mut output = Vec::new();
    RtfWriter::new(&mut output)
        .write_document(document)
        .unwrap();
    output
}

fn has_page(events: &[StoryEvent], position: usize) -> bool {
    events
        .iter()
        .any(|event| matches!(event, StoryEvent::PageBreak(page) if page.position == position))
}

#[test]
fn body_page_is_zero_width_ordered_and_round_trips_canonically() {
    let document = RtfDocument::parse(r#"{\rtf1\ansi A\page B}"#).unwrap();
    assert_eq!(document.text(), "AB");
    assert!(matches!(
        document.body_story_events(),
        [BodyStoryEvent::PageBreak(page)] if page.position == 1
    ));

    let output = write(&document);
    let serialized = String::from_utf8(output.clone()).unwrap();
    assert!(serialized.contains("\\page "));
    assert!(!serialized.contains("A\\par B"));
    let reparsed = RtfDocument::parse_bytes(&output).unwrap();
    assert_eq!(reparsed.text(), document.text());
    assert_eq!(reparsed.body_story_events(), document.body_story_events());
}

#[test]
fn visible_story_owners_preserve_page_events() {
    let source = concat!(
        r#"{\rtf1\ansi{\header H\page T}"#,
        r#"A{\footnote N\page O}"#,
        r#"{\field{\*\fldinst TEST}{\fldrslt F\page R}}"#,
        r#"{\shp{\*\shpinst{\sp{\sn shapeType}{\sv 202}}{\shptxt S\page T}}}"#,
        r#"{\*\atnid I}{\*\atnauthor A}\chatn{\*\annotation C\page D}"#,
        r#"{\*\do\dptxbx{\dptxbxtext L\page M}}}"#,
    );
    let document = RtfDocument::parse(source).unwrap();
    assert!(has_page(
        &document.sections()[0].headers_footers[0].story_events,
        1
    ));
    assert!(has_page(&document.notes()[0].story_events, 1));
    assert!(has_page(&document.fields()[0].result_events, 1));
    assert!(has_page(&document.shapes()[0].text_story_events, 1));
    assert!(has_page(&document.annotations()[0].story_events, 1));
    assert!(has_page(&document.legacy_text_boxes()[0].story_events, 1));

    let output = write(&document);
    let reparsed = RtfDocument::parse_bytes(&output).unwrap();
    assert!(has_page(
        &reparsed.sections()[0].headers_footers[0].story_events,
        1
    ));
    assert!(has_page(&reparsed.notes()[0].story_events, 1));
    assert!(has_page(&reparsed.fields()[0].result_events, 1));
    assert!(has_page(&reparsed.shapes()[0].text_story_events, 1));
    assert!(has_page(&reparsed.annotations()[0].story_events, 1));
    assert!(has_page(&reparsed.legacy_text_boxes()[0].story_events, 1));
}

#[test]
fn table_and_nested_table_page_events_round_trip_in_place() {
    let source = r#"{\rtf1\trowd\cellx5000\intbl\itap1 Before\page After \intbl\itap2 Inner\page Tail\nestcell{\*\nesttableprops\itap2\trowd\cellx1000\nestrow}{\nonesttables\par}\intbl\itap1 End\cell\row}"#;
    let document = RtfDocument::parse(source).unwrap();
    let outer = &document.tables()[0].rows()[0].cells()[0];
    assert!(
        outer
            .story_events()
            .iter()
            .any(|event| matches!(event, CellStoryEvent::PageBreak(page) if page.position == 6),)
    );
    let inner = &outer.nested_tables()[0].table.rows()[0].cells()[0];
    assert!(
        inner
            .story_events()
            .iter()
            .any(|event| matches!(event, CellStoryEvent::PageBreak(page) if page.position == 5),)
    );

    let output = write(&document);
    let reparsed = RtfDocument::parse_bytes(&output).unwrap();
    let outer = &reparsed.tables()[0].rows()[0].cells()[0];
    let inner = &outer.nested_tables()[0].table.rows()[0].cells()[0];
    assert!(
        outer
            .story_events()
            .iter()
            .any(|event| matches!(event, CellStoryEvent::PageBreak(_)))
    );
    assert!(
        inner
            .story_events()
            .iter()
            .any(|event| matches!(event, CellStoryEvent::PageBreak(_)))
    );
}

#[test]
fn rejects_parameters_and_non_story_destinations_but_skips_unknown_groups() {
    for source in [
        r#"{\rtf1 A\page1 B}"#,
        r#"{\rtf1{\header A\page-1 B}x}"#,
        r#"{\rtf1{\stylesheet{\s0\page Normal;}}x}"#,
        r#"{\rtf1{\*\defchp\page}x}"#,
    ] {
        assert!(RtfDocument::parse(source).is_err(), "accepted {source}");
    }

    let document = RtfDocument::parse(r#"{\rtf1 A{\*\unknown\page1 ignored}B}"#).unwrap();
    assert_eq!(document.text(), "AB");
    assert!(document.page_breaks().next().is_none());
}
