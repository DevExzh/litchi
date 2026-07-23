use litchi_ooxml::pptx::{
    Package, PptxSlideShowEventKind, PptxSlideShowTrigger, SHOW_EVENT_EXTENSION_URI,
};
use litchi_ooxml::{OoxmlError, PackURI};
use tempfile::NamedTempFile;

const LOCAL_SHOW_EVENTS: &[u8] =
    include_bytes!("../../../test-data/ooxml/pptx/show-events/basic_show_events.xml");

#[test]
fn package_inventory_reports_local_show_events() {
    let package = package_with_local_show_events();

    let events = package.show_events().unwrap();
    assert_eq!(events.len(), 7);

    let trigger = &events[0];
    assert_eq!(trigger.slide_index(), 0);
    assert_eq!(trigger.event_index(), 0);
    assert_eq!(
        trigger.kind(),
        PptxSlideShowEventKind::Trigger(PptxSlideShowTrigger::OnClick)
    );
    assert_eq!(trigger.time(), "6950");
    assert_eq!(trigger.object_id(), 6);
    assert_eq!(trigger.seek_time(), None);

    assert_eq!(events[1].kind(), PptxSlideShowEventKind::Play);
    assert_eq!(events[2].kind(), PptxSlideShowEventKind::Pause);
    assert_eq!(events[3].kind(), PptxSlideShowEventKind::Seek);
    assert_eq!(events[3].time(), "38839");
    assert_eq!(events[3].object_id(), 4);
    assert_eq!(events[3].seek_time(), Some("10379"));
    assert_eq!(events[4].kind(), PptxSlideShowEventKind::Resume);
    assert_eq!(events[5].kind(), PptxSlideShowEventKind::Stop);
    assert_eq!(events[6].kind(), PptxSlideShowEventKind::Null);
    assert_eq!(events[6].time(), "50000ms");

    assert_eq!(
        package.presentation().unwrap().show_events().unwrap(),
        events
    );
}

#[test]
fn package_inventory_rejects_malformed_show_event_times() {
    let malformed = format!(
        r#"<p:extLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main"><p:ext uri="{SHOW_EVENT_EXTENSION_URI}"><p14:showEvtLst><p14:playEvt time="bad-time" objId="4"/></p14:showEvtLst></p:ext></p:extLst>"#
    );
    let package = package_with_show_event_extension(&malformed);

    assert!(matches!(
        package.show_events(),
        Err(OoxmlError::InvalidFormat(message))
            if message.contains("universal time offset")
    ));
}

#[test]
fn package_inventory_ignores_events_under_other_extensions() {
    let ignored = r#"<p:extLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main"><p:ext uri="urn:litchi:not-show-events"><p14:showEvtLst><p14:playEvt time="1" objId="4"/></p14:showEvtLst></p:ext></p:extLst>"#;
    let package = package_with_show_event_extension(ignored);

    assert!(package.show_events().unwrap().is_empty());
}

fn package_with_local_show_events() -> Package {
    package_with_show_event_extension(std::str::from_utf8(LOCAL_SHOW_EVENTS).unwrap())
}

fn package_with_show_event_extension(extension: &str) -> Package {
    let output = NamedTempFile::with_suffix(".pptx").unwrap();
    let mut package = Package::new().unwrap();
    package.presentation_mut().unwrap().add_slide().unwrap();
    package.save(output.path()).unwrap();

    let mut package = Package::open(output.path()).unwrap();
    let slide_name = PackURI::new("/ppt/slides/slide1.xml").unwrap();
    let slide = package.opc_package_mut().get_part_mut(&slide_name).unwrap();
    let xml = std::str::from_utf8(slide.blob()).unwrap();
    let updated = xml.replacen("</p:sld>", &format!("{extension}</p:sld>"), 1);
    assert_ne!(updated, xml);
    slide.set_blob(updated.into_bytes());
    package
}
