#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

use litchi_opc::{OpcPackage, PackURI};
use litchi_pptx::presentation_properties::metadata::events::{Draft, EXTENSION_URI, Kind, Trigger};
use litchi_pptx::time::Offset;
use litchi_pptx::{Error, Package, Result};

const LOCAL_SHOW_EVENTS: &[u8] =
    include_bytes!("../../../test-data/ooxml/pptx/show-events/basic_show_events.xml");

#[test]
fn event_owner_stores_the_local_show_event_inventory() {
    assert_eq!(
        std::str::from_utf8(LOCAL_SHOW_EVENTS)
            .unwrap()
            .matches("<p14:")
            .count(),
        8
    );
    let mut package = package_with_slide();
    let slide_name = PackURI::new("/ppt/slides/slide1.xml").unwrap();
    let events = event_drafts();

    store(&mut package, &slide_name, &events).unwrap();

    let xml = std::str::from_utf8(package.get_part(&slide_name).unwrap().blob()).unwrap();
    assert_eq!(xml.matches("<p14:").count(), 8); // showEvtLst plus seven events
    assert!(xml.contains(r#"type="onClick" time="6950" objId="6""#));
    assert!(xml.contains(r#"<p14:seekEvt time="38839" objId="4" seek="10379"/>"#));
    assert!(xml.contains(r#"<p14:nullEvt time="50000" objId="4"/>"#));
    assert!(matches!(events[0].kind(), Kind::Trigger(Trigger::OnClick)));
    assert_eq!(events[0].time(), &Offset::ms(6950));
    assert_eq!(events[3].seek_time(), Some(&Offset::ms(10379)));
}

#[test]
fn event_owner_rejects_malformed_existing_show_event_times() {
    let malformed = format!(
        r#"<p:extLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main"><p:ext uri="{EXTENSION_URI}"><p14:showEvtLst><p14:playEvt time="bad-time" objId="4"/></p14:showEvtLst></p:ext></p:extLst>"#
    );
    let mut package = package_with_show_event_extension(&malformed);
    let slide_name = PackURI::new("/ppt/slides/slide1.xml").unwrap();

    assert!(matches!(
        store(&mut package, &slide_name, &[Draft::play(Offset::ms(1), 4)]),
        Err(Error::Invalid(message)) if message.contains("universal time offset")
    ));
}

#[test]
fn event_owner_ignores_events_under_other_extensions() {
    let ignored = r#"<p:extLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main"><p:ext uri="urn:litchi:not-show-events"><p14:showEvtLst><p14:playEvt time="1" objId="4"/></p14:showEvtLst></p:ext></p:extLst>"#;
    let mut package = package_with_show_event_extension(ignored);
    let slide_name = PackURI::new("/ppt/slides/slide1.xml").unwrap();

    store(&mut package, &slide_name, &[Draft::play(Offset::ms(1), 4)]).unwrap();
    let xml = std::str::from_utf8(package.get_part(&slide_name).unwrap().blob()).unwrap();
    assert!(xml.contains("urn:litchi:not-show-events"));
    assert!(xml.contains(EXTENSION_URI));
}

/// Store `events` on one slide of a raw graph through the package family.
fn store(opc: &mut OpcPackage, slide: &PackURI, events: &[Draft]) -> Result<()> {
    let package = Package::from_opc_package(opc.clone())?;
    let mut edit = package.edit_show_events()?;
    // Every fixture here has one slide.
    assert_eq!(slide.as_str(), "/ppt/slides/slide1.xml");
    edit.set(0_usize, events)?;
    *opc = edit.commit()?.into_package().opc()?.clone();
    Ok(())
}

fn event_drafts() -> Vec<Draft> {
    vec![
        Draft::trigger(Trigger::OnClick, Offset::ms(6950), 6),
        Draft::play(Offset::ms(12722), 4),
        Draft::pause(Offset::ms(38839), 4),
        Draft::seek(Offset::ms(38839), 4, Offset::ms(10379)),
        Draft::resume(Offset::ms(38859), 4),
        Draft::stop(Offset::ms(49628), 4),
        Draft::null(Offset::ms(50000), 4),
    ]
}

fn package_with_slide() -> OpcPackage {
    let mut builder = litchi_pptx::builder::Builder::new().unwrap();
    builder.presentation().unwrap().add_slide().unwrap();
    let authored = builder.build().unwrap();
    OpcPackage::from_bytes(
        &authored.to_bytes().unwrap(),
        litchi_opc::SourceLimits::default(),
    )
    .unwrap()
}

fn package_with_show_event_extension(extension: &str) -> OpcPackage {
    let mut builder = litchi_pptx::builder::Builder::new().unwrap();
    builder.presentation().unwrap().add_slide().unwrap();
    let authored = builder.build().unwrap();
    let bytes = authored.to_bytes().unwrap();
    let mut package = OpcPackage::from_bytes(&bytes, litchi_opc::SourceLimits::default()).unwrap();
    let slide_name = PackURI::new("/ppt/slides/slide1.xml").unwrap();
    let slide = package.get_part_mut(&slide_name).unwrap();
    let xml = std::str::from_utf8(slide.blob()).unwrap();
    let updated = xml.replacen("</p:sld>", &format!("{extension}</p:sld>"), 1);
    assert_ne!(updated, xml);
    slide.set_blob(updated.into_bytes());
    package
}
