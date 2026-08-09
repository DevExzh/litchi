#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

use litchi_pptx::Error;
use litchi_pptx::transition::{Kind, Ms, Ripple, Speed, Transition, read, write};

const LOCAL_P14_RIPPLE: &[u8] =
    include_bytes!("../../../test-data/ooxml/pptx/transitions/p14_ripple.xml");
const PRESENTATIONML_NAMESPACE: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";

#[test]
fn transition_reader_uses_the_local_powerpoint_2010_choice() {
    let fragment = std::str::from_utf8(LOCAL_P14_RIPPLE).unwrap();
    let transition = read(slide_xml(fragment).as_bytes()).unwrap().unwrap();
    assert_eq!(transition.speed(), Speed::Slow);
    assert_eq!(transition.duration().map(Ms::get), Some(1500));
    assert!(!transition.click());
    assert_eq!(transition.after().map(Ms::get), Some(4250));
    assert_eq!(transition.kind(), &Kind::Ripple(Ripple::LeftDown));
}

#[test]
fn transition_reader_rejects_invalid_powerpoint_2010_ripple_direction() {
    let invalid = std::str::from_utf8(LOCAL_P14_RIPPLE).unwrap().replacen(
        "dir=\"ld\"",
        "dir=\"not-a-direction\"",
        1,
    );

    assert!(matches!(
        read(slide_xml(&invalid).as_bytes()),
        Err(Error::Invalid(message)) if message.contains("ripple transition direction")
    ));
}

#[test]
fn transition_writer_round_trips_powerpoint_2010_ripple_with_fade_fallback() {
    let expected = Transition::new(Kind::Ripple(Ripple::LeftDown))
        .with_speed(Speed::Slow)
        .with_duration(Ms::new(1500).unwrap())
        .with_click(false)
        .with_after(Ms::new(4250).unwrap());
    let xml = write(&expected).unwrap();
    assert!(xml.contains("<mc:AlternateContent"));
    assert!(xml.contains(r#"<p14:ripple dir="ld"/>"#));
    assert!(xml.contains("<p:fade/>"));
    assert!(
        read(slide_xml(&xml).as_bytes())
            .unwrap()
            .unwrap()
            .same_semantics(&expected)
    );
}

#[test]
fn transition_writer_round_trips_custom_duration_through_compatibility_markup() {
    let expected = Transition::new(Kind::Fade { black: None })
        .with_speed(Speed::Fast)
        .with_duration(Ms::new(750).unwrap())
        .with_click(false)
        .with_after(Ms::new(1250).unwrap());
    let xml = write(&expected).unwrap();
    assert!(xml.contains(r#"p14:dur="750""#));
    assert!(xml.contains("<mc:Fallback>"));
    assert!(
        read(slide_xml(&xml).as_bytes())
            .unwrap()
            .unwrap()
            .same_semantics(&expected)
    );
}

fn slide_xml(fragment: &str) -> String {
    format!(r#"<p:sld xmlns:p="{PRESENTATIONML_NAMESPACE}">{fragment}</p:sld>"#)
}
