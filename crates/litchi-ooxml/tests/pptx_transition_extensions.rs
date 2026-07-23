use litchi_ooxml::pptx::{Package, RippleDirection, TransitionSpeed, TransitionType};
use litchi_ooxml::{OoxmlError, PackURI};
use tempfile::NamedTempFile;

const LOCAL_P14_RIPPLE: &[u8] =
    include_bytes!("../../../test-data/ooxml/pptx/transitions/p14_ripple.xml");

#[test]
fn slide_transition_uses_local_powerpoint_2010_choice() {
    let package = package_with_transition_fragment(std::str::from_utf8(LOCAL_P14_RIPPLE).unwrap());

    let transition = first_transition(&package);
    assert_eq!(transition.speed, TransitionSpeed::Slow);
    assert_eq!(transition.duration_ms, Some(1500));
    assert!(!transition.advance_on_click);
    assert_eq!(transition.advance_after_ms, Some(4250));
    assert_eq!(
        transition.transition_type,
        TransitionType::Ripple {
            direction: RippleDirection::LeftDown,
        }
    );
}

#[test]
fn slide_transition_rejects_invalid_powerpoint_2010_ripple_direction() {
    let invalid = std::str::from_utf8(LOCAL_P14_RIPPLE).unwrap().replacen(
        "dir=\"ld\"",
        "dir=\"not-a-direction\"",
        1,
    );
    let package = package_with_transition_fragment(&invalid);
    let presentation = package.presentation().unwrap();
    let slides = presentation.slides().unwrap();

    assert!(matches!(
        slides[0].transition(),
        Err(OoxmlError::InvalidFormat(message)) if message.contains("ripple direction")
    ));
}

fn first_transition(package: &Package) -> litchi_ooxml::pptx::SlideTransition {
    let presentation = package.presentation().unwrap();
    let slides = presentation.slides().unwrap();
    slides[0].transition().unwrap().unwrap()
}

fn package_with_transition_fragment(fragment: &str) -> Package {
    let output = NamedTempFile::with_suffix(".pptx").unwrap();
    let mut package = Package::new().unwrap();
    package.presentation_mut().unwrap().add_slide().unwrap();
    package.save(output.path()).unwrap();

    let mut package = Package::open(output.path()).unwrap();
    let slide_name = PackURI::new("/ppt/slides/slide1.xml").unwrap();
    let slide = package.opc_package_mut().get_part_mut(&slide_name).unwrap();
    let xml = std::str::from_utf8(slide.blob()).unwrap();
    let updated = xml.replacen("</p:sld>", &format!("{fragment}</p:sld>"), 1);
    assert_ne!(updated, xml);
    slide.set_blob(updated.into_bytes());
    package
}
