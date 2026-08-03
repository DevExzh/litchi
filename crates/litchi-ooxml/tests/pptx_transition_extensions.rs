use litchi_ooxml::pptx::Package;
use litchi_ooxml::{OoxmlError, PackURI};
use litchi_pptx::transition::{Kind, Ms, Ripple, Speed, Transition};
use tempfile::NamedTempFile;

const LOCAL_P14_RIPPLE: &[u8] =
    include_bytes!("../../../test-data/ooxml/pptx/transitions/p14_ripple.xml");

#[test]
fn slide_transition_uses_local_powerpoint_2010_choice() {
    let package = package_with_transition_fragment(std::str::from_utf8(LOCAL_P14_RIPPLE).unwrap());

    let transition = first_transition(&package);
    assert_eq!(transition.speed(), Speed::Slow);
    assert_eq!(transition.duration().map(Ms::get), Some(1500));
    assert!(!transition.click());
    assert_eq!(transition.after().map(Ms::get), Some(4250));
    assert_eq!(transition.kind(), &Kind::Ripple(Ripple::LeftDown));
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
        Err(OoxmlError::Pptx(litchi_pptx::Error::Invalid(message)))
            if message.contains("ripple transition direction")
    ));
}

#[test]
fn writer_round_trips_powerpoint_2010_ripple_with_fade_fallback() {
    let expected = Transition::new(Kind::Ripple(Ripple::LeftDown))
        .with_speed(Speed::Slow)
        .with_duration(Ms::new(1500).unwrap())
        .with_click(false)
        .with_after(Ms::new(4250).unwrap());
    let output = NamedTempFile::with_suffix(".pptx").unwrap();
    let mut package = Package::new().unwrap();
    package
        .presentation_mut()
        .unwrap()
        .add_slide()
        .unwrap()
        .set_transition(expected.clone());
    package.save(output.path()).unwrap();

    let package = Package::open(output.path()).unwrap();
    assert!(first_transition(&package).same_semantics(&expected));

    let slide_name = PackURI::new("/ppt/slides/slide1.xml").unwrap();
    let slide = package.opc().unwrap().get_part(&slide_name).unwrap();
    let xml = std::str::from_utf8(slide.blob()).unwrap();
    assert!(xml.contains("<mc:AlternateContent"));
    assert!(xml.contains(r#"<p14:ripple dir="ld"/>"#));
    assert!(xml.contains("<p:fade/>"));
}

#[test]
fn writer_round_trips_custom_duration_through_compatibility_markup() {
    let expected = Transition::new(Kind::Fade { black: None })
        .with_speed(Speed::Fast)
        .with_duration(Ms::new(750).unwrap())
        .with_click(false)
        .with_after(Ms::new(1250).unwrap());
    let output = NamedTempFile::with_suffix(".pptx").unwrap();
    let mut package = Package::new().unwrap();
    package
        .presentation_mut()
        .unwrap()
        .add_slide()
        .unwrap()
        .set_transition(expected.clone());
    package.save(output.path()).unwrap();

    let package = Package::open(output.path()).unwrap();
    assert!(first_transition(&package).same_semantics(&expected));

    let slide_name = PackURI::new("/ppt/slides/slide1.xml").unwrap();
    let slide = package.opc().unwrap().get_part(&slide_name).unwrap();
    let xml = std::str::from_utf8(slide.blob()).unwrap();
    assert!(xml.contains(r#"p14:dur="750""#));
    assert!(xml.contains("<mc:Fallback>"));
}

fn first_transition(package: &Package) -> Transition {
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
    package
        .edit_opc(|opc| {
            let slide = opc.get_part_mut(&slide_name)?;
            let xml = std::str::from_utf8(slide.blob()).unwrap();
            let updated = xml.replacen("</p:sld>", &format!("{fragment}</p:sld>"), 1);
            assert_ne!(updated, xml);
            slide.set_blob(updated.into_bytes());
            Ok(())
        })
        .unwrap();
    package
}
