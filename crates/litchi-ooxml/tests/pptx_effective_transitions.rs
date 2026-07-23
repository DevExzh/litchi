use litchi_ooxml::PackURI;
use litchi_ooxml::pptx::{
    Package, RippleDirection, SlideTransition, TransitionDirection, TransitionType,
};
use tempfile::NamedTempFile;

const LOCAL_P14_RIPPLE: &str =
    include_str!("../../../test-data/ooxml/pptx/transitions/p14_ripple.xml");
const STANDARD_COVER: &str =
    include_str!("../../../test-data/ooxml/pptx/transitions/standard_cover.xml");

#[test]
fn effective_slide_transition_prefers_the_slide_over_layout_and_master() {
    let package = package_with_transition_fragments(&[
        (
            "/ppt/slides/slide1.xml",
            "</p:sld>",
            transition_fragment(STANDARD_COVER),
        ),
        (
            "/ppt/slideLayouts/slideLayout1.xml",
            "</p:sldLayout>",
            LOCAL_P14_RIPPLE,
        ),
        (
            "/ppt/slideMasters/slideMaster1.xml",
            "</p:sldMaster>",
            LOCAL_P14_RIPPLE,
        ),
    ]);

    let presentation = package.presentation().unwrap();
    let slides = presentation.slides().unwrap();
    assert_cover(&slides[0].effective_transition().unwrap().unwrap());
}

#[test]
fn effective_slide_transition_prefers_the_layout_over_the_master() {
    let package = package_with_transition_fragments(&[
        (
            "/ppt/slideLayouts/slideLayout1.xml",
            "</p:sldLayout>",
            LOCAL_P14_RIPPLE,
        ),
        (
            "/ppt/slideMasters/slideMaster1.xml",
            "</p:sldMaster>",
            transition_fragment(STANDARD_COVER),
        ),
    ]);

    let presentation = package.presentation().unwrap();
    let slides = presentation.slides().unwrap();
    assert_ripple(
        &slides[0]
            .layout()
            .unwrap()
            .effective_transition()
            .unwrap()
            .unwrap(),
    );
    assert_ripple(&slides[0].effective_transition().unwrap().unwrap());
}

#[test]
fn effective_slide_transition_falls_back_to_the_master() {
    let package = package_with_transition_fragments(&[(
        "/ppt/slideMasters/slideMaster1.xml",
        "</p:sldMaster>",
        transition_fragment(STANDARD_COVER),
    )]);

    let presentation = package.presentation().unwrap();
    let slides = presentation.slides().unwrap();
    assert_cover(
        &slides[0]
            .layout()
            .unwrap()
            .effective_transition()
            .unwrap()
            .unwrap(),
    );
    assert_cover(&slides[0].effective_transition().unwrap().unwrap());
}

fn assert_cover(transition: &SlideTransition) {
    assert_eq!(
        transition.transition_type,
        TransitionType::Cover {
            direction: TransitionDirection::RightDown,
        }
    );
}

fn assert_ripple(transition: &SlideTransition) {
    assert_eq!(transition.duration_ms, Some(1500));
    assert_eq!(
        transition.transition_type,
        TransitionType::Ripple {
            direction: RippleDirection::LeftDown,
        }
    );
}

fn package_with_transition_fragments(fragments: &[(&str, &str, &str)]) -> Package {
    let output = NamedTempFile::with_suffix(".pptx").unwrap();
    let mut package = Package::new().unwrap();
    package.presentation_mut().unwrap().add_slide().unwrap();
    package.save(output.path()).unwrap();

    let mut package = Package::open(output.path()).unwrap();
    for (part_name, end_tag, fragment) in fragments {
        let part_name = PackURI::new(*part_name).unwrap();
        let part = package.opc_package_mut().get_part_mut(&part_name).unwrap();
        let xml = std::str::from_utf8(part.blob()).unwrap();
        let updated = xml.replacen(end_tag, &format!("{fragment}{end_tag}"), 1);
        assert_ne!(updated, xml);
        part.set_blob(updated.into_bytes());
    }
    package
}

fn transition_fragment(xml: &str) -> &str {
    xml.strip_prefix("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n")
        .unwrap_or(xml)
}
