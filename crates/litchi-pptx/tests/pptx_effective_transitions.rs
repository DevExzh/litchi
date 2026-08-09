#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

use litchi_opc::{OpcPackage, PackURI};
use litchi_pptx::Package;
use litchi_pptx::transition::{Kind, Ms, Origin, Ripple, Transition, read};

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

    assert_cover(
        &effective_transition(&package, "/ppt/slides/slide1.xml")
            .unwrap()
            .unwrap(),
    );
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

    assert_ripple(
        &effective_transition(&package, "/ppt/slideLayouts/slideLayout1.xml")
            .unwrap()
            .unwrap(),
    );
    assert_ripple(
        &effective_transition(&package, "/ppt/slides/slide1.xml")
            .unwrap()
            .unwrap(),
    );
}

#[test]
fn effective_slide_transition_falls_back_to_the_master() {
    let package = package_with_transition_fragments(&[(
        "/ppt/slideMasters/slideMaster1.xml",
        "</p:sldMaster>",
        transition_fragment(STANDARD_COVER),
    )]);

    assert_cover(
        &effective_transition(&package, "/ppt/slideLayouts/slideLayout1.xml")
            .unwrap()
            .unwrap(),
    );
    assert_cover(
        &effective_transition(&package, "/ppt/slides/slide1.xml")
            .unwrap()
            .unwrap(),
    );
}

fn assert_cover(transition: &Transition) {
    assert_eq!(transition.kind(), &Kind::Cover(Origin::RightDown));
}

fn assert_ripple(transition: &Transition) {
    assert_eq!(transition.duration().map(Ms::get), Some(1500));
    assert_eq!(transition.kind(), &Kind::Ripple(Ripple::LeftDown));
}

fn effective_transition(
    package: &OpcPackage,
    part_name: &str,
) -> litchi_pptx::Result<Option<Transition>> {
    let candidates = match part_name {
        "/ppt/slides/slide1.xml" => [
            Some(part_name),
            Some("/ppt/slideLayouts/slideLayout1.xml"),
            Some("/ppt/slideMasters/slideMaster1.xml"),
        ],
        "/ppt/slideLayouts/slideLayout1.xml" => [
            Some(part_name),
            Some("/ppt/slideMasters/slideMaster1.xml"),
            None,
        ],
        _ => [Some(part_name), None, None],
    };
    for name in candidates.into_iter().flatten() {
        let part = package.get_part(&PackURI::new(name).unwrap())?;
        if let Some(transition) = read(part.blob())? {
            return Ok(Some(transition));
        }
    }
    Ok(None)
}

fn package_with_transition_fragments(fragments: &[(&str, &str, &str)]) -> OpcPackage {
    let mut package = Package::new().unwrap();
    package.presentation_mut().unwrap().add_slide().unwrap();
    let bytes = package.to_bytes().unwrap();
    let mut package = OpcPackage::from_vec(bytes).unwrap();
    for (part_name, end_tag, fragment) in fragments {
        let part_name = PackURI::new(*part_name).unwrap();
        let part = package.get_part_mut(&part_name).unwrap();
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
