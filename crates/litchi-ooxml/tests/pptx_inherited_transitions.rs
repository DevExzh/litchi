use litchi_ooxml::PackURI;
use litchi_ooxml::pptx::parts::SlideLayoutPart;
use litchi_ooxml::pptx::{Package, RippleDirection, SlideTransition, TransitionType};
use tempfile::NamedTempFile;

const LOCAL_P14_RIPPLE: &[u8] =
    include_bytes!("../../../test-data/ooxml/pptx/transitions/p14_ripple.xml");

#[test]
fn layout_and_master_read_local_powerpoint_2010_transition_choices() {
    let package =
        package_with_inherited_transition_fragment(std::str::from_utf8(LOCAL_P14_RIPPLE).unwrap());

    let presentation = package.presentation().unwrap();
    let masters = presentation.slide_masters().unwrap();
    assert_ripple(&masters[0].transition().unwrap().unwrap());

    let layout_name = PackURI::new("/ppt/slideLayouts/slideLayout1.xml").unwrap();
    let layout_part =
        SlideLayoutPart::from_part(package.opc_package().get_part(&layout_name).unwrap()).unwrap();
    assert_ripple(&layout_part.transition().unwrap().unwrap());
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

fn package_with_inherited_transition_fragment(fragment: &str) -> Package {
    let output = NamedTempFile::with_suffix(".pptx").unwrap();
    let mut package = Package::new().unwrap();
    package.save(output.path()).unwrap();

    let mut package = Package::open(output.path()).unwrap();
    for (part_name, end_tag) in [
        ("/ppt/slideLayouts/slideLayout1.xml", "</p:sldLayout>"),
        ("/ppt/slideMasters/slideMaster1.xml", "</p:sldMaster>"),
    ] {
        let part_name = PackURI::new(part_name).unwrap();
        let part = package.opc_package_mut().get_part_mut(&part_name).unwrap();
        let xml = std::str::from_utf8(part.blob()).unwrap();
        let updated = xml.replacen(end_tag, &format!("{fragment}{end_tag}"), 1);
        assert_ne!(updated, xml);
        part.set_blob(updated.into_bytes());
    }
    package
}
