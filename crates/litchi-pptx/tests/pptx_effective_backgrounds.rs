use litchi_opc::{OpcPackage, PackURI};
use litchi_pptx::Package;
use litchi_pptx::SlideBackground;

const SOLID_BACKGROUND: &str = include_str!("../../../test-data/ooxml/pptx/backgrounds/solid.xml");

#[test]
fn effective_slide_background_prefers_the_slide_over_layout_and_master() {
    let package = package_with_backgrounds(&[
        ("/ppt/slides/slide1.xml", "AABBCC"),
        ("/ppt/slideLayouts/slideLayout1.xml", "445566"),
        ("/ppt/slideMasters/slideMaster1.xml", "112233"),
    ]);

    assert_solid(
        effective_background(&package, "/ppt/slides/slide1.xml").unwrap(),
        "AABBCC",
    );
}

#[test]
fn effective_slide_background_prefers_the_layout_over_the_master() {
    let package = package_with_backgrounds(&[
        ("/ppt/slideLayouts/slideLayout1.xml", "445566"),
        ("/ppt/slideMasters/slideMaster1.xml", "112233"),
    ]);

    assert_solid(
        effective_background(&package, "/ppt/slideLayouts/slideLayout1.xml").unwrap(),
        "445566",
    );
    assert_solid(
        effective_background(&package, "/ppt/slides/slide1.xml").unwrap(),
        "445566",
    );
}

#[test]
fn effective_slide_background_falls_back_to_the_master() {
    let package = package_with_backgrounds(&[("/ppt/slideMasters/slideMaster1.xml", "112233")]);

    assert_solid(
        effective_background(&package, "/ppt/slideLayouts/slideLayout1.xml").unwrap(),
        "112233",
    );
    assert_solid(
        effective_background(&package, "/ppt/slides/slide1.xml").unwrap(),
        "112233",
    );
}

fn assert_solid(background: SlideBackground, expected_color: &str) {
    assert_eq!(
        background,
        SlideBackground::Solid {
            color: expected_color.to_string(),
        }
    );
}

fn effective_background(package: &OpcPackage, part_name: &str) -> Option<SlideBackground> {
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
    candidates.into_iter().flatten().find_map(|name| {
        let part = package.get_part(&PackURI::new(name).unwrap()).unwrap();
        SlideBackground::from_xml(part.blob()).unwrap()
    })
}

fn package_with_backgrounds(backgrounds: &[(&str, &str)]) -> OpcPackage {
    let mut package = Package::new().unwrap();
    package.presentation_mut().unwrap().add_slide().unwrap();
    let bytes = package.to_bytes().unwrap();
    let mut package = OpcPackage::from_vec(bytes).unwrap();

    for (part_name, color) in backgrounds {
        let part_name = PackURI::new(*part_name).unwrap();
        let part = package.get_part_mut(&part_name).unwrap();
        let xml = std::str::from_utf8(part.blob()).unwrap();
        let background = SOLID_BACKGROUND.replacen("112233", color, 1);
        let updated = xml.replacen("<p:spTree>", &format!("{background}<p:spTree>"), 1);
        assert_ne!(updated, xml);
        part.set_blob(updated.into_bytes());
    }
    package
}
