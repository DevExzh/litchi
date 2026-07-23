use litchi_ooxml::PackURI;
use litchi_ooxml::pptx::{Package, SlideBackground};
use tempfile::NamedTempFile;

const SOLID_BACKGROUND: &str = include_str!("../../../test-data/ooxml/pptx/backgrounds/solid.xml");

#[test]
fn effective_slide_background_prefers_the_slide_over_layout_and_master() {
    let package = package_with_backgrounds(&[
        ("/ppt/slides/slide1.xml", "AABBCC"),
        ("/ppt/slideLayouts/slideLayout1.xml", "445566"),
        ("/ppt/slideMasters/slideMaster1.xml", "112233"),
    ]);

    let presentation = package.presentation().unwrap();
    let slides = presentation.slides().unwrap();
    assert_solid(slides[0].effective_background().unwrap().unwrap(), "AABBCC");
}

#[test]
fn effective_slide_background_prefers_the_layout_over_the_master() {
    let package = package_with_backgrounds(&[
        ("/ppt/slideLayouts/slideLayout1.xml", "445566"),
        ("/ppt/slideMasters/slideMaster1.xml", "112233"),
    ]);

    let presentation = package.presentation().unwrap();
    let slides = presentation.slides().unwrap();
    assert_solid(
        slides[0]
            .layout()
            .unwrap()
            .effective_background()
            .unwrap()
            .unwrap(),
        "445566",
    );
    assert_solid(slides[0].effective_background().unwrap().unwrap(), "445566");
}

#[test]
fn effective_slide_background_falls_back_to_the_master() {
    let package = package_with_backgrounds(&[("/ppt/slideMasters/slideMaster1.xml", "112233")]);

    let presentation = package.presentation().unwrap();
    let slides = presentation.slides().unwrap();
    assert_solid(
        slides[0]
            .layout()
            .unwrap()
            .effective_background()
            .unwrap()
            .unwrap(),
        "112233",
    );
    assert_solid(slides[0].effective_background().unwrap().unwrap(), "112233");
}

fn assert_solid(background: SlideBackground, expected_color: &str) {
    assert_eq!(
        background,
        SlideBackground::Solid {
            color: expected_color.to_string(),
        }
    );
}

fn package_with_backgrounds(backgrounds: &[(&str, &str)]) -> Package {
    let output = NamedTempFile::with_suffix(".pptx").unwrap();
    let mut package = Package::new().unwrap();
    package.presentation_mut().unwrap().add_slide().unwrap();
    package.save(output.path()).unwrap();

    let mut package = Package::open(output.path()).unwrap();
    for (part_name, color) in backgrounds {
        let part_name = PackURI::new(*part_name).unwrap();
        let part = package.opc_package_mut().get_part_mut(&part_name).unwrap();
        let xml = std::str::from_utf8(part.blob()).unwrap();
        let background = SOLID_BACKGROUND.replacen("112233", color, 1);
        let updated = xml.replacen("<p:spTree>", &format!("{background}<p:spTree>"), 1);
        assert_ne!(updated, xml);
        part.set_blob(updated.into_bytes());
    }
    package
}
