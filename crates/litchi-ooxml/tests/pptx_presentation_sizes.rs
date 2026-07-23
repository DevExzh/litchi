use litchi_ooxml::PackURI;
use litchi_ooxml::pptx::Package;
use tempfile::NamedTempFile;

const DEFINED_XML: &[u8] =
    include_bytes!("../../../test-data/ooxml/pptx/presentation-sizes/defined.xml");
const ABSENT_XML: &[u8] =
    include_bytes!("../../../test-data/ooxml/pptx/presentation-sizes/absent.xml");

#[test]
fn presentation_surface_sizes_are_exposed() {
    let package = package_with_presentation_xml(DEFINED_XML);
    let presentation = package.presentation().unwrap();

    assert_eq!(
        presentation.slide_size().unwrap(),
        Some((12_192_000, 6_858_000))
    );
    let slide_size = presentation.slide_size_metadata().unwrap().unwrap();
    assert_eq!(slide_size.width(), 12_192_000);
    assert_eq!(slide_size.height(), 6_858_000);
    assert_eq!(slide_size.size_type(), Some("screen16x9"));

    let notes_size = presentation.notes_size().unwrap().unwrap();
    assert_eq!(notes_size.width(), 6_858_000);
    assert_eq!(notes_size.height(), 9_144_000);
}

#[test]
fn absent_presentation_surface_sizes_return_none() {
    let package = package_with_presentation_xml(ABSENT_XML);
    let presentation = package.presentation().unwrap();

    assert_eq!(presentation.slide_size().unwrap(), None);
    assert_eq!(presentation.slide_size_metadata().unwrap(), None);
    assert_eq!(presentation.notes_size().unwrap(), None);
}

fn package_with_presentation_xml(xml: &[u8]) -> Package {
    let output = NamedTempFile::with_suffix(".pptx").unwrap();
    let mut package = Package::new().unwrap();
    package.presentation_mut().unwrap().add_slide().unwrap();
    package.save(output.path()).unwrap();

    let mut package = Package::open(output.path()).unwrap();
    let part_name = PackURI::new("/ppt/presentation.xml").unwrap();
    package
        .opc_package_mut()
        .get_part_mut(&part_name)
        .unwrap()
        .set_blob(xml.to_vec());
    package
}
