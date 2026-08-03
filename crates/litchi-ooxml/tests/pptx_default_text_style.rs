use litchi_ooxml::PackURI;
use litchi_ooxml::pptx::Package;
use tempfile::NamedTempFile;

const PRESENTATION_XML: &[u8] =
    include_bytes!("../../../test-data/ooxml/pptx/default-text-style/presentation.xml");

#[test]
fn presentation_default_text_style_is_exposed() {
    let package = package_with_presentation_xml();
    let style = package
        .presentation()
        .unwrap()
        .default_text_style()
        .unwrap()
        .unwrap();

    assert!(style.has_default_paragraph_properties());
    assert_eq!(style.levels(), [2, 5]);
    assert!(style.has_level(5));
    assert!(!style.has_level(1));
}

fn package_with_presentation_xml() -> Package {
    let output = NamedTempFile::with_suffix(".pptx").unwrap();
    let mut package = Package::new().unwrap();
    package.save(output.path()).unwrap();

    let mut package = Package::open(output.path()).unwrap();
    let part_name = PackURI::new("/ppt/presentation.xml").unwrap();
    package
        .edit_opc(|opc| {
            opc.get_part_mut(&part_name)
                .unwrap()
                .set_blob(PRESENTATION_XML.to_vec());
            Ok(())
        })
        .unwrap();
    package
}
