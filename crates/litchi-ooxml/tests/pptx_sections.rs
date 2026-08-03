use litchi_ooxml::PackURI;
use litchi_ooxml::pptx::Package;
use tempfile::NamedTempFile;

const PRESENTATION_XML: &[u8] =
    include_bytes!("../../../test-data/ooxml/pptx/sections/presentation.xml");

#[test]
fn presentation_sections_include_slide_membership() {
    let package = package_with_presentation_xml();
    let presentation = package.presentation().unwrap();

    assert_eq!(presentation.slide_ids().unwrap(), [256, 257, 258]);

    let sections = presentation.sections().unwrap();
    assert_eq!(sections.sections().len(), 2);
    assert_eq!(sections.sections()[0].name.as_deref(), Some("Opening"));
    assert_eq!(
        sections.sections()[0].id.as_deref(),
        Some("{11111111-1111-1111-1111-111111111111}")
    );
    assert_eq!(sections.sections()[0].slide_ids, [256, 258]);
    assert_eq!(sections.sections()[1].name.as_deref(), Some("Recap"));
    assert_eq!(sections.sections()[1].slide_ids, [257]);

    assert_eq!(
        presentation.get_sections().unwrap(),
        [
            ("Opening".to_string(), vec![0, 2]),
            ("Recap".to_string(), vec![1]),
        ]
    );
}

fn package_with_presentation_xml() -> Package {
    let output = NamedTempFile::with_suffix(".pptx").unwrap();
    let mut package = Package::new().unwrap();
    package.save(output.path()).unwrap();

    let mut package = Package::open(output.path()).unwrap();
    let part_name = PackURI::new("/ppt/presentation.xml").unwrap();
    package
        .edit_opc(|opc| {
            opc.get_part_mut(&part_name)?
                .set_blob(PRESENTATION_XML.to_vec());
            Ok(())
        })
        .unwrap();
    package
}
