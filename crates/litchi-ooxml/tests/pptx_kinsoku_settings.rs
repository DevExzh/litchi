use litchi_ooxml::PackURI;
use litchi_ooxml::pptx::Package;
use tempfile::NamedTempFile;

const PRESENTATION_XML: &[u8] =
    include_bytes!("../../../test-data/ooxml/pptx/kinsoku/presentation.xml");

#[test]
fn presentation_kinsoku_settings_are_exposed() {
    let package = package_with_presentation_xml();
    let settings = package
        .presentation()
        .unwrap()
        .kinsoku_settings()
        .unwrap()
        .unwrap();

    assert_eq!(settings.language(), Some("ja-jp"));
    assert_eq!(settings.invalid_start_characters(), "、。）］");
    assert_eq!(settings.invalid_end_characters(), "（［");
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
