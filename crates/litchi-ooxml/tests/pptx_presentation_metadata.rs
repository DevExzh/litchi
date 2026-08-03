use litchi_ooxml::PackURI;
use litchi_ooxml::pptx::{Package, PresentationConformance};
use tempfile::NamedTempFile;

const PRESENTATION_XML: &[u8] =
    include_bytes!("../../../test-data/ooxml/pptx/presentation-metadata/presentation.xml");

#[test]
fn presentation_root_metadata_is_exposed() {
    let package = package_with_presentation_xml();
    let metadata = package.presentation().unwrap().metadata().unwrap();

    assert_eq!(metadata.server_zoom(), 125_000);
    assert_eq!(metadata.first_slide_number(), 7);
    assert!(!metadata.shows_special_placeholders_on_title_slide());
    assert!(metadata.is_right_to_left());
    assert!(metadata.removes_personal_info_on_save());
    assert!(metadata.is_compatibility_mode());
    assert!(!metadata.uses_strict_first_and_last_chars());
    assert!(metadata.embeds_true_type_fonts());
    assert!(metadata.saves_subset_fonts());
    assert!(!metadata.automatically_compresses_pictures());
    assert_eq!(metadata.bookmark_id_seed(), 42);
    assert_eq!(metadata.conformance(), PresentationConformance::Strict);
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
