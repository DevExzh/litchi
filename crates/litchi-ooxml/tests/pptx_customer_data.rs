use litchi_ooxml::PackURI;
use litchi_ooxml::pptx::Package;
use tempfile::NamedTempFile;

const PRESENTATION_XML: &[u8] =
    include_bytes!("../../../test-data/ooxml/pptx/customer-data/presentation.xml");

#[test]
fn presentation_customer_data_relationships_are_exposed() {
    let package = package_with_presentation_xml();
    let customer_data = package
        .presentation()
        .unwrap()
        .customer_data()
        .unwrap()
        .unwrap();

    assert_eq!(
        customer_data.custom_data_relationship_ids(),
        ["rIdCustomerDataOne", "rIdCustomerDataTwo"]
    );
    assert_eq!(
        customer_data.tags_relationship_id(),
        Some("rIdCustomerDataTags")
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
