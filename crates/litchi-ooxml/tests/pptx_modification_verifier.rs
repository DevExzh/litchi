use litchi_ooxml::PackURI;
use litchi_ooxml::pptx::Package;
use tempfile::NamedTempFile;

const PRESENTATION_XML: &[u8] =
    include_bytes!("../../../test-data/ooxml/pptx/modification-verifier/presentation.xml");

#[test]
fn presentation_modification_verifier_is_exposed() {
    let package = package_with_presentation_xml();
    let verifier = package
        .presentation()
        .unwrap()
        .modification_verifier()
        .unwrap()
        .unwrap();

    assert_eq!(verifier.cryptographic_provider_type(), Some("rsaAES"));
    assert_eq!(verifier.cryptographic_algorithm_class(), Some("hash"));
    assert_eq!(verifier.cryptographic_algorithm_type(), Some("typeAny"));
    assert_eq!(verifier.cryptographic_provider(), Some("Microsoft"));
    assert_eq!(verifier.algorithm_name(), None);
    assert_eq!(verifier.crypt_algorithm_sid(), Some(14));
    assert_eq!(verifier.hash_data(), Some("legacy-hash"));
    assert_eq!(verifier.hash_value(), None);
    assert_eq!(verifier.salt_data(), Some("legacy-salt"));
    assert_eq!(verifier.salt_value(), None);
    assert_eq!(verifier.spin_count(), Some(100_000));
    assert_eq!(verifier.spin_value(), None);
}

fn package_with_presentation_xml() -> Package {
    let output = NamedTempFile::with_suffix(".pptx").unwrap();
    let mut package = Package::new().unwrap();
    package.save(output.path()).unwrap();

    let mut package = Package::open(output.path()).unwrap();
    let part_name = PackURI::new("/ppt/presentation.xml").unwrap();
    package
        .opc_package_mut()
        .get_part_mut(&part_name)
        .unwrap()
        .set_blob(PRESENTATION_XML.to_vec());
    package
}
