use litchi_opc::{OpcPackage, PackURI};
use litchi_pptx::Package;

const PRESENTATION_XML: &[u8] =
    include_bytes!("../../../test-data/ooxml/pptx/photo-album/presentation.xml");

#[test]
fn presentation_photo_album_metadata_is_preserved_as_an_opaque_extension() {
    let package = package_with_presentation_xml();
    let presentation = package
        .opc()
        .unwrap()
        .get_part(&PackURI::new("/ppt/presentation.xml").unwrap())
        .unwrap();

    // Photo-album is not currently a modeled standalone PresentationML owner.
    // The lossless OPC facade still preserves and exposes its exact source
    // bytes, so migration does not discard the producer's metadata.
    let xml = presentation.blob();
    assert!(xml.windows(b"<p:photoAlbum".len()).any(|window| window == b"<p:photoAlbum"));
    assert!(xml.windows(b"bw=\"1\"".len()).any(|window| window == b"bw=\"1\""));
    assert!(xml
        .windows(b"showCaptions=\"true\"".len())
        .any(|window| window == b"showCaptions=\"true\""));
    assert!(xml
        .windows(b"layout=\"2picTitle\"".len())
        .any(|window| window == b"layout=\"2picTitle\""));
    assert!(xml
        .windows(b"frame=\"frameStyle5\"".len())
        .any(|window| window == b"frame=\"frameStyle5\""));
}

fn package_with_presentation_xml() -> Package {
    let mut package = Package::new().unwrap();
    let bytes = package.to_bytes().unwrap();
    let mut opc = OpcPackage::from_bytes(&bytes).unwrap();
    opc.get_part_mut(&PackURI::new("/ppt/presentation.xml").unwrap())
        .unwrap()
        .set_blob(PRESENTATION_XML.to_vec());
    Package::from_opc_package(opc).unwrap()
}
