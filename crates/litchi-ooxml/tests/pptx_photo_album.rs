use litchi_ooxml::PackURI;
use litchi_ooxml::pptx::{Package, PhotoAlbumFrame, PhotoAlbumLayout};
use tempfile::NamedTempFile;

const PRESENTATION_XML: &[u8] =
    include_bytes!("../../../test-data/ooxml/pptx/photo-album/presentation.xml");

#[test]
fn presentation_photo_album_metadata_is_exposed() {
    let package = package_with_presentation_xml();
    let photo_album = package
        .presentation()
        .unwrap()
        .photo_album()
        .unwrap()
        .unwrap();

    assert!(photo_album.is_black_and_white());
    assert!(photo_album.shows_captions());
    assert_eq!(photo_album.layout(), PhotoAlbumLayout::TwoPicturesWithTitle);
    assert_eq!(photo_album.frame(), PhotoAlbumFrame::CompoundBlack);
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
