//! Image authoring for presentations: `MutablePresentation::insert_image`
//! stores sniffed payloads under `Pictures/` and the frames stay discoverable
//! through the packaged read APIs after save and reopen.

use litchi_odf::{
    MutablePresentation, OdfImageSource, OdfLength, OpenDocumentPackage, OwnedPackage, Presentation,
};

const PNG_PAYLOAD: &[u8] = b"\x89PNG\r\n\x1a\nfake-png-payload";
const GIF_PAYLOAD: &[u8] = b"GIF89afake-gif-payload";

#[test]
fn insert_image_round_trips_discoverable_package_pictures() {
    let mut mutable = MutablePresentation::new();
    mutable.add_slide("First", "Body one").unwrap();
    mutable.add_slide("Second", "Body two").unwrap();

    let first = mutable
        .insert_image(
            0,
            PNG_PAYLOAD,
            &OdfLength::centimeters(2.0),
            &OdfLength::centimeters(3.0),
            &OdfLength::centimeters(10.0),
            &OdfLength::centimeters(5.0),
        )
        .unwrap();
    let second = mutable
        .insert_image(
            1,
            GIF_PAYLOAD,
            &OdfLength::centimeters(1.0),
            &OdfLength::centimeters(1.0),
            &OdfLength::centimeters(4.0),
            &OdfLength::centimeters(4.0),
        )
        .unwrap();
    assert_eq!(first, "Pictures/image1.png");
    assert_eq!(second, "Pictures/image2.gif");

    let bytes = mutable.to_bytes().unwrap();

    // Payloads are stored verbatim with manifest media types.
    let package = OwnedPackage::from_bytes(bytes.clone()).unwrap();
    assert_eq!(package.get_file(&first).unwrap(), PNG_PAYLOAD);
    assert_eq!(package.get_file(&second).unwrap(), GIF_PAYLOAD);
    let manifest = package.package().unwrap();
    assert_eq!(
        manifest.manifest().get_media_type(&first),
        Some("image/png")
    );
    assert_eq!(
        manifest.manifest().get_media_type(&second),
        Some("image/gif")
    );

    // The media-level scan attributes frames to their pages with geometry.
    let generic = OpenDocumentPackage::from_bytes(bytes.clone()).unwrap();
    let scanned = generic.images().unwrap();
    assert_eq!(scanned.len(), 2);
    let first_frame = scanned[0].frame.as_ref().unwrap();
    assert_eq!(first_frame.page_name.as_deref(), Some("page1"));
    assert_eq!(first_frame.name.as_deref(), Some("Image 1"));
    assert_eq!(first_frame.x.as_deref(), Some("2cm"));
    assert_eq!(first_frame.y.as_deref(), Some("3cm"));
    assert_eq!(first_frame.width.as_deref(), Some("10cm"));
    assert_eq!(first_frame.height.as_deref(), Some("5cm"));
    assert_eq!(scanned[1].frame.as_ref().unwrap().page_name.as_deref(), Some("page2"));
    assert!(
        matches!(&scanned[0].source, OdfImageSource::PackagePart { path, .. } if path == &first)
    );
    assert_eq!(
        generic.image_bytes(&scanned[0]).unwrap().as_deref(),
        Some(PNG_PAYLOAD)
    );

    // The slide model exposes the pictures; slide text stays intact.
    let presentation = Presentation::from_bytes(bytes).unwrap();
    let slides = presentation.slides().unwrap();
    assert_eq!(slides.len(), 2);
    assert_eq!(slides[0].text().unwrap(), "Body one");
    assert_eq!(slides[1].text().unwrap(), "Body two");
    let picture = &slides[0].shapes().unwrap()[0];
    assert_eq!(picture.image_href(), Some(first.as_str()));
    assert_eq!(picture.x.as_deref(), Some("2cm"));
    assert_eq!(picture.width.as_deref(), Some("10cm"));
    assert_eq!(slides[1].shapes().unwrap()[0].image_href(), Some(second.as_str()));
}

#[test]
fn insert_image_continues_numbering_across_generations() {
    let mut mutable = MutablePresentation::new();
    mutable.add_slide("Title", "Body").unwrap();
    let first = mutable
        .insert_image(
            0,
            PNG_PAYLOAD,
            &OdfLength::centimeters(1.0),
            &OdfLength::centimeters(1.0),
            &OdfLength::centimeters(1.0),
            &OdfLength::centimeters(1.0),
        )
        .unwrap();
    let second = mutable
        .insert_image(
            0,
            GIF_PAYLOAD,
            &OdfLength::centimeters(2.0),
            &OdfLength::centimeters(2.0),
            &OdfLength::centimeters(2.0),
            &OdfLength::centimeters(2.0),
        )
        .unwrap();
    let first_generation = mutable.to_bytes().unwrap();

    let presentation = Presentation::from_bytes(first_generation).unwrap();
    let mut mutable = MutablePresentation::from_presentation(presentation).unwrap();
    let third = mutable
        .insert_image(
            0,
            PNG_PAYLOAD,
            &OdfLength::centimeters(3.0),
            &OdfLength::centimeters(3.0),
            &OdfLength::centimeters(3.0),
            &OdfLength::centimeters(3.0),
        )
        .unwrap();
    // Existing parts block both their own stems and sibling extensions.
    assert_eq!(third, "Pictures/image3.png");

    let bytes = mutable.to_bytes().unwrap();
    let package = OwnedPackage::from_bytes(bytes.clone()).unwrap();
    assert_eq!(package.get_file(&first).unwrap(), PNG_PAYLOAD);
    assert_eq!(package.get_file(&second).unwrap(), GIF_PAYLOAD);
    assert_eq!(package.get_file(&third).unwrap(), PNG_PAYLOAD);
    let reparsed = Presentation::from_bytes(bytes).unwrap();
    assert_eq!(reparsed.slides().unwrap()[0].shapes().unwrap().len(), 3);
}

#[test]
fn insert_image_preserves_existing_fixture_parts() {
    let fixture =
        include_bytes!("../../../test-data/libreoffice-core/sd/qa/unit/data/odp/ole_icon.odp");
    let presentation = Presentation::from_bytes(fixture.to_vec()).unwrap();
    assert_eq!(presentation.slide_count().unwrap(), 1);

    let mut mutable = MutablePresentation::from_presentation(presentation).unwrap();
    let path = mutable
        .insert_image(
            0,
            GIF_PAYLOAD,
            &OdfLength::centimeters(1.0),
            &OdfLength::centimeters(1.0),
            &OdfLength::centimeters(2.0),
            &OdfLength::centimeters(2.0),
        )
        .unwrap();
    assert_eq!(path, "Pictures/image1.gif");

    let bytes = mutable.to_bytes().unwrap();
    let package = OwnedPackage::from_bytes(bytes.clone()).unwrap();
    // The embedded OLE object of the source package survives the edit.
    assert!(package.has_file("Object 1/content.xml").unwrap());
    assert!(package.has_file("ObjectReplacements/Object 1").unwrap());
    assert_eq!(package.get_file(&path).unwrap(), GIF_PAYLOAD);

    let reparsed = Presentation::from_bytes(bytes).unwrap();
    assert_eq!(reparsed.slide_count().unwrap(), 1);
    let slides = reparsed.slides().unwrap();
    assert!(
        slides[0]
            .shapes()
            .unwrap()
            .iter()
            .any(|shape| shape.image_href() == Some(path.as_str()))
    );
}

#[test]
fn insert_image_rejects_bad_input() {
    let mut mutable = MutablePresentation::new();
    mutable.add_slide("Title", "Body").unwrap();
    let origin = &OdfLength::centimeters(0.0);
    let size = &OdfLength::centimeters(1.0);

    // Slide bounds are enforced before anything is staged.
    assert!(
        mutable
            .insert_image(9, PNG_PAYLOAD, origin, origin, size, size)
            .is_err()
    );
    // Unsupported payload formats are rejected.
    assert!(mutable.insert_image(0, b"BM bitmap", origin, origin, size, size).is_err());
    assert!(mutable.insert_image(0, b"", origin, origin, size, size).is_err());
    // Oversized payloads stay within the shared 64 MiB bound.
    let oversized = vec![0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]
        .into_iter()
        .chain(std::iter::repeat_n(0u8, 64 * 1024 * 1024 + 1))
        .collect::<Vec<_>>();
    assert!(
        mutable
            .insert_image(0, &oversized, origin, origin, size, size)
            .is_err()
    );

    // Nothing was staged by the rejected inserts.
    assert_eq!(mutable.slides()[0].shapes.len(), 0);
    let bytes = mutable.to_bytes().unwrap();
    let package = OwnedPackage::from_bytes(bytes).unwrap();
    assert!(!package.has_file("Pictures/image1.png").unwrap());
}
