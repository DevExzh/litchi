#![allow(
    clippy::unwrap_used,
    reason = "integration-test assertions panic on failure by design"
)]

use litchi_odp::core::OwnedPackage;
use litchi_odp::{Builder, Presentation, image::Length as OdfLength};

const PNG: &[u8] = b"\x89PNG\r\n\x1a\nimage";
const JPEG: &[u8] = b"\xff\xd8\xff\xe0image";
const GIF: &[u8] = b"GIF89aimage";

fn dimensions() -> (OdfLength, OdfLength) {
    (OdfLength::centimeters(1.25), OdfLength::centimeters(4.5))
}

#[test]
fn builder_inserts_a_picture_that_round_trips_with_exact_payload_and_mime() {
    let mut builder = Builder::new();
    builder.add_slide("Image").unwrap();
    let (origin, size) = dimensions();
    builder
        .insert_image(0, PNG, &origin, &origin, &size, &size)
        .unwrap();

    let bytes = builder.build().unwrap();
    let package = OwnedPackage::from_bytes(bytes.clone()).unwrap();
    assert_eq!(package.get_file("Pictures/image1.png").unwrap(), PNG);
    let manifest = String::from_utf8(package.get_file("META-INF/manifest.xml").unwrap()).unwrap();
    assert!(manifest.contains("Pictures/image1.png"));
    assert!(manifest.contains("image/png"));
    let content = String::from_utf8(package.get_file("content.xml").unwrap()).unwrap();
    assert!(!content.contains('\n'));
    for attribute in [
        "svg:x=\"1.25cm\"",
        "svg:y=\"1.25cm\"",
        "svg:width=\"4.5cm\"",
        "svg:height=\"4.5cm\"",
    ] {
        assert!(content.contains(attribute), "missing {attribute}");
    }

    let presentation = Presentation::from_bytes(bytes).unwrap();
    let slides = presentation.slides().unwrap();
    let shapes = slides[0].shapes().unwrap();
    assert_eq!(shapes[0].image_href(), Some("Pictures/image1.png"));
}

#[test]
fn builder_allocates_picture_numbers_across_image_extensions() {
    let mut builder = Builder::new();
    builder.add_slide("Images").unwrap();
    let (origin, size) = dimensions();
    for bytes in [PNG, JPEG, GIF] {
        builder
            .insert_image(0, bytes, &origin, &origin, &size, &size)
            .unwrap();
    }

    let package = OwnedPackage::from_bytes(builder.build().unwrap()).unwrap();
    assert_eq!(package.get_file("Pictures/image1.png").unwrap(), PNG);
    assert_eq!(package.get_file("Pictures/image2.jpg").unwrap(), JPEG);
    assert_eq!(package.get_file("Pictures/image3.gif").unwrap(), GIF);
}

#[test]
fn rejected_image_inputs_do_not_stage_media_or_mutate_the_slide() {
    let mut builder = Builder::new();
    builder.add_slide("Image").unwrap();
    let (origin, size) = dimensions();

    assert!(
        builder
            .insert_image(0, b"BM bitmap", &origin, &origin, &size, &size)
            .is_err()
    );
    assert!(
        builder
            .insert_image(1, PNG, &origin, &origin, &size, &size)
            .is_err()
    );
    let oversized = [PNG, &vec![0; 64 * 1024 * 1024 + 1]].concat();
    assert!(
        builder
            .insert_image(0, &oversized, &origin, &origin, &size, &size)
            .is_err()
    );

    builder
        .insert_image(0, PNG, &origin, &origin, &size, &size)
        .unwrap();
    let package = OwnedPackage::from_bytes(builder.build().unwrap()).unwrap();
    assert_eq!(package.get_file("Pictures/image1.png").unwrap(), PNG);
    assert!(!package.has_file("Pictures/image2.png").unwrap());
    let content = String::from_utf8(package.get_file("content.xml").unwrap()).unwrap();
    assert_eq!(content.matches("<draw:image ").count(), 1);
}
