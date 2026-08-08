#![allow(clippy::unwrap_used, reason = "test assertions use unwrap for clarity")]

use litchi_core::Error;
use litchi_odi::{FlatImage, source::Source};

const FLAT_IMAGE: &str = concat!(
    r#"<?xml version="1.0" encoding="UTF-8"?>"#,
    r#"<office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" "#,
    r#"xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" "#,
    r#"xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" "#,
    r#"office:mimetype="application/vnd.oasis.opendocument.image"><office:body><office:image>"#,
    r#"<draw:frame draw:name="flat-image" svg:width="1cm" svg:height="1cm">"#,
    r#"<draw:image><office:binary-data>aQ==</office:binary-data></draw:image>"#,
    r#"</draw:frame></office:image></office:body></office:document>"#,
);

#[test]
fn flat_image_exposes_frame_and_preserves_bytes() {
    let image = FlatImage::from_bytes(FLAT_IMAGE.as_bytes().to_vec()).unwrap();
    assert_eq!(image.frames().len(), 1);
    assert_eq!(image.frames()[0].name(), Some("flat-image"));
    assert_eq!(image.frames()[0].source(), &Source::Embedded(vec![b'i']));
    assert_eq!(image.as_bytes(), FLAT_IMAGE.as_bytes());
}

#[test]
fn flat_image_rejects_wrong_family() {
    let wrong = FLAT_IMAGE.replace(
        "application/vnd.oasis.opendocument.image",
        "application/vnd.oasis.opendocument.presentation",
    );
    assert!(matches!(
        FlatImage::from_bytes(wrong.into_bytes()),
        Err(Error::InvalidFormat(_))
    ));
}

#[test]
fn flat_image_requires_namespace_aware_body_and_image_placement() {
    let cases = [
        FLAT_IMAGE
            .replace("<office:body>", "")
            .replace("</office:body>", ""),
        FLAT_IMAGE
            .replace("<office:image>", "")
            .replace("</office:image>", ""),
        FLAT_IMAGE
            .replace("<office:body><office:image>", "<office:image><office:body>")
            .replace(
                "</office:image></office:body>",
                "</office:body></office:image>",
            ),
        FLAT_IMAGE.replace(
            "</office:image></office:body>",
            "</office:image><draw:image/></office:body>",
        ),
    ];
    for case in cases {
        assert!(matches!(
            FlatImage::from_bytes(case.into_bytes()),
            Err(Error::InvalidFormat(_))
        ));
    }
}
