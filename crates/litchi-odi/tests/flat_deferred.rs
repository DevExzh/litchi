#![allow(clippy::unwrap_used, reason = "test assertions use unwrap for clarity")]

use litchi_core::Error as CoreError;
use litchi_odi::flat::{Error, Image as FlatImage};
use litchi_odi::source::Source;

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
    let image = FlatImage::from_bytes(FLAT_IMAGE.as_bytes()).unwrap();
    assert_eq!(image.frames().len(), 1);
    assert_eq!(image.frames()[0].name(), Some("flat-image"));
    assert_eq!(image.frames()[0].width(), Some("1cm"));
    assert_eq!(image.frames()[0].height(), Some("1cm"));
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
        FlatImage::from_bytes(wrong.as_bytes()),
        Err(Error::Core(CoreError::InvalidFormat(_)))
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
            FlatImage::from_bytes(case.as_bytes()),
            Err(Error::Core(CoreError::InvalidFormat(_)))
        ));
    }
}

#[test]
fn flat_image_enforces_the_normative_single_frame_single_image_shape() {
    let no_frame = FLAT_IMAGE.replace(
        r#"<draw:frame draw:name="flat-image" svg:width="1cm" svg:height="1cm"><draw:image><office:binary-data>aQ==</office:binary-data></draw:image></draw:frame>"#,
        "",
    );
    let empty_frame = FLAT_IMAGE.replace(
        "<draw:image><office:binary-data>aQ==</office:binary-data></draw:image>",
        "",
    );
    let two_images = FLAT_IMAGE.replace(
        "</draw:frame>",
        r#"<draw:image xlink:href="second.png" xmlns:xlink="http://www.w3.org/1999/xlink"/></draw:frame>"#,
    );
    let two_frames = FLAT_IMAGE.replace(
        "</office:image>",
        r#"<draw:frame><draw:image xlink:href="second.png" xmlns:xlink="http://www.w3.org/1999/xlink"/></draw:frame></office:image>"#,
    );
    let spoofed = FLAT_IMAGE
        .replace("<draw:frame ", "<fake:frame xmlns:fake=\"urn:fake\" ")
        .replace("</draw:frame>", "</fake:frame>");
    for case in [no_frame, empty_frame, two_images, two_frames, spoofed] {
        assert!(matches!(
            FlatImage::from_bytes(case.as_bytes()),
            Err(Error::Core(CoreError::InvalidFormat(_)))
        ));
    }
}

#[test]
fn flat_image_rejects_excessive_xml_depth() {
    let opening = "<foreign:n>".repeat(4_097);
    let closing = "</foreign:n>".repeat(4_097);
    let deep = FLAT_IMAGE
        .replace("xmlns:svg=", r#"xmlns:foreign="urn:foreign" xmlns:svg="#)
        .replace("<draw:image>", &format!("<draw:image>{opening}"))
        .replace("</draw:image>", &format!("{closing}</draw:image>"));
    assert!(matches!(
        FlatImage::from_bytes(deep.as_bytes()),
        Err(Error::Core(CoreError::InvalidFormat(_)))
    ));
}

#[test]
fn flat_image_takes_caller_limits_through_open_options() {
    use litchi_odi::{LimitKind, OpenOptions, SemanticLimitKind, SemanticLimits};

    let image = FlatImage::from_bytes(FLAT_IMAGE.as_bytes()).unwrap();
    assert_eq!(image.options(), OpenOptions::new());
    let opening = "<foreign:n>".repeat(300);
    let closing = "</foreign:n>".repeat(300);
    let deep = FLAT_IMAGE
        .replace("xmlns:svg=", r#"xmlns:foreign="urn:foreign" xmlns:svg="#)
        .replace("<draw:image>", &format!("<draw:image>{opening}"))
        .replace("</draw:image>", &format!("{closing}</draw:image>"));
    let parsed = FlatImage::from_bytes(deep.as_bytes());
    assert!(parsed.is_ok(), "{:?}", parsed.err());
    let shallow = OpenOptions::new().semantic_limits(
        SemanticLimits::builder()
            .max_depth(256)
            .unwrap()
            .build()
            .unwrap(),
    );
    assert!(matches!(
        FlatImage::from_bytes_with_options(deep.as_bytes(), shallow),
        Err(Error::Limit {
            kind: LimitKind::Semantic(SemanticLimitKind::Depth),
            ..
        })
    ));
    let no_frames = OpenOptions::new().semantic_limits(
        SemanticLimits::builder()
            .max_frames(1)
            .unwrap()
            .build()
            .unwrap(),
    );
    assert!(FlatImage::from_bytes_with_options(FLAT_IMAGE.as_bytes(), no_frames).is_ok());
    let mut edit = image.edit();
    edit.set_frame_name(0, Some("renamed".into())).unwrap();
    edit.reset();
    let noop = edit.commit().unwrap();
    assert!(!noop.diagnostics().changed());
    assert!(noop.patch().is_noop());
    let reread = FlatImage::from_reader(std::io::Cursor::new(FLAT_IMAGE.as_bytes())).unwrap();
    assert_eq!(reread.as_bytes(), image.as_bytes());
}
