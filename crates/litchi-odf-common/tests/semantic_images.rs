use litchi_odf_common::{
    drawing::Part,
    media::{Image, Source, scan_flat, scan_package},
    package::PackageLookup,
};

const ODT_LINK: &str = r#"<office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:xlink="http://www.w3.org/1999/xlink"><office:body><office:text><text:p><draw:frame><draw:image xlink:href="http://192.0.2.1:12345/tracking-pixel.png" xlink:actuate="onLoad"/></draw:frame></text:p></office:text></office:body></office:document>"#;
const ODT_INLINE: &str = r#"<office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:xlink="http://www.w3.org/1999/xlink"><office:body><office:text><text:p><draw:frame draw:name="img1"><draw:image><office:binary-data>iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAFgAH/q8422QAAAABJRU5ErkJggg==</office:binary-data></draw:image></draw:frame></text:p></office:text></office:body></office:document>"#;
const ODS_LINK: &str = r#"<office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:xlink="http://www.w3.org/1999/xlink"><office:body><office:spreadsheet><table:table table:name="Sheet1"><table:shapes><draw:frame><draw:image xlink:href="http://192.0.2.1:12345/tracking-pixel.png" xlink:actuate="onLoad"/></draw:frame></table:shapes></table:table></office:spreadsheet></office:body></office:document>"#;
const ODP_LINK: &str = r#"<office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:xlink="http://www.w3.org/1999/xlink"><office:body><office:presentation><draw:page draw:name="page1"><draw:frame><draw:image xlink:href="http://192.0.2.1:12345/tracking-pixel.png" xlink:actuate="onLoad"/></draw:frame></draw:page></office:presentation></office:body></office:document>"#;
const ODT_HYPERLINK_ALT: &str = r#"<office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"><office:body><office:text><text:p><draw:frame><draw:image xlink:href="https://example.invalid/ship.png"/><svg:title>Ship drawing</svg:title><svg:desc>Very cute</svg:desc></draw:frame></text:p></office:text></office:body></office:document>"#;
const ODP_ALTERNATIVE_TEXT: &str = r#"<office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"><office:body><office:presentation><draw:page draw:name="Slide1"><draw:frame><draw:image xlink:href="https://example.invalid/slide.png"/><svg:title>This is the text alternative</svg:title><svg:desc>This is the description</svg:desc></draw:frame></draw:page></office:presentation></office:body></office:document>"#;

struct Lookup<'a> {
    entries: &'a [(&'a str, Option<&'a str>)],
}

impl PackageLookup for Lookup<'_> {
    fn has_file(&self, path: &str) -> bool {
        self.entries.iter().any(|(entry, _)| *entry == path)
    }

    fn media_type(&self, path: &str) -> Option<&str> {
        self.entries
            .iter()
            .find_map(|(entry, media_type)| (*entry == path).then_some(*media_type).flatten())
    }
}

#[test]
fn libreoffice_flat_links_are_typed_and_inert_across_families() {
    for (xml, page, sheet) in [
        (ODT_LINK, None, None),
        (ODS_LINK, None, Some("Sheet1")),
        (ODP_LINK, Some("page1"), None),
    ] {
        let images = scan_flat(xml).unwrap();
        assert_eq!(images.len(), 1);
        assert_eq!(images[0].part, Part::FlatDocument);
        assert_eq!(images[0].actuate.as_deref(), Some("onLoad"));
        assert!(matches!(
            &images[0].source,
            Source::Linked { href } if href == "http://192.0.2.1:12345/tracking-pixel.png"
        ));
        let frame = images[0].frame.as_ref().unwrap();
        assert_eq!(frame.page_name.as_deref(), page);
        assert_eq!(frame.sheet_name.as_deref(), sheet);
    }
}

#[test]
fn libreoffice_inline_image_is_strictly_decoded_without_following_links() {
    let images = scan_flat(ODT_INLINE).unwrap();
    assert_eq!(images.len(), 1);
    assert_eq!(
        &images[0].inline_bytes().unwrap()[..8],
        b"\x89PNG\r\n\x1a\n"
    );
    assert_eq!(
        images[0].frame.as_ref().unwrap().name.as_deref(),
        Some("img1")
    );
}

#[test]
fn packaged_images_join_references_and_manifest_metadata() {
    let lookup = Lookup {
        entries: &[("Pictures/no-extension", Some("image/png"))],
    };
    for body in [
        "<office:text><text:p>{image}</text:p></office:text>",
        "<office:spreadsheet><table:table table:name=\"Sheet1\"><table:shapes>{image}</table:shapes></table:table></office:spreadsheet>",
        "<office:presentation><draw:page draw:name=\"Slide1\">{image}</draw:page></office:presentation>",
    ] {
        let images = scan_package(&content_xml(body), None, &lookup).unwrap();
        assert_eq!(images.len(), 1);
        assert_packaged_image(&images[0]);
    }
}

#[test]
fn inline_data_wins_and_unsafe_package_traversal_is_rejected() {
    let inline = ODT_INLINE.replace(
        "<draw:image>",
        "<draw:image xlink:href=\"http://192.0.2.1/ignored.png\">",
    );
    assert!(matches!(
        &scan_flat(&inline).unwrap()[0].source,
        Source::Inline { ignored_href: Some(href), .. }
            if href == "http://192.0.2.1/ignored.png"
    ));

    let lookup = Lookup { entries: &[] };
    let xml = content_xml("<office:text><text:p>{image}</text:p></office:text>")
        .replace("Pictures/no-extension", "%2e%2e/secret.png");
    assert!(scan_package(&xml, None, &lookup).is_err());
}

#[test]
fn malformed_inline_base64_is_rejected() {
    let xml = ODT_INLINE.replace(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8",
        "not!base64",
    );
    assert!(scan_flat(&xml).is_err());
}

#[test]
fn libreoffice_accessibility_metadata_is_typed_after_image_content() {
    for (xml, title, description) in [
        (ODT_HYPERLINK_ALT, "Ship drawing", "Very cute"),
        (
            ODP_ALTERNATIVE_TEXT,
            "This is the text alternative",
            "This is the description",
        ),
    ] {
        let images = scan_flat(xml).unwrap();
        let image = images
            .iter()
            .find(|image| {
                image
                    .frame
                    .as_ref()
                    .and_then(|frame| frame.title.as_deref())
                    == Some(title)
            })
            .unwrap();
        assert_eq!(
            image.frame.as_ref().unwrap().description.as_deref(),
            Some(description)
        );
    }
}

#[test]
fn accessibility_metadata_is_shared_and_deferred_to_frame_close() {
    for body in [
        "<office:text><text:p>{frame}</text:p></office:text>",
        "<office:spreadsheet><table:table table:name=\"Sheet1\"><table:shapes>{frame}</table:shapes></table:table></office:spreadsheet>",
        "<office:presentation><draw:page draw:name=\"Slide1\">{frame}</draw:page></office:presentation>",
    ] {
        let frame = "<draw:frame><draw:image xlink:href=\"first.png\"/><draw:image xlink:href=\"second.png\"/><svg:title>Before &amp; &#x41;<![CDATA[ <raw>]]></svg:title><svg:desc/></draw:frame>";
        let images = scan_flat(&flat_document(&body.replace("{frame}", frame))).unwrap();
        assert_eq!(images.len(), 2);
        for (index, image) in images.iter().enumerate() {
            assert_eq!(image.alternative_index, index);
            let frame = image.frame.as_ref().unwrap();
            assert_eq!(frame.title.as_deref(), Some("Before & A <raw>"));
            assert_eq!(frame.description.as_deref(), Some(""));
        }
    }
}

#[test]
fn accessibility_metadata_is_direct_text_only_unique_and_bounded() {
    let scoped = "<office:text><text:p><draw:frame><draw:custom-shape><svg:desc>shape only</svg:desc></draw:custom-shape><draw:image xlink:href=\"image.png\"/><svg:title>image title</svg:title></draw:frame></text:p></office:text>";
    let images = scan_flat(&flat_document(scoped)).unwrap();
    let frame = images[0].frame.as_ref().unwrap();
    assert_eq!(frame.title.as_deref(), Some("image title"));
    assert_eq!(frame.description, None);

    for accessibility in [
        "<svg:title>one</svg:title><svg:title>two</svg:title>".to_string(),
        "<svg:desc>one</svg:desc><svg:desc>two</svg:desc>".to_string(),
        "<svg:title>one<text:span>two</text:span></svg:title>".to_string(),
        format!("<svg:title>{}</svg:title>", "x".repeat(64 * 1024 + 1)),
    ] {
        let body = format!(
            "<office:text><text:p><draw:frame><draw:image xlink:href=\"image.png\"/>{accessibility}</draw:frame></text:p></office:text>"
        );
        assert!(scan_flat(&flat_document(&body)).is_err());
    }
}

fn assert_packaged_image(image: &Image) {
    assert_eq!(image.part, Part::Content);
    assert_eq!(
        image.frame.as_ref().unwrap().name.as_deref(),
        Some("Image1")
    );
    assert!(matches!(
        &image.source,
        Source::PackagePart {
            href,
            path,
            manifest_media_type: Some(media_type),
        } if href == "Pictures/no-extension"
            && path == "Pictures/no-extension"
            && media_type == "image/png"
    ));
}

fn content_xml(body: &str) -> String {
    let image = "<draw:frame draw:name=\"Image1\" svg:width=\"1cm\" svg:height=\"2cm\"><draw:image xlink:href=\"Pictures/no-extension\" xlink:type=\"simple\" xlink:show=\"embed\"/></draw:frame>";
    format!(
        "<office:document-content xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" xmlns:text=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\" xmlns:table=\"urn:oasis:names:tc:opendocument:xmlns:table:1.0\" xmlns:draw=\"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0\" xmlns:xlink=\"http://www.w3.org/1999/xlink\" xmlns:svg=\"urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0\"><office:body>{}</office:body></office:document-content>",
        body.replace("{image}", image)
    )
}

fn flat_document(body: &str) -> String {
    format!(
        "<office:document xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" xmlns:text=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\" xmlns:table=\"urn:oasis:names:tc:opendocument:xmlns:table:1.0\" xmlns:draw=\"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0\" xmlns:xlink=\"http://www.w3.org/1999/xlink\" xmlns:svg=\"urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0\" office:mimetype=\"application/vnd.oasis.opendocument.text\"><office:body>{body}</office:body></office:document>"
    )
}
