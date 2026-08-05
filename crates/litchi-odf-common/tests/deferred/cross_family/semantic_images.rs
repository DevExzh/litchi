use litchi_odt::{
    Document, FlatDocument, ImagePart, ImageSource, Family, Package,
    Presentation, Spreadsheet,
};
use std::io::{Cursor, Write};
use zip::CompressionMethod;
use zip::write::SimpleFileOptions;

const ODT_LINK: &str = include_str!(
    "../../../test-data/libreoffice-core/sw/qa/extras/odfimport/data/draw-image-link.fodt"
);
const ODT_INLINE: &str = include_str!(
    "../../../test-data/libreoffice-core/sw/qa/extras/odfimport/data/draw-image-embedded.fodt"
);
const ODS_LINK: &str =
    include_str!("../../../test-data/libreoffice-core/sc/qa/unit/data/draw-image-link.fods");
const ODP_LINK: &str =
    include_str!("../../../test-data/libreoffice-core/sd/qa/unit/data/draw-image-link.fodp");
const ODT_HYPERLINK_ALT: &str = include_str!(
    "../../../test-data/libreoffice-core/vcl/qa/cppunit/pdfexport/data/image-hyperlink-alttext.fodt"
);
const ODP_ALTERNATIVE_TEXT: &str = include_str!(
    "../../../test-data/libreoffice-core/vcl/qa/cppunit/pdfexport/data/alternativeText.fodp"
);

#[test]
fn libreoffice_flat_links_are_typed_and_remain_inert_across_families() {
    for (xml, family, page, sheet) in [
        (ODT_LINK, Family::Text, None, None),
        (
            ODS_LINK,
            Family::Spreadsheet,
            None,
            Some("Sheet1"),
        ),
        (
            ODP_LINK,
            Family::Presentation,
            Some("page1"),
            None,
        ),
    ] {
        let document = FlatDocument::from_bytes(xml.as_bytes().to_vec()).unwrap();
        assert_eq!(document.family(), family);
        let images = document.images().unwrap();
        assert_eq!(images.len(), 1);
        assert_eq!(images[0].part, ImagePart::FlatDocument);
        assert_eq!(images[0].actuate.as_deref(), Some("onLoad"));
        assert!(matches!(
            &images[0].source,
            ImageSource::Linked { href }
                if href == "http://192.0.2.1:12345/tracking-pixel.png"
        ));
        let frame = images[0].frame.as_ref().unwrap();
        assert_eq!(frame.page_name.as_deref(), page);
        assert_eq!(frame.sheet_name.as_deref(), sheet);
    }
}

#[test]
fn libreoffice_inline_image_is_strictly_decoded_without_following_links() {
    let document = FlatDocument::from_bytes(ODT_INLINE.as_bytes().to_vec()).unwrap();
    let images = document.images().unwrap();
    assert_eq!(images.len(), 1);
    let bytes = images[0].inline_bytes().unwrap();
    assert_eq!(&bytes[..8], b"\x89PNG\r\n\x1a\n");
    assert_eq!(
        images[0].frame.as_ref().unwrap().name.as_deref(),
        Some("img1")
    );
}

#[test]
fn packaged_images_join_references_manifest_and_bytes_for_all_specialized_families() {
    for (mimetype, content) in [
        (
            "application/vnd.oasis.opendocument.text",
            content_xml("<office:text><text:p>{image}</text:p></office:text>"),
        ),
        (
            "application/vnd.oasis.opendocument.spreadsheet",
            content_xml(
                "<office:spreadsheet><table:table table:name=\"Sheet1\"><table:shapes>{image}</table:shapes></table:table></office:spreadsheet>",
            ),
        ),
        (
            "application/vnd.oasis.opendocument.presentation",
            content_xml(
                "<office:presentation><draw:page draw:name=\"Slide1\">{image}</draw:page></office:presentation>",
            ),
        ),
    ] {
        let bytes = package(mimetype, &content, "Pictures/no-extension");
        let generic = Package::from_bytes(bytes.clone()).unwrap();
        let images = generic.images().unwrap();
        assert_packaged_image(&images[0]);
        assert_eq!(
            generic.image_bytes(&images[0]).unwrap(),
            Some(b"image-data".to_vec())
        );
        assert!(
            generic
                .media_files()
                .unwrap()
                .contains(&"Pictures/unused.png".to_string())
        );
        assert_eq!(
            images.len(),
            1,
            "unreferenced package media is not an occurrence"
        );

        let specialized = if mimetype.ends_with(".text") {
            Document::from_bytes(bytes).unwrap().images().unwrap()
        } else if mimetype.ends_with(".spreadsheet") {
            Spreadsheet::from_bytes(bytes).unwrap().images().unwrap()
        } else {
            Presentation::from_bytes(bytes).unwrap().images().unwrap()
        };
        assert_packaged_image(&specialized[0]);
    }
}

#[test]
fn inline_data_wins_and_unsafe_package_traversal_is_rejected() {
    let inline = ODT_INLINE.replace(
        "<draw:image>",
        "<draw:image xlink:href=\"http://192.0.2.1/ignored.png\">",
    );
    let document = FlatDocument::from_bytes(inline.into_bytes()).unwrap();
    assert!(matches!(
        &document.images().unwrap()[0].source,
        ImageSource::Inline { ignored_href: Some(href), .. }
            if href == "http://192.0.2.1/ignored.png"
    ));

    let content = content_xml("<office:text><text:p>{image}</text:p></office:text>")
        .replace("Pictures/no-extension", "%2e%2e/secret.png");
    let document = Package::from_bytes(package(
        "application/vnd.oasis.opendocument.text",
        &content,
        "Pictures/no-extension",
    ))
    .unwrap();
    assert!(document.images().is_err());
}

#[test]
fn malformed_inline_base64_is_rejected() {
    let xml = ODT_INLINE.replace(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8",
        "not!base64",
    );
    let document = FlatDocument::from_bytes(xml.into_bytes()).unwrap();
    assert!(document.images().is_err());
}

#[test]
fn libreoffice_image_accessibility_metadata_is_typed_after_image_content() {
    for (xml, title, description) in [
        (ODT_HYPERLINK_ALT, "Ship drawing", "Very cute"),
        (
            ODP_ALTERNATIVE_TEXT,
            "This is the text alternative",
            "This is the description",
        ),
    ] {
        let document = FlatDocument::from_bytes(xml.as_bytes().to_vec()).unwrap();
        let images = document.images().unwrap();
        let image = images
            .iter()
            .find(|image| {
                image
                    .frame
                    .as_ref()
                    .and_then(|frame| frame.title.as_deref())
                    == Some(title)
            })
            .expect("fixture image with alternative title");
        let frame = image.frame.as_ref().unwrap();
        assert_eq!(frame.description.as_deref(), Some(description));
    }
}

#[test]
fn accessibility_metadata_is_shared_across_flat_document_families() {
    for (mimetype, body, family) in [
        (
            "application/vnd.oasis.opendocument.text",
            "<office:text><text:p>{frame}</text:p></office:text>",
            Family::Text,
        ),
        (
            "application/vnd.oasis.opendocument.spreadsheet",
            "<office:spreadsheet><table:table table:name=\"Sheet1\"><table:shapes>{frame}</table:shapes></table:table></office:spreadsheet>",
            Family::Spreadsheet,
        ),
        (
            "application/vnd.oasis.opendocument.presentation",
            "<office:presentation><draw:page draw:name=\"Slide1\">{frame}</draw:page></office:presentation>",
            Family::Presentation,
        ),
    ] {
        let frame = "<draw:frame><s:title>Before &amp; &#x41;<![CDATA[ <raw>]]></s:title><s:desc/><draw:image xlink:href=\"https://example.invalid/image.png\"/></draw:frame>";
        let xml = flat_document(mimetype, &body.replace("{frame}", frame));
        let document = FlatDocument::from_bytes(xml.into_bytes()).unwrap();
        assert_eq!(document.family(), family);
        let images = document.images().unwrap();
        let frame = images[0].frame.as_ref().unwrap();
        assert_eq!(frame.title.as_deref(), Some("Before & A <raw>"));
        assert_eq!(frame.description.as_deref(), Some(""));
    }
}

#[test]
fn metadata_after_multiple_image_alternatives_is_deferred_to_frame_close() {
    let body = "<office:text><text:p><draw:frame><draw:image xlink:href=\"first.png\"/><draw:image xlink:href=\"second.png\"/><s:title>shared title</s:title><s:desc>shared description</s:desc></draw:frame></text:p></office:text>";
    let document = FlatDocument::from_bytes(
        flat_document("application/vnd.oasis.opendocument.text", body).into_bytes(),
    )
    .unwrap();
    let images = document.images().unwrap();
    assert_eq!(images.len(), 2);
    for (index, image) in images.iter().enumerate() {
        assert_eq!(image.alternative_index, index);
        let frame = image.frame.as_ref().unwrap();
        assert_eq!(frame.title.as_deref(), Some("shared title"));
        assert_eq!(frame.description.as_deref(), Some("shared description"));
    }
}

#[test]
fn accessibility_metadata_is_direct_text_only_and_unique() {
    let scoped = "<office:text><text:p><draw:frame><draw:custom-shape><s:desc>shape only</s:desc></draw:custom-shape><draw:image xlink:href=\"image.png\"/><s:title>image title</s:title></draw:frame></text:p></office:text>";
    let document = FlatDocument::from_bytes(
        flat_document("application/vnd.oasis.opendocument.text", scoped).into_bytes(),
    )
    .unwrap();
    let frame = document.images().unwrap()[0].frame.clone().unwrap();
    assert_eq!(frame.title.as_deref(), Some("image title"));
    assert_eq!(frame.description, None);

    for accessibility in [
        "<s:title>one</s:title><s:title>two</s:title>",
        "<s:desc>one</s:desc><s:desc>two</s:desc>",
        "<s:title>one<text:span>two</text:span></s:title>",
    ] {
        let body = format!(
            "<office:text><text:p><draw:frame><draw:image xlink:href=\"image.png\"/>{accessibility}</draw:frame></text:p></office:text>"
        );
        let document = FlatDocument::from_bytes(
            flat_document("application/vnd.oasis.opendocument.text", &body).into_bytes(),
        )
        .unwrap();
        assert!(document.images().is_err());
    }
}

#[test]
fn accessibility_text_limits_are_enforced_per_field_and_in_aggregate() {
    let oversized = "x".repeat(64 * 1024 + 1);
    let body = format!(
        "<office:text><text:p><draw:frame><draw:image xlink:href=\"image.png\"/><s:title>{oversized}</s:title></draw:frame></text:p></office:text>"
    );
    let document = FlatDocument::from_bytes(
        flat_document("application/vnd.oasis.opendocument.text", &body).into_bytes(),
    )
    .unwrap();
    assert!(document.images().is_err());

    let chunk = "x".repeat(64 * 1024);
    let mut frames = String::new();
    for _ in 0..64 {
        frames.push_str(&format!(
            "<draw:frame><draw:image xlink:href=\"image.png\"/><s:title>{chunk}</s:title><s:desc>{chunk}</s:desc></draw:frame>"
        ));
    }
    frames.push_str(
        "<draw:frame><draw:image xlink:href=\"image.png\"/><s:title>x</s:title></draw:frame>",
    );
    let body = format!("<office:text><text:p>{frames}</text:p></office:text>");
    let document = FlatDocument::from_bytes(
        flat_document("application/vnd.oasis.opendocument.text", &body).into_bytes(),
    )
    .unwrap();
    assert!(document.images().is_err());
}

fn assert_packaged_image(image: &litchi_odt::Image) {
    assert_eq!(image.part, ImagePart::Content);
    assert_eq!(
        image.frame.as_ref().unwrap().name.as_deref(),
        Some("Image1")
    );
    assert!(matches!(
        &image.source,
        ImageSource::PackagePart {
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
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?><office:document-content xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" xmlns:text=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\" xmlns:table=\"urn:oasis:names:tc:opendocument:xmlns:table:1.0\" xmlns:draw=\"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0\" xmlns:xlink=\"http://www.w3.org/1999/xlink\" xmlns:svg=\"urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0\"><office:body>{}</office:body></office:document-content>",
        body.replace("{image}", image)
    )
}

fn flat_document(mimetype: &str, body: &str) -> String {
    format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?><office:document xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" xmlns:text=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\" xmlns:table=\"urn:oasis:names:tc:opendocument:xmlns:table:1.0\" xmlns:draw=\"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0\" xmlns:xlink=\"http://www.w3.org/1999/xlink\" xmlns:s=\"urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0\" office:mimetype=\"{mimetype}\"><office:body>{body}</office:body></office:document>"
    )
}

fn package(mimetype: &str, content: &str, image_path: &str) -> Vec<u8> {
    let cursor = Cursor::new(Vec::new());
    let mut zip = zip::ZipWriter::new(cursor);
    let stored = SimpleFileOptions::default().compression_method(CompressionMethod::Stored);
    let deflated = SimpleFileOptions::default().compression_method(CompressionMethod::Deflated);

    zip.start_file("mimetype", stored).unwrap();
    zip.write_all(mimetype.as_bytes()).unwrap();
    zip.start_file("content.xml", deflated).unwrap();
    zip.write_all(content.as_bytes()).unwrap();
    zip.start_file(image_path, deflated).unwrap();
    zip.write_all(b"image-data").unwrap();
    zip.start_file("Pictures/unused.png", deflated).unwrap();
    zip.write_all(b"unused").unwrap();

    let manifest = format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?><manifest:manifest xmlns:manifest=\"urn:oasis:names:tc:opendocument:xmlns:manifest:1.0\" manifest:version=\"1.3\"><manifest:file-entry manifest:full-path=\"/\" manifest:media-type=\"{mimetype}\"/><manifest:file-entry manifest:full-path=\"content.xml\" manifest:media-type=\"text/xml\"/><manifest:file-entry manifest:full-path=\"{image_path}\" manifest:media-type=\"image/png\"/><manifest:file-entry manifest:full-path=\"Pictures/unused.png\" manifest:media-type=\"image/png\"/></manifest:manifest>"
    );
    zip.start_file("META-INF/manifest.xml", deflated).unwrap();
    zip.write_all(manifest.as_bytes()).unwrap();
    zip.finish().unwrap().into_inner()
}
