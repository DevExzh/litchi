use super::{Image, Source, scan_flat, scan_package};
use crate::drawing::Part;
use crate::package::PackageLookup;

const PREFIX: &str = r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:d="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:s="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:t="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:x="http://www.w3.org/1999/xlink" xmlns:xml="http://www.w3.org/XML/1998/namespace"><o:body>"#;
const SUFFIX: &str = "</o:body></o:document-content>";

struct Lookup;

impl PackageLookup for Lookup {
    fn has_file(&self, path: &str) -> bool {
        path == "Pictures/photo.png"
    }

    fn media_type(&self, path: &str) -> Option<&str> {
        (path == "Pictures/photo.png").then_some("image/png")
    }
}

#[test]
fn package_scan_retains_frame_context_and_classifies_sources() -> litchi_core::Result<()> {
    let xml = format!(
        r#"{PREFIX}<d:page d:name="Page 1"><d:frame d:name="Photo" xml:id="photo" t:end-cell-address="B2" s:x="1cm" s:width="2cm"><s:title>A &amp; B</s:title><s:desc><![CDATA[Portrait]]></s:desc><d:image x:href="Pictures/photo.png" x:type="simple"/><d:image x:href="ignored.bin"><o:binary-data>AQID</o:binary-data></d:image></d:frame></d:page><t:table t:name="Sheet 1"><t:shapes><d:frame><d:image x:href="Pictures/missing.png"/></d:frame></t:shapes></t:table>{SUFFIX}"#
    );

    let images = scan_package(&xml, None, &Lookup)?;
    assert_eq!(images.len(), 3);
    assert_eq!(images[0].part, Part::Content);
    assert_eq!(images[0].alternative_index, 0);
    let frame = images[0].frame.as_ref().ok_or_else(|| {
        litchi_core::Error::InvalidFormat("expected the first image to have a frame".into())
    })?;
    assert_eq!(frame.name.as_deref(), Some("Photo"));
    assert_eq!(frame.xml_id.as_deref(), Some("photo"));
    assert_eq!(frame.title.as_deref(), Some("A & B"));
    assert_eq!(frame.description.as_deref(), Some("Portrait"));
    assert_eq!(frame.page_name.as_deref(), Some("Page 1"));
    assert_eq!(frame.end_cell_address.as_deref(), Some("B2"));
    assert!(matches!(
        images[0].source,
        Source::PackagePart {
            ref path,
            manifest_media_type: Some(ref media),
            ..
        } if path == "Pictures/photo.png" && media == "image/png"
    ));
    assert_eq!(images[1].alternative_index, 1);
    assert_eq!(images[1].inline_bytes(), Some(&[1, 2, 3][..]));
    assert!(matches!(
        images[1].source,
        Source::Inline {
            ref ignored_href,
            ..
        } if ignored_href.as_deref() == Some("ignored.bin")
    ));
    let sheet_frame = images[2].frame.as_ref().ok_or_else(|| {
        litchi_core::Error::InvalidFormat("expected the third image to have a frame".into())
    })?;
    assert_eq!(sheet_frame.sheet_name.as_deref(), Some("Sheet 1"));
    assert!(sheet_frame.sheet_shape);
    Ok(())
}

#[test]
fn flat_scan_keeps_external_links_inert_and_rejects_dtds() -> litchi_core::Result<()> {
    let xml = format!(
        r#"{PREFIX}<d:frame><d:image x:href="https://example.invalid/image.png"/><d:image/></d:frame>{SUFFIX}"#
    );
    let images = scan_flat(&xml)?;
    assert!(matches!(
        images[0].source,
        Source::Linked { ref href } if href == "https://example.invalid/image.png"
    ));
    assert!(matches!(images[1].source, Source::Missing));

    let dtd = r#"<!DOCTYPE office:document-content><o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"/>"#;
    assert!(scan_flat(dtd).is_err());
    Ok(())
}

#[test]
fn malformed_inline_data_is_rejected_without_execution() {
    let xml = format!(
        r"{PREFIX}<d:frame><d:image><o:binary-data>not-base64!</o:binary-data></d:image></d:frame>{SUFFIX}"
    );
    assert!(scan_flat(&xml).is_err());
}

#[test]
fn image_inventory_helpers_are_borrowed() {
    let image = Image {
        part: Part::FlatDocument,
        source: Source::Inline {
            bytes: vec![4, 5],
            ignored_href: None,
        },
        frame: None,
        xml_id: None,
        filter_name: None,
        declared_media_type: None,
        link_type: None,
        show: None,
        actuate: None,
        alternative_index: 0,
    };
    assert_eq!(image.inline_bytes(), Some(&[4, 5][..]));
    assert_eq!(image.package_path(), None);
}
