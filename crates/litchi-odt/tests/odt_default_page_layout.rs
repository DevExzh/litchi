use litchi_odt::generic::FlatDocument;
use std::error::Error;

const STYLE_NS: &str = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
const PREFIX: &str = "<?xml version=\"1.0\" encoding=\"UTF-8\"?><office:document xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" xmlns:style=\"urn:oasis:names:tc:opendocument:xmlns:style:1.0\" office:mimetype=\"application/vnd.oasis.opendocument.text\" office:version=\"1.4\"><office:styles>";
const SUFFIX: &str = "</office:styles><office:body><office:text/></office:body></office:document>";

fn flat(styles: &str) -> Result<FlatDocument, Box<dyn Error>> {
    Ok(FlatDocument::from_bytes(
        format!("{PREFIX}{styles}{SUFFIX}").into_bytes(),
    )?)
}

#[test]
fn reads_default_page_layout_from_a_flat_document() -> Result<(), Box<dyn Error>> {
    let document = flat(
        "<style:default-page-layout><style:page-layout-properties style:writing-mode=\"lr-tb\"/></style:default-page-layout>",
    )?;
    let layout = document
        .default_page_layout()?
        .ok_or_else(|| std::io::Error::other("fixture declares a default page layout"))?;
    assert!(layout.name.is_empty(), "default layout is unnamed");
    let properties = layout
        .properties
        .as_ref()
        .ok_or_else(|| std::io::Error::other("layout properties"))?;
    assert_eq!(
        properties.attribute(Some(STYLE_NS), "writing-mode"),
        Some("lr-tb")
    );
    assert_eq!(
        layout.xml,
        "<style:default-page-layout><style:page-layout-properties style:writing-mode=\"lr-tb\"/></style:default-page-layout>"
    );
    Ok(())
}

#[test]
fn documents_without_default_page_layout_report_none() -> Result<(), Box<dyn Error>> {
    assert!(flat("")?.default_page_layout()?.is_none());
    Ok(())
}
