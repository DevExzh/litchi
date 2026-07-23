use litchi_ooxml::pptx::{EmbeddedFontStyle, Package};
use litchi_ooxml::{OoxmlError, PackURI};
use litchi_opc::constants::relationship_type as rt;
use litchi_opc::part::BlobPart;

const PRESENTATION_XML: &[u8] =
    include_bytes!("../../../test-data/ooxml/pptx/embedded-fonts/presentation.xml");
const FONT_DATA: &[u8] =
    include_bytes!("../../../test-data/ooxml/pptx/embedded-fonts/example.odttf");
const FONT_CONTENT_TYPE: &str = "application/x-fontdata";

#[test]
fn presentation_embedded_fonts_resolve_inert_resources() {
    let package = package_with_embedded_font();
    let embedded_fonts = package
        .presentation()
        .unwrap()
        .embedded_fonts()
        .unwrap()
        .unwrap();

    assert_eq!(embedded_fonts.fonts.len(), 1);
    assert_eq!(embedded_fonts.fonts[0].typeface, "Example Sans");
    assert_eq!(embedded_fonts.fonts[0].faces.len(), 1);
    assert_eq!(
        embedded_fonts.fonts[0].faces[0].style,
        EmbeddedFontStyle::Regular
    );
    assert_eq!(
        embedded_fonts.fonts[0].faces[0].relationship_id,
        "rIdFontRegular"
    );
    let resource = embedded_fonts.fonts[0].faces[0].resource.as_ref().unwrap();
    assert_eq!(resource.part_name, "/ppt/fonts/example.odttf");
    assert_eq!(resource.content_type, FONT_CONTENT_TYPE);
    assert_eq!(resource.data, FONT_DATA);
}

#[test]
fn presentation_embedded_fonts_validate_font_relationships() {
    let mut package = package_with_embedded_font();
    let presentation_name = PackURI::new("/ppt/presentation.xml").unwrap();
    let presentation = package
        .opc_package_mut()
        .get_part_mut(&presentation_name)
        .unwrap();
    presentation.rels_mut().remove("rIdFontRegular");
    presentation.rels_mut().add_relationship(
        rt::THEME.to_string(),
        "fonts/example.odttf".to_string(),
        "rIdFontRegular".to_string(),
        false,
    );

    assert!(matches!(
        package.presentation().unwrap().embedded_fonts(),
        Err(OoxmlError::InvalidFormat(message))
            if message.contains("is not a font relationship")
    ));
}

fn package_with_embedded_font() -> Package {
    let mut package = Package::new().unwrap();
    let presentation_name = PackURI::new("/ppt/presentation.xml").unwrap();
    let font_name = PackURI::new("/ppt/fonts/example.odttf").unwrap();

    {
        let presentation = package
            .opc_package_mut()
            .get_part_mut(&presentation_name)
            .unwrap();
        presentation.set_blob(PRESENTATION_XML.to_vec());
        presentation.rels_mut().add_relationship(
            rt::FONT.to_string(),
            "fonts/example.odttf".to_string(),
            "rIdFontRegular".to_string(),
            false,
        );
    }
    package.opc_package_mut().add_part(Box::new(BlobPart::new(
        font_name,
        FONT_CONTENT_TYPE.to_string(),
        FONT_DATA.to_vec(),
    )));
    package
}
