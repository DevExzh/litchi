#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

use litchi_opc::constants::relationship_type as rt;
use litchi_opc::part::BlobPart;
use litchi_opc::{OpcPackage, PackURI};
use litchi_pptx::font;
use litchi_pptx::font::Style;
use litchi_pptx::{Error, Package};

const PRESENTATION_XML: &[u8] =
    include_bytes!("../../../test-data/ooxml/pptx/embedded-fonts/presentation.xml");
const FONT_DATA: &[u8] =
    include_bytes!("../../../test-data/ooxml/pptx/embedded-fonts/example.odttf");
const FONT_CONTENT_TYPE: &str = "application/x-fontdata";

#[test]
fn presentation_embedded_fonts_resolve_inert_resources() {
    let package = package_with_embedded_font();
    let fonts = font::load(package.opc().unwrap()).unwrap().unwrap();

    assert_eq!(fonts.len(), 1);
    let font = fonts.get("Example Sans").unwrap();
    assert_eq!(font.name(), "Example Sans");
    assert_eq!(font.faces().len(), 1);
    assert_eq!(font.faces()[0].style(), Style::Regular);
    assert_eq!(font.faces()[0].data().bytes(), FONT_DATA);
}

#[test]
fn presentation_embedded_fonts_validate_font_relationships() {
    let mut package = package_with_embedded_font();
    let presentation_name = PackURI::new("/ppt/presentation.xml").unwrap();
    package = edit_package(package, |opc| {
        let presentation = opc.get_part_mut(&presentation_name).unwrap();
        presentation.rels_mut().remove("rIdFontRegular");
        presentation.rels_mut().add_relationship(
            rt::THEME.to_string(),
            "fonts/example.odttf".to_string(),
            "rIdFontRegular".to_string(),
            false,
        );
    });

    assert!(matches!(
        font::load(package.opc().unwrap()),
        Err(Error::Invalid(message))
            if message.contains("does not match the presentation conformance")
    ));
}

fn package_with_embedded_font() -> Package {
    let mut package = Package::new().unwrap();
    let presentation_name = PackURI::new("/ppt/presentation.xml").unwrap();
    let font_name = PackURI::new("/ppt/fonts/example.odttf").unwrap();

    let bytes = package.to_bytes().unwrap();
    let mut opc = OpcPackage::from_bytes(&bytes).unwrap();
    {
        let presentation = opc.get_part_mut(&presentation_name).unwrap();
        presentation.set_blob(PRESENTATION_XML.to_vec());
        presentation.rels_mut().add_relationship(
            rt::FONT.to_string(),
            "fonts/example.odttf".to_string(),
            "rIdFontRegular".to_string(),
            false,
        );
        opc.add_part(Box::new(BlobPart::new(
            font_name,
            FONT_CONTENT_TYPE.to_string(),
            FONT_DATA.to_vec(),
        )));
    }
    Package::from_opc_package(opc).unwrap()
}

fn edit_package(mut package: Package, edit: impl FnOnce(&mut OpcPackage)) -> Package {
    let bytes = package.to_bytes().unwrap();
    let mut opc = OpcPackage::from_bytes(&bytes).unwrap();
    edit(&mut opc);
    Package::from_opc_package(opc).unwrap()
}
