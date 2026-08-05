use litchi_opc::{OpcPackage, PackURI, Part};
use litchi_pptx::{
    Error, Package,
    presentation_properties::{
        self, PrintColorMode, PrintOutput, ShowMode, SlideSelection, ShowExtension,
    },
    view_properties::{self, ViewKind},
};

const VIEW_PROPERTIES_RELATIONSHIP_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/viewProps";
const PRESENTATION_PROPERTIES_RELATIONSHIP_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps";
const LOCAL_VIEW_PROPERTIES: &[u8] =
    include_bytes!("../../../test-data/ooxml/pptx/view-properties/basic_view.xml");
const LOCAL_PRESENTATION_PROPERTIES: &[u8] =
    include_bytes!("../../../test-data/ooxml/pptx/presentation-properties/basic_presentation.xml");

#[test]
fn package_loads_local_view_and_presentation_settings() {
    let package = package_with_local_settings();
    let opc = package.opc().unwrap();

    let view_properties = view_properties::load_from_package(opc).unwrap().unwrap();
    assert_eq!(view_properties.last_view, Some(ViewKind::Slide));
    assert_eq!(view_properties.show_comments, Some(true));
    assert_eq!(view_properties.grid_spacing.as_ref().unwrap().cx, 72_000);

    let presentation_properties =
        presentation_properties::load_from_package(opc).unwrap().unwrap();
    let web = presentation_properties.web.as_ref().unwrap();
    assert_eq!(web.allow_png, Some(true));
    assert_eq!(
        presentation_properties.print.as_ref().unwrap().output,
        Some(PrintOutput::Handouts6)
    );
    assert_eq!(
        presentation_properties.print.as_ref().unwrap().color_mode,
        Some(PrintColorMode::Gray)
    );
    let show = presentation_properties.show.as_ref().unwrap();
    assert_eq!(show.mode, Some(ShowMode::Kiosk { restart: Some(5) }));
    assert_eq!(
        show.extensions,
        vec![ShowExtension::BrowseMode {
            show_status: Some(false)
        }]
    );
    assert_eq!(
        show.slides,
        Some(SlideSelection::Range { start: 2, end: 4 })
    );
}

#[test]
fn package_readers_report_absent_settings() {
    let mut opc = base_opc();
    remove_presentation_relationships(&mut opc, VIEW_PROPERTIES_RELATIONSHIP_TYPE);
    remove_presentation_relationships(&mut opc, PRESENTATION_PROPERTIES_RELATIONSHIP_TYPE);

    assert_eq!(view_properties::load_from_package(&opc).unwrap(), None);
    assert_eq!(presentation_properties::load_from_package(&opc).unwrap(), None);
}

#[test]
fn package_readers_reject_external_settings_relationships() {
    let mut opc = base_opc();
    let presentation_name = PackURI::new("/ppt/presentation.xml").unwrap();
    remove_presentation_relationships(&mut opc, VIEW_PROPERTIES_RELATIONSHIP_TYPE);
    opc.get_part_mut(&presentation_name)
        .unwrap()
        .rels_mut()
        .add_relationship(
            VIEW_PROPERTIES_RELATIONSHIP_TYPE.to_string(),
            "https://example.invalid/view-properties.xml".to_string(),
            "rIdExternalViewProperties".to_string(),
            true,
        );

    assert!(matches!(
        view_properties::load_from_package(&opc),
        Err(Error::Invalid(message)) if message.contains("cannot be external")
    ));

    remove_presentation_relationships(&mut opc, PRESENTATION_PROPERTIES_RELATIONSHIP_TYPE);
    opc.get_part_mut(&presentation_name)
        .unwrap()
        .rels_mut()
        .add_relationship(
            PRESENTATION_PROPERTIES_RELATIONSHIP_TYPE.to_string(),
            "https://example.invalid/presentation-properties.xml".to_string(),
            "rIdExternalPresentationProperties".to_string(),
            true,
        );

    assert!(matches!(
        presentation_properties::load_from_package(&opc),
        Err(Error::Invalid(message)) if message.contains("cannot be external")
    ));
}

fn base_opc() -> OpcPackage {
    let mut package = Package::new().unwrap();
    let package_bytes = package.to_bytes().unwrap();
    OpcPackage::from_bytes(&package_bytes).unwrap()
}

fn package_with_local_settings() -> Package {
    let mut opc = base_opc();
    let view_properties_name = PackURI::new("/ppt/viewProps.xml").unwrap();
    let presentation_properties_name = PackURI::new("/ppt/presProps.xml").unwrap();
    opc.get_part_mut(&view_properties_name)
        .unwrap()
        .set_blob(LOCAL_VIEW_PROPERTIES.to_vec());
    opc.get_part_mut(&presentation_properties_name)
        .unwrap()
        .set_blob(LOCAL_PRESENTATION_PROPERTIES.to_vec());
    Package::from_opc_package(opc).unwrap()
}

fn remove_presentation_relationships(package: &mut OpcPackage, relationship_type: &str) {
    let presentation_name = PackURI::new("/ppt/presentation.xml").unwrap();
    let presentation = package.get_part_mut(&presentation_name).unwrap();
    let relationship_ids = presentation
        .rels()
        .iter()
        .filter(|relationship| relationship.reltype() == relationship_type)
        .map(|relationship| relationship.r_id().to_string())
        .collect::<Vec<_>>();
    for relationship_id in relationship_ids {
        presentation.rels_mut().remove(&relationship_id);
    }
}
