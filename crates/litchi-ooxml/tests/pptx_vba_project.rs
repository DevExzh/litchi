use litchi_ooxml::pptx::Package;
use litchi_ooxml::{OoxmlError, PackURI};
use litchi_opc::OpcPackage;
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::part::{BlobPart, Part};

const PRESENTATION_XML: &[u8] =
    include_bytes!("../../../test-data/ooxml/pptx/vba-project/presentation.xml");

#[test]
fn presentation_discovers_inert_vba_project_metadata() {
    let package = package_with_vba_project(false);

    let project = package
        .presentation()
        .unwrap()
        .vba_project()
        .unwrap()
        .unwrap();
    assert_eq!(project.source_part_name().as_str(), "/ppt/presentation.xml");
    assert_eq!(project.relationship_id(), "rIdVbaProject");
    assert_eq!(project.project_part_name().as_str(), "/ppt/vbaProject.bin");
}

#[test]
fn presentation_rejects_external_vba_project_relationships() {
    let package = package_with_vba_project(true);

    assert!(matches!(
        package.presentation().unwrap().vba_project(),
        Err(OoxmlError::InvalidFormat(message)) if message.contains("cannot be external")
    ));
}

fn package_with_vba_project(external: bool) -> Package {
    let presentation_name = PackURI::new("/ppt/presentation.xml").unwrap();
    let project_name = PackURI::new("/ppt/vbaProject.bin").unwrap();
    let mut presentation = BlobPart::new(
        presentation_name,
        ct::PML_PRES_MACRO_MAIN.to_string(),
        PRESENTATION_XML.to_vec(),
    );
    presentation.rels_mut().add_relationship(
        rt::VBA_PROJECT.to_string(),
        if external {
            "https://example.invalid/vbaProject.bin"
        } else {
            "vbaProject.bin"
        }
        .to_string(),
        "rIdVbaProject".to_string(),
        external,
    );

    let mut opc = OpcPackage::new();
    opc.add_part(Box::new(presentation));
    if !external {
        opc.add_part(Box::new(BlobPart::new(
            project_name,
            ct::OFC_VBA_PROJECT.to_string(),
            b"opaque macro payload".to_vec(),
        )));
    }
    opc.relate_to("ppt/presentation.xml", rt::OFFICE_DOCUMENT);
    Package::from_opc_package(opc).unwrap()
}
