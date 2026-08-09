#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::part::BlobPart;
use litchi_opc::{OpcPackage, PackURI};
use litchi_pptx::model3d::{Asset, Data, Preview};
use litchi_pptx::{Error, Package};

const MODEL_RELATIONSHIP: &str =
    "http://schemas.microsoft.com/office/2017/06/relationships/model3d";
const MODEL_CONTENT_TYPE: &str = "model/gltf-binary";
const MODEL_PART: &str = "/ppt/media/model3d.glb";
const PREVIEW_PART: &str = "/ppt/media/model3d-preview.png";
const SLIDE_PART: &str = "/ppt/slides/slide1.xml";

const MODEL_FRAME: &str = r#"<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="42" name="3D Model"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/drawing/2017/model3d"><am3d:model3d xmlns:am3d="http://schemas.microsoft.com/office/drawing/2017/model3d" xmlns:f="urn:future" r:embed="rIdModel"><am3d:spPr/><am3d:camera/><am3d:trans/><am3d:raster rName="Office Renderer" rVer="17.0"><am3d:blip r:embed="rIdPreview"/><f:future f:keep="yes"/></am3d:raster><am3d:objViewport/><f:future f:top="yes"/></am3d:model3d></a:graphicData></a:graphic></p:graphicFrame>"#;

#[test]
fn package_model3d_round_trip_preserves_payloads_relationships_and_unknowns() {
    let mut package = package_with_model();
    let model = package.model3d(0, "3D Model").unwrap().unwrap();

    assert_eq!(model.shape().name(), Some("3D Model"));
    assert_eq!(model.asset().embedded_data().unwrap().as_slice(), b"glb-v1");
    assert_eq!(
        model.preview().unwrap().data().unwrap().as_slice(),
        b"png-v1"
    );
    assert_eq!(model.preview().unwrap().content_type(), Some(ct::PNG));
    assert_eq!(model.scene().unknown_children().count(), 1);
    assert!(
        model
            .scene()
            .unknown_children()
            .any(|value| { value.local_name() == "future" && value.namespace() == "urn:future" })
    );

    let mut replacement = model.clone();
    replacement.set_asset(Asset::embedded(Data::new(b"glb-v2".to_vec()).unwrap()));
    let previous = package
        .put_model3d(0, "3D Model", replacement)
        .unwrap()
        .unwrap();
    assert_eq!(
        previous.asset().embedded_data().unwrap().as_slice(),
        b"glb-v1"
    );

    let bytes = package.to_bytes().unwrap();
    let mut reopened = Package::from_bytes(&bytes).unwrap();
    let model = reopened.model3d(0, "3D Model").unwrap().unwrap();
    assert_eq!(model.asset().embedded_data().unwrap().as_slice(), b"glb-v2");
    assert_eq!(
        model.preview().unwrap().data().unwrap().as_slice(),
        b"png-v1"
    );
    assert_eq!(model.scene().unknown_children().count(), 1);
    assert!(
        !reopened
            .opc()
            .unwrap()
            .contains_part(&PackURI::new(MODEL_PART).unwrap())
    );
    assert!(
        reopened
            .opc()
            .unwrap()
            .iter_parts()
            .any(|part| part.content_type() == MODEL_CONTENT_TYPE && part.blob() == b"glb-v2")
    );

    let removed = reopened.remove_model3d(0, "3D Model").unwrap().unwrap();
    assert_eq!(removed.shape().name(), Some("3D Model"));
    let bytes = reopened.to_bytes().unwrap();
    let reopened = Package::from_bytes(&bytes).unwrap();
    assert!(reopened.model3d(0, "3D Model").unwrap().is_none());
    assert!(
        !reopened
            .opc()
            .unwrap()
            .iter_parts()
            .any(|part| part.content_type() == MODEL_CONTENT_TYPE)
    );
    assert!(
        !reopened
            .opc()
            .unwrap()
            .iter_parts()
            .any(|part| part.content_type() == ct::PNG && part.blob() == b"png-v1")
    );
}

#[test]
fn package_model3d_rejects_wrong_payload_content_type() {
    let mut package = package_with_model();
    package
        .edit_opc(|opc| {
            opc.get_part_mut(&PackURI::new(MODEL_PART).unwrap())?
                .set_content_type("application/octet-stream".to_owned())?;
            Ok(())
        })
        .unwrap();

    assert!(matches!(
        package.model3d(0, "3D Model"),
        Err(Error::ContentType { expected, actual })
            if expected == MODEL_CONTENT_TYPE && actual == "application/octet-stream"
    ));
}

#[test]
fn package_model3d_rejects_missing_relationship_and_missing_target() {
    let mut missing_relationship = package_with_model();
    missing_relationship
        .edit_opc(|opc| {
            opc.get_part_mut(&PackURI::new(SLIDE_PART).unwrap())?
                .rels_mut()
                .remove("rIdModel");
            Ok(())
        })
        .unwrap();
    assert!(matches!(
        missing_relationship.model3d(0, "3D Model"),
        Err(Error::Invalid(message)) if message.contains("rIdModel")
    ));

    let mut missing_target = package_with_model();
    missing_target
        .edit_opc(|opc| {
            assert!(opc.remove_part(&PackURI::new(MODEL_PART).unwrap()));
            Ok(())
        })
        .unwrap();
    assert!(matches!(
        missing_target.model3d(0, "3D Model"),
        Err(Error::Opc(litchi_opc::OpcError::PartNotFound(message)))
            if message.contains(MODEL_PART)
    ));
}

#[test]
fn package_model3d_failed_edit_restores_the_snapshot() {
    let mut package = package_with_model();
    let model = package.model3d(0, "3D Model").unwrap().unwrap();
    let mut invalid = model.clone();
    invalid.set_preview(Some(
        Preview::new(
            Data::new(b"not-an-image".to_vec()).unwrap(),
            "application/octet-stream",
        )
        .unwrap(),
    ));
    let before_xml = package
        .opc()
        .unwrap()
        .get_part(&PackURI::new(SLIDE_PART).unwrap())
        .unwrap()
        .blob()
        .to_vec();
    let before_parts = package.opc().unwrap().part_count();
    assert!(package.put_model3d(0, "3D Model", invalid).is_err());
    assert_eq!(
        package
            .opc()
            .unwrap()
            .get_part(&PackURI::new(SLIDE_PART).unwrap())
            .unwrap()
            .blob(),
        before_xml.as_slice()
    );
    assert_eq!(package.opc().unwrap().part_count(), before_parts);
}

fn package_with_model() -> Package {
    let mut package = Package::new().unwrap();
    package.presentation_mut().unwrap().add_slide().unwrap();
    let bytes = package.to_bytes().unwrap();
    let mut opc = OpcPackage::from_bytes(&bytes).unwrap();
    install_model(&mut opc);
    Package::from_opc_package(opc).unwrap()
}

fn install_model(opc: &mut OpcPackage) {
    let slide_name = PackURI::new(SLIDE_PART).unwrap();
    let slide = opc.get_part_mut(&slide_name).unwrap();
    let xml = std::str::from_utf8(slide.blob()).unwrap();
    let updated = xml.replacen("</p:spTree>", &format!("{MODEL_FRAME}</p:spTree>"), 1);
    assert_ne!(updated, xml);
    slide.set_blob(updated.into_bytes());
    slide.rels_mut().add_relationship(
        MODEL_RELATIONSHIP.to_owned(),
        "../media/model3d.glb".to_owned(),
        "rIdModel".to_owned(),
        false,
    );
    slide.rels_mut().add_relationship(
        rt::IMAGE.to_owned(),
        "../media/model3d-preview.png".to_owned(),
        "rIdPreview".to_owned(),
        false,
    );
    opc.add_part(Box::new(BlobPart::new(
        PackURI::new(MODEL_PART).unwrap(),
        MODEL_CONTENT_TYPE.to_owned(),
        b"glb-v1".to_vec(),
    )));
    opc.add_part(Box::new(BlobPart::new(
        PackURI::new(PREVIEW_PART).unwrap(),
        ct::PNG.to_owned(),
        b"png-v1".to_vec(),
    )));
}
