#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

use super::super::model::{Frame, Kind, Mode, Target};
use super::{Definition, Snapshot};
use crate::presentation::embedded::ole::slide::apply_patch;
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::part::BlobPart;
use litchi_opc::{OpcPackage, PackURI, Part};

const SLIDE: &str = "/ppt/slides/slide1.xml";
const PAYLOAD: &str = "/ppt/embeddings/oleObject1.bin";
const PREVIEW: &str = "/ppt/media/image1.png";

fn slide_xml(linked: bool) -> Vec<u8> {
    let branch = if linked {
        r#"<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:v="urn:unsupported"><mc:Choice Requires="v"><p:embed r:id="rIdIgnored"/></mc:Choice><mc:Fallback><p:link/></mc:Fallback></mc:AlternateContent>"#
    } else {
        r#"<p:embed/><p:pic><p:blipFill><a:blip r:embed="rIdPreview"/></p:blipFill></p:pic>"#
    };
    format!(
        r#"<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:x="urn:opaque"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name="root"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="101" name="Frame"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="10" y="20"/><a:ext cx="300" cy="400"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/presentationml/2006/ole"><p:oleObj name="Old" progId="Old.App" showAsIcon="0" imgW="30" imgH="40" r:id="rIdOle"><x:opaque>keep</x:opaque>{branch}</p:oleObj></a:graphicData></a:graphic></p:graphicFrame></p:spTree></p:cSld></p:sld>"#
    )
    .into_bytes()
}

fn package(linked: bool) -> OpcPackage {
    let slide_name = PackURI::new(SLIDE).unwrap();
    let mut slide = BlobPart::new(slide_name, ct::PML_SLIDE.into(), slide_xml(linked));
    slide.rels_mut().add_relationship(
        rt::OLE_OBJECT.into(),
        if linked {
            "https://example.invalid/source"
        } else {
            "../embeddings/oleObject1.bin"
        }
        .into(),
        "rIdOle".into(),
        linked,
    );
    if !linked {
        slide.rels_mut().add_relationship(
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image".into(),
            "../media/image1.png".into(),
            "rIdPreview".into(),
            false,
        );
    }
    let mut package = OpcPackage::new();
    package.add_part(Box::new(slide));
    if !linked {
        package.add_part(Box::new(BlobPart::new(
            PackURI::new(PAYLOAD).unwrap(),
            ct::OFC_OLE_OBJECT.into(),
            b"opaque-ole-payload".to_vec(),
        )));
        package.add_part(Box::new(BlobPart::new(
            PackURI::new(PREVIEW).unwrap(),
            "image/png".into(),
            b"opaque-preview".to_vec(),
        )));
    }
    package
}

fn snapshot(package: &OpcPackage) -> Snapshot {
    Snapshot::load(package, 0, &PackURI::new(SLIDE).unwrap()).unwrap()
}

#[test]
fn exact_noop_is_byte_identical() {
    let mut package = package(false);
    let before = snapshot(&package);
    let source = before.source_xml().to_vec();
    let commit = before.edit().commit().unwrap();
    assert!(!commit.is_changed());
    assert!(commit.patch().is_empty());
    apply_patch(&mut package, commit.patch()).unwrap();
    assert_eq!(snapshot(&package).source_xml(), source.as_slice());
    assert_eq!(
        snapshot(&package).payload(0).unwrap(),
        Some(&b"opaque-ole-payload"[..])
    );
}

#[test]
fn typed_edits_preserve_opaque_xml_and_binary_parts() {
    let mut package = package(false);
    let before = snapshot(&package);
    let mut edit = before.edit();
    edit.set_name(0, Some("Edited & safe".into())).unwrap();
    edit.set_program_id(0, Some("Word.Document.12".into()))
        .unwrap();
    edit.set_show_as_icon(0, Some(true)).unwrap();
    edit.set_preview_size(0, Some((90, 100))).unwrap();
    edit.set_anchor(0, Frame::new(-10, 20, 500, 600)).unwrap();
    edit.replace_payload(0, b"replaced-opaque-payload".to_vec())
        .unwrap();
    let commit = edit.commit().unwrap();
    let xml = std::str::from_utf8(commit.snapshot().source_xml()).unwrap();
    assert!(xml.contains("<x:opaque>keep</x:opaque>"));
    assert!(xml.contains("name=\"Edited &amp; safe\""));
    assert!(xml.contains("x=\"-10\""));
    apply_patch(&mut package, commit.patch()).unwrap();
    let after = snapshot(&package);
    assert_eq!(after.objects()[0].name(), Some("Edited & safe"));
    assert_eq!(after.objects()[0].program_id(), Some("Word.Document.12"));
    assert_eq!(after.objects()[0].show_as_icon(), Some(true));
    assert_eq!(
        after.objects()[0].anchor(),
        Some(Frame::new(-10, 20, 500, 600))
    );
    assert_eq!(
        after.payload(0).unwrap(),
        Some(&b"replaced-opaque-payload"[..])
    );
    assert_eq!(
        package
            .get_part(&PackURI::new(PREVIEW).unwrap())
            .unwrap()
            .blob(),
        b"opaque-preview"
    );
}

#[test]
fn add_remove_and_inverse_manage_payload_ownership() {
    let mut package = package_without_object();
    let before = snapshot(&package);
    let mut edit = before.edit();
    let index = edit
        .add(
            Definition::embedded(Kind::Package, Frame::new(1, 2, 300, 400), b"new-package")
                .set_name("Package"),
        )
        .unwrap();
    assert_eq!(index, 0);
    let commit = edit.commit().unwrap();
    apply_patch(&mut package, commit.patch()).unwrap();
    let added = snapshot(&package);
    assert_eq!(added.objects().len(), 1);
    assert_eq!(added.objects()[0].mode(), Mode::Embedded);
    assert_eq!(added.objects()[0].kind(), Some(Kind::Package));
    let payload_name = added.objects()[0]
        .target()
        .unwrap()
        .part_name()
        .unwrap()
        .clone();
    assert_eq!(added.payload(0).unwrap(), Some(&b"new-package"[..]));

    let mut remove = added.edit();
    remove.detach(0).unwrap();
    let removal = remove.commit().unwrap();
    apply_patch(&mut package, removal.patch()).unwrap();
    assert!(snapshot(&package).objects().is_empty());
    assert!(package.get_part(&payload_name).is_err());
    apply_patch(&mut package, &removal.patch().inverse()).unwrap();
    assert_eq!(
        snapshot(&package).payload(0).unwrap(),
        Some(&b"new-package"[..])
    );
}

#[test]
fn linked_target_edits_stay_external_and_inert() {
    let mut package = package(true);
    let before = snapshot(&package);
    assert!(matches!(
        before.objects()[0].target(),
        Some(Target::External { target, .. }) if target == "https://example.invalid/source"
    ));
    let mut edit = before.edit();
    edit.set_link_target(0, "https://example.invalid/changed")
        .unwrap();
    let patch = edit.commit().unwrap().into_patch();
    apply_patch(&mut package, &patch).unwrap();
    let current = snapshot(&package);
    let xml = std::str::from_utf8(current.source_xml()).unwrap();
    assert!(xml.contains("rIdIgnored"));
    assert!(xml.contains("<mc:Fallback>"));
    assert!(matches!(
        snapshot(&package).objects()[0].target(),
        Some(Target::External { target, .. }) if target == "https://example.invalid/changed"
    ));
    assert!(package.get_part(&PackURI::new(PAYLOAD).unwrap()).is_err());
}

#[test]
fn stale_and_invalid_edits_are_failure_atomic() {
    let mut stale_package = package(false);
    let before = snapshot(&stale_package);
    let mut edit = before.edit();
    edit.set_name(0, Some("new".into())).unwrap();
    let patch = edit.commit().unwrap().into_patch();
    let slide = PackURI::new(SLIDE).unwrap();
    stale_package
        .get_part_mut(&slide)
        .unwrap()
        .set_blob(b"stale".to_vec());
    assert!(apply_patch(&mut stale_package, &patch).is_err());
    assert_eq!(stale_package.get_part(&slide).unwrap().blob(), b"stale");

    let clean_package = package(false);
    let source = snapshot(&clean_package);
    let mut invalid = source.edit();
    assert!(invalid.replace_payload(0, Vec::<u8>::new()).is_err());
    assert_eq!(source.source_xml(), snapshot(&clean_package).source_xml());
}

fn package_without_object() -> OpcPackage {
    let mut package = package(false);
    let slide = PackURI::new(SLIDE).unwrap();
    let xml = std::str::from_utf8(package.get_part(&slide).unwrap().blob())
        .unwrap()
        .replace(
            "<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id=\"101\" name=\"Frame\"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x=\"10\" y=\"20\"/><a:ext cx=\"300\" cy=\"400\"/></p:xfrm><a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/presentationml/2006/ole\"><p:oleObj name=\"Old\" progId=\"Old.App\" showAsIcon=\"0\" imgW=\"30\" imgH=\"40\" r:id=\"rIdOle\"><x:opaque>keep</x:opaque><p:embed/><p:pic><p:blipFill><a:blip r:embed=\"rIdPreview\"/></p:blipFill></p:pic></p:oleObj></a:graphicData></a:graphic></p:graphicFrame>",
            "",
        );
    package
        .get_part_mut(&slide)
        .unwrap()
        .set_blob(xml.into_bytes());
    package
        .get_part_mut(&slide)
        .unwrap()
        .rels_mut()
        .remove("rIdOle");
    package
        .get_part_mut(&slide)
        .unwrap()
        .rels_mut()
        .remove("rIdPreview");
    package.remove_part(&PackURI::new(PAYLOAD).unwrap());
    package.remove_part(&PackURI::new(PREVIEW).unwrap());
    package
}
