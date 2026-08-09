#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

use super::codec::{add_anchor, load, put, remove, remove_anchor, scan_layout, selected_raw_span};
use super::validation::P14;
use crate::tag::Tag;
use crate::tag::{CONTENT_TYPE, Conformance, List};
use litchi_opc::{OpcPackage, PackURI, XmlPart};
use std::sync::Arc;

const PML: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const STRICT_PML: &str = "http://purl.oclc.org/ooxml/presentationml/main";
const STRICT_REL: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships";
const MC: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";

fn list(name: &str, value: &str) -> List {
    let mut list = List::new();
    list.add(Tag::new(name, value).expect("valid tag"))
        .expect("unique tag");
    list
}

fn owner_package(xml: Vec<u8>) -> (OpcPackage, PackURI) {
    let owner = PackURI::new("/ppt/slides/slide1.xml").expect("owner URI");
    let mut package = OpcPackage::new();
    package.add_part(Box::new(XmlPart::new(
        owner.clone(),
        "application/vnd.openxmlformats-officedocument.presentationml.slide+xml".into(),
        xml,
    )));
    (package, owner)
}

fn shape_xml(name: &str, id: u32) -> String {
    format!(
        r#"<p:sp><p:nvSpPr><p:cNvPr id="{id}" name="{name}"/><p:cNvSpPr/><p:nvPr><p:ph/></p:nvPr></p:nvSpPr><p:spPr/></p:sp>"#
    )
}

fn mce_shared_anchor_package() -> (OpcPackage, PackURI, PackURI) {
    let anchor = r#"<p:custDataLst><p:tags r:id="rId1"/></p:custDataLst>"#;
    let active = shape_xml("Active", 2).replace("<p:ph/>", anchor);
    let inactive = shape_xml("Inactive", 3).replace("<p:ph/>", anchor);
    let xml = format!(
            r#"<p:sld xmlns:p="{PML}" xmlns:r="{REL}" xmlns:mc="{MC}" xmlns:p14="{P14}"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/><mc:AlternateContent><mc:Choice Requires="p14">{active}</mc:Choice><mc:Fallback>{inactive}</mc:Fallback></mc:AlternateContent></p:spTree></p:cSld></p:sld>"#
        )
        .into_bytes();
    let (mut package, owner) = owner_package(xml);
    package
        .get_part_mut(&owner)
        .expect("owner")
        .rels_mut()
        .add_relationship(
            crate::tag::TAG_REL.into(),
            "../tags/tag1.xml".into(),
            "rId1".into(),
            false,
        );
    let part = PackURI::new("/ppt/tags/tag1.xml").expect("part URI");
    package.add_part(Box::new(XmlPart::new(
        part.clone(),
        CONTENT_TYPE.into(),
        format!(r#"<p:tagLst xmlns:p="{PML}"><p:tag name="Owner" val="Alice"/></p:tagLst>"#)
            .into_bytes(),
    )));
    (package, owner, part)
}

#[test]
fn maps_all_five_families_and_nested_groups_to_raw_source() {
    let xml = format!(
            r#"<p:sld xmlns:p="{PML}" xmlns:r="{REL}"><p:cSld><p:spTree>
                <p:nvGrpSpPr/><p:grpSpPr/>
                {}
                <p:pic><p:nvPicPr><p:cNvPr id="3" name="Picture"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill/><p:spPr/></p:pic>
                <p:cxnSp><p:nvCxnSpPr><p:cNvPr id="4" name="Connector"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr><p:spPr/></p:cxnSp>
                <p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="5" name="Frame"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm/></p:graphicFrame>
                <p:grpSp><p:nvGrpSpPr><p:cNvPr id="6" name="Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>{}</p:grpSp>
            </p:spTree></p:cSld></p:sld>"#,
            shape_xml("Auto", 2),
            shape_xml("Nested", 7),
        )
        .into_bytes();

    let expected = [
        ("Auto", b"<p:sp>".as_slice()),
        ("Picture", b"<p:pic>".as_slice()),
        ("Connector", b"<p:cxnSp>".as_slice()),
        ("Frame", b"<p:graphicFrame>".as_slice()),
        ("Group", b"<p:grpSp>".as_slice()),
        ("Nested", b"<p:sp>".as_slice()),
    ];
    for (index, (name, opening)) in expected.iter().enumerate() {
        let by_name =
            selected_raw_span(&xml, crate::shape::Key::Name(name)).expect("name maps to raw span");
        let by_index = selected_raw_span(&xml, crate::shape::Key::Index(index))
            .expect("index maps to raw span");
        assert_eq!(by_name, by_index);
        assert!(xml[by_name].starts_with(opening));

        let layout = scan_layout(&xml, by_index).expect("shape layout");
        let staged = add_anchor(&xml, &layout, "rId9").expect("anchor insertion");
        let staged_span =
            selected_raw_span(&staged, crate::shape::Key::Name(name)).expect("staged shape maps");
        let staged_layout = scan_layout(&staged, staged_span).expect("staged layout");
        assert_eq!(
            staged_layout
                .anchor
                .as_ref()
                .map(|anchor| anchor.id.as_str()),
            Some("rId9")
        );
        let removed = remove_anchor(&staged, &staged_layout).expect("anchor removal");
        let removed_span =
            selected_raw_span(&removed, crate::shape::Key::Name(name)).expect("removed shape maps");
        assert!(
            scan_layout(&removed, removed_span)
                .expect("removed layout")
                .anchor
                .is_none()
        );
    }
}

#[test]
fn preserves_customer_data_and_inserts_before_extensions() {
    let xml = format!(
            r#"<p:sld xmlns:p="{PML}" xmlns:r="{REL}"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>
            <p:sp><p:nvSpPr><p:cNvPr id="2" name="Ordered"/><p:cNvSpPr/><p:nvPr><p:ph/><p:audioFile r:link="rIdAudio"/><p:custDataLst keep="yes"><p:custData r:id="rIdData"/></p:custDataLst><!--keep--><p:extLst/></p:nvPr></p:nvSpPr><p:spPr/></p:sp>
            </p:spTree></p:cSld></p:sld>"#
        )
        .into_bytes();
    let span = selected_raw_span(&xml, crate::shape::Key::Name("Ordered")).expect("shape");
    let layout = scan_layout(&xml, span).expect("layout");
    let staged = add_anchor(&xml, &layout, "rIdTags").expect("insert");
    let text = std::str::from_utf8(&staged).expect("UTF-8 fixture");
    assert!(text.find("<p:custData ").unwrap() < text.find("<p:tags ").unwrap());
    assert!(text.find("<p:tags ").unwrap() < text.find("</p:custDataLst>").unwrap());
    assert!(text.find("</p:custDataLst>").unwrap() < text.find("<!--keep-->").unwrap());
    assert!(text.find("<!--keep-->").unwrap() < text.find("<p:extLst").unwrap());

    let staged_span =
        selected_raw_span(&staged, crate::shape::Key::Name("Ordered")).expect("shape");
    let staged_layout = scan_layout(&staged, staged_span).expect("layout");
    let removed = remove_anchor(&staged, &staged_layout).expect("remove");
    assert_eq!(removed, xml);
}

#[test]
fn mce_mapping_edits_only_the_active_raw_branch() {
    let p14 = P14;
    let xml = format!(
            r#"<p:sld xmlns:p="{PML}" xmlns:r="{REL}" xmlns:mc="{MC}" xmlns:p14="{p14}"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>
            <mc:AlternateContent><mc:Choice Requires="p14">{}</mc:Choice><mc:Fallback><p:pic><p:nvPicPr><p:cNvPr id="3" name="Inactive"/><p:cNvPicPr/><p:nvPr><p:custDataLst><p:tags r:id="rIdInactive"/></p:custDataLst></p:nvPr></p:nvPicPr><p:blipFill/><p:spPr/></p:pic></mc:Fallback></mc:AlternateContent>
            {}
            </p:spTree></p:cSld></p:sld>"#,
            shape_xml("Active", 2),
            shape_xml("Second", 4),
        )
        .into_bytes();

    let active_span = selected_raw_span(&xml, crate::shape::Key::Index(0)).expect("active span");
    assert!(
        xml[active_span.clone()]
            .windows(13)
            .any(|bytes| bytes == b"name=\"Active\"")
    );
    assert!(
        !xml[active_span.clone()]
            .windows(15)
            .any(|bytes| bytes == b"name=\"Inactive\"")
    );
    let layout = scan_layout(&xml, active_span).expect("active layout");
    let staged = add_anchor(&xml, &layout, "rIdTags").expect("active insertion");
    let text = std::str::from_utf8(&staged).expect("UTF-8 fixture");
    assert_eq!(text.matches("<p:tags ").count(), 2);
    assert!(text.contains(r#"<p:tags r:id="rIdInactive"/>"#));
    let inactive = text.find("name=\"Inactive\"").expect("inactive branch");
    let inactive_end = text[inactive..].find("</p:pic>").expect("picture end") + inactive;
    assert!(!text[inactive..inactive_end].contains("rIdTags"));
}

#[test]
fn inactive_mce_anchor_forces_replacement_fork_and_removal_retention() {
    let (mut package, owner, original_part) = mce_shared_anchor_package();

    let old = put(&mut package, &owner, "Active", list("Reviewer", "Bob"))
        .expect("replace active attachment")
        .expect("old list");
    assert_eq!(old.get("owner").expect("old tag").value(), "Alice");
    let active = load(&package, &owner, "Active")
        .expect("load active attachment")
        .expect("active attachment");
    assert_ne!(active.rel(), "rId1");
    assert_ne!(active.part(), &original_part);
    assert_eq!(
        active.list().get("reviewer").expect("new tag").value(),
        "Bob"
    );
    let owner_part = package.get_part(&owner).expect("owner");
    assert!(owner_part.rels().get("rId1").is_some());
    assert!(package.get_part(&original_part).is_ok());
    let owner_xml = std::str::from_utf8(owner_part.blob()).expect("UTF-8 fixture");
    assert_eq!(owner_xml.matches(r#"r:id="rId1""#).count(), 1);
    let inactive = owner_xml
        .find("name=\"Inactive\"")
        .expect("inactive branch");
    let inactive_end = owner_xml[inactive..]
        .find("</p:sp>")
        .expect("inactive shape end")
        + inactive;
    assert!(owner_xml[inactive..inactive_end].contains(r#"r:id="rId1""#));

    let (mut package, owner, original_part) = mce_shared_anchor_package();
    let removed = remove(&mut package, &owner, "Active")
        .expect("remove active attachment")
        .expect("old list");
    assert_eq!(removed.get("owner").expect("old tag").value(), "Alice");
    assert!(
        load(&package, &owner, "Active")
            .expect("load active")
            .is_none()
    );
    let owner_part = package.get_part(&owner).expect("owner");
    assert!(owner_part.rels().get("rId1").is_some());
    assert!(package.get_part(&original_part).is_ok());
    let owner_xml = std::str::from_utf8(owner_part.blob()).expect("UTF-8 fixture");
    assert_eq!(owner_xml.matches(r#"r:id="rId1""#).count(), 1);
    let inactive = owner_xml
        .find("name=\"Inactive\"")
        .expect("inactive branch");
    let inactive_end = owner_xml[inactive..]
        .find("</p:sp>")
        .expect("inactive shape end")
        + inactive;
    assert!(owner_xml[inactive..inactive_end].contains(r#"r:id="rId1""#));
}

#[test]
fn shape_crud_is_atomic_noop_safe_and_move_based() {
    let xml = format!(
            r#"<p:sld xmlns:p="{PML}" xmlns:r="{REL}"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>{}</p:spTree></p:cSld></p:sld>"#,
            shape_xml("Title", 2),
        )
        .into_bytes();
    let (mut package, owner) = owner_package(xml);

    assert!(load(&package, &owner, "Title").expect("load").is_none());
    assert_eq!(
        put(&mut package, &owner, "Title", list("Owner", "Alice")).expect("create"),
        None
    );
    let source = load(&package, &owner, 0_usize)
        .expect("load")
        .expect("attachment");
    assert_eq!(source.list().get("owner").expect("tag").value(), "Alice");

    let owner_before = package.get_part(&owner).expect("owner").blob_arc();
    let part_before = package
        .get_part(source.part())
        .expect("tag part")
        .blob_arc();
    package.relate_to(
        "_xmlsignatures/origin.sigs",
        litchi_opc::constants::relationship_type::DIGITAL_SIGNATURE_ORIGIN,
    );
    assert!(package.is_signed());
    let old = put(&mut package, &owner, "Title", list("Owner", "Alice"))
        .expect("no-op")
        .expect("old list");
    assert_eq!(old.get("OWNER").expect("old tag").value(), "Alice");
    assert!(package.is_signed());
    assert!(Arc::ptr_eq(
        &owner_before,
        &package.get_part(&owner).expect("owner").blob_arc()
    ));
    assert!(Arc::ptr_eq(
        &part_before,
        &package
            .get_part(source.part())
            .expect("tag part")
            .blob_arc()
    ));

    assert!(put(&mut package, &owner, "Missing", List::new()).is_err());
    assert!(package.is_signed());
    assert!(Arc::ptr_eq(
        &owner_before,
        &package.get_part(&owner).expect("owner").blob_arc()
    ));

    let removed = remove(&mut package, &owner, "Title")
        .expect("remove")
        .expect("old list");
    assert_eq!(removed.get("owner").expect("tag").value(), "Alice");
    assert!(!package.is_signed());
    assert!(load(&package, &owner, "Title").expect("load").is_none());
    assert!(
        remove(&mut package, &owner, "Title")
            .expect("idempotent remove")
            .is_none()
    );
}

#[test]
fn strict_shape_crud_uses_strict_namespaces_and_relationship_type() {
    let xml = format!(
            r#"<p:sld xmlns:p="{STRICT_PML}" xmlns:r="{STRICT_REL}"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>{}</p:spTree></p:cSld></p:sld>"#,
            shape_xml("Strict", 2),
        )
        .into_bytes();
    let (mut package, owner) = owner_package(xml);
    assert_eq!(
        put(&mut package, &owner, "Strict", List::new()).expect("strict create"),
        None
    );
    let source = load(&package, &owner, "Strict")
        .expect("strict load")
        .expect("strict attachment");
    assert_eq!(source.conformance(), Conformance::Strict);
    assert_eq!(
        package
            .get_part(&owner)
            .expect("owner")
            .rels()
            .get(source.rel())
            .expect("relationship")
            .reltype(),
        crate::tag::STRICT_TAG_REL
    );
    let owner_xml = std::str::from_utf8(package.get_part(&owner).expect("owner").blob())
        .expect("UTF-8 fixture");
    assert!(owner_xml.contains(STRICT_PML));
    assert!(owner_xml.contains(STRICT_REL));
    let part_xml = std::str::from_utf8(package.get_part(source.part()).expect("tag part").blob())
        .expect("UTF-8 tag part");
    assert!(part_xml.contains(STRICT_PML));
}

#[test]
fn shared_shape_anchor_forks_then_collects_each_orphan() {
    let anchor = r#"<p:custDataLst><p:tags r:id="rId1"/></p:custDataLst>"#;
    let first = shape_xml("First", 2).replace("<p:ph/>", anchor);
    let second = shape_xml("Second", 3).replace("<p:ph/>", anchor);
    let xml = format!(
            r#"<p:sld xmlns:p="{PML}" xmlns:r="{REL}"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>{first}{second}</p:spTree></p:cSld></p:sld>"#
        )
        .into_bytes();
    let (mut package, owner) = owner_package(xml);
    package
        .get_part_mut(&owner)
        .expect("owner")
        .rels_mut()
        .add_relationship(
            crate::tag::TAG_REL.into(),
            "../tags/tag1.xml".into(),
            "rId1".into(),
            false,
        );
    let original_part = PackURI::new("/ppt/tags/tag1.xml").expect("part URI");
    package.add_part(Box::new(XmlPart::new(
        original_part.clone(),
        CONTENT_TYPE.into(),
        format!(r#"<p:tagLst xmlns:p="{PML}"><p:tag name="Owner" val="Alice"/></p:tagLst>"#)
            .into_bytes(),
    )));

    let old = put(&mut package, &owner, "First", list("Reviewer", "Bob"))
        .expect("fork")
        .expect("old list");
    assert_eq!(old.get("owner").expect("old tag").value(), "Alice");
    let first = load(&package, &owner, "First")
        .expect("first load")
        .expect("first attachment");
    let second = load(&package, &owner, "Second")
        .expect("second load")
        .expect("second attachment");
    assert_ne!(first.rel(), second.rel());
    assert_ne!(first.part(), second.part());
    assert_eq!(
        first.list().get("reviewer").expect("new tag").value(),
        "Bob"
    );
    assert_eq!(
        second.list().get("owner").expect("old tag").value(),
        "Alice"
    );

    let forked_part = first.part().clone();
    assert!(
        remove(&mut package, &owner, "First")
            .expect("remove first")
            .is_some()
    );
    assert!(package.get_part(&forked_part).is_err());
    assert!(package.get_part(&original_part).is_ok());
    assert!(
        remove(&mut package, &owner, "Second")
            .expect("remove second")
            .is_some()
    );
    assert!(package.get_part(&original_part).is_err());
}
