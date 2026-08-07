use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{BlobPart, OpcPackage, PackURI, Part};

use super::*;

const P: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const P202: &str = "http://schemas.microsoft.com/office/powerpoint/2020/02/main";
const URI: &str = "{E3EDB536-0D56-4F60-86BA-61A60CA02DAB}";

fn package(hosts: &str) -> OpcPackage {
    let mut package = OpcPackage::new();
    let presentation_name = PackURI::new("/ppt/presentation.xml").unwrap();
    let mut presentation = BlobPart::new(
        presentation_name,
        ct::PML_PRESENTATION_MAIN.into(),
        format!(
            "<p:presentation xmlns:p=\"{P}\" xmlns:r=\"{R}\" xmlns:p202=\"{P202}\"><p:sldIdLst>{hosts}</p:sldIdLst><p:extLst><p:ext uri=\"urn:opaque\"><x:v xmlns:x=\"urn:x\"/></p:ext></p:extLst></p:presentation>"
        )
        .into_bytes(),
    );
    let first = presentation.relate_to("slides/slide1.xml", rt::SLIDE);
    let second = presentation.relate_to("slides/slide2.xml", rt::SLIDE);
    assert_eq!(first, "rId1");
    assert_eq!(second, "rId2");
    package.add_part(Box::new(presentation));
    for index in 1..=2 {
        package.add_part(Box::new(BlobPart::new(
            PackURI::new(&format!("/ppt/slides/slide{index}.xml")).unwrap(),
            ct::PML_SLIDE.into(),
            format!("<p:sld xmlns:p=\"{P}\"/>").into_bytes(),
        )));
    }
    package.relate_to("ppt/presentation.xml", rt::OFFICE_DOCUMENT);
    package
}

fn tag(name: &str, value: &str) -> Tag {
    Tag::new(name, value).unwrap()
}

#[test]
fn inventories_absent_empty_duplicates_and_order() {
    let hosts = format!(
        "<p:sldId id=\"256\" r:id=\"rId1\"><p:extLst><!--keep--><p:ext uri=\"{URI}\"><p202:designTagLst/></p:ext><p:ext uri=\"{URI}\"><p202:designTagLst><p202:designTag name=\"n\" val=\"1\"/><p202:designTag name=\"n\" val=\"2\"/></p202:designTagLst></p:ext></p:extLst></p:sldId><p:sldId id=\"257\" r:id=\"rId2\"/>"
    );
    let package = package(&hosts);
    let duplicate = load_snapshot(&package, 256).unwrap();
    assert_eq!(duplicate.occurrence_count(), 2);
    assert!(duplicate.occurrences().next().unwrap().is_empty());
    assert_eq!(
        duplicate.occurrences().nth(1).unwrap().as_slice()[1].value(),
        "2"
    );
    assert!(duplicate.tags().is_err());
    assert!(duplicate.edit().is_err());
    assert!(
        load_snapshot(&package, 257)
            .unwrap()
            .tags()
            .unwrap()
            .is_none()
    );
}

#[test]
fn inserts_replaces_and_removes_without_touching_opaque_bytes() {
    let hosts =
        "<p:sldId data=\"keep\" id=\"256\" r:id=\"rId1\"/><p:sldId id=\"257\" r:id=\"rId2\"/>";
    let mut package = package(hosts);
    let mut tags = Tags::new();
    tags.push(tag("same", "1")).unwrap();
    tags.push(tag("same", "2")).unwrap();
    assert!(store(&mut package, 256, &tags).unwrap().is_none());
    let xml = package.main_document_part().unwrap().blob();
    assert!(String::from_utf8_lossy(xml).contains("data=\"keep\""));
    assert!(String::from_utf8_lossy(xml).contains("urn:opaque"));
    assert_eq!(load(&package, 256).unwrap().unwrap(), tags);

    let before_noop = package.main_document_part().unwrap().blob_arc();
    package.relate_to("_xmlsignatures/origin.sigs", rt::DIGITAL_SIGNATURE_ORIGIN);
    assert!(package.is_signed());
    store(&mut package, 256, &tags).unwrap();
    assert!(package.is_signed());
    assert_eq!(
        package.main_document_part().unwrap().blob_arc(),
        before_noop
    );

    assert_eq!(remove(&mut package, 256).unwrap(), Some(tags));
    assert!(!package.is_signed());
    let xml = String::from_utf8_lossy(package.main_document_part().unwrap().blob());
    assert!(xml.contains("<p:sldId data=\"keep\" id=\"256\" r:id=\"rId1\"/>"));
    assert!(xml.contains("urn:opaque"));
}

#[test]
fn patch_allows_reorder_and_inverse_but_refuses_stale_or_rebound_host() {
    let hosts = "<p:sldId id=\"256\" r:id=\"rId1\"/><p:sldId id=\"257\" r:id=\"rId2\"/>";
    let mut package = package(hosts);
    let snapshot = load_snapshot(&package, 256).unwrap();
    let mut edit = snapshot.edit().unwrap();
    let mut tags = Tags::new();
    tags.push(tag("k", "v")).unwrap();
    edit.set(tags.clone()).unwrap();
    let patch = edit.commit().unwrap().into_patch();

    let name = PackURI::new("/ppt/presentation.xml").unwrap();
    let reordered = String::from_utf8(package.get_part(&name).unwrap().blob().to_vec())
        .unwrap()
        .replace(
            "<p:sldId id=\"256\" r:id=\"rId1\"/><p:sldId id=\"257\" r:id=\"rId2\"/>",
            "<p:sldId id=\"257\" r:id=\"rId2\"/><p:sldId id=\"256\" r:id=\"rId1\"/>",
        );
    package
        .get_part_mut(&name)
        .unwrap()
        .set_blob(reordered.into_bytes());
    patch.apply(&mut package).unwrap();
    assert_eq!(load(&package, 256).unwrap(), Some(tags));
    patch.inverse().apply(&mut package).unwrap();
    assert!(load(&package, 256).unwrap().is_none());

    let mut stale = String::from_utf8(package.get_part(&name).unwrap().blob().to_vec()).unwrap();
    stale = stale.replace("id=\"256\" r:id=\"rId1\"/>", "id=\"256\" r:id=\"rId2\"/>");
    package
        .get_part_mut(&name)
        .unwrap()
        .set_blob(stale.into_bytes());
    assert!(patch.apply(&mut package).is_err());
}

#[test]
fn rejects_spoofing_forbidden_markup_bad_graph_and_limits() {
    let spoof = format!(
        "<p:sldId id=\"256\" r:id=\"rId1\"><p:extLst><p:ext uri=\"{URI}\"><evil:designTagLst xmlns:evil=\"urn:evil\"/></p:ext></p:extLst></p:sldId><p:sldId id=\"257\" r:id=\"rId2\"/>"
    );
    assert!(load_snapshot(&package(&spoof), 256).is_err());

    let duplicate = "<p:sldId id=\"256\" r:id=\"rId1\"/><p:sldId id=\"256\" r:id=\"rId2\"/>";
    assert!(load_snapshot(&package(duplicate), 256).is_err());

    let dtd = package("<p:sldId id=\"256\" r:id=\"rId1\"/><p:sldId id=\"257\" r:id=\"rId2\"/>");
    let name = PackURI::new("/ppt/presentation.xml").unwrap();
    let xml = dtd.get_part(&name).unwrap().blob();
    let mut with_dtd = b"<!DOCTYPE p:presentation>".to_vec();
    with_dtd.extend_from_slice(xml);
    let mut dtd = dtd;
    dtd.get_part_mut(&name).unwrap().set_blob(with_dtd);
    assert!(load_snapshot(&dtd, 256).is_err());

    let exact = package("<p:sldId id=\"256\" r:id=\"rId1\"/><p:sldId id=\"257\" r:id=\"rId2\"/>");
    let bytes = exact.main_document_part().unwrap().blob().len();
    assert!(
        load_snapshot_with_limits(&exact, 256, Limits::default().with_xml_bytes(bytes)).is_ok()
    );
    assert!(
        load_snapshot_with_limits(&exact, 256, Limits::default().with_xml_bytes(bytes - 1))
            .is_err()
    );
    assert!(load_snapshot_with_limits(&exact, 256, Limits::default().with_xml_nodes(1)).is_err());
}

#[test]
fn accepts_inherited_default_and_alias_bindings_but_rejects_local_rebinding() {
    let aliases = format!(
        "<p:sldId id=\"256\" r:id=\"rId1\"><p:extLst><p:ext uri=\"{URI}\"><a:designTagLst><b:designTag name=\"n\" val=\"v\"/></a:designTagLst></p:ext></p:extLst></p:sldId><p:sldId id=\"257\" r:id=\"rId2\"/>"
    );
    let mut aliases_package = package(&aliases);
    let name = PackURI::new("/ppt/presentation.xml").unwrap();
    let xml = String::from_utf8(aliases_package.get_part(&name).unwrap().blob().to_vec())
        .unwrap()
        .replace(
            &format!("xmlns:p202=\"{P202}\""),
            &format!("xmlns:p202=\"{P202}\" xmlns:a=\"{P202}\" xmlns:b=\"{P202}\""),
        );
    aliases_package
        .get_part_mut(&name)
        .unwrap()
        .set_blob(xml.into_bytes());
    assert_eq!(
        load(&aliases_package, 256).unwrap().unwrap().as_slice()[0].value(),
        "v"
    );

    let default = format!(
        "<p:sldId id=\"256\" r:id=\"rId1\"><p:extLst><p:ext uri=\"{URI}\"><designTagLst><designTag name=\"n\" val=\"d\"/></designTagLst></p:ext></p:extLst></p:sldId><p:sldId id=\"257\" r:id=\"rId2\"/>"
    );
    let mut default_package = package(&default);
    let xml = String::from_utf8(default_package.get_part(&name).unwrap().blob().to_vec())
        .unwrap()
        .replace(
            &format!("xmlns:p202=\"{P202}\""),
            &format!("xmlns:p202=\"{P202}\" xmlns=\"{P202}\""),
        );
    default_package
        .get_part_mut(&name)
        .unwrap()
        .set_blob(xml.into_bytes());
    assert_eq!(
        load(&default_package, 256).unwrap().unwrap().as_slice()[0].value(),
        "d"
    );

    let rebound = aliases.replace("<b:designTag", "<b:designTag xmlns:b=\"urn:evil\"");
    let mut rebound_package = package(&rebound);
    let xml = String::from_utf8(rebound_package.get_part(&name).unwrap().blob().to_vec())
        .unwrap()
        .replace(
            &format!("xmlns:p202=\"{P202}\""),
            &format!("xmlns:p202=\"{P202}\" xmlns:a=\"{P202}\" xmlns:b=\"{P202}\""),
        );
    rebound_package
        .get_part_mut(&name)
        .unwrap()
        .set_blob(xml.into_bytes());
    assert!(load(&rebound_package, 256).is_err());
}

#[test]
fn duplicate_mutations_and_deleted_host_are_atomic_and_signatures_follow_changes() {
    let duplicate = format!(
        "<p:sldId id=\"256\" r:id=\"rId1\"><p:extLst><p:ext uri=\"{URI}\"><p202:designTagLst/></p:ext><p:ext uri=\"{URI}\"><p202:designTagLst/></p:ext></p:extLst></p:sldId><p:sldId id=\"257\" r:id=\"rId2\"/>"
    );
    let mut duplicate_package = package(&duplicate);
    duplicate_package.relate_to("_xmlsignatures/origin.sigs", rt::DIGITAL_SIGNATURE_ORIGIN);
    let before = duplicate_package.main_document_part().unwrap().blob_arc();
    assert!(store(&mut duplicate_package, 256, &Tags::new()).is_err());
    assert!(remove(&mut duplicate_package, 256).is_err());
    assert!(duplicate_package.is_signed());
    assert_eq!(
        duplicate_package.main_document_part().unwrap().blob_arc(),
        before
    );

    let hosts = "<p:sldId id=\"256\" r:id=\"rId1\"/><p:sldId id=\"257\" r:id=\"rId2\"/>";
    let mut package = package(hosts);
    let snapshot = load_snapshot(&package, 256).unwrap();
    let mut edit = snapshot.edit().unwrap();
    edit.set(Tags::new()).unwrap();
    let patch = edit.commit().unwrap().into_patch();
    let name = PackURI::new("/ppt/presentation.xml").unwrap();
    let xml = String::from_utf8(package.get_part(&name).unwrap().blob().to_vec())
        .unwrap()
        .replace("<p:sldId id=\"256\" r:id=\"rId1\"/>", "");
    package
        .get_part_mut(&name)
        .unwrap()
        .set_blob(xml.into_bytes());
    assert!(patch.apply(&mut package).is_err());
}

#[test]
fn store_validates_default_limits_before_mutating() {
    let hosts = "<p:sldId id=\"256\" r:id=\"rId1\"/><p:sldId id=\"257\" r:id=\"rId2\"/>";
    let mut package = package(hosts);
    let before = package.main_document_part().unwrap().blob_arc();
    let oversized = Tag::new_with_limits(
        "x".repeat(Limits::default().string_bytes() + 1),
        "v",
        Limits::default().with_string_bytes(Limits::default().string_bytes() + 1),
    )
    .unwrap();
    let mut tags = Tags::new();
    tags.push_with_limits(
        oversized,
        Limits::default().with_string_bytes(Limits::default().string_bytes() + 1),
    )
    .unwrap();
    assert!(store(&mut package, 256, &tags).is_err());
    assert_eq!(package.main_document_part().unwrap().blob_arc(), before);
}
