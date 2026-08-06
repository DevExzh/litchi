use super::*;

const PML: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
const DML: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const MC: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const MAIN: &str = "http://schemas.microsoft.com/office/powerpoint/2016/6/main";
const SECTION: &str = "http://schemas.microsoft.com/office/powerpoint/2016/sectionzoom";
const SLIDE: &str = "http://schemas.microsoft.com/office/powerpoint/2016/slidezoom";
const SUMMARY: &str = "http://schemas.microsoft.com/office/powerpoint/2016/summaryzoom";

const SECTION_ID: &str = "{11111111-1111-1111-1111-111111111111}";
const SLIDE_OBJECT_ID: &str = "{22222222-2222-2222-2222-222222222222}";
const SUMMARY_OBJECT_ID: &str = "{33333333-3333-3333-3333-333333333333}";

fn blip_fill(relationship: Option<&str>) -> String {
    let blip = relationship
        .map(|id| format!(" r:embed=\"{id}\""))
        .unwrap_or_default();
    format!(r#"<p166:blipFill><a:blip{blip}/><a:stretch><a:fillRect/></a:stretch></p166:blipFill>"#)
}

fn shape_properties() -> &'static str {
    r#"<p166:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1" cy="1"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p166:spPr>"#
}

fn properties(id: &str, relationship: Option<&str>, attributes: &str) -> String {
    format!(
        r#"<p166:zmPr id="{id}"{attributes}>{}{}</p166:zmPr>"#,
        blip_fill(relationship),
        shape_properties()
    )
}

fn owner_xml() -> Vec<u8> {
    let section_properties = properties(
        SECTION_ID,
        Some("rIdZoomCover"),
        " returnToParent=\"false\" imageType=\"cover\" transitionDur=\"1s\" showBg=\"false\"",
    );
    let slide_properties = properties(SLIDE_OBJECT_ID, None, "");
    let summary_properties = properties(SUMMARY_OBJECT_ID, None, "");
    format!(
        r#"<p:spTree xmlns:p="{PML}" xmlns:a="{DML}" xmlns:r="{REL}" xmlns:mc="{MC}" xmlns:p166="{MAIN}" xmlns:psez="{SECTION}" xmlns:pslz="{SLIDE}" xmlns:psuz="{SUMMARY}" xmlns:future="urn:litchi:pptx:future">
  <p:nvGrpSpPr/><p:grpSpPr/>
  <mc:AlternateContent>
    <mc:Choice Requires="psez"><psez:sectionZm><psez:sectionZmObj sectionId="{SECTION_ID}">{section_properties}<p:extLst><p:ext uri="urn:litchi:section"/></p:extLst></psez:sectionZmObj></psez:sectionZm></mc:Choice>
    <mc:Choice Requires="future"><future:zoomExtension token="kept"/></mc:Choice>
    <mc:Fallback><p:pic/></mc:Fallback>
  </mc:AlternateContent>
  <mc:AlternateContent>
    <mc:Choice Requires="pslz"><pslz:sldZm><pslz:sldZmObj sldId="256" cId="7">{slide_properties}</pslz:sldZmObj></pslz:sldZm></mc:Choice>
    <mc:Fallback><p:pic/></mc:Fallback>
  </mc:AlternateContent>
  <mc:AlternateContent>
    <mc:Choice Requires="psuz"><psuz:summaryZm><psuz:summaryZmObj sectionId="{SECTION_ID}" title="Agenda" descr="Overview" offsetFactorX="-2500" scaleFactorX="125000">{summary_properties}</psuz:summaryZmObj><psuz:gridLayout/><p:extLst><p:ext uri="urn:litchi:summary"/></p:extLst></psuz:summaryZm></mc:Choice>
    <mc:Fallback><p:grpSp/></mc:Fallback>
  </mc:AlternateContent>
  <p:sp/>
</p:spTree>"#
    )
    .into_bytes()
}

fn owner_text() -> String {
    String::from_utf8(owner_xml()).expect("valid fixture XML")
}

fn standalone_properties(id: &str, relationship: Option<&str>) -> Properties {
    Properties::new(
        id,
        blip_fill(relationship).into_bytes(),
        shape_properties().as_bytes().to_vec(),
    )
    .expect("valid zoom properties")
}

#[test]
fn reads_typed_zoom_shapes_and_preserves_fallback_and_unknown_choices() {
    let xml = owner_xml();
    let owner = Owner::read(&xml).expect("zoom owner");
    assert_eq!(owner.len(), 3);
    assert_eq!(owner.to_xml().expect("round trip"), xml);

    let Zoom::Section(section) = owner.get(0).expect("section zoom") else {
        panic!("expected section zoom");
    };
    assert_eq!(section.section_id(), SECTION_ID);
    assert_eq!(section.properties().image_type(), ImageType::Cover);
    assert!(!section.properties().return_to_parent());
    assert_eq!(
        section
            .properties()
            .transition()
            .expect("transition")
            .as_str(),
        "1000"
    );
    assert_eq!(
        section
            .properties()
            .image_relationship()
            .expect("relationship")
            .id(),
        "rIdZoomCover"
    );
    assert_eq!(section.fallback_xml(), b"<p:pic/>");
    assert_eq!(section.unknown_xml().len(), 1);
    assert!(
        std::str::from_utf8(&section.unknown_xml()[0])
            .expect("unknown XML")
            .contains("future:zoomExtension")
    );

    let Zoom::Slide(slide) = owner.get(1).expect("slide zoom") else {
        panic!("expected slide zoom");
    };
    assert_eq!(slide.slide_id(), 256);
    assert_eq!(slide.creation_id(), Some(7));

    let Zoom::Summary(summary) = owner.get(2).expect("summary zoom") else {
        panic!("expected summary zoom");
    };
    assert_eq!(summary.layout(), Layout::Grid);
    assert_eq!(summary.items().len(), 1);
    assert_eq!(summary.items()[0].title(), "Agenda");
    assert_eq!(summary.items()[0].offset_x(), Percentage::new(-2500));
    assert_eq!(summary.items()[0].scale_x(), Percentage::new(125000));
}

#[test]
fn owner_crud_patches_only_zoom_entries() {
    let original = owner_xml();
    let mut owner = Owner::read(&original).expect("zoom owner");
    let replacement = owner.get(0).expect("section").clone();
    let removed = owner.remove(1).expect("remove slide zoom");
    assert!(matches!(removed, Zoom::Slide(_)));
    let replaced = owner.replace(0, replacement).expect("replace section zoom");
    assert!(matches!(replaced, Zoom::Section(_)));
    assert_eq!(owner.len(), 2);
    assert!(
        std::str::from_utf8(owner.xml())
            .expect("owner XML")
            .contains("<p:sp/>")
    );

    let unknown_xml = br#"<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice Requires="future"><future:zoom xmlns:future="urn:litchi:pptx:future"/></mc:Choice><mc:Fallback><p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/></mc:Fallback></mc:AlternateContent>"#;
    let index = owner
        .add(Zoom::Unknown(
            Unknown::new(unknown_xml.to_vec()).expect("unknown"),
        ))
        .expect("add unknown");
    assert_eq!(index, 2);
    assert!(matches!(owner.get(index), Some(Zoom::Unknown(_))));

    owner.clear().expect("clear zooms");
    assert!(owner.is_empty());
    assert!(
        std::str::from_utf8(owner.xml())
            .expect("owner XML")
            .contains("<p:sp/>")
    );
    assert!(
        !std::str::from_utf8(owner.xml())
            .expect("owner XML")
            .contains("AlternateContent")
    );
}

#[test]
fn malformed_zoom_content_is_rejected() {
    let cases = [
        (
            "missing fallback",
            owner_text().replace("<mc:Fallback><p:pic/></mc:Fallback>", ""),
        ),
        (
            "invalid section GUID",
            owner_text().replace(SECTION_ID, "not-a-guid"),
        ),
        (
            "invalid image type",
            owner_text().replace("imageType=\"cover\"", "imageType=\"poster\""),
        ),
        (
            "invalid summary layout",
            owner_text().replace("<psuz:gridLayout/>", "<future:otherLayout/>"),
        ),
        (
            "multiple roots",
            format!("{}<p:spTree xmlns:p=\"{PML}\"/>", owner_text()),
        ),
    ];
    for (name, xml) in cases {
        assert!(
            Owner::read(xml.as_bytes()).is_err(),
            "accepted malformed case: {name}"
        );
    }
}

#[test]
fn constructors_and_typed_mutations_validate_domains() {
    assert!(Percentage::new(i32::MAX).value() > 0);
    assert!(Properties::new("bad", b"<p166:blipFill/>", b"<p166:spPr/>").is_err());
    assert!(
        Slide::new(
            255,
            standalone_properties(SLIDE_OBJECT_ID, None),
            b"<p:pic/>"
        )
        .is_err()
    );
    assert!(Section::new("bad", standalone_properties(SECTION_ID, None), b"<p:pic/>").is_err());
    assert!(Item::new("bad", standalone_properties(SUMMARY_OBJECT_ID, None)).is_err());
}

#[test]
fn package_context_resolves_targets_and_rejects_dangling_image_relationships() {
    let mut package = crate::Package::new().expect("new package");
    package
        .presentation_mut()
        .expect("writer")
        .add_slide()
        .expect("slide");
    let bytes = package.to_bytes().expect("serialize");
    let mut package = crate::Package::from_bytes(&bytes).expect("reopen");

    let mut owner = package.zooms(0).expect("empty owner");
    let slide = Slide::new(
        256,
        standalone_properties(SLIDE_OBJECT_ID, None),
        b"<p:pic/>",
    )
    .expect("slide zoom");
    owner.add(Zoom::Slide(slide)).expect("add slide zoom");
    let previous = package.put_zooms(0, owner).expect("store zoom");
    assert!(previous.expect("previous owner").is_empty());
    assert!(matches!(
        package.zooms(0).expect("loaded owner").get(0),
        Some(Zoom::Slide(_))
    ));

    let mut owner = package.zooms(0).expect("loaded owner");
    owner.clear().expect("clear");
    let bad_properties = standalone_properties(SLIDE_OBJECT_ID, Some("rIdMissing"));
    owner
        .add(Zoom::Slide(
            Slide::new(256, bad_properties, b"<p:pic/>").expect("bad relationship reference"),
        ))
        .expect("stage bad relationship");
    assert!(package.put_zooms(0, owner).is_err());
}
