use litchi_opc::constants::relationship_type as rt;
use litchi_opc::{OpcPackage, PackURI};
use litchi_pptx::Error;
use litchi_pptx::Package;
use litchi_pptx::parts::{PresentationPart, SlideMasterPart, SlidePart};
use litchi_pptx::transition::{Kind, Ms, Ripple, Transition, read};

const LOCAL_P14_RIPPLE: &[u8] =
    include_bytes!("../../../test-data/ooxml/pptx/transitions/p14_ripple.xml");

#[test]
fn layout_and_master_read_local_powerpoint_2010_transition_choices() {
    let package =
        package_with_inherited_transition_fragment(std::str::from_utf8(LOCAL_P14_RIPPLE).unwrap());

    let presentation = PresentationPart::from_package(&package).unwrap();
    let master = SlideMasterPart::from_part(
        package
            .get_part(&PackURI::new("/ppt/slideMasters/slideMaster1.xml").unwrap())
            .unwrap(),
    )
    .unwrap();
    assert_ripple(&read(master.part().blob()).unwrap().unwrap());

    let layouts = master.layouts(&package).unwrap();
    let layout = layouts
        .iter()
        .find(|layout| layout.part().partname().as_str() == "/ppt/slideLayouts/slideLayout1.xml")
        .unwrap();
    assert_ripple(&read(layout.part().blob()).unwrap().unwrap());
    let inherited_master = layout.master(&package).unwrap();
    assert_ripple(&read(inherited_master.part().blob()).unwrap().unwrap());

    let reference = presentation.slide_references().unwrap().remove(0);
    let slide = SlidePart::from_part(
        package
            .get_part(&PackURI::new("/ppt/slides/slide1.xml").unwrap())
            .unwrap(),
    )
    .unwrap();
    assert_eq!(reference.relationship_id(), "rId4");
    let layout = slide.layout(&package).unwrap().unwrap();
    assert_ripple(&read(layout.part().blob()).unwrap().unwrap());
    assert_ripple(
        &read(layout.master(&package).unwrap().part().blob())
            .unwrap()
            .unwrap(),
    );
}

#[test]
fn master_layout_inventory_rejects_external_layout_relationships() {
    let mut package = package_with_slides(0);
    let master_name = PackURI::new("/ppt/slideMasters/slideMaster1.xml").unwrap();
    let relationship_id = package
        .get_part(&master_name)
        .unwrap()
        .rels()
        .iter()
        .find(|relationship| relationship.reltype() == rt::SLIDE_LAYOUT)
        .unwrap()
        .r_id()
        .to_string();
    let master = package.get_part_mut(&master_name).unwrap();
    master.rels_mut().remove(&relationship_id);
    master.rels_mut().add_relationship(
        rt::SLIDE_LAYOUT.to_string(),
        "https://example.invalid/slide-layout.xml".to_string(),
        relationship_id,
        true,
    );

    let master = SlideMasterPart::from_part(package.get_part(&master_name).unwrap()).unwrap();
    let result = master.layouts(&package);
    assert!(matches!(
        result,
        Err(Error::Relationship(message)) if message.contains("must be internal")
    ));
}

#[test]
fn slide_layout_accessor_rejects_external_layout_relationships() {
    let mut package = package_with_slides(1);
    let slide_name = PackURI::new("/ppt/slides/slide1.xml").unwrap();
    let relationship_id = package
        .get_part(&slide_name)
        .unwrap()
        .rels()
        .iter()
        .find(|relationship| relationship.reltype() == rt::SLIDE_LAYOUT)
        .unwrap()
        .r_id()
        .to_string();
    let slide = package.get_part_mut(&slide_name).unwrap();
    slide.rels_mut().remove(&relationship_id);
    slide.rels_mut().add_relationship(
        rt::SLIDE_LAYOUT.to_string(),
        "https://example.invalid/slide-layout.xml".to_string(),
        relationship_id,
        true,
    );

    let slide = SlidePart::from_part(package.get_part(&slide_name).unwrap()).unwrap();
    let result = slide.layout(&package);
    assert!(matches!(
        result,
        Err(Error::Relationship(message)) if message.contains("must be internal")
    ));
}

#[test]
fn layout_master_accessor_rejects_external_master_relationships() {
    let mut package = package_with_slides(1);
    let slide_name = PackURI::new("/ppt/slides/slide1.xml").unwrap();
    let layout_name = package
        .get_part(&slide_name)
        .unwrap()
        .rels()
        .iter()
        .find(|relationship| relationship.reltype() == rt::SLIDE_LAYOUT)
        .unwrap()
        .target_partname()
        .unwrap();
    let relationship_id = package
        .get_part(&layout_name)
        .unwrap()
        .rels()
        .iter()
        .find(|relationship| relationship.reltype() == rt::SLIDE_MASTER)
        .unwrap()
        .r_id()
        .to_string();
    let layout = package.get_part_mut(&layout_name).unwrap();
    layout.rels_mut().remove(&relationship_id);
    layout.rels_mut().add_relationship(
        rt::SLIDE_MASTER.to_string(),
        "https://example.invalid/slide-master.xml".to_string(),
        relationship_id,
        true,
    );

    let slide = SlidePart::from_part(package.get_part(&slide_name).unwrap()).unwrap();
    let result = slide.layout(&package).unwrap().unwrap().master(&package);
    assert!(matches!(
        result,
        Err(Error::Relationship(message)) if message.contains("must be internal")
    ));
}

fn assert_ripple(transition: &Transition) {
    assert_eq!(transition.duration().map(Ms::get), Some(1500));
    assert_eq!(transition.kind(), &Kind::Ripple(Ripple::LeftDown));
}

fn package_with_inherited_transition_fragment(fragment: &str) -> OpcPackage {
    let mut package = package_with_slides(1);
    for (part_name, end_tag) in [
        ("/ppt/slideLayouts/slideLayout1.xml", "</p:sldLayout>"),
        ("/ppt/slideMasters/slideMaster1.xml", "</p:sldMaster>"),
    ] {
        let part_name = PackURI::new(part_name).unwrap();
        let part = package.get_part_mut(&part_name).unwrap();
        let xml = std::str::from_utf8(part.blob()).unwrap();
        let updated = xml.replacen(end_tag, &format!("{fragment}{end_tag}"), 1);
        assert_ne!(updated, xml);
        part.set_blob(updated.into_bytes());
    }
    package
}

fn package_with_slides(count: usize) -> OpcPackage {
    let mut package = Package::new().unwrap();
    for _ in 0..count {
        package.presentation_mut().unwrap().add_slide().unwrap();
    }
    OpcPackage::from_vec(package.to_bytes().unwrap()).unwrap()
}
