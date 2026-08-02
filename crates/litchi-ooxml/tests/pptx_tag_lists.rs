use litchi_ooxml::pptx::Package;
use litchi_ooxml::{OoxmlError, PackURI};
use litchi_opc::part::BlobPart;
use litchi_pptx::tag::{self, List, Tag};
use tempfile::NamedTempFile;

const TAG_RELATIONSHIP_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/tags";
const TAG_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.presentationml.tags+xml";
const LOCAL_PRIMARY_TAGS: &[u8] =
    include_bytes!("../../../test-data/ooxml/pptx/tags/basic_tags.xml");
const LOCAL_SECONDARY_TAGS: &[u8] =
    include_bytes!("../../../test-data/ooxml/pptx/tags/secondary_tags.xml");

#[test]
fn package_reads_direct_slide_tag_lists() {
    let package = package_with_direct_tag_lists();

    let primary = package.tags(0).unwrap().unwrap();
    assert_eq!(package.tags("Overview").unwrap(), Some(primary.clone()));
    assert_eq!(primary.tags().len(), 2);
    assert_eq!(primary.tags()[0].name(), "OWNER");
    assert_eq!(primary.tags()[0].value(), "Alice");
    assert_eq!(primary.tags()[1].value(), "<not-a-command/>");
    assert_eq!(primary.attrs()[0].qualified_name(), "ext:origin");
    assert_eq!(primary.attrs()[0].value(), "local");
    assert_eq!(
        package
            .presentation()
            .unwrap()
            .slide(0)
            .unwrap()
            .unwrap()
            .tags()
            .unwrap(),
        Some(primary)
    );

    let secondary = package.tags(1).unwrap().unwrap();
    assert_eq!(secondary.tags()[0].name(), "STATUS");
    assert_eq!(secondary.tags()[0].value(), "Review");
}

#[test]
fn slide_selectors_report_missing_ambiguous_and_out_of_bounds() {
    let mut package = package_with_direct_tag_lists();

    assert!(matches!(
        package.tags("Missing"),
        Err(OoxmlError::Pptx(litchi_pptx::Error::SlideNameNotFound(name)))
            if name == "Missing"
    ));
    assert!(matches!(
        package.tags(2),
        Err(OoxmlError::Pptx(
            litchi_pptx::Error::SlideIndexOutOfBounds { index: 2, len: 2 }
        ))
    ));

    name_slide(&mut package, "/ppt/slides/slide2.xml", "Overview");
    assert!(matches!(
        package.tags("Overview"),
        Err(OoxmlError::Pptx(litchi_pptx::Error::AmbiguousSlideName {
            name,
            matches: 2
        })) if name == "Overview"
    ));
}

#[test]
fn numeric_slide_lookup_does_not_process_unrelated_slides() {
    let mut package = package_with_direct_tag_lists();
    let unrelated = PackURI::new("/ppt/slides/slide2.xml").unwrap();
    package
        .opc_package_mut()
        .get_part_mut(&unrelated)
        .unwrap()
        .set_blob(b"<not-valid-xml".to_vec());

    let selected = package.tags(0).unwrap().unwrap();
    assert_eq!(selected.get("owner").unwrap().value(), "Alice");
}

#[test]
fn direct_anchor_rejects_an_external_tag_relationship() {
    let mut package = package_with_direct_tag_lists();
    let slide_name = PackURI::new("/ppt/slides/slide1.xml").unwrap();
    let source = tag::load(package.opc_package(), &slide_name)
        .unwrap()
        .unwrap();
    let relationship_id = source.rel().to_owned();
    let slide = package.opc_package_mut().get_part_mut(&slide_name).unwrap();
    let relationship = slide.rels_mut().remove(&relationship_id).unwrap();
    slide.rels_mut().add_relationship(
        relationship.reltype().to_owned(),
        "https://example.invalid/tags.xml".to_owned(),
        relationship_id,
        true,
    );

    assert!(matches!(
        package.tags(0),
        Err(OoxmlError::Pptx(litchi_pptx::Error::Invalid(message)))
            if message.contains("cannot be external")
    ));
}

#[test]
fn direct_anchor_rejects_the_wrong_tag_content_type() {
    let mut package = package_with_direct_tag_lists();
    let slide_name = PackURI::new("/ppt/slides/slide1.xml").unwrap();
    let part_name = tag::load(package.opc_package(), &slide_name)
        .unwrap()
        .unwrap()
        .part()
        .clone();
    assert!(package.opc_package_mut().remove_part(&part_name));
    package.opc_package_mut().add_part(Box::new(BlobPart::new(
        part_name,
        "application/xml".to_owned(),
        LOCAL_PRIMARY_TAGS.to_vec(),
    )));

    assert!(matches!(
        package.tags(0),
        Err(OoxmlError::Pptx(litchi_pptx::Error::ContentType { expected, actual }))
            if expected == TAG_CONTENT_TYPE && actual == "application/xml"
    ));
}

#[test]
fn unanchored_relationships_stay_in_the_low_level_inventory() {
    let mut package = package_with_named_slides();
    install_unanchored_tag_relationship(&mut package);

    assert_eq!(package.tags(0).unwrap(), None);
    let presentation = package.presentation().unwrap();
    let slides = presentation.slides().unwrap();
    let sources = slides[0].tag_inventory().unwrap();
    assert_eq!(sources.len(), 1);
    assert_eq!(sources[0].list().get("owner").unwrap().value(), "Alice");
}

#[test]
fn real_shape_tags_are_not_flattened_into_slide_tags() {
    let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../..");
    let package =
        Package::open(root.join("test-data/libreoffice-core/sd/qa/unit/data/pptx/tdf103477.pptx"))
            .unwrap();

    assert_eq!(package.tags(0).unwrap(), None);
    let presentation = package.presentation().unwrap();
    let slides = presentation.slides().unwrap();
    let sources = slides[0].tag_inventory().unwrap();
    assert_eq!(sources.len(), 7);
    assert!(sources.iter().all(|source| !source.list().is_empty()));
}

#[test]
fn package_tag_put_and_remove_are_singleton_and_persistent() {
    let mut package = package_with_named_slides();
    let created = tag_list("Priority", "High");
    assert_eq!(package.put_tags("Overview", created.clone()).unwrap(), None);
    assert_eq!(package.tags("Overview").unwrap(), Some(created.clone()));

    let previous = package.put_tags("Overview", List::new()).unwrap();
    assert_eq!(previous, Some(created));
    assert_eq!(package.tags("Overview").unwrap(), Some(List::new()));

    let output = NamedTempFile::with_suffix(".pptx").unwrap();
    package.save(output.path()).unwrap();
    let mut reopened = Package::open(output.path()).unwrap();
    assert_eq!(reopened.tags("Overview").unwrap(), Some(List::new()));
    assert_eq!(reopened.remove_tags("Overview").unwrap(), Some(List::new()));
    assert_eq!(reopened.remove_tags("Overview").unwrap(), None);
    reopened.save(output.path()).unwrap();

    let reopened = Package::open(output.path()).unwrap();
    assert_eq!(reopened.tags("Overview").unwrap(), None);
}

#[test]
fn package_tag_edits_reject_a_dirty_legacy_writer() {
    let mut package = Package::new().unwrap();
    package.presentation_mut().unwrap().add_slide().unwrap();

    assert!(matches!(
        package.put_tags(0, tag_list("Owner", "Alice")),
        Err(OoxmlError::UnsafeEdit {
            format: "PPTX",
            operation: "put_tags",
            ..
        })
    ));
}

fn package_with_direct_tag_lists() -> Package {
    let mut package = package_with_named_slides();
    assert_eq!(
        package
            .put_tags("Overview", tag::parse(LOCAL_PRIMARY_TAGS).unwrap())
            .unwrap(),
        None
    );
    assert_eq!(
        package
            .put_tags("Appendix", tag::parse(LOCAL_SECONDARY_TAGS).unwrap())
            .unwrap(),
        None
    );
    package
}

fn package_with_named_slides() -> Package {
    let output = NamedTempFile::with_suffix(".pptx").unwrap();
    let mut package = Package::new().unwrap();
    package.presentation_mut().unwrap().add_slide().unwrap();
    package.presentation_mut().unwrap().add_slide().unwrap();
    package.save(output.path()).unwrap();

    let mut package = Package::open(output.path()).unwrap();
    name_slide(&mut package, "/ppt/slides/slide1.xml", "Overview");
    name_slide(&mut package, "/ppt/slides/slide2.xml", "Appendix");
    package
}

fn name_slide(package: &mut Package, slide_name: &str, name: &str) {
    let slide_name = PackURI::new(slide_name).unwrap();
    let slide = package.opc_package_mut().get_part_mut(&slide_name).unwrap();
    let xml = std::str::from_utf8(slide.blob()).unwrap();
    let named = if xml.contains("<p:cSld>") {
        xml.replacen("<p:cSld>", &format!(r#"<p:cSld name="{name}">"#), 1)
    } else {
        let marker = " name=\"";
        let root = xml.find("<p:cSld ").expect("generated slide has p:cSld");
        let end = root
            + xml[root..]
                .find('>')
                .expect("generated p:cSld has a closing delimiter");
        let value_start = root
            + xml[root..end]
                .find(marker)
                .expect("named slide has a name attribute")
            + marker.len();
        let value_end = value_start
            + xml[value_start..end]
                .find('"')
                .expect("slide name attribute is quoted");
        let mut named = xml.to_owned();
        named.replace_range(value_start..value_end, name);
        named
    };
    assert_ne!(named, xml, "generated slide must contain p:cSld");
    slide.set_blob(named.into_bytes());
}

fn install_unanchored_tag_relationship(package: &mut Package) {
    let slide_name = PackURI::new("/ppt/slides/slide1.xml").unwrap();
    let part_name = PackURI::new("/ppt/tags/tag99.xml").unwrap();
    package.opc_package_mut().add_part(Box::new(BlobPart::new(
        part_name,
        TAG_CONTENT_TYPE.to_owned(),
        LOCAL_PRIMARY_TAGS.to_vec(),
    )));
    package
        .opc_package_mut()
        .get_part_mut(&slide_name)
        .unwrap()
        .rels_mut()
        .add_relationship(
            TAG_RELATIONSHIP_TYPE.to_owned(),
            "../tags/tag99.xml".to_owned(),
            "rIdDanglingTags".to_owned(),
            false,
        );
}

fn tag_list(name: &str, value: &str) -> List {
    let mut list = List::new();
    list.add(Tag::new(name, value).unwrap()).unwrap();
    list
}
