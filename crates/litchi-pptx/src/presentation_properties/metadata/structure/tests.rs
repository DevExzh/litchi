#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{BlobPart, OpcPackage, PackURI};

use super::*;
use crate::presentation_properties::metadata::custom_show::Show;
use crate::presentation_properties::metadata::sections::Section;

const P: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
const PS: &str = "http://purl.oclc.org/ooxml/presentationml/main";
const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const RS: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships";
const P14: &str = "http://schemas.microsoft.com/office/powerpoint/2010/main";
const SECTION_URI: &str = "{521415D9-36F7-43E2-AB2F-B90AF26B5E84}";
const SECTION_ONE: &str = "{11111111-1111-1111-1111-111111111111}";
const SECTION_TWO: &str = "{22222222-2222-2222-2222-222222222222}";

fn fixture() -> OpcPackage {
    let mut package = OpcPackage::new();
    package.rels_mut().add_relationship(
        rt::OFFICE_DOCUMENT.to_owned(),
        "ppt/presentation.xml".to_owned(),
        "rIdPresentation".to_owned(),
        false,
    );
    package.add_part(Box::new(BlobPart::new(
        PackURI::new("/ppt/presentation.xml").unwrap(),
        ct::PML_PRESENTATION_MAIN.into(),
        format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:p="{P}" xmlns:r="{R}" xmlns:p14="{P14}" xmlns:x="urn:future">
  <p:sldIdLst>
    <p:sldId id="256" r:id="rIdSlideOne">
      <p:extLst><p:ext uri="{{slide-opaque}}"><x:future>keep-slide-extension</x:future></p:ext></p:extLst>
    </p:sldId>
    <p:sldId id="257" r:id="rIdSlideTwo"/>
    <p:sldId id="258" r:id="rIdSlideThree"/>
  </p:sldIdLst>
  <p:custShowLst>
    <p:custShow name="Opening" id="7" future="keep">
      <p:sldLst><p:sld r:id="rIdSlideOne"/><p:sld r:id="rIdSlideThree"/></p:sldLst>
    </p:custShow>
    <p:custShow name="Recap" id="8"><p:sldLst><p:sld r:id="rIdSlideTwo"/></p:sldLst></p:custShow>
  </p:custShowLst>
  <p:extLst>
    <p:ext uri="{SECTION_URI}">
      <p14:sectionLst>
        <p14:section name="Opening" id="{SECTION_ONE}"><p14:sldIdLst><p14:sldId id="256"/><p14:sldId id="258"/></p14:sldIdLst></p14:section>
        <p14:section name="Recap" id="{SECTION_TWO}"><p14:sldIdLst><p14:sldId id="257"/></p14:sldIdLst></p14:section>
      </p14:sectionLst>
    </p:ext>
    <p:ext uri="{{opaque}}"><x:future>keep</x:future></p:ext>
  </p:extLst>
</p:presentation>"#
        )
        .into_bytes(),
    )));

    let presentation = package
        .get_part_mut(&PackURI::new("/ppt/presentation.xml").unwrap())
        .unwrap();
    for (relationship_id, target) in [
        ("rIdSlideOne", "slides/slide1.xml"),
        ("rIdSlideTwo", "slides/slide2.xml"),
        ("rIdSlideThree", "slides/slide3.xml"),
    ] {
        presentation.rels_mut().add_relationship(
            rt::SLIDE.to_owned(),
            target.to_owned(),
            relationship_id.to_owned(),
            false,
        );
    }
    presentation.rels_mut().add_relationship(
        rt::THEME.to_owned(),
        "theme/theme1.xml".to_owned(),
        "rIdTheme".to_owned(),
        false,
    );

    for name in [
        "/ppt/slides/slide1.xml",
        "/ppt/slides/slide2.xml",
        "/ppt/slides/slide3.xml",
    ] {
        package.add_part(Box::new(BlobPart::new(
            PackURI::new(name).unwrap(),
            ct::PML_SLIDE.into(),
            b"<p:sld xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\"/>"
                .to_vec(),
        )));
    }
    package
}

fn presentation_xml(package: &OpcPackage) -> Vec<u8> {
    package
        .get_part(&PackURI::new("/ppt/presentation.xml").unwrap())
        .unwrap()
        .blob()
        .to_vec()
}

fn set_presentation_xml(package: &mut OpcPackage, xml: Vec<u8>) {
    package
        .get_part_mut(&PackURI::new("/ppt/presentation.xml").unwrap())
        .unwrap()
        .set_blob(xml);
}

#[test]
fn non_empty_slide_id_is_shared_by_ordering_custom_shows_and_sections() {
    let package = fixture();
    let graph = load(&package).unwrap();

    assert_eq!(
        graph
            .slides
            .iter()
            .map(|slide| slide.slide_id)
            .collect::<Vec<_>>(),
        [256, 257, 258]
    );
    assert_eq!(
        graph.custom_shows.get_by_id(7).unwrap().slide_ids,
        [256, 258]
    );
    assert_eq!(
        graph.sections.get_by_id(SECTION_ONE).unwrap().slide_ids,
        [256, 258]
    );
}

#[test]
fn non_empty_slide_id_uses_the_same_duplicate_and_relationship_validation() {
    let cases = [
        (
            r#"<p:sldId id="257" r:id="rIdSlideTwo"/>"#,
            r#"<p:sldId id="256" r:id="rIdSlideTwo"/>"#,
        ),
        (
            r#"<p:sldId id="257" r:id="rIdSlideTwo"/>"#,
            r#"<p:sldId id="257" r:id="rIdSlideOne"/>"#,
        ),
        (
            r#"<p:sldId id="256" r:id="rIdSlideOne">"#,
            r#"<p:sldId id="256" r:id="rIdMissing">"#,
        ),
    ];

    for (source, replacement) in cases {
        let mut package = fixture();
        let xml = String::from_utf8(presentation_xml(&package)).unwrap();
        let malformed = xml.replacen(source, replacement, 1).into_bytes();
        set_presentation_xml(&mut package, malformed);
        assert!(load(&package).is_err());
    }
}

#[test]
fn structure_names_are_resolved_and_accept_transitional_and_strict_aliases() {
    for (presentationml, relationships) in [(P, R), (PS, RS)] {
        let xml = format!(
            r#"<main:presentation xmlns:main="{presentationml}" xmlns:rel="{relationships}"><main:sldIdLst><main:sldId id="256" rel:id="slide"/></main:sldIdLst><main:custShowLst><main:custShow name="show" id="1"><main:sldLst><main:sld rel:id="slide"/></main:sldLst></main:custShow></main:custShowLst></main:presentation>"#
        );
        let (slides, shows) = codec::parse_core(xml.as_bytes()).unwrap();
        assert_eq!(slides, [(256, "slide".to_owned())]);
        assert_eq!(shows.len(), 1);
        assert_eq!(shows[0].relationship_ids, ["slide"]);
    }
}

#[test]
fn structure_parser_rejects_namespace_lookalikes_and_arbitrary_id_prefixes() {
    let cases = [
        (
            r#"<p:sldId id="257" r:id="rIdSlideTwo"/>"#,
            r#"<x:sldId id="257" r:id="rIdSlideTwo"/>"#,
        ),
        (
            r#"<p:sldId id="257" r:id="rIdSlideTwo"/>"#,
            r#"<p:sldId id="257" x:id="rIdSlideTwo"/>"#,
        ),
        (
            "<p:sldIdLst>",
            r#"<x:sldIdLst xmlns:x="urn:not-presentation">"#,
        ),
    ];

    for (source, replacement) in cases {
        let mut package = fixture();
        let xml = String::from_utf8(presentation_xml(&package)).unwrap();
        set_presentation_xml(
            &mut package,
            xml.replacen(source, replacement, 1).into_bytes(),
        );
        assert!(load(&package).is_err());
    }
}

#[test]
fn presentation_slide_id_enforces_the_full_schema_range() {
    for id in ["255", "2147483648", "4294967295"] {
        let xml = format!(
            r#"<p:presentation xmlns:p="{P}" xmlns:r="{R}"><p:sldIdLst><p:sldId id="{id}" r:id="slide"/></p:sldIdLst></p:presentation>"#
        );
        assert!(
            codec::parse_core(xml.as_bytes()).is_err(),
            "accepted slide ID {id}"
        );
    }

    for id in ["256", "2147483647"] {
        let xml = format!(
            r#"<p:presentation xmlns:p="{P}" xmlns:r="{R}"><p:sldIdLst><p:sldId id="{id}" r:id="slide"/></p:sldIdLst></p:presentation>"#
        );
        assert_eq!(
            codec::parse_core(xml.as_bytes()).unwrap().0[0]
                .0
                .to_string(),
            id
        );
    }
}

#[test]
fn no_op_and_inverse_patches_preserve_exact_source_and_opaque_xml() {
    let mut package = fixture();
    let source = Snapshot::load(&package).unwrap();
    let source_xml = source.source_xml().to_vec();

    let no_op = source.edit().commit().unwrap();
    assert!(!no_op.is_changed());
    assert!(no_op.patch().is_empty());
    apply_commit(&mut package, no_op).unwrap();
    assert_eq!(presentation_xml(&package), source_xml);

    let mut section_only = source.edit();
    section_only
        .edit_sections(|sections| {
            sections.get_by_id_mut(SECTION_TWO).unwrap().name = Some("Closing".into());
            Ok(())
        })
        .unwrap();
    let section_commit = section_only.commit().unwrap();
    assert!(
        String::from_utf8_lossy(section_commit.snapshot().source_xml())
            .contains(r#"future="keep""#)
    );

    let mut edit = source.edit();
    edit.edit_custom_shows(|shows| {
        shows.get_by_id_mut(7).unwrap().name = "Opening Updated".into();
        Ok(())
    })
    .unwrap();
    edit.edit_sections(|sections| {
        sections.get_by_id_mut(SECTION_TWO).unwrap().name = Some("Closing".into());
        Ok(())
    })
    .unwrap();
    let commit = edit.commit().unwrap();
    let after_xml = commit.snapshot().source_xml().to_vec();
    assert!(commit.is_changed());
    assert!(String::from_utf8_lossy(&after_xml).contains("<x:future>keep</x:future>"));

    apply_commit(&mut package, commit.clone()).unwrap();
    assert_eq!(presentation_xml(&package), after_xml);
    commit.patch().inverse().apply(&mut package).unwrap();
    assert_eq!(presentation_xml(&package), source_xml);
}

#[test]
fn stale_source_and_relationship_topology_are_rejected_without_mutation() {
    let mut package = fixture();
    let source = Snapshot::load(&package).unwrap();
    let mut edit = source.edit();
    edit.edit_custom_shows(|shows| {
        shows.get_by_id_mut(8).unwrap().slide_ids.push(256);
        Ok(())
    })
    .unwrap();
    let patch = edit.commit().unwrap().into_patch();

    let before = presentation_xml(&package);
    package
        .get_part_mut(&PackURI::new("/ppt/presentation.xml").unwrap())
        .unwrap()
        .rels_mut()
        .add_relationship(
            "urn:future".into(),
            "future.bin".into(),
            "rIdFuture".into(),
            false,
        );
    assert!(patch.apply(&mut package).is_err());
    assert_eq!(presentation_xml(&package), before);

    let mut stale = fixture();
    let mut changed = presentation_xml(&stale);
    changed.extend_from_slice(b"\n");
    stale
        .get_part_mut(&PackURI::new("/ppt/presentation.xml").unwrap())
        .unwrap()
        .set_blob(changed.clone());
    assert!(patch.apply(&mut stale).is_err());
    assert_eq!(presentation_xml(&stale), changed);
}

#[test]
fn custom_show_and_section_edits_round_trip_as_one_validated_graph() {
    let mut package = fixture();
    let snapshot = Snapshot::load(&package).unwrap();
    let mut edit = snapshot.edit();
    edit.edit(|graph| {
        graph
            .custom_shows
            .add(Show::new(9, "All slides").with_slides(vec![256, 257, 258]));
        graph
            .sections
            .get_by_id_mut(SECTION_ONE)
            .unwrap()
            .slide_ids
            .retain(|slide_id| *slide_id != 258);
        graph.sections.add_section(
            Section::new("Tail", "{33333333-3333-3333-3333-333333333333}").with_slides([258]),
        );
        Ok(())
    })
    .unwrap();

    let commit = edit.commit().unwrap();
    let published = apply_commit(&mut package, commit).unwrap();
    assert_eq!(published.custom_shows().shows.len(), 3);
    assert_eq!(
        published.custom_shows().get_by_id(9).unwrap().slide_ids,
        [256, 257, 258]
    );
    assert_eq!(published.sections().sections().len(), 3);
    assert_eq!(load(&package).unwrap(), published.graph().clone());
}
