#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]
//! The presentation-metadata and comment-collaboration edit families of
//! record 0356: each commit publishes a new package with a reversible patch,
//! the source package never changes, and an unchanged edit publishes the
//! source bytes.

use litchi_pptx::Package;
use litchi_pptx::presentation_properties::metadata::color_map::{Role, Slot};
use litchi_pptx::presentation_properties::metadata::custom_show::List as CustomShows;
use litchi_pptx::presentation_properties::metadata::events::Draft as Event;
use litchi_pptx::presentation_properties::metadata::guides::{Guides, Limits, ListKind};
use litchi_pptx::presentation_properties::metadata::sections::Section;
use litchi_pptx::presentation_properties::metadata::structure::{self, Limits as StructureLimits};
use litchi_pptx::time::Offset;

const GUIDES_XML: &[u8] =
    include_bytes!("../../../test-data/ooxml/pptx/extended-guides/presentation.xml");

fn deck() -> Package {
    let mut builder = litchi_pptx::builder::Builder::new().unwrap();
    builder.presentation().unwrap().add_slide().unwrap();
    builder.build().unwrap()
}

fn slide_id(package: &Package) -> u32 {
    package
        .with_opc(|opc| Ok(structure::load(opc, StructureLimits::DEFAULT)?.slides[0].slide_id))
        .unwrap()
}

#[test]
fn guide_edits_publish_a_new_package_and_invert() {
    let package = deck();
    let guides = Guides::from_xml(GUIDES_XML, Limits::DEFAULT).unwrap();
    let mut edit = package.edit_guides().unwrap();
    edit.set(&guides).unwrap();
    assert_eq!(edit.guides().unwrap(), guides);
    let commit = edit.commit().unwrap();
    assert!(commit.diagnostics().changed());
    let (after, patch) = commit.into_parts();
    assert_eq!(after.edit_guides().unwrap().guides().unwrap(), guides);
    assert_eq!(
        package.edit_guides().unwrap().guides().unwrap(),
        Guides::default()
    );

    let restored = after.apply_guides(&patch.inverse()).unwrap().into_package();
    assert!(restored.same_content(&package).unwrap());
    assert!(
        after.apply_guides(&patch).is_err(),
        "a patch replays only onto its base"
    );

    let mut edit = after.edit_guides().unwrap();
    edit.update(|draft| draft.remove(ListKind::Slide, 0).map(drop))
        .unwrap();
    let removed = edit.commit().unwrap().into_package();
    let staged = removed.edit_guides().unwrap().guides().unwrap();
    assert!(staged.slide.unwrap().guides.is_empty());
}

#[test]
fn an_unchanged_metadata_edit_publishes_the_source_bytes() {
    let package = deck();
    let bytes = package.to_bytes().unwrap();
    let mut edit = package.edit_custom_shows().unwrap();
    assert!(edit.remove().unwrap().is_none());
    let commit = edit.commit().unwrap();
    assert!(!commit.diagnostics().changed());
    assert!(commit.patch().is_noop());
    assert_eq!(commit.package().to_bytes().unwrap(), bytes);
}

#[test]
fn custom_show_and_section_edits_share_the_structure_graph() {
    let package = deck();
    let id = slide_id(&package);
    let mut shows = CustomShows::new();
    shows.create("Short", vec![id]);
    let mut edit = package.edit_custom_shows().unwrap();
    edit.set(&shows).unwrap();
    let after = edit.commit().unwrap().into_package();
    assert_eq!(
        after.edit_custom_shows().unwrap().custom_shows().unwrap(),
        shows
    );

    let mut edit = after.edit_structure().unwrap();
    let section = edit
        .add_section(Section::new(
            "Intro",
            "{8D3E1A50-3F0B-4C6E-9B0F-1C2D3E4F5A6B}",
        ))
        .unwrap();
    edit.add_section_slide(&section, id).unwrap();
    let after = edit.commit().unwrap().into_package();
    let graph = after.edit_structure().unwrap().structure().unwrap();
    assert_eq!(graph.custom_shows, shows);
    assert_eq!(graph.sections.len(), 1);
    assert!(
        package
            .edit_structure()
            .unwrap()
            .find_section(&section)
            .unwrap()
            .is_none()
    );
}

#[test]
fn slide_show_events_are_staged_per_slide() {
    let package = deck();
    let slide = 0_usize;
    let mut edit = package.edit_show_events().unwrap();
    edit.set(slide, &[Event::play(Offset::ms(5), 2)]).unwrap();
    let after = edit.commit().unwrap().into_package();
    let events = after
        .edit_show_events()
        .unwrap()
        .show_events(slide)
        .unwrap();
    assert_eq!(events.unwrap().len(), 1);
    assert!(
        package
            .edit_show_events()
            .unwrap()
            .show_events(slide)
            .unwrap()
            .is_none()
    );

    let mut edit = after.edit_show_events().unwrap();
    edit.update(slide, |draft| draft.push(Event::pause(Offset::ms(9), 2)))
        .unwrap();
    let extended = edit.commit().unwrap().into_package();
    let events = extended
        .edit_show_events()
        .unwrap()
        .show_events(slide)
        .unwrap();
    assert_eq!(events.unwrap().len(), 2);
    let mut edit = extended.edit_show_events().unwrap();
    assert!(edit.remove(slide).unwrap().is_some());
    let removed = edit.commit().unwrap().into_package();
    let events = removed
        .edit_show_events()
        .unwrap()
        .show_events(slide)
        .unwrap();
    assert!(events.is_none());
}

#[test]
fn master_color_map_edits_rewrite_only_the_master_map() {
    let package = deck();
    let master = package.presentation().unwrap().slide_masters().unwrap()[0]
        .id()
        .unwrap();
    let edit = package.edit_color_maps().unwrap();
    let before = edit.color_map(master).unwrap().map().unwrap();
    drop(edit);
    let role = if before.color(Slot::Background1) == Role::Accent1 {
        Role::Accent2
    } else {
        Role::Accent1
    };
    let mut edit = package.edit_color_maps().unwrap();
    assert!(
        edit.update(master, |draft| draft.set_color(Slot::Background1, role))
            .unwrap()
    );
    let commit = edit.commit().unwrap();
    assert!(commit.diagnostics().changed());
    let after = commit.package();
    let map = after
        .edit_color_maps()
        .unwrap()
        .color_map(master)
        .unwrap()
        .map()
        .unwrap();
    assert_eq!(map.color(Slot::Background1), role);
    assert_eq!(map.color(Slot::Text1), before.color(Slot::Text1));

    let layout = package.presentation().unwrap().slide_layouts().unwrap()[0]
        .id()
        .unwrap();
    let mut edit = package.edit_color_maps().unwrap();
    assert!(edit.color_map(0_usize).is_ok());
    assert!(edit.color_map(layout).is_ok());
    assert!(matches!(
        edit.update(litchi_pptx::slide::MasterId(1), |_| Ok(())),
        Err(litchi_pptx::Error::Refused(refusal))
            if refusal.kind() == litchi_core::edit::RefusalKind::NotFound
    ));
}

#[test]
fn comment_collaboration_edits_refuse_a_missing_author() {
    let package = deck();
    let mut edit = package.edit_author_presence().unwrap();
    assert!(edit.presence(0).unwrap().is_none());
    assert!(edit.remove(0).unwrap().is_none());
    let value = litchi_pptx::PresenceInfo {
        user_id: "user".to_owned(),
        provider_id: "AD".to_owned(),
    };
    assert!(edit.set(0, value).is_err());
    let commit = edit.commit().unwrap();
    assert!(!commit.diagnostics().changed());

    let mut edit = package.edit_comment_threading().unwrap();
    assert!(edit.remove(0_usize, 0, 1).unwrap().is_none());
}
