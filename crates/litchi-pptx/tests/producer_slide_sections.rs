#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "the corpus fixtures and assertions intentionally panic"
)]

//! Source-backed slide removal and order changes keep the `PowerPoint` 2010
//! section list (`p14:sectionLst`) consistent (record 0303).
//!
//! A removed slide leaves its section, and a section that loses its last
//! slide stays, as `PowerPoint` keeps an empty section. A moved slide joins
//! the section of its new neighbour; a swap keeps every slide in its section
//! and re-sorts the sections. Inverse patches restore the source bytes.

use std::sync::Arc;

use litchi_core::{OwnedSource, Position};
use litchi_pptx::source::{Package, Patch};
use soapberry_zip::office::ArchiveReader;

const SECTIONS: &[u8] =
    include_bytes!("../../../test-data/libreoffice-core/sd/qa/unit/data/pptx/slide-sections.pptx");
const SECTION_TEST: &[u8] = include_bytes!(
    "../../../test-data/libreoffice-core/sd/qa/unit/data/pptx/slide-section-test.pptx"
);

fn open(bytes: &[u8]) -> Package {
    Package::from_read_at(Arc::new(OwnedSource::new(bytes.to_vec()))).unwrap()
}

fn presentation_xml(bytes: &[u8]) -> String {
    String::from_utf8(
        ArchiveReader::new(bytes, soapberry_zip::office::ArchiveLimits::default())
            .unwrap()
            .read("ppt/presentation.xml")
            .unwrap(),
    )
    .unwrap()
}

/// Each section's slide identifiers, in document order.
fn sections(bytes: &[u8]) -> Vec<Vec<u32>> {
    let xml = presentation_xml(bytes);
    let list = &xml[xml.find("<p14:sectionLst").unwrap()..];
    list.split("<p14:section ")
        .skip(1)
        .map(|section| {
            section
                .split("<p14:sldId id=\"")
                .skip(1)
                .map(|entry| entry[..entry.find('"').unwrap()].parse().unwrap())
                .collect()
        })
        .collect()
}

/// Commit `edit`, check the sections, and prove the local inverse and the
/// durable inverse after reopening restore the source.
fn check(
    source_bytes: &[u8],
    edit: impl FnOnce(&mut litchi_pptx::source::Edit),
    expected: &[&[u32]],
) {
    let source = open(source_bytes);
    let mut staged = source.edit().unwrap();
    edit(&mut staged);
    let (updated, patch) = staged.commit().unwrap().into_parts();
    let output = updated.to_bytes().unwrap();
    assert_eq!(sections(&output), expected);
    assert_eq!(
        updated
            .apply(&patch.inverse())
            .unwrap()
            .into_package()
            .to_bytes()
            .unwrap(),
        source.to_bytes().unwrap()
    );
    let reopened = open(&output);
    let inverse = Patch::from_bytes(&patch.inverse().to_bytes().unwrap()).unwrap();
    let restored = reopened
        .apply(&inverse)
        .unwrap()
        .into_package()
        .to_bytes()
        .unwrap();
    assert_eq!(presentation_xml(&restored), presentation_xml(source_bytes));
    let forward = Patch::from_bytes(&patch.to_bytes().unwrap()).unwrap();
    let replayed = open(source_bytes)
        .apply(&forward)
        .unwrap()
        .into_package()
        .to_bytes()
        .unwrap();
    assert_eq!(presentation_xml(&replayed), presentation_xml(&output));
}

#[test]
fn a_removed_slide_leaves_its_section() {
    // Slide 4 (id 402) opens the second section.
    check(
        SECTIONS,
        |edit| {
            edit.remove_slide(Position::new(3)).unwrap();
        },
        &[&[401, 377, 366], &[403, 404, 309]],
    );
}

#[test]
fn a_section_that_loses_its_last_slide_stays_empty() {
    check(
        SECTION_TEST,
        |edit| {
            edit.remove_slide(Position::new(12)).unwrap();
            edit.remove_slide(Position::new(11)).unwrap();
        },
        &[
            &[256, 257, 258, 259],
            &[260, 261, 262, 263, 264, 265, 266],
            &[],
        ],
    );
}

#[test]
fn a_moved_slide_joins_its_new_neighbours_section() {
    // Slide 3 (id 366, the end of the first section) moves after slide 4
    // (id 402, the start of the second).
    check(
        SECTIONS,
        |edit| {
            edit.move_slide(Position::new(2), Position::new(3)).unwrap();
        },
        &[&[401, 377], &[402, 366, 403, 404, 309]],
    );
}

#[test]
fn moving_a_slide_out_and_back_restores_the_source() {
    let source = open(SECTIONS);
    let mut edit = source.edit().unwrap();
    edit.move_slide(Position::new(2), Position::new(3)).unwrap();
    edit.move_slide(Position::new(3), Position::new(2)).unwrap();
    let (updated, _) = edit.commit().unwrap().into_parts();
    assert_eq!(updated.to_bytes().unwrap(), source.to_bytes().unwrap());
}

#[test]
fn a_swap_keeps_each_slide_in_its_section() {
    // Swapping across sections re-sorts each section; membership is kept.
    check(
        SECTIONS,
        |edit| {
            edit.swap_slides(Position::new(0), Position::new(1))
                .unwrap();
        },
        &[&[377, 401, 366], &[402, 403, 404, 309]],
    );
}

#[test]
fn removal_after_a_move_composes() {
    check(
        SECTIONS,
        |edit| {
            edit.move_slide(Position::new(2), Position::new(3)).unwrap();
            edit.remove_slide(Position::new(0)).unwrap();
        },
        &[&[377], &[402, 366, 403, 404, 309]],
    );
}

#[test]
fn moving_out_and_back_across_a_boundary_changes_only_membership() {
    // Slide 4 (id 402) opens the second section. Moved before slide 3 it
    // joins the first section, and moved back after slide 3 it stays there:
    // the order is the source order, only the membership changed.
    check(
        SECTIONS,
        |edit| {
            edit.move_slide(Position::new(3), Position::new(2)).unwrap();
            edit.move_slide(Position::new(2), Position::new(3)).unwrap();
        },
        &[&[401, 377, 366, 402], &[403, 404, 309]],
    );
}

#[test]
fn an_inserted_slide_joins_the_section_of_the_slide_before_it() {
    let source = open(SECTION_TEST);
    let mut edit = source.edit().unwrap();
    edit.insert_slide(Position::new(4), "Default").unwrap();
    let (updated, patch) = edit.commit().unwrap().into_parts();
    let output = updated.to_bytes().unwrap();
    let sections = sections(&output);
    assert_eq!(sections.len(), 3);
    assert_eq!(sections[0].len(), 5, "{sections:?}");
    assert_eq!(&sections[0][..4], &[256, 257, 258, 259]);
    assert_eq!(sections[1], [260, 261, 262, 263, 264, 265, 266]);
    assert_eq!(
        updated
            .apply(&patch.inverse())
            .unwrap()
            .into_package()
            .to_bytes()
            .unwrap(),
        source.to_bytes().unwrap()
    );
}

#[test]
fn the_eager_editor_removes_the_section_entry_too() {
    let package = litchi_pptx::Package::from_reader(std::io::Cursor::new(SECTIONS)).unwrap();
    let mut edit = package.edit().unwrap();
    edit.remove_slide(3_usize).unwrap();
    let (updated, _) = edit.commit().unwrap().into_parts();
    assert_eq!(
        sections(&updated.to_bytes().unwrap()),
        [vec![401, 377, 366], vec![403, 404, 309]]
    );
}
