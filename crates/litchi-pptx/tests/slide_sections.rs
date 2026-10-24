//! Record 0252: p14 section membership follows slides that are inserted,
//! moved, swapped or reordered through the primary edits.
#![allow(clippy::unwrap_used, clippy::expect_used, reason = "test assertions")]
use std::sync::Arc;

use litchi_core::{OwnedSource, Position};
use litchi_pptx::Package;
use soapberry_zip::office::{ArchiveReader, StreamingArchiveWriter};

const SECTION_URI: &str = "{521415D9-36F7-43E2-AB2F-B90AF26B5E84}";
const P14: &str = "http://schemas.microsoft.com/office/powerpoint/2010/main";

fn blank_layout(package: &Package) -> String {
    package
        .presentation()
        .unwrap()
        .slide_layouts()
        .unwrap()
        .into_iter()
        .find(|layout| layout.kind().unwrap().as_deref() == Some("blank"))
        .unwrap()
        .name()
        .unwrap()
}

/// A three-slide deck whose sections are "a" = first two slides, "b" = third.
fn deck() -> Vec<u8> {
    let mut builder = litchi_pptx::builder::Builder::new().unwrap();
    for _ in 0..3 {
        builder.presentation().unwrap().add_slide().unwrap();
    }
    let package = builder.build().unwrap();
    let mut edit = package.edit().unwrap();
    let ids: Vec<u32> = edit
        .slides()
        .unwrap()
        .iter()
        .map(|slide| slide.id())
        .collect();
    let bytes = package.to_bytes().unwrap();
    let archive =
        ArchiveReader::new(&bytes, soapberry_zip::office::ArchiveLimits::default()).unwrap();
    let mut writer =
        StreamingArchiveWriter::new(soapberry_zip::office::StreamingArchiveLimits::default());
    for name in archive.file_names() {
        let mut data = archive.read(name).unwrap();
        if name == "ppt/presentation.xml" {
            let xml = String::from_utf8(data).unwrap();
            let entry = |id: u32| format!(r#"<p14:sldId id="{id}"/>"#);
            let sections = format!(
                r#"<p:extLst><p:ext uri="{SECTION_URI}"><p14:sectionLst xmlns:p14="{P14}"><p14:section name="a" id="{{11111111-1111-1111-1111-111111111111}}"><p14:sldIdLst>{}{}</p14:sldIdLst></p14:section><p14:section name="b" id="{{22222222-2222-2222-2222-222222222222}}"><p14:sldIdLst>{}</p14:sldIdLst></p14:section></p14:sectionLst></p:ext></p:extLst>"#,
                entry(ids[0]),
                entry(ids[1]),
                entry(ids[2])
            );
            assert!(!xml.contains("<p:extLst>"));
            data = xml
                .replacen(
                    "</p:presentation>",
                    &format!("{sections}</p:presentation>"),
                    1,
                )
                .into_bytes();
        }
        writer.write_stored(name, &data).unwrap();
    }
    writer.finish().unwrap().into_inner()
}

/// Section membership as slide positions (in the presentation order).
fn membership(package: &mut Package) -> Vec<Vec<usize>> {
    let mut edit = package.edit().unwrap();
    let ids: Vec<u32> = edit
        .slides()
        .unwrap()
        .iter()
        .map(|slide| slide.id())
        .collect();
    let bytes = package.to_bytes().unwrap();
    let archive =
        ArchiveReader::new(&bytes, soapberry_zip::office::ArchiveLimits::default()).unwrap();
    let xml = String::from_utf8(archive.read("ppt/presentation.xml").unwrap()).unwrap();
    xml.split("<p14:section ")
        .skip(1)
        .map(|section| {
            section
                .split("</p14:section>")
                .next()
                .unwrap()
                .split("<p14:sldId id=\"")
                .skip(1)
                .map(|entry| {
                    let id: u32 = entry.split('"').next().unwrap().parse().unwrap();
                    ids.iter().position(|candidate| *candidate == id).unwrap()
                })
                .collect()
        })
        .collect()
}

#[test]
fn eager_insertion_joins_the_previous_slides_section() {
    let source = Package::from_bytes(&deck()).unwrap();
    let layout = blank_layout(&source);
    let mut edit = source.edit().unwrap();
    edit.insert_slide(Position::new(2), &layout).unwrap();
    edit.insert_slide(Position::new(0), &layout).unwrap();
    let mut package = edit.commit().unwrap().into_package();
    // New slides at 0 (before "a") and 3 (after the last of "a").
    assert_eq!(membership(&mut package), [vec![0, 1, 2, 3], vec![4]]);
}

#[test]
fn eager_moves_and_reorders_keep_sections_coherent() {
    let source = Package::from_bytes(&deck()).unwrap();
    let mut edit = source.edit().unwrap();
    // Move the first slide to the end: it joins "b".
    assert!(edit.move_slide(Position::new(0), Position::new(2)).unwrap());
    let mut package = edit.commit().unwrap().into_package();
    assert_eq!(membership(&mut package), [vec![0], vec![1, 2]]);
    // A swap keeps each slide in its section, re-sorted.
    let mut edit = package.edit().unwrap();
    assert!(
        edit.swap_slides(Position::new(1), Position::new(2))
            .unwrap()
    );
    let mut swapped = edit.commit().unwrap().into_package();
    assert_eq!(membership(&mut swapped), [vec![0], vec![1, 2]]);
}

#[test]
fn source_insertion_places_the_new_slide_in_a_section() {
    // The source-backed editor maintains the section list (record 0303):
    // a slide inserted at the end joins the section of the slide before it.
    let bytes = deck();
    let layout = blank_layout(&Package::from_bytes(&bytes).unwrap());
    let source =
        litchi_pptx::source::Package::from_read_at(Arc::new(OwnedSource::new(bytes))).unwrap();
    let mut edit = source.edit().unwrap();
    edit.insert_slide(Position::new(3), &layout).unwrap();
    let (updated, patch) = edit.commit().unwrap().into_parts();
    let mut package = Package::from_bytes(&updated.to_bytes().unwrap()).unwrap();
    assert_eq!(membership(&mut package), [vec![0, 1], vec![2, 3]]);
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
