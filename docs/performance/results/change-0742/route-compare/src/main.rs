//! Change 0742 route comparison probe (informational).
//!
//! Builds a media-rich deck pair with the harness corpus's shape (three source
//! slides with eight 256 KiB incompressible PNG pictures on the third, two
//! destination slides with same-named pictures on the first), copies source
//! slide 2 to destination position 1 after destination slide 1 through both
//! the owned route (`plan_cross_slide_copy` + `apply_cross_slide_copy_plan`)
//! and the source-backed route (`SourceBackedPresentationEditor`), and prints
//! every member whose raw local record differs between the two outputs.

use std::collections::BTreeMap;
use std::error::Error;
use std::sync::Arc;

use litchi_core::OwnedSource;
use litchi_pptx::media_parts::Resource;
use litchi_pptx::{Package, SourceBackedPresentation, SourceBackedPresentationEditor};
use soapberry_zip::ZipArchive;

type BoxResult<T> = Result<T, Box<dyn Error>>;

fn payload(index: usize) -> Vec<u8> {
    let mut state = 0x9e37_79b9_u32 ^ (index as u32);
    let mut bytes = Vec::with_capacity(256 * 1024);
    while bytes.len() < 256 * 1024 {
        state ^= state << 13;
        state ^= state >> 17;
        state ^= state << 5;
        bytes.push((state >> 24) as u8);
    }
    bytes[..8].copy_from_slice(b"\x89PNG\r\n\x1a\n");
    bytes
}

fn deck(slides: usize, media_slide: usize, prefix: &str) -> BoxResult<Vec<u8>> {
    let mut authored = Package::new()?;
    let presentation = authored.presentation_mut()?;
    for index in 0..slides {
        let slide = presentation.add_slide()?;
        slide.set_title(&format!("route-compare-{prefix}-title-{index}"));
    }
    let mut package = Package::from_bytes(&authored.to_bytes()?)?;
    let mut edit = package.opened_presentation_transaction()?;
    for index in 0..8 {
        edit.add_picture(
            media_slide,
            format!("route-compare-picture-{index:02}"),
            &Resource::new(
                format!("/ppt/media/route-compare-media-{index:02}.png"),
                "image/png",
                payload(index),
            ),
            (800, 800, 72, 72),
        )?;
    }
    let commit = edit.commit()?;
    package.apply_opened_presentation_commit(commit)?;
    Ok(package.to_bytes()?)
}

#[derive(Debug, PartialEq, Eq)]
struct Member {
    method: u16,
    flags: u16,
    compressed: Vec<u8>,
    local_record: Vec<u8>,
}

fn members(archive: &[u8]) -> BoxResult<BTreeMap<String, Member>> {
    let zip = ZipArchive::from_slice(archive)?;
    let mut output = BTreeMap::new();
    for entry in zip.entries() {
        let entry = entry?;
        let name = String::from_utf8(entry.file_path().as_ref().to_vec())?;
        let local = usize::try_from(entry.local_header_offset())?;
        let data = zip.get_entry(entry.wayfinder())?;
        let (_, end) = data.compressed_data_range();
        output.insert(
            name,
            Member {
                method: u16::from_le_bytes([archive[local + 8], archive[local + 9]]),
                flags: u16::from_le_bytes([archive[local + 6], archive[local + 7]]),
                compressed: data.data().to_vec(),
                local_record: archive[local..usize::try_from(end)?].to_vec(),
            },
        );
    }
    Ok(output)
}

fn main() -> BoxResult<()> {
    let source = deck(3, 2, "source")?;
    let destination = deck(2, 0, "destination")?;

    let owned_source = Package::from_vec(source.clone())?;
    let mut owned_destination = Package::from_vec(destination.clone())?;
    let plan = owned_destination
        .opened_presentation()?
        .plan_cross_slide_copy(&owned_source.opened_presentation()?, 2_usize, 1_usize, 1)?;
    owned_destination.apply_cross_slide_copy_plan(&owned_source, &plan)?;
    let owned = owned_destination.to_bytes()?;

    let backed_source =
        SourceBackedPresentation::from_read_at(Arc::new(OwnedSource::new(source.clone())))?;
    let editor = SourceBackedPresentationEditor::from_read_at(Arc::new(OwnedSource::new(
        destination.clone(),
    )))?;
    let backed_plan = editor.plan_cross_slide_copy(&backed_source, 2, 1, 1)?;
    let mut backed = Vec::new();
    editor.publish_cross_slide_copy_to_stream(&mut backed, &backed_plan)?;

    println!("transfers_source_compressed_media\t{}", plan.transfers_source_compressed_media());
    println!("owned_bytes\t{}\nsource_backed_bytes\t{}", owned.len(), backed.len());
    println!("byte_identical\t{}", owned == backed);
    let owned_members = members(&owned)?;
    let backed_members = members(&backed)?;
    let identical_media = owned_members
        .iter()
        .filter(|(name, member)| {
            name.starts_with("ppt/media/") && backed_members.get(*name) == Some(*member)
        })
        .count();
    let owned_media = owned_members
        .keys()
        .filter(|name| name.starts_with("ppt/media/"))
        .count();
    println!("media_members_identical\t{identical_media}/{owned_media}");
    for (name, member) in &owned_members {
        match backed_members.get(name) {
            None => println!("only-owned\t{name}"),
            Some(other) if other == member => {},
            Some(other) => println!(
                "differs\t{name}\towned(method={},flags={:#06x},compressed={})\tsource_backed(method={},flags={:#06x},compressed={})\tcompressed_equal={}",
                member.method,
                member.flags,
                member.compressed.len(),
                other.method,
                other.flags,
                other.compressed.len(),
                member.compressed == other.compressed
            ),
        }
    }
    for name in backed_members.keys() {
        if !owned_members.contains_key(name) {
            println!("only-source-backed\t{name}");
        }
    }
    let owned_order: Vec<_> = ZipArchive::from_slice(&owned)?
        .entries()
        .map(|entry| entry.map(|entry| String::from_utf8_lossy(entry.file_path().as_ref()).into_owned()))
        .collect::<Result<_, _>>()?;
    let backed_order: Vec<_> = ZipArchive::from_slice(&backed)?
        .entries()
        .map(|entry| entry.map(|entry| String::from_utf8_lossy(entry.file_path().as_ref()).into_owned()))
        .collect::<Result<_, _>>()?;
    println!("central_order_equal\t{}", owned_order == backed_order);
    Ok(())
}
