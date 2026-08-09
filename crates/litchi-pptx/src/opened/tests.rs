//! Opened-presentation transaction regression tests.

use std::collections::BTreeMap;

use litchi_opc::PackURI;

use super::{History, Limits, Patch};
use crate::{Error, Package, Result};

fn opened_two_slide_package() -> Result<Package> {
    let mut package = Package::new()?;
    {
        let presentation = package.presentation_mut()?;
        let first = presentation.add_slide()?;
        first.set_title("First title");
        first.add_text_box("First body", 10, 20, 300, 400);
        first.set_notes("First notes");
        let second = presentation.add_slide()?;
        second.set_title("Second title");
        second.add_text_box("Second body", 10, 20, 300, 400);
        second.set_notes("Second notes");
    }
    let bytes = package.to_bytes()?;
    Package::from_bytes(&bytes)
}

fn part_states(package: &Package) -> BTreeMap<String, (Vec<u8>, Vec<String>)> {
    package
        .opc
        .iter_parts()
        .map(|part| {
            let mut relationships: Vec<_> = part
                .rels()
                .iter()
                .map(|relationship| {
                    format!(
                        "{}\0{}\0{}\0{}",
                        relationship.r_id(),
                        relationship.reltype(),
                        relationship.target_ref(),
                        relationship.is_external()
                    )
                })
                .collect();
            relationships.sort_unstable();
            (
                part.partname().as_str().to_owned(),
                (part.blob().to_vec(), relationships),
            )
        })
        .collect()
}

#[test]
fn composes_order_shape_and_notes_in_one_durable_inverse_commit() -> Result<()> {
    let mut package = opened_two_slide_package()?;
    let source = package.opened_presentation()?;
    let source_revision = source.revision();
    let before = part_states(&package);
    let mut edit = source.edit();
    assert!(edit.move_slide(0, 1)?);
    assert!(edit.set_shape_text(0_usize, 0_usize, "Edited <second> & title")?);
    assert!(edit.set_notes_text(0_usize, "Edited & second notes")?);
    let commit = edit.commit()?;
    assert_eq!(commit.patch().resource_count(), 3);
    let durable = commit.patch().to_bytes()?;
    let decoded = Patch::from_bytes(&durable)?;
    assert_eq!(&decoded, commit.patch());
    let inverse = decoded.inverse();
    let touched: Vec<_> = decoded
        .resources()
        .map(|name| name.as_str().to_owned())
        .collect();

    let published = package.apply_opened_presentation_patch(&decoded)?;
    assert_eq!(published.slides()[0].name(), "Slide 257");
    let presentation = package.presentation()?;
    assert_eq!(
        presentation
            .slide(0)?
            .ok_or_else(|| Error::Invalid("published slide disappeared".into()))?
            .text()?,
        "Edited <second> & title\nSecond body"
    );
    let notes = package
        .notes()?
        .ok_or_else(|| Error::Invalid("published notes disappeared".into()))?;
    assert_eq!(
        notes.slides()[0].text()?.as_deref(),
        Some("Edited & second notes")
    );
    let after = part_states(&package);
    for (name, state) in &before {
        if !touched.contains(name) {
            assert_eq!(after.get(name), Some(state), "untouched part {name}");
        }
    }

    let bytes = package.to_bytes()?;
    let mut reopened = Package::from_bytes(&bytes)?;
    assert_eq!(reopened.opened_presentation()?.slides()[0].id(), 257);
    let restored = reopened.apply_opened_presentation_patch(&inverse)?;
    assert_eq!(restored.revision(), source_revision);
    assert_eq!(part_states(&reopened), before);
    Ok(())
}

#[test]
fn disjoint_patches_join_and_same_part_edits_conflict() -> Result<()> {
    let mut package = opened_two_slide_package()?;
    let source = package.opened_presentation()?;

    let mut shape = source.edit();
    shape.set_shape_text(0_usize, 0_usize, "Joined shape")?;
    let shape_patch = shape.commit()?.into_patch();

    let mut notes = source.edit();
    notes.set_notes_text(1_usize, "Joined notes")?;
    let notes_patch = notes.commit()?.into_patch();
    assert!(!shape_patch.conflicts_with(&notes_patch));
    let joined = shape_patch.join(&notes_patch)?;
    assert_eq!(joined.resource_count(), 2);
    package.apply_opened_presentation_patch(&joined)?;
    assert!(
        package
            .presentation()?
            .slide(0)?
            .is_some_and(|slide| { slide.text().is_ok_and(|text| text.contains("Joined shape")) })
    );

    let mut competing = source.edit();
    competing.set_shape_text(0_usize, 1_usize, "Competing shape")?;
    let competing = competing.commit()?.into_patch();
    assert!(shape_patch.conflicts_with(&competing));
    assert!(shape_patch.join(&competing).is_err());
    Ok(())
}

#[test]
fn stale_and_unsupported_raw_xml_fail_before_publication() -> Result<()> {
    let mut package = opened_two_slide_package()?;
    let source = package.opened_presentation()?;
    let mut edit = source.edit();
    edit.set_shape_text(0_usize, 0_usize, "Stale target")?;
    let patch = edit.commit()?.into_patch();
    let slide_name = patch
        .resources()
        .find(|name| name.as_str().contains("/slides/"))
        .cloned()
        .ok_or_else(|| Error::Invalid("shape patch has no slide resource".into()))?;
    let mut changed = package.opc.get_part(&slide_name)?.blob().to_vec();
    changed.extend_from_slice(b" ");
    package
        .opc
        .get_part_mut(&slide_name)?
        .set_blob(changed.clone());
    assert!(package.apply_opened_presentation_patch(&patch).is_err());
    assert_eq!(package.opc.get_part(&slide_name)?.blob(), changed);

    let mut package = opened_two_slide_package()?;
    let main = PackURI::new("/ppt/presentation.xml").map_err(Error::Invalid)?;
    let xml = std::str::from_utf8(package.opc.get_part(&main)?.blob())
        .map_err(|error| Error::Xml(error.to_string()))?
        .replacen("</p:sldIdLst>", "<p:ext/></p:sldIdLst>", 1)
        .into_bytes();
    package.opc.get_part_mut(&main)?.set_blob(xml);
    let source = package.opened_presentation()?;
    let mut edit = source.edit();
    assert!(edit.move_slide(0, 1).is_err());
    assert!(!edit.is_changed());

    let mut package = opened_two_slide_package()?;
    let slide = package.opened_presentation()?.slides()[0]
        .part_name()
        .clone();
    let xml = std::str::from_utf8(package.opc.get_part(&slide)?.blob())
        .map_err(|error| Error::Xml(error.to_string()))?
        .replacen(
            "<a:t>First title</a:t>",
            "<a:t><x:v xmlns:x=\"urn:hostile\"/></a:t>",
            1,
        )
        .into_bytes();
    package.opc.get_part_mut(&slide)?.set_blob(xml);
    let source = package.opened_presentation()?;
    let mut edit = source.edit();
    assert!(edit.set_shape_text(0_usize, 0_usize, "rejected").is_err());
    assert!(!edit.is_changed());
    Ok(())
}

#[test]
fn durable_decoder_and_history_enforce_finite_bounds() -> Result<()> {
    let package = opened_two_slide_package()?;
    let source = package.opened_presentation()?;
    let mut edit = source.edit();
    edit.set_shape_text(0_usize, 0_usize, "History")?;
    let patch = edit.commit()?.into_patch();
    let mut trailing = patch.to_bytes()?;
    trailing.push(0);
    assert!(Patch::from_bytes(&trailing).is_err());
    let tiny = Limits::new(1, 16, 1, 1, 1)
        .ok_or_else(|| Error::Invalid("test limits are invalid".into()))?;
    assert!(Patch::from_bytes_with_limits(&patch.to_bytes()?, tiny).is_err());

    let history_limits = Limits::new(
        4_096,
        128 * 1024 * 1024,
        8 * 1024 * 1024,
        2,
        256 * 1024 * 1024,
    )
    .ok_or_else(|| Error::Invalid("test history limits are invalid".into()))?;
    let mut history = History::new(history_limits);
    history.push(patch.clone())?;
    history.push(patch.clone())?;
    history.push(patch.clone())?;
    assert_eq!(history.len(), 2);
    assert!(history.encoded_bytes() > 0);
    assert_eq!(history.pop_inverse(), Some(patch.inverse()));
    assert_eq!(history.len(), 1);
    Ok(())
}

#[test]
fn real_powerpoint_fixture_round_trips_an_exact_no_op() -> Result<()> {
    let bytes = std::fs::read("../../test-data/ooxml/pptx/shapes.pptx")?;
    let mut package = Package::from_bytes(&bytes)?;
    let source = package.opened_presentation()?;
    assert!(source.slides().len() >= 2);
    let original_first = source.slides()[0].id();
    let slide = package
        .presentation()?
        .slide(0)?
        .ok_or_else(|| Error::Invalid("fixture slide disappeared".into()))?;
    let (shape_position, original_text) = slide
        .shapes()?
        .iter()
        .enumerate()
        .find_map(|(position, shape)| {
            shape
                .common()
                .text()
                .map(|text| (position, text.to_owned()))
        })
        .ok_or_else(|| Error::Invalid("fixture slide has no text shape".into()))?;
    let mut edit = source.edit();
    assert!(!edit.move_slide(0, 0)?);
    assert!(!edit.set_shape_text(0_usize, shape_position, &original_text)?);
    let patch = Patch::from_bytes(&edit.commit()?.into_patch().to_bytes()?)?;
    assert!(patch.is_empty());
    package.apply_opened_presentation_patch(&patch)?;
    let saved = package.to_bytes()?;
    let reopened = Package::from_bytes(&saved)?;
    assert_eq!(
        reopened.opened_presentation()?.slides()[0].id(),
        original_first
    );
    Ok(())
}
