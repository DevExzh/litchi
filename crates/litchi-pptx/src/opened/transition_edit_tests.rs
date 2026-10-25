//! Slide transitions staged through the eager edit and the source-backed
//! primary edit (record 0269), read back through the transition codec.

use std::sync::Arc;

use litchi_core::Position;

use crate::transition::{Kind, Limits, Morph, Side, Speed, Transition};
use crate::{Package, Result};

fn deck() -> Result<Package> {
    let mut builder = crate::builder::Builder::new()?;
    builder
        .presentation()?
        .add_slide()?
        .add_text_box("alpha", 0, 0, 914_400, 914_400);
    builder.build()
}

fn slide_xml(package: &Package) -> Result<Vec<u8>> {
    let mut edit = package.edit()?;
    let part = edit.slides()?[0].part_name().clone();
    Ok(package.opc()?.get_part(&part)?.blob().to_vec())
}

fn read(package: &Package) -> Result<Option<Transition>> {
    crate::presentation::transition::read_direct_with_limits(&slide_xml(package)?, Limits::DEFAULT)
}

fn set(package: &Package, value: Option<&Transition>) -> Result<(Package, bool)> {
    let mut edit = package.edit()?;
    let changed = edit.set_slide_transition(0_usize, value)?;
    Ok((edit.commit()?.into_package(), changed))
}

#[test]
fn eager_edit_sets_replaces_and_clears_transitions() -> Result<()> {
    let package = deck()?;
    let before = slide_xml(&package)?;
    assert_eq!(read(&package)?, None);

    let push = Transition::new(Kind::Push(Side::Right)).with_speed(Speed::Slow);
    let (pushed, changed) = set(&package, Some(&push))?;
    assert!(changed);
    assert_eq!(
        read(&pushed)?.as_ref().map(Transition::kind),
        Some(push.kind())
    );

    // The same value again is an exact no-op.
    let (again, changed) = set(&pushed, Some(&push))?;
    assert!(!changed);
    assert_eq!(slide_xml(&again)?, slide_xml(&pushed)?);

    // Morph needs its markup-compatibility owner; replacing keeps one owner.
    let morph = Transition::new(Kind::Morph(Morph::ByWord));
    let (morphed, changed) = set(&pushed, Some(&morph))?;
    assert!(changed);
    assert_eq!(
        read(&morphed)?.as_ref().map(Transition::kind),
        Some(morph.kind())
    );
    let xml = String::from_utf8(slide_xml(&morphed)?).expect("UTF-8 slide");
    assert_eq!(
        xml.matches("<p:transition").count(),
        2,
        "choice and fallback"
    );
    assert!(!xml.contains("<p:push"));

    // Clearing removes the direct transition; clearing again changes nothing.
    let (cleared, changed) = set(&morphed, None)?;
    assert!(changed);
    assert_eq!(read(&cleared)?, None);
    let (_, changed) = set(&cleared, None)?;
    assert!(!changed);
    // The source package is untouched throughout.
    assert_eq!(slide_xml(&package)?, before);
    Ok(())
}

#[test]
fn source_edit_matches_the_eager_edit_and_inverts() -> Result<()> {
    let deck = deck()?;
    let bytes = deck.to_bytes()?;
    let push = Transition::new(Kind::Push(Side::Left));
    let (eager, _) = set(&Package::from_bytes(&bytes)?, Some(&push))?;
    let source =
        crate::source::Package::from_read_at(Arc::new(litchi_core::OwnedSource::new(bytes)))?;
    let mut edit = source.edit()?;
    assert!(edit.set_slide_transition(Position::new(0), Some(&push))?);
    let (after, patch) = edit.commit()?.into_parts();
    assert!(!patch.is_noop());
    let reopened = Package::from_bytes(&after.to_bytes()?)?;
    assert_eq!(slide_xml(&reopened)?, slide_xml(&eager)?);
    assert_eq!(
        after.apply(&patch.inverse())?.into_package().to_bytes()?,
        source.to_bytes()?
    );
    // An absent transition cleared is a no-op on the source owner too.
    let mut edit = source.edit()?;
    assert!(!edit.set_slide_transition(Position::new(0), None)?);
    assert!(edit.commit()?.into_parts().1.is_noop());
    Ok(())
}

#[test]
fn out_of_range_slides_leave_the_edit_unchanged() -> Result<()> {
    let package = deck()?;
    let mut edit = package.edit()?;
    let fade = Transition::new(Kind::Fade { black: None });
    assert!(edit.set_slide_transition(3_usize, Some(&fade)).is_err());
    assert!(!edit.commit()?.diagnostics().changed());
    Ok(())
}

#[test]
fn eager_edit_retains_custom_transition_profile() -> Result<()> {
    let bytes = deck()?.to_bytes()?;
    let output_limit = Limits::DEFAULT.with_output_bytes(1).expect("nonzero limit");
    let package = Package::from_bytes_with_options(
        &bytes,
        crate::OpenOptions::default().transition_limits(output_limit),
    )?;
    let mut edit = package.edit()?;
    let push = Transition::new(Kind::Push(Side::Right));
    assert!(matches!(
        edit.set_slide_transition(0_usize, Some(&push)),
        Err(crate::Error::Limit {
            kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::OutputBytes),
            observed,
                maximum: 1,
        }) if observed > 1
    ));
    Ok(())
}

#[test]
fn eager_edit_retains_custom_shape_profile() -> Result<()> {
    let bytes = deck()?.to_bytes()?;
    let defaults = crate::shape::Limits::DEFAULT;
    let shape_limits = crate::shape::Limits::new(
        defaults.input_bytes(),
        defaults.output_bytes(),
        defaults.depth(),
        1,
        defaults.shapes(),
        defaults.retained_text_bytes(),
    )
    .expect("nonzero shape limits");
    let package = Package::from_bytes_with_options(
        &bytes,
        crate::OpenOptions::default().shape_limits(shape_limits),
    )?;
    let mut edit = package.edit()?;
    assert!(
        edit.add_text_box(0_usize, "beta", (0, 0, 914_400, 914_400))
            .is_err()
    );
    Ok(())
}

#[test]
fn borrowed_designer_projection_retains_custom_shape_profile() -> Result<()> {
    let mut builder = crate::builder::Builder::new()?;
    builder
        .presentation()?
        .add_slide()?
        .add_text_box("Designer", 0, 0, 914_400, 914_400)
        .set_designer_properties(crate::shape::designer::DrawingProperties::new())?;
    let bytes = builder.build()?.to_bytes()?;
    let defaults = crate::shape::Limits::DEFAULT;
    let shape_limits = crate::shape::Limits::new(
        defaults.input_bytes(),
        defaults.output_bytes(),
        defaults.depth(),
        1,
        defaults.shapes(),
        defaults.retained_text_bytes(),
    )
    .expect("tight shape profile");
    let package = Package::from_bytes_with_options(
        &bytes,
        crate::OpenOptions::default().shape_limits(shape_limits),
    )?;
    let slides = package.presentation()?.slides()?;
    assert!(matches!(
        slides[0].shape_designer_properties("Designer"),
        Err(crate::Error::Limit { .. })
    ));
    Ok(())
}

#[test]
fn shape_edits_work_after_an_extension_transition() -> Result<()> {
    // Markup compatibility after `cSld` (the Morph owner) leaves the scene
    // in source offsets, so byte-splicing shape edits stay available.
    let morph = Transition::new(Kind::Morph(Morph::ByObject));
    let (morphed, _) = set(&deck()?, Some(&morph))?;
    let xml = slide_xml(&morphed)?;
    assert!(!crate::shape::Scene::read(&xml, crate::shape::Limits::DEFAULT)?.is_rewritten());
    let mut edit = morphed.edit()?;
    edit.add_text_box(0_usize, "beta", (0, 0, 914_400, 914_400))?;
    edit.set_text_properties(
        0_usize,
        &[super::TextAssignment::new(
            super::TextTarget::Shape {
                shape: Position::new(0),
            },
            super::TextProperties::default().with_bold(Some(true)),
        )],
    )?;
    let after = edit.commit()?.into_package();
    let xml = String::from_utf8(slide_xml(&after)?).expect("UTF-8 slide");
    assert!(xml.contains("beta") && xml.contains("b=\"1\""));
    assert_eq!(
        crate::presentation::transition::read_direct_with_limits(xml.as_bytes(), Limits::DEFAULT)?
            .as_ref()
            .map(Transition::kind),
        Some(morph.kind())
    );
    // Markup compatibility inside the shape tree still rewrites the scene.
    let inside = xml.replacen(
        "<p:spTree>",
        "<p:spTree><mc:AlternateContent xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\"><mc:Choice xmlns:p14=\"http://schemas.microsoft.com/office/powerpoint/2010/main\" Requires=\"p14\"/><mc:Fallback/></mc:AlternateContent>",
        1,
    );
    assert!(
        crate::shape::Scene::read(inside.as_bytes(), crate::shape::Limits::DEFAULT)?.is_rewritten()
    );
    Ok(())
}
