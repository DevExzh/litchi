//! Picture geometry, image replacement and descriptions.

use litchi_core::Position;

use super::ShapePropertiesReplacement;
use crate::{Package, Result};

const PNG_A: &[u8] = b"\x89PNG\r\n\x1a\nfirst-image";
const PNG_B: &[u8] = b"\x89PNG\r\n\x1a\nsecond-image";

fn resource(data: &[u8]) -> crate::media_parts::Resource {
    crate::media_parts::Resource::new("/ppt/media/image1.png", "image/png", data.to_vec())
}

fn deck() -> Result<Package> {
    let source = Package::new()?;
    let mut edit = source.edit()?;
    edit.insert_slide(Position::new(0), "Blank")?;
    edit.add_picture(0usize, "Logo", &resource(PNG_A), (0, 0, 914_400, 914_400))?;
    Ok(edit.commit()?.into_package())
}

fn slide_xml(package: &Package) -> Result<Vec<u8>> {
    let mut edit = package.edit()?;
    let part = edit.slides()?[0].part_name().clone();
    Ok(package.opc()?.get_part(&part)?.blob().to_vec())
}

fn media(package: &Package) -> Result<Vec<(String, Vec<u8>)>> {
    let opc = package.opc()?;
    let names: Vec<_> = opc
        .iter_parts()
        .map(|part| part.partname().clone())
        .filter(|name| name.as_str().starts_with("/ppt/media/"))
        .collect();
    names
        .into_iter()
        .map(|name| Ok((name.to_string(), opc.get_part(&name)?.blob().to_vec())))
        .collect()
}

#[test]
fn picture_bounds_are_set_lexically() -> Result<()> {
    let package = deck()?;
    let mut edit = package.edit()?;
    let changed = edit.set_shape_properties(
        0usize,
        &[ShapePropertiesReplacement::new(Position::new(0))
            .with_bounds((100, 200, 300_000, 400_000))],
    )?;
    assert_eq!(changed, 1);
    let after = edit.commit()?.into_package();
    let xml = slide_xml(&after)?;
    let scene = crate::shape::Scene::read(&xml, crate::shape::Limits::DEFAULT)?;
    let bounds = scene.at(0)?.common().bounds().expect("bounds");
    assert_eq!(
        (bounds.x(), bounds.y(), bounds.width(), bounds.height()),
        (100, 200, 300_000, 400_000)
    );
    Ok(())
}

#[test]
fn picture_images_are_replaced_and_the_old_part_dropped() -> Result<()> {
    let package = deck()?;
    let mut edit = package.edit()?;
    assert!(!edit.set_picture_image(0usize, 0usize, &resource(PNG_A))?);
    assert!(edit.set_picture_image(0usize, 0usize, &resource(PNG_B))?);
    let after = edit.commit()?.into_package();
    let parts = media(&after)?;
    assert_eq!(parts.len(), 1, "{parts:?}");
    assert_eq!(parts[0].1, PNG_B);
    // The picture reads its new image back.
    let reopened = Package::from_bytes(&after.to_bytes()?)?;
    assert_eq!(media(&reopened)?[0].1, PNG_B);
    Ok(())
}

#[test]
fn an_image_shared_with_another_picture_is_kept() -> Result<()> {
    let package = deck()?;
    let mut edit = package.edit()?;
    edit.add_picture(0usize, "Copy", &resource(PNG_A), (0, 0, 914_400, 914_400))?;
    let two = edit.commit()?.into_package();
    let before = media(&two)?.len();
    let mut edit = two.edit()?;
    edit.set_picture_image(0usize, 1usize, &resource(PNG_B))?;
    let after = edit.commit()?.into_package();
    assert_eq!(media(&after)?.len(), before);
    Ok(())
}

#[test]
fn descriptions_are_set_and_removed() -> Result<()> {
    let package = deck()?;
    let mut edit = package.edit()?;
    assert!(edit.set_shape_description(0usize, 0usize, Some("A \"logo\" & mark"))?);
    let described = edit.commit()?.into_package();
    let xml = slide_xml(&described)?;
    let scene = crate::shape::Scene::read(&xml, crate::shape::Limits::DEFAULT)?;
    assert_eq!(
        scene.at(0)?.description()?.as_deref(),
        Some("A \"logo\" & mark")
    );
    let mut edit = described.edit()?;
    assert!(!edit.set_shape_description(0usize, 0usize, Some("A \"logo\" & mark"))?);
    assert!(edit.set_shape_description(0usize, 0usize, None)?);
    let cleared = edit.commit()?.into_package();
    assert_eq!(slide_xml(&cleared)?, slide_xml(&package)?);
    Ok(())
}

#[test]
fn shapes_are_renamed() -> Result<()> {
    let package = deck()?;
    let mut edit = package.edit()?;
    assert!(edit.set_shape_name(0usize, 0usize, "Logo & mark")?);
    assert!(!edit.set_shape_name(0usize, 0usize, "Logo & mark")?);
    assert!(edit.set_shape_name(0usize, 0usize, "").is_err());
    let renamed = edit.commit()?.into_package();
    let xml = slide_xml(&renamed)?;
    assert!(
        xml.windows(b"name=\"Logo &amp; mark\"".len())
            .any(|window| window == b"name=\"Logo &amp; mark\"")
    );
    Ok(())
}

/// An edit of one shape keeps the space-only text run of another: the
/// commit's compaction drops formatting between elements, never the text of
/// a run.
#[test]
fn an_edit_keeps_space_only_runs_of_other_shapes() -> Result<()> {
    let package = deck()?;
    let mut edit = package.edit()?;
    edit.add_text_box(0usize, "ab", (0, 0, 914_400, 914_400))?;
    let mut package = edit.commit()?.into_package();
    let part = {
        let mut edit = package.edit()?;
        edit.slides()?[0].part_name().clone()
    };
    let xml = String::from_utf8(package.opc()?.get_part(&part)?.blob().to_vec()).unwrap();
    let spaced = xml.replacen(
        r#"<a:r><a:t xml:space="preserve">ab</a:t></a:r>"#,
        "<a:r><a:t>a</a:t></a:r><a:r><a:t> </a:t></a:r><a:r><a:t>b</a:t></a:r>",
        1,
    );
    assert_ne!(spaced, xml);
    std::sync::Arc::make_mut(&mut package.opc)
        .get_part_mut(&part)?
        .set_blob(spaced.into_bytes());
    let package = Package::from_bytes(&package.to_bytes()?)?;
    let mut edit = package.edit()?;
    assert!(edit.set_shape_description(0usize, 0usize, Some("Logo"))?);
    let after = edit.commit()?.into_package();
    let xml = slide_xml(&after)?;
    let text = std::str::from_utf8(&xml).unwrap();
    assert!(
        text.contains("<a:r><a:t>a</a:t></a:r><a:r><a:t> </a:t></a:r><a:r><a:t>b</a:t></a:r>"),
        "{text}"
    );
    Ok(())
}
