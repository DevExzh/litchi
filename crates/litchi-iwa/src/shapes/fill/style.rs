//! Copy-on-write shape-style lifecycle for native fills.

use std::collections::HashSet;

use prost::Message;

use crate::data_reference_registry::{
    add_component_data_reference, remove_component_data_reference,
};
use crate::package_metadata::{
    add_component_external_reference, add_component_object_uuids, component_identifier_for_entry,
    next_object_identifier, set_package_last_object_identifier,
};
use crate::protobuf::tswp;
use crate::{Error, IWorkMediaEditor, IWorkPackage, MediaType, Result};

use super::super::line_end::{
    ShapeStyleOverrides, ShapeStyleVariationLocation, collapse_style_variation,
    direct_shape_style_overrides, insert_style_variation, object_archive_name,
    patch_shape_style_reference, replace_style_variation, shape_payload, shape_style,
    shape_style_is_exclusive, shape_style_message, shape_style_variation_object,
};
use super::super::{DrawableSize, RgbaColor};
use super::native::{fill_from_native, fill_to_native, image_data_identifier};
use super::{ShapeFill, ShapeImageDataIdentifier, ShapeImageFill, ShapeImageFillTechnique};

const MAX_STYLE_INHERITANCE_DEPTH: usize = 64;

pub(crate) fn shape_fill(
    package: &IWorkPackage,
    archive_name: &str,
    drawable_id: u64,
) -> Result<ShapeFill> {
    let shape = shape_payload(package, archive_name, drawable_id)?;
    let style_id = shape
        .super_
        .style
        .as_ref()
        .map(|reference| reference.identifier)
        .ok_or_else(|| Error::InvalidFormat(format!("iWork shape {drawable_id} has no style")))?;
    inherited_shape_fill(package, style_id)
}

pub(crate) fn set_shape_fill(
    package: &mut IWorkPackage,
    archive_name: &str,
    drawable_id: u64,
    fill: &ShapeFill,
) -> Result<()> {
    let current_fill = shape_fill(package, archive_name, drawable_id)?;
    if &current_fill == fill {
        return Ok(());
    }
    let old_data_identifier = image_data_identifier(&current_fill);
    validate_image_asset(package, fill)?;
    let shape = shape_payload(package, archive_name, drawable_id)?;
    let old_style_id = shape
        .super_
        .style
        .as_ref()
        .map(|reference| reference.identifier)
        .ok_or_else(|| Error::InvalidFormat(format!("iWork shape {drawable_id} has no style")))?;
    let style_archive_name = object_archive_name(package, old_style_id)?;
    let old_style_message = shape_style_message(package, &style_archive_name, old_style_id)?;
    let old_style = tswp::ShapeStyleArchive::decode(old_style_message.data.as_slice())?;
    let stylesheet_id = stylesheet_id(package, &style_archive_name, old_style_id, &old_style)?;
    let direct = direct_shape_style_overrides(&old_style, &old_style_message.data)?;
    let exclusive = direct.is_some() && shape_style_is_exclusive(package, old_style_id)?;
    let native = fill_to_native(fill);

    if exclusive {
        let parent_style_id = parent_style_id(old_style_id, &old_style)?;
        let mut overrides = direct.ok_or_else(|| {
            Error::InvalidFormat(format!(
                "iWork fill variation {old_style_id} lost its direct overrides"
            ))
        })?;
        overrides.fill = Some(native);
        let replacement =
            shape_style_variation_object(old_style_id, parent_style_id, stylesheet_id, overrides)?;
        let mut staged = package.clone();
        replace_style_variation(&mut staged, &style_archive_name, old_style_id, replacement)?;
        adjust_style_data_reference(
            &mut staged,
            &style_archive_name,
            old_style_id,
            old_data_identifier,
            image_data_identifier(fill),
        )?;
        validate_fill(&staged, archive_name, drawable_id, fill)?;
        remove_orphaned_image_asset(&mut staged, old_data_identifier)?;
        *package = staged;
        return Ok(());
    }

    insert_fill_variation(
        package,
        archive_name,
        drawable_id,
        &style_archive_name,
        stylesheet_id,
        old_style_id,
        ShapeStyleOverrides {
            fill: Some(native),
            ..Default::default()
        },
        fill,
    )
}

/// Embed image bytes and attach them as one shape's native image fill.
pub(crate) fn set_shape_image_fill_data(
    package: &mut IWorkPackage,
    archive_name: &str,
    drawable_id: u64,
    preferred_filename: &str,
    data: &[u8],
    technique: ShapeImageFillTechnique,
    fill_size: DrawableSize,
    tint: Option<RgbaColor>,
) -> Result<ShapeImageFill> {
    let mut media = IWorkMediaEditor::from_package(package.clone())?;
    let asset = media.insert_unreferenced(preferred_filename, data)?;
    if asset.media_type != MediaType::Image {
        return Err(Error::Bundle(format!(
            "Shape image fills require image data, not {}",
            asset.media_type.name()
        )));
    }
    let mut staged = media.into_package();
    let mut image = ShapeImageFill::embedded(
        ShapeImageDataIdentifier::new(asset.data_identifier)?,
        technique,
        fill_size,
    )?;
    if let Some(tint) = tint {
        image = image.with_tint(tint);
    }
    set_shape_fill(
        &mut staged,
        archive_name,
        drawable_id,
        &ShapeFill::Image(image.clone()),
    )?;
    let package_path = asset.package_path.as_deref().ok_or_else(|| {
        Error::InvalidFormat("Shape image-fill asset is not materialized".to_owned())
    })?;
    if staged.entry(package_path) != Some(data) {
        return Err(Error::InvalidFormat(
            "Shape image-fill insertion failed validation".to_owned(),
        ));
    }
    *package = staged;
    Ok(image)
}

/// Remove a direct fill override and restore the inherited appearance.
pub(crate) fn reset_shape_fill(
    package: &mut IWorkPackage,
    archive_name: &str,
    drawable_id: u64,
) -> Result<bool> {
    let shape = shape_payload(package, archive_name, drawable_id)?;
    let old_style_id = shape
        .super_
        .style
        .as_ref()
        .map(|reference| reference.identifier)
        .ok_or_else(|| Error::InvalidFormat(format!("iWork shape {drawable_id} has no style")))?;
    let style_archive_name = object_archive_name(package, old_style_id)?;
    let old_style_message = shape_style_message(package, &style_archive_name, old_style_id)?;
    let old_style = tswp::ShapeStyleArchive::decode(old_style_message.data.as_slice())?;
    let Some(direct_native_fill) = old_style
        .super_
        .shape_properties
        .as_ref()
        .and_then(|properties| properties.fill.as_ref())
    else {
        return Ok(false);
    };
    let direct_data_identifier = image_data_identifier(&fill_from_native(direct_native_fill)?);
    let mut direct = direct_shape_style_overrides(&old_style, &old_style_message.data)?
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "iWork shape style {old_style_id} has a fill plus unsupported direct overrides"
            ))
        })?;
    let stylesheet_id = stylesheet_id(package, &style_archive_name, old_style_id, &old_style)?;
    let parent_style_id = parent_style_id(old_style_id, &old_style)?;
    let inherited = inherited_shape_fill(package, parent_style_id)?;

    if shape_style_is_exclusive(package, old_style_id)? {
        direct.fill = None;
        let mut staged = package.clone();
        if direct.is_empty() {
            remove_style_data_reference(
                &mut staged,
                &style_archive_name,
                old_style_id,
                direct_data_identifier,
            )?;
            collapse_style_variation(
                &mut staged,
                ShapeStyleVariationLocation {
                    drawable_archive_name: archive_name,
                    style_archive_name: &style_archive_name,
                    drawable_id,
                    stylesheet_id,
                    style_id: old_style_id,
                    parent_style_id,
                },
            )?;
        } else {
            let replacement =
                shape_style_variation_object(old_style_id, parent_style_id, stylesheet_id, direct)?;
            replace_style_variation(&mut staged, &style_archive_name, old_style_id, replacement)?;
            remove_style_data_reference(
                &mut staged,
                &style_archive_name,
                old_style_id,
                direct_data_identifier,
            )?;
        }
        validate_fill(&staged, archive_name, drawable_id, &inherited)?;
        remove_orphaned_image_asset(&mut staged, direct_data_identifier)?;
        *package = staged;
        return Ok(true);
    }

    // A shared variation cannot be removed for just one drawable. Override it
    // with the parent's effective value in a private child instead.
    let reset_native = fill_to_native(&inherited);
    insert_fill_variation(
        package,
        archive_name,
        drawable_id,
        &style_archive_name,
        stylesheet_id,
        old_style_id,
        ShapeStyleOverrides {
            fill: Some(reset_native),
            ..Default::default()
        },
        &inherited,
    )?;
    Ok(true)
}

fn insert_fill_variation(
    package: &mut IWorkPackage,
    drawable_archive_name: &str,
    drawable_id: u64,
    style_archive_name: &str,
    stylesheet_id: u64,
    parent_style_id: u64,
    overrides: ShapeStyleOverrides,
    expected: &ShapeFill,
) -> Result<()> {
    let new_style_id = next_object_identifier(package)?;
    let new_style =
        shape_style_variation_object(new_style_id, parent_style_id, stylesheet_id, overrides)?;
    let mut staged = package.clone();
    patch_shape_style_reference(
        &mut staged,
        drawable_archive_name,
        drawable_id,
        parent_style_id,
        new_style_id,
    )?;
    insert_style_variation(
        &mut staged,
        style_archive_name,
        stylesheet_id,
        parent_style_id,
        new_style_id,
        new_style,
    )?;
    let expected_data_identifier = image_data_identifier(expected);
    if let Some(style_component) = component_identifier_for_entry(&staged, style_archive_name)? {
        add_component_object_uuids(&mut staged, style_component, &[new_style_id])?;
        if let Some(data_identifier) = expected_data_identifier {
            add_component_data_reference(
                &mut staged,
                style_component,
                data_identifier,
                new_style_id,
            )?;
        }
        if let Some(drawable_component) =
            component_identifier_for_entry(&staged, drawable_archive_name)?
            && drawable_component != style_component
        {
            add_component_external_reference(
                &mut staged,
                drawable_component,
                style_component,
                new_style_id,
            )?;
        }
    } else if expected_data_identifier.is_some() {
        return Err(Error::InvalidFormat(
            "iWork stylesheet has no component for an image-fill reference".to_owned(),
        ));
    }
    set_package_last_object_identifier(&mut staged, new_style_id)?;
    validate_fill(&staged, drawable_archive_name, drawable_id, expected)?;
    *package = staged;
    Ok(())
}

pub(crate) fn validate_image_asset(package: &IWorkPackage, fill: &ShapeFill) -> Result<()> {
    let Some(data_identifier) = image_data_identifier(fill) else {
        return Ok(());
    };
    let assets = crate::media::embedded_assets(package)?;
    let asset = assets
        .iter()
        .find(|asset| asset.data_identifier == data_identifier)
        .ok_or_else(|| {
            Error::Bundle(format!(
                "Shape image fill references missing data identifier {data_identifier}"
            ))
        })?;
    if asset.media_type != MediaType::Image || !asset.is_materialized() {
        return Err(Error::Bundle(format!(
            "Shape image fill data identifier {data_identifier} must be a materialized image"
        )));
    }
    Ok(())
}

fn adjust_style_data_reference(
    package: &mut IWorkPackage,
    style_archive_name: &str,
    style_id: u64,
    old_data_identifier: Option<u64>,
    new_data_identifier: Option<u64>,
) -> Result<()> {
    if old_data_identifier == new_data_identifier {
        return Ok(());
    }
    let style_component = component_identifier_for_entry(package, style_archive_name)?
        .ok_or_else(|| Error::InvalidFormat("iWork stylesheet has no component".to_owned()))?;
    if let Some(identifier) = old_data_identifier {
        remove_component_data_reference(package, style_component, identifier, style_id)?;
    }
    if let Some(identifier) = new_data_identifier {
        add_component_data_reference(package, style_component, identifier, style_id)?;
    }
    Ok(())
}

fn remove_style_data_reference(
    package: &mut IWorkPackage,
    style_archive_name: &str,
    style_id: u64,
    data_identifier: Option<u64>,
) -> Result<()> {
    let Some(data_identifier) = data_identifier else {
        return Ok(());
    };
    let style_component = component_identifier_for_entry(package, style_archive_name)?
        .ok_or_else(|| Error::InvalidFormat("iWork stylesheet has no component".to_owned()))?;
    remove_component_data_reference(package, style_component, data_identifier, style_id)
}

pub(crate) fn remove_orphaned_image_asset(
    package: &mut IWorkPackage,
    data_identifier: Option<u64>,
) -> Result<()> {
    let Some(data_identifier) = data_identifier else {
        return Ok(());
    };
    let mut media = IWorkMediaEditor::from_package(package.clone())?;
    if media
        .asset(data_identifier)
        .is_some_and(|asset| !asset.is_referenced())
    {
        media.remove_unreferenced(data_identifier)?;
        *package = media.into_package();
    }
    Ok(())
}

fn validate_fill(
    package: &IWorkPackage,
    archive_name: &str,
    drawable_id: u64,
    expected: &ShapeFill,
) -> Result<()> {
    if &shape_fill(package, archive_name, drawable_id)? != expected {
        return Err(Error::InvalidFormat(
            "iWork shape fill update failed validation".to_owned(),
        ));
    }
    Ok(())
}

fn stylesheet_id(
    package: &IWorkPackage,
    style_archive_name: &str,
    style_id: u64,
    style: &tswp::ShapeStyleArchive,
) -> Result<u64> {
    let identifier = style
        .super_
        .super_
        .stylesheet
        .as_ref()
        .map(|reference| reference.identifier)
        .ok_or_else(|| {
            Error::InvalidFormat(format!("iWork shape style {style_id} has no stylesheet"))
        })?;
    if object_archive_name(package, identifier)? != style_archive_name {
        return Err(Error::InvalidFormat(format!(
            "iWork shape style {style_id} is not stored with stylesheet {identifier}"
        )));
    }
    Ok(identifier)
}

fn parent_style_id(style_id: u64, style: &tswp::ShapeStyleArchive) -> Result<u64> {
    style
        .super_
        .super_
        .parent
        .as_ref()
        .map(|reference| reference.identifier)
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "iWork shape-style variation {style_id} has no parent"
            ))
        })
}

fn inherited_shape_fill(package: &IWorkPackage, first_style_id: u64) -> Result<ShapeFill> {
    let mut visited = HashSet::new();
    let mut style_id = Some(first_style_id);
    for _ in 0..MAX_STYLE_INHERITANCE_DEPTH {
        let Some(identifier) = style_id else {
            return Ok(ShapeFill::None);
        };
        if !visited.insert(identifier) {
            return Err(Error::InvalidFormat(format!(
                "iWork shape style inheritance cycles at {identifier}"
            )));
        }
        let archive_name = object_archive_name(package, identifier)?;
        let style = shape_style(package, &archive_name, identifier)?;
        if let Some(fill) = style
            .super_
            .shape_properties
            .as_ref()
            .and_then(|properties| properties.fill.as_ref())
        {
            return fill_from_native(fill);
        }
        style_id = style
            .super_
            .super_
            .parent
            .map(|reference| reference.identifier);
    }
    Err(Error::InvalidFormat(format!(
        "iWork shape style inheritance exceeds {MAX_STYLE_INHERITANCE_DEPTH} levels"
    )))
}
