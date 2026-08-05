//! Copy-on-write shape-style lifecycle for shadows.

use std::collections::HashSet;

use prost::Message;

use crate::package_metadata::{
    add_component_external_reference, add_component_object_uuids, component_identifier_for_entry,
    next_object_identifier, set_package_last_object_identifier,
};
use crate::protobuf::tswp;
use crate::{Error, IWorkPackage, Result};

use super::super::line_end::{
    ShapeStyleOverrides, ShapeStyleVariationLocation, collapse_style_variation,
    direct_shape_style_overrides, insert_style_variation, object_archive_name,
    patch_shape_style_reference, replace_style_variation, shape_payload, shape_style,
    shape_style_is_exclusive, shape_style_message, shape_style_variation_object,
};
use super::Shadow;
use super::native::{shadow_from_native, shadow_to_native};

const MAX_STYLE_INHERITANCE_DEPTH: usize = 64;

pub(crate) fn shape_shadow(
    package: &IWorkPackage,
    archive_name: &str,
    drawable_id: u64,
) -> Result<Shadow> {
    let style_id = shape_payload(package, archive_name, drawable_id)?
        .super_
        .style
        .as_ref()
        .map(|reference| reference.identifier)
        .ok_or_else(|| Error::InvalidFormat(format!("iWork shape {drawable_id} has no style")))?;
    inherited_shape_shadow(package, style_id)
}

pub(crate) fn set_shape_shadow(
    mut package: IWorkPackage,
    archive_name: &str,
    drawable_id: u64,
    shadow: Shadow,
) -> Result<IWorkPackage> {
    if shape_shadow(&package, archive_name, drawable_id)? == shadow {
        return Ok(package);
    }
    let old_style_id = shape_payload(&package, archive_name, drawable_id)?
        .super_
        .style
        .as_ref()
        .map(|reference| reference.identifier)
        .ok_or_else(|| Error::InvalidFormat(format!("iWork shape {drawable_id} has no style")))?;
    let style_archive_name = object_archive_name(&package, old_style_id)?;
    let old_style_message = shape_style_message(&package, &style_archive_name, old_style_id)?;
    let old_style = tswp::ShapeStyleArchive::decode(old_style_message.data.as_slice())?;
    let stylesheet_id = stylesheet_id(&package, &style_archive_name, old_style_id, &old_style)?;
    let direct = direct_shape_style_overrides(&old_style, &old_style_message.data)?;
    let exclusive = direct.is_some() && shape_style_is_exclusive(&package, old_style_id)?;

    if exclusive {
        let parent_style_id = parent_style_id(old_style_id, &old_style)?;
        let mut overrides = direct.ok_or_else(|| {
            Error::InvalidFormat(format!(
                "iWork shadow variation {old_style_id} lost its direct overrides"
            ))
        })?;
        overrides.shadow = Some(shadow_to_native(shadow));
        let replacement =
            shape_style_variation_object(old_style_id, parent_style_id, stylesheet_id, overrides)?;
        replace_style_variation(&mut package, &style_archive_name, old_style_id, replacement)?;
        validate_shadow(&package, archive_name, drawable_id, shadow)?;
        return Ok(package);
    }

    insert_shadow_variation(
        package,
        ShadowVariationLocation {
            drawable_archive_name: archive_name,
            drawable_id,
            style_archive_name: &style_archive_name,
            stylesheet_id,
            parent_style_id: old_style_id,
        },
        shadow_overrides(shadow),
        shadow,
    )
}

/// Remove the direct shadow override and restore inherited shadow state.
pub(crate) fn reset_shape_shadow(
    mut package: IWorkPackage,
    archive_name: &str,
    drawable_id: u64,
) -> Result<(IWorkPackage, bool)> {
    let old_style_id = shape_payload(&package, archive_name, drawable_id)?
        .super_
        .style
        .as_ref()
        .map(|reference| reference.identifier)
        .ok_or_else(|| Error::InvalidFormat(format!("iWork shape {drawable_id} has no style")))?;
    let style_archive_name = object_archive_name(&package, old_style_id)?;
    let old_style_message = shape_style_message(&package, &style_archive_name, old_style_id)?;
    let old_style = tswp::ShapeStyleArchive::decode(old_style_message.data.as_slice())?;
    if old_style
        .super_
        .shape_properties
        .as_ref()
        .is_none_or(|properties| properties.shadow.is_none())
    {
        return Ok((package, false));
    }
    let mut direct = direct_shape_style_overrides(&old_style, &old_style_message.data)?
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "iWork shape style {old_style_id} has a shadow plus unsupported direct overrides"
            ))
        })?;
    let stylesheet_id = stylesheet_id(&package, &style_archive_name, old_style_id, &old_style)?;
    let parent_style_id = parent_style_id(old_style_id, &old_style)?;
    let inherited = inherited_shape_shadow(&package, parent_style_id)?;

    if shape_style_is_exclusive(&package, old_style_id)? {
        direct.shadow = None;
        if direct.is_empty() {
            collapse_style_variation(
                &mut package,
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
            replace_style_variation(&mut package, &style_archive_name, old_style_id, replacement)?;
        }
        validate_shadow(&package, archive_name, drawable_id, inherited)?;
        return Ok((package, true));
    }

    // Preserve a shared variation and mask only this drawable with the
    // parent's effective value in a new private child.
    package = insert_shadow_variation(
        package,
        ShadowVariationLocation {
            drawable_archive_name: archive_name,
            drawable_id,
            style_archive_name: &style_archive_name,
            stylesheet_id,
            parent_style_id: old_style_id,
        },
        shadow_overrides(inherited),
        inherited,
    )?;
    Ok((package, true))
}

fn shadow_overrides(shadow: Shadow) -> ShapeStyleOverrides {
    ShapeStyleOverrides {
        shadow: Some(shadow_to_native(shadow)),
        ..Default::default()
    }
}

struct ShadowVariationLocation<'a> {
    drawable_archive_name: &'a str,
    drawable_id: u64,
    style_archive_name: &'a str,
    stylesheet_id: u64,
    parent_style_id: u64,
}

fn insert_shadow_variation(
    mut package: IWorkPackage,
    location: ShadowVariationLocation<'_>,
    overrides: ShapeStyleOverrides,
    expected: Shadow,
) -> Result<IWorkPackage> {
    let ShadowVariationLocation {
        drawable_archive_name,
        drawable_id,
        style_archive_name,
        stylesheet_id,
        parent_style_id,
    } = location;
    let new_style_id = next_object_identifier(&package)?;
    let new_style =
        shape_style_variation_object(new_style_id, parent_style_id, stylesheet_id, overrides)?;
    patch_shape_style_reference(
        &mut package,
        drawable_archive_name,
        drawable_id,
        parent_style_id,
        new_style_id,
    )?;
    insert_style_variation(
        &mut package,
        style_archive_name,
        stylesheet_id,
        parent_style_id,
        new_style_id,
        new_style,
    )?;
    if let Some(style_component) = component_identifier_for_entry(&package, style_archive_name)? {
        add_component_object_uuids(&mut package, style_component, &[new_style_id])?;
        if let Some(drawable_component) =
            component_identifier_for_entry(&package, drawable_archive_name)?
            && drawable_component != style_component
        {
            add_component_external_reference(
                &mut package,
                drawable_component,
                style_component,
                new_style_id,
            )?;
        }
    }
    set_package_last_object_identifier(&mut package, new_style_id)?;
    validate_shadow(&package, drawable_archive_name, drawable_id, expected)?;
    Ok(package)
}

fn validate_shadow(
    package: &IWorkPackage,
    archive_name: &str,
    drawable_id: u64,
    expected: Shadow,
) -> Result<()> {
    if shape_shadow(package, archive_name, drawable_id)? != expected {
        return Err(Error::InvalidFormat(
            "iWork shape shadow update failed validation".to_owned(),
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

fn inherited_shape_shadow(package: &IWorkPackage, first_style_id: u64) -> Result<Shadow> {
    let mut visited = HashSet::new();
    let mut style_id = Some(first_style_id);
    for _ in 0..MAX_STYLE_INHERITANCE_DEPTH {
        let Some(identifier) = style_id else {
            return Ok(Shadow::Disabled);
        };
        if !visited.insert(identifier) {
            return Err(Error::InvalidFormat(format!(
                "iWork shape style inheritance cycles at {identifier}"
            )));
        }
        let archive_name = object_archive_name(package, identifier)?;
        let style = shape_style(package, &archive_name, identifier)?;
        if let Some(native) = style
            .super_
            .shape_properties
            .as_ref()
            .and_then(|properties| properties.shadow.as_ref())
        {
            return shadow_from_native(native);
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
