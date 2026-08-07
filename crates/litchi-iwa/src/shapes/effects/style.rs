//! Copy-on-write shape-style lifecycle for object effects.

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
use super::native::{
    opacity_from_native, opacity_to_native, reflection_from_native, reflection_to_native,
};
use litchi_iwa_common::shape::effects::{Effects, Opacity, Reflection};

const MAX_STYLE_INHERITANCE_DEPTH: usize = 64;

pub(crate) fn shape_effects(
    package: &IWorkPackage,
    archive_name: &str,
    drawable_id: u64,
) -> Result<Effects> {
    let style_id = shape_payload(package, archive_name, drawable_id)?
        .super_
        .style
        .as_ref()
        .map(|reference| reference.identifier)
        .ok_or_else(|| Error::InvalidFormat(format!("iWork shape {drawable_id} has no style")))?;
    inherited_shape_effects(package, style_id)
}

pub(crate) fn set_shape_effects(
    package: &mut IWorkPackage,
    archive_name: &str,
    drawable_id: u64,
    effects: Effects,
) -> Result<()> {
    if shape_effects(package, archive_name, drawable_id)? == effects {
        return Ok(());
    }
    let old_style_id = shape_payload(package, archive_name, drawable_id)?
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

    if exclusive {
        let parent_style_id = parent_style_id(old_style_id, &old_style)?;
        let mut overrides = direct.ok_or_else(|| {
            Error::InvalidFormat(format!(
                "iWork effect variation {old_style_id} lost its direct overrides"
            ))
        })?;
        apply_effects(&mut overrides, effects);
        let replacement =
            shape_style_variation_object(old_style_id, parent_style_id, stylesheet_id, overrides)?;
        let mut staged = package.clone();
        replace_style_variation(&mut staged, &style_archive_name, old_style_id, replacement)?;
        validate_effects(&staged, archive_name, drawable_id, effects)?;
        *package = staged;
        return Ok(());
    }

    insert_effect_variation(
        package,
        EffectVariationLocation {
            drawable_archive_name: archive_name,
            drawable_id,
            style_archive_name: &style_archive_name,
            stylesheet_id,
            parent_style_id: old_style_id,
        },
        effect_overrides(effects),
        effects,
    )
}

/// Remove direct opacity and reflection overrides and restore inheritance.
pub(crate) fn reset_shape_effects(
    package: &mut IWorkPackage,
    archive_name: &str,
    drawable_id: u64,
) -> Result<bool> {
    let old_style_id = shape_payload(package, archive_name, drawable_id)?
        .super_
        .style
        .as_ref()
        .map(|reference| reference.identifier)
        .ok_or_else(|| Error::InvalidFormat(format!("iWork shape {drawable_id} has no style")))?;
    let style_archive_name = object_archive_name(package, old_style_id)?;
    let old_style_message = shape_style_message(package, &style_archive_name, old_style_id)?;
    let old_style = tswp::ShapeStyleArchive::decode(old_style_message.data.as_slice())?;
    let has_direct_effect = old_style
        .super_
        .shape_properties
        .as_ref()
        .is_some_and(|properties| properties.opacity.is_some() || properties.reflection.is_some());
    if !has_direct_effect {
        return Ok(false);
    }
    let mut direct = direct_shape_style_overrides(&old_style, &old_style_message.data)?
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "iWork shape style {old_style_id} has effects plus unsupported direct overrides"
            ))
        })?;
    let stylesheet_id = stylesheet_id(package, &style_archive_name, old_style_id, &old_style)?;
    let parent_style_id = parent_style_id(old_style_id, &old_style)?;
    let inherited = inherited_shape_effects(package, parent_style_id)?;

    if shape_style_is_exclusive(package, old_style_id)? {
        direct.opacity = None;
        direct.reflection = None;
        let mut staged = package.clone();
        if direct.is_empty() {
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
        }
        validate_effects(&staged, archive_name, drawable_id, inherited)?;
        *package = staged;
        return Ok(true);
    }

    // Preserve a shared variation and mask only this drawable with the
    // parent's effective values in a new private child.
    insert_effect_variation(
        package,
        EffectVariationLocation {
            drawable_archive_name: archive_name,
            drawable_id,
            style_archive_name: &style_archive_name,
            stylesheet_id,
            parent_style_id: old_style_id,
        },
        effect_overrides(inherited),
        inherited,
    )?;
    Ok(true)
}

fn effect_overrides(effects: Effects) -> ShapeStyleOverrides {
    let mut overrides = ShapeStyleOverrides::default();
    apply_effects(&mut overrides, effects);
    overrides
}

fn apply_effects(overrides: &mut ShapeStyleOverrides, effects: Effects) {
    overrides.opacity = Some(opacity_to_native(effects.opacity()));
    overrides.reflection = Some(reflection_to_native(effects.reflection()));
}

struct EffectVariationLocation<'a> {
    drawable_archive_name: &'a str,
    drawable_id: u64,
    style_archive_name: &'a str,
    stylesheet_id: u64,
    parent_style_id: u64,
}

fn insert_effect_variation(
    package: &mut IWorkPackage,
    location: EffectVariationLocation<'_>,
    overrides: ShapeStyleOverrides,
    expected: Effects,
) -> Result<()> {
    let EffectVariationLocation {
        drawable_archive_name,
        drawable_id,
        style_archive_name,
        stylesheet_id,
        parent_style_id,
    } = location;
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
    if let Some(style_component) = component_identifier_for_entry(&staged, style_archive_name)? {
        add_component_object_uuids(&mut staged, style_component, &[new_style_id])?;
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
    }
    set_package_last_object_identifier(&mut staged, new_style_id)?;
    validate_effects(&staged, drawable_archive_name, drawable_id, expected)?;
    *package = staged;
    Ok(())
}

fn validate_effects(
    package: &IWorkPackage,
    archive_name: &str,
    drawable_id: u64,
    expected: Effects,
) -> Result<()> {
    if shape_effects(package, archive_name, drawable_id)? != expected {
        return Err(Error::InvalidFormat(
            "iWork shape effect update failed validation".to_owned(),
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

fn inherited_shape_effects(package: &IWorkPackage, first_style_id: u64) -> Result<Effects> {
    let mut visited = HashSet::new();
    let mut style_id = Some(first_style_id);
    let mut opacity = None;
    let mut reflection = None;
    for _ in 0..MAX_STYLE_INHERITANCE_DEPTH {
        let Some(identifier) = style_id else {
            return Ok(Effects::new(
                opacity.unwrap_or(Opacity::OPAQUE),
                reflection.unwrap_or(Reflection::Disabled),
            ));
        };
        if !visited.insert(identifier) {
            return Err(Error::InvalidFormat(format!(
                "iWork shape style inheritance cycles at {identifier}"
            )));
        }
        let archive_name = object_archive_name(package, identifier)?;
        let style = shape_style(package, &archive_name, identifier)?;
        if let Some(properties) = style.super_.shape_properties.as_ref() {
            if opacity.is_none()
                && let Some(native) = properties.opacity
            {
                opacity = Some(opacity_from_native(native)?);
            }
            if reflection.is_none()
                && let Some(native) = properties.reflection.as_ref()
            {
                reflection = Some(reflection_from_native(native)?);
            }
        }
        if let (Some(opacity), Some(reflection)) = (opacity, reflection) {
            return Ok(Effects::new(opacity, reflection));
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
