//! Copy-on-write shape-style lifecycle for native strokes.

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
use super::ShapeStroke;
use super::native::{empty_stroke_archive, stroke_from_native, stroke_to_native};

const MAX_STYLE_INHERITANCE_DEPTH: usize = 64;

pub(crate) fn shape_stroke(
    package: &IWorkPackage,
    archive_name: &str,
    drawable_id: u64,
) -> Result<Option<ShapeStroke>> {
    let shape = shape_payload(package, archive_name, drawable_id)?;
    let style_id = shape
        .super_
        .style
        .as_ref()
        .map(|reference| reference.identifier)
        .ok_or_else(|| Error::InvalidFormat(format!("iWork shape {drawable_id} has no style")))?;
    inherited_shape_stroke(package, style_id)
}

pub(crate) fn set_shape_stroke(
    package: &mut IWorkPackage,
    archive_name: &str,
    drawable_id: u64,
    stroke: ShapeStroke,
) -> Result<()> {
    if shape_stroke(package, archive_name, drawable_id)? == Some(stroke) {
        return Ok(());
    }
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
    let native = stroke_to_native(stroke);

    if exclusive {
        let parent_style_id = parent_style_id(old_style_id, &old_style)?;
        let mut overrides = direct.ok_or_else(|| {
            Error::InvalidFormat(format!(
                "iWork stroke variation {old_style_id} lost its direct overrides"
            ))
        })?;
        overrides.stroke = Some(native);
        let replacement =
            shape_style_variation_object(old_style_id, parent_style_id, stylesheet_id, overrides)?;
        let mut staged = package.clone();
        replace_style_variation(&mut staged, &style_archive_name, old_style_id, replacement)?;
        validate_stroke(&staged, archive_name, drawable_id, Some(stroke))?;
        *package = staged;
        return Ok(());
    }

    insert_stroke_variation(
        package,
        archive_name,
        drawable_id,
        &style_archive_name,
        stylesheet_id,
        old_style_id,
        ShapeStyleOverrides {
            stroke: Some(native),
            ..Default::default()
        },
        Some(stroke),
    )
}

/// Remove a direct stroke override and restore the inherited appearance.
pub(crate) fn reset_shape_stroke(
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
    let direct_native_stroke = old_style
        .super_
        .shape_properties
        .as_ref()
        .and_then(|properties| properties.stroke.as_ref());
    if direct_native_stroke.is_none() {
        return Ok(false);
    }
    let mut direct = direct_shape_style_overrides(&old_style, &old_style_message.data)?
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "iWork shape style {old_style_id} has a stroke plus unsupported direct overrides"
            ))
        })?;
    let stylesheet_id = stylesheet_id(package, &style_archive_name, old_style_id, &old_style)?;
    let parent_style_id = parent_style_id(old_style_id, &old_style)?;
    let inherited = inherited_shape_stroke(package, parent_style_id)?;

    if shape_style_is_exclusive(package, old_style_id)? {
        direct.stroke = None;
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
        validate_stroke(&staged, archive_name, drawable_id, inherited)?;
        *package = staged;
        return Ok(true);
    }

    // A shared variation cannot be removed for just one drawable. Override it
    // with the parent's effective value in a private child instead.
    let reset_native = inherited.map_or_else(empty_stroke_archive, stroke_to_native);
    insert_stroke_variation(
        package,
        archive_name,
        drawable_id,
        &style_archive_name,
        stylesheet_id,
        old_style_id,
        ShapeStyleOverrides {
            stroke: Some(reset_native),
            ..Default::default()
        },
        inherited,
    )?;
    Ok(true)
}

fn insert_stroke_variation(
    package: &mut IWorkPackage,
    drawable_archive_name: &str,
    drawable_id: u64,
    style_archive_name: &str,
    stylesheet_id: u64,
    parent_style_id: u64,
    overrides: ShapeStyleOverrides,
    expected: Option<ShapeStroke>,
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
    validate_stroke(&staged, drawable_archive_name, drawable_id, expected)?;
    *package = staged;
    Ok(())
}

fn validate_stroke(
    package: &IWorkPackage,
    archive_name: &str,
    drawable_id: u64,
    expected: Option<ShapeStroke>,
) -> Result<()> {
    if shape_stroke(package, archive_name, drawable_id)? != expected {
        return Err(Error::InvalidFormat(
            "iWork shape stroke update failed validation".to_owned(),
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

fn inherited_shape_stroke(
    package: &IWorkPackage,
    first_style_id: u64,
) -> Result<Option<ShapeStroke>> {
    let mut visited = HashSet::new();
    let mut style_id = Some(first_style_id);
    for _ in 0..MAX_STYLE_INHERITANCE_DEPTH {
        let Some(identifier) = style_id else {
            return Ok(None);
        };
        if !visited.insert(identifier) {
            return Err(Error::InvalidFormat(format!(
                "iWork shape style inheritance cycles at {identifier}"
            )));
        }
        let archive_name = object_archive_name(package, identifier)?;
        let style = shape_style(package, &archive_name, identifier)?;
        if let Some(stroke) = style
            .super_
            .shape_properties
            .as_ref()
            .and_then(|properties| properties.stroke.as_ref())
        {
            return stroke_from_native(stroke);
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
