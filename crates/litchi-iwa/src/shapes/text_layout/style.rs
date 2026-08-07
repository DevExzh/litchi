//! Copy-on-write shape-style lifecycle for frame-level text layout.

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
    auto_size_from_native, auto_size_to_native, insets_from_native, insets_to_native,
    vertical_alignment_from_native, vertical_alignment_to_native,
};
use litchi_iwa_common::text::layout::{AutoSize, Insets, Layout, VerticalAlignment};

const MAX_STYLE_INHERITANCE_DEPTH: usize = 64;

pub(crate) fn shape_text_layout(
    package: &IWorkPackage,
    archive_name: &str,
    drawable_id: u64,
) -> Result<Layout> {
    let style_id = shape_payload(package, archive_name, drawable_id)?
        .super_
        .style
        .as_ref()
        .map(|reference| reference.identifier)
        .ok_or_else(|| Error::InvalidFormat(format!("iWork shape {drawable_id} has no style")))?;
    inherited_shape_text_layout(package, style_id)
}

pub(crate) fn set_shape_text_layout(
    mut package: IWorkPackage,
    archive_name: &str,
    drawable_id: u64,
    layout: Layout,
) -> Result<IWorkPackage> {
    if shape_text_layout(&package, archive_name, drawable_id)? == layout {
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
                "iWork text-layout variation {old_style_id} lost its direct overrides"
            ))
        })?;
        apply_layout(&mut overrides, layout);
        let replacement =
            shape_style_variation_object(old_style_id, parent_style_id, stylesheet_id, overrides)?;
        replace_style_variation(&mut package, &style_archive_name, old_style_id, replacement)?;
        validate_layout(&package, archive_name, drawable_id, layout)?;
        return Ok(package);
    }

    insert_layout_variation(
        package,
        LayoutVariationLocation {
            drawable_archive_name: archive_name,
            drawable_id,
            style_archive_name: &style_archive_name,
            stylesheet_id,
            parent_style_id: old_style_id,
        },
        layout_overrides(layout),
        layout,
    )
}

/// Remove direct frame-layout overrides and restore inherited values.
pub(crate) fn reset_shape_text_layout(
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
    let Some(mut direct) = direct_shape_style_overrides(&old_style, &old_style_message.data)?
    else {
        return Ok((package, false));
    };
    let has_direct_layout = old_style
        .shape_properties
        .as_ref()
        .is_some_and(|properties| {
            properties.shrink_to_fit.is_some()
                || properties.vertical_alignment.is_some()
                || properties.padding_null.is_some()
                || properties.padding.is_some()
        });
    if !has_direct_layout {
        return Ok((package, false));
    }
    let stylesheet_id = stylesheet_id(&package, &style_archive_name, old_style_id, &old_style)?;
    let parent_style_id = parent_style_id(old_style_id, &old_style)?;
    let inherited = inherited_shape_text_layout(&package, parent_style_id)?;

    if shape_style_is_exclusive(&package, old_style_id)? {
        clear_layout(&mut direct);
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
        validate_layout(&package, archive_name, drawable_id, inherited)?;
        return Ok((package, true));
    }

    package = insert_layout_variation(
        package,
        LayoutVariationLocation {
            drawable_archive_name: archive_name,
            drawable_id,
            style_archive_name: &style_archive_name,
            stylesheet_id,
            parent_style_id: old_style_id,
        },
        layout_overrides(inherited),
        inherited,
    )?;
    Ok((package, true))
}

fn layout_overrides(layout: Layout) -> ShapeStyleOverrides {
    let mut overrides = ShapeStyleOverrides::default();
    apply_layout(&mut overrides, layout);
    overrides
}

fn apply_layout(overrides: &mut ShapeStyleOverrides, layout: Layout) {
    overrides.shrink_to_fit = Some(auto_size_to_native(layout.auto_size()));
    overrides.vertical_alignment = Some(vertical_alignment_to_native(layout.vertical_alignment()));
    overrides.padding_null = None;
    overrides.padding = Some(insets_to_native(layout.insets()));
}

fn clear_layout(overrides: &mut ShapeStyleOverrides) {
    overrides.shrink_to_fit = None;
    overrides.vertical_alignment = None;
    overrides.padding_null = None;
    overrides.padding = None;
}

struct LayoutVariationLocation<'a> {
    drawable_archive_name: &'a str,
    drawable_id: u64,
    style_archive_name: &'a str,
    stylesheet_id: u64,
    parent_style_id: u64,
}

fn insert_layout_variation(
    mut package: IWorkPackage,
    location: LayoutVariationLocation<'_>,
    overrides: ShapeStyleOverrides,
    expected: Layout,
) -> Result<IWorkPackage> {
    let LayoutVariationLocation {
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
    validate_layout(&package, drawable_archive_name, drawable_id, expected)?;
    Ok(package)
}

fn validate_layout(
    package: &IWorkPackage,
    archive_name: &str,
    drawable_id: u64,
    expected: Layout,
) -> Result<()> {
    if shape_text_layout(package, archive_name, drawable_id)? != expected {
        return Err(Error::InvalidFormat(
            "iWork shape text-layout update failed validation".to_owned(),
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

fn inherited_shape_text_layout(package: &IWorkPackage, first_style_id: u64) -> Result<Layout> {
    let mut visited = HashSet::new();
    let mut style_id = Some(first_style_id);
    let mut vertical_alignment = None;
    let mut insets = None;
    let mut auto_size = None;
    for _ in 0..MAX_STYLE_INHERITANCE_DEPTH {
        let Some(identifier) = style_id else {
            return Ok(Layout::new(
                vertical_alignment.unwrap_or(VerticalAlignment::Top),
                insets.unwrap_or(Insets::ZERO),
                auto_size.unwrap_or(AutoSize::Fixed),
            ));
        };
        if !visited.insert(identifier) {
            return Err(Error::InvalidFormat(format!(
                "iWork shape style inheritance cycles at {identifier}"
            )));
        }
        let archive_name = object_archive_name(package, identifier)?;
        let style = shape_style(package, &archive_name, identifier)?;
        if let Some(properties) = style.shape_properties.as_ref() {
            if auto_size.is_none()
                && let Some(value) = properties.shrink_to_fit
            {
                auto_size = Some(auto_size_from_native(value));
            }
            if vertical_alignment.is_none()
                && let Some(value) = properties.vertical_alignment
            {
                vertical_alignment = Some(vertical_alignment_from_native(value)?);
            }
            if insets.is_none() {
                insets = direct_insets(properties)?;
            }
        }
        if let (Some(vertical_alignment), Some(insets), Some(auto_size)) =
            (vertical_alignment, insets, auto_size)
        {
            return Ok(Layout::new(vertical_alignment, insets, auto_size));
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

fn direct_insets(properties: &tswp::ShapeStylePropertiesArchive) -> Result<Option<Insets>> {
    match (properties.padding_null, properties.padding.as_ref()) {
        (Some(true), None) => Ok(Some(Insets::ZERO)),
        (Some(true), Some(_)) => Err(Error::InvalidFormat(
            "iWork shape text padding is both null and populated".to_owned(),
        )),
        (Some(false), None) => Err(Error::InvalidFormat(
            "iWork shape text padding is marked non-null but missing".to_owned(),
        )),
        (_, Some(padding)) => Ok(Some(insets_from_native(padding)?)),
        (None, None) => Ok(None),
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::keynote::KeynoteEditor;
    use crate::numbers::NumbersEditor;
    use crate::pages::PagesEditor;
    use crate::shapes::{DrawablePoint, DrawableSize};
    use litchi_iwa_common::text::layout::{AutoSize, Inset, Insets, Layout, VerticalAlignment};
    use litchi_iwa_text::columns::{Columns, Count};

    const POSITION: DrawablePoint = DrawablePoint { x: 96.0, y: 120.0 };
    const SIZE: DrawableSize = DrawableSize {
        width: 360.0,
        height: 180.0,
    };

    #[test]
    fn malformed_padding_presence_is_rejected() {
        let padding = tswp::PaddingArchive::default();
        assert!(
            direct_insets(&tswp::ShapeStylePropertiesArchive {
                padding_null: Some(true),
                padding: Some(padding),
                ..Default::default()
            })
            .is_err()
        );
        assert!(
            direct_insets(&tswp::ShapeStylePropertiesArchive {
                padding_null: Some(false),
                padding: None,
                ..Default::default()
            })
            .is_err()
        );
        assert_eq!(
            direct_insets(&tswp::ShapeStylePropertiesArchive {
                padding_null: Some(true),
                padding: None,
                ..Default::default()
            })
            .unwrap(),
            Some(Insets::ZERO)
        );
    }

    #[test]
    fn scratch_suite_text_boxes_support_composable_layout_crud() {
        let inset = Inset::from_points(9.0).unwrap();
        let layout = Layout::new(
            VerticalAlignment::Middle,
            Insets::uniform(inset),
            AutoSize::ShrinkToFit,
        );
        let columns = Columns::equal(Count::new(2).unwrap(), None);

        let mut pages = PagesEditor::create_with_text("Body").unwrap();
        let pages_box = pages
            .add_text_box(4, "Pages layout", POSITION, SIZE)
            .unwrap();
        let pages_inherited = pages
            .text_box_text_layout(pages_box.drawable_object_id)
            .unwrap();
        pages
            .set_text_box_columns(pages_box.drawable_object_id, &columns)
            .unwrap();
        pages
            .set_text_box_text_layout(pages_box.drawable_object_id, layout)
            .unwrap();
        let mut pages = PagesEditor::from_bytes(&pages.to_bytes().unwrap()).unwrap();
        assert_eq!(
            pages
                .text_box_text_layout(pages_box.drawable_object_id)
                .unwrap(),
            layout
        );
        assert_eq!(
            pages
                .text_box_columns(pages_box.drawable_object_id)
                .unwrap(),
            columns
        );
        assert!(
            pages
                .reset_text_box_text_layout(pages_box.drawable_object_id)
                .unwrap()
        );
        assert_eq!(
            pages
                .text_box_text_layout(pages_box.drawable_object_id)
                .unwrap(),
            pages_inherited
        );
        assert_eq!(
            pages
                .text_box_columns(pages_box.drawable_object_id)
                .unwrap(),
            columns
        );
        assert!(
            !pages
                .reset_text_box_text_layout(pages_box.drawable_object_id)
                .unwrap()
        );

        let mut numbers = NumbersEditor::create().unwrap();
        let sheet_id = numbers.sheets().unwrap()[0].object_id;
        let numbers_box = numbers
            .add_sheet_text_box(sheet_id, "Numbers layout", POSITION, SIZE)
            .unwrap();
        numbers
            .set_sheet_text_box_text_layout(sheet_id, numbers_box.drawable_object_id, layout)
            .unwrap();
        let mut numbers = NumbersEditor::from_bytes(&numbers.to_bytes().unwrap()).unwrap();
        assert_eq!(
            numbers
                .sheet_text_box_text_layout(sheet_id, numbers_box.drawable_object_id)
                .unwrap(),
            layout
        );
        assert!(
            numbers
                .reset_sheet_text_box_text_layout(sheet_id, numbers_box.drawable_object_id)
                .unwrap()
        );

        let mut keynote = KeynoteEditor::create().unwrap();
        let keynote_box = keynote
            .add_slide_text_box(0, "Keynote layout", POSITION, SIZE)
            .unwrap();
        keynote
            .set_slide_text_box_text_layout(0, keynote_box.drawable_object_id, layout)
            .unwrap();
        let mut keynote = KeynoteEditor::from_bytes(&keynote.to_bytes().unwrap()).unwrap();
        assert_eq!(
            keynote
                .slide_text_box_text_layout(0, keynote_box.drawable_object_id)
                .unwrap(),
            layout
        );
        assert!(
            keynote
                .reset_slide_text_box_text_layout(0, keynote_box.drawable_object_id)
                .unwrap()
        );
    }
}
