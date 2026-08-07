//! Copy-on-write shape-style lifecycle for text-frame columns.

use std::collections::HashSet;

use prost::Message;

use crate::package_metadata::{
    add_component_external_reference, add_component_object_uuids, component_identifier_for_entry,
    next_object_identifier, set_package_last_object_identifier,
};
use crate::protobuf::tswp;
use crate::text::columns::{from_native, to_native};
use crate::{Error, IWorkPackage, Result};
use litchi_iwa_text::columns::Columns;

use super::line_end::{
    ShapeStyleOverrides, ShapeStyleVariationLocation, collapse_style_variation,
    direct_shape_style_overrides, insert_style_variation, object_archive_name,
    patch_shape_style_reference, replace_style_variation, shape_payload, shape_style,
    shape_style_is_exclusive, shape_style_message, shape_style_variation_object,
};

const MAX_STYLE_INHERITANCE_DEPTH: usize = 64;

pub(crate) fn shape_text_columns(
    package: &IWorkPackage,
    archive_name: &str,
    drawable_id: u64,
) -> Result<Columns> {
    let style_id = shape_payload(package, archive_name, drawable_id)?
        .super_
        .style
        .as_ref()
        .map(|reference| reference.identifier)
        .ok_or_else(|| Error::InvalidFormat(format!("iWork shape {drawable_id} has no style")))?;
    inherited_columns(package, style_id)
}

pub(crate) fn set_shape_text_columns(
    mut package: IWorkPackage,
    archive_name: &str,
    drawable_id: u64,
    columns: &Columns,
) -> Result<IWorkPackage> {
    if &shape_text_columns(&package, archive_name, drawable_id)? == columns {
        return Ok(package);
    }
    let old_style_id = shape_payload(&package, archive_name, drawable_id)?
        .super_
        .style
        .as_ref()
        .map(|reference| reference.identifier)
        .ok_or_else(|| Error::InvalidFormat(format!("iWork shape {drawable_id} has no style")))?;
    let style_archive_name = object_archive_name(&package, old_style_id)?;
    let old_message = shape_style_message(&package, &style_archive_name, old_style_id)?;
    let old_style = tswp::ShapeStyleArchive::decode(old_message.data.as_slice())?;
    let stylesheet_id = stylesheet_id(&package, &style_archive_name, old_style_id, &old_style)?;
    let direct = direct_shape_style_overrides(&old_style, &old_message.data)?;

    if let Some(mut overrides) = direct
        && shape_style_is_exclusive(&package, old_style_id)?
    {
        let parent_style_id = parent_style_id(old_style_id, &old_style)?;
        apply_columns(&mut overrides, columns);
        let replacement =
            shape_style_variation_object(old_style_id, parent_style_id, stylesheet_id, overrides)?;
        replace_style_variation(&mut package, &style_archive_name, old_style_id, replacement)?;
        validate_columns(&package, archive_name, drawable_id, columns)?;
        return Ok(package);
    }

    insert_columns_variation(
        package,
        ColumnsVariationLocation {
            drawable_archive_name: archive_name,
            drawable_id,
            style_archive_name: &style_archive_name,
            stylesheet_id,
            parent_style_id: old_style_id,
        },
        columns_overrides(columns),
        columns,
    )
}

pub(crate) fn reset_shape_text_columns(
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
    let old_message = shape_style_message(&package, &style_archive_name, old_style_id)?;
    let old_style = tswp::ShapeStyleArchive::decode(old_message.data.as_slice())?;
    let Some(mut direct) = direct_shape_style_overrides(&old_style, &old_message.data)? else {
        return Ok((package, false));
    };
    let has_direct = old_style
        .shape_properties
        .as_ref()
        .is_some_and(|properties| {
            properties.columns_null.is_some() || properties.columns.is_some()
        });
    if !has_direct {
        return Ok((package, false));
    }
    let stylesheet_id = stylesheet_id(&package, &style_archive_name, old_style_id, &old_style)?;
    let parent_style_id = parent_style_id(old_style_id, &old_style)?;
    let inherited = inherited_columns(&package, parent_style_id)?;

    if shape_style_is_exclusive(&package, old_style_id)? {
        clear_columns(&mut direct);
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
        validate_columns(&package, archive_name, drawable_id, &inherited)?;
        return Ok((package, true));
    }

    package = insert_columns_variation(
        package,
        ColumnsVariationLocation {
            drawable_archive_name: archive_name,
            drawable_id,
            style_archive_name: &style_archive_name,
            stylesheet_id,
            parent_style_id: old_style_id,
        },
        columns_overrides(&inherited),
        &inherited,
    )?;
    Ok((package, true))
}

fn columns_overrides(columns: &Columns) -> ShapeStyleOverrides {
    ShapeStyleOverrides {
        columns: Some(to_native(columns)),
        ..Default::default()
    }
}

fn apply_columns(overrides: &mut ShapeStyleOverrides, columns: &Columns) {
    overrides.columns_null = None;
    overrides.columns = Some(to_native(columns));
}

fn clear_columns(overrides: &mut ShapeStyleOverrides) {
    overrides.columns_null = None;
    overrides.columns = None;
}

struct ColumnsVariationLocation<'a> {
    drawable_archive_name: &'a str,
    drawable_id: u64,
    style_archive_name: &'a str,
    stylesheet_id: u64,
    parent_style_id: u64,
}

fn insert_columns_variation(
    mut package: IWorkPackage,
    location: ColumnsVariationLocation<'_>,
    overrides: ShapeStyleOverrides,
    expected: &Columns,
) -> Result<IWorkPackage> {
    let new_style_id = next_object_identifier(&package)?;
    let new_style = shape_style_variation_object(
        new_style_id,
        location.parent_style_id,
        location.stylesheet_id,
        overrides,
    )?;
    patch_shape_style_reference(
        &mut package,
        location.drawable_archive_name,
        location.drawable_id,
        location.parent_style_id,
        new_style_id,
    )?;
    insert_style_variation(
        &mut package,
        location.style_archive_name,
        location.stylesheet_id,
        location.parent_style_id,
        new_style_id,
        new_style,
    )?;
    if let Some(style_component) =
        component_identifier_for_entry(&package, location.style_archive_name)?
    {
        add_component_object_uuids(&mut package, style_component, &[new_style_id])?;
        if let Some(drawable_component) =
            component_identifier_for_entry(&package, location.drawable_archive_name)?
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
    validate_columns(
        &package,
        location.drawable_archive_name,
        location.drawable_id,
        expected,
    )?;
    Ok(package)
}

fn inherited_columns(package: &IWorkPackage, first_style_id: u64) -> Result<Columns> {
    let mut visited = HashSet::new();
    let mut style_id = Some(first_style_id);
    for _ in 0..MAX_STYLE_INHERITANCE_DEPTH {
        let Some(identifier) = style_id else {
            return Ok(Columns::default());
        };
        if !visited.insert(identifier) {
            return Err(Error::InvalidFormat(format!(
                "iWork shape style inheritance cycles at {identifier}"
            )));
        }
        let archive_name = object_archive_name(package, identifier)?;
        let style = shape_style(package, &archive_name, identifier)?;
        if let Some(properties) = style.shape_properties.as_ref()
            && let Some(columns) = direct_columns(properties)?
        {
            return Ok(columns);
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

fn direct_columns(properties: &tswp::ShapeStylePropertiesArchive) -> Result<Option<Columns>> {
    match (properties.columns_null, properties.columns.as_ref()) {
        (Some(true), None) => Ok(Some(Columns::default())),
        (Some(true), Some(_)) => Err(Error::InvalidFormat(
            "iWork shape text columns are both null and populated".into(),
        )),
        (Some(false), None) => Err(Error::InvalidFormat(
            "iWork shape text columns are marked non-null but missing".into(),
        )),
        (_, Some(columns)) => Ok(Some(from_native(columns)?)),
        (None, None) => Ok(None),
    }
}

fn stylesheet_id(
    package: &IWorkPackage,
    archive_name: &str,
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
    if object_archive_name(package, identifier)? != archive_name {
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

fn validate_columns(
    package: &IWorkPackage,
    archive_name: &str,
    drawable_id: u64,
    expected: &Columns,
) -> Result<()> {
    if &shape_text_columns(package, archive_name, drawable_id)? != expected {
        return Err(Error::InvalidFormat(
            "iWork text-box column update failed validation".into(),
        ));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::keynote::KeynoteEditor;
    use crate::numbers::NumbersEditor;
    use crate::pages::PagesEditor;
    use crate::shapes::{DrawablePoint, DrawableSize};
    use litchi_iwa_text::columns::{Count, Gap};

    const POSITION: DrawablePoint = DrawablePoint { x: 96.0, y: 120.0 };
    const SIZE: DrawableSize = DrawableSize {
        width: 360.0,
        height: 180.0,
    };

    #[test]
    fn malformed_column_presence_is_rejected() {
        assert!(
            direct_columns(&tswp::ShapeStylePropertiesArchive {
                columns_null: Some(true),
                columns: Some(tswp::ColumnsArchive::default()),
                ..Default::default()
            })
            .is_err()
        );
    }

    #[test]
    fn scratch_suite_text_boxes_support_column_crud() {
        let two = Columns::equal(
            Count::new(2).unwrap(),
            Some(Gap::from_points(12.0).unwrap()),
        );
        let three = Columns::equal(Count::new(3).unwrap(), None);

        let mut pages = PagesEditor::create_with_text("Body").unwrap();
        let pages_box = pages
            .add_text_box(4, "Pages columns", POSITION, SIZE)
            .unwrap();
        pages
            .set_text_box_columns(pages_box.drawable_object_id, &two)
            .unwrap();
        pages
            .set_text_box_columns(pages_box.drawable_object_id, &three)
            .unwrap();
        let mut pages = PagesEditor::from_bytes(&pages.to_bytes().unwrap()).unwrap();
        assert_eq!(
            pages
                .text_box_columns(pages_box.drawable_object_id)
                .unwrap(),
            three
        );
        assert!(
            pages
                .reset_text_box_columns(pages_box.drawable_object_id)
                .unwrap()
        );
        assert_eq!(
            pages
                .text_box_columns(pages_box.drawable_object_id)
                .unwrap(),
            Columns::default()
        );
        assert!(
            !pages
                .reset_text_box_columns(pages_box.drawable_object_id)
                .unwrap()
        );

        let mut numbers = NumbersEditor::create().unwrap();
        let sheet_id = numbers.sheets().unwrap()[0].object_id;
        let numbers_box = numbers
            .add_sheet_text_box(sheet_id, "Numbers columns", POSITION, SIZE)
            .unwrap();
        numbers
            .set_sheet_text_box_columns(sheet_id, numbers_box.drawable_object_id, &two)
            .unwrap();
        let mut numbers = NumbersEditor::from_bytes(&numbers.to_bytes().unwrap()).unwrap();
        assert_eq!(
            numbers
                .sheet_text_box_columns(sheet_id, numbers_box.drawable_object_id)
                .unwrap(),
            two
        );
        assert!(
            numbers
                .reset_sheet_text_box_columns(sheet_id, numbers_box.drawable_object_id)
                .unwrap()
        );

        let mut keynote = KeynoteEditor::create().unwrap();
        let keynote_box = keynote
            .add_slide_text_box(0, "Keynote columns", POSITION, SIZE)
            .unwrap();
        keynote
            .set_slide_text_box_columns(0, keynote_box.drawable_object_id, &three)
            .unwrap();
        let mut keynote = KeynoteEditor::from_bytes(&keynote.to_bytes().unwrap()).unwrap();
        assert_eq!(
            keynote
                .slide_text_box_columns(0, keynote_box.drawable_object_id)
                .unwrap(),
            three
        );
        assert!(
            keynote
                .reset_slide_text_box_columns(0, keynote_box.drawable_object_id)
                .unwrap()
        );
    }
}
