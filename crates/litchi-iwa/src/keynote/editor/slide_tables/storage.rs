//! Table geometry, physical dimension, and object-storage mutation.

use std::collections::HashMap;

use super::*;
use crate::numbers::{NumbersTableDimension, NumbersTableDimensionSize};

pub(super) fn set_table_geometry_in_package(
    package: &mut IWorkPackage,
    drawable_object_id: u64,
    geometry: DrawableGeometry,
) -> Result<()> {
    let graph = ObjectGraph::read(package)?;
    let archive_name = graph.archive_name(drawable_object_id)?.to_owned();
    package.update_archive(&archive_name, |archive| {
        let object = archive.object_mut(drawable_object_id).ok_or_else(|| {
            Error::InvalidFormat(format!("Keynote table {drawable_object_id} is missing"))
        })?;
        let message_indexes = object
            .messages
            .iter()
            .enumerate()
            .filter_map(|(index, message)| {
                (message.type_ == TABLE_INFO_MESSAGE_TYPE).then_some(index)
            })
            .collect::<Vec<_>>();
        let [message_index] = message_indexes.as_slice() else {
            return Err(Error::InvalidFormat(format!(
                "Keynote table {drawable_object_id} must contain exactly one table-info payload"
            )));
        };
        let original = object.messages[*message_index].clone();
        let data = transform_length_delimited_field(&original.data, 1, |drawable| {
            crate::shapes::patch_drawable_geometry(drawable, geometry)
        })?;
        object.replace_message(
            *message_index,
            RawMessage {
                type_: original.type_,
                data,
            },
        )?;
        Ok(())
    })
}

pub(super) fn set_uniform_table_dimensions(
    package: &mut IWorkPackage,
    model_object_id: u64,
    rows: usize,
    columns: usize,
    size: DrawableSize,
) -> Result<()> {
    let row_height = NumbersTableDimensionSize::points(size.height / rows as f32)?;
    let column_width = NumbersTableDimensionSize::points(size.width / columns as f32)?;
    for row in 0..rows {
        crate::numbers::editor::set_table_dimension_size_in_package(
            package,
            model_object_id,
            NumbersTableDimension::Row(row),
            row_height,
        )?;
    }
    for column in 0..columns {
        crate::numbers::editor::set_table_dimension_size_in_package(
            package,
            model_object_id,
            NumbersTableDimension::Column(column),
            column_width,
        )?;
    }
    Ok(())
}

pub(super) fn remove_objects(package: &mut IWorkPackage, identifiers: &[u64]) -> Result<()> {
    let graph = ObjectGraph::read(package)?;
    let mut by_archive = HashMap::<String, Vec<u64>>::new();
    for &identifier in identifiers {
        by_archive
            .entry(graph.archive_name(identifier)?.to_owned())
            .or_default()
            .push(identifier);
    }
    for (archive_name, object_ids) in by_archive {
        let component_id = component_identifier_for_entry(package, &archive_name)?;
        if let Some(component_id) = component_id {
            let registered = component_uuid_identifiers(package, component_id)?.unwrap_or_default();
            let uuid_ids = object_ids
                .iter()
                .copied()
                .filter(|identifier| registered.contains(identifier))
                .collect::<Vec<_>>();
            if !uuid_ids.is_empty() {
                remove_component_object_uuids(package, component_id, &uuid_ids)?;
            }
        }
        let mut archive = package.archive(&archive_name)?;
        for identifier in &object_ids {
            archive.remove_object(*identifier).ok_or_else(|| {
                Error::InvalidFormat(format!("Keynote table object {identifier} is missing"))
            })?;
        }
        if archive.objects.is_empty() {
            package.remove_entry(&archive_name).ok_or_else(|| {
                Error::InvalidFormat(format!("Keynote table component {archive_name} is missing"))
            })?;
            if let Some(component_id) = component_id {
                remove_component_registration(package, component_id)?;
            }
        } else {
            package.replace_archive(&archive_name, &archive)?;
        }
    }
    for &identifier in identifiers {
        if package_references_object(package, identifier)? {
            return Err(Error::InvalidFormat(format!(
                "Keynote table object {identifier} remains referenced after deletion"
            )));
        }
    }
    Ok(())
}
