use std::collections::HashSet;

use litchi_iwa_common::wire::WireView;

use super::super::{FORM_BASED_SHEET_MESSAGE_TYPE, Package};
use super::error::{map_read_error, map_wire_error};
use super::resolve::{
    canonical_varint, local_reference_identifier, repeated_length_payloads,
    require_declared_reference, resolved_object, singular_length_payload, unique_message_index,
    validate_message_metadata,
};
use super::{
    CATEGORY_OWNER_REFERENCE_MESSAGE_TYPE, Error, GROUP_BY_MESSAGE_TYPE, Path, Settings, Target,
};

pub(in crate::package) fn validate_dependencies(
    source: &Package,
    target: Target,
    before: Settings,
    after: Settings,
) -> Result<(), Error> {
    let path = Path::Table {
        sheet: target.sheet_position,
        table: target.table_position,
    };
    let model_object = source
        .state
        .components
        .catalog()
        .get_index(target.component_index)
        .and_then(|component| component.archive().objects.get(target.object_index))
        .ok_or(Error::InvalidSource {
            path: Path::Package,
        })?;
    let model = model_object
        .messages
        .get(target.message_index)
        .ok_or(Error::InvalidSource {
            path: Path::Package,
        })?;
    let model_view = WireView::parse(&model.data).map_err(map_wire_error)?;
    let header_counts_changed =
        before.header_rows != after.header_rows || before.header_columns != after.header_columns;
    let section_counts_changed = header_counts_changed || before.footer_rows != after.footer_rows;
    let mut active_pivot_or_group = false;
    let mut header_count_dependency = false;
    let mut pivot_seen = false;
    let mut unsupported_pivot_dependency = false;
    let mut group_seen = false;
    let mut other_seen = [false; 3];
    for field in model_view.fields() {
        match field.number() {
            83 if !group_seen && field.wire_type() == 2 => {
                field.validate_canonical_framing().map_err(map_wire_error)?;
                group_seen = true;
                active_pivot_or_group |= !field.payload().is_empty();
                header_count_dependency |= !field.payload().is_empty();
            },
            85 if !pivot_seen && field.wire_type() == 2 => {
                field.validate_canonical_framing().map_err(map_wire_error)?;
                pivot_seen = true;
                let identifier = local_reference_identifier(field.payload())?;
                if identifier == 1
                    || identifier == target.sheet_identifier
                    || identifier == target.drawable_identifier
                    || identifier == target.model_identifier
                {
                    return Err(Error::InvalidSource {
                        path: Path::Package,
                    });
                }
                require_declared_reference(model_object, target.message_index, identifier, &[85])?;
                unsupported_pivot_dependency = true;
                active_pivot_or_group = true;
                header_count_dependency = true;
            },
            81 | 84 | 86 if field.wire_type() == 2 => {
                let slot = match field.number() {
                    81 => 0,
                    84 => 1,
                    86 => 2,
                    _ => {
                        return Err(Error::InvalidSource {
                            path: Path::Package,
                        });
                    },
                };
                if other_seen[slot] {
                    return Err(Error::InvalidSource {
                        path: Path::Package,
                    });
                }
                other_seen[slot] = true;
                field.validate_canonical_framing().map_err(map_wire_error)?;
                header_count_dependency = true;
                if field.number() == 81 {
                    active_pivot_or_group |= deprecated_category_grouping_active(field.payload())?;
                } else if field.number() == 86 {
                    active_pivot_or_group |=
                        category_owner_reference_active(source, target, field.payload())?;
                }
            },
            81 | 83 | 84 | 85 | 86 => {
                return Err(Error::InvalidSource {
                    path: Path::Package,
                });
            },
            _ => {},
        }
    }
    if unsupported_pivot_dependency {
        return Err(Error::UnsupportedDependency { path });
    }
    if header_counts_changed && header_count_dependency {
        return Err(Error::UnsupportedDependency { path });
    }
    if section_counts_changed && active_pivot_or_group {
        return Err(Error::UnsupportedDependency { path });
    }
    if table_info_has_count_dependency(
        source,
        target,
        header_counts_changed,
        section_counts_changed,
    )? {
        return Err(Error::UnsupportedDependency { path });
    }
    let repeating_changed = before.repeating_header_rows_enabled
        != after.repeating_header_rows_enabled
        || before.repeating_header_columns_enabled != after.repeating_header_columns_enabled;
    if repeating_changed {
        let sheet = source
            .state
            .components
            .catalog()
            .get_index(target.sheet_component_index)
            .and_then(|component| component.archive().objects.get(target.sheet_object_index))
            .and_then(|object| object.messages.get(target.sheet_message_index))
            .ok_or(Error::InvalidSource {
                path: Path::Package,
            })?;
        let payload = if target.sheet_message_type == FORM_BASED_SHEET_MESSAGE_TYPE {
            singular_length_payload(&sheet.data, 1)?
        } else {
            &sheet.data
        };
        if WireView::parse(payload)
            .map_err(map_wire_error)?
            .fields()
            .any(|field| field.number() == 4)
        {
            return Err(Error::UnsupportedDependency { path });
        }
    }
    if header_counts_changed && has_rooted_header_name_manager(source)? {
        return Err(Error::UnsupportedDependency { path });
    }
    Ok(())
}

fn table_info_has_count_dependency(
    source: &Package,
    target: Target,
    header_counts_changed: bool,
    section_counts_changed: bool,
) -> Result<bool, Error> {
    let object = source
        .state
        .components
        .catalog()
        .get_index(target.info_component_index)
        .and_then(|component| component.archive().objects.get(target.info_object_index))
        .ok_or(Error::InvalidSource {
            path: Path::Package,
        })?;
    let message = object
        .messages
        .get(target.info_message_index)
        .ok_or(Error::InvalidSource {
            path: Path::Package,
        })?;
    let view = WireView::parse(&message.data).map_err(map_wire_error)?;
    let mut seen = [false; 7];
    let mut header_active = false;
    let mut section_active = false;
    for field in view.fields() {
        let slot_option = match field.number() {
            4 => Some(0),
            5 => Some(1),
            7 => Some(2),
            8 => Some(3),
            15 => Some(4),
            16 => Some(5),
            17 => Some(6),
            _ => None,
        };
        let Some(slot_index) = slot_option else {
            continue;
        };
        if seen[slot_index] {
            return Err(Error::InvalidSource {
                path: Path::Package,
            });
        }
        seen[slot_index] = true;
        field.validate_canonical_framing().map_err(map_wire_error)?;
        match field.number() {
            4 | 5 | 15 | 17 if field.wire_type() == 2 => {
                let identifier = local_reference_identifier(field.payload())?;
                if identifier == 1
                    || identifier == target.sheet_identifier
                    || identifier == target.drawable_identifier
                    || identifier == target.model_identifier
                {
                    return Err(Error::InvalidSource {
                        path: Path::Package,
                    });
                }
                require_declared_reference(
                    object,
                    target.info_message_index,
                    identifier,
                    &[field.number()],
                )?;
                header_active = true;
                section_active |= matches!(field.number(), 5 | 15 | 17);
            },
            7 | 8 if field.wire_type() == 2 => header_active = true,
            16 if field.wire_type() == 0 => {
                let value = canonical_varint(field.payload())?;
                if value > 1 {
                    return Err(Error::InvalidSource {
                        path: Path::Package,
                    });
                }
                header_active |= value == 1;
                section_active |= value == 1;
            },
            _ => {
                return Err(Error::InvalidSource {
                    path: Path::Package,
                });
            },
        }
    }
    Ok((header_counts_changed && header_active) || (section_counts_changed && section_active))
}

pub(in crate::package) fn deprecated_category_grouping_active(
    source: &[u8],
) -> Result<bool, Error> {
    let view = WireView::parse(source).map_err(map_wire_error)?;
    let mut active = false;
    for field in view.fields() {
        field.validate_canonical_framing().map_err(map_wire_error)?;
        if field.number() == 2 {
            if field.wire_type() != 2 {
                return Err(Error::InvalidSource {
                    path: Path::Package,
                });
            }
            active |= group_by_enabled(field.payload())?.ok_or(Error::InvalidSource {
                path: Path::Package,
            })?;
        }
    }
    Ok(active)
}

fn group_by_enabled(source: &[u8]) -> Result<Option<bool>, Error> {
    let view = WireView::parse(source).map_err(map_wire_error)?;
    let mut enabled = None;
    for field in view.fields() {
        field.validate_canonical_framing().map_err(map_wire_error)?;
        if field.number() == 6 {
            if enabled.is_some() || field.wire_type() != 0 {
                return Err(Error::InvalidSource {
                    path: Path::Package,
                });
            }
            let value = canonical_varint(field.payload())?;
            if value > 1 {
                return Err(Error::InvalidSource {
                    path: Path::Package,
                });
            }
            enabled = Some(value == 1);
        }
    }
    Ok(enabled)
}

pub(in crate::package) fn category_owner_reference_active(
    source: &Package,
    target: Target,
    payload: &[u8],
) -> Result<bool, Error> {
    let owner_identifier = local_reference_identifier(payload)?;
    if owner_identifier == 1
        || owner_identifier == target.sheet_identifier
        || owner_identifier == target.drawable_identifier
        || owner_identifier == target.model_identifier
    {
        return Err(Error::InvalidSource {
            path: Path::Package,
        });
    }
    let model_object = source
        .state
        .components
        .catalog()
        .get_index(target.component_index)
        .and_then(|component| component.archive().objects.get(target.object_index))
        .ok_or(Error::InvalidSource {
            path: Path::Package,
        })?;
    require_declared_reference(model_object, target.message_index, owner_identifier, &[86])?;
    let owner = source
        .state
        .index
        .resolve_ref_id(&source.state.components, owner_identifier)
        .map_err(map_read_error)?
        .ok_or(Error::InvalidSource {
            path: Path::Package,
        })?;
    let owner_object = resolved_object(source, owner)?;
    let (owner_message_index, owner_message) =
        unique_message_index(owner.messages, CATEGORY_OWNER_REFERENCE_MESSAGE_TYPE)?.ok_or(
            Error::InvalidSource {
                path: Path::Package,
            },
        )?;
    let references = repeated_length_payloads(&owner_message.data, 1)?;
    validate_message_metadata(owner_object, owner_message_index)?;
    let mut group_identifiers = HashSet::new();
    group_identifiers
        .try_reserve(references.len())
        .map_err(|_allocation| Error::Allocation {
            amount: references.len(),
            path: Path::Package,
        })?;
    for reference in references {
        let identifier = local_reference_identifier(reference)?;
        if identifier == 1
            || identifier == target.sheet_identifier
            || identifier == target.drawable_identifier
            || identifier == target.model_identifier
            || identifier == owner_identifier
            || !group_identifiers.insert(identifier)
        {
            return Err(Error::InvalidSource {
                path: Path::Package,
            });
        }
    }
    require_declared_group_references(owner_object, owner_message_index, &group_identifiers)?;
    let mut active = false;
    for identifier in group_identifiers {
        let group = source
            .state
            .index
            .resolve_ref_id(&source.state.components, identifier)
            .map_err(map_read_error)?
            .ok_or(Error::InvalidSource {
                path: Path::Package,
            })?;
        let group_object = resolved_object(source, group)?;
        let (message_index, message) = unique_message_index(group.messages, GROUP_BY_MESSAGE_TYPE)?
            .ok_or(Error::InvalidSource {
                path: Path::Package,
            })?;
        let enabled = group_by_enabled(&message.data)?.ok_or(Error::InvalidSource {
            path: Path::Package,
        })?;
        validate_message_metadata(group_object, message_index)?;
        active |= enabled;
    }
    Ok(active)
}

fn require_declared_group_references(
    object: &litchi_iwa_core::ArchiveObject,
    message_index: usize,
    identifiers: &HashSet<u64>,
) -> Result<(), Error> {
    let info =
        object
            .archive_info
            .message_infos
            .get(message_index)
            .ok_or(Error::InvalidSource {
                path: Path::Package,
            })?;
    let mut aggregate_seen = HashSet::new();
    let mut field_seen = HashSet::new();
    aggregate_seen
        .try_reserve(identifiers.len())
        .map_err(|_allocation| Error::Allocation {
            amount: identifiers.len(),
            path: Path::Package,
        })?;
    field_seen
        .try_reserve(identifiers.len())
        .map_err(|_allocation| Error::Allocation {
            amount: identifiers.len(),
            path: Path::Package,
        })?;
    for identifier in &info.object_references {
        if identifiers.contains(identifier) && !aggregate_seen.insert(*identifier) {
            return Err(Error::InvalidSource {
                path: Path::Package,
            });
        }
    }
    if aggregate_seen.len() != identifiers.len() {
        return Err(Error::InvalidSource {
            path: Path::Package,
        });
    }
    for field in &info.field_infos {
        for identifier in &field.object_references {
            if identifiers.contains(identifier)
                && (field.path.as_slice() != [1] || !field_seen.insert(*identifier))
            {
                return Err(Error::InvalidSource {
                    path: Path::Package,
                });
            }
        }
    }
    Ok(())
}

/// Return whether the rooted document owns the native header-name index.
///
/// Cell mutations use this narrow topology proof to refuse changes to header
/// coordinates.  The index contains formula-name fragments that cannot be
/// safely regenerated by a scalar storage rewrite.
pub(in crate::package) fn has_rooted_header_name_manager(source: &Package) -> Result<bool, Error> {
    let document_object = source
        .state
        .components
        .get_archive("Index/Document.iwa")
        .and_then(|archive| archive.object(1))
        .ok_or(Error::InvalidSource {
            path: Path::Package,
        })?;
    let (document_index, document_message) = unique_message_index(
        &document_object.messages,
        super::super::DOCUMENT_MESSAGE_TYPE,
    )?
    .ok_or(Error::InvalidSource {
        path: Path::Package,
    })?;
    let legacy = repeated_length_payloads(&document_message.data, 3)?;
    let super_payload = singular_length_payload(&document_message.data, 8)?;
    let primary = repeated_length_payloads(super_payload, 4)?;
    let (engine_payload, engine_path): (&[u8], &[u32]) =
        match (primary.as_slice(), legacy.as_slice()) {
            ([], []) => return Ok(false),
            ([payload], []) => (*payload, &[8, 4]),
            ([], [payload]) => (*payload, &[3]),
            _ => {
                return Err(Error::InvalidSource {
                    path: Path::Package,
                });
            },
        };
    let engine_identifier = local_reference_identifier(engine_payload)?;
    require_declared_reference(
        document_object,
        document_index,
        engine_identifier,
        engine_path,
    )?;
    let engine = source
        .state
        .index
        .resolve_ref_id(&source.state.components, engine_identifier)
        .map_err(map_read_error)?
        .ok_or(Error::InvalidSource {
            path: Path::Package,
        })?;
    let (engine_index, engine_message) =
        unique_message_index(engine.messages, 4_000)?.ok_or(Error::InvalidSource {
            path: Path::Package,
        })?;
    let engine_object = resolved_object(source, engine)?;
    validate_message_metadata(engine_object, engine_index)?;
    let manager_payloads = repeated_length_payloads(&engine_message.data, 14)?;
    let manager_payload = match manager_payloads.as_slice() {
        [] => return Ok(false),
        [payload] => *payload,
        _ => {
            return Err(Error::InvalidSource {
                path: Path::Package,
            });
        },
    };
    let manager_identifier = local_reference_identifier(manager_payload)?;
    require_declared_reference(engine_object, engine_index, manager_identifier, &[14])?;
    let manager = source
        .state
        .index
        .resolve_ref_id(&source.state.components, manager_identifier)
        .map_err(map_read_error)?
        .ok_or(Error::InvalidSource {
            path: Path::Package,
        })?;
    let (manager_index, _message) =
        unique_message_index(manager.messages, 6_366)?.ok_or(Error::InvalidSource {
            path: Path::Package,
        })?;
    validate_message_metadata(resolved_object(source, manager)?, manager_index)?;
    Ok(true)
}
