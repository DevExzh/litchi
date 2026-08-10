use litchi_iwa_protos::table_info_codec;

use super::super::{
    FORM_BASED_SHEET_MESSAGE_TYPE, Package, SHEET_MESSAGE_TYPE, table_info_decode_options,
};
use super::error::{map_read_error, map_table_info_codec_error, usize_as_u64};
use super::resolve::{
    local_reference_identifier, repeated_length_payloads, require_declared_reference,
    require_local_reference, sheet_drawable_payloads, singular_length_payload,
    unique_message_index, unique_sheet_message_index, unique_table_info, validate_message_metadata,
};
use super::{Error, LimitKind, Path, Target};

pub(super) fn validate_selected_ownership(source: &Package, target: Target) -> Result<(), Error> {
    let document_object = source
        .state
        .components
        .get_archive("Index/Document.iwa")
        .and_then(|archive| archive.object(1))
        .ok_or(Error::InvalidSource {
            path: Path::Package,
        })?;
    let (document_message_index, document_message) = unique_message_index(
        &document_object.messages,
        super::super::DOCUMENT_MESSAGE_TYPE,
    )?
    .ok_or(Error::InvalidSource {
        path: Path::Package,
    })?;
    validate_message_metadata(document_object, document_message_index)?;
    let mut wire_work = 0usize;
    let maximum_work = source.state.options.archive().max_iwa_stream_bytes();
    charge_work(
        &mut wire_work,
        document_message.data.len().saturating_mul(2),
        maximum_work,
    )?;
    let sheet_payloads = repeated_length_payloads(&document_message.data, 1)?;
    require_local_reference(
        sheet_payloads
            .get(target.sheet_position)
            .ok_or(Error::InvalidSource {
                path: Path::Package,
            })?,
        target.sheet_identifier,
    )?;
    require_declared_reference(
        document_object,
        document_message_index,
        target.sheet_identifier,
        &[1],
    )?;

    let sheet_object = source
        .state
        .components
        .catalog()
        .get_index(target.sheet_component_index)
        .and_then(|component| component.archive().objects.get(target.sheet_object_index))
        .ok_or(Error::InvalidSource {
            path: Path::Package,
        })?;
    if sheet_object.archive_info.identifier != Some(target.sheet_identifier) {
        return Err(Error::InvalidSource {
            path: Path::Package,
        });
    }
    validate_message_metadata(sheet_object, target.sheet_message_index)?;
    let sheet_message = sheet_object
        .messages
        .get(target.sheet_message_index)
        .ok_or(Error::InvalidSource {
            path: Path::Package,
        })?;
    if sheet_message.type_ != target.sheet_message_type {
        return Err(Error::InvalidSource {
            path: Path::Package,
        });
    }
    let drawable_payloads =
        sheet_drawable_payloads(target.sheet_message_type, &sheet_message.data)?;
    require_local_reference(
        drawable_payloads
            .get(target.drawable_position)
            .ok_or(Error::InvalidSource {
                path: Path::Package,
            })?,
        target.drawable_identifier,
    )?;
    let sheet_path: &[u32] = match target.sheet_message_type {
        SHEET_MESSAGE_TYPE => &[2],
        FORM_BASED_SHEET_MESSAGE_TYPE => &[1, 2],
        _ => {
            return Err(Error::InvalidSource {
                path: Path::Package,
            });
        },
    };
    require_declared_reference(
        sheet_object,
        target.sheet_message_index,
        target.drawable_identifier,
        sheet_path,
    )?;

    let info_object = source
        .state
        .components
        .catalog()
        .get_index(target.info_component_index)
        .and_then(|component| component.archive().objects.get(target.info_object_index))
        .ok_or(Error::InvalidSource {
            path: Path::Package,
        })?;
    if info_object.archive_info.identifier != Some(target.drawable_identifier) {
        return Err(Error::InvalidSource {
            path: Path::Package,
        });
    }
    validate_message_metadata(info_object, target.info_message_index)?;
    let info_message =
        info_object
            .messages
            .get(target.info_message_index)
            .ok_or(Error::InvalidSource {
                path: Path::Package,
            })?;
    if info_message.type_ != target.info_message_type {
        return Err(Error::InvalidSource {
            path: Path::Package,
        });
    }
    require_local_reference(
        singular_length_payload(&info_message.data, 2)?,
        target.model_identifier,
    )?;
    require_declared_reference(
        info_object,
        target.info_message_index,
        target.model_identifier,
        &[2],
    )?;
    let model_object = source
        .state
        .components
        .catalog()
        .get_index(target.component_index)
        .and_then(|component| component.archive().objects.get(target.object_index))
        .ok_or(Error::InvalidSource {
            path: Path::Package,
        })?;
    if model_object.archive_info.identifier != Some(target.model_identifier) {
        return Err(Error::InvalidSource {
            path: Path::Package,
        });
    }
    validate_message_metadata(model_object, target.message_index)?;

    let mut owners = 0usize;
    if target.sheet_identifier == 1
        || target.drawable_identifier == 1
        || target.model_identifier == 1
    {
        return Err(Error::InvalidSource {
            path: Path::Package,
        });
    }
    for sheet_payload in sheet_payloads {
        charge_work(&mut wire_work, sheet_payload.len(), maximum_work)?;
        let sheet_identifier = local_reference_identifier(sheet_payload)?;
        if sheet_identifier == 1
            || sheet_identifier == target.drawable_identifier
            || sheet_identifier == target.model_identifier
        {
            return Err(Error::InvalidSource {
                path: Path::Package,
            });
        }
        let sheet = source
            .state
            .index
            .resolve_ref_id(&source.state.components, sheet_identifier)
            .map_err(map_read_error)?
            .ok_or(Error::InvalidSource {
                path: Path::Package,
            })?;
        let sheet_message_index = unique_sheet_message_index(sheet.messages)?;
        let owner_sheet_message =
            sheet
                .messages
                .get(sheet_message_index)
                .ok_or(Error::InvalidSource {
                    path: Path::Package,
                })?;
        let sheet_multiplier = if owner_sheet_message.type_ == FORM_BASED_SHEET_MESSAGE_TYPE {
            4
        } else {
            2
        };
        charge_work(
            &mut wire_work,
            owner_sheet_message
                .data
                .len()
                .saturating_mul(sheet_multiplier),
            maximum_work,
        )?;
        for drawable_payload in
            sheet_drawable_payloads(owner_sheet_message.type_, &owner_sheet_message.data)?
        {
            charge_work(&mut wire_work, drawable_payload.len(), maximum_work)?;
            let drawable_identifier = local_reference_identifier(drawable_payload)?;
            if drawable_identifier == target.sheet_identifier
                || drawable_identifier == target.model_identifier
            {
                return Err(Error::InvalidSource {
                    path: Path::Package,
                });
            }
            let info = source
                .state
                .index
                .resolve_ref_id(&source.state.components, drawable_identifier)
                .map_err(map_read_error)?
                .ok_or(Error::InvalidSource {
                    path: Path::Package,
                })?;
            let Some((_index, message)) = unique_table_info(info)? else {
                continue;
            };
            charge_work(
                &mut wire_work,
                message.data.len().saturating_mul(4),
                maximum_work,
            )?;
            let snapshot = table_info_codec::decode_table_info(
                &message.data,
                table_info_decode_options(&message.data),
            )
            .map_err(map_table_info_codec_error)?;
            let model_identifier = snapshot.table_model().identifier().get();
            if model_identifier == target.sheet_identifier
                || model_identifier == target.drawable_identifier
            {
                return Err(Error::InvalidSource {
                    path: Path::Package,
                });
            }
            if model_identifier == target.model_identifier {
                owners = owners.checked_add(1).ok_or(Error::InvalidSource {
                    path: Path::Package,
                })?;
                if owners > 1 {
                    return Err(Error::InvalidSource {
                        path: Path::Package,
                    });
                }
            }
        }
    }
    if owners != 1 {
        return Err(Error::InvalidSource {
            path: Path::Package,
        });
    }
    Ok(())
}

pub(super) fn charge_work(total: &mut usize, amount: usize, maximum: usize) -> Result<(), Error> {
    *total = total.checked_add(amount).ok_or(Error::LimitExceeded {
        kind: LimitKind::TransactionWork,
        observed: u64::MAX,
        maximum: usize_as_u64(maximum),
        path: Path::Package,
    })?;
    if *total > maximum {
        return Err(Error::LimitExceeded {
            kind: LimitKind::TransactionWork,
            observed: usize_as_u64(*total),
            maximum: usize_as_u64(maximum),
            path: Path::Package,
        });
    }
    Ok(())
}
