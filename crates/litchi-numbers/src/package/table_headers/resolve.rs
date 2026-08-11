use litchi_iwa_common::{
    WireLimits, decode_varint_from_bytes, varint::encoded_len, wire::WireView,
};
use litchi_iwa_core::RawMessage;
use litchi_iwa_protos::{numbers_table_header_settings_codec, table_info_codec};

use super::super::{
    FORM_BASED_SHEET_MESSAGE_TYPE, LEGACY_TABLE_INFO_MESSAGE_TYPE, Package, Resolved,
    SHEET_MESSAGE_TYPE, TABLE_INFO_MESSAGE_TYPE, TABLE_MODEL_MESSAGE_TYPE,
    table_info_decode_options,
};
use super::error::{
    map_header_codec_error, map_read_error, map_table_info_codec_error, map_wire_error,
};
use super::{
    Error, InvalidReason, LEGACY_TABLE_MODEL_MESSAGE_TYPE, MIN_SIGN_EXTENDED_I32, Path, Settings,
    Target,
};
use crate::table::lock::State as LockState;

pub(in crate::package) fn resolve_target(
    source: &Package,
    sheet_position: usize,
    table_position: usize,
) -> Result<Target, Error> {
    let document_object = source
        .state
        .components
        .get_archive("Index/Document.iwa")
        .and_then(|archive| archive.object(1))
        .ok_or(Error::InvalidSource {
            path: Path::Package,
        })?;
    let (_document_index, document_message) = unique_message_index(
        &document_object.messages,
        super::super::DOCUMENT_MESSAGE_TYPE,
    )?
    .ok_or(Error::InvalidSource {
        path: Path::Package,
    })?;
    let sheet_payloads = repeated_length_payloads(&document_message.data, 1)?;
    let sheet_identifier = local_reference_identifier(
        sheet_payloads
            .get(sheet_position)
            .ok_or(Error::SheetNotFound)?,
    )?;
    let sheet = source
        .state
        .index
        .resolve_ref_id(&source.state.components, sheet_identifier)
        .map_err(map_read_error)?
        .ok_or(Error::InvalidSource {
            path: Path::Package,
        })?;
    let sheet_message_index = unique_sheet_message_index(sheet.messages)?;
    let sheet_message = sheet
        .messages
        .get(sheet_message_index)
        .ok_or(Error::InvalidSource {
            path: Path::Package,
        })?;
    let drawable_payloads = sheet_drawable_payloads(sheet_message.type_, &sheet_message.data)?;
    let mut semantic_table = 0usize;
    for (drawable_position, drawable_payload) in drawable_payloads.iter().enumerate() {
        let drawable_identifier = local_reference_identifier(drawable_payload)?;
        let info = source
            .state
            .index
            .resolve_ref_id(&source.state.components, drawable_identifier)
            .map_err(map_read_error)?
            .ok_or(Error::InvalidSource {
                path: Path::Package,
            })?;
        let Some((info_message_index, info_message)) = unique_table_info(info)? else {
            continue;
        };
        let info_snapshot = table_info_codec::decode_table_info(
            &info_message.data,
            table_info_decode_options(&info_message.data),
        )
        .map_err(map_table_info_codec_error)?;
        if semantic_table != table_position {
            semantic_table = semantic_table.checked_add(1).ok_or(Error::InvalidSource {
                path: Path::Package,
            })?;
            continue;
        }
        let model_identifier = info_snapshot.table_model().identifier().get();
        let model = source
            .state
            .index
            .resolve_ref_id(&source.state.components, model_identifier)
            .map_err(map_read_error)?
            .ok_or(Error::InvalidSource {
                path: Path::Package,
            })?;
        let (message_index, message) = unique_table_model(model.messages)?;
        let decoded = decode_settings(&message.data)?;
        let settings = settings_from_snapshot(&decoded)?;
        validate_stored(settings, decoded.rows(), decoded.columns())?;
        if sheet_identifier == drawable_identifier
            || sheet_identifier == model_identifier
            || drawable_identifier == model_identifier
        {
            return Err(Error::InvalidSource {
                path: Path::Package,
            });
        }
        return Ok(Target {
            sheet_position,
            table_position,
            model_identifier,
            sheet_identifier,
            drawable_identifier,
            drawable_position,
            sheet_component_index: sheet.component_index,
            sheet_object_index: sheet.object_index,
            sheet_message_index,
            sheet_message_type: sheet_message.type_,
            info_component_index: info.component_index,
            info_object_index: info.object_index,
            info_message_index,
            info_message_type: info_message.type_,
            component_index: model.component_index,
            object_index: model.object_index,
            message_index,
            message_type: message.type_,
            settings,
            rows: decoded.rows(),
            columns: decoded.columns(),
            locked: LockState::from_locked(info_snapshot.locked().unwrap_or(false)),
        });
    }
    Err(Error::TableNotFound)
}

pub(super) fn settings_at_target(source: &Package, target: Target) -> Result<Settings, Error> {
    let object = source
        .state
        .components
        .catalog()
        .get_index(target.component_index)
        .and_then(|component| component.archive().objects.get(target.object_index))
        .ok_or(Error::InvalidSource {
            path: Path::Package,
        })?;
    if object.archive_info.identifier != Some(target.model_identifier) {
        return Err(Error::InvalidSource {
            path: Path::Package,
        });
    }
    validate_message_metadata(object, target.message_index)?;
    let message = object
        .messages
        .get(target.message_index)
        .ok_or(Error::InvalidSource {
            path: Path::Package,
        })?;
    if message.type_ != target.message_type {
        return Err(Error::InvalidSource {
            path: Path::Package,
        });
    }
    let snapshot = decode_settings(&message.data)?;
    let settings = settings_from_snapshot(&snapshot)?;
    validate_stored(settings, snapshot.rows(), snapshot.columns())?;
    Ok(settings)
}

pub(in crate::package) fn resolved_object<'a>(
    source: &'a Package,
    resolved: Resolved<'a>,
) -> Result<&'a litchi_iwa_core::ArchiveObject, Error> {
    source
        .state
        .components
        .catalog()
        .get_index(resolved.component_index)
        .and_then(|component| component.archive().objects.get(resolved.object_index))
        .ok_or(Error::InvalidSource {
            path: Path::Package,
        })
}

pub(in crate::package) fn unique_sheet_message_index(
    messages: &[RawMessage],
) -> Result<usize, Error> {
    let sheet = unique_message_index(messages, SHEET_MESSAGE_TYPE)?;
    let form = unique_message_index(messages, FORM_BASED_SHEET_MESSAGE_TYPE)?;
    match (sheet, form) {
        (Some(_), Some(_)) | (None, None) => Err(Error::InvalidSource {
            path: Path::Package,
        }),
        (Some((index, _)), None) | (None, Some((index, _))) => Ok(index),
    }
}

pub(in crate::package) fn unique_table_info(
    resolved: Resolved<'_>,
) -> Result<Option<(usize, &RawMessage)>, Error> {
    let canonical = unique_message_index(resolved.messages, TABLE_INFO_MESSAGE_TYPE)?;
    let legacy = unique_message_index(resolved.messages, LEGACY_TABLE_INFO_MESSAGE_TYPE)?;
    match (canonical, legacy) {
        (Some(_), Some(_)) => Err(Error::InvalidSource {
            path: Path::Package,
        }),
        (Some(message), None) | (None, Some(message)) => Ok(Some(message)),
        (None, None) => Ok(None),
    }
}

pub(in crate::package) fn unique_table_model(
    messages: &[RawMessage],
) -> Result<(usize, &RawMessage), Error> {
    let canonical = unique_message_index(messages, TABLE_MODEL_MESSAGE_TYPE)?;
    let legacy = unique_message_index(messages, LEGACY_TABLE_MODEL_MESSAGE_TYPE)?;
    match (canonical, legacy) {
        (Some(_), Some(_)) | (None, None) => Err(Error::InvalidSource {
            path: Path::Package,
        }),
        (Some(message), None) | (None, Some(message)) => Ok(message),
    }
}

pub(in crate::package) fn unique_message_index(
    messages: &[RawMessage],
    message_type: u32,
) -> Result<Option<(usize, &RawMessage)>, Error> {
    let mut matches = messages
        .iter()
        .enumerate()
        .filter(|(_, message)| message.type_ == message_type);
    let first = matches.next();
    if matches.next().is_some() {
        return Err(Error::InvalidSource {
            path: Path::Package,
        });
    }
    Ok(first)
}

pub(super) fn decode_settings(
    source: &[u8],
) -> Result<numbers_table_header_settings_codec::TableHeaderSettingsSnapshot, Error> {
    let options = numbers_table_header_settings_codec::DecodeOptions::new(
        source.len().max(1),
        source.len().max(1),
        source.len().saturating_mul(4).max(1),
        u32::try_from(WireLimits::MAX_NESTING).unwrap_or(u32::MAX),
    );
    numbers_table_header_settings_codec::decode_table_header_settings(source, options)
        .map_err(map_header_codec_error)
}

pub(super) fn settings_from_snapshot(
    snapshot: &numbers_table_header_settings_codec::TableHeaderSettingsSnapshot,
) -> Result<Settings, Error> {
    fn count(raw: Option<u32>) -> Result<Option<crate::table::headers::Count>, Error> {
        raw.map(|raw_count| {
            usize::try_from(raw_count)
                .ok()
                .and_then(|count| crate::table::headers::Count::new(count).ok())
                .ok_or(Error::InvalidSource {
                    path: Path::Package,
                })
        })
        .transpose()
    }
    Ok(Settings {
        header_rows: count(snapshot.header_rows())?,
        header_columns: count(snapshot.header_columns())?,
        footer_rows: count(snapshot.footer_rows())?,
        header_rows_frozen: snapshot.header_rows_frozen(),
        header_columns_frozen: snapshot.header_columns_frozen(),
        repeating_header_rows_enabled: snapshot.repeating_header_rows_enabled(),
        repeating_header_columns_enabled: snapshot.repeating_header_columns_enabled(),
    })
}

pub(super) fn validate_requested(
    settings: Settings,
    rows: u32,
    columns: u32,
    path: Path,
) -> Result<(), Error> {
    let header_rows = u8::try_from(settings.header_row_count()).unwrap_or(u8::MAX);
    let footer_rows = u8::try_from(settings.footer_row_count()).unwrap_or(u8::MAX);
    if u16::from(header_rows).saturating_add(u16::from(footer_rows))
        > u16::try_from(rows).unwrap_or(u16::MAX)
    {
        return Err(Error::InvalidSettings {
            path,
            reason: InvalidReason::RowSectionsExceedTable {
                header_rows,
                footer_rows,
                table_rows: rows,
            },
        });
    }
    let header_columns = u8::try_from(settings.header_column_count()).unwrap_or(u8::MAX);
    if u32::from(header_columns) > columns {
        return Err(Error::InvalidSettings {
            path,
            reason: InvalidReason::HeaderColumnsExceedTable {
                header_columns,
                table_columns: columns,
            },
        });
    }
    Ok(())
}

pub(in crate::package) fn validate_message_metadata(
    object: &litchi_iwa_core::ArchiveObject,
    message_index: usize,
) -> Result<(), Error> {
    let info =
        object
            .archive_info
            .message_infos
            .get(message_index)
            .ok_or(Error::InvalidSource {
                path: Path::Package,
            })?;
    let message = object
        .messages
        .get(message_index)
        .ok_or(Error::InvalidSource {
            path: Path::Package,
        })?;
    if info.type_ != message.type_
        || object.archive_info.should_merge == Some(true)
        || info.base_message_index.is_some()
        || !info.diff_merge_version.is_empty()
        || info.diff_field_path.is_some()
        || !info.fields_to_remove.is_empty()
        || !info.diff_read_version.is_empty()
    {
        return Err(Error::InvalidSource {
            path: Path::Package,
        });
    }
    Ok(())
}

pub(in crate::package) fn require_declared_reference(
    object: &litchi_iwa_core::ArchiveObject,
    message_index: usize,
    identifier: u64,
    accepted_path: &[u32],
) -> Result<(), Error> {
    let info =
        object
            .archive_info
            .message_infos
            .get(message_index)
            .ok_or(Error::InvalidSource {
                path: Path::Package,
            })?;
    if info
        .object_references
        .iter()
        .filter(|candidate| **candidate == identifier)
        .count()
        != 1
    {
        return Err(Error::InvalidSource {
            path: Path::Package,
        });
    }
    let mut field_occurrence = false;
    for field in &info.field_infos {
        let count = field
            .object_references
            .iter()
            .filter(|candidate| **candidate == identifier)
            .count();
        if count != 0 {
            if count != 1 || field_occurrence || field.path.as_slice() != accepted_path {
                return Err(Error::InvalidSource {
                    path: Path::Package,
                });
            }
            field_occurrence = true;
        }
    }
    Ok(())
}

pub(in crate::package) fn repeated_length_payloads(
    source: &[u8],
    field_number: u32,
) -> Result<Vec<&[u8]>, Error> {
    let view = WireView::parse(source).map_err(map_wire_error)?;
    let count = view
        .fields()
        .filter(|field| field.number() == field_number)
        .count();
    let mut values = Vec::new();
    values
        .try_reserve_exact(count)
        .map_err(|_allocation| Error::Allocation {
            amount: count,
            path: Path::Package,
        })?;
    for field in view.fields().filter(|field| field.number() == field_number) {
        if field.wire_type() != 2 {
            return Err(Error::InvalidSource {
                path: Path::Package,
            });
        }
        field.validate_canonical_framing().map_err(map_wire_error)?;
        values.push(field.payload());
    }
    Ok(values)
}

pub(in crate::package) fn singular_length_payload(
    source: &[u8],
    field_number: u32,
) -> Result<&[u8], Error> {
    let values = repeated_length_payloads(source, field_number)?;
    match values.as_slice() {
        [value] => Ok(*value),
        _ => Err(Error::InvalidSource {
            path: Path::Package,
        }),
    }
}

pub(in crate::package) fn sheet_drawable_payloads(
    message_type: u32,
    source: &[u8],
) -> Result<Vec<&[u8]>, Error> {
    match message_type {
        SHEET_MESSAGE_TYPE => repeated_length_payloads(source, 2),
        FORM_BASED_SHEET_MESSAGE_TYPE => {
            repeated_length_payloads(singular_length_payload(source, 1)?, 2)
        },
        _ => Err(Error::InvalidSource {
            path: Path::Package,
        }),
    }
}

pub(super) fn canonical_varint(source: &[u8]) -> Result<u64, Error> {
    let (value, length) =
        decode_varint_from_bytes(source).map_err(|_error| Error::InvalidSource {
            path: Path::Package,
        })?;
    if length != source.len() || encoded_len(value) != length {
        return Err(Error::InvalidSource {
            path: Path::Package,
        });
    }
    Ok(value)
}

pub(in crate::package) fn require_local_reference(
    source: &[u8],
    expected: u64,
) -> Result<(), Error> {
    if local_reference_identifier(source)? != expected {
        return Err(Error::InvalidSource {
            path: Path::Package,
        });
    }
    Ok(())
}

pub(in crate::package) fn local_reference_identifier(source: &[u8]) -> Result<u64, Error> {
    let view = WireView::parse(source).map_err(map_wire_error)?;
    let mut identifier = None;
    let mut deprecated_type = None;
    let mut external = None;
    for field in view.fields() {
        field.validate_canonical_framing().map_err(map_wire_error)?;
        match field.number() {
            1 if identifier.is_none() && field.wire_type() == 0 => {
                identifier = Some(canonical_varint(field.payload())?);
            },
            2 if deprecated_type.is_none() && field.wire_type() == 0 => {
                let value = canonical_varint(field.payload())?;
                if value > u64::from(i32::MAX.unsigned_abs()) && value < MIN_SIGN_EXTENDED_I32 {
                    return Err(Error::InvalidSource {
                        path: Path::Package,
                    });
                }
                deprecated_type = Some(value);
            },
            3 if external.is_none() && field.wire_type() == 0 => {
                let value = canonical_varint(field.payload())?;
                if value > 1 {
                    return Err(Error::InvalidSource {
                        path: Path::Package,
                    });
                }
                external = Some(value != 0);
            },
            1..=3 => {
                return Err(Error::InvalidSource {
                    path: Path::Package,
                });
            },
            _ => {},
        }
    }
    let resolved_identifier = identifier.ok_or(Error::InvalidSource {
        path: Path::Package,
    })?;
    if resolved_identifier == 0 || external == Some(true) {
        return Err(Error::InvalidSource {
            path: Path::Package,
        });
    }
    Ok(resolved_identifier)
}

fn validate_stored(settings: Settings, rows: u32, columns: u32) -> Result<(), Error> {
    if u64::try_from(settings.header_row_count())
        .unwrap_or(u64::MAX)
        .saturating_add(u64::try_from(settings.footer_row_count()).unwrap_or(u64::MAX))
        > u64::from(rows)
        || u64::try_from(settings.header_column_count()).unwrap_or(u64::MAX) > u64::from(columns)
    {
        return Err(Error::InvalidSource {
            path: Path::Package,
        });
    }
    Ok(())
}
