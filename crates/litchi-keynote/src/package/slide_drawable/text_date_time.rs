//! Date & Time fields of a text box or shape text (record 0372; the Keynote
//! leaf of the host's `*_slide_text_box_date_time_field*` families over the
//! shared W1-D smart-field engine). A field is a
//! `TSWP.DateTimeSmartFieldArchive` covering a range of the storage's
//! smart-field table, addressed by its position among the storage's
//! date-time fields in text order.

use litchi_iwa_core::{ArchiveObject, RawMessage};
use litchi_iwa_text::date_time::Settings;
use litchi_iwa_text::position::TextRange;
use litchi_iwa_text_wire::ranged::{self, Table};
use litchi_iwa_text_wire::smart_field;

use super::super::structure::workspace::Workspace;
use super::text_hyperlinks::{attribute_uuid, replace_storage};
use super::text_lists::{component_of, remove_edges_to};
use super::text_properties::storage_payload;
use super::{Error, Refusal, Unsupported, lift};

/// One staged date-time field change of one storage.
#[derive(Clone)]
pub(super) enum DateTimeChange {
    Add(TextRange, Settings),
    Update(usize, TextRange, Settings),
    Remove(usize),
}

fn refused() -> Error {
    Error::Refused(Refusal::Unsupported {
        reason: Unsupported::TextStyling,
    })
}

fn from_engine<E>(_error: E) -> Error {
    refused()
}

/// The storage's date-time fields in text order: (object, range, settings).
pub(super) fn fields(
    workspace: &Workspace<'_>,
    storage: u64,
) -> Result<Vec<(u64, TextRange, Settings)>, Error> {
    let limits = workspace.wire;
    let (_, payload) = storage_payload(workspace, storage)?;
    let member = workspace
        .component_of(storage)
        .ok_or(Error::InvalidSource)?;
    let mut found = Vec::new();
    for entry in ranged::entries(&payload, Table::SmartField, limits).map_err(from_engine)? {
        let object = entry.object();
        if workspace.component_of(object).as_deref() != Some(member.as_str()) {
            continue;
        }
        let Some(message) = workspace.object(object).and_then(|object| {
            object
                .messages
                .iter()
                .find(|message| message.type_ == smart_field::DATE_TIME_MESSAGE_TYPE)
        }) else {
            continue;
        };
        let (_, settings) = smart_field::date_time(&message.data, limits).map_err(from_engine)?;
        found.push((object, entry.range(), settings));
    }
    Ok(found)
}

/// Publish one date-time field change of `storage`.
pub(super) fn apply(
    workspace: &mut Workspace<'_>,
    storage: u64,
    change: &DateTimeChange,
) -> Result<(), Error> {
    let limits = workspace.wire;
    let (index, payload) = storage_payload(workspace, storage)?;
    match change {
        DateTimeChange::Add(range, settings) => {
            let object = workspace.allocate().map_err(lift)?;
            let rewrite = ranged::rewrite(
                &payload,
                Table::SmartField,
                &ranged::Change::Add {
                    object,
                    range: *range,
                },
                limits,
            )
            .map_err(from_engine)?;
            let (added, removed) = (rewrite.added(), rewrite.removed());
            replace_storage(
                workspace,
                storage,
                index,
                rewrite.into_payload(),
                added,
                removed,
            )?;
            let uuid = attribute_uuid(workspace, object)?;
            let field =
                smart_field::encode_date_time(&uuid, settings, limits).map_err(from_engine)?;
            let new = ArchiveObject::new(
                object,
                vec![RawMessage {
                    type_: smart_field::DATE_TIME_MESSAGE_TYPE,
                    data: field,
                }],
                workspace.archive_limits(),
            )
            .map_err(|_| Error::InvalidSource)?;
            let member = workspace
                .component_of(storage)
                .ok_or(Error::InvalidSource)?;
            workspace.insert(&member, new).map_err(lift)
        },
        DateTimeChange::Update(position, range, settings) => {
            let object = fields(workspace, storage)?
                .get(*position)
                .map(|entry| entry.0)
                .ok_or(refused())?;
            let rewrite = ranged::rewrite(
                &payload,
                Table::SmartField,
                &ranged::Change::Move {
                    object,
                    range: *range,
                },
                limits,
            )
            .map_err(from_engine)?;
            replace_storage(
                workspace,
                storage,
                index,
                rewrite.into_payload(),
                None,
                None,
            )?;
            let archive_limits = workspace.archive_limits();
            let field = workspace.object_mut(object).map_err(lift)?;
            let message = field
                .messages
                .iter()
                .position(|message| message.type_ == smart_field::DATE_TIME_MESSAGE_TYPE)
                .ok_or(Error::InvalidSource)?;
            let data = smart_field::date_time_with_settings(
                &field.messages[message].data,
                settings,
                limits,
            )
            .map_err(from_engine)?;
            field
                .replace_message(
                    message,
                    RawMessage {
                        type_: smart_field::DATE_TIME_MESSAGE_TYPE,
                        data,
                    },
                    archive_limits,
                )
                .map_err(|_| Error::InvalidSource)?;
            Ok(())
        },
        DateTimeChange::Remove(position) => {
            let object = fields(workspace, storage)?
                .get(*position)
                .map(|entry| entry.0)
                .ok_or(refused())?;
            let rewrite = ranged::rewrite(
                &payload,
                Table::SmartField,
                &ranged::Change::Remove { object },
                limits,
            )
            .map_err(from_engine)?;
            let removed = rewrite.removed();
            replace_storage(
                workspace,
                storage,
                index,
                rewrite.into_payload(),
                None,
                removed,
            )?;
            let component = component_of(workspace, storage)?;
            workspace.remove_object(object).map_err(lift)?;
            remove_edges_to(workspace, component, object, None)?;
            let registered = workspace
                .registry
                .components()
                .find(|candidate| candidate.uuid_of(object).is_some())
                .map(|candidate| candidate.identifier);
            if let Some(registered) = registered {
                workspace
                    .registry
                    .component_mut(registered)
                    .ok_or(Error::InvalidSource)?
                    .remove_object_uuids(&[object]);
            }
            workspace.release_removed_identifier_suffix();
            Ok(())
        },
    }
}
