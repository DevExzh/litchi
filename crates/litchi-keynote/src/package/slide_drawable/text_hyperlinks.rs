//! Hyperlinks of a text box or shape text (record 0372; the Keynote leaf of
//! the host's `*_slide_text_box_hyperlink*` families over the shared W1-D
//! ranged smart-field engine, `litchi_iwa_text_wire::{ranged,
//! smart_field}`).
//!
//! A hyperlink is a `TSWP.HyperlinkFieldArchive` object in the storage's
//! component, covering one range of the storage's smart-field table. It is
//! addressed by its position among the storage's hyperlinks in text order.

use litchi_iwa_core::{ArchiveObject, RawMessage};
use litchi_iwa_text::hyperlink::TextHyperlinkTarget;
use litchi_iwa_text::position::TextRange;
use litchi_iwa_text_wire::ranged::{self, Table};
use litchi_iwa_text_wire::smart_field::{self, TextAttributeUuid};

use super::super::structure::workspace::Workspace;
use super::text_lists::{component_of, remove_edges_to};
use super::text_properties::storage_payload;
use super::{Error, Refusal, Unsupported, lift};

const STORAGE_MESSAGE_TYPE: u32 = 2_001;

/// One staged hyperlink change of one storage.
#[derive(Clone)]
pub(super) enum HyperlinkChange {
    Add(TextRange, TextHyperlinkTarget),
    Update(usize, TextRange, TextHyperlinkTarget),
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

/// The storage's hyperlinks in text order: (object, range, target).
pub(super) fn hyperlinks(
    workspace: &Workspace<'_>,
    storage: u64,
) -> Result<Vec<(u64, TextRange, TextHyperlinkTarget)>, Error> {
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
                .find(|message| message.type_ == smart_field::HYPERLINK_MESSAGE_TYPE)
        }) else {
            continue;
        };
        let (_, target) = smart_field::hyperlink(&message.data, limits).map_err(from_engine)?;
        found.push((object, entry.range(), target));
    }
    Ok(found)
}

pub(super) fn replace_storage(
    workspace: &mut Workspace<'_>,
    storage: u64,
    index: usize,
    payload: Vec<u8>,
    added: Option<u64>,
    removed: Option<u64>,
) -> Result<(), Error> {
    let limits = workspace.archive_limits();
    let object = workspace.object_mut(storage).map_err(lift)?;
    object
        .replace_message(
            index,
            RawMessage {
                type_: STORAGE_MESSAGE_TYPE,
                data: payload,
            },
            limits,
        )
        .map_err(|_| Error::InvalidSource)?;
    let info = object
        .archive_info
        .message_infos
        .get_mut(index)
        .ok_or(Error::InvalidSource)?;
    if let Some(added) = added {
        if info.object_references.contains(&added) {
            return Err(Error::InvalidSource);
        }
        info.object_references.push(added);
    }
    if let Some(removed) = removed {
        info.object_references
            .retain(|reference| *reference != removed);
        for field in &mut info.field_infos {
            field
                .object_references
                .retain(|reference| *reference != removed);
        }
    }
    Ok(())
}

/// A deterministic text-attribute UUID for a new field (the host draws a
/// random one).
pub(super) fn attribute_uuid(
    workspace: &mut Workspace<'_>,
    object: u64,
) -> Result<TextAttributeUuid, Error> {
    let uuid = workspace.clone_uuid(None, object).map_err(lift)?;
    let (high, low) = (uuid.upper, uuid.lower);
    let text = format!(
        "{:08x}-{:04x}-{:04x}-{:04x}-{:012x}",
        high >> 32,
        (high >> 16) & 0xffff,
        high & 0xffff,
        low >> 48,
        low & 0xffff_ffff_ffff
    );
    TextAttributeUuid::new(&text).map_err(|_| Error::InvalidSource)
}

/// Publish one hyperlink change of `storage`.
pub(super) fn apply(
    workspace: &mut Workspace<'_>,
    storage: u64,
    change: &HyperlinkChange,
) -> Result<(), Error> {
    let limits = workspace.wire;
    let (index, payload) = storage_payload(workspace, storage)?;
    match change {
        HyperlinkChange::Add(range, target) => {
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
                smart_field::encode_hyperlink(&uuid, target, limits).map_err(from_engine)?;
            let new = ArchiveObject::new(
                object,
                vec![RawMessage {
                    type_: smart_field::HYPERLINK_MESSAGE_TYPE,
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
        HyperlinkChange::Update(position, range, target) => {
            let (object, current) = hyperlinks(workspace, storage)?
                .get(*position)
                .map(|(object, range, _)| (*object, *range))
                .ok_or(refused())?;
            if current != *range {
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
            }
            let archive_limits = workspace.archive_limits();
            let field = workspace.object_mut(object).map_err(lift)?;
            let message = field
                .messages
                .iter()
                .position(|message| message.type_ == smart_field::HYPERLINK_MESSAGE_TYPE)
                .ok_or(Error::InvalidSource)?;
            let data =
                smart_field::hyperlink_with_target(&field.messages[message].data, target, limits)
                    .map_err(from_engine)?;
            field
                .replace_message(
                    message,
                    RawMessage {
                        type_: smart_field::HYPERLINK_MESSAGE_TYPE,
                        data,
                    },
                    archive_limits,
                )
                .map_err(|_| Error::InvalidSource)?;
            Ok(())
        },
        HyperlinkChange::Remove(position) => {
            let object = hyperlinks(workspace, storage)?
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
