//! Paragraph lists of a text box or shape text (record 0372; the Keynote
//! leaf of the host's `set_slide_text_box_paragraph_list*` families over
//! the shared W1-D list engine, `litchi_iwa_text_wire::list`).
//!
//! The engine plans every change over the storage's list-style table; this
//! module publishes the plans on the structural workspace exactly as the
//! host does: list presets of the stylesheet are reused or created,
//! per-paragraph variations are rewritten in place, derived beside their
//! stylesheet or collapsed back into their parent, and the storage's
//! list-style references follow the table in place.

use litchi_iwa_common::WireLimits;
use litchi_iwa_core::{ArchiveObject, RawMessage};
use litchi_iwa_text::paragraph::list::{
    ParagraphList, ParagraphListLevel, ParagraphListLevelPlacement, ParagraphListNumbering,
};
use litchi_iwa_text::position::TextPosition;
use litchi_iwa_text_wire::list::{self as engine, Change, Plan, Property, PropertyKind, Styles};

use super::super::structure::metadata::ExternalReference;
use super::super::structure::wire::reference;
use super::super::structure::workspace::Workspace;
use super::{Error, Refusal, Unsupported, lift};

const STORAGE_MESSAGE_TYPE: u32 = 2_001;
const STYLESHEET_MESSAGE_TYPE: u32 = 401;
const STYLESHEET_STYLES: u32 = 1;

/// One staged list change of one storage.
#[derive(Clone)]
pub(super) enum ListChange {
    List(ParagraphList),
    Level(TextPosition, ParagraphListLevel),
    Numbering(TextPosition, ParagraphListNumbering),
    Property(TextPosition, Change),
}

struct WorkspaceStyles<'w, 'p>(&'w Workspace<'p>);

impl Styles for WorkspaceStyles<'_, '_> {
    fn list_style(&self, identifier: u64) -> Option<&[u8]> {
        self.0
            .object(identifier)?
            .messages
            .iter()
            .find(|message| message.type_ == engine::MESSAGE_TYPE)
            .map(|message| message.data.as_slice())
    }
}

fn refused() -> Error {
    Error::Refused(Refusal::Unsupported {
        reason: Unsupported::TextStyling,
    })
}

fn from_engine(error: engine::Error) -> Error {
    match error {
        engine::Error::Refused(_) | engine::Error::InvalidSource { .. } => refused(),
        _ => Error::InvalidSource,
    }
}

fn storage_payload(workspace: &Workspace<'_>, storage: u64) -> Result<(usize, Vec<u8>), Error> {
    let object = workspace.require(storage).map_err(lift)?;
    object
        .messages
        .iter()
        .enumerate()
        .find(|(_, message)| message.type_ == STORAGE_MESSAGE_TYPE)
        .map(|(index, message)| (index, message.data.clone()))
        .ok_or(Error::InvalidSource)
}

/// Whether `style` is referenced by one storage boundary only and is
/// nobody's parent.
fn exclusive(workspace: &Workspace<'_>, style: u64, limits: WireLimits) -> bool {
    let mut references = 0usize;
    for component in workspace.package().state.source.components().iter() {
        for source in &component.archive().objects {
            let Some(object) = source
                .archive_info
                .identifier
                .and_then(|identifier| workspace.object(identifier))
            else {
                continue;
            };
            for message in &object.messages {
                if message.type_ == STORAGE_MESSAGE_TYPE {
                    references += engine::boundaries(&message.data, limits)
                        .map(|entries| entries.iter().filter(|entry| entry.1 == style).count())
                        .unwrap_or(0);
                } else if message.type_ == engine::MESSAGE_TYPE
                    && litchi_iwa_text_wire::character_style_parent(&message.data, limits)
                        .ok()
                        .flatten()
                        == Some(style)
                {
                    return false;
                }
            }
        }
    }
    references == 1
}

pub(super) fn component_of(workspace: &Workspace<'_>, object: u64) -> Result<u64, Error> {
    let member = workspace.component_of(object).ok_or(Error::InvalidSource)?;
    Ok(workspace
        .registry
        .component_for_member(&member)
        .ok_or(Error::InvalidSource)?
        .identifier)
}

fn sequence(references: &mut Vec<u64>, old: &[u64], new: &[u64]) -> Result<(), Error> {
    let positions = references
        .iter()
        .enumerate()
        .filter_map(|(index, reference)| old.contains(reference).then_some(index))
        .collect::<Vec<_>>();
    if positions.len() != old.len() {
        return Err(refused());
    }
    let insertion = positions.first().copied().unwrap_or(references.len());
    references.retain(|reference| !old.contains(reference));
    references.splice(insertion..insertion, new.iter().copied());
    Ok(())
}

/// Replace the storage payload; its list-style references follow the table
/// in place (the host's `replace_reference_sequence`).
fn replace_storage(
    workspace: &mut Workspace<'_>,
    storage: u64,
    index: usize,
    payload: Vec<u8>,
    old: &[u64],
    new: &[u64],
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
    if !old.is_empty() || !new.is_empty() {
        sequence(&mut info.object_references, old, new)?;
        for field in &mut info.field_infos {
            if field
                .object_references
                .iter()
                .any(|reference| old.contains(reference))
            {
                sequence(&mut field.object_references, old, new)?;
            }
        }
    }
    Ok(())
}

pub(super) fn replace_stylesheet(
    workspace: &mut Workspace<'_>,
    stylesheet: u64,
    data: Vec<u8>,
    removed: Option<u64>,
    appended: Option<u64>,
) -> Result<(), Error> {
    let limits = workspace.archive_limits();
    let object = workspace.object_mut(stylesheet).map_err(lift)?;
    let index = object
        .messages
        .iter()
        .position(|message| message.type_ == STYLESHEET_MESSAGE_TYPE)
        .ok_or(Error::InvalidSource)?;
    object
        .replace_message(
            index,
            RawMessage {
                type_: STYLESHEET_MESSAGE_TYPE,
                data,
            },
            limits,
        )
        .map_err(|_| Error::InvalidSource)?;
    let info = object
        .archive_info
        .message_infos
        .get_mut(index)
        .ok_or(Error::InvalidSource)?;
    if let Some(removed) = removed {
        info.object_references
            .retain(|reference| *reference != removed);
        for field in &mut info.field_infos {
            field
                .object_references
                .retain(|reference| *reference != removed);
        }
    }
    if let Some(appended) = appended {
        if info.object_references.contains(&appended) {
            return Err(Error::InvalidSource);
        }
        info.object_references.push(appended);
    }
    Ok(())
}

fn new_style(
    workspace: &Workspace<'_>,
    identifier: u64,
    payload: Vec<u8>,
    references: &[u64],
) -> Result<ArchiveObject, Error> {
    let mut object = ArchiveObject::new(
        identifier,
        vec![RawMessage {
            type_: engine::MESSAGE_TYPE,
            data: payload,
        }],
        workspace.archive_limits(),
    )
    .map_err(|_| Error::InvalidSource)?;
    let info = object
        .archive_info
        .message_infos
        .first_mut()
        .ok_or(Error::InvalidSource)?;
    info.versions = engine::MESSAGE_VERSION.to_vec().into();
    info.object_references = references.to_vec();
    Ok(object)
}

pub(super) fn add_edge(
    workspace: &mut Workspace<'_>,
    from: u64,
    to: u64,
    object: u64,
) -> Result<(), Error> {
    if from == to {
        return Ok(());
    }
    workspace
        .registry
        .component_mut(from)
        .ok_or(Error::InvalidSource)?
        .add_external_reference(ExternalReference {
            component: to,
            object: Some(object),
            is_weak: None,
        });
    Ok(())
}

/// The host's `register_private_style`: a UUID in the style's component
/// and an owner edge to it.
pub(super) fn register_private_style(
    workspace: &mut Workspace<'_>,
    owner: u64,
    style_component: u64,
    style: u64,
) -> Result<(), Error> {
    let uuid = workspace.clone_uuid(None, style).map_err(lift)?;
    workspace
        .registry
        .component_mut(style_component)
        .ok_or(Error::InvalidSource)?
        .add_object_uuid(style, uuid);
    add_edge(workspace, owner, style_component, style)
}

/// The host's `unregister_private_style`.
pub(super) fn unregister_private_style(
    workspace: &mut Workspace<'_>,
    owner: u64,
    style_component: u64,
    style: u64,
    replacement: Option<u64>,
) -> Result<(), Error> {
    workspace
        .registry
        .component_mut(style_component)
        .ok_or(Error::InvalidSource)?
        .remove_object_uuids(&[style]);
    remove_edges_to(workspace, style_component, style, None)?;
    if let Some(replacement) = replacement {
        add_edge(workspace, owner, style_component, replacement)?;
    }
    Ok(())
}

/// Remove the external references to (`component`, `object`), from every
/// component or only from `from`.
pub(super) fn remove_edges_to(
    workspace: &mut Workspace<'_>,
    component: u64,
    object: u64,
    from: Option<u64>,
) -> Result<(), Error> {
    let referrers = workspace
        .registry
        .components()
        .filter(|candidate| from.is_none_or(|from| candidate.identifier == from))
        .filter(|candidate| {
            candidate.external_references().iter().any(|reference| {
                reference.component == component && reference.object == Some(object)
            })
        })
        .map(|candidate| candidate.identifier)
        .collect::<Vec<_>>();
    for referrer in referrers {
        let entry = workspace
            .registry
            .component_mut(referrer)
            .ok_or(Error::InvalidSource)?;
        let kept = entry
            .external_references()
            .iter()
            .filter(|reference| {
                !(reference.component == component && reference.object == Some(object))
            })
            .copied()
            .collect();
        entry.set_external_references(kept);
    }
    Ok(())
}

/// The host's `unregister_owner_reference_if_unused`: drop the owner's
/// edge to `style` when no object of the owner still references it.
pub(super) fn unregister_if_unused(
    workspace: &mut Workspace<'_>,
    owner: u64,
    style_component: u64,
    style: u64,
) -> Result<(), Error> {
    if owner == style_component {
        return Ok(());
    }
    let member = workspace
        .registry
        .component(owner)
        .ok_or(Error::InvalidSource)?
        .member_name();
    for object in workspace.member_objects(&member) {
        let used = workspace.object(object).is_some_and(|object| {
            object
                .archive_info
                .message_infos
                .iter()
                .any(|info| info.object_references.contains(&style))
        });
        if used {
            return Ok(());
        }
    }
    remove_edges_to(workspace, style_component, style, Some(owner))
}

/// The stylesheet's preset object for `preset`, created when missing.
fn preset_object(
    workspace: &mut Workspace<'_>,
    owner: u64,
    style_member: &str,
    stylesheet: u64,
    preset: ParagraphList,
    limits: WireLimits,
) -> Result<u64, Error> {
    let candidates = workspace
        .member_objects(style_member)
        .into_iter()
        .filter_map(|identifier| {
            workspace.object(identifier).and_then(|object| {
                object
                    .messages
                    .iter()
                    .find(|message| message.type_ == engine::MESSAGE_TYPE)
                    .map(|message| (identifier, message.data.clone()))
            })
        })
        .collect::<Vec<_>>();
    let style_component = workspace
        .registry
        .component_for_member(style_member)
        .ok_or(Error::InvalidSource)?
        .identifier;
    let existing = engine::find_preset(
        candidates
            .iter()
            .map(|(identifier, payload)| (*identifier, payload.as_slice())),
        stylesheet,
        preset,
        limits,
    )
    .map_err(from_engine)?;
    if let Some(identifier) = existing {
        add_edge(workspace, owner, style_component, identifier)?;
        return Ok(identifier);
    }
    let created = workspace.allocate().map_err(lift)?;
    let object = new_style(
        workspace,
        created,
        engine::preset_payload(stylesheet, preset, limits).map_err(from_engine)?,
        &[],
    )?;
    // The host's `insert_private_style`: appended to the stylesheet's styles.
    let sheet = workspace
        .payload(stylesheet, STYLESHEET_MESSAGE_TYPE)
        .map_err(lift)?
        .to_vec();
    let data = litchi_iwa_common::wire::append_repeated_length_delimited_field(
        &sheet,
        STYLESHEET_STYLES,
        &reference(created),
    )
    .map_err(|_| Error::InvalidSource)?;
    replace_stylesheet(workspace, stylesheet, data, None, Some(created))?;
    workspace.insert(style_member, object).map_err(lift)?;
    register_private_style(workspace, owner, style_component, created)?;
    Ok(created)
}

fn set_list(
    workspace: &mut Workspace<'_>,
    storage: u64,
    preset: ParagraphList,
    limits: WireLimits,
) -> Result<(), Error> {
    let (index, payload) = storage_payload(workspace, storage)?;
    let style = engine::uniform_style(&payload, limits).map_err(from_engine)?;
    if engine::preset(&WorkspaceStyles(workspace), style, limits).ok() == Some(preset) {
        return Ok(());
    }
    let style_member = workspace.component_of(style).ok_or(Error::InvalidSource)?;
    let stylesheet =
        engine::stylesheet(&WorkspaceStyles(workspace), style, limits).map_err(from_engine)?;
    let owner = component_of(workspace, storage)?;
    let target = preset_object(workspace, owner, &style_member, stylesheet, preset, limits)?;
    let (_, payload) = storage_payload(workspace, storage)?;
    let rewritten =
        engine::replace_boundaries(&payload, &[(0, target)], limits).map_err(from_engine)?;
    replace_storage(workspace, storage, index, rewritten, &[style], &[target])?;
    let style_component = component_of(workspace, style)?;
    unregister_if_unused(workspace, owner, style_component, style)
}

fn apply_plan(
    workspace: &mut Workspace<'_>,
    storage: u64,
    paragraph: TextPosition,
    plan: Plan,
    limits: WireLimits,
) -> Result<bool, Error> {
    let (index, payload) = storage_payload(workspace, storage)?;
    let owner = component_of(workspace, storage)?;
    match plan {
        Plan::Unchanged => return Ok(false),
        Plan::Rewrite { style, payload } => {
            let archive_limits = workspace.archive_limits();
            let object = workspace.object_mut(style).map_err(lift)?;
            let message = object
                .messages
                .iter()
                .position(|message| message.type_ == engine::MESSAGE_TYPE)
                .ok_or(Error::InvalidSource)?;
            object
                .replace_message(
                    message,
                    RawMessage {
                        type_: engine::MESSAGE_TYPE,
                        data: payload,
                    },
                    archive_limits,
                )
                .map_err(|_| Error::InvalidSource)?;
        },
        Plan::Derive {
            parent,
            stylesheet,
            payload: variation,
        } => {
            let style_member = workspace.component_of(parent).ok_or(Error::InvalidSource)?;
            let style_component = component_of(workspace, parent)?;
            let created = workspace.allocate().map_err(lift)?;
            let object = new_style(workspace, created, variation, &[parent])?;
            let sheet = workspace
                .payload(stylesheet, STYLESHEET_MESSAGE_TYPE)
                .map_err(lift)?
                .to_vec();
            let data = litchi_iwa_text_wire::paragraph_style::stylesheet_with_variation(
                &sheet, parent, created, limits,
            )
            .map_err(|_| Error::InvalidSource)?;
            replace_stylesheet(workspace, stylesheet, data, None, Some(created))?;
            workspace.insert(&style_member, object).map_err(lift)?;
            register_private_style(workspace, owner, style_component, created)?;
            let old = engine::boundaries(&payload, limits)
                .map_err(from_engine)?
                .into_iter()
                .map(|entry| entry.1)
                .collect::<Vec<_>>();
            let new = engine::boundaries_with_style(&payload, paragraph, created, limits)
                .map_err(from_engine)?;
            let rewritten =
                engine::replace_boundaries(&payload, &new, limits).map_err(from_engine)?;
            let new = new.into_iter().map(|entry| entry.1).collect::<Vec<_>>();
            replace_storage(workspace, storage, index, rewritten, &old, &new)?;
        },
        Plan::Collapse {
            style,
            parent,
            stylesheet,
        } => {
            let style_component = component_of(workspace, style)?;
            let old = engine::boundaries(&payload, limits)
                .map_err(from_engine)?
                .into_iter()
                .map(|entry| entry.1)
                .collect::<Vec<_>>();
            let new = engine::boundaries_with_style(&payload, paragraph, parent, limits)
                .map_err(from_engine)?;
            let rewritten =
                engine::replace_boundaries(&payload, &new, limits).map_err(from_engine)?;
            let new = new.into_iter().map(|entry| entry.1).collect::<Vec<_>>();
            replace_storage(workspace, storage, index, rewritten, &old, &new)?;
            let sheet = workspace
                .payload(stylesheet, STYLESHEET_MESSAGE_TYPE)
                .map_err(lift)?
                .to_vec();
            let data = litchi_iwa_text_wire::paragraph_style::stylesheet_without_variation(
                &sheet, parent, style, limits,
            )
            .map_err(|_| Error::InvalidSource)?;
            replace_stylesheet(workspace, stylesheet, data, Some(style), None)?;
            workspace.remove_object(style).map_err(lift)?;
            unregister_private_style(workspace, owner, style_component, style, Some(parent))?;
            workspace.release_removed_identifier_suffix();
        },
        _ => return Err(refused()),
    }
    Ok(true)
}

/// Publish one list change of `storage`.
pub(super) fn apply(
    workspace: &mut Workspace<'_>,
    storage: u64,
    change: &ListChange,
) -> Result<(), Error> {
    let limits = workspace.wire;
    match change {
        ListChange::List(preset) => set_list(workspace, storage, *preset, limits),
        ListChange::Level(paragraph, level) => {
            let (index, payload) = storage_payload(workspace, storage)?;
            if let Some(rewritten) =
                engine::with_level(&payload, *paragraph, *level, limits).map_err(from_engine)?
            {
                replace_storage(workspace, storage, index, rewritten, &[], &[])?;
            }
            Ok(())
        },
        ListChange::Numbering(paragraph, numbering) => {
            let (index, payload) = storage_payload(workspace, storage)?;
            if let Some(rewritten) = engine::with_numbering(
                &payload,
                &WorkspaceStyles(workspace),
                *paragraph,
                *numbering,
                limits,
            )
            .map_err(from_engine)?
            {
                replace_storage(workspace, storage, index, rewritten, &[], &[])?;
            }
            Ok(())
        },
        ListChange::Property(paragraph, change) => {
            let (_, payload) = storage_payload(workspace, storage)?;
            let plan = engine::plan(
                &payload,
                &WorkspaceStyles(workspace),
                *paragraph,
                change,
                |style| exclusive(workspace, style, limits),
                limits,
            )
            .map_err(from_engine)?;
            let changed = apply_plan(workspace, storage, *paragraph, plan, limits)?;
            if let Change::Reset(kind) = change
                && changed
            {
                let (_, payload) = storage_payload(workspace, storage)?;
                let plan = engine::plan_collapse(
                    &payload,
                    &WorkspaceStyles(workspace),
                    *paragraph,
                    *kind,
                    |style| exclusive(workspace, style, limits),
                    limits,
                )
                .map_err(from_engine)?;
                apply_plan(workspace, storage, *paragraph, plan, limits)?;
            }
            Ok(())
        },
    }
}

/// The storage's uniform list preset.
pub(super) fn list(workspace: &Workspace<'_>, storage: u64) -> Result<ParagraphList, Error> {
    let limits = workspace.wire;
    let (_, payload) = storage_payload(workspace, storage)?;
    let style = engine::uniform_style(&payload, limits).map_err(from_engine)?;
    engine::preset(&WorkspaceStyles(workspace), style, limits).map_err(from_engine)
}

/// Every paragraph's list level.
pub(super) fn levels(
    workspace: &Workspace<'_>,
    storage: u64,
) -> Result<Vec<ParagraphListLevelPlacement>, Error> {
    let (_, payload) = storage_payload(workspace, storage)?;
    engine::levels(&payload, workspace.wire).map_err(from_engine)
}

/// One paragraph's list level.
pub(super) fn level(
    workspace: &Workspace<'_>,
    storage: u64,
    paragraph: TextPosition,
) -> Result<ParagraphListLevel, Error> {
    let (_, payload) = storage_payload(workspace, storage)?;
    engine::level(&payload, paragraph, workspace.wire).map_err(from_engine)
}

/// One paragraph's list numbering.
pub(super) fn numbering(
    workspace: &Workspace<'_>,
    storage: u64,
    paragraph: TextPosition,
) -> Result<ParagraphListNumbering, Error> {
    let (_, payload) = storage_payload(workspace, storage)?;
    engine::numbering(&payload, paragraph, workspace.wire).map_err(from_engine)
}

/// One paragraph's effective list property.
pub(super) fn property(
    workspace: &Workspace<'_>,
    storage: u64,
    paragraph: TextPosition,
    kind: PropertyKind,
) -> Result<Property, Error> {
    let (_, payload) = storage_payload(workspace, storage)?;
    engine::property(
        &payload,
        &WorkspaceStyles(workspace),
        paragraph,
        kind,
        workspace.wire,
    )
    .map_err(from_engine)
}
