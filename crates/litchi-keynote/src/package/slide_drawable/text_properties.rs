//! Character and paragraph properties of a whole text box or shape text
//! (record 0372; the Keynote leaf of the host's `set_slide_text_box_*` /
//! `reset_slide_text_box_*` families over the shared W1-D paragraph-style
//! engine, `litchi_iwa_text_wire::paragraph_style`).
//!
//! A drawable's storage is uniformly styled by one paragraph style. Each
//! change runs the engine's plan over that style (exclusive when the
//! storage is its only referrer and no style names it as parent) and
//! publishes it:
//!
//! - **Rewrite**: the exclusive variation takes the new payload in place;
//! - **Derive**: a new child variation is stored beside its stylesheet,
//!   registered there, and the storage points at it;
//! - **Collapse**: the variation overrides nothing any more, so the storage
//!   points back at its parent and the variation is removed with its
//!   stylesheet registration.

use litchi_iwa_common::WireLimits;
use litchi_iwa_core::{
    ArchiveObject, FieldInfo, FieldPath, FieldType, RawMessage, UnknownFieldRule,
};
use litchi_iwa_text_wire::paragraph_style::{
    self as engine, Change, MESSAGE_VERSION, Plan, Property, PropertyKind, Styles, Target,
};

use super::super::structure::metadata::ExternalReference;
use super::super::structure::workspace::Workspace;
use super::{Error, Refusal, Unsupported, lift};

const STORAGE_MESSAGE_TYPE: u32 = 2_001;
const STYLESHEET_MESSAGE_TYPE: u32 = 401;
/// The text fill, stroke, shadow and background paths a variation manages.
const MANAGED_PATHS: [[u32; 2]; 4] = [[11, 46], [11, 44], [11, 21], [11, 26]];

/// The staged package's paragraph styles, as the engine reads them.
struct WorkspaceStyles<'w, 'p>(&'w Workspace<'p>);

impl Styles for WorkspaceStyles<'_, '_> {
    fn paragraph_style(&self, identifier: u64) -> Option<&[u8]> {
        self.0
            .object(identifier)?
            .messages
            .iter()
            .find(|message| message.type_ == engine::MESSAGE_TYPE)
            .map(|message| message.data.as_slice())
    }
}

fn from_engine(error: engine::Error) -> Error {
    match error {
        engine::Error::InvalidSource { .. } => Error::Refused(Refusal::Unsupported {
            reason: Unsupported::TextStyling,
        }),
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

/// Whether `storage` is the only storage whose table names `style` and no
/// paragraph style names it as parent (the host's exclusivity rule).
fn exclusive(
    workspace: &Workspace<'_>,
    storage: u64,
    style: u64,
    limits: WireLimits,
) -> Result<bool, Error> {
    let package = workspace.package();
    for component in package.state.source.components().iter() {
        for source in &component.archive().objects {
            let Some(identifier) = source.archive_info.identifier else {
                continue;
            };
            let Some(object) = workspace.object(identifier) else {
                continue;
            };
            for (message, info) in object
                .messages
                .iter()
                .zip(&object.archive_info.message_infos)
            {
                if message.type_ == STORAGE_MESSAGE_TYPE
                    && identifier != storage
                    && info.object_references.contains(&style)
                {
                    return Ok(false);
                }
                if message.type_ == engine::MESSAGE_TYPE
                    && litchi_iwa_text_wire::character_style_parent(&message.data, limits)
                        .ok()
                        .flatten()
                        == Some(style)
                {
                    return Ok(false);
                }
            }
        }
    }
    Ok(true)
}

fn sync_managed(
    infos: &mut Vec<FieldInfo>,
    managed: [bool; 4],
    limits: litchi_iwa_core::Limits,
) -> Result<(), Error> {
    for (present, path) in managed.into_iter().zip(MANAGED_PATHS) {
        let is_path = |field: &FieldInfo| field.path.as_slice() == path;
        if present {
            if !infos.iter().any(is_path) {
                infos.push(managed_info(path, limits)?);
            }
        } else {
            infos.retain(|field| !is_path(field));
        }
    }
    Ok(())
}

fn managed_info(path: [u32; 2], limits: litchi_iwa_core::Limits) -> Result<FieldInfo, Error> {
    let mut info = FieldInfo::new(
        FieldPath::new(path.to_vec(), limits).map_err(|_| Error::InvalidSource)?,
        limits,
    )
    .map_err(|_| Error::InvalidSource)?;
    info.r#type = Some(FieldType::Message);
    info.unknown_field_rule = Some(UnknownFieldRule::IgnoreAndPreserve);
    Ok(info)
}

fn managed(variation: &engine::Variation) -> [bool; 4] {
    let fields = variation.managed_fields();
    [
        fields.text_fill(),
        fields.text_stroke(),
        fields.text_shadow(),
        fields.text_background(),
    ]
}

fn component_of(workspace: &Workspace<'_>, object: u64) -> Result<u64, Error> {
    let member = workspace.component_of(object).ok_or(Error::InvalidSource)?;
    Ok(workspace
        .registry
        .component_for_member(&member)
        .ok_or(Error::InvalidSource)?
        .identifier)
}

/// Replace the storage's payload, repointing its header reference.
fn repoint_storage(
    workspace: &mut Workspace<'_>,
    storage: u64,
    index: usize,
    payload: Vec<u8>,
    from: u64,
    to: u64,
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
    let mut replaced = 0;
    for reference in &mut info.object_references {
        if *reference == from {
            *reference = to;
            replaced += 1;
        }
    }
    for field in &mut info.field_infos {
        for reference in &mut field.object_references {
            if *reference == from {
                *reference = to;
            }
        }
    }
    if replaced != 1 {
        return Err(Error::Refused(Refusal::Unsupported {
            reason: Unsupported::TextStyling,
        }));
    }
    Ok(())
}

fn replace_stylesheet(
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
        info.object_references.push(appended);
    }
    Ok(())
}

/// The effective `kind` of the storage's uniform paragraph style.
pub(super) fn effective(
    workspace: &Workspace<'_>,
    storage: u64,
    kind: PropertyKind,
) -> Result<Property, Error> {
    let limits = workspace.wire;
    let (_, payload) = storage_payload(workspace, storage)?;
    let style = engine::uniform_style(&payload, limits).map_err(from_engine)?;
    engine::effective(&WorkspaceStyles(workspace), style, kind, limits).map_err(from_engine)
}

/// Plan and publish one change of the storage's paragraph style.
pub(super) fn apply(
    workspace: &mut Workspace<'_>,
    storage: u64,
    change: &Change,
) -> Result<(), Error> {
    let limits = workspace.wire;
    let archive_limits = workspace.archive_limits();
    let (storage_index, payload) = storage_payload(workspace, storage)?;
    let style = engine::uniform_style(&payload, limits).map_err(from_engine)?;
    let target = Target::new(style, exclusive(workspace, storage, style, limits)?);
    let plan =
        engine::plan(&WorkspaceStyles(workspace), target, change, limits).map_err(from_engine)?;
    match plan {
        Plan::Unchanged => Ok(()),
        Plan::Rewrite(variation) => {
            let flags = managed(&variation);
            let object = workspace.object_mut(style).map_err(lift)?;
            let index = object
                .messages
                .iter()
                .position(|message| message.type_ == engine::MESSAGE_TYPE)
                .ok_or(Error::InvalidSource)?;
            object
                .replace_message(
                    index,
                    RawMessage {
                        type_: engine::MESSAGE_TYPE,
                        data: variation.into_payload(),
                    },
                    archive_limits,
                )
                .map_err(|_| Error::InvalidSource)?;
            let info = object
                .archive_info
                .message_infos
                .get_mut(index)
                .ok_or(Error::InvalidSource)?;
            sync_managed(&mut info.field_infos, flags, archive_limits)
        },
        Plan::Derive(variation) => {
            let stylesheet = variation.stylesheet();
            let style_member = workspace.component_of(style).ok_or(Error::InvalidSource)?;
            if workspace.component_of(stylesheet) != Some(style_member.clone()) {
                return Err(Error::Refused(Refusal::Unsupported {
                    reason: Unsupported::TextStyling,
                }));
            }
            let created = workspace.allocate().map_err(lift)?;
            let repointed =
                engine::repoint(&payload, style, created, limits).map_err(from_engine)?;
            repoint_storage(workspace, storage, storage_index, repointed, style, created)?;
            let mut object = ArchiveObject::new(
                created,
                vec![RawMessage {
                    type_: engine::MESSAGE_TYPE,
                    data: variation.payload().to_vec(),
                }],
                archive_limits,
            )
            .map_err(|_| Error::InvalidSource)?;
            {
                let info = object
                    .archive_info
                    .message_infos
                    .first_mut()
                    .ok_or(Error::InvalidSource)?;
                info.versions = MESSAGE_VERSION.to_vec().into();
                info.object_references = variation.references().to_vec();
                sync_managed(&mut info.field_infos, managed(&variation), archive_limits)?;
            }
            let sheet = workspace
                .payload(stylesheet, STYLESHEET_MESSAGE_TYPE)
                .map_err(lift)?
                .to_vec();
            let data = engine::stylesheet_with_variation(&sheet, style, created, limits)
                .map_err(from_engine)?;
            replace_stylesheet(workspace, stylesheet, data, None, Some(created))?;
            workspace.insert(&style_member, object).map_err(lift)?;
            let style_component = component_of(workspace, stylesheet)?;
            let owner = component_of(workspace, storage)?;
            let uuid = workspace.clone_uuid(None, created).map_err(lift)?;
            workspace
                .registry
                .component_mut(style_component)
                .ok_or(Error::InvalidSource)?
                .add_object_uuid(created, uuid);
            if owner != style_component {
                workspace
                    .registry
                    .component_mut(owner)
                    .ok_or(Error::InvalidSource)?
                    .add_external_reference(ExternalReference {
                        component: style_component,
                        object: Some(created),
                        is_weak: None,
                    });
            }
            Ok(())
        },
        Plan::Collapse { parent, stylesheet } => {
            let repointed =
                engine::repoint(&payload, style, parent, limits).map_err(from_engine)?;
            repoint_storage(workspace, storage, storage_index, repointed, style, parent)?;
            let sheet = workspace
                .payload(stylesheet, STYLESHEET_MESSAGE_TYPE)
                .map_err(lift)?
                .to_vec();
            let data = engine::stylesheet_without_variation(&sheet, parent, style, limits)
                .map_err(from_engine)?;
            let style_component = component_of(workspace, stylesheet)?;
            let owner = component_of(workspace, storage)?;
            replace_stylesheet(workspace, stylesheet, data, Some(style), None)?;
            workspace.remove_object(style).map_err(lift)?;
            workspace
                .registry
                .component_mut(style_component)
                .ok_or(Error::InvalidSource)?
                .remove_object_uuids(&[style]);
            let referrers = workspace
                .registry
                .components()
                .filter(|component| {
                    component.external_references().iter().any(|reference| {
                        reference.component == style_component && reference.object == Some(style)
                    })
                })
                .map(|component| component.identifier)
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
                        !(reference.component == style_component && reference.object == Some(style))
                    })
                    .copied()
                    .collect();
                entry.set_external_references(kept);
            }
            if owner != style_component {
                workspace
                    .registry
                    .component_mut(owner)
                    .ok_or(Error::InvalidSource)?
                    .add_external_reference(ExternalReference {
                        component: style_component,
                        object: Some(parent),
                        is_weak: None,
                    });
            }
            workspace.release_removed_identifier_suffix();
            Ok(())
        },
        _ => Err(Error::Refused(Refusal::Unsupported {
            reason: Unsupported::TextStyling,
        })),
    }
}
