//! Plain-text Drop Caps of a text box or shape text (record 0372; the
//! Keynote leaf of the host's `*_slide_text_box_paragraph_drop_cap*`
//! families over the shared W1-D engine, `litchi_iwa_text_wire::drop_cap`).
//!
//! A paragraph's Drop Cap is a variation of the stylesheet's Drop Cap base
//! (the theme preset that is a root style of the stylesheet) named by the
//! storage's Drop Cap table. The engine plans a rewrite of an exclusive
//! variation, a derived variation, or detaching the paragraph (removing the
//! variation when exclusive); this module publishes the plan.

use std::collections::BTreeMap;

use litchi_iwa_core::{ArchiveObject, RawMessage};
use litchi_iwa_text::paragraph::drop_cap::DropCap;
use litchi_iwa_text::position::TextPosition;
use litchi_iwa_text_wire::drop_cap::{self as engine, Change, Plan};

use super::super::structure::wire::Fields;
use super::super::structure::workspace::Workspace;
use super::text_lists::{
    component_of, register_private_style, replace_stylesheet, unregister_if_unused,
    unregister_private_style,
};
use super::text_properties::storage_payload;
use super::{Error, Refusal, Unsupported, lift};

const STORAGE_MESSAGE_TYPE: u32 = 2_001;
const STYLESHEET_MESSAGE_TYPE: u32 = 401;
const THEME_DROP_CAP_PRESETS: u32 = 8;
const STYLE_ARCHIVE: u32 = 1;
const STYLE_IS_VARIATION: u32 = 4;
const STYLE_STYLESHEET: u32 = 5;
const STYLE_PARENT: u32 = 3;

fn refused() -> Error {
    Error::Refused(Refusal::Unsupported {
        reason: Unsupported::TextStyling,
    })
}

fn from_engine<E>(_error: E) -> Error {
    refused()
}

/// Every source object holding `message_type`, staged.
fn messages(workspace: &Workspace<'_>, message_type: u32) -> Vec<(u64, Vec<u8>)> {
    let mut found = Vec::new();
    for component in workspace.package().state.source.components().iter() {
        for source in &component.archive().objects {
            let Some(identifier) = source.archive_info.identifier else {
                continue;
            };
            let Some(object) = workspace.object(identifier) else {
                continue;
            };
            for message in &object.messages {
                if message.type_ == message_type {
                    found.push((identifier, message.data.clone()));
                }
            }
        }
    }
    found
}

struct Styles {
    payloads: BTreeMap<u64, Vec<u8>>,
    exclusive: BTreeMap<u64, bool>,
}

fn styles(workspace: &Workspace<'_>) -> Result<Styles, Error> {
    let limits = workspace.wire;
    let payloads = messages(workspace, engine::MESSAGE_TYPE)
        .into_iter()
        .collect::<BTreeMap<_, _>>();
    let parents = payloads
        .values()
        .filter_map(|payload| {
            litchi_iwa_text_wire::character_style_parent(payload, limits)
                .ok()
                .flatten()
        })
        .collect::<Vec<_>>();
    let mut references = BTreeMap::<u64, usize>::new();
    for (_, payload) in messages(workspace, STORAGE_MESSAGE_TYPE) {
        if let Ok(placements) = engine::placements(&payload, limits) {
            for (_, style) in placements.entries() {
                if let Some(style) = style {
                    *references.entry(*style).or_default() += 1;
                }
            }
        }
    }
    let exclusive = payloads
        .keys()
        .map(|style| {
            (
                *style,
                references.get(style) == Some(&1) && !parents.contains(style),
            )
        })
        .collect();
    Ok(Styles {
        payloads,
        exclusive,
    })
}

/// Every Drop Cap preset the package's themes list.
fn presets(workspace: &Workspace<'_>) -> Result<Vec<u64>, Error> {
    let wire = workspace.wire;
    let mut presets = Vec::new();
    for message_type in litchi_iwa_text_wire::paragraph_style::named::THEME_MESSAGE_TYPES {
        for (_, payload) in messages(workspace, message_type) {
            let length = payload.len().max(1);
            let options = litchi_iwa_protos::theme_codec::DecodeOptions::new(
                length,
                length.saturating_mul(4),
                length.saturating_mul(16),
                64,
            );
            let Ok(theme) = litchi_iwa_protos::theme_codec::decode_theme_archive(&payload, options)
            else {
                continue;
            };
            let Some(text) = theme.text_presets() else {
                continue;
            };
            let fields = Fields::parse(text, wire).map_err(lift)?;
            for field in fields.all(THEME_DROP_CAP_PRESETS) {
                let identifier = super::super::structure::wire::reference_identifier(
                    field.message().map_err(lift)?,
                    wire,
                )
                .map_err(lift)?;
                if identifier != 0 {
                    presets.push(identifier);
                }
            }
        }
    }
    Ok(presets)
}

/// The stylesheet's Drop Cap base (the host's `base_style`).
fn base(workspace: &Workspace<'_>, styles: &Styles, stylesheet: u64) -> Result<u64, Error> {
    let wire = workspace.wire;
    let presets = presets(workspace)?;
    let mut found = Vec::new();
    for (identifier, payload) in &styles.payloads {
        if !presets.contains(identifier) {
            continue;
        }
        let fields = Fields::parse(payload, wire).map_err(lift)?;
        let Some(style) = fields.unique(STYLE_ARCHIVE).map_err(lift)? else {
            continue;
        };
        let style = Fields::parse(style.message().map_err(lift)?, wire).map_err(lift)?;
        let reference = |number| -> Result<Option<u64>, Error> {
            style
                .unique(number)
                .map_err(lift)?
                .map(|field| {
                    super::super::structure::wire::reference_identifier(
                        field.message().map_err(lift)?,
                        wire,
                    )
                    .map_err(lift)
                })
                .transpose()
        };
        if reference(STYLE_PARENT)?.is_none()
            && style.unique(STYLE_IS_VARIATION).map_err(lift)?.is_none()
            && reference(STYLE_STYLESHEET)? == Some(stylesheet)
        {
            found.push(*identifier);
        }
    }
    match found.as_slice() {
        [base] => Ok(*base),
        _ => Err(refused()),
    }
}

/// The storage's Drop Caps: (paragraph, value).
pub(super) fn drop_caps(
    workspace: &Workspace<'_>,
    storage: u64,
) -> Result<Vec<(TextPosition, DropCap)>, Error> {
    let limits = workspace.wire;
    let (_, payload) = storage_payload(workspace, storage)?;
    let placements = engine::placements(&payload, limits).map_err(from_engine)?;
    if placements
        .entries()
        .iter()
        .all(|(_, style)| style.is_none())
    {
        return Ok(Vec::new());
    }
    let styles = styles(workspace)?;
    let base = base(workspace, &styles, placements.stylesheet())?;
    let mut found = Vec::new();
    for (paragraph, style) in placements.entries() {
        let Some(style) = style.filter(|style| *style != base) else {
            continue;
        };
        let payload = styles.payloads.get(&style).ok_or(Error::InvalidSource)?;
        found.push((
            *paragraph,
            engine::model(payload, limits)
                .map_err(from_engine)?
                .drop_cap(),
        ));
    }
    Ok(found)
}

/// Replace the storage payload, keeping the Drop Cap table's field
/// references in step (the host's table-reference bookkeeping).
fn replace_storage(
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
    if let Some(old) = removed {
        if info.object_references.iter().filter(|r| **r == old).count() != 1 {
            return Err(refused());
        }
        info.object_references.retain(|reference| *reference != old);
        for field in &mut info.field_infos {
            if field.path.as_slice().first() == Some(&engine::TABLE_FIELD) {
                field
                    .object_references
                    .retain(|reference| *reference != old);
            }
        }
    }
    if let Some(new) = added {
        if info.object_references.contains(&new) {
            return Err(refused());
        }
        info.object_references.push(new);
        for field in &mut info.field_infos {
            if field.path.as_slice().first() == Some(&engine::TABLE_FIELD)
                && !field.object_references.contains(&new)
            {
                field.object_references.push(new);
            }
        }
    }
    Ok(())
}

/// Publish one Drop Cap change of the paragraph at `paragraph`.
pub(super) fn apply(
    workspace: &mut Workspace<'_>,
    storage: u64,
    paragraph: TextPosition,
    change: Change,
) -> Result<(), Error> {
    let limits = workspace.wire;
    let (index, payload) = storage_payload(workspace, storage)?;
    let placements = engine::placements(&payload, limits).map_err(from_engine)?;
    let stylesheet = placements.stylesheet();
    let styles = styles(workspace)?;
    let base = base(workspace, &styles, stylesheet)?;
    let plan = engine::plan(
        &payload,
        paragraph,
        base,
        |style| styles.payloads.get(&style).map(Vec::as_slice),
        |style| styles.exclusive.get(&style).copied().unwrap_or(false),
        change,
        limits,
    )
    .map_err(from_engine)?;
    let owner = component_of(workspace, storage)?;
    match plan {
        Plan::Unchanged => Ok(()),
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
            Ok(())
        },
        Plan::Derive { previous, payload } => {
            let base_member = workspace.component_of(base).ok_or(Error::InvalidSource)?;
            let base_component = component_of(workspace, base)?;
            let created = workspace.allocate().map_err(lift)?;
            let mut object = ArchiveObject::new(
                created,
                vec![RawMessage {
                    type_: engine::MESSAGE_TYPE,
                    data: payload,
                }],
                workspace.archive_limits(),
            )
            .map_err(|_| Error::InvalidSource)?;
            {
                let info = object
                    .archive_info
                    .message_infos
                    .first_mut()
                    .ok_or(Error::InvalidSource)?;
                info.versions = engine::MESSAGE_VERSION.to_vec().into();
                info.object_references = vec![base];
            }
            let sheet = workspace
                .payload(stylesheet, STYLESHEET_MESSAGE_TYPE)
                .map_err(lift)?
                .to_vec();
            let data = litchi_iwa_text_wire::paragraph_style::stylesheet_with_variation(
                &sheet, base, created, limits,
            )
            .map_err(from_engine)?;
            replace_stylesheet(workspace, stylesheet, data, None, Some(created))?;
            workspace.insert(&base_member, object).map_err(lift)?;
            let (_, current) = storage_payload(workspace, storage)?;
            let rewrite = engine::point(&current, paragraph, previous, Some(created), limits)
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
            register_private_style(workspace, owner, base_component, created)?;
            if let Some(previous) = previous {
                let component = component_of(workspace, previous)?;
                unregister_if_unused(workspace, owner, component, previous)?;
            }
            Ok(())
        },
        Plan::Detach {
            style,
            parent,
            exclusive,
        } => {
            let component = component_of(workspace, style)?;
            let rewrite = engine::point(&payload, paragraph, Some(style), None, limits)
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
            if exclusive {
                let sheet = workspace
                    .payload(stylesheet, STYLESHEET_MESSAGE_TYPE)
                    .map_err(lift)?
                    .to_vec();
                let data = litchi_iwa_text_wire::paragraph_style::stylesheet_without_variation(
                    &sheet, parent, style, limits,
                )
                .map_err(from_engine)?;
                replace_stylesheet(workspace, stylesheet, data, Some(style), None)?;
                workspace.remove_object(style).map_err(lift)?;
                unregister_private_style(workspace, owner, component, style, None)?;
                workspace.release_removed_identifier_suffix();
            } else {
                unregister_if_unused(workspace, owner, component, style)?;
            }
            Ok(())
        },
        _ => Err(refused()),
    }
}
