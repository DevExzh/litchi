//! Named paragraph styles of a text box or shape text (record 0372; the
//! Keynote leaf of the host's `*_slide_text_box_named_paragraph_style*`
//! families over the shared W1-D engine,
//! `litchi_iwa_text_wire::paragraph_style::named`).
//!
//! The catalog is the union of the paragraph presets of the themes whose
//! document stylesheet is the storage's. Apply repoints the storage (and
//! removes its exclusive override chain), create copies a named style under
//! a new name and lists it in the themes, rename rewrites the name, delete
//! unregisters an unreferenced style everywhere, and redefine folds the
//! storage's overrides into its named style before applying it.

use std::collections::BTreeMap;

use litchi_iwa_core::{ArchiveObject, RawMessage};
use litchi_iwa_text::paragraph::style::raw::native_id;
use litchi_iwa_text::paragraph::style::{
    NamedParagraphStyle, ParagraphStyleId, ParagraphStyleName,
};
use litchi_iwa_text_wire::paragraph_style::named::{self, ApplyPlan, Catalog};
use litchi_iwa_text_wire::paragraph_style::{self as engine};

use super::super::structure::wire::Fields;
use super::super::structure::workspace::Workspace;
use super::text_lists::{
    add_edge, component_of, remove_edges_to, replace_stylesheet, unregister_if_unused,
    unregister_private_style,
};
use super::text_properties::{exclusive, repoint_storage, storage_payload, sync_managed};
use super::{Error, Refusal, Unsupported, lift};

const STORAGE_MESSAGE_TYPE: u32 = 2_001;
const STYLESHEET_MESSAGE_TYPE: u32 = 401;
const STORAGE_PARAGRAPH_TABLE: u32 = 5;

/// One staged named-style change of one storage.
#[derive(Clone)]
pub(super) enum NamedChange {
    Apply(ParagraphStyleId),
    Create(ParagraphStyleId, ParagraphStyleName),
    Rename(ParagraphStyleId, ParagraphStyleName),
    Delete(ParagraphStyleId),
    DeleteApplied(ParagraphStyleId, ParagraphStyleId),
    Redefine,
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

/// Every staged object of the source package that holds `message_type`:
/// (identifier, message index, payload).
fn messages(workspace: &Workspace<'_>, message_type: u32) -> Vec<(u64, usize, Vec<u8>)> {
    let mut found = Vec::new();
    for component in workspace.package().state.source.components().iter() {
        for source in &component.archive().objects {
            let Some(identifier) = source.archive_info.identifier else {
                continue;
            };
            let Some(object) = workspace.object(identifier) else {
                continue;
            };
            for (index, message) in object.messages.iter().enumerate() {
                if message.type_ == message_type {
                    found.push((identifier, index, message.data.clone()));
                }
            }
        }
    }
    found
}

/// The paragraph styles reachable by the engine, by identifier.
fn paragraph_styles(workspace: &Workspace<'_>) -> BTreeMap<u64, Vec<u8>> {
    messages(workspace, engine::MESSAGE_TYPE)
        .into_iter()
        .map(|(identifier, _, payload)| (identifier, payload))
        .collect()
}

struct Context {
    first: u64,
    styles: BTreeMap<u64, Vec<u8>>,
    catalog: Catalog,
    /// Themes listing the stylesheet's presets: (object, message index,
    /// message type, payload, presets).
    themes: Vec<(u64, usize, u32, Vec<u8>, Vec<u64>)>,
}

impl Context {
    fn load(workspace: &Workspace<'_>, storage: u64) -> Result<Self, Error> {
        let limits = workspace.wire;
        let (_, payload) = storage_payload(workspace, storage)?;
        let first = engine::uniform_style(&payload, limits).map_err(from_engine)?;
        let styles = paragraph_styles(workspace);
        let stylesheet = named::stylesheet(&styles, first, limits).map_err(from_engine)?;
        let sheet = workspace
            .payload(stylesheet, STYLESHEET_MESSAGE_TYPE)
            .map_err(lift)?
            .to_vec();
        let mut themes = Vec::new();
        let mut presets = Vec::new();
        for message_type in named::THEME_MESSAGE_TYPES {
            for (identifier, index, payload) in messages(workspace, message_type) {
                if let Some(listed) =
                    named::theme_presets(&payload, stylesheet, limits).map_err(from_engine)?
                {
                    presets.extend_from_slice(&listed);
                    themes.push((identifier, index, message_type, payload, listed));
                }
            }
        }
        let catalog =
            named::catalog(&styles, stylesheet, &sheet, &presets, limits).map_err(from_engine)?;
        let sheet_member = workspace.component_of(stylesheet);
        for style in catalog.styles() {
            if workspace.component_of(native_id(style.id())) != sheet_member {
                return Err(refused());
            }
        }
        Ok(Self {
            first,
            styles,
            catalog,
            themes,
        })
    }

    fn themes_with(&self, style: u64) -> Result<Vec<(u64, usize, u32, Vec<u8>)>, Error> {
        let found = self
            .themes
            .iter()
            .filter(|theme| theme.4.contains(&style))
            .map(|theme| (theme.0, theme.1, theme.2, theme.3.clone()))
            .collect::<Vec<_>>();
        if found.is_empty() {
            return Err(refused());
        }
        Ok(found)
    }
}

/// The paragraph-style references of every storage's paragraph table.
fn storage_tables(workspace: &Workspace<'_>) -> Result<Vec<(u64, Vec<u64>)>, Error> {
    let wire = workspace.wire;
    let mut tables = Vec::new();
    for (identifier, _, payload) in messages(workspace, STORAGE_MESSAGE_TYPE) {
        let fields = Fields::parse(&payload, wire).map_err(lift)?;
        let mut styles = Vec::new();
        if let Some(table) = fields.unique(STORAGE_PARAGRAPH_TABLE).map_err(lift)? {
            let table = Fields::parse(table.message().map_err(lift)?, wire).map_err(lift)?;
            for entry in table.all(1) {
                let entry = Fields::parse(entry.message().map_err(lift)?, wire).map_err(lift)?;
                if let Some(reference) = entry.unique(2).map_err(lift)? {
                    styles.push(
                        super::super::structure::wire::reference_identifier(
                            reference.message().map_err(lift)?,
                            wire,
                        )
                        .map_err(lift)?,
                    );
                }
            }
        }
        tables.push((identifier, styles));
    }
    Ok(tables)
}

fn replace_message(
    workspace: &mut Workspace<'_>,
    object: u64,
    index: usize,
    message_type: u32,
    payload: Vec<u8>,
    edit: impl FnOnce(&mut litchi_iwa_core::MessageInfo) -> Result<(), Error>,
) -> Result<(), Error> {
    let limits = workspace.archive_limits();
    let target = workspace.object_mut(object).map_err(lift)?;
    target
        .replace_message(
            index,
            RawMessage {
                type_: message_type,
                data: payload,
            },
            limits,
        )
        .map_err(|_| Error::InvalidSource)?;
    edit(
        target
            .archive_info
            .message_infos
            .get_mut(index)
            .ok_or(Error::InvalidSource)?,
    )
}

fn drop_reference(info: &mut litchi_iwa_core::MessageInfo, target: u64) {
    info.object_references
        .retain(|reference| *reference != target);
    for field in &mut info.field_infos {
        field
            .object_references
            .retain(|reference| *reference != target);
    }
}

fn apply(
    workspace: &mut Workspace<'_>,
    storage: u64,
    target: ParagraphStyleId,
) -> Result<(), Error> {
    let limits = workspace.wire;
    let context = Context::load(workspace, storage)?;
    let tables = storage_tables(workspace)?;
    let mut parents = Vec::new();
    for (style, payload) in &context.styles {
        if let Some(parent) = litchi_iwa_text_wire::character_style_parent(payload, limits)
            .ok()
            .flatten()
        {
            parents.push((*style, parent));
        }
    }
    let is_exclusive = exclusive(workspace, storage, context.first, limits)?;
    let plan = named::plan_apply(
        &context.styles,
        &context.catalog,
        context.first,
        target,
        is_exclusive,
        |style, removed| {
            !tables
                .iter()
                .any(|(owner, styles)| *owner != storage && styles.contains(&style))
                && !parents
                    .iter()
                    .any(|(child, parent)| *parent == style && !removed.contains(child))
        },
        limits,
    )
    .map_err(from_engine)?;
    let (index, payload) = storage_payload(workspace, storage)?;
    let owner = component_of(workspace, storage)?;
    match plan {
        ApplyPlan::Unchanged => {},
        ApplyPlan::Reference { from, to } => {
            let repointed = engine::repoint(&payload, from, to, limits).map_err(from_engine)?;
            repoint_storage(workspace, storage, index, repointed, from, to)?;
            let to_component = component_of(workspace, to)?;
            add_edge(workspace, owner, to_component, to)?;
            let from_component = component_of(workspace, from)?;
            unregister_if_unused(workspace, owner, from_component, from)?;
        },
        ApplyPlan::Replace { from, to, removals } => {
            let repointed = engine::repoint(&payload, from, to, limits).map_err(from_engine)?;
            repoint_storage(workspace, storage, index, repointed, from, to)?;
            for removal in removals {
                let style_component = component_of(workspace, removal.style)?;
                let sheet = workspace
                    .payload(removal.stylesheet, STYLESHEET_MESSAGE_TYPE)
                    .map_err(lift)?
                    .to_vec();
                let data = engine::stylesheet_without_variation(
                    &sheet,
                    removal.parent,
                    removal.style,
                    limits,
                )
                .map_err(from_engine)?;
                replace_stylesheet(
                    workspace,
                    removal.stylesheet,
                    data,
                    Some(removal.style),
                    None,
                )?;
                workspace.remove_object(removal.style).map_err(lift)?;
                unregister_private_style(
                    workspace,
                    owner,
                    style_component,
                    removal.style,
                    Some(to),
                )?;
            }
            workspace.release_removed_identifier_suffix();
        },
        _ => return Err(refused()),
    }
    Ok(())
}

fn create(
    workspace: &mut Workspace<'_>,
    storage: u64,
    source: ParagraphStyleId,
    name: &ParagraphStyleName,
) -> Result<(), Error> {
    let limits = workspace.wire;
    let context = Context::load(workspace, storage)?;
    context
        .catalog
        .plan_create(source, name)
        .map_err(from_engine)?;
    let source_id = native_id(source);
    let stylesheet = context.catalog.stylesheet();
    let themes = context.themes_with(source_id)?;
    let created = workspace.allocate().map_err(lift)?;
    let identifier = named::generated_identifier(created);
    let payload = named::named_copy(
        context.styles.get(&source_id).ok_or(Error::InvalidSource)?,
        name,
        &identifier,
        limits,
    )
    .map_err(from_engine)?;
    let sheet = workspace
        .payload(stylesheet, STYLESHEET_MESSAGE_TYPE)
        .map_err(lift)?
        .to_vec();
    let sheet_payload =
        named::stylesheet_with_named(&sheet, created, &identifier, limits).map_err(from_engine)?;
    let original = workspace.require(source_id).map_err(lift)?.clone();
    let source_index = original
        .messages
        .iter()
        .position(|message| message.type_ == engine::MESSAGE_TYPE)
        .ok_or(Error::InvalidSource)?;
    let mut object = ArchiveObject::new(
        created,
        original.messages.clone(),
        workspace.archive_limits(),
    )
    .map_err(|_| Error::InvalidSource)?;
    object.archive_info.message_infos = original.archive_info.message_infos.clone();
    object.archive_info.should_merge = original.archive_info.should_merge;
    object
        .replace_message(
            source_index,
            RawMessage {
                type_: engine::MESSAGE_TYPE,
                data: payload,
            },
            workspace.archive_limits(),
        )
        .map_err(|_| Error::InvalidSource)?;
    replace_stylesheet(workspace, stylesheet, sheet_payload, None, Some(created))?;
    let member = workspace
        .component_of(stylesheet)
        .ok_or(Error::InvalidSource)?;
    workspace.insert(&member, object).map_err(lift)?;
    for (theme, index, message_type, payload) in &themes {
        let rewritten =
            named::theme_with_preset(payload, stylesheet, created, limits).map_err(from_engine)?;
        replace_message(
            workspace,
            *theme,
            *index,
            *message_type,
            rewritten,
            |info| {
                info.object_references.push(created);
                Ok(())
            },
        )?;
    }
    let style_component = component_of(workspace, stylesheet)?;
    let uuid = workspace.clone_uuid(None, created).map_err(lift)?;
    workspace
        .registry
        .component_mut(style_component)
        .ok_or(Error::InvalidSource)?
        .add_object_uuid(created, uuid);
    for (theme, ..) in &themes {
        let theme_component = component_of(workspace, *theme)?;
        add_edge(workspace, theme_component, style_component, created)?;
    }
    Ok(())
}

fn rename(
    workspace: &mut Workspace<'_>,
    storage: u64,
    target: ParagraphStyleId,
    name: &ParagraphStyleName,
) -> Result<(), Error> {
    let limits = workspace.wire;
    let context = Context::load(workspace, storage)?;
    if context
        .catalog
        .plan_rename(target, name)
        .map_err(from_engine)?
        .is_none()
    {
        return Ok(());
    }
    let id = native_id(target);
    let object = workspace.require(id).map_err(lift)?;
    let index = object
        .messages
        .iter()
        .position(|message| message.type_ == engine::MESSAGE_TYPE)
        .ok_or(Error::InvalidSource)?;
    let payload =
        named::renamed(&object.messages[index].data, name, limits).map_err(from_engine)?;
    replace_message(workspace, id, index, engine::MESSAGE_TYPE, payload, |_| {
        Ok(())
    })
}

fn delete(
    workspace: &mut Workspace<'_>,
    storage: u64,
    target: ParagraphStyleId,
) -> Result<(), Error> {
    let limits = workspace.wire;
    let context = Context::load(workspace, storage)?;
    context.catalog.plan_delete(target).map_err(from_engine)?;
    let id = native_id(target);
    let stylesheet = context.catalog.stylesheet();
    let themes = context.themes_with(id)?;
    // Nothing but the stylesheet and the listing themes may still name it.
    for component in workspace.package().state.source.components().iter() {
        for source in &component.archive().objects {
            let Some(object_id) = source.archive_info.identifier else {
                continue;
            };
            let Some(object) = workspace.object(object_id) else {
                continue;
            };
            for (message, info) in object
                .messages
                .iter()
                .zip(&object.archive_info.message_infos)
            {
                if !info.object_references.contains(&id) {
                    continue;
                }
                let is_sheet = object_id == stylesheet && message.type_ == STYLESHEET_MESSAGE_TYPE;
                let is_theme = themes
                    .iter()
                    .any(|theme| theme.0 == object_id && theme.2 == message.type_);
                if !is_sheet && !is_theme {
                    return Err(refused());
                }
            }
        }
    }
    for (theme, index, message_type, payload) in &themes {
        let rewritten =
            named::theme_without_preset(payload, stylesheet, id, limits).map_err(from_engine)?;
        replace_message(
            workspace,
            *theme,
            *index,
            *message_type,
            rewritten,
            |info| {
                drop_reference(info, id);
                Ok(())
            },
        )?;
    }
    let sheet = workspace
        .payload(stylesheet, STYLESHEET_MESSAGE_TYPE)
        .map_err(lift)?
        .to_vec();
    let sheet_payload = named::stylesheet_without_named(&sheet, id, limits).map_err(from_engine)?;
    replace_stylesheet(workspace, stylesheet, sheet_payload, Some(id), None)?;
    let component = component_of(workspace, stylesheet)?;
    workspace.remove_object(id).map_err(lift)?;
    remove_edges_to(workspace, component, id, None)?;
    workspace
        .registry
        .component_mut(component)
        .ok_or(Error::InvalidSource)?
        .remove_object_uuids(&[id]);
    workspace.release_removed_identifier_suffix();
    Ok(())
}

fn redefine(workspace: &mut Workspace<'_>, storage: u64) -> Result<(), Error> {
    let limits = workspace.wire;
    let archive_limits = workspace.archive_limits();
    let context = Context::load(workspace, storage)?;
    let applied = named::applied(&context.styles, &context.catalog, context.first, limits)
        .map_err(from_engine)?;
    let redefinition = named::redefine(&context.styles, &context.catalog, context.first, limits)
        .map_err(from_engine)?;
    let style = redefinition.style();
    let index = workspace
        .require(style)
        .map_err(lift)?
        .messages
        .iter()
        .position(|message| message.type_ == engine::MESSAGE_TYPE)
        .ok_or(Error::InvalidSource)?;
    let managed = redefinition.managed_fields();
    let flags = [
        managed.text_fill(),
        managed.text_stroke(),
        managed.text_shadow(),
        managed.text_background(),
    ];
    let (previous, following) = (redefinition.previous_following(), redefinition.following());
    replace_message(
        workspace,
        style,
        index,
        engine::MESSAGE_TYPE,
        redefinition.payload().to_vec(),
        |info| {
            if let Some(previous) = previous {
                info.object_references
                    .retain(|reference| *reference != previous);
            }
            if let Some(following) = following
                && !info.object_references.contains(&following)
            {
                info.object_references.push(following);
            }
            sync_managed(&mut info.field_infos, flags, archive_limits)
        },
    )?;
    apply(workspace, storage, applied.style().id())
}

/// Publish one named-style change of `storage`.
pub(super) fn change(
    workspace: &mut Workspace<'_>,
    storage: u64,
    change: &NamedChange,
) -> Result<(), Error> {
    match change {
        NamedChange::Apply(target) => apply(workspace, storage, *target),
        NamedChange::Create(source, name) => create(workspace, storage, *source, name),
        NamedChange::Rename(target, name) => rename(workspace, storage, *target, name),
        NamedChange::Delete(target) => delete(workspace, storage, *target),
        NamedChange::DeleteApplied(target, replacement) => {
            let limits = workspace.wire;
            let context = Context::load(workspace, storage)?;
            let applied = named::applied(&context.styles, &context.catalog, context.first, limits)
                .map_err(from_engine)?;
            context
                .catalog
                .plan_replacement(&applied, *target, *replacement)
                .map_err(from_engine)?;
            apply(workspace, storage, *replacement)?;
            delete(workspace, storage, *target)
        },
        NamedChange::Redefine => redefine(workspace, storage),
    }
}

/// The storage stylesheet's named paragraph styles, in theme preset order.
pub(super) fn catalog(
    workspace: &Workspace<'_>,
    storage: u64,
) -> Result<Vec<NamedParagraphStyle>, Error> {
    Ok(Context::load(workspace, storage)?.catalog.styles().to_vec())
}

/// The storage's applied named style and whether it overrides it.
pub(super) fn applied(
    workspace: &Workspace<'_>,
    storage: u64,
) -> Result<(NamedParagraphStyle, bool), Error> {
    let context = Context::load(workspace, storage)?;
    let applied = named::applied(
        &context.styles,
        &context.catalog,
        context.first,
        workspace.wire,
    )
    .map_err(from_engine)?;
    Ok((applied.style().clone(), applied.has_overrides()))
}
