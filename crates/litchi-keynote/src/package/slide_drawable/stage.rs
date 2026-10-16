//! Staging of geometry, removal and z-order operations in a structural
//! workspace.

use std::collections::{BTreeSet, HashMap};

use litchi_iwa_common::WireLimits;
use litchi_iwa_core::RawMessage;

use super::super::structure::StructureError;
use super::super::structure::refs;
use super::super::structure::wire::{self, Fields, LENGTH_DELIMITED};
use super::super::structure::workspace::Workspace;
use super::inventory::{
    self, SLIDE, SLIDE_OWNED_DRAWABLES, SLIDE_Z_ORDER, drawable_path, unique_message,
};
use super::{Error, Geometry, Unsupported, lift};

/// Re-encode `payload` with `edit` applied to the message at `path`.
pub(super) fn edit_nested(
    payload: &[u8],
    path: &[u32],
    wire: WireLimits,
    edit: &mut dyn FnMut(&mut Fields) -> Result<(), StructureError>,
) -> Result<Vec<u8>, StructureError> {
    let mut fields = Fields::parse(payload, wire)?;
    if let Some((first, rest)) = path.split_first() {
        let inner = fields
            .unique(*first)?
            .ok_or(StructureError::Unexpected("nested message"))?
            .message()?
            .to_vec();
        let edited = edit_nested(&inner, rest, wire, edit)?;
        fields.set_message(*first, &edited);
    } else {
        edit(&mut fields)?;
    }
    Ok(fields.encode())
}

/// Replace one payload, keeping every untouched header byte.
pub(super) fn replace_payload(
    workspace: &mut Workspace<'_>,
    identifier: u64,
    index: usize,
    message_type: u32,
    data: Vec<u8>,
) -> Result<(), StructureError> {
    let limits = workspace.archive_limits();
    workspace
        .object_mut(identifier)?
        .replace_message_preserving_header_with_limits(
            index,
            RawMessage {
                type_: message_type,
                data,
            },
            limits,
        )?;
    Ok(())
}

/// Patch a `TSD.GeometryArchive` field list in place.
pub(super) fn patch_geometry(
    fields: &mut Fields,
    geometry: Geometry,
    wire: WireLimits,
) -> Result<(), StructureError> {
    for (number, first, second) in [
        (1, geometry.x, geometry.y),
        (2, geometry.width, geometry.height),
    ] {
        let mut inner = match fields.unique(number)? {
            Some(field) => Fields::parse(field.message()?, wire)?,
            None => Fields::parse(&[], wire)?,
        };
        inner.set_float(1, first);
        inner.set_float(2, second);
        fields.set_message(number, &inner.encode());
    }
    match geometry.rotation {
        Some(rotation) => fields.set_float(4, rotation),
        None => {
            fields.remove(4);
        },
    }
    Ok(())
}

/// Set the geometry of one drawable payload.
pub(super) fn set_geometry(
    workspace: &mut Workspace<'_>,
    drawable: u64,
    message_index: usize,
    message_type: u32,
    geometry: Geometry,
) -> Result<(), Error> {
    let wire = workspace.wire;
    let payload = workspace
        .require(drawable)
        .map_err(lift)?
        .messages
        .get(message_index)
        .ok_or(Error::InvalidSource)?
        .data
        .clone();
    let mut path = drawable_path(message_type).to_vec();
    path.push(1);
    let data = edit_nested(&payload, &path, wire, &mut |fields| {
        patch_geometry(fields, geometry, wire)
    })
    .map_err(lift)?;
    replace_payload(workspace, drawable, message_index, message_type, data).map_err(lift)
}

/// Every object reference held by one staged object: the header's aggregate
/// and field lists, and the schema walk of admitted payloads.
fn outgoing(workspace: &Workspace<'_>, identifier: u64) -> Result<Vec<u64>, StructureError> {
    let object = workspace.require(identifier)?;
    let mut found = BTreeSet::new();
    for (message, info) in object
        .messages
        .iter()
        .zip(&object.archive_info.message_infos)
    {
        found.extend(info.object_references.iter().copied());
        for field in &info.field_infos {
            found.extend(field.object_references.iter().copied());
        }
        if refs::admitted(message.type_) {
            found.extend(refs::object_references(
                message.type_,
                &message.data,
                workspace.wire,
            )?);
        }
    }
    found.remove(&identifier);
    Ok(found.into_iter().collect())
}

const SLIDE_TAGS: u32 = 28;
const SLIDE_INSTRUCTIONAL_TEXT: u32 = 45;

/// Whether message `entry` holds a reference to `drawable` in field `number`.
fn entry_names(
    entry: &[u8],
    number: u32,
    drawable: u64,
    wire: WireLimits,
) -> Result<bool, StructureError> {
    match Fields::parse(entry, wire)?.unique(number)? {
        Some(field) => Ok(wire::reference_identifier(field.message()?, wire)? == drawable),
        None => Ok(false),
    }
}

/// What a removal took out of the package.
pub(super) struct Removal {
    pub(super) objects: usize,
}

/// Remove one drawable and its private object closure from its slide.
pub(super) fn remove(
    workspace: &mut Workspace<'_>,
    slide: u64,
    drawable: u64,
) -> Result<Removal, Error> {
    let wire = workspace.wire;
    let member = workspace.component_of(slide).ok_or(Error::InvalidSource)?;
    if workspace.component_of(drawable).as_deref() != Some(member.as_str()) {
        return Err(Error::Unsupported {
            reason: Unsupported::ForeignComponent,
        });
    }
    let members = workspace.member_objects(&member);
    let mut graph = HashMap::with_capacity(members.len());
    for identifier in &members {
        graph.insert(*identifier, outgoing(workspace, *identifier).map_err(lift)?);
    }
    // Closure of objects reachable from the drawable inside its component,
    // shrunk until no retained member outside it references a member.
    let mut shared = BTreeSet::from([slide]);
    let closure = loop {
        let mut closure = BTreeSet::from([drawable]);
        let mut pending = vec![drawable];
        while let Some(current) = pending.pop() {
            for target in graph.get(&current).map(Vec::as_slice).unwrap_or(&[]) {
                if graph.contains_key(target) && !shared.contains(target) && closure.insert(*target)
                {
                    pending.push(*target);
                }
            }
        }
        let mut grew = false;
        for (source, targets) in &graph {
            if closure.contains(source) {
                continue;
            }
            for target in targets {
                if !closure.contains(target) {
                    continue;
                }
                if *target == drawable {
                    if *source != slide {
                        return Err(Error::Unsupported {
                            reason: Unsupported::ExternalReferences,
                        });
                    }
                } else if shared.insert(*target) {
                    grew = true;
                }
            }
        }
        if !grew {
            break closure;
        }
    };
    let closure = closure.into_iter().collect::<Vec<_>>();
    let component_id = workspace
        .registry
        .component_for_member(&member)
        .ok_or(Error::InvalidSource)?
        .identifier;
    for other in workspace.registry.components() {
        if other.identifier == component_id {
            continue;
        }
        if other.external_references().iter().any(|reference| {
            reference.component == component_id
                && reference
                    .object
                    .is_some_and(|object| closure.contains(&object))
        }) {
            return Err(Error::Unsupported {
                reason: Unsupported::ExternalReferences,
            });
        }
    }

    // The slide may reference the drawable only from its owned list and
    // z-order.
    let slide_object = workspace.require(slide).map_err(lift)?;
    let (slide_index, slide_payload) = unique_message(slide_object, SLIDE).map_err(lift)?;
    let slide_payload = slide_payload.to_vec();
    let mut fields = Fields::parse(&slide_payload, wire).map_err(lift)?;
    let owned = inventory::references(&fields, SLIDE_OWNED_DRAWABLES, wire).map_err(lift)?;
    let z_order = inventory::references(&fields, SLIDE_Z_ORDER, wire).map_err(lift)?;
    // Template tags (`sage_tag_to_info_map`, field 28) and instructional
    // text entries (field 45) that name the drawable go with it.
    let tag_entries = fields
        .all(SLIDE_TAGS)
        .map(|field| field.message().map(<[u8]>::to_vec))
        .collect::<Result<Vec<_>, _>>()
        .map_err(lift)?;
    let mut kept_tags = Vec::with_capacity(tag_entries.len());
    for entry in tag_entries {
        if entry_names(&entry, 2, drawable, wire).map_err(lift)? {
            continue;
        }
        kept_tags.push(entry);
    }
    if kept_tags.len() != fields.all(SLIDE_TAGS).count() {
        fields.replace_all(SLIDE_TAGS, LENGTH_DELIMITED, &kept_tags);
    }
    if let Some(map) = fields.unique(SLIDE_INSTRUCTIONAL_TEXT).map_err(lift)? {
        let mut map = Fields::parse(map.message().map_err(lift)?, wire).map_err(lift)?;
        let entries = map
            .all(1)
            .map(|field| field.message().map(<[u8]>::to_vec))
            .collect::<Result<Vec<_>, _>>()
            .map_err(lift)?;
        let mut kept = Vec::with_capacity(entries.len());
        for entry in &entries {
            if !entry_names(entry, 1, drawable, wire).map_err(lift)? {
                kept.push(entry.clone());
            }
        }
        if kept.len() != entries.len() {
            map.replace_all(1, LENGTH_DELIMITED, &kept);
            fields.set_message(SLIDE_INSTRUCTIONAL_TEXT, &map.encode());
        }
    }
    let walked = refs::object_references(SLIDE, &fields.encode(), wire)
        .map_err(StructureError::from)
        .map_err(lift)?;
    let listed = owned.iter().filter(|id| **id == drawable).count()
        + z_order.iter().filter(|id| **id == drawable).count();
    if owned.iter().filter(|id| **id == drawable).count() != 1
        || walked.iter().filter(|id| **id == drawable).count() != listed
    {
        return Err(Error::Unsupported {
            reason: Unsupported::ExternalReferences,
        });
    }
    let encode = |values: &[u64]| {
        values
            .iter()
            .filter(|id| **id != drawable)
            .map(|id| wire::reference(*id))
            .collect::<Vec<_>>()
    };
    fields.replace_all(SLIDE_OWNED_DRAWABLES, LENGTH_DELIMITED, &encode(&owned));
    fields.replace_all(SLIDE_Z_ORDER, LENGTH_DELIMITED, &encode(&z_order));
    let limits = workspace.archive_limits();
    workspace
        .object_mut(slide)
        .map_err(lift)?
        .replace_message_pruning_object_references_preserving_header_with_limits(
            slide_index,
            RawMessage {
                type_: SLIDE,
                data: fields.encode(),
            },
            &[drawable],
            limits,
        )
        .map_err(StructureError::from)
        .map_err(lift)?;

    {
        let component = workspace
            .registry
            .component_mut(component_id)
            .ok_or(Error::InvalidSource)?;
        component.remove_object_uuids(&closure);
        for identifier in &closure {
            component.remove_data_owner(*identifier);
        }
    }
    for identifier in &closure {
        workspace.remove_object(*identifier).map_err(lift)?;
    }
    Ok(Removal {
        objects: closure.len(),
    })
}

/// Move one drawable to `position` in its slide's back-to-front z-order.
pub(super) fn move_to(
    workspace: &mut Workspace<'_>,
    slide: u64,
    drawable: u64,
    position: usize,
) -> Result<(), Error> {
    let wire = workspace.wire;
    let slide_object = workspace.require(slide).map_err(lift)?;
    let (index, payload) = unique_message(slide_object, SLIDE).map_err(lift)?;
    let mut fields = Fields::parse(payload, wire).map_err(lift)?;
    let mut order = inventory::references(&fields, SLIDE_Z_ORDER, wire).map_err(lift)?;
    let current = order
        .iter()
        .position(|id| *id == drawable)
        .ok_or(Error::Unsupported {
            reason: Unsupported::ExternalReferences,
        })?;
    if position >= order.len() {
        return Err(Error::ZPositionOutOfRange {
            position: litchi_core::Position::new(position),
            count: order.len(),
        });
    }
    let moved = order.remove(current);
    order.insert(position, moved);
    let encoded = order
        .iter()
        .map(|id| wire::reference(*id))
        .collect::<Vec<_>>();
    fields.replace_all(SLIDE_Z_ORDER, LENGTH_DELIMITED, &encoded);
    replace_payload(workspace, slide, index, SLIDE, fields.encode()).map_err(lift)
}

/// The changed middles of `before` and `after`: what is left of each once
/// their common prefix and suffix (whole characters) are removed.
pub(super) fn changed_middle<'a, 'b>(before: &'a str, after: &'b str) -> (&'a str, &'b str) {
    let prefix = before
        .char_indices()
        .zip(after.chars())
        .find(|((_, left), right)| left != right)
        .map_or_else(|| before.len().min(after.len()), |((index, _), _)| index);
    // `prefix` is a character boundary of both (the characters before it are
    // equal).
    let (before_rest, after_rest) = (&before[prefix..], &after[prefix..]);
    let suffix = before_rest
        .chars()
        .rev()
        .zip(after_rest.chars().rev())
        .take_while(|(left, right)| left == right)
        .map(|(left, _)| left.len_utf8())
        .sum::<usize>();
    (
        &before_rest[..before_rest.len() - suffix],
        &after_rest[..after_rest.len() - suffix],
    )
}

/// Refuse a storage other objects or components reference, or one outside
/// the slide's component; return the slide's component.
fn own_storage(
    workspace: &Workspace<'_>,
    slide: u64,
    drawable: u64,
    storage: u64,
) -> Result<String, Error> {
    let member = workspace.component_of(slide).ok_or(Error::InvalidSource)?;
    if workspace.component_of(drawable).as_deref() != Some(member.as_str())
        || workspace.component_of(storage).as_deref() != Some(member.as_str())
    {
        return Err(Error::Unsupported {
            reason: Unsupported::ForeignComponent,
        });
    }
    // The storage is the drawable's own: no other object of the component
    // and no other component references it.
    for identifier in workspace.member_objects(&member) {
        if identifier == drawable || identifier == storage {
            continue;
        }
        if outgoing(workspace, identifier)
            .map_err(lift)?
            .contains(&storage)
        {
            return Err(Error::Unsupported {
                reason: Unsupported::ExternalReferences,
            });
        }
    }
    let component_id = workspace
        .registry
        .component_for_member(&member)
        .ok_or(Error::InvalidSource)?
        .identifier;
    for other in workspace.registry.components() {
        if other.identifier != component_id
            && other.external_references().iter().any(|reference| {
                reference.component == component_id && reference.object == Some(storage)
            })
        {
            return Err(Error::Unsupported {
                reason: Unsupported::ExternalReferences,
            });
        }
    }
    Ok(member)
}

/// Rewrite the text of one text box or shape: only the changed span of its
/// own storage, which no object but the drawable may reference.
pub(super) fn set_text(
    workspace: &mut Workspace<'_>,
    slide: u64,
    drawable: u64,
    storage: u64,
    text: &str,
    source: &super::Package,
) -> Result<(), Error> {
    let wire = workspace.wire;
    own_storage(workspace, slide, drawable, storage)?;
    let before = inventory::storage_text(workspace, storage, wire).map_err(lift)?;
    let (removed, inserted) = changed_middle(&before, text);
    if [removed, inserted]
        .iter()
        .any(|middle| middle.contains(['\u{fffc}', '\u{000e}']))
    {
        return Err(Error::Unsupported {
            reason: Unsupported::InlineObjects,
        });
    }
    // UTF-16 offsets of the changed span in the storage text.
    let prefix_bytes = removed.as_ptr() as usize - before.as_ptr() as usize;
    let start = before[..prefix_bytes].encode_utf16().count();
    let end = start + removed.encode_utf16().count();
    let object = workspace.require(storage).map_err(lift)?;
    let (index, payload) = unique_message(object, inventory::STORAGE).map_err(lift)?;
    let limits = rewrite_limits(source, workspace.archive_limits().max_message_bytes())?;
    let rewrite = litchi_iwa_text_wire::rewrite_storage_text_with_behavior_and_limits(
        payload,
        start..end,
        inserted,
        litchi_iwa_text_wire::RewriteBehavior::PreserveOnEqualText,
        limits,
    )
    .map_err(map_rewrite_error)?;
    if !rewrite.removed_object_references().is_empty()
        || !rewrite.removed_object_references_by_field().is_empty()
        || rewrite.object_reference_occurrences_before()
            != rewrite.object_reference_occurrences_after()
    {
        return Err(Error::Unsupported {
            reason: Unsupported::InlineObjects,
        });
    }
    replace_payload(
        workspace,
        storage,
        index,
        inventory::STORAGE,
        rewrite.into_bytes(),
    )
    .map_err(lift)
}

/// The text-rewrite bounds: the wire defaults narrowed by the package's
/// semantic text limits and the archive's message ceiling.
fn rewrite_limits(
    source: &super::Package,
    max_message_bytes: usize,
) -> Result<litchi_iwa_text_wire::RewriteLimits, Error> {
    let defaults = litchi_iwa_text_wire::RewriteLimits::default();
    let message_bytes = defaults.max_message_bytes().min(max_message_bytes);
    let semantic = source.semantic_limits();
    litchi_iwa_text_wire::RewriteLimits::new(
        message_bytes,
        defaults.max_fields(),
        defaults.max_nesting(),
        defaults.max_fragments().min(semantic.max_text_fragments()),
        defaults
            .max_text_bytes()
            .min(semantic.max_text_bytes())
            .min(message_bytes),
        defaults.max_table_entries(),
        defaults.max_object_references(),
        defaults.max_output_bytes().min(message_bytes),
        defaults.max_rewrite_work(),
    )
    .map_err(map_rewrite_error)
}

fn map_rewrite_error(error: litchi_iwa_text_wire::RewriteError) -> Error {
    use litchi_iwa_text_wire::RewriteError;
    match error {
        RewriteError::LimitExceeded {
            resource,
            observed,
            limit,
        } => Error::LimitExceeded {
            kind: match resource {
                "text bytes" => super::LimitKind::TextBytes,
                "text fragments" => super::LimitKind::TextFragments,
                "fields" | "aggregate field count" | "generated text fields" => {
                    super::LimitKind::WireFields
                },
                "nesting" => super::LimitKind::WireNesting,
                "rewrite work" | "aggregate nested scan bytes" => super::LimitKind::Work,
                "output bytes" | "rewritten storage bytes" | "appended text field" => {
                    super::LimitKind::OutputBytes
                },
                _ => super::LimitKind::WireBytes,
            },
            observed: observed as u64,
            maximum: limit as u64,
        },
        RewriteError::Allocation { amount, .. } => Error::Allocation { amount },
        _ => Error::InvalidSource,
    }
}

/// `TSWP.CharacterStyleArchive`.
const CHARACTER_STYLE: u32 = 2_021;
/// `TSS.StylesheetArchive`.
const STYLESHEET: u32 = 401;
/// `TSWP.StorageArchive.style_sheet`.
const STORAGE_STYLESHEET: u32 = 2;
/// The message versions of a new character style when the stylesheet has
/// none to copy.
const CHARACTER_STYLE_VERSION: [u32; 3] = [1, 0, 5];

fn character_properties(style: &super::Style) -> litchi_iwa_text_wire::CharacterProperties {
    let mut properties = litchi_iwa_text_wire::CharacterProperties::new();
    properties.bold = style.bold;
    properties.italic = style.italic;
    properties.underline = style.underline;
    properties.size = style.font_size;
    properties.font_name.clone_from(&style.font_name);
    properties.color = style.color;
    properties
}

/// Style one span (characters) of a text box's or shape's own storage:
/// each character-style run inside the span gets a new variation of its
/// style (or a parentless one) carrying `style`, registered in the
/// storage's stylesheet; the storage's run table is split at the span's
/// ends.
pub(super) fn set_text_style(
    workspace: &mut Workspace<'_>,
    slide: u64,
    drawable: u64,
    storage: u64,
    span: super::TextSpan,
    style: &super::Style,
) -> Result<(), Error> {
    let wire = workspace.wire;
    let limits = workspace.archive_limits();
    let member = own_storage(workspace, slide, drawable, storage)?;
    let text = inventory::storage_text(workspace, storage, wire).map_err(lift)?;
    let utf16 = |chars: usize| -> usize { text.chars().take(chars).map(char::len_utf16).sum() };
    let range = utf16(span.start())..utf16(span.end());
    let text_len = text.encode_utf16().count();
    let object = workspace.require(storage).map_err(lift)?;
    let (storage_index, payload) = unique_message(object, inventory::STORAGE).map_err(lift)?;
    let payload = payload.to_vec();
    let storage_refs_before = object
        .archive_info
        .message_infos
        .get(storage_index)
        .ok_or(Error::InvalidSource)?
        .object_references
        .clone();
    let stylesheet = Fields::parse(&payload, wire)
        .map_err(lift)?
        .unique(STORAGE_STYLESHEET)
        .map_err(lift)?
        .map(|field| {
            field
                .message()
                .and_then(|reference| wire::reference_identifier(reference, wire))
        })
        .transpose()
        .map_err(lift)?
        .ok_or(Error::Unsupported {
            reason: Unsupported::TextStyling,
        })?;
    let sheet_member = workspace
        .component_of(stylesheet)
        .ok_or(Error::InvalidSource)?;
    let map_style = |_error| Error::InvalidSource;
    let previous = litchi_iwa_text_wire::styles_in_span(&payload, text_len, range.clone())
        .map_err(map_style)?;
    let properties = character_properties(style);
    // Versions of an existing character style of the stylesheet.
    let versions = workspace
        .member_objects(&sheet_member)
        .into_iter()
        .filter_map(|identifier| workspace.object(identifier))
        .find_map(|object| {
            object
                .messages
                .iter()
                .zip(&object.archive_info.message_infos)
                .find(|(message, _)| message.type_ == CHARACTER_STYLE)
                .map(|(_, info)| info.versions.clone())
        })
        .unwrap_or_else(|| CHARACTER_STYLE_VERSION.to_vec());
    let mut created: Vec<(Option<u64>, u64)> = Vec::with_capacity(previous.len());
    for parent in &previous {
        if let Some(parent) = parent
            && workspace.component_of(*parent).as_deref() != Some(sheet_member.as_str())
        {
            return Err(Error::Unsupported {
                reason: Unsupported::ForeignComponent,
            });
        }
        let identifier = workspace.allocate().map_err(lift)?;
        let data =
            litchi_iwa_text_wire::character_style_variation(*parent, stylesheet, &properties);
        let mut references = Vec::with_capacity(2);
        references.extend(*parent);
        references.push(stylesheet);
        let mut object = litchi_iwa_core::ArchiveObject::new_with_limits(
            identifier,
            vec![RawMessage {
                type_: CHARACTER_STYLE,
                data,
            }],
            limits,
        )
        .map_err(|error| lift(StructureError::from(error)))?;
        let info = object
            .archive_info
            .message_infos
            .first_mut()
            .ok_or(Error::InvalidSource)?;
        info.versions.clone_from(&versions);
        info.object_references = references;
        workspace.insert(&sheet_member, object).map_err(lift)?;
        // Register the style in its stylesheet.
        let sheet_object = workspace.require(stylesheet).map_err(lift)?;
        let (sheet_index, sheet_payload) =
            unique_message(sheet_object, STYLESHEET).map_err(lift)?;
        let sheet_refs = sheet_object
            .archive_info
            .message_infos
            .get(sheet_index)
            .ok_or(Error::InvalidSource)?
            .object_references
            .clone();
        let sheet_data =
            litchi_iwa_text_wire::stylesheet_with_style(sheet_payload, identifier, *parent)
                .map_err(map_style)?;
        let mut sheet_after = sheet_refs.clone();
        sheet_after.push(identifier);
        workspace
            .object_mut(stylesheet)
            .map_err(lift)?
            .replace_message_transitioning_object_references_preserving_header_with_limits(
                sheet_index,
                RawMessage {
                    type_: STYLESHEET,
                    data: sheet_data,
                },
                litchi_iwa_core::archive::ObjectReferenceTransition {
                    aggregate_before: &sheet_refs,
                    aggregate_after: &sheet_after,
                    fields: &[],
                },
                limits,
            )
            .map_err(|error| lift(StructureError::from(error)))?;
        created.push((*parent, identifier));
    }
    let restyled = litchi_iwa_text_wire::restyle_span(&payload, text_len, range, |style| {
        created
            .iter()
            .find(|(parent, _)| *parent == style)
            .map(|(_, created)| *created)
    })
    .map_err(map_style)?;
    // The storage now references the new styles and no longer the styles
    // the span covered completely.
    let runs = litchi_iwa_text_wire::char_style_runs(&restyled).map_err(map_style)?;
    let still: BTreeSet<u64> = runs.iter().filter_map(|run| run.style).collect();
    let mut storage_after: Vec<u64> = storage_refs_before
        .iter()
        .copied()
        .filter(|identifier| !previous.contains(&Some(*identifier)) || still.contains(identifier))
        .collect();
    for (_, identifier) in &created {
        storage_after.push(*identifier);
    }
    workspace
        .object_mut(storage)
        .map_err(lift)?
        .replace_message_transitioning_object_references_preserving_header_with_limits(
            storage_index,
            RawMessage {
                type_: inventory::STORAGE,
                data: restyled,
            },
            litchi_iwa_core::archive::ObjectReferenceTransition {
                aggregate_before: &storage_refs_before,
                aggregate_after: &storage_after,
                fields: &[],
            },
            limits,
        )
        .map_err(|error| lift(StructureError::from(error)))?;
    // Registry: the new styles' UUIDs in the stylesheet's component, and
    // the slide component's references to them.
    let sheet_component = workspace
        .registry
        .component_for_member(&sheet_member)
        .ok_or(Error::InvalidSource)?
        .identifier;
    let slide_component = workspace
        .registry
        .component_for_member(&member)
        .ok_or(Error::InvalidSource)?
        .identifier;
    let mut uuids = Vec::with_capacity(created.len());
    for (_, identifier) in &created {
        uuids.push((
            *identifier,
            workspace.clone_uuid(None, *identifier).map_err(lift)?,
        ));
    }
    let entry = workspace
        .registry
        .component_mut(sheet_component)
        .ok_or(Error::InvalidSource)?;
    for (identifier, uuid) in uuids {
        entry.add_object_uuid(identifier, uuid);
    }
    if slide_component != sheet_component {
        let entry = workspace
            .registry
            .component_mut(slide_component)
            .ok_or(Error::InvalidSource)?;
        for (_, identifier) in &created {
            entry.add_external_reference(super::super::structure::metadata::ExternalReference {
                component: sheet_component,
                object: Some(*identifier),
                is_weak: None,
            });
        }
    }
    Ok(())
}
