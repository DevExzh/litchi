//! Batch source-preserving host authoring over one immutable edit snapshot.

use std::ops::Range;
use std::sync::Arc;

use litchi_opc::{OwnedElementEdit, OwnedElementUpdate};

use super::drawing_ids::DrawingIds;
use super::{
    Commit, Edit, GraphChange, Patch, State, StoryChange, bound, is_annotation, unsupported,
};
use crate::ink::codec::Form;
use crate::ink::graph::{AddDelta, Binding, Delta, NewTarget};
use crate::ink::{Location, authoring, host, placement};
use crate::{Error, Result};

struct ReplacementPlan {
    host_index: usize,
    ink_target: usize,
    image_target: Option<usize>,
}

struct AdditionPlan {
    addition_index: usize,
    ink_target: usize,
    image_target: Option<usize>,
}

enum StructuralEdit {
    Remove,
    Append(Vec<u8>),
}

pub(super) fn commit(mut edit: Edit) -> Result<Commit> {
    bound(
        "annotations",
        edit.base
            .snapshot
            .annotations()
            .len()
            .saturating_sub(edit.removals.len())
            .saturating_add(edit.additions.len()),
        edit.limits.inventory.max_annotations,
    )?;
    let stories =
        crate::package::story::capture(&edit.base.package, edit.limits.inventory.stories)?;
    let mut candidate = edit.base.package.clone();
    let mut drawing_ids = DrawingIds::new(edit.limits.inventory.max_xml_nodes);
    if edit
        .additions
        .iter()
        .any(|addition| addition.style.fallback().is_some())
    {
        for story in stories.stories() {
            drawing_ids.observe(story.source())?;
        }
    }
    let mut addition_order = Vec::new();
    addition_order
        .try_reserve_exact(edit.additions.len())
        .map_err(|source| Error::Allocation {
            resource: "Ink insertion story index",
            source,
        })?;
    addition_order.extend(0..edit.additions.len());
    addition_order.sort_unstable_by_key(|&index| {
        let location = edit.additions[index].destination.story;
        (
            role_index(location.kind()),
            location.position().get(),
            index,
        )
    });
    let mut ordinal = 0usize;
    let mut roles = [0usize; 7];
    let mut inserted = 0usize;
    let mut replaced = 0usize;
    let mut removed = 0usize;
    let mut changes = Vec::new();
    let mut staged_bytes = edit.input_bytes;
    for story in stories.stories() {
        let role = role_index(story.kind());
        let location = Location::new(story.kind(), litchi_core::Position::new(roles[role]));
        roles[role] += 1;
        let hosts = host::capture(story.source(), stories.dialect(), edit.limits.inventory)?;
        let mut removal_hosts = Vec::new();
        let mut replacement_plans = Vec::new();
        let mut addition_plans = Vec::new();
        let mut targets = Vec::new();
        for (host_index, entry) in hosts.iter().enumerate() {
            if !is_annotation(&edit.base.package, story.part(), &entry.anchor)? {
                continue;
            }
            let selected = ordinal;
            ordinal += 1;
            if edit.removals.binary_search(&selected).is_ok() {
                if !entry.removable {
                    return Err(unsupported(
                        "annotation host contains unmodeled sibling or fallback dependencies",
                    ));
                }
                push(&mut removal_hosts, host_index, "Ink selected removals")?;
            }
            if let Ok(replacement_index) = edit
                .replacements
                .binary_search_by_key(&selected, |value| value.position)
            {
                let replacement = &edit.replacements[replacement_index];
                if !entry.removable {
                    return Err(unsupported(
                        "annotation replacement requires a modeled host and complete fallback closure",
                    ));
                }
                match (entry.anchor.form, replacement.fallback.as_ref()) {
                    (Form::Base, Some(_)) => {
                        return Err(unsupported(
                            "base Ink content parts do not have a drawing fallback",
                        ));
                    },
                    (Form::Base, None) => {},
                    (_, None) => {
                        return Err(unsupported(
                            "changed drawing Ink requires a complete replacement fallback image",
                        ));
                    },
                    (_, Some(_)) => {},
                }
                let owner = edit.base.package.get_part(story.part())?;
                let relationship =
                    owner
                        .rels()
                        .get(&entry.anchor.relationship_id)
                        .ok_or_else(|| {
                            Error::InvalidRelationship(
                                "Ink replacement relationship disappeared".into(),
                            )
                        })?;
                let previous = edit
                    .base
                    .package
                    .get_part(&relationship.target_partname()?)?;
                let ink_target = targets.len();
                push(
                    &mut targets,
                    NewTarget::ink(replacement.payload.shared_source(), previous.content_type()),
                    "Ink authored targets",
                )?;
                let image_target = if let Some(image) = &replacement.fallback {
                    let index = targets.len();
                    push(
                        &mut targets,
                        NewTarget::image(image.shared_source(), image.content_type()),
                        "Ink authored targets",
                    )?;
                    Some(index)
                } else {
                    None
                };
                push(
                    &mut replacement_plans,
                    ReplacementPlan {
                        host_index,
                        ink_target,
                        image_target,
                    },
                    "Ink replacement plans",
                )?;
            }
        }
        let key = (role, location.position().get());
        let insertion_start = addition_order.partition_point(|&index| {
            let location = edit.additions[index].destination.story;
            (role_index(location.kind()), location.position().get()) < key
        });
        let insertion_end = addition_order.partition_point(|&index| {
            let location = edit.additions[index].destination.story;
            (role_index(location.kind()), location.position().get()) <= key
        });
        for &addition_index in &addition_order[insertion_start..insertion_end] {
            let addition = &edit.additions[addition_index];
            let ink_target = targets.len();
            push(
                &mut targets,
                NewTarget::ink(
                    addition.payload.shared_source(),
                    addition.style.content_type(),
                ),
                "Ink authored targets",
            )?;
            let image_target = if let Some(image) = addition.style.fallback() {
                let index = targets.len();
                push(
                    &mut targets,
                    NewTarget::image(image.shared_source(), image.content_type()),
                    "Ink authored targets",
                )?;
                Some(index)
            } else {
                None
            };
            push(
                &mut addition_plans,
                AdditionPlan {
                    addition_index,
                    ink_target,
                    image_target,
                },
                "Ink insertion plans",
            )?;
        }
        if removal_hosts.is_empty() && replacement_plans.is_empty() && addition_plans.is_empty() {
            continue;
        }
        let before = candidate.source_xml_part(story.part())?;
        let initial = before
            .bytes()
            .len()
            .saturating_mul(2)
            .saturating_add(
                candidate
                    .source_relationships(story.part())?
                    .bytes()
                    .len()
                    .saturating_mul(2),
            )
            .saturating_add(
                candidate
                    .source_content_types()?
                    .bytes()
                    .len()
                    .saturating_mul(2),
            );
        bound(
            "edit staged bytes",
            staged_bytes.saturating_add(initial),
            edit.limits.max_staged_bytes,
        )?;
        let (added_targets, added) = AddDelta::insert_many(
            &mut candidate,
            story.part(),
            &targets,
            edit.limits.inventory,
        )?;
        let mut retargets = Vec::new();
        for plan in &replacement_plans {
            push(
                &mut retargets,
                (
                    hosts[plan.host_index].anchor_span.clone(),
                    added_targets[plan.ink_target].relationship_id(),
                ),
                "Ink relationship retargets",
            )?;
        }
        let retargeted = placement::retarget_many(
            &before,
            &retargets,
            stories.dialect(),
            edit.limits.inventory.stories.max_story_bytes,
        )?;
        let current_hosts = if retargets.is_empty() {
            hosts
        } else {
            let current =
                host::capture(retargeted.bytes(), stories.dialect(), edit.limits.inventory)?;
            if current.len() != hosts.len() {
                return Err(invalid("retargeting changed the Ink host inventory"));
            }
            current
        };
        let mut fallback_targets = Vec::new();
        for plan in &replacement_plans {
            if let Some(image_target) = plan.image_target {
                let range = current_hosts[plan.host_index]
                    .removal_span
                    .clone()
                    .ok_or_else(|| unsupported("Ink replacement has no complete fallback host"))?;
                push(
                    &mut fallback_targets,
                    (range, added_targets[image_target].relationship_id()),
                    "Ink fallback retargets",
                )?;
            }
        }
        let retargeted = placement::retarget_fallbacks(
            &retargeted,
            &fallback_targets,
            stories.dialect(),
            edit.limits.inventory,
        )?;
        let current_hosts = if fallback_targets.is_empty() {
            current_hosts
        } else {
            let current =
                host::capture(retargeted.bytes(), stories.dialect(), edit.limits.inventory)?;
            if current.len() != current_hosts.len() {
                return Err(invalid(
                    "fallback retargeting changed the Ink host inventory",
                ));
            }
            current
        };
        let mut structural = Vec::new();
        for index in &removal_hosts {
            let entry = &current_hosts[*index];
            let tag = entry
                .removal_start_tag
                .clone()
                .ok_or_else(|| unsupported("Ink removal has no complete source host"))?;
            push_structural(&mut structural, tag, StructuralEdit::Remove)?;
        }
        for plan in &replacement_plans {
            let entry = &current_hosts[plan.host_index];
            if entry.anchor.relationship_id != added_targets[plan.ink_target].relationship_id() {
                return Err(invalid(
                    "Ink retarget readback does not match the selected relationship",
                ));
            }
        }
        if !addition_plans.is_empty() {
            let paragraphs = placement::paragraph_records(
                retargeted.bytes(),
                stories.dialect(),
                edit.limits.inventory,
            )?;
            for plan in &addition_plans {
                let addition = &edit.additions[plan.addition_index];
                let paragraph = paragraphs
                    .get(addition.destination.paragraph.get())
                    .ok_or_else(|| invalid("Ink insertion paragraph selector is out of range"))?;
                if !paragraph.eligible {
                    return Err(unsupported(
                        "Ink insertion paragraph has unmodeled alternative-content dependencies",
                    ));
                }
                let image_id = plan
                    .image_target
                    .map(|index| added_targets[index].relationship_id());
                let drawing_id = if image_id.is_some() {
                    drawing_ids.allocate()?
                } else {
                    0
                };
                let fragment = authoring::render(
                    &addition.style,
                    stories.dialect(),
                    added_targets[plan.ink_target].relationship_id(),
                    image_id,
                    drawing_id,
                )?;
                let run = wrap_run(
                    &fragment,
                    stories.dialect(),
                    edit.limits.inventory.stories.max_story_bytes,
                )?;
                push_structural(
                    &mut structural,
                    paragraph.span.clone(),
                    StructuralEdit::Append(run),
                )?;
            }
        }
        structural.sort_unstable_by_key(|(range, order, _)| (range.start, *order));
        let mut updates = Vec::new();
        for (tag, _, operation) in &structural {
            let operation = match operation {
                StructuralEdit::Remove => OwnedElementEdit::Remove,
                StructuralEdit::Append(bytes) => OwnedElementEdit::AppendChild(bytes),
            };
            push(
                &mut updates,
                OwnedElementUpdate {
                    start_tag: tag.clone(),
                    edit: operation,
                },
                "Ink XML updates",
            )?;
        }
        let after =
            retargeted.update_elements(&updates, edit.limits.inventory.stories.max_story_bytes)?;
        let mut previous_refs = host::reference_values(before.bytes(), edit.limits.inventory)?;
        previous_refs.sort_unstable();
        let mut remaining_refs = host::reference_values(after.bytes(), edit.limits.inventory)?;
        remaining_refs.sort_unstable();
        let mut removed_ids = Vec::new();
        for relationship in candidate.get_part(story.part())?.rels().iter() {
            if previous_refs
                .binary_search_by(|value| value.as_str().cmp(relationship.r_id()))
                .is_ok()
                && remaining_refs
                    .binary_search_by(|value| value.as_str().cmp(relationship.r_id()))
                    .is_err()
            {
                push(
                    &mut removed_ids,
                    relationship.r_id().to_owned(),
                    "Ink released relationship IDs",
                )?;
            }
        }
        removed_ids.sort_unstable();
        candidate.try_replace_owned_xml_part(before.bytes(), after.clone())?;
        let cleanup = Delta::remove(
            &mut candidate,
            story.part(),
            &removed_ids,
            edit.limits.inventory,
        )?;
        staged_bytes = staged_bytes
            .saturating_add(before.bytes().len())
            .saturating_add(after.bytes().len())
            .saturating_add(added.retained_bytes()?)
            .saturating_add(cleanup.retained_bytes()?);
        bound(
            "edit staged bytes",
            staged_bytes,
            edit.limits.max_staged_bytes,
        )?;
        for (target, requested) in added_targets.iter().zip(&targets) {
            if candidate.get_part(target.part())?.blob() != requested.as_bytes() {
                return Err(invalid(
                    "authored Ink resource readback differs from its prepared input",
                ));
            }
        }
        inserted += addition_plans.len();
        replaced += replacement_plans.len();
        removed += removal_hosts.len();
        push(
            &mut changes,
            StoryChange {
                owner: story.part().clone(),
                before,
                after,
                graph: GraphChange::Authored {
                    added,
                    removed: cleanup,
                },
            },
            "Ink authored story changes",
        )?;
    }
    if inserted != edit.additions.len()
        || replaced != edit.replacements.len()
        || removed != edit.removals.len()
        || ordinal != edit.base.snapshot.annotations().len()
    {
        return Err(invalid(
            "Ink authoring selectors did not resolve against the complete source snapshot",
        ));
    }
    let snapshot = crate::ink::package::load(&candidate, edit.limits.inventory)?;
    if snapshot.annotations().len() != ordinal - removed + inserted {
        return Err(invalid(
            "Ink authoring changed an unexpected annotation count",
        ));
    }
    let binding = Binding::capture(&candidate, edit.limits.inventory)?;
    let intent = edit.take_semantic_intent()?;
    let after = Arc::new(State {
        package: candidate,
        binding,
        snapshot,
    });
    Ok(Commit {
        patch: Patch {
            before: edit.base,
            after,
            changes: Arc::new(changes),
            reversed: false,
            limits: edit.limits,
            intent,
        },
    })
}

fn wrap_run(
    fragment: &[u8],
    dialect: crate::package::story::StoryDialect,
    maximum: usize,
) -> Result<Vec<u8>> {
    let opening: &[u8] = match dialect {
        crate::package::story::StoryDialect::Transitional => {
            b"<w:r xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">"
        },
        crate::package::story::StoryDialect::Strict => {
            b"<w:r xmlns:w=\"http://purl.oclc.org/ooxml/wordprocessingml/main\">"
        },
    };
    let length = opening
        .len()
        .saturating_add(fragment.len())
        .saturating_add(6);
    bound("Ink authored run bytes", length, maximum)?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(length)
        .map_err(|source| Error::Allocation {
            resource: "Ink authored run",
            source,
        })?;
    output.extend_from_slice(opening);
    output.extend_from_slice(fragment);
    output.extend_from_slice(b"</w:r>");
    Ok(output)
}

fn push_structural(
    items: &mut Vec<(Range<usize>, usize, StructuralEdit)>,
    tag: Range<usize>,
    operation: StructuralEdit,
) -> Result<()> {
    let order = items.len();
    push(items, (tag, order, operation), "Ink structural edits")
}

fn push<T>(items: &mut Vec<T>, value: T, resource: &'static str) -> Result<()> {
    items
        .try_reserve(1)
        .map_err(|source| Error::Allocation { resource, source })?;
    items.push(value);
    Ok(())
}

fn role_index(kind: crate::package::story::StoryKind) -> usize {
    use crate::package::story::StoryKind;
    match kind {
        StoryKind::Main => 0,
        StoryKind::Header => 1,
        StoryKind::Footer => 2,
        StoryKind::Footnotes => 3,
        StoryKind::Endnotes => 4,
        StoryKind::Comments => 5,
        StoryKind::Glossary => 6,
    }
}

fn invalid(message: &'static str) -> Error {
    Error::Invalid(message.into())
}
