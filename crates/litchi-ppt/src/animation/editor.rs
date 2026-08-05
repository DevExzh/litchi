//! Transactional animation/timing mutation for persisted slide and master records.
//!
//! Actions, hyperlinks, sounds, commands, and media references remain inert
//! record metadata. This module never resolves or executes them.

use super::{
    AnimationInfo, BuildList, BuildListEntry, ExtendedTimeNode, SlideAnimationExtension,
    TimeModifier, TimeNodeBehavior, TimeSubEffectBehavior, TimeVisualElement, parse_animation_info,
    parse_slide_animation_extension, write_animation_info, write_build_list,
    write_extended_time_node,
};
use crate::consts::PptRecordType;
use crate::embedded::object::editor::Editor as ObjectEditor;
use crate::package::{PptError, Result};
use crate::records::PptRecord;
use std::collections::{BTreeSet, HashSet};

const ESCHER_SP_CONTAINER: u16 = 0xF004;
const ESCHER_SP: u16 = 0xF00A;
const ESCHER_CLIENT_DATA: u16 = 0xF011;

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Scope {
    Slide,
    MainMaster,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct EditorLimits {
    pub max_persist_records: usize,
    pub max_record_bytes: usize,
    pub max_timeline_nodes: usize,
    pub max_timeline_depth: usize,
    pub max_build_entries: usize,
    pub max_shapes: usize,
}

impl Default for EditorLimits {
    fn default() -> Self {
        Self {
            max_persist_records: 65_536,
            max_record_bytes: 64 * 1024 * 1024,
            max_timeline_nodes: 65_536,
            max_timeline_depth: 128,
            max_build_entries: 65_536,
            max_shapes: 65_536,
        }
    }
}

#[derive(Clone, Debug, PartialEq)]
pub struct Timeline {
    pub persist_id: u32,
    pub scope: Scope,
    pub extension: SlideAnimationExtension,
}

#[derive(Clone, Debug)]
pub struct LegacyShapeAnimation {
    pub persist_id: u32,
    pub scope: Scope,
    pub shape_id: u32,
    pub animation: AnimationInfo,
}

#[derive(Clone)]
struct PersistAnimation {
    persist_id: u32,
    scope: Scope,
    record: Vec<u8>,
    extension_payload: Option<Vec<u8>>,
    extension: SlideAnimationExtension,
    shape_ids: BTreeSet<u32>,
    legacy: Vec<LegacyShapeAnimation>,
}

#[derive(Clone)]
pub struct Editor {
    package: ObjectEditor,
    entries: Vec<PersistAnimation>,
    limits: EditorLimits,
    changed: bool,
}

impl Editor {
    pub fn open(bytes: Vec<u8>, limits: EditorLimits) -> Result<Self> {
        validate_limits(limits)?;
        let package = ObjectEditor::open_records(bytes)?;
        let ids = package.persist_ids();
        if ids.len() > limits.max_persist_records {
            return invalid("persisted-record count exceeds animation editor limit");
        }
        let mut entries = Vec::new();
        for persist_id in ids {
            let record = package.persisted_record(persist_id)?;
            if record.len() > limits.max_record_bytes {
                return invalid("persisted record exceeds animation editor limit");
            }
            let kind = raw_type(&record)?;
            let scope = if kind == PptRecordType::Slide.as_u16() {
                Scope::Slide
            } else if kind == PptRecordType::MainMaster.as_u16() {
                Scope::MainMaster
            } else {
                continue;
            };
            let (parsed, consumed) = PptRecord::parse(&record, 0)?;
            if consumed != record.len() || parsed.data_length as usize + 8 != record.len() {
                return corrupted("persisted slide/master record has trailing or truncated bytes");
            }
            let extension_payload = find_ppt10_payload(&parsed)?;
            let extension = extension_payload
                .as_deref()
                .map(parse_slide_animation_extension)
                .transpose()?
                .unwrap_or_default();
            let (shape_ids, legacy) =
                collect_shapes_and_legacy(persist_id, scope, &parsed, limits)?;
            validate_extension(&extension, &shape_ids, limits)?;
            entries.push(PersistAnimation {
                persist_id,
                scope,
                record,
                extension_payload,
                extension,
                shape_ids,
                legacy,
            });
        }
        Ok(Self {
            package,
            entries,
            limits,
            changed: false,
        })
    }

    pub fn is_changed(&self) -> bool {
        self.changed
    }

    pub fn timelines(&self) -> Vec<Timeline> {
        self.entries
            .iter()
            .map(|entry| Timeline {
                persist_id: entry.persist_id,
                scope: entry.scope,
                extension: entry.extension.clone(),
            })
            .collect()
    }

    pub fn find(&self, persist_id: u32) -> Option<Timeline> {
        self.entries
            .iter()
            .find(|entry| entry.persist_id == persist_id)
            .map(|entry| Timeline {
                persist_id,
                scope: entry.scope,
                extension: entry.extension.clone(),
            })
    }

    pub fn legacy_shape_animations(&self) -> Vec<LegacyShapeAnimation> {
        self.entries
            .iter()
            .flat_map(|entry| entry.legacy.clone())
            .collect()
    }

    pub fn find_shape(&self, persist_id: u32, shape_id: u32) -> Option<LegacyShapeAnimation> {
        self.entries
            .iter()
            .find(|entry| entry.persist_id == persist_id)
            .and_then(|entry| entry.legacy.iter().find(|value| value.shape_id == shape_id))
            .cloned()
    }

    /// Adds a child to the root timing container at `index`.
    pub fn add(&mut self, persist_id: u32, index: usize, node: ExtendedTimeNode) -> Result<()> {
        let mut candidate = self.clone();
        let entry = candidate.entry_mut(persist_id)?;
        let mut root = entry.extension.time_node.clone().unwrap_or_default();
        if index > root.children.len() {
            return invalid("timeline insertion index is out of range");
        }
        root.children.insert(index, node);
        entry.extension.time_node = Some(root);
        candidate.stage(persist_id)?;
        *self = candidate;
        Ok(())
    }

    /// Updates one root child in place.
    pub fn update(&mut self, persist_id: u32, index: usize, node: ExtendedTimeNode) -> Result<()> {
        let mut candidate = self.clone();
        let entry = candidate.entry_mut(persist_id)?;
        let root = entry
            .extension
            .time_node
            .as_mut()
            .ok_or_else(|| PptError::InvalidFormat("timeline has no root node".into()))?;
        let slot = root.children.get_mut(index).ok_or_else(|| {
            PptError::InvalidFormat("timeline child index is out of range".into())
        })?;
        *slot = node;
        candidate.stage(persist_id)?;
        *self = candidate;
        Ok(())
    }

    /// Replaces the root timeline and build list atomically.
    pub fn replace(
        &mut self,
        persist_id: u32,
        root: Option<ExtendedTimeNode>,
        builds: Option<BuildList>,
    ) -> Result<()> {
        let mut candidate = self.clone();
        let entry = candidate.entry_mut(persist_id)?;
        entry.extension.time_node = root;
        entry.extension.build_list = builds;
        candidate.stage(persist_id)?;
        *self = candidate;
        Ok(())
    }

    pub fn remove(&mut self, persist_id: u32, index: usize) -> Result<ExtendedTimeNode> {
        let mut candidate = self.clone();
        let entry = candidate.entry_mut(persist_id)?;
        let root = entry
            .extension
            .time_node
            .as_mut()
            .ok_or_else(|| PptError::InvalidFormat("timeline has no root node".into()))?;
        if index >= root.children.len() {
            return invalid("timeline child index is out of range");
        }
        let removed = root.children.remove(index);
        candidate.stage(persist_id)?;
        *self = candidate;
        Ok(removed)
    }

    pub fn reorder(&mut self, persist_id: u32, order: &[usize]) -> Result<()> {
        let mut candidate = self.clone();
        let entry = candidate.entry_mut(persist_id)?;
        let root = entry
            .extension
            .time_node
            .as_mut()
            .ok_or_else(|| PptError::InvalidFormat("timeline has no root node".into()))?;
        if order.len() != root.children.len() {
            return invalid("timeline reorder must include every child");
        }
        let mut seen = HashSet::new();
        if order
            .iter()
            .any(|index| *index >= root.children.len() || !seen.insert(*index))
        {
            return invalid("timeline reorder contains an invalid or duplicate index");
        }
        let old = root.children.clone();
        root.children = order.iter().map(|index| old[*index].clone()).collect();
        candidate.stage(persist_id)?;
        *self = candidate;
        Ok(())
    }

    pub fn add_build(
        &mut self,
        persist_id: u32,
        index: usize,
        build: BuildListEntry,
    ) -> Result<()> {
        let mut candidate = self.clone();
        let entry = candidate.entry_mut(persist_id)?;
        let list = entry
            .extension
            .build_list
            .get_or_insert_with(BuildList::new);
        if index > list.builds.len() {
            return invalid("build insertion index is out of range");
        }
        list.builds.insert(index, build);
        candidate.stage(persist_id)?;
        *self = candidate;
        Ok(())
    }

    pub fn update_build(
        &mut self,
        persist_id: u32,
        index: usize,
        build: BuildListEntry,
    ) -> Result<()> {
        let mut candidate = self.clone();
        let entry = candidate.entry_mut(persist_id)?;
        let list = entry
            .extension
            .build_list
            .as_mut()
            .ok_or_else(|| PptError::InvalidFormat("slide/master has no build list".into()))?;
        let slot = list
            .builds
            .get_mut(index)
            .ok_or_else(|| PptError::InvalidFormat("build index is out of range".into()))?;
        *slot = build;
        candidate.stage(persist_id)?;
        *self = candidate;
        Ok(())
    }

    pub fn remove_build(&mut self, persist_id: u32, index: usize) -> Result<BuildListEntry> {
        let mut candidate = self.clone();
        let entry = candidate.entry_mut(persist_id)?;
        let list = entry
            .extension
            .build_list
            .as_mut()
            .ok_or_else(|| PptError::InvalidFormat("slide/master has no build list".into()))?;
        if index >= list.builds.len() {
            return invalid("build index is out of range");
        }
        let removed = list.builds.remove(index);
        candidate.stage(persist_id)?;
        *self = candidate;
        Ok(removed)
    }

    pub fn reorder_builds(&mut self, persist_id: u32, order: &[usize]) -> Result<()> {
        let mut candidate = self.clone();
        let entry = candidate.entry_mut(persist_id)?;
        let list = entry
            .extension
            .build_list
            .as_mut()
            .ok_or_else(|| PptError::InvalidFormat("slide/master has no build list".into()))?;
        if order.len() != list.builds.len() {
            return invalid("build reorder must include every entry");
        }
        let mut seen = HashSet::new();
        if order
            .iter()
            .any(|index| *index >= list.builds.len() || !seen.insert(*index))
        {
            return invalid("build reorder contains an invalid or duplicate index");
        }
        let old = list.builds.clone();
        list.builds = order.iter().map(|index| old[*index].clone()).collect();
        candidate.stage(persist_id)?;
        *self = candidate;
        Ok(())
    }

    pub fn replace_shape_animation(
        &mut self,
        persist_id: u32,
        shape_id: u32,
        animation: Option<AnimationInfo>,
    ) -> Result<()> {
        let mut candidate = self.clone();
        let index = candidate
            .entries
            .iter()
            .position(|entry| entry.persist_id == persist_id)
            .ok_or_else(|| PptError::InvalidFormat("animation persist ID was not found".into()))?;
        if !candidate.entries[index].shape_ids.contains(&shape_id) {
            return invalid("animation target shape does not exist");
        }
        let bytes = animation
            .as_ref()
            .map(write_animation_info)
            .transpose()?
            .map(|value| value.0);
        let (record, found) =
            rewrite_shape_animation(&candidate.entries[index].record, shape_id, bytes.as_deref())?;
        if !found {
            return corrupted("target OfficeArt shape container was not found");
        }
        candidate
            .package
            .replace_persisted_record(persist_id, record.clone())?;
        candidate.entries[index].record = record;
        candidate.entries[index]
            .legacy
            .retain(|value| value.shape_id != shape_id);
        if let Some(animation) = animation {
            let scope = candidate.entries[index].scope;
            candidate.entries[index].legacy.push(LegacyShapeAnimation {
                persist_id,
                scope,
                shape_id,
                animation,
            });
            candidate.entries[index]
                .legacy
                .sort_by_key(|value| value.shape_id);
        }
        candidate.changed = true;
        *self = candidate;
        Ok(())
    }

    pub fn finish(self) -> Result<Vec<u8>> {
        self.package.finish()
    }

    fn entry_mut(&mut self, persist_id: u32) -> Result<&mut PersistAnimation> {
        self.entries
            .iter_mut()
            .find(|entry| entry.persist_id == persist_id)
            .ok_or_else(|| PptError::InvalidFormat("animation persist ID was not found".into()))
    }

    fn stage(&mut self, persist_id: u32) -> Result<()> {
        let index = self
            .entries
            .iter()
            .position(|entry| entry.persist_id == persist_id)
            .ok_or_else(|| PptError::InvalidFormat("animation persist ID was not found".into()))?;
        validate_extension(
            &self.entries[index].extension,
            &self.entries[index].shape_ids,
            self.limits,
        )?;
        let payload = rewrite_extension_payload(
            self.entries[index].extension_payload.as_deref(),
            self.entries[index].extension.time_node.as_ref(),
            self.entries[index].extension.build_list.as_ref(),
        )?;
        if payload.len() > self.limits.max_record_bytes {
            return invalid("animation extension exceeds record limit");
        }
        let record = rewrite_ppt10_payload(&self.entries[index].record, &payload)?;
        let (parsed, consumed) = PptRecord::parse(&record, 0)?;
        if consumed != record.len() || parsed.data_length as usize + 8 != record.len() {
            return corrupted("rewritten slide/master record failed length validation");
        }
        let reparsed_payload = find_ppt10_payload(&parsed)?
            .ok_or_else(|| PptError::Corrupted("rewritten ___PPT10 payload is missing".into()))?;
        let reparsed = parse_slide_animation_extension(&reparsed_payload)?;
        if reparsed.time_node != self.entries[index].extension.time_node
            || reparsed.build_list != self.entries[index].extension.build_list
        {
            return corrupted("rewritten animation timeline failed round-trip validation");
        }
        self.package
            .replace_persisted_record(persist_id, record.clone())?;
        self.entries[index].record = record;
        self.entries[index].extension_payload = Some(payload);
        self.entries[index].extension = reparsed;
        self.changed = true;
        Ok(())
    }
}

fn validate_limits(limits: EditorLimits) -> Result<()> {
    if limits.max_persist_records == 0
        || limits.max_record_bytes < 8
        || limits.max_timeline_nodes == 0
        || limits.max_timeline_depth == 0
        || limits.max_build_entries == 0
        || limits.max_shapes == 0
    {
        return invalid("all animation editor limits must be nonzero");
    }
    Ok(())
}

fn validate_extension(
    extension: &SlideAnimationExtension,
    shapes: &BTreeSet<u32>,
    limits: EditorLimits,
) -> Result<()> {
    let mut count = 0usize;
    if let Some(root) = &extension.time_node {
        validate_node(root, 1, &mut count, shapes, limits)?;
        let _ = write_extended_time_node(root)?;
    }
    if let Some(builds) = &extension.build_list {
        if builds.builds.len() > limits.max_build_entries {
            return invalid("build list exceeds resource limit");
        }
        let mut ids = HashSet::new();
        for build in &builds.builds {
            let atom = match build {
                BuildListEntry::Paragraph(value) => &value.atom,
                BuildListEntry::Chart(value) => &value.atom,
                BuildListEntry::Diagram(value) => &value.atom,
            };
            if !shapes.contains(&atom.shape_id_ref) {
                return invalid("build atom references a missing shape");
            }
            if !ids.insert(atom.build_id) {
                return invalid("build list contains duplicate build IDs");
            }
        }
        let _ = write_build_list(builds)?;
    }
    Ok(())
}

fn validate_node(
    node: &ExtendedTimeNode,
    depth: usize,
    count: &mut usize,
    shapes: &BTreeSet<u32>,
    limits: EditorLimits,
) -> Result<()> {
    *count = count
        .checked_add(1)
        .ok_or_else(|| PptError::InvalidFormat("timeline node count overflow".into()))?;
    if depth > limits.max_timeline_depth || *count > limits.max_timeline_nodes {
        return invalid("timeline nesting or node count exceeds limits");
    }
    if let Some(target) = &node.visual_target {
        validate_target(target, shapes)?;
    }
    if let Some(behavior) = &node.behavior {
        let target = match behavior {
            TimeNodeBehavior::Animate(v) => &v.behavior.target,
            TimeNodeBehavior::Color(v) => &v.behavior.target,
            TimeNodeBehavior::Effect(v) => &v.behavior.target,
            TimeNodeBehavior::Motion(v) => &v.behavior.target,
            TimeNodeBehavior::Rotation(v) => &v.behavior.target,
            TimeNodeBehavior::Scale(v) => &v.behavior.target,
            TimeNodeBehavior::Set(v) => &v.behavior.target,
            TimeNodeBehavior::Command(v) => &v.behavior.target,
        };
        validate_target(target, shapes)?;
    }
    for modifier in &node.modifiers {
        validate_modifier(modifier)?;
    }
    for effect in &node.sub_effects {
        if let Some(target) = &effect.visual_target {
            validate_target(target, shapes)?;
        }
        if let Some(behavior) = &effect.behavior {
            let target = match behavior {
                TimeSubEffectBehavior::Color(v) => &v.behavior.target,
                TimeSubEffectBehavior::Set(v) => &v.behavior.target,
                TimeSubEffectBehavior::Command(v) => &v.behavior.target,
            };
            validate_target(target, shapes)?;
        }
        for modifier in &effect.modifiers {
            validate_modifier(modifier)?;
        }
    }
    for child in &node.children {
        validate_node(child, depth + 1, count, shapes, limits)?;
    }
    Ok(())
}

fn validate_target(target: &TimeVisualElement, shapes: &BTreeSet<u32>) -> Result<()> {
    match target {
        TimeVisualElement::Shape {
            shape_id_ref,
            data1,
            data2,
            ..
        } => {
            if !shapes.contains(shape_id_ref) {
                return invalid("behavior references a missing shape");
            }
            if *data1 < -1 || *data2 < -1 {
                return invalid("text-range target contains an invalid range");
            }
        },
        TimeVisualElement::Chart { shape_id_ref, .. } if !shapes.contains(shape_id_ref) => {
            return invalid("chart behavior references a missing shape");
        },
        _ => {},
    }
    Ok(())
}

fn validate_modifier(_value: &TimeModifier) -> Result<()> {
    Ok(())
}

fn find_ppt10_payload(record: &PptRecord) -> Result<Option<Vec<u8>>> {
    for prog_tags in record.find_children(PptRecordType::ProgTags) {
        for binary in prog_tags.find_children(PptRecordType::ProgBinaryTag) {
            let Some(name) = binary.find_child(PptRecordType::CString) else {
                continue;
            };
            if is_ppt10_name(&name.data) {
                let data = binary
                    .find_child(PptRecordType::BinaryTagData)
                    .ok_or_else(|| {
                        PptError::Corrupted("___PPT10 tag is missing BinaryTagData".into())
                    })?;
                return Ok(Some(data.data.clone()));
            }
        }
    }
    Ok(None)
}

fn is_ppt10_name(data: &[u8]) -> bool {
    data == "___PPT10"
        .encode_utf16()
        .flat_map(u16::to_le_bytes)
        .collect::<Vec<_>>()
}

fn rewrite_extension_payload(
    original: Option<&[u8]>,
    root: Option<&ExtendedTimeNode>,
    builds: Option<&BuildList>,
) -> Result<Vec<u8>> {
    let mut output = Vec::new();
    let mut saw_root = false;
    let mut saw_builds = false;
    let data = original.unwrap_or(&[]);
    let mut offset = 0;
    while offset < data.len() {
        let raw = raw_record(data, offset)?;
        let kind = raw_type(raw)?;
        if kind == PptRecordType::ExtTimeNode.as_u16() {
            if saw_root {
                return invalid("extension contains duplicate root timelines");
            }
            saw_root = true;
            if let Some(value) = root {
                output.extend(write_extended_time_node(value)?);
            }
        } else if kind == PptRecordType::BuildList.as_u16() {
            if saw_builds {
                return invalid("extension contains duplicate build lists");
            }
            saw_builds = true;
            if let Some(value) = builds {
                output.extend(write_build_list(value)?);
            }
        } else {
            output.extend_from_slice(raw);
        }
        offset += raw.len();
    }
    if !saw_root && let Some(value) = root {
        output.extend(write_extended_time_node(value)?);
    }
    if !saw_builds && let Some(value) = builds {
        output.extend(write_build_list(value)?);
    }
    Ok(output)
}

fn rewrite_ppt10_payload(root: &[u8], payload: &[u8]) -> Result<Vec<u8>> {
    let (rewritten, found) = rewrite_ppt_record(root, &mut |record| {
        if raw_type(record).ok() != Some(PptRecordType::ProgBinaryTag.as_u16()) {
            return Ok(None);
        }
        if !prog_binary_is_ppt10(record)? {
            return Ok(None);
        }
        Ok(Some(rewrite_binary_tag_data(record, payload)?))
    })?;
    if found {
        return Ok(rewritten);
    }
    if raw_version(&rewritten)? != 0x0F {
        return corrupted("slide/master persist record is not a container");
    }
    let mut binary_children = atom(
        0,
        0,
        PptRecordType::CString.as_u16(),
        &"___PPT10"
            .encode_utf16()
            .flat_map(u16::to_le_bytes)
            .collect::<Vec<_>>(),
    )?;
    binary_children.extend(atom(0, 0, PptRecordType::BinaryTagData.as_u16(), payload)?);
    let binary = container(0, PptRecordType::ProgBinaryTag.as_u16(), &binary_children)?;
    let tags = container(0, PptRecordType::ProgTags.as_u16(), &binary)?;
    append_container_child(&rewritten, &tags)
}

fn prog_binary_is_ppt10(record: &[u8]) -> Result<bool> {
    let mut offset = 8;
    while offset < record.len() {
        let child = raw_record(record, offset)?;
        if raw_type(child)? == PptRecordType::CString.as_u16() {
            return Ok(is_ppt10_name(&child[8..]));
        }
        offset += child.len();
    }
    Ok(false)
}
fn rewrite_binary_tag_data(record: &[u8], payload: &[u8]) -> Result<Vec<u8>> {
    let mut out = Vec::new();
    let mut offset = 8;
    let mut found = false;
    while offset < record.len() {
        let child = raw_record(record, offset)?;
        if raw_type(child)? == PptRecordType::BinaryTagData.as_u16() {
            if found {
                return invalid("___PPT10 contains duplicate BinaryTagData");
            }
            out.extend(atom(0, 0, PptRecordType::BinaryTagData.as_u16(), payload)?);
            found = true;
        } else {
            out.extend_from_slice(child);
        }
        offset += child.len();
    }
    if !found {
        return corrupted("___PPT10 tag is missing BinaryTagData");
    }
    rebuild_record(record, &out)
}

fn collect_shapes_and_legacy(
    persist_id: u32,
    scope: Scope,
    record: &PptRecord,
    limits: EditorLimits,
) -> Result<(BTreeSet<u32>, Vec<LegacyShapeAnimation>)> {
    let mut ids = BTreeSet::new();
    let mut legacy = Vec::new();
    for drawing in record.find_children(PptRecordType::PPDrawing) {
        collect_escher(
            &drawing.data,
            persist_id,
            scope,
            &mut ids,
            &mut legacy,
            limits,
        )?;
    }
    Ok((ids, legacy))
}
fn collect_escher(
    data: &[u8],
    persist: u32,
    scope: Scope,
    ids: &mut BTreeSet<u32>,
    legacy: &mut Vec<LegacyShapeAnimation>,
    limits: EditorLimits,
) -> Result<()> {
    let mut offset = 0;
    while offset < data.len() {
        let record = raw_record(data, offset)?;
        if raw_type(record)? == ESCHER_SP_CONTAINER {
            let mut shape_id = None;
            let mut child = 8;
            while child < record.len() {
                let value = raw_record(record, child)?;
                match raw_type(value)? {
                    ESCHER_SP if value.len() >= 16 => shape_id = Some(u32_at(value, 8)?),
                    ESCHER_CLIENT_DATA => {
                        let mut ppt = 8;
                        while ppt < value.len() {
                            let item = raw_record(value, ppt)?;
                            if raw_type(item)? == PptRecordType::AnimationInfo.as_u16() {
                                let (parsed, used) = PptRecord::parse(item, 0)?;
                                if used != item.len() {
                                    return corrupted("AnimationInfo length mismatch");
                                }
                                if let Some(id) = shape_id {
                                    legacy.push(LegacyShapeAnimation {
                                        persist_id: persist,
                                        scope,
                                        shape_id: id,
                                        animation: parse_animation_info(&parsed)?,
                                    });
                                }
                            }
                            ppt += item.len();
                        }
                    },
                    _ => {},
                }
                child += value.len();
            }
            if let Some(id) = shape_id
                && (ids.len() >= limits.max_shapes || !ids.insert(id))
            {
                return invalid("shape identifier limit or uniqueness violation");
            }
        }
        if raw_version(record)? == 0x0F {
            collect_escher(&record[8..], persist, scope, ids, legacy, limits)?;
        }
        offset += record.len();
    }
    Ok(())
}

fn rewrite_shape_animation(
    root: &[u8],
    shape_id: u32,
    animation: Option<&[u8]>,
) -> Result<(Vec<u8>, bool)> {
    rewrite_ppt_record(root, &mut |record| {
        if raw_type(record).ok() != Some(PptRecordType::PPDrawing.as_u16()) {
            return Ok(None);
        };
        let (data, found) = rewrite_escher_shapes(&record[8..], shape_id, animation)?;
        if found {
            Ok(Some(rebuild_record(record, &data)?))
        } else {
            Ok(None)
        }
    })
}
fn rewrite_escher_shapes(
    data: &[u8],
    shape_id: u32,
    animation: Option<&[u8]>,
) -> Result<(Vec<u8>, bool)> {
    let mut out = Vec::new();
    let mut offset = 0;
    let mut found = false;
    while offset < data.len() {
        let record = raw_record(data, offset)?;
        let mut changed = None;
        if raw_type(record)? == ESCHER_SP_CONTAINER {
            let mut id = None;
            let mut child = 8;
            while child < record.len() {
                let value = raw_record(record, child)?;
                if raw_type(value)? == ESCHER_SP && value.len() >= 16 {
                    id = Some(u32_at(value, 8)?);
                }
                child += value.len();
            }
            if id == Some(shape_id) {
                changed = Some(rewrite_shape_container(record, animation)?);
                found = true;
            }
        }
        if changed.is_none() && raw_version(record)? == 0x0F {
            let (payload, hit) = rewrite_escher_shapes(&record[8..], shape_id, animation)?;
            if hit {
                changed = Some(rebuild_record(record, &payload)?);
                found = true;
            }
        }
        out.extend_from_slice(changed.as_deref().unwrap_or(record));
        offset += record.len();
    }
    Ok((out, found))
}
fn rewrite_shape_container(record: &[u8], animation: Option<&[u8]>) -> Result<Vec<u8>> {
    let mut out = Vec::new();
    let mut offset = 8;
    let mut client_found = false;
    while offset < record.len() {
        let child = raw_record(record, offset)?;
        if raw_type(child)? == ESCHER_CLIENT_DATA {
            if client_found {
                return invalid("shape has multiple ClientData records");
            }
            client_found = true;
            let mut payload = Vec::new();
            let mut ppt = 8;
            let mut anim_found = false;
            while ppt < child.len() {
                let item = raw_record(child, ppt)?;
                if raw_type(item)? == PptRecordType::AnimationInfo.as_u16() {
                    if anim_found {
                        return invalid("shape has duplicate AnimationInfo records");
                    }
                    anim_found = true;
                    if let Some(value) = animation {
                        payload.extend_from_slice(value);
                    }
                } else {
                    payload.extend_from_slice(item);
                }
                ppt += item.len();
            }
            if !anim_found && let Some(value) = animation {
                payload.extend_from_slice(value);
            }
            out.extend(escher_record(0x0f, 0, ESCHER_CLIENT_DATA, &payload)?);
        } else {
            out.extend_from_slice(child);
        }
        offset += child.len();
    }
    if !client_found && let Some(value) = animation {
        out.extend(escher_record(0x0f, 0, ESCHER_CLIENT_DATA, value)?);
    }
    rebuild_record(record, &out)
}

fn rewrite_ppt_record(
    record: &[u8],
    edit: &mut impl FnMut(&[u8]) -> Result<Option<Vec<u8>>>,
) -> Result<(Vec<u8>, bool)> {
    if let Some(value) = edit(record)? {
        return Ok((value, true));
    }
    if raw_version(record)? != 0x0F {
        return Ok((record.to_vec(), false));
    }
    let mut out = Vec::new();
    let mut offset = 8;
    let mut found = false;
    while offset < record.len() {
        let child = raw_record(record, offset)?;
        let (value, hit) = rewrite_ppt_record(child, edit)?;
        out.extend(value);
        found |= hit;
        offset += child.len();
    }
    if found {
        Ok((rebuild_record(record, &out)?, true))
    } else {
        Ok((record.to_vec(), false))
    }
}
fn append_container_child(record: &[u8], child: &[u8]) -> Result<Vec<u8>> {
    let mut payload = record
        .get(8..)
        .ok_or_else(|| PptError::Corrupted("container header is truncated".into()))?
        .to_vec();
    payload.extend_from_slice(child);
    rebuild_record(record, &payload)
}
fn rebuild_record(record: &[u8], payload: &[u8]) -> Result<Vec<u8>> {
    let mut out = record
        .get(..4)
        .ok_or_else(|| PptError::Corrupted("record header is truncated".into()))?
        .to_vec();
    out.extend_from_slice(
        &u32::try_from(payload.len())
            .map_err(|_| PptError::InvalidFormat("record payload exceeds u32".into()))?
            .to_le_bytes(),
    );
    out.extend_from_slice(payload);
    Ok(out)
}
fn atom(version: u16, instance: u16, kind: u16, payload: &[u8]) -> Result<Vec<u8>> {
    record_bytes(version, instance, kind, payload)
}
fn container(instance: u16, kind: u16, payload: &[u8]) -> Result<Vec<u8>> {
    record_bytes(0x0F, instance, kind, payload)
}
fn escher_record(version: u16, instance: u16, kind: u16, payload: &[u8]) -> Result<Vec<u8>> {
    record_bytes(version, instance, kind, payload)
}
fn record_bytes(version: u16, instance: u16, kind: u16, payload: &[u8]) -> Result<Vec<u8>> {
    let mut out = ((instance << 4) | (version & 0xF)).to_le_bytes().to_vec();
    out.extend_from_slice(&kind.to_le_bytes());
    out.extend_from_slice(
        &u32::try_from(payload.len())
            .map_err(|_| PptError::InvalidFormat("record payload exceeds u32".into()))?
            .to_le_bytes(),
    );
    out.extend_from_slice(payload);
    Ok(out)
}
fn raw_record(data: &[u8], offset: usize) -> Result<&[u8]> {
    let len = u32_at(data, offset + 4)? as usize;
    let end = offset
        .checked_add(8)
        .and_then(|v| v.checked_add(len))
        .ok_or_else(|| PptError::Corrupted("record length overflow".into()))?;
    data.get(offset..end)
        .ok_or_else(|| PptError::Corrupted("record is truncated".into()))
}
fn raw_type(record: &[u8]) -> Result<u16> {
    record
        .get(2..4)
        .map(|v| u16::from_le_bytes([v[0], v[1]]))
        .ok_or_else(|| PptError::Corrupted("record type is truncated".into()))
}
fn raw_version(record: &[u8]) -> Result<u16> {
    record
        .get(0..2)
        .map(|v| u16::from_le_bytes([v[0], v[1]]) & 0xF)
        .ok_or_else(|| PptError::Corrupted("record version is truncated".into()))
}
fn u32_at(data: &[u8], offset: usize) -> Result<u32> {
    data.get(offset..offset + 4)
        .map(|v| u32::from_le_bytes(v.try_into().unwrap()))
        .ok_or_else(|| PptError::Corrupted("record u32 is truncated".into()))
}
fn invalid<T>(message: impl Into<String>) -> Result<T> {
    Err(PptError::InvalidFormat(message.into()))
}
fn corrupted<T>(message: impl Into<String>) -> Result<T> {
    Err(PptError::Corrupted(message.into()))
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::animation::{TimeNodeAtom, TimeNodeKind};

    #[test]
    fn ppt10_payload_replacement_preserves_unknown_records() {
        let unknown = atom(0, 0, 0x7777, b"opaque").unwrap();
        let old = write_extended_time_node(&ExtendedTimeNode::default()).unwrap();
        let mut payload = unknown.clone();
        payload.extend(old);
        let rewritten = rewrite_extension_payload(
            Some(&payload),
            Some(&ExtendedTimeNode {
                atom: TimeNodeAtom {
                    node_type: Some(TimeNodeKind::Sequential),
                    duration_ms: Some(500),
                    ..Default::default()
                },
                ..Default::default()
            }),
            None,
        )
        .unwrap();
        assert!(
            rewritten
                .windows(unknown.len())
                .any(|value| value == unknown)
        );
        assert_eq!(
            parse_slide_animation_extension(&rewritten)
                .unwrap()
                .time_node
                .unwrap()
                .atom
                .duration_ms,
            Some(500)
        );
    }

    #[test]
    fn nested_timeline_limits_and_bad_reorders_roll_back_before_staging() {
        let mut node = ExtendedTimeNode::default();
        node.children.push(ExtendedTimeNode::default());
        let mut count = 0;
        assert!(
            validate_node(
                &node,
                1,
                &mut count,
                &BTreeSet::new(),
                EditorLimits {
                    max_timeline_depth: 1,
                    ..Default::default()
                }
            )
            .is_err()
        );
    }

    #[test]
    fn malformed_ppt10_payload_is_rejected() {
        assert!(rewrite_extension_payload(Some(&[0; 7]), None, None).is_err());
    }

    #[test]
    fn legacy_shape_edit_preserves_inert_interactive_records() {
        let interactive =
            atom(0, 0, PptRecordType::InteractiveInfoAtom.as_u16(), &[0; 16]).unwrap();
        let animation = write_animation_info(&AnimationInfo::new()).unwrap().0;
        let mut client_payload = interactive.clone();
        client_payload.extend(animation);
        let client = escher_record(0x0f, 0, ESCHER_CLIENT_DATA, &client_payload).unwrap();
        let mut shape_payload = escher_record(2, 0, ESCHER_SP, &[42, 0, 0, 0, 0, 0, 0, 0]).unwrap();
        shape_payload.extend(client);
        let shape = escher_record(0x0f, 0, ESCHER_SP_CONTAINER, &shape_payload).unwrap();
        let drawing = atom(0, 0, PptRecordType::PPDrawing.as_u16(), &shape).unwrap();
        let slide = container(0, PptRecordType::Slide.as_u16(), &drawing).unwrap();

        let (rewritten, found) = rewrite_shape_animation(&slide, 42, None).unwrap();
        assert!(found);
        assert!(
            rewritten
                .windows(interactive.len())
                .any(|value| value == interactive)
        );
        let (record, used) = PptRecord::parse(&rewritten, 0).unwrap();
        assert_eq!(used, rewritten.len());
        let (shapes, legacy) =
            collect_shapes_and_legacy(1, Scope::Slide, &record, EditorLimits::default()).unwrap();
        assert!(shapes.contains(&42));
        assert!(legacy.is_empty());
    }
}
